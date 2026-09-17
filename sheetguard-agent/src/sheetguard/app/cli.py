"""SheetGuard V1 的安全 Typer 命令行接口。"""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Annotated

import typer
from openpyxl import load_workbook

from sheetguard.app.report import render_report
from sheetguard.app import review_store as rs
from sheetguard.app.rounds import (
    BATCH_SUCCESS_STATUSES,
    base_dir as _base_dir,
    round_base as _round_base,
    round_dir as _round_dir,
    write_json as _write_json,
    write_round_artifacts as _write_round_artifacts,  # re-export：旧测试断言 cli._write_round_artifacts
)
from sheetguard.app.inspection import inspection_payload as _inspection_payload
from sheetguard.app.review_service import (
    _error_type_of as _error_type_of,  # re-export：旧测试仍用 cli._error_type_of
    _find_audit_entry as _find_audit_entry,
    _find_dismissed_reason as _find_dismissed_reason,
    build_review_items as build_review_items,  # re-export：旧测试/调用方仍用 cli.build_review_items
    collect_history as _collect_history,
    load_review_state as _load_review_state,
    refresh_workbook_index as _refresh_workbook_index_service,
    run_next_round as _run_next_round,
    save_review_draft as _save_review_draft,
    submit_review as _submit_review,
    ReviewStateError as _ReviewStateError,
)
from sheetguard.agent.baseline import BaselineAgent
from sheetguard.agent.graph import SheetGuardGraph
from sheetguard.agent.llm_config import build_model
from sheetguard.config import ConfigError, SheetGuardConfig, load_config
from evaluation.generation.workbook import build_gold_model
from sheetguard.spreadsheet.dependency_graph import DependencyGraph
from sheetguard.spreadsheet.parser import parse_workbook

app = typer.Typer(
    name="sheetguard",
    help="Diagnose and safely propose repairs for Excel formulas.",
    no_args_is_help=True,
    add_completion=False,
)

def _fail(message: str, exit_code: int = 1) -> None:
    typer.echo(f"Error: {message}", err=True)
    raise typer.Exit(exit_code)


def _workbook_path(path: Path) -> Path:
    if not path.exists() or not path.is_file():
        _fail(f"workbook not found: {path}")
    if path.suffix.lower() != ".xlsx":
        _fail(f"workbook must be an .xlsx file: {path}")
    return path


def _require_model_configuration() -> None:
    """在未配置真实 provider key/model 时，在任何网络调用前直接失败。"""
    from sheetguard.agent.llm_config import model_configured

    if not model_configured():
        _fail(
            "LLM configuration is required. Set OPENAI_API_KEY "
            "in the environment or .env file."
        )


def _resolve_config(
    config_path: Path | None,
    max_attempts: int | None,
    max_retries: int | None,
    min_anomaly_score: float | None,
    max_candidates: int | None,
) -> "SheetGuardConfig":
    """按「CLI 覆盖 > YAML > 默认值」解析运行配置；非法时给出友好报错。"""
    overrides: dict[str, int | float | None] = {}
    if max_attempts is not None:
        overrides["max_attempts"] = max_attempts
    elif max_retries is not None:
        overrides["max_attempts"] = max_retries
    overrides["min_anomaly_score"] = min_anomaly_score
    overrides["max_candidates"] = max_candidates
    try:
        return load_config(config_path, **overrides)
    except ConfigError as exc:
        _fail(f"invalid config: {exc}")


@app.command()
def generate(
    output: Annotated[Path, typer.Argument(help="Destination .xlsx workbook path")],
    seed: Annotated[int, typer.Option(help="Deterministic generator seed")] = 0,
    months: Annotated[int, typer.Option(min=1, help="Number of monthly columns")] = 12,
    force: Annotated[bool, typer.Option("--force", help="Replace an existing workbook")] = False,
) -> None:
    """生成确定性的合成 gold 工作簿；不需要 LLM。"""
    if output.exists() and not force:
        _fail(f"output already exists: {output} (use --force to replace it)")
    if output.suffix.lower() != ".xlsx":
        _fail(f"output must end in .xlsx: {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    build_gold_model(seed=seed, n_months=months).save(output)
    typer.echo(f"Generated gold workbook: {output}")


@app.command()
def inspect(
    workbook: Annotated[Path, typer.Argument(help="Workbook to inspect")],
    output: Annotated[Path | None, typer.Option("--output", "-o", help="Write JSON report here")] = None,
    json_output: Annotated[bool, typer.Option("--json", help="Print JSON instead of a summary")] = False,
) -> None:
    """运行离线解析、依赖分析和静态异常检测。"""
    path = _workbook_path(workbook)
    try:
        payload = _inspection_payload(path)
    except Exception as exc:
        _fail(f"could not inspect {path}: {type(exc).__name__}: {exc}")

    if output is not None:
        _write_json(output, payload)
        typer.echo(f"Inspection report saved to: {output}")
    if json_output:
        typer.echo(json.dumps(payload, indent=2, ensure_ascii=False))
        return
    typer.echo(f"SheetGuard inspection: {path.name}")
    typer.echo(f"  Sheets: {payload['sheets']}")
    typer.echo(f"  Formula cells: {payload['formula_cells']}")
    typer.echo(f"  Cross-sheet dependencies: {payload['cross_sheet_dependencies']}")
    typer.echo(f"  Suspects: {len(payload['candidates'])}")
    for candidate in payload["candidates"][:10]:
        typer.echo(
            f"    {candidate['cell']}  score={candidate['score']:.2f} "
            f"signals={','.join(candidate['signals'])}"
        )


@app.command()
def run(
    workbook: Annotated[Path, typer.Argument(help="Workbook to analyze")],
    output: Annotated[Path | None, typer.Option("--output", "-o", help="Write JSON report here (default: out/<workbook>/round-N/audit.json)")] = None,
    agent_type: Annotated[str, typer.Option("--agent-type", help="baseline or structured")] = "structured",
    config_path: Annotated[
        Path | None,
        typer.Option("--config", help="YAML 行为配置；未填用内置默认（键见 sheetguard.config）"),
    ] = None,
    max_attempts: Annotated[int | None, typer.Option("--max-attempts", min=1, help="Max repair attempts per candidate (includes the first)")] = None,
    max_retries: Annotated[int | None, typer.Option("--max-retries", min=1, help="Deprecated alias of --max-attempts")] = None,
    min_anomaly_score: Annotated[float | None, typer.Option("--min-anomaly-score", min=0.0, max=1.0, help="Seed candidate score threshold (overrides config)")] = None,
    max_candidates: Annotated[int | None, typer.Option("--max-candidates", min=1, help="Active batch size for one run (overrides config)")] = None,
    repaired_workbook: Annotated[Path | None, typer.Option("--repaired-workbook", help="Export the repaired workbook here (default: out/<workbook>/round-N/repaired.xlsx)")] = None,
    report: Annotated[Path | None, typer.Option("--report", help="Markdown 报告路径（default: out/<workbook>/round-N/report.md）")] = None,
) -> None:
    """分析一个工作簿并按依赖关系批量修复多个候选，不修改原始文件。"""
    path = _workbook_path(workbook)
    normalized_agent_type = agent_type.lower()
    if normalized_agent_type not in {"baseline", "structured"}:
        _fail("--agent-type must be baseline or structured")
    _require_model_configuration()

    # structured 分支的三个产物默认集中到 out/<工作簿名>/round-N/；显式路径覆盖默认位置。
    # baseline 分支没有默认产物，audit/report 仅在显式传参时写出。
    audit_path = output
    report_path = report
    try:
        index = parse_workbook(str(path))
        graph = DependencyGraph(index)
        graph.build()
        model = build_model()
        if normalized_agent_type == "baseline":
            result = BaselineAgent(index, graph, model).run(str(path))
            audit = {"workbook": str(path), "agent_type": "baseline", "proposal": result}
        else:
            cfg = _resolve_config(
                config_path, max_attempts, max_retries,
                min_anomaly_score, max_candidates,
            )
            artifacts_dir = _round_dir(path)
            audit_path = output if output is not None else artifacts_dir / "audit.json"
            report_path = report if report is not None else artifacts_dir / "report.md"
            repaired_path = (
                repaired_workbook
                if repaired_workbook is not None
                else artifacts_dir / "repaired.xlsx"
            )
            agent = SheetGuardGraph(model, config=cfg)
            try:
                result = agent.invoke(str(path), repaired_workbook=repaired_path)
            finally:
                agent.cleanup()
            audit = {"workbook": str(path), "agent_type": "structured", **result.get("audit", {})}
    except Exception as exc:
        _fail(f"run failed safely: {type(exc).__name__}: {exc}")

    if output is not None or normalized_agent_type != "baseline":
        _write_json(audit_path, audit)
        typer.echo(f"Audit report saved to: {audit_path}")
    if normalized_agent_type != "baseline" or report is not None:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(render_report(audit), encoding="utf-8")
        typer.echo(f"Human-readable report saved to: {report_path}")
    typer.echo(f"Run agent: {normalized_agent_type}")
    if normalized_agent_type == "baseline":
        typer.echo(f"Proposed target: {audit['proposal'].get('target', 'unknown')}")
    else:
        typer.echo(f"Status: {audit.get('status', 'unknown')}")
        typer.echo(
            f"Fixed: {audit.get('initial_active_candidate_count', 0)} active / "
            f"{len(audit.get('fixed', []))} repaired; "
            f"deferred: {audit.get('deferred_count', 0)}"
        )


@app.command()
def repair(
    workbook: Annotated[Path, typer.Argument(help="Workbook to diagnose")],
    output: Annotated[Path | None, typer.Option("--output", "-o", help="Write JSON audit report here")] = None,
    config_path: Annotated[
        Path | None,
        typer.Option("--config", help="YAML 行为配置；未填用内置默认（键见 sheetguard.config）"),
    ] = None,
    max_attempts: Annotated[int | None, typer.Option("--max-attempts", min=1, help="Max repair attempts per candidate (includes the first)")] = None,
    max_retries: Annotated[int | None, typer.Option("--max-retries", min=1, help="Deprecated alias of --max-attempts")] = None,
    min_anomaly_score: Annotated[float | None, typer.Option("--min-anomaly-score", min=0.0, max=1.0, help="Seed candidate score threshold (overrides config)")] = None,
    max_candidates: Annotated[int | None, typer.Option("--max-candidates", min=1, help="Active batch size for one run (overrides config)")] = None,
    repaired_workbook: Annotated[Path | None, typer.Option("--repaired-workbook", help="Export the repaired workbook here (default: out/<workbook>/round-N/repaired.xlsx)")] = None,
    report: Annotated[Path | None, typer.Option("--report", help="Markdown 报告路径（default: out/<workbook>/round-N/report.md）")] = None,
) -> None:
    """运行有界的 multi-cell 批处理修复工作流，不修改源文件。"""
    path = _workbook_path(workbook)
    _require_model_configuration()
    # 三个产物默认集中到 out/<工作簿名>/round-N/；显式路径覆盖默认位置。
    artifacts_dir = _round_dir(path)
    audit_path = output if output is not None else artifacts_dir / "audit.json"
    report_path = report if report is not None else artifacts_dir / "report.md"
    repaired_path = (
        repaired_workbook
        if repaired_workbook is not None
        else artifacts_dir / "repaired.xlsx"
    )
    agent = None
    try:
        cfg = _resolve_config(
            config_path, max_attempts, max_retries,
            min_anomaly_score, max_candidates,
        )
        agent = SheetGuardGraph(build_model(), config=cfg)
        result = agent.invoke(str(path), repaired_workbook=repaired_path)
    except Exception as exc:
        _fail(f"repair failed safely: {type(exc).__name__}: {exc}")
    finally:
        if agent is not None:
            agent.cleanup()

    audit = {"workbook": str(path), **result.get("audit", {})}
    _write_json(audit_path, audit)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(render_report(audit), encoding="utf-8")
    typer.echo(f"Human-readable report saved to: {report_path}")
    typer.echo(f"Repair status: {result.get('status', 'unknown')}")
    if audit.get("repaired_workbook"):
        typer.echo(f"Repaired workbook: {audit['repaired_workbook']}")
    typer.echo(f"Audit report saved to: {audit_path}")
    if result.get("error"):
        typer.echo(f"Details: {result['error']}", err=True)
    if result.get("status") not in BATCH_SUCCESS_STATUSES:
        raise typer.Exit(1)


def _prompt_verdict(entry: dict, kind: str) -> tuple[str, str | None] | str:
    """交互判定单条候选（v1.11 统一键位）；返回 (verdict, remark) 或 'q'/'x'。

    kind ∈ {"fixed", "failed", "skipped"}：键位含义全类型一致
    （1=提案对 / 2=确实坏但没修对 / 3=本来就没坏 / s=悬着），只有
    「有有效提案」的条目才提供 1（fixed 正常、skipped 带历史提案）；
    failed 与无提案的 skipped（历史 proposal=null 落盘后重现）均不提供
    1——按 1 会令 reviews_payload proposal=None → save_feedback 抛
    FeedbackError，整轮判定丢失。无提案条目渲染 fallback：新公式行
    显示"—"。

    kind = "dismissed"（v1.12）走独立键位 2/s/q/x：agent 调查后认为
    该格没有错误，没有提案可判，1/3 无意义——2 = 推翻（强制重查），
    s = 同意（默认通过，之后不重现）。
    """
    if kind == "dismissed":
        typer.echo(f"\n### {entry.get('target', '—')}")
        typer.echo(f"  当前公式：{entry.get('formula') or '—'}")
        if entry.get("reason"):
            typer.echo(f"  dismissed 原因：{entry['reason']}")
        typer.echo("  agent 调查后认为此格没有错误；s = 同意（默认通过），2 = 其实坏了。")
        options = "[2=其实坏了 / s=同意没坏 / q=存草稿退出 / x=提前提交]"
        while True:
            choice = typer.prompt(f"判定 {options}", default="s").strip().lower()
            if choice == "s":
                return "skipped", None
            if choice == "q":
                return "q"
            if choice == "x":
                return "x"
            if choice == "2":
                remark = typer.prompt(
                    "备注（可选，直接回车跳过；会作为重查的线索）", default=""
                ).strip()
                return "incorrect", (remark or None)
            typer.echo("无效输入，请输入 2/s/q/x")
            continue
    has_proposal = bool(entry.get("new_formula") or entry.get("proposal"))
    offer_one = kind in {"fixed", "skipped"} and has_proposal
    typer.echo(f"\n### {entry.get('target', '—')}"
               + (f"（来自 round-{entry.get('src_round')} 的提案）" if kind == "skipped" else ""))
    typer.echo(f"  原公式：{entry.get('old_formula') or '—'}")
    if kind == "failed":
        typer.echo(f"  最后尝试：{entry.get('last_formula') or '—'}")
        if entry.get("reason"):
            typer.echo(f"  失败原因：{entry['reason']}")
    else:
        typer.echo(f"  新公式：{entry.get('new_formula') or entry.get('proposal') or '—'}")
    hypothesis = entry.get("hypothesis") or {}
    if isinstance(hypothesis, dict) and hypothesis.get("hypothesis"):
        typer.echo(f"  诊断：{hypothesis['hypothesis']}")
    options = (
        "[1=对 / 2=错 / 3=不该修 / s=跳过 / q=存草稿退出 / x=提前提交]"
        if offer_one
        else "[2=确实坏了 / 3=本来就没坏 / s=跳过 / q=存草稿退出 / x=提前提交]"
    )
    while True:
        choice = typer.prompt(f"判定 {options}", default="s").strip().lower()
        if choice == "1" and offer_one:
            return "correct", None
        if choice == "s":
            return "skipped", None
        if choice == "q":
            return "q"
        if choice == "x":
            return "x"
        if choice == "2":
            remark = typer.prompt(
                "备注（可选，直接回车跳过；会作为下一轮修复的线索）", default=""
            ).strip()
            return "incorrect", (remark or None)
        if choice == "3":
            remark = typer.prompt(
                "备注（可选，直接回车跳过；说明为什么这格本来就没坏）", default=""
            ).strip()
            return "false_positive", (remark or None)
        typer.echo(f"无效输入，请输入 {'1/2/3/s/q/x' if offer_one else '2/3/s/q/x'}")
        continue


def _refresh_workbook_index(
    base: Path,
    wb_name: str,
    dataset_dir: Path | None,
    cert_path: Path,
    fp_path: Path,
) -> tuple[str | None, list[str], list[str]]:
    """兼容包装：服务层纯函数 + 降级警告直接 echo stderr（finalize/label 用）。"""
    warnings: list[str] = []
    status, pending, contested = _refresh_workbook_index_service(
        base, wb_name, dataset_dir, cert_path, fp_path, warnings
    )
    for warning in warnings:
        typer.echo(warning, err=True)
    return status, pending, contested


def _cli_repair_fn(cfg: "SheetGuardConfig"):
    """构造下一轮修复执行闭包：延迟读取模块全局，保证测试可 monkeypatch。"""
    def repair_fn(workbook_path, repaired_path, review_context, progress_callback=None):
        agent = None
        try:
            kwargs = {"review_context": review_context}
            if progress_callback is not None:
                kwargs["progress_callback"] = progress_callback
            agent = SheetGuardGraph(build_model(), config=cfg)
            return agent.invoke(str(workbook_path), repaired_workbook=repaired_path, **kwargs)
        finally:
            if agent is not None:
                agent.cleanup()
    return repair_fn


@app.command()
def review(
    target: Annotated[Path, typer.Argument(help="工作簿 .xlsx 或 out/<工作簿名>/ 轮次父目录")],
    config_path: Annotated[
        Path | None,
        typer.Option("--config", help="YAML 行为配置（读取 max_review_rounds）"),
    ] = None,
    dataset_dir: Annotated[
        Path | None,
        typer.Option("--dataset-dir", help="用户数据集目录（默认 evaluation/data/user_feedback/）"),
    ] = None,
) -> None:
    """逐条审查最新一轮的修复提案；提交后对被否条目触发下一轮重修。"""
    base = _base_dir(target)
    # 配置在提交反馈之前加载：非法配置直接失败，避免「反馈已提交但重修不触发」。
    try:
        cfg = load_config(config_path)
    except ConfigError as exc:
        _fail(f"invalid config: {exc}")
    try:
        state = _load_review_state(base)
    except _ReviewStateError as exc:
        _fail(str(exc))
    verdicts: dict[str, tuple[str, str | None]] = state["verdicts"]
    review_items: list[tuple[str, dict]] = state["review_items"]

    for kind, entry in review_items:
        if entry["target"] in verdicts:
            continue
        outcome = _prompt_verdict(entry, kind)
        if outcome == "q":
            _save_review_draft(state, verdicts)
            typer.echo("已存草稿退出；下次 review 将从断点续审（未触发任何修复）。")
            return
        if outcome == "x":
            break
        verdict, remark = outcome
        verdicts[entry["target"]] = (verdict, remark)

    counts = {"correct": 0, "incorrect": 0, "false_positive": 0, "skipped": 0}
    for verdict, _ in verdicts.values():
        counts[verdict] += 1
    unjudged = sum(1 for _, entry in review_items if entry["target"] not in verdicts)
    typer.echo(
        f"\n汇总：对 {counts['correct']} · 错 {counts['incorrect']} · "
        f"误修 {counts['false_positive']} · 跳过 {counts['skipped']} · 未判 {unjudged}"
    )
    while True:
        confirm = typer.prompt(
            "确认提交？提交后：对的认证入库，误修的记录在案，错的触发重新修复 [Y/n]",
            default="Y",
        ).strip().lower()
        if confirm in {"y", "yes", ""}:
            break
        if confirm in {"n", "no"}:
            _save_review_draft(state, verdicts)
            typer.echo("未提交；判定已存草稿。")
            return
        typer.echo("无效输入，请输入 Y 或 n")

    result = _submit_review(state, verdicts, cfg, dataset_dir)
    for line in result["logs"]:
        typer.echo(line)
    for warning in result["warnings"]:
        typer.echo(warning, err=True)
    if result["review_context"] is None:
        if result["incorrect_count"] == 0:
            typer.echo("没有被否条目，流程结束。")
        return

    _require_model_configuration()
    try:
        outcome = _run_next_round(base, result["round_no"], result["review_context"], _cli_repair_fn(cfg))
    except _ReviewStateError as exc:
        _fail(str(exc))
    except Exception as exc:
        _fail(f"review-triggered repair failed safely: {type(exc).__name__}: {exc}")
    typer.echo(f"下一轮修复完成：{outcome['dir'] / 'report.md'}")
    if outcome["status"] not in BATCH_SUCCESS_STATUSES:
        typer.echo(f"警告：下一轮修复状态为 {outcome['status']}，请检查报告", err=True)
    typer.echo(f"状态：{outcome['status']}；运行 sheetguard review 继续审查。")


@app.command()
def finalize(
    target: Annotated[Path, typer.Argument(help="工作簿 .xlsx 或 out/<工作簿名>/ 轮次父目录")],
    dataset_dir: Annotated[
        Path | None,
        typer.Option("--dataset-dir", help="用户数据集目录（默认 evaluation/data/user_feedback/）"),
    ] = None,
) -> None:
    """把 broken_source + 全部认证修复组装为认证成品 certified.xlsx。

    数据源 = certified_cases.json（含 label 补标样本）。workbook_index
    状态如实报告：fully_certified 干净导出；partial 警告并列出仍保持
    原始公式的未终局格（可能是坏公式）。幂等：重复运行直接覆盖。
    """
    base = _base_dir(target)
    wb_name = base.name
    broken_source = base / "broken_source.xlsx"
    if not broken_source.exists():
        _fail(f"broken_source.xlsx 不存在（还没有运行过 repair/run 的留档）：{broken_source}")
    ds_root = dataset_dir
    cert_path = (
        ds_root / rs.CERTIFIED_NAME if ds_root is not None
        else rs.dataset_dir() / rs.CERTIFIED_NAME
    )
    fp_path = (
        ds_root / rs.FALSE_POSITIVE_NAME if ds_root is not None
        else rs.dataset_dir() / rs.FALSE_POSITIVE_NAME
    )
    try:
        samples = [
            s for s in rs.load_samples(cert_path) if s.get("workbook") == wb_name
        ]
    except rs.FeedbackError as exc:
        _fail(f"certified_cases.json 损坏或不可读：{exc}")
    if not samples:
        _fail(f"没有可应用的认证修复（certified_cases.json 中无 {wb_name} 的样本）")
    # 每 target 取最新认证公式：数据集按轮追加，后写覆盖先写。
    apply_map: dict[str, str] = {}
    for s in samples:
        cell = s.get("target_cell")
        gold = s.get("gold_formula")
        if cell and gold:
            apply_map[cell] = gold
    wb = load_workbook(broken_source)
    for cell, gold in apply_map.items():
        sheet_name, _, cell_addr = cell.rpartition("!")
        try:
            wb[sheet_name][cell_addr] = gold
        except KeyError:
            _fail(
                f"认证成品组装失败：工作簿中找不到工作表 {sheet_name!r}"
                f"（来源样本 target_cell={cell}）"
            )
    out_path = base / "certified.xlsx"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    typer.echo(f"认证成品已导出：{out_path}（应用 {len(apply_map)} 条认证修复）")
    status, pending, contested = _refresh_workbook_index(
        base, wb_name, ds_root, cert_path, fp_path
    )
    if status == "partial":
        typer.echo(
            "警告：工作簿状态为 partial，以下格子保持原始公式（可能是坏公式）：",
            err=True,
        )
        for t in pending + contested:
            typer.echo(f"  - {t}", err=True)
    elif status == "unknown":
        typer.echo(
            "警告：workbook_index 状态未知（轮次数据不全），未终局格可能仍是原始公式",
            err=True,
        )
    else:
        typer.echo("警告：workbook_index 状态计算失败，未终局格状态未知", err=True)


@app.command()
def label(
    target: Annotated[Path, typer.Argument(help="工作簿 .xlsx 或 out/<工作簿名>/ 轮次父目录")],
    dataset_dir: Annotated[
        Path | None,
        typer.Option("--dataset-dir", help="用户数据集目录（默认 evaluation/data/user_feedback/）"),
    ] = None,
) -> None:
    """把 failed_cases 条目人工补标正确公式并晋升进 certified_cases。

    补标即 resolved：workbook_index 随即重算，该表可能解锁
    fully_certified 上传门。只动数据文件，不碰工作簿——成品由
    finalize 组装。
    """
    base = _base_dir(target)
    wb_name = base.name
    ds_root = dataset_dir
    fail_path = (
        ds_root / rs.FAILED_NAME if ds_root is not None
        else rs.dataset_dir() / rs.FAILED_NAME
    )
    cert_path = (
        ds_root / rs.CERTIFIED_NAME if ds_root is not None
        else rs.dataset_dir() / rs.CERTIFIED_NAME
    )
    fp_path = (
        ds_root / rs.FALSE_POSITIVE_NAME if ds_root is not None
        else rs.dataset_dir() / rs.FALSE_POSITIVE_NAME
    )
    try:
        all_failed = rs.load_samples(fail_path)
    except rs.FeedbackError as exc:
        _fail(f"failed_cases.json 损坏或不可读：{exc}")
    samples = [s for s in all_failed if s.get("workbook") == wb_name]
    if not samples:
        _fail(f"没有可补标的条目（failed_cases.json 中无 {wb_name} 的样本）")
    for i, s in enumerate(samples, 1):
        typer.echo(
            f"[{i}] {s.get('target_cell')} · error_type={s.get('error_type')}"
            f" · {len(s.get('trajectory') or [])} 轮轨迹"
        )
    while True:
        raw = typer.prompt("选择编号（q=退出）", default="q").strip().lower()
        if raw == "q":
            typer.echo("未做任何修改。")
            return
        if raw.isdigit() and 1 <= int(raw) <= len(samples):
            chosen = samples[int(raw) - 1]
            break
        typer.echo(f"无效编号，请输入 1-{len(samples)} 或 q")
    typer.echo(f"补标 {chosen.get('target_cell')}")
    for t in chosen.get("trajectory") or []:
        suffix = f"（{t.get('remark')}）" if t.get("remark") else ""
        typer.echo(
            f"  round {t.get('round')}: {t.get('proposal') or '—'}"
            f" → {t.get('verdict')}{suffix}"
        )
    while True:
        gold = typer.prompt("正确公式（以 = 开头）", default="").strip()
        if gold.startswith("="):
            break
        typer.echo("公式必须以 = 开头，请重新输入")
    confirm = typer.prompt(
        f"确认把 {chosen.get('target_cell')} 补标为 {gold}？[Y/n]", default="Y"
    ).strip().lower()
    if confirm not in {"y", "yes", ""}:
        typer.echo("已取消，未做任何修改。")
        return
    promoted = rs.build_promoted_sample(
        chosen, gold,
        promoted_at=datetime.now().astimezone().isoformat(timespec="seconds"),
    )
    rs.remove_sample(fail_path, chosen.get("case_id"))
    # 先删 certified 中的同 case_id 旧晋升样本再写入：重补标 = 覆盖（最新
    # 用户裁决为准），否则 append_samples 去重会静默丢弃新公式。
    if rs.remove_sample(cert_path, promoted["case_id"]):
        typer.echo(f"重新补标：已覆盖旧晋升样本 {promoted['case_id']}")
    rs.append_samples(cert_path, [promoted])
    typer.echo(f"已晋升：{chosen.get('case_id')} → {promoted['case_id']} → {cert_path}")
    # 老样本的 broken_path 可能缺失：补挂数据集副本（幂等，已有则跳过）。
    broken_source = base / "broken_source.xlsx"
    if broken_source.exists() and not promoted.get("broken_path"):
        wb_copy = rs.ensure_workbook_copy(broken_source, wb_name, ds_root)
        if wb_copy is not None:
            typer.echo(f"工作簿副本：{wb_copy}")
    status, pending, contested = _refresh_workbook_index(
        base, wb_name, ds_root, cert_path, fp_path
    )
    if status:
        bits = [f"工作簿状态：{status}"]
        if pending:
            bits.append(f"pending={pending}")
        if contested:
            bits.append(f"contested={contested}")
        typer.echo("；".join(bits))


@app.command()
def web(
    port: Annotated[int, typer.Option("--port", min=1, help="监听端口")] = 8000,
    no_open: Annotated[bool, typer.Option("--no-open", help="不自动打开浏览器")] = False,
) -> None:
    """启动本地网页入口（127.0.0.1）。需要 pip install -e ".\\[web]"。"""
    try:
        import uvicorn
    except ImportError as exc:
        _fail(f"web 依赖未安装：python -m pip install -e \".[web]\"（{exc}）")
    import threading
    import webbrowser

    url = f"http://127.0.0.1:{port}"
    if not no_open:
        threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    typer.echo(f"SheetGuard web: {url}（Ctrl+C 停止）")
    uvicorn.run("sheetguard.web.app:app", host="127.0.0.1", port=port, log_level="info")


def main() -> None:
    """控制台脚本入口。"""
    app()


if __name__ == "__main__":
    main()
