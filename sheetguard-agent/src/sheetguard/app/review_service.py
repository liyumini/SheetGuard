"""审查流程服务层：CLI review 与 web 审查端点共用。

交互式提问（_prompt_verdict、Y/n 确认）留在 cli.review；本模块只做
可程序化的状态读取、草稿存取、提交落盘与下一轮重修组装。
禁止 import typer（web 层复用）。
"""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Callable

from sheetguard.app import review_store as rs
from sheetguard.app.rounds import (
    load_audit,
    write_round_artifacts,
)


class ReviewStateError(ValueError):
    """审查状态不可用（无待审轮次 / audit 不可读 / 无可审条目 / 参数非法）。"""


def collect_history(base: Path, upto_round: int) -> list[dict]:
    """历史轮缓存：audit 与 feedback 一次性读入（损坏容错为 None，不阻塞构造）。"""
    history: list[dict] = []
    for d in rs.round_dirs(base):
        no = int(d.name.split("-")[1])
        if no >= upto_round:
            break
        try:
            round_audit = json.loads(
                (d / "audit.json").read_text(encoding="utf-8")
            )
        except (OSError, ValueError):
            round_audit = None
        try:
            round_fb = rs.load_feedback(d / rs.FEEDBACK_NAME)
        except rs.FeedbackError:
            round_fb = None
        history.append({"round": no, "audit": round_audit, "feedback": round_fb})
    return history


def build_review_items(
    audit: dict, history: list[dict]
) -> tuple[list[dict], list[dict], list[dict], list[dict]]:
    """构造四类审查集：本轮 fixed/failed + 历史跳过未决格 + dismissed 条目。

    history 每项 {"round": int, "audit": dict | None, "feedback": dict | None}
    （历史轮缓存，audit/feedback 损坏为 None，不阻塞构造）。

    历史跳过格去重规则（v1.11）：每个 target 取历轮 feedback 中最新一条
    判定——最新 verdict 仍是 skipped 才以 ③ 类重现（src_round 取最新轮）；
    已 correct/false_positive 的不重现（已终局），incorrect 的也不重现
    （它会走正常 incorrect 路由，或已在本轮 audit 中以 ①/② 呈现）。
    dismissed 的 skipped verdict=默认通过，不重现（v1.12）。
    """
    items = [
        e for e in (audit.get("fixed") or [])
        if isinstance(e, dict) and e.get("target")
    ]
    failed_items = [
        e for e in (audit.get("failed") or [])
        if isinstance(e, dict) and e.get("target")
    ]
    dismissed_items = [
        e for e in (audit.get("dismissed") or [])
        if isinstance(e, dict) and e.get("target")
    ]
    seen_targets = {e["target"] for e in items} | {e["target"] for e in failed_items}
    # 升序遍历 + 同 target 覆盖 → 每 target 留最新一轮的判定记录。
    latest: dict[str, tuple[int, dict]] = {}
    for record in history:
        src_round = record.get("round")
        feedback = record.get("feedback") or {}
        for r in feedback.get("reviews") or []:
            if isinstance(r, dict) and r.get("target"):
                latest[r["target"]] = (src_round, r)
    skipped_items = [
        dict(r, src_round=src_round)
        for target, (src_round, r) in latest.items()
        if r.get("verdict") == "skipped"
        and r.get("kind", "fixed") != "dismissed"
        and target not in seen_targets
    ]
    return items, failed_items, skipped_items, dismissed_items


def load_review_state(base: Path) -> dict:
    """读轮次 + audit + 历史 + 草稿，产出审查所需全部状态。

    base = out/<工作簿名>/ 目录（调用方先归一）。不可用时抛
    ReviewStateError（文案与原 CLI _fail 消息一致）。
    """
    base = Path(base)
    rnd = rs.latest_reviewable_round(base)
    if rnd is None:
        raise ReviewStateError(
            "没有待审查的轮次（全部已提交，或还没有运行过 repair/run，"
            "或最新一轮的反馈文件已存在）"
        )
    round_no = int(rnd.name.split("-")[1])
    try:
        audit = load_audit(rnd / "audit.json")
    except ValueError as exc:
        raise ReviewStateError(str(exc)) from exc
    history = collect_history(base, round_no)
    items, failed_items, skipped_items, dismissed_items = build_review_items(audit, history)
    if not items and not failed_items and not skipped_items and not dismissed_items:
        raise ReviewStateError(
            f"round-{round_no} 没有可审查的条目"
            "（fixed/failed 均空，历史无跳过格，无 dismissed）"
        )

    # 草稿续审：已判条目直接采用，不再询问。
    draft = rs.load_draft(rnd / rs.DRAFT_NAME) or {}
    verdicts: dict[str, tuple[str, str | None]] = {
        item["target"]: (item["verdict"], item.get("remark"))
        for item in draft.get("reviews") or []
        if isinstance(item, dict) and item.get("verdict") in rs.VERDICTS
    }
    proposal_of = {e["target"]: e.get("new_formula") for e in items}
    proposal_of.update({e["target"]: e.get("last_formula") for e in failed_items})
    proposal_of.update({e["target"]: e.get("proposal") for e in skipped_items})
    review_items: list[tuple[str, dict]] = (
        [("fixed", e) for e in items]
        + [("failed", e) for e in failed_items]
        + [("skipped", e) for e in skipped_items]
        + [("dismissed", e) for e in dismissed_items]
    )
    kind_of = {e["target"]: kind for kind, e in review_items}
    return {
        "base": base,
        "round": round_no,
        "round_dir": rnd,
        "audit": audit,
        "history": history,
        "items": items,
        "failed": failed_items,
        "skipped": skipped_items,
        "dismissed": dismissed_items,
        "verdicts": verdicts,
        "proposal_of": proposal_of,
        "kind_of": kind_of,
        "review_items": review_items,
    }


def save_review_draft(state: dict, verdicts: dict[str, tuple[str, str | None]]) -> None:
    """把当前已判条目写入轮次草稿（CLI q/n 分支与 web 实时存草稿共用格式）。"""
    rnd: Path = state["round_dir"]
    rs.save_draft(rnd / rs.DRAFT_NAME, {
        "round": state["round"],
        "reviews": [
            {"target": t, "proposal": state["proposal_of"].get(t),
             "verdict": v, "remark": m,
             "kind": state["kind_of"].get(t, "fixed")}
            for t, (v, m) in verdicts.items()
        ],
    })


def save_draft_entry(base: Path, target: str, verdict: str, remark: str | None) -> None:
    """单条裁决实时存草稿（web 用）：读现有草稿，同 target 覆盖。"""
    if verdict not in rs.VERDICTS:
        raise ReviewStateError(f"invalid verdict: {verdict}")
    state = load_review_state(base)
    draft_path = state["round_dir"] / rs.DRAFT_NAME
    draft = rs.load_draft(draft_path) or {"round": state["round"], "reviews": []}
    reviews = [r for r in (draft.get("reviews") or [])
               if isinstance(r, dict) and r.get("target") != target]
    reviews.append({
        "target": target,
        "proposal": state["proposal_of"].get(target),
        "verdict": verdict,
        "remark": remark,
        "kind": state["kind_of"].get(target, "fixed"),
    })
    rs.save_draft(draft_path, {"round": state["round"], "reviews": reviews})


def _error_type_of(history: list[dict], kind: str, entry: dict) -> dict | None:
    """取条目的 hypothesis dict：skipped 形态从历史轮 audit 补，其余直读。"""
    if kind == "skipped":
        audit_entry = _find_audit_entry(history, entry["target"], entry.get("src_round"))
        error_type = ((audit_entry or {}).get("hypothesis") or {}).get("error_type")
        return {"error_type": error_type} if error_type else {}
    return entry.get("hypothesis") or {}


def _find_audit_entry(
    history: list[dict], target: str, src_round: int | None = None
) -> dict | None:
    """从历史轮 audit 的 fixed/failed 列表中找 target 条目。

    传 src_round 时优先精确匹配该轮（提案所在轮的 audit）；找不到
    再回退最早优先（容错：audit 缺失或轮号对不上时不阻塞取值）。
    """
    fallback = None
    for record in history:
        audit = record.get("audit") or {}
        for section in ("fixed", "failed"):
            for e in audit.get(section) or []:
                if isinstance(e, dict) and e.get("target") == target:
                    if src_round is not None and record.get("round") == src_round:
                        return e
                    if fallback is None:
                        fallback = e
    return fallback


def _find_dismissed_reason(
    history: list[dict], target: str, dismissed_round: int | None
) -> str | None:
    """从争议发生轮的 audit dismissed 列表取 reason；audit 缺失记 None。"""
    for record in history:
        if dismissed_round is not None and record.get("round") != dismissed_round:
            continue
        for e in (record.get("audit") or {}).get("dismissed") or []:
            if isinstance(e, dict) and e.get("target") == target:
                return e.get("reason")
    return None


def refresh_workbook_index(
    base: Path,
    wb_name: str,
    dataset_dir: Path | None,
    cert_path: Path,
    fp_path: Path,
    warnings: list[str],
) -> tuple[str | None, list[str], list[str]]:
    """重算并记录 workbook_index 条目；失败降级：警告文案 append 进 warnings。

    数据源 = 该表历轮 audit+feedback 与 certified/false_positive 数据集；
    任一轮不可读 → unknown（保守不上传）。位于提交管线中 feedback 落盘
    之后、重修触发之前：计算失败不卡半提交状态。
    """
    try:
        rounds_data: list[dict] = []
        rounds_readable = True
        for d in rs.round_dirs(base):
            try:
                round_audit = json.loads(
                    (d / "audit.json").read_text(encoding="utf-8")
                )
                round_fb = rs.load_feedback(d / rs.FEEDBACK_NAME)
            except (OSError, ValueError):
                # JSONDecodeError/FeedbackError/UnicodeDecodeError ⊂ ValueError。
                rounds_readable = False
                break
            rounds_data.append({
                "round": int(d.name.split("-")[1]),
                "fixed": [
                    e.get("target")
                    for e in (round_audit.get("fixed") or []) if e.get("target")
                ],
                "failed": [
                    e.get("target")
                    for e in (round_audit.get("failed") or []) if e.get("target")
                ],
                "rejected": [
                    r.get("target")
                    for r in (round_fb.get("reviews") or [])
                    if r.get("verdict") == "incorrect" and r.get("target")
                ],
            })
        # resolved 集 = certified ∪ false_positive（两类都是终局）。
        all_samples = rs.load_samples(cert_path)
        wb_cert = [s for s in all_samples if s.get("workbook") == wb_name]
        fp_samples_all = rs.load_samples(fp_path)
        wb_fp = [s for s in fp_samples_all if s.get("workbook") == wb_name]
        status, pending, contested = rs.decide_workbook_status(
            rounds_data,
            {s.get("target_cell") for s in wb_cert}
            | {s.get("target_cell") for s in wb_fp},
            rounds_readable,
        )
        rs.record_workbook_entry(
            None if dataset_dir is None else dataset_dir,
            wb_name,
            status=status,
            pending_targets=pending,
            contested_targets=contested,
            sample_case_ids=[
                s["case_id"] for s in wb_cert if s.get("case_id")
            ],
            false_positive_case_ids=[
                s["case_id"] for s in wb_fp if s.get("case_id")
            ],
        )
        return status, pending, contested
    except (OSError, ValueError, TypeError) as exc:
        # FeedbackError ⊂ ValueError（certified_cases.json 损坏）；
        # decide/record 内部 sorted() 遇非字符串条目抛 TypeError。
        warnings.append(
            f"警告：workbook_index 状态计算失败，本轮保持未知（{exc}）"
        )
        return None, [], []


def submit_review(
    state: dict,
    verdicts: dict[str, tuple[str, str | None]],
    cfg,
    dataset_dir: Path | None = None,
) -> dict:
    """提交一轮审查裁决：feedback 落盘 → 样本沉淀 → index 刷新 → 重修上下文。

    verdicts = {target: (verdict, remark | None)}；未出现在 verdicts 的
    审查条目按 skipped 落盘（未判等价跳过，轨迹完整）。
    返回 dict：{round_no, wb_name, counts, logs, warnings, incorrect_count,
    review_context, exhausted}。review_context 非 None ⇒ 调用方应触发下一轮。
    """
    base: Path = state["base"]
    rnd: Path = state["round_dir"]
    audit: dict = state["audit"]
    history: list[dict] = state["history"]
    review_items: list[tuple[str, dict]] = state["review_items"]
    proposal_of = state["proposal_of"]
    round_no: int = state["round"]
    wb_name = base.name
    logs: list[str] = []
    warnings: list[str] = []

    counts = {"correct": 0, "incorrect": 0, "false_positive": 0, "skipped": 0}
    for verdict, _ in verdicts.values():
        counts[verdict] += 1

    # 提交：feedback.json（未判条目等价 skipped，一并落盘，保证轨迹完整）。
    reviews_payload = [
        {
            "target": e["target"],
            "old_formula": e.get("old_formula"),
            "proposal": proposal_of.get(e["target"]),
            "verdict": verdicts.get(e["target"], ("skipped", None))[0],
            "remark": verdicts.get(e["target"], (None, None))[1],
            "kind": kind,
        }
        for kind, e in review_items
    ]
    submitted_at = datetime.now().astimezone().isoformat(timespec="seconds")
    agent_ver = rs.agent_version()
    rs.save_feedback(
        rnd / rs.FEEDBACK_NAME,
        {
            "workbook": audit.get("workbook"),
            "round": round_no,
            "submitted_at": submitted_at,
            "agent_version": agent_ver,
            "reviews": reviews_payload,
        },
    )
    (rnd / rs.DRAFT_NAME).unlink(missing_ok=True)
    logs.append(f"反馈已提交：{rnd / rs.FEEDBACK_NAME}")

    # 显式传 dataset_dir 时目录即数据集根；否则用包内默认 evaluation/data/user_feedback/。
    if dataset_dir is not None:
        cert_path = dataset_dir / rs.CERTIFIED_NAME
        fail_path = dataset_dir / rs.FAILED_NAME
        fp_path = dataset_dir / rs.FALSE_POSITIVE_NAME
        missed_path = dataset_dir / rs.MISSED_NAME
    else:
        cert_path = rs.dataset_dir() / rs.CERTIFIED_NAME
        fail_path = rs.dataset_dir() / rs.FAILED_NAME
        fp_path = rs.dataset_dir() / rs.FALSE_POSITIVE_NAME
        missed_path = rs.dataset_dir() / rs.MISSED_NAME

    feedback = {
        "round": round_no,
        "submitted_at": submitted_at,
        "agent_version": agent_ver,
        "reviews": reviews_payload,
    }
    # 按轮号过滤而非切片：轮次目录可能不连续，切片会漏掉编号间隙后的历史反馈。
    prior_feedbacks = [
        rs.load_feedback(d / rs.FEEDBACK_NAME)
        for d in rs.round_dirs(base)
        if d.name.split("-")[1].isdigit()
        and int(d.name.split("-")[1]) < round_no
        and (d / rs.FEEDBACK_NAME).exists()
    ]
    broken_source = base / "broken_source.xlsx"
    # 数据集自包含：把 broken_source 副本按工作簿落进数据集 workbooks/。
    wb_copy = rs.ensure_workbook_copy(
        broken_source, wb_name, None if dataset_dir is None else dataset_dir
    )
    copy_path = str(wb_copy) if wb_copy is not None else None

    # —— 认证样本 ——
    cert_samples = []
    for kind, entry in review_items:
        verdict, _ = verdicts.get(entry["target"], ("skipped", None))
        if verdict != "correct":
            continue
        item = {"target": entry["target"], "old_formula": entry.get("old_formula"),
                "proposal": proposal_of.get(entry["target"]),
                "verdict": "correct",
                "remark": verdicts.get(entry["target"], (None, None))[1],
                "kind": kind}
        trajectory = rs.build_trajectory(
            prior_feedbacks + [feedback], entry["target"]
        )
        if kind == "skipped":
            cert_entry = {
                "target": entry["target"],
                "old_formula": entry.get("old_formula"),
                "hypothesis": _error_type_of(history, "skipped", entry),
            }
            cert_samples.append(rs.build_certified_sample(
                wb_name, cert_entry, item, round_no, copy_path, trajectory,
                certified_at=submitted_at, agent_version=agent_ver,
                config=cfg.to_dict(),
            ))
        else:
            cert_samples.append(rs.build_certified_sample(
                wb_name, entry, item, round_no, copy_path, trajectory,
                certified_at=submitted_at, agent_version=agent_ver,
                config=cfg.to_dict(),
            ))
    if cert_samples:
        rs.append_samples(cert_path, cert_samples)
        logs.append(f"认证样本 {len(cert_samples)} 条 → {cert_path}")

    # —— 误修样本 ——
    fp_samples = []
    for kind, entry in review_items:
        verdict, _ = verdicts.get(entry["target"], ("skipped", None))
        if verdict != "false_positive":
            continue
        item = {"target": entry["target"], "old_formula": entry.get("old_formula"),
                "proposal": proposal_of.get(entry["target"]),
                "verdict": "false_positive",
                "remark": verdicts.get(entry["target"], (None, None))[1],
                "kind": kind}
        trajectory = rs.build_trajectory(
            prior_feedbacks + [feedback], entry["target"]
        )
        error_type = (_error_type_of(history, kind, entry) or {}).get("error_type")
        # skipped 条目的历史 feedback 可能没有 old_formula 字段（如
        # proposal=null 落盘后又判 3 的最小 fixture）：fp 样本的
        # old_formula/gold_formula 回退到提案公式——该格现值即提案。
        fp_old_formula = item.get("old_formula") or (
            item.get("proposal") if kind == "skipped" else None
        )
        fp_samples.append(rs.build_false_positive_sample(
            wb_name,
            target=entry["target"],
            old_formula=fp_old_formula,
            error_type=error_type,
            round_no=round_no,
            broken_source=copy_path,
            trajectory=trajectory,
            decided_at=submitted_at, agent_version=agent_ver,
            config=cfg.to_dict(),
        ))
    if fp_samples:
        rs.append_samples(fp_path, fp_samples)
        logs.append(f"误修样本 {len(fp_samples)} 条 → {fp_path}")

    # —— 漏检样本 ——
    missed_samples = []
    for kind, entry in review_items:
        verdict, _ = verdicts.get(entry["target"], ("skipped", None))
        if verdict != "correct":
            continue
        dispute_round = rs.find_dismissal_dispute(prior_feedbacks, entry["target"])
        if dispute_round is None:
            continue
        item = {"target": entry["target"], "old_formula": entry.get("old_formula"),
                "proposal": proposal_of.get(entry["target"]),
                "verdict": "correct",
                "remark": verdicts.get(entry["target"], (None, None))[1],
                "kind": kind}
        trajectory = rs.build_trajectory(
            prior_feedbacks + [feedback], entry["target"]
        )
        # 形态归一：skipped 形态 entry（历史 feedback item）无 hypothesis 键，
        # error_type 从历史轮 audit 补齐，与同格 cert 样本保持一致。
        miss_entry = dict(entry, hypothesis=_error_type_of(history, kind, entry))
        missed_samples.append(rs.build_missed_sample(
            wb_name, miss_entry, item, round_no, copy_path, trajectory,
            dismissed_round=dispute_round,
            dismissed_reason=_find_dismissed_reason(
                history, entry["target"], dispute_round
            ),
            certified_at=submitted_at, agent_version=agent_ver,
            config=cfg.to_dict(),
        ))
    if missed_samples:
        rs.append_samples(missed_path, missed_samples)
        logs.append(f"漏检样本 {len(missed_samples)} 条 → {missed_path}")

    # —— 全认证门状态计算（失败降级警告，不卡半提交）——
    refresh_workbook_index(base, wb_name, dataset_dir, cert_path, fp_path, warnings)

    # —— 错项处理：超限 → failed；否则组装 review_context ——
    incorrect = [
        (kind, entry) for kind, entry in review_items
        if verdicts.get(entry["target"], ("skipped", None))[0] == "incorrect"
    ]
    if not incorrect:
        return {
            "round_no": round_no, "wb_name": wb_name, "counts": counts,
            "logs": logs, "warnings": warnings,
            "incorrect_count": 0, "review_context": None, "exhausted": False,
        }

    if round_no + 1 > cfg.max_review_rounds:
        failed_samples = []
        for kind, entry in incorrect:
            trajectory = rs.build_trajectory(
                prior_feedbacks + [feedback], entry["target"]
            )
            failed_samples.append(rs.build_failed_sample(
                wb_name,
                (
                    # dismissed 条目形状是 {target, reason, formula}：
                    # old_formula 回退 formula（该格从未被修改，当前公式即
                    # 原公式）；error_type 无源可查（无 hypothesis），保持 None。
                    dict(entry, old_formula=entry.get("old_formula") or entry.get("formula"))
                    if kind == "dismissed" else entry
                ),
                trajectory, cfg.max_review_rounds,
                copy_path, recorded_at=submitted_at,
                agent_version=agent_ver, config=cfg.to_dict(),
            ))
        rs.append_samples(fail_path, failed_samples)
        logs.append(
            f"已达轮次上限 {cfg.max_review_rounds}，"
            f"{len(failed_samples)} 条记为 rejected_exhausted → {fail_path}"
        )
        return {
            "round_no": round_no, "wb_name": wb_name, "counts": counts,
            "logs": logs, "warnings": warnings,
            "incorrect_count": len(incorrect), "review_context": None,
            "exhausted": True,
        }

    # 认证格跨轮累积（I1）：上一轮 audit 透传的 certified 条目合并进来，
    # 同格冲突时当前轮优先；旧条目保留其原有 round 键（若 audit 有）。
    # 遍历 review_items（而非本轮 fixed）：历史跳过格在本轮判对同样入列。
    current_certified = [
        {"target": t, "formula": proposal_of.get(t), "round": round_no}
        for t, (v, _) in verdicts.items()
        if v == "correct" and t in {e["target"] for _, e in review_items}
    ]
    current_targets = {c["target"] for c in current_certified}
    prior_certified = [
        c for c in (audit.get("certified") or [])
        if isinstance(c, dict) and c.get("target") not in current_targets
    ]
    # 排除集按 kind 取公式源：fixed=new_formula；failed=历轮 repair_history
    # 该格全部尝试 ∪ last_formula；skipped=历史提案本身。
    rejected_context: dict[str, dict] = {}
    for kind, entry in incorrect:
        target = entry["target"]
        if kind == "dismissed":
            # 被推翻的 dismissed：无提案需排除；remark 组合 agent 原因与用户
            # 质疑，注入重查提示词（设计 2.3）。
            remark_bits = [f"上轮 dismissed 原因：{entry.get('reason') or '—'}"]
            user_remark = verdicts.get(target, (None, None))[1]
            if user_remark:
                remark_bits.append(f"用户坚持此格有错：{user_remark}")
            rejected_context[target] = {
                "rejected_formulas": [],
                "remark": "；".join(remark_bits),
            }
            continue
        formulas: list[str] = []
        if kind == "fixed":
            formulas.append(entry.get("new_formula") or "")
        elif kind == "failed":
            # 历轮 + 当前轮 repair_history 中该格的全部尝试 + last_formula。
            # 被否 failed 条目的失败尝试发生在当前轮（当前轮 audit 的
            # repair_history），只扫 history 会漏掉当轮尝试（如首轮已判
            # 非法的公式），下一轮 agent 可能重试。
            audit_sources = [h.get("audit") or {} for h in history] + [audit]
            for src in audit_sources:
                for h in src.get("repair_history") or []:
                    if isinstance(h, dict) and h.get("target") == target and h.get("formula"):
                        formulas.append(h["formula"])
            formulas.append(entry.get("last_formula") or "")
        else:  # skipped
            formulas.append(entry.get("proposal") or "")
        rejected_context[target] = {
            "rejected_formulas": [f for f in formulas if f],
            "remark": verdicts.get(target, (None, None))[1],
        }
    review_context = {
        "certified": prior_certified + current_certified,
        "rejected": rejected_context,
    }
    return {
        "round_no": round_no, "wb_name": wb_name, "counts": counts,
        "logs": logs, "warnings": warnings,
        "incorrect_count": len(incorrect),
        "review_context": review_context, "exhausted": False,
    }


def run_next_round(
    base: Path,
    round_no: int,
    review_context: dict,
    repair_fn: Callable,
    progress_callback=None,
) -> dict:
    """以上一轮 repaired.xlsx 为输入跑第 N+1 轮修复并落盘轮次产物。

    repair_fn(workbook_path, repaired_workbook_path, review_context,
    progress_callback) -> agent result dict（含 "audit" 键）。
    CLI 传入闭包读取 cli 模块全局（SheetGuardGraph/build_model 可被
    测试 monkeypatch）；web 传入携带 progress_callback 的工厂。
    """
    rnd = base / f"round-{round_no}"
    prev_repaired = rnd / "repaired.xlsx"
    next_input = prev_repaired if prev_repaired.exists() else base / "broken_source.xlsx"
    if not next_input.exists():
        raise ReviewStateError(f"找不到下一轮输入工作簿：{next_input}")
    next_dir = base / f"round-{round_no + 1}"
    result = repair_fn(
        str(next_input), next_dir / "repaired.xlsx", review_context, progress_callback
    )
    next_audit = {"workbook": str(next_input), **result.get("audit", {})}
    next_dir.mkdir(parents=True, exist_ok=True)
    write_round_artifacts(next_dir, next_audit)
    return {
        "round": round_no + 1,
        "dir": next_dir,
        "status": result.get("status", "unknown"),
        "audit": next_audit,
    }
