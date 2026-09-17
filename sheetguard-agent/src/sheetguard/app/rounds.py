"""轮次产物目录与读写工具：run/repair/review/finalize/web 共用。

自 cli.py 机械搬出（cli 保留同名 re-export）；此处禁止 import typer，
保证可被 web 层无副作用地复用。
"""
from __future__ import annotations

import json
import shutil
from pathlib import Path

from sheetguard.app.report import render_report

# 批处理结果中视为成功的最终状态；其余状态（failed 等）返回非零退出码。
BATCH_SUCCESS_STATUSES = {
    "success", "partial_success", "completed_without_repairs", "no_candidates",
}


class AuditReadError(ValueError):
    """审计 JSON 缺失或损坏。"""


def round_base(workbook: Path) -> Path:
    """一轮运行的全部产物（audit/report/repaired）所在的父目录。"""
    return Path("out") / workbook.stem


def round_dir(workbook: Path) -> Path:
    """out/<工作簿名>/round-N：N 取已有最大轮次 +1；首轮留档 broken_source。"""
    base = round_base(workbook)
    existing = [
        int(d.name.split("-")[1])
        for d in base.glob("round-*")
        if d.is_dir() and d.name.split("-")[1].isdigit()
    ]
    target = base / f"round-{max(existing, default=0) + 1}"
    source_copy = base / "broken_source.xlsx"
    if target.name == "round-1" and not source_copy.exists():
        source_copy.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(workbook, source_copy)
    return target


def write_json(path: Path, payload: dict | list) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


def write_round_artifacts(artifacts_dir: Path, audit: dict) -> None:
    """审计 JSON 与 Markdown 报告写入轮次目录（调用方先建目录）。"""
    write_json(artifacts_dir / "audit.json", audit)
    (artifacts_dir / "report.md").write_text(
        render_report(audit), encoding="utf-8"
    )


def load_audit(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise AuditReadError(f"cannot read audit file {path}: {exc}") from exc


def base_dir(arg: Path) -> Path:
    """review/finalize 入参归一：.xlsx/.xlsm 视为工作簿 → out/<stem>/；目录直接作为 base。

    目录不要求已存在：不存在的目录交给调用方各自的缺失检查给出更贴切的
    错误（如 finalize 的 broken_source.xlsx 不存在；review 无轮次可审）。
    """
    if arg.suffix.lower() in {".xlsx", ".xlsm"}:
        return round_base(arg)
    return arg
