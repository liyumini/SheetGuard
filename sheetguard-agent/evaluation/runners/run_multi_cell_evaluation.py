"""运行 multi-cell Langfuse 实验，检查 SheetGuard 的批量修复结果（设计 10/11）。"""
from __future__ import annotations

import io
import json
import sys
import threading
import time
from pathlib import Path
import argparse

from dotenv import load_dotenv
from langfuse import Evaluation, get_client
from rich.console import Console
from rich.progress import (
    BarColumn,
    Progress,
    SpinnerColumn,
    TaskProgressColumn,
    TextColumn,
    TimeElapsedColumn,
)

from sheetguard.agent.graph import SheetGuardGraph
from sheetguard.agent.llm_config import build_model
from sheetguard.config import load_config_from_env
from evaluation.runners.create_dataset import DATASET_NAME
from evaluation.runners.multi_cell_evaluators import (
    evaluate_attempt_cost,
    evaluate_confirmation_recall,
    evaluate_deferred_counts,
    evaluate_dependency_skip_accuracy,
    evaluate_dismissal_accuracy,
    evaluate_expected_deferred,
    evaluate_expected_dismissal,
    evaluate_expected_skip,
    evaluate_false_repair,
    evaluate_formula_accuracy,
    evaluate_formula_equivalent,
    evaluate_outcome_counts,
    evaluate_repair_precision,
    evaluate_repair_recall,
    evaluate_static_candidate_counts,
    evaluate_workbook_status,
)
from evaluation.runners.score_integrity import verify_experiment_records, verify_item_results

_console = Console(
    file=io.TextIOWrapper(
        sys.stdout.buffer if hasattr(sys.stdout, "buffer") else sys.stdout,
        encoding="utf-8",
        errors="replace",
        line_buffering=True,
    )
)
_progress = Progress(
    SpinnerColumn(),
    TextColumn("[progress.description]{task.description}"),
    BarColumn(),
    TaskProgressColumn(),
    TextColumn("("),
    TimeElapsedColumn(),
    TextColumn(")"),
    console=_console,
)
_progress_lock = threading.Lock()
_progress_task_id: int | None = None

load_dotenv()

# multi-cell 数据集在 Langfuse 中的名称。
MULTI_CELL_DATASET_NAME = "sheetguard-repair-multi-cell-v1"

# 修复副本的导出目录（与 CWD 无关，已被 .gitignore 排除）。
ROOT = Path(__file__).resolve().parents[2]
REPAIRED_DIR = ROOT / "evaluation/data/repaired"


def target(*, item, **kwargs) -> dict:
    """对一条 multi-cell 数据集 item 运行一次批处理工作流。

    返回值直接作为评测器的 ``output``（集合级评测器契约：audit dict
    层级，见 ``multi_cell_evaluators.py``）；``case_id`` 用于在 Langfuse
    输出中定位案例。修复副本统一导出到 ``evaluation/data/repaired/``：
    相对路径会落在进程 CWD（曾把运行产物写进仓库根并随提交入库），
    绝对路径让产物位置与运行目录无关，且该目录已被 .gitignore 排除。
    """
    inputs = item.input
    broken_path = inputs["broken_path"]
    repaired_path = inputs.get("repaired_workbook") or str(
        REPAIRED_DIR / f"{inputs['case_id']}.xlsx"
    )

    # 行为参数跟 sheetguard.config 的默认值走（评测测的是"当前项目"；
    # 实验换配置用 SHEETGUARD_CONFIG 环境变量，指纹随审计 JSON 记录）。
    agent = SheetGuardGraph(build_model(), config=load_config_from_env())
    try:
        result = agent.invoke(broken_path, repaired_workbook=repaired_path)
        audit = result.get("audit", {})
        return {**audit, "case_id": inputs["case_id"]}
    finally:
        agent.cleanup()
        with _progress_lock:
            if _progress_task_id is not None:
                _progress.update(_progress_task_id, advance=1)


EVALUATORS = [
    evaluate_static_candidate_counts,
    evaluate_confirmation_recall,
    evaluate_dismissal_accuracy,
    evaluate_repair_recall,
    evaluate_repair_precision,
    evaluate_false_repair,
    evaluate_formula_accuracy,
    evaluate_formula_equivalent,
    evaluate_dependency_skip_accuracy,
    evaluate_deferred_counts,
    evaluate_outcome_counts,
    evaluate_workbook_status,
    evaluate_attempt_cost,
    evaluate_expected_dismissal,
    evaluate_expected_skip,
    evaluate_expected_deferred,
]


def _get_dataset_with_retry(langfuse, name: str, *, attempts: int = 5, delay: float = 2.0):
    """带线性退避重试的数据集获取。

    到 Langfuse 云端（us.cloud.langfuse.com）的 TLS 连接在当前网络下会被
    间歇掐断（SSL EOF）；SDK 的 API 客户端对连接类错误不自动重试，而
    启动期数据集获取是只读幂等调用，重试即可恢复，避免一次网络抖动
    中断整个评测。重试耗尽后抛出最后一次的原始异常。
    """
    last_error: Exception | None = None
    for attempt in range(attempts):
        try:
            return langfuse.get_dataset(name)
        except Exception as exc:  # httpx.ConnectError / Timeout / 协议错误等
            last_error = exc
            if attempt + 1 < attempts:
                wait = delay * (attempt + 1)
                _console.print(
                    f"[yellow]get_dataset 网络错误（{attempt + 1}/{attempts}）："
                    f"{type(exc).__name__}，{wait:.0f}s 后重试[/yellow]"
                )
                time.sleep(wait)
    assert last_error is not None
    raise last_error


def main():
    """启动 multi-cell Langfuse 评测实验。

    ``--dataset`` 选择数据集（默认 sheetguard-repair-multi-cell-v1；
    难题数据集传 sheetguard-repair-multi-cell-hard-v1），评测器与
    完整性校验共用同一套。
    """
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--dataset",
        default=MULTI_CELL_DATASET_NAME,
        help="Langfuse 数据集名称（默认 multi-cell 常规集）",
    )
    args = parser.parse_args()

    langfuse = get_client()
    dataset = _get_dataset_with_retry(langfuse, args.dataset)
    total_items = len(dataset.items)

    global _progress_task_id
    _progress_task_id = _progress.add_task(
        description=f"正在评测 {args.dataset}（{total_items} 个案例）",
        total=total_items,
    )

    _progress.start()
    try:
        results = dataset.run_experiment(
            name=args.dataset,
            description=(
                "SheetGuard multi-cell batch repair evaluation: candidate / "
                "confirmation / repair / false-repair / dependency-blocker / "
                "deferred / workbook-level status"
            ),
            task=target,
            evaluators=EVALUATORS,
            max_concurrency=1,
        )
    finally:
        _progress.stop()
        langfuse.flush()

    # 服务端评分对账 + 缺失回填。results.format() 用的是内存评测结果，
    # 服务端丢分在此不可见（见 score_integrity.py 模块注释）。先打印
    # 结果再对账：回填重试耗尽时 verify_item_results 显式抛错终止评测，
    # 但结果仍留在终端可见。
    _console.print(results.format())
    _console.print("[dim]正在对账服务端评分完整性……[/dim]")
    verify_item_results(langfuse, results.item_results, console=_console)
    _console.print("[green]评分完整性校验通过。[/green]")
    # 服务端偶发把一次运行拆成多条同名实验记录（见 verify_experiment_records）。
    verify_experiment_records(
        langfuse,
        run_name=results.run_name,
        expected_items=total_items,
        console=_console,
    )


if __name__ == "__main__":
    main()
