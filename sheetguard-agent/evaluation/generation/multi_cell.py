"""多错误工作簿数据集生成（设计 10.2）。

从 gold 工作簿生成多错误 broken 副本，每个案例的 gold 描述：
- 工作簿中的真实错误单元格集合；
- 每个错误单元格的 gold 公式；
- workbook-level 期望结果。

数据集只描述"哪些单元格确实坏了、各自的 gold 公式是什么"；
Agent 的 dismissed/unresolved/skip/deferred 等判定由 evaluators 对照
这份 gold 评估（静态检测的候选是全集，Agent 自行判断真假阳性）。
"""
from __future__ import annotations

import json
import random
from pathlib import Path

from evaluation.generation.fault_injection import (
    inject_cross_sheet_error,
    inject_missing_formula,
    inject_wrong_cell_reference,
    inject_wrong_operator,
    inject_wrong_range,
    inject_faults,
)

# 每个案例从不同注入器组合中抽取，制造 2-4 个真实错误单元格。
_CASE_INJECTOR_PLANS = [
    [inject_missing_formula, inject_wrong_operator],
    [inject_wrong_range, inject_missing_formula],
    [inject_cross_sheet_error, inject_wrong_operator],
    [inject_missing_formula, inject_wrong_cell_reference, inject_wrong_operator],
    [inject_wrong_range, inject_cross_sheet_error],
]


def build_multi_error_case(
    gold_path: str,
    broken_dir: str,
    case_id: str,
    specs: list[tuple],
) -> dict:
    """生成一个多错误案例：broken 副本 + gold 描述（不修改 gold）。"""
    broken_dir = Path(broken_dir)
    broken_dir.mkdir(parents=True, exist_ok=True)
    broken_path = broken_dir / f"{case_id}.xlsx"
    cases = inject_faults(gold_path, str(broken_path), specs)
    return {
        "case_id": case_id,
        "gold_path": str(gold_path),
        "broken_path": str(broken_path),
        "errors": [
            {
                "target_cell": case.target_cell,
                "error_type": case.error_type,
                "old_formula": case.old_formula,
                "gold_formula": case.gold_formula,
            }
            for case in cases
        ],
    }


def build_multi_error_dataset(
    gold_path: str,
    broken_dir: str,
    output_json: str,
    n_cases: int = 10,
    seed: int = 0,
) -> list[dict]:
    """批量生成多错误案例并写出数据集清单 JSON。"""
    rng = random.Random(seed)
    records: list[dict] = []
    for i in range(n_cases):
        plan = _CASE_INJECTOR_PLANS[i % len(_CASE_INJECTOR_PLANS)]
        specs = [(injector, rng.randint(0, 10_000_000)) for injector in plan]
        records.append(
            build_multi_error_case(
                gold_path, broken_dir, f"multi_{i:03d}", specs,
            )
        )
    output = Path(output_json)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(
        json.dumps(records, indent=2, ensure_ascii=False), encoding="utf-8",
    )
    return records
