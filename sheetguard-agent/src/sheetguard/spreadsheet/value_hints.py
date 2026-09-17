"""missing_formula 候选的确定性值反推提示。

missing_formula 的目标格是数值常量，正确公式必须复现它。当同列（或
同行）家族的兄弟聚合公式（如 Agg!N 列的 MIN/MAX/AVERAGE/SUMIF/
COUNTIF）覆盖同形明细范围时，聚合函数可由"值反推"确定：把兄弟聚合
的范围平移到目标所在行/列，对引擎支持的无条件聚合函数（SUM/AVERAGE/
MIN/MAX/COUNT）逐一求值，与原常量比对。整个过程纯确定性、无 LLM，
修复提示与数据集 gold 唯一性守门共用同一份枚举，保证口径一致。
"""
from __future__ import annotations

import math

from openpyxl.utils import get_column_letter

from sheetguard.spreadsheet.aggregates import aggregate_span
from sheetguard.spreadsheet.dependency_graph import DependencyGraph
from sheetguard.spreadsheet.model import CellRef, WorkbookIndex
from sheetguard.spreadsheet.recalc import RecalcEngine

# 可枚举的无条件聚合函数（引擎支持且无需 criteria 实参）。
_HINT_FUNCTIONS = ("SUM", "AVERAGE", "MIN", "MAX", "COUNT")


def aggregate_value_hints(
    index: WorkbookIndex,
    graph: DependencyGraph,
    target_addr: str,
    rel_tol: float = 1e-6,
    abs_tol: float = 1e-3,
) -> list[dict]:
    """对常量目标枚举可平移聚合范围上的无条件聚合函数值。

    返回 ``[{"formula", "value", "matches_constant"}]``；范围取自同轴
    聚合兄弟的 ``aggregate_span`` 平移到目标行/列后的多数派形状。目标
    是公式格、非数值常量，或家族中没有聚合兄弟时返回 []（不瞎猜范围）。
    """
    info = index.cell(target_addr)
    if info is None or info.is_formula:
        return []
    if not isinstance(info.value, (int, float)) or isinstance(info.value, bool):
        return []
    constant = float(info.value)

    sheet = index._sheet_by_name(info.ref.sheet)
    if sheet is None:
        return []
    spans: dict[tuple[int, int, int, int], int] = {}
    for cell in sheet.cells.values():
        if not cell.is_formula or not cell.formula:
            continue
        if (cell.ref.row, cell.ref.col) == (info.ref.row, info.ref.col):
            continue
        span = aggregate_span(index, cell)
        if span is None:
            continue
        axis, (min_col, min_row, max_col, max_row) = span
        if axis == "row" and cell.ref.col == info.ref.col:
            # 同列的行聚合兄弟：范围平移到目标所在行。
            shifted = (min_col, info.ref.row, max_col, info.ref.row)
        elif axis == "col" and cell.ref.row == info.ref.row:
            # 同行的列聚合兄弟：范围平移到目标所在列。
            shifted = (info.ref.col, min_row, info.ref.col, max_row)
        else:
            continue
        spans[shifted] = spans.get(shifted, 0) + 1
    if not spans:
        return []
    (min_col, min_row, max_col, max_row), _count = max(
        spans.items(), key=lambda item: item[1]
    )
    range_str = (
        f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
    )

    engine = RecalcEngine(index, graph)
    hints: list[dict] = []
    for fn in _HINT_FUNCTIONS:
        formula = f"={fn}({range_str})"
        try:
            value = float(engine.evaluate_expression(formula, info.ref))
        except Exception:
            # 单个函数枚举失败（如范围形态被引擎拒绝）不阻塞其余函数。
            continue
        hints.append({
            "formula": formula,
            "value": value,
            "matches_constant": math.isclose(
                value, constant, rel_tol=rel_tol, abs_tol=abs_tol
            ),
        })
    return hints


def format_value_hints(hints: list[dict], constant: float | None) -> str:
    """把提示条目格式化为修复提示词文本（中文，与项目提示词风格一致）。

    ``hints`` 为空或无可用常量时返回"无"——明细行家族的 missing_formula
    没有可反推的聚合范围，提示词保持干净，不制造噪音。
    """
    if not hints:
        return "无"
    parts = []
    for hint in hints:
        mark = " ←与原常量一致" if hint["matches_constant"] else ""
        parts.append(f"{hint['formula']}→{hint['value']:.6g}{mark}")
    if not any(hint["matches_constant"] for hint in hints):
        parts.append(
            "以上无条件聚合都不等于原常量：优先考虑带条件的聚合"
            "（SUMIF/COUNTIF）或引用/查找型公式；"
            "不要用加减乘除组合硬凑原常量——值恰好相等的算式不是原公式"
        )
    if constant is not None:
        return "；".join(parts) + f"（原常量 {constant:.6g}）"
    return "；".join(parts)
