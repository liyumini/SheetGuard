r"""complete 数据集的新注入器:函数过滤包装器 + 孤立汇总泄漏 + 邻居不一致。

复用 fault_injection 的五类注入器与工具;新增三个入口:

- has_function(formula, fn)      精确函数匹配(拒绝 SUMIF/COUNTIF 对
                                 IF/SUM 的假阳性,正则 (?<![A-Z0-9])FN\()。
- with_function(injector, fn)    把注入器的目标挑选限定在「gold 公式含 fn
                                 且注入后静态候选集恰为注入目标」的单元格上:
                                 对连续种子在 probe 副本上重试注入,直到命中;
                                 再用胜出种子在正式路径上重新注入(确定性:
                                 同 seed 复现)。守门是必需的——例如 VLOOKUP
                                 的种子可能落在 Revenue 行 4 的装饰行上,
                                 硬编码后静态候选集为空,不能作为载体。
- inject_singleton_sum_leak      行尾汇总 =SUM(B..M) 收缩为前 3 列,右端
                                 紧贴明细公式、范围没盖全 → 触发
                                 _signal_singleton_sum_boundary
                                 (range_boundary)。探针锚定形态:
                                 Revenue!N2 = =SUM(B2:D2) → 候选
                                 {Revenue!N2},信号含 range_boundary。
- inject_neighbor_mismatch       家族行中部格 ± 互换,左右对向邻居模板一致
                                 (各自相对自身坐标归一化)且宿主模板与之一致,
                                 换号后宿主偏离 → 触发
                                 _signal_neighbor_inconsistency
                                 (neighbor_mismatch)。探针锚定形态:
                                 家族行中部格 ± 互换 → 单候选,信号
                                 [pattern_anomaly, neighbor_mismatch]。
"""
from __future__ import annotations

import os
import random
import re

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, range_boundaries

from evaluation.generation.fault_injection import _copy_gold
from evaluation.generation.models import CaseMetadata
from sheetguard.spreadsheet.anomaly_detector import StaticInspector
from sheetguard.spreadsheet.formula_parser import to_relative_template
from sheetguard.spreadsheet.model import CellRef
from sheetguard.spreadsheet.parser import parse_workbook

# with_function 的种子重试上限。gold_v2 的 207 个公式格中,SUMIF/COUNTIF/
# AVERAGE/MIN/MAX 各只有一个载体格(Agg N2..N6,约 1/207 命中率),
# 离线探针测得其最低命中种子分布在 95..689,预算必须盖过整个挑选空间。
_MAX_SEED_TRIES = 1000


def has_function(formula: str, fn: str) -> bool:
    """精确匹配公式中的函数名(拒绝 SUMIF/COUNTIF 对 IF/SUM 的假阳性)。"""
    return re.search(rf"(?<![A-Z0-9]){fn}\(", (formula or "").upper()) is not None


def _static_gate_ok(broken_path: str, target_cell: str) -> bool:
    """守门:注入后静态候选集必须恰为注入目标,否则该种子不能作为案例。"""
    idx = parse_workbook(broken_path)
    return [c["cell"] for c in StaticInspector(idx).detect_all()] == [target_cell]


def with_function(injector, fn: str):
    """包装注入器:重试种子直到注入目标的 gold 公式含指定函数且守门通过。"""

    def wrapped(gold_path, broken_path, seed: int = 0, copy_gold: bool = True) -> CaseMetadata:
        probe = f"{broken_path}.probe.xlsx"
        try:
            winning_seed = None
            s = seed
            for _ in range(_MAX_SEED_TRIES):
                try:
                    meta = injector(gold_path, probe, seed=s, copy_gold=True)
                except ValueError:
                    s += 1
                    continue
                if not has_function(meta.gold_formula, fn):
                    s += 1
                    continue
                # 守门:静态候选集不恰为注入目标的种子(如落在无家族
                # 上下文的装饰行)同样不可用,继续换种子。
                if not _static_gate_ok(probe, meta.target_cell):
                    s += 1
                    continue
                winning_seed = s
                break
            if winning_seed is None:
                raise ValueError(
                    f"with_function({injector.__name__}, {fn}): "
                    f"{_MAX_SEED_TRIES} 个种子内找不到含 {fn} 且通过守门的注入目标"
                )
            # 用胜出种子在正式路径上重新注入(全新 copy_gold,保证 broken 干净)。
            return injector(gold_path, broken_path, seed=winning_seed, copy_gold=copy_gold)
        finally:
            try:
                os.unlink(probe)
            except FileNotFoundError:
                pass

    return wrapped


def inject_singleton_sum_leak(gold_path: str, broken_path: str, seed: int = 0, copy_gold: bool = True) -> CaseMetadata:
    """把行尾汇总收缩为前三列,右侧紧贴明细公式,触发孤立汇总边界检测。"""
    rng = random.Random(seed)
    if copy_gold:
        _copy_gold(gold_path, broken_path)
    wb = load_workbook(broken_path)

    def has_sum_range(formula: str) -> bool:
        return "SUM(" in formula.upper() and ":" in formula

    candidates = []
    for ws_name in wb.sheetnames:
        ws = wb[ws_name]
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                    if has_sum_range(cell.value):
                        candidates.append((ws_name, cell.coordinate, cell.value))
    if not candidates:
        raise ValueError("No row-end SUM cells found")
    sheet, coord, gold_formula = rng.choice(candidates)
    match = re.search(r"SUM\(([^)]+)\)", gold_formula)
    if not match or ":" not in match.group(1):
        raise ValueError(f"Could not inject singleton leak at {sheet}!{coord}")
    min_col, min_row, max_col, max_row = range_boundaries(match.group(1))
    if max_col - min_col + 1 <= 3:
        raise ValueError("Range too narrow to shrink")
    new_range = (
        f"{get_column_letter(min_col)}{min_row}:{get_column_letter(min_col + 2)}{max_row}"
    )
    old_formula = gold_formula.replace(match.group(1), new_range)
    wb[sheet][coord] = old_formula
    wb.save(broken_path)
    return CaseMetadata(
        case_id=f"singleton_sum_{seed}",
        error_type="wrong_range",
        target_cell=f"{sheet}!{coord}",
        old_formula=old_formula,
        gold_formula=gold_formula,
        gold_path=gold_path,
        broken_path=broken_path,
    )


def inject_neighbor_mismatch(gold_path: str, broken_path: str, seed: int = 0, copy_gold: bool = True) -> CaseMetadata:
    """换家族行中部格的加减号,左右邻居模板一致 → 触发邻居不一致检测。"""
    rng = random.Random(seed)
    if copy_gold:
        _copy_gold(gold_path, broken_path)
    wb = load_workbook(broken_path)

    candidates = []
    for ws_name in wb.sheetnames:
        ws = wb[ws_name]
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not (v and isinstance(v, str) and v.startswith("=")):
                    continue
                if cell.column <= 1:
                    continue
                left = ws.cell(cell.row, cell.column - 1).value
                right = ws.cell(cell.row, cell.column + 1).value
                if not (
                    isinstance(left, str) and left.startswith("=")
                    and isinstance(right, str) and right.startswith("=")
                ):
                    continue
                try:
                    # 各模板都相对该单元格自身的坐标归一化:同形家族的
                    # 左右邻居与宿主才能得到同一模板串。
                    t_left = to_relative_template(
                        left, CellRef(sheet=ws_name, row=cell.row, col=cell.column - 1)
                    )
                    t_right = to_relative_template(
                        right, CellRef(sheet=ws_name, row=cell.row, col=cell.column + 1)
                    )
                    t_host = to_relative_template(
                        v, CellRef(sheet=ws_name, row=cell.row, col=cell.column)
                    )
                except (IndexError, ValueError, TypeError):
                    continue
                if t_left != t_right or t_host != t_left:
                    continue
                # +/- 位置必须落在 SUM(...) 之外(span 判定,精确排除)。
                sum_spans = [m.span() for m in re.finditer(r"SUM\([^)]*\)", v)]
                positions = [
                    (m.start(), m.group())
                    for m in re.finditer(r"[+-]", v)
                    if not any(start <= m.start() < end for start, end in sum_spans)
                ]
                if not positions:
                    continue
                candidates.append((ws_name, cell.coordinate, v, positions))
    if not candidates:
        raise ValueError("No neighbor-homogeneous cells found")
    sheet, coord, gold_formula, positions = rng.choice(candidates)
    pos, op = rng.choice(positions)
    replacement = "-" if op == "+" else "+"
    old_formula = gold_formula[:pos] + replacement + gold_formula[pos + 1:]
    wb[sheet][coord] = old_formula
    wb.save(broken_path)
    return CaseMetadata(
        case_id=f"neighbor_mismatch_{seed}",
        error_type="wrong_operator",
        target_cell=f"{sheet}!{coord}",
        old_formula=old_formula,
        gold_formula=gold_formula,
        gold_path=gold_path,
        broken_path=broken_path,
    )
