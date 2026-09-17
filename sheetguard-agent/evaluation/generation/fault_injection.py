"""从正确 gold 工作簿复制并制造五类公式错误。

这些函数用于准备 LangSmith 的 broken 案例：原始 gold 文件始终不被修改，
每次注入都会先复制，再只改动副本中的一个目标单元格。
"""
from __future__ import annotations
import random
import shutil
import json
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from sheetguard.spreadsheet.parser import parse_workbook
from sheetguard.spreadsheet.dependency_graph import DependencyGraph
from sheetguard.spreadsheet.recalc import RecalcEngine
from sheetguard.spreadsheet.formula_parser import extract_refs
from evaluation.generation.models import CaseMetadata


def _pick_formula_cell(wb, rng: random.Random, predicate=None) -> tuple[str, str, str]:
    """从工作簿中随机挑选符合条件的公式单元格。

    返回值依次为工作表名、单元格坐标和原始正确公式；predicate 可限制
    可选公式，例如只选择包含 SUM 范围的公式。
    """
    candidates = []
    for ws_name in wb.sheetnames:
        ws = wb[ws_name]
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                    if predicate is None or predicate(cell.value):
                        candidates.append((ws_name, cell.coordinate, cell.value))
    if not candidates:
        raise ValueError("No matching formula cells found")
    return rng.choice(candidates)


def _copy_gold(gold_path: str, broken_path: str) -> None:
    """先完整复制正确工作簿，确保错误只出现在 broken 副本。"""
    shutil.copy2(gold_path, broken_path)


def inject_wrong_range(gold_path: str, broken_path: str, seed: int = 0, copy_gold: bool = True) -> CaseMetadata:
    """制造“SUM 少算一个单元格”的范围错误，并返回案例说明。"""
    rng = random.Random(seed)
    if copy_gold:
        _copy_gold(gold_path, broken_path)
    wb = load_workbook(broken_path)

    def has_sum_range(formula: str) -> bool:
        return "SUM(" in formula.upper() and ":" in formula

    sheet, coord, gold_formula = _pick_formula_cell(wb, rng, has_sum_range)
    import re
    range_match = re.search(r'SUM\(([^)]+)\)', gold_formula)
    if range_match:
        range_str = range_match.group(1)
        if ":" in range_str:
            from openpyxl.utils import range_boundaries
            try:
                min_col, min_row, max_col, max_row = range_boundaries(range_str)
                if max_col > min_col:
                    new_max_col = max_col - 1
                    new_max_row = max_row
                elif max_row > min_row:
                    new_max_row = max_row - 1
                    new_max_col = max_col
                else:
                    raise ValueError("Cannot shrink single-cell range")
                new_range = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(new_max_col)}{new_max_row}"
                old_formula = gold_formula.replace(range_str, new_range)
                wb[sheet][coord] = old_formula
                wb.save(broken_path)
                return CaseMetadata(
                    case_id=f"wrong_range_{seed}",
                    error_type="wrong_range",
                    target_cell=f"{sheet}!{coord}",
                    old_formula=old_formula,
                    gold_formula=gold_formula,
                    gold_path=gold_path,
                    broken_path=broken_path,
                )
            except Exception:
                pass
    raise ValueError(f"Could not inject wrong_range in {sheet}!{coord}")


def inject_wrong_cell_reference(gold_path: str, broken_path: str, seed: int = 0, copy_gold: bool = True) -> CaseMetadata:
    """制造“引用了相邻错误单元格”的错误，并返回案例说明。"""
    rng = random.Random(seed)
    if copy_gold:
        _copy_gold(gold_path, broken_path)
    wb = load_workbook(broken_path)

    def has_cell_ref(formula: str) -> bool:
        import re
        return bool(re.search(r'[A-Z]+\d+', formula))

    sheet, coord, gold_formula = _pick_formula_cell(wb, rng, has_cell_ref)
    import re
    refs = re.findall(r'[A-Z]+\d+', gold_formula)
    target_ref = rng.choice(refs)
    from openpyxl.utils import coordinate_to_tuple
    row, col = coordinate_to_tuple(target_ref)
    if rng.random() < 0.5:
        col += 1 if col < 26 else -1
    else:
        row += 1 if row < 100 else -1
    new_ref = f"{get_column_letter(col)}{row}"
    old_formula = gold_formula.replace(target_ref, new_ref, 1)
    wb[sheet][coord] = old_formula
    wb.save(broken_path)
    return CaseMetadata(
        case_id=f"wrong_cell_ref_{seed}",
        error_type="wrong_cell_reference",
        target_cell=f"{sheet}!{coord}",
        old_formula=old_formula,
        gold_formula=gold_formula,
        gold_path=gold_path,
        broken_path=broken_path,
    )


def inject_wrong_operator(gold_path: str, broken_path: str, seed: int = 0, copy_gold: bool = True) -> CaseMetadata:
    """制造加减号互换的运算符错误，并返回案例说明。"""
    rng = random.Random(seed)
    if copy_gold:
        _copy_gold(gold_path, broken_path)
    wb = load_workbook(broken_path)

    def has_op(formula: str) -> bool:
        # 检查 SUM() 之外的运算符
        import re
        stripped = re.sub(r'SUM\([^)]+\)', '', formula)
        return '+' in stripped or '-' in stripped

    sheet, coord, gold_formula = _pick_formula_cell(wb, rng, has_op)
    import re
    # 在 SUM(...) 之外查找 + 或 -
    plus_positions = [m.start() for m in re.finditer(r'\+', gold_formula)]
    minus_positions = [m.start() for m in re.finditer(r'(?<!SUM\()-(?!\d)', gold_formula)]
    candidates = []
    for pos in plus_positions:
        candidates.append(('+', pos))
    for pos in minus_positions:
        candidates.append(('-', pos))
    if candidates:
        op, pos = rng.choice(candidates)
        replacement = '-' if op == '+' else '+'
        old_formula = gold_formula[:pos] + replacement + gold_formula[pos + 1:]
        wb[sheet][coord] = old_formula
        wb.save(broken_path)
        return CaseMetadata(
            case_id=f"wrong_operator_{seed}",
            error_type="wrong_operator",
            target_cell=f"{sheet}!{coord}",
            old_formula=old_formula,
            gold_formula=gold_formula,
            gold_path=gold_path,
            broken_path=broken_path,
        )
    raise ValueError(f"Could not inject wrong_operator in {sheet}!{coord}")


def inject_missing_formula(gold_path: str, broken_path: str, seed: int = 0, copy_gold: bool = True) -> CaseMetadata:
    """用计算结果替换公式，模拟公式被硬编码数值覆盖的情况。"""
    rng = random.Random(seed)
    if copy_gold:
        _copy_gold(gold_path, broken_path)
    wb = load_workbook(broken_path)

    idx = parse_workbook(gold_path)
    graph = DependencyGraph(idx)
    graph.build()
    engine = RecalcEngine(idx, graph)

    sheet, coord, gold_formula = _pick_formula_cell(wb, rng)
    target = f"{sheet}!{coord}"
    try:
        value = engine.evaluate_cell(target)
    except Exception:
        value = 0.0

    wb[sheet][coord] = value
    wb.save(broken_path)
    return CaseMetadata(
        case_id=f"missing_formula_{seed}",
        error_type="missing_formula",
        target_cell=target,
        old_formula=str(value),
        gold_formula=gold_formula,
        gold_path=gold_path,
        broken_path=broken_path,
    )


def inject_cross_sheet_error(gold_path: str, broken_path: str, seed: int = 0, copy_gold: bool = True) -> CaseMetadata:
    """把跨工作表引用改到错误工作表，模拟跨表引用错误。"""
    rng = random.Random(seed)
    if copy_gold:
        _copy_gold(gold_path, broken_path)
    wb = load_workbook(broken_path)

    def has_cross_sheet(formula: str) -> bool:
        return "!" in formula

    sheet, coord, gold_formula = _pick_formula_cell(wb, rng, has_cross_sheet)
    sheet_refs = [ref_sheet for ref_sheet, _ in extract_refs(gold_formula) if ref_sheet]
    all_sheets = set(wb.sheetnames)
    for ref_sheet in sheet_refs:
        other_sheets = all_sheets - {ref_sheet}
        if other_sheets:
            wrong_sheet = rng.choice(sorted(other_sheets))
            old_ref = f"'{ref_sheet}'!" if f"'{ref_sheet}'!" in gold_formula else f"{ref_sheet}!"
            new_ref = f"'{wrong_sheet}'!" if not wrong_sheet.replace("_", "").isalnum() else f"{wrong_sheet}!"
            old_formula = gold_formula.replace(old_ref, new_ref, 1)
            wb[sheet][coord] = old_formula
            wb.save(broken_path)
            return CaseMetadata(
                case_id=f"cross_sheet_error_{seed}",
                error_type="cross_sheet_error",
                target_cell=f"{sheet}!{coord}",
                old_formula=old_formula,
                gold_formula=gold_formula,
                gold_path=gold_path,
                broken_path=broken_path,
            )
    raise ValueError(f"Could not inject cross_sheet_error in {sheet}!{coord}")


# 生成案例时按此列表轮换错误类型，避免数据集只包含一种错误。
INJECTORS = [
    inject_wrong_range,
    inject_wrong_cell_reference,
    inject_wrong_operator,
    inject_missing_formula,
    inject_cross_sheet_error,
]


def generate_cases(
    gold_path: str,
    broken_dir: str,
    cases_json_path: str,
    n_cases: int = 100,
    seed: int = 0,
    max_attempts_per_case: int = 10,
) -> list[CaseMetadata]:
    """批量生成 broken 工作簿和对应的 cases.json 案例索引。

    每个案例记录正确公式与错误公式，供 LangSmith 上传数据集时作为标准答案。
    当某种错误不适用于随机选中的公式时，会在限制次数内重新选择。
    """
    if n_cases < 0:
        raise ValueError("n_cases must be non-negative")
    if max_attempts_per_case < 1:
        raise ValueError("max_attempts_per_case must be at least 1")

    rng = random.Random(seed)
    broken_dir = Path(broken_dir)
    broken_dir.mkdir(parents=True, exist_ok=True)

    cases: list[CaseMetadata] = []
    for i in range(n_cases):
        # 轮换选择注入器，让各错误类型尽量均匀地出现在数据集中。
        injector = INJECTORS[i % len(INJECTORS)]
        broken_path = broken_dir / f"broken_{i:03d}.xlsx"
        last_error = None
        for _attempt in range(max_attempts_per_case):
            case_seed = rng.randint(0, 10_000_000)
            try:
                meta = injector(gold_path, str(broken_path), seed=case_seed)
                meta.case_id = f"case_{i:03d}"
                meta.gold_path = gold_path
                meta.broken_path = str(broken_path)
                cases.append(meta)
                break
            except ValueError as exc:
                last_error = exc
                if broken_path.exists():
                    broken_path.unlink()
        else:
            raise RuntimeError(
                f"Could not generate case {i} with {injector.__name__} "
                f"after {max_attempts_per_case} attempts: {last_error}"
            ) from last_error

    # cases.json 是 LangSmith 上传脚本读取的案例清单。
    with open(cases_json_path, "w") as f:
        json.dump([c.to_dict() for c in cases], f, indent=2)

    return cases

def inject_faults(
    gold_path: str,
    broken_path: str,
    specs: list[tuple],
    max_attempts_per_fault: int = 20,
) -> list[CaseMetadata]:
    """在同一 broken 副本上依次注入多个错误，制造 multi-cell 案例。

    ``specs`` 为 (injector, seed) 列表；每个注入器只改动一个单元格，
    目标单元格互不相同时注入成功。原始 gold 文件始终不被修改。
    """
    if not specs:
        raise ValueError("specs must contain at least one (injector, seed) pair")
    _copy_gold(gold_path, broken_path)
    cases: list[CaseMetadata] = []
    for injector, seed in specs:
        last_error = None
        for _attempt in range(max_attempts_per_fault):
            try:
                meta = injector(gold_path, broken_path, seed=seed, copy_gold=False)
            except ValueError as exc:
                last_error = exc
                continue
            if any(case.target_cell == meta.target_cell for case in cases):
                last_error = ValueError(
                    f"{injector.__name__} picked already broken cell {meta.target_cell}"
                )
                continue
            cases.append(meta)
            break
        else:
            raise RuntimeError(
                f"Could not inject {injector.__name__} (seed={seed}) after "
                f"{max_attempts_per_fault} attempts: {last_error}"
            )
    return cases
