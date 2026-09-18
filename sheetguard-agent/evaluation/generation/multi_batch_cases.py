"""multi-batch 数据集:30 个多错误单元格工作簿(每表 3~6 个错误)。

错误模板池取自 coverage 矩阵的 33 个已验证注入(coverage_cases),按错误
类型轮转采样、单元格不冲突注入;每个 case 生成期走与 coverage 相同的
双守门:静态候选集必须与注入目标精确一致 + gold 必须通过 Verifier 六项
验证。错误级记录附 carrier_function 元数据,与 coverage/multi-cell 的
item 结构兼容(errors[].target_cell/gold_formula + expected_status)。

固定种子 → 同参数重生成结果逐字节一致;守门不过自动换种子重采样。
"""
from __future__ import annotations

import json
import random
import shutil
from pathlib import Path

from openpyxl import load_workbook

from evaluation.generation.complete_cases import (
    _assert_gold_repairable,
    _hardcode,
)
from sheetguard.spreadsheet.anomaly_detector import StaticInspector
from sheetguard.spreadsheet.parser import parse_workbook

# 错误模板池:与 coverage 矩阵一致;(sheet, coord, broken_formula, etype, fn)。
# missing_formula 的 broken 公式由 RecalcEngine 求值写入(None 占位)。
FAULT_POOL: dict[str, list[tuple[str, str, str | None, str]]] = {
    "wrong_range": [
        ("Revenue", "N3", "=SUM(B3:L3)", "SUM"),
        ("Agg", "N4", "=AVERAGE(B4:L4)", "AVERAGE"),
        ("Agg", "N2", "=MIN(B2:L2)", "MIN"),
        ("Agg", "N3", "=MAX(B3:L3)", "MAX"),
        ("Agg", "N7", "=COUNT(B7:L7)", "COUNT"),
        ("Lk", "B7", "=VLOOKUP(C7, C7:C7, 2, FALSE)", "VLOOKUP"),
        ("Agg", "N5", '=SUMIF(B5:L5, ">=700", B5:L5)', "SUMIF"),
        ("Agg", "N6", '=COUNTIF(B6:L6, ">1000")', "COUNTIF"),
    ],
    "wrong_cell_reference": [
        ("Revenue", "D3", "=IF(RawData!E2>1000, Revenue!D2*0.9, Revenue!D2)", "IF"),
        ("Lk", "B7", "=VLOOKUP(C8, C7:D7, 2, FALSE)", "VLOOKUP"),
        ("Agg", "N5", '=SUMIF(B5:M5, ">=700", C5:M5)', "SUMIF"),
        ("Revenue", "E2", "=RawData!F2*RawData!E3", "Arith"),
    ],
    "wrong_operator": [
        ("P&L", "C2", "=Revenue!C2+Cost!C2", "Arith"),
        ("Revenue", "D3", "=IF(RawData!D2<1000, Revenue!D2*0.9, Revenue!D2)", "IF"),
    ],
    "missing_formula": [
        ("Revenue", "N3", None, "SUM"),
        ("Agg", "N2", None, "MIN"),
        ("Agg", "N3", None, "MAX"),
        ("Agg", "N4", None, "AVERAGE"),
        ("Agg", "N7", None, "COUNT"),
        ("Revenue", "D3", None, "IF"),
        ("Lk", "B7", None, "VLOOKUP"),
        ("Agg", "N5", None, "SUMIF"),
        ("Agg", "N6", None, "COUNTIF"),
        ("P&L", "C2", None, "Arith"),
    ],
    "cross_sheet_error": [
        ("Revenue", "N3", "=SUM(Cost!B3:M3)", "SUM"),
        ("Agg", "N4", "=AVERAGE(Revenue!B4:M4)", "AVERAGE"),
        ("Agg", "N2", "=MIN(Cost!B2:M2)", "MIN"),
        ("Agg", "N3", "=MAX(Revenue!B3:M3)", "MAX"),
        ("Revenue", "D3", "=IF(Dashboard!D2>1000, Revenue!D2*0.9, Revenue!D2)", "IF"),
        ("Lk", "F7", "=VLOOKUP(C7, RawData!A7:B7, 2, FALSE)", "VLOOKUP"),
        ("Agg", "N5", '=SUMIF(Revenue!B5:M5, ">=700", Revenue!B5:M5)', "SUMIF"),
        ("Agg", "N6", '=COUNTIF(Revenue!B6:M6, ">1000")', "COUNTIF"),
        ("Cost", "E2", "=Revenue!E2*Revenue!E4+Revenue!E5", "Arith"),
    ],
}
_ETYPE_ROTATION = list(FAULT_POOL.keys())
DEFAULT_SEED = 20260917


def build_multi_batch_cases(
    gold_path: str,
    broken_dir: str,
    output_json: str,
    n_cases: int = 30,
    seed: int = DEFAULT_SEED,
    max_attempts: int = 50,
) -> list[dict]:
    """构造 n_cases 个多错误工作簿并写出清单 JSON;每 case 先过双守门。"""
    gold_path = Path(gold_path)
    broken_dir = Path(broken_dir)
    broken_dir.mkdir(parents=True, exist_ok=True)
    gold_wb = load_workbook(gold_path)

    def gold_formula(sheet: str, coord: str) -> str:
        return str(gold_wb[sheet][coord].value)

    rng = random.Random(seed)
    records: list[dict] = []
    for i in range(n_cases):
        case_id = f"mbatch_{i:02d}"
        broken_path = broken_dir / f"{case_id}.xlsx"
        for _attempt in range(max_attempts):
            # 错误数量 3~6,类型按轮转洗牌采样,保证单表内类型多样。
            k = rng.choice([3, 4, 5, 6])
            order = _ETYPE_ROTATION[:]
            rng.shuffle(order)
            picked: list[tuple[str, str, str | None, str, str]] = []
            occupied: set[str] = set()
            for slot in range(k):
                etype = order[slot % len(order)]
                pool = [
                    t for t in FAULT_POOL[etype]
                    if f"{t[0]}!{t[1]}" not in occupied
                ]
                if not pool:
                    continue
                sheet, coord, broken, fn = rng.choice(pool)
                occupied.add(f"{sheet}!{coord}")
                picked.append((sheet, coord, broken, etype, fn))
            if len(picked) < 3:
                continue  # 模板冲突过多,换种子重采样

            shutil.copy2(gold_path, broken_path)
            wb = load_workbook(broken_path)
            errors = []
            try:
                for sheet, coord, broken, etype, fn in picked:
                    gold = gold_formula(sheet, coord)
                    if etype == "missing_formula":
                        # 同 hard/complete 注入:写 gold 的重算正确值。
                        value = _hardcode(str(gold_path), f"{sheet}!{coord}")
                        old = str(value)
                        wb[sheet][coord] = value
                    else:
                        old = broken
                        wb[sheet][coord] = broken
                    errors.append({
                        "target_cell": f"{sheet}!{coord}",
                        "error_type": etype,
                        "old_formula": old,
                        "gold_formula": gold,
                        "carrier_function": fn,
                    })
                wb.save(broken_path)
            except Exception:
                continue

            idx = parse_workbook(str(broken_path))
            got = {c["cell"] for c in StaticInspector(idx).detect_all()}
            targets = {e["target_cell"] for e in errors}
            if got != targets:
                continue
            try:
                _assert_gold_repairable(broken_path, errors)
            except AssertionError:
                continue
            records.append({
                "case_id": case_id,
                "gold_path": str(gold_path),
                "broken_path": str(broken_path),
                "errors": errors,
                "expected_status": "success",
                "static_candidate_count": len(errors),
            })
            break
        else:
            raise RuntimeError(
                f"{case_id}: {max_attempts} 次采样内守门不通过"
            )

    output = Path(output_json)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(records, indent=2, ensure_ascii=False), encoding="utf-8")
    return records


if __name__ == "__main__":
    root = Path(__file__).resolve().parents[2]
    cases = build_multi_batch_cases(
        str(root / "evaluation/data/coverage/gold.xlsx"),
        str(root / "evaluation/data/multi_batch/broken"),
        str(root / "evaluation/data/multi_batch_cases.json"),
    )
    total = sum(len(c["errors"]) for c in cases)
    from collections import Counter

    et_counter = Counter(e["error_type"] for c in cases for e in c["errors"])
    print(f"multi-batch cases: {len(cases)} total_errors={total} by_type={dict(et_counter)}")
