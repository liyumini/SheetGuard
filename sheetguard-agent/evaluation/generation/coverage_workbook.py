"""gold_v3:coverage 数据集(函数×错误类型覆盖矩阵)的载体工作簿。

在 build_gold_v2_model(seed=0) 上追加(原有行列一律不动,保持 complete-v1
的 gold 与探针边界不被扰动):

- Agg 行 7:COUNT 载体行——明细 ='Revenue'!{col}3(IF 家族结果,数值),
  行尾 =COUNT(B7:M7);补齐 SUPPORTED_FUNCTIONS 里唯一没有载体的 COUNT;
- Lk F 列:跨表 VLOOKUP 载体——=VLOOKUP(C{r}, Rate!A{r}:B{r}, 2, FALSE),
  键 C{r}=r-1 与 Rate!A{r}=r-1 对齐,查找范围随行平移使纵向模板一致;
  Rate!B{r} 系数与 Lk!D{r} 不同,F 列是真正依赖 Rate 数据的跨表查找家族,
  覆盖 Lk B 本地家族给不了的 cross_sheet_error 对。

生成守门(与 complete_workbook 同口径):全表重算零异常 + 干净静态候选集为空。
"""
from __future__ import annotations

from openpyxl import Workbook

from evaluation.generation.complete_workbook import (
    build_gold_v2_model,
    clean_candidates,
    recalc_failures,
)

# 与 build_gold_model / complete_workbook 的月份布局一致:CS=2(B),Total 列 = N
_COLS = ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]


def build_gold_v3_model(seed: int = 0) -> Workbook:
    """在 gold_v2 上追加 COUNT 与跨表 VLOOKUP 载体;原有行列零改动。"""
    wb = build_gold_v2_model(seed=seed)

    # ── Agg 行 7:COUNT 载体(明细引用 Revenue 行 3 的 IF 家族结果)──
    agg = wb["Agg"]
    agg["A7"] = "Line6"
    for col in _COLS:
        agg[f"{col}7"] = f"='Revenue'!{col}3"
    agg["N7"] = "=COUNT(B7:M7)"

    # ── Lk F 列:跨表 VLOOKUP 家族(查找表随行平移,纵向模板一致)──
    lk = wb["Lk"]
    for r in range(2, 14):
        lk[f"F{r}"] = f"=VLOOKUP(C{r}, Rate!A{r}:B{r}, 2, FALSE)"

    return wb


if __name__ == "__main__":
    import sys
    from pathlib import Path

    root = Path(__file__).resolve().parents[2]
    target = root / "evaluation/data/coverage/gold.xlsx"
    target.parent.mkdir(parents=True, exist_ok=True)
    wb = build_gold_v3_model(seed=0)
    assert recalc_failures(wb) == [], recalc_failures(wb)
    assert clean_candidates(wb) == [], clean_candidates(wb)
    wb.save(target)
    print(f"gold_v3 saved: {target}")
