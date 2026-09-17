"""生成用于 LangSmith 测评的正确（gold）财务工作簿。

此文件不创建错误案例；它只负责制作一份公式完全正确的模板，
随后由 fault_injection.py 复制该模板并人为制造错误。
"""
import random
import calendar
from openpyxl import Workbook
from openpyxl.utils import get_column_letter


def _month_col(m: int, col_start: int = 2) -> int:
    """月份编号（从 1 开始）→ 列索引。"""
    return col_start + m - 1


def _month_col_letter(m: int, col_start: int = 2) -> str:
    """把月份编号转换成 Excel 列字母，例如 1 对应 B。"""
    return get_column_letter(_month_col(m, col_start))


def _total_col(n_months: int, col_start: int = 2) -> int:
    """返回 Total 汇总列的列号；它位于最后一个月份列之后。"""
    return col_start + n_months


def _total_col_letter(n_months: int, col_start: int = 2) -> str:
    """返回 Total 汇总列的 Excel 列字母。"""
    return get_column_letter(_total_col(n_months, col_start))


def build_gold_model(seed: int = 0, n_months: int = 12) -> Workbook:
    """生成确定性的合成财务工作簿。

    工作表包括：RawData、Revenue、Cost、P&L、Dashboard。
    每个工作表都有月份列（B..M）和一个 Total 汇总列。
    """
    # 使用独立且固定种子的随机数生成器，保证同一个 seed 总能得到相同工作簿。
    rng = random.Random(seed)
    wb = Workbook()
    wb.remove(wb.active)

    months = list(range(1, n_months + 1))
    month_names = [calendar.month_abbr[m] for m in months]
    CS = 2  # column start (B)

    # ── RawData：原始业务输入，不含公式。──
    raw = wb.create_sheet("RawData")
    raw.cell(1, 1, "Label")
    for m in months:
        raw.cell(1, _month_col(m, CS), month_names[m - 1])

    raw.cell(2, 1, "Units")
    for m in months:
        raw.cell(2, _month_col(m, CS), rng.randint(800, 1200))

    raw.cell(3, 1, "UnitPrice")
    for m in months:
        raw.cell(3, _month_col(m, CS), round(rng.uniform(45, 65), 2))

    raw.cell(4, 1, "UnitCost")
    for m in months:
        raw.cell(4, _month_col(m, CS), round(rng.uniform(25, 40), 2))

    raw.cell(5, 1, "Opex")
    for m in months:
        raw.cell(5, _month_col(m, CS), rng.randint(5000, 12000))

    # ── Revenue：用销量 × 单价计算每月收入，并在最后一列汇总。──
    rev = wb.create_sheet("Revenue")
    rev.cell(1, 1, "Label")
    for m in months:
        rev.cell(1, _month_col(m, CS), month_names[m - 1])
    rev.cell(1, _total_col(n_months, CS), "Total")

    rev.cell(2, 1, "GrossRevenue")
    for m in months:
        col = _month_col_letter(m, CS)
        rev[f"{col}2"] = f"=RawData!{col}2*RawData!{col}3"
    tc = _total_col_letter(n_months, CS)
    rev[f"{tc}2"] = f"=SUM(B2:{get_column_letter(_total_col(n_months, CS) - 1)}2)"

    # ── Cost：用销量 × 单位成本 + 运营费用计算每月成本。──
    cost = wb.create_sheet("Cost")
    cost.cell(1, 1, "Label")
    for m in months:
        cost.cell(1, _month_col(m, CS), month_names[m - 1])
    cost.cell(1, _total_col(n_months, CS), "Total")

    cost.cell(2, 1, "TotalCost")
    for m in months:
        col = _month_col_letter(m, CS)
        cost[f"{col}2"] = f"=RawData!{col}2*RawData!{col}4+RawData!{col}5"
    cost[f"{tc}2"] = f"=SUM(B2:{get_column_letter(_total_col(n_months, CS) - 1)}2)"

    # ── P&L：计算毛利、税费和净利润。──
    pl = wb.create_sheet("P&L")
    pl.cell(1, 1, "Label")
    for m in months:
        pl.cell(1, _month_col(m, CS), month_names[m - 1])
    pl.cell(1, _total_col(n_months, CS), "Total")

    pl.cell(2, 1, "GrossProfit")
    for m in months:
        col = _month_col_letter(m, CS)
        pl[f"{col}2"] = f"=Revenue!{col}2-Cost!{col}2"

    pl.cell(3, 1, "Tax")
    for m in months:
        col = _month_col_letter(m, CS)
        pl[f"{col}3"] = f"='P&L'!{col}2*0.25"

    pl.cell(4, 1, "NetProfit")
    for m in months:
        col = _month_col_letter(m, CS)
        pl[f"{col}4"] = f"='P&L'!{col}2-'P&L'!{col}3"

    for r in (2, 3, 4):
        pl[f"{tc}{r}"] = f"=SUM(B{r}:{get_column_letter(_total_col(n_months, CS) - 1)}{r})"

    # ── Dashboard：引用 P&L、Revenue、Cost 中的结果，提供跨表引用案例。──
    dash = wb.create_sheet("Dashboard")
    dash.cell(1, 1, "Label")
    for m in months:
        dash.cell(1, _month_col(m, CS), month_names[m - 1])
    dash.cell(1, _total_col(n_months, CS), "Total")

    dash.cell(2, 1, "NetProfit")
    for m in months:
        col = _month_col_letter(m, CS)
        dash[f"{col}2"] = f"='P&L'!{col}4"

    dash.cell(3, 1, "Revenue")
    for m in months:
        col = _month_col_letter(m, CS)
        dash[f"{col}3"] = f"=Revenue!{col}2"

    dash.cell(4, 1, "Cost")
    for m in months:
        col = _month_col_letter(m, CS)
        dash[f"{col}4"] = f"=Cost!{col}2"

    for r in (2, 3, 4):
        dash[f"{tc}{r}"] = f"=SUM(B{r}:{get_column_letter(_total_col(n_months, CS) - 1)}{r})"

    # 此时工作簿仍只存在于内存；调用方决定保存到 gold.xlsx 还是临时测试文件。
    return wb