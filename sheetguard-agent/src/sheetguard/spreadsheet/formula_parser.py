"""公式引用提取与模板归一化。

本模块提供两个核心能力：
1. 提取并展开单元格/范围引用，用于构建依赖图（dependency_graph.py）
2. 将公式归一化为相对偏移模板，用于公式模式异常检测(当前文件的to_relative_template方法)

示例工作流：
假设公式 `=SUM('P&L'!B2:B5) + C4` 位于 `P&L!C6`（当前单元格行=6, 列=3）：

1. 用于依赖图：extract_refs() → expand_range()
   - extract_refs() → [("P&L", "B2:B5"), (None, "C4")]
   - expand_range("P&L", "B2:B5") → [("P&L", "B2"), ("P&L", "B3"), ..., ("P&L", "B5")]
   → 结果得到该公式依赖的所有单元格

2. 用于模式检测：to_relative_template()
   - 将引用替换为相对于当前单元格(6, 3)的偏移量：
   - 'P&L'!B2:B5 → S(P&L)RR(-4, -1, -1, -1)
   - C4 → R(-2, 0)
   - 最终模板：`=SUM(S(P&L)RR(-4,-1,-1,-1)) + R(-2,0)`
   → 同行/同列相邻公式模板相同则为一致，模板不匹配则标记为异常候选。
"""
from __future__ import annotations
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils import range_boundaries, get_column_letter
from sheetguard.spreadsheet.model import CellRef


def _split_sheet_ref(raw: str) -> tuple[str | None, str]:
    """拆分引用字符串，分离表名和单元格/范围部分。

    例子:
        "Revenue!B2"     → ("Revenue", "B2")
        "'P&L'!B2:B20"   → ("P&L", "B2:B20")
        "B2"             → (None, "B2")
    """
    if "!" in raw:
        idx = raw.rindex("!")  # 取最后一个 !，兼容表名本身包含 ! 的情况
        sheet = raw[:idx]
        if sheet.startswith("'") and sheet.endswith("'"):
            sheet = sheet[1:-1]  # 去掉引号，例如 "'P&L' → P&L"
        return (sheet, raw[idx + 1:])
    return (None, raw)


def _sanitize_range_str(s: str) -> str:
    """去除范围字符串中的绝对引用符号 $，统一格式。

    例子:
        "$B$2:$B$20" → "B2:B20"
    """
    return s.replace("$", "")


def extract_refs(formula: str) -> list[tuple[str | None, str]]:
    """从公式文本中提取所有单元格/范围引用。

    返回 (sheet, range_string) 元组列表：
        - 对于当前表内引用，sheet 为 None
        - 对于跨表引用，sheet 为目标表名

    例子:
        "=A1+Sheet1!B2" → [(None, "A1"), ("Sheet1", "B2")]
    """
    if not formula.startswith("="):
        return []
    refs: list[tuple[str | None, str]] = []
    seen = set()
    # 使用 openpyxl 分词器切分公式，只取出范围/单元格引用 token
    for tok in Tokenizer(formula).items:
        if tok.type == "OPERAND" and tok.subtype == "RANGE":
            sheet, rest = _split_sheet_ref(tok.value)
            rest = _sanitize_range_str(rest)
            key = (sheet, rest)
            if key not in seen:
                seen.add(key)
                refs.append(key)
    return refs


def expand_range(sheet: str, range_str: str) -> list[tuple[str, str]]:
    """将范围字符串展开为其中每个单元格的 (sheet, 地址) 元组列表。

    例子:
        "Revenue", "B2:B4" → [("Revenue", "B2"), ("Revenue", "B3"), ("Revenue", "B4")]
    """
    clean = _sanitize_range_str(range_str)
    try:
        min_col, min_row, max_col, max_row = range_boundaries(clean)
    except Exception:
        # 范围解析失败，尝试按单个单元格解析
        from openpyxl.utils import coordinate_to_tuple
        try:
            row, col = coordinate_to_tuple(clean)
            return [(sheet, clean)]
        except Exception:
            return []
    cells = []
    # 遍历范围中每一行每一列，生成所有单元格
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            addr = f"{get_column_letter(c)}{r}"
            cells.append((sheet, addr))
    return cells


def to_relative_template(formula: str, host: CellRef) -> str:
    """将公式归一化为规范模板字符串：把所有单元格引用替换为相对于当前单元格的偏移量。

    相同结构的公式会生成完全相同的模板，用于公式模式异常检测和验证：
    如果相邻公式属于同一个家族，它们的模板应该相同；不同则表示异常。

    例子:
        当前单元格 host: B4 (row=4, col=2)
        公式: =SUM(A2:A4) + C4
        结果: =SUM(RR(-2,-1,0,-1)) + R(0,+1)

        如果引用跨表: Revenue!B2 → S(Revenue)R(-2,-1)
    """
    if not formula.startswith("="):
        return formula
    parts: list[str] = []
    for tok in Tokenizer(formula).items:
        if tok.type == "OPERAND" and tok.subtype == "RANGE":
            sheet, rest = _split_sheet_ref(tok.value)
            rest = _sanitize_range_str(rest)
            try:
                rng = range_boundaries(rest)
            except Exception:
                parts.append(tok.value)
                continue
            min_col, min_row, max_col, max_row = rng
            # 计算相对于当前单元格 host 的行列偏移
            r1_off = min_row - host.row
            c1_off = min_col - host.col
            r2_off = max_row - host.row
            c2_off = max_col - host.col
            # 跨表引用加前缀 S(表名)，本地引用不加
            prefix = f"S({sheet})" if sheet else ""
            if (min_col, min_row) == (max_col, max_row):
                # 单个单元格 → R(行偏移, 列偏移)
                parts.append(f"{prefix}R({r1_off},{c1_off})")
            else:
                # 范围 → RR(min行偏移, min列偏移, max行偏移, max列偏移)
                parts.append(f"{prefix}RR({r1_off},{c1_off},{r2_off},{c2_off})")
        elif tok.value == "SUM(":
            parts.append("SUM(")
        elif tok.type == "FUNC":
            parts.append(tok.value)
        else:
            # 运算符、括号等直接保留
            parts.append(tok.value)
    return "".join(parts)
