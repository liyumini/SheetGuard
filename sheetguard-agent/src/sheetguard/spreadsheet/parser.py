"""基于 openpyxl 的工作簿解析器：生成 WorkbookIndex。"""
from __future__ import annotations
from openpyxl import load_workbook
from sheetguard.spreadsheet.model import CellRef, CellInfo, SheetInfo, WorkbookIndex


def parse_workbook(path: str, read_values: bool = False) -> WorkbookIndex:
    """解析 .xlsx 工作簿，将其转换为结构化的 WorkbookIndex 内存索引。

    Args:
        path: .xlsx 文件的路径
        read_values: 如果为 True，则加载 Excel 文件中缓存的公式计算结果
                     （使用 data_only=True 参数）。对于 openpyxl 新写入的文件，
                     如果没有在 Excel 中打开保存过，公式单元格的缓存值可能是 None。
                     我们的项目默认使用 False，让公式留空，自己重算更准确。
    """
    # 使用 openpyxl 打开 Excel 文件，data_only 参数控制是否读取缓存计算值
    wb = load_workbook(path, data_only=read_values)
    sheets: list[SheetInfo] = []  # 准备空列表，存储所有解析好的工作表

    # 遍历工作簿中每个工作表
    for ws_name in wb.sheetnames:
        ws = wb[ws_name]  # 获取当前工作表对象
        cells: dict[str, CellInfo] = {}  # 准备空字典，存储当前工作表的所有单元格
        dim = ws.calculate_dimension()  # 计算工作表使用范围（例如 "A1:C10"）
        if dim:  # 只有当工作表不为空时才处理
            # 按行遍历工作表中所有单元格
            for row in ws.iter_rows():
                for cell in row:
                    # 只处理有内容的单元格，跳过空单元格
                    if cell.value is not None:
                        addr = cell.coordinate  # 得到单元格地址，例如 "A1"
                        # 判断是否是公式单元格：cell.value 是字符串，且以 = 开头
                        is_formula = isinstance(cell.value, str) and cell.value.startswith("=")
                        # 构造 CellInfo 并存入字典
                        # 公式单元格：formula 存文本，value 留 None
                        # 常量单元格：value 存值，formula 留 None
                        cells[addr] = CellInfo(
                            ref=CellRef(ws_name, cell.row, cell.column),  # 生成坐标对象
                            formula=cell.value if is_formula else None,   # 公式文本（如果是公式）
                            value=cell.value if not is_formula else None, # 常量值（如果不是公式）
                            is_formula=is_formula,                         # 标记是否为公式
                        )
        # 将当前工作表包装成 SheetInfo，加入工作表列表
        sheets.append(SheetInfo(name=ws_name, cells=cells))

    # 处理命名区域：将 Excel 定义的命名区域读入字典 {名称 -> 引用文本}
    named_ranges: dict[str, str] = {}
    for name, dn in wb.defined_names.items():
        # 兼容不同版本 openpyxl 的属性名差异：优先用 attr_text，回退到 value
        if hasattr(dn, "attr_text") and dn.attr_text:
            named_ranges[name] = dn.attr_text
        elif hasattr(dn, "value"):
            named_ranges[name] = str(dn.value)

    wb.close()  # 关闭打开的 Excel 文件
    # 打包所有解析结果，返回 WorkbookIndex
    return WorkbookIndex(path=path, sheets=sheets, named_ranges=named_ranges)
