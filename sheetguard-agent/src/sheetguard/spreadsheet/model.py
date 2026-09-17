"""SheetGuard 的核心数据模型。

本模块定义了表示 Excel 工作簿的四层内存结构：
    CellRef -> CellInfo -> SheetInfo -> WorkbookIndex

所有下游逻辑（解析、依赖图、静态检查、重算、修复、验证）都只与这套结构
打交道，而不直接操作 Excel 文件，从而把"读取 Excel"与"分析逻辑"解耦。
"""
from __future__ import annotations
from dataclasses import dataclass, field
from typing import Any


@dataclass(frozen=True)
class CellRef:
    """单元格坐标（不可变）。

    sheet 为工作表名（如 "Revenue"），row/col 均为从 1 开始计数的行列号。
    由于使用了 frozen=True，对象创建后不可修改，可以安全地作为 dict 的 key，
    也便于放入集合或作为依赖图的节点。
    """

    sheet: str
    row: int   # 行号，从 1 开始
    col: int   # 列号，从 1 开始

    @property
    def address(self) -> str:
        """把行列号转换为 Excel 地址，如 (row=3, col=2) -> "B3"。"""
        from openpyxl.utils import get_column_letter
        return f"{get_column_letter(self.col)}{self.row}"

    @property
    def full_address(self) -> str:
        """返回带工作表名的完整地址，如 "Revenue!B3"。"""
        return f"{self.sheet}!{self.address}"

    @classmethod
    def parse(cls, raw: str) -> "CellRef":
        """从字符串解析单元格坐标。

        支持两种格式：
          - 无表名："B3"            -> sheet 取默认值 "Sheet"
          - 带表名："Revenue!B3"     -> 自动拆分表名与坐标
        对带引号的表名（如 "'P&L'!B2"）会去掉外层单引号，因为 Excel 中
        含特殊字符的表名必须用引号括起来。
        """
        raw = raw.strip()
        if "!" in raw:
            # Handle quoted sheet names: 'Sheet Name'!A1
            # 处理带引号的工作表名，例如 'Sheet Name'!A1
            idx = raw.rindex("!")          # 取最后一个 !，表名本身可能含 !
            sheet_part = raw[:idx]         # ! 之前是工作表名部分
            cell_part = raw[idx + 1:]      # ! 之后是单元格坐标部分
            if sheet_part.startswith("'") and sheet_part.endswith("'"):
                sheet_part = sheet_part[1:-1]  # 去掉外层单引号
        else:
            sheet_part = "Sheet"           # 没有表名时使用默认工作表名
            cell_part = raw

        from openpyxl.utils import coordinate_to_tuple
        try:
            row, col = coordinate_to_tuple(cell_part)
        except Exception:
            raise ValueError(f"Invalid cell reference: {raw}")
        return cls(sheet=sheet_part, row=row, col=col)

    def __hash__(self) -> int:
        """基于 (sheet, row, col) 生成哈希，使 CellRef 可作为集合/字典 key。"""
        return hash((self.sheet, self.row, self.col))

    def __eq__(self, other: object) -> bool:
        """按 (sheet, row, col) 判断两个 CellRef 是否相等。"""
        if not isinstance(other, CellRef):
            return NotImplemented
        return (self.sheet, self.row, self.col) == (other.sheet, other.row, other.col)


@dataclass
class CellInfo:
    """单个单元格的内容信息。

    ref        : 单元格坐标
    formula    : 公式文本（若为公式单元格），否则为 None
    value      : 单元格的缓存值（可能是数字/文本/布尔等）
    is_formula : 标记该单元格是否为公式单元格
    """

    ref: CellRef
    formula: str | None
    value: Any = None
    is_formula: bool = False

    def to_dict(self) -> dict:
        """把对象序列化为纯 dict，方便写入 JSON 报告或跨进程传递。"""
        return {
            "ref": self.ref.full_address,
            "formula": self.formula,
            "value": self.value,
            "is_formula": self.is_formula,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "CellInfo":
        """从 dict 还原 CellInfo 对象（to_dict 的逆操作）。"""
        return cls(
            ref=CellRef.parse(d["ref"]),
            formula=d.get("formula"),
            value=d.get("value"),
            is_formula=d.get("is_formula", False),
        )


@dataclass
class SheetInfo:
    """单个工作表的索引。

    name  : 工作表名
    cells : 该表内所有已索引单元格，key 为不含表名的地址（如 "A1"），
            value 为 CellInfo。表名放在本类 name 字段，避免重复存储。
    """

    name: str
    cells: dict[str, CellInfo] = field(default_factory=dict)  # key: "A1" address

    def to_dict(self) -> dict:
        return {
            "name": self.name,
            "cells": {k: v.to_dict() for k, v in self.cells.items()},
        }

    @classmethod
    def from_dict(cls, d: dict) -> "SheetInfo":
        return cls(
            name=d["name"],
            cells={k: CellInfo.from_dict(v) for k, v in d.get("cells", {}).items()},
        )


@dataclass
class WorkbookIndex:
    """整个工作簿在内存中的统一索引。

    这是项目的核心入口数据结构，由 parser 从 .xlsx 构建，
    供依赖图、静态检查器、重算器、调查工具、验证器等模块共同使用。

    path         : 源 .xlsx 文件路径
    sheets       : 所有工作表的索引列表
    named_ranges : 命名区域映射（名称 -> 引用文本）
    """

    path: str
    sheets: list[SheetInfo] = field(default_factory=list)
    named_ranges: dict[str, str] = field(default_factory=dict)

    def cell(self, full_address: str) -> CellInfo | None:
        """按完整地址（如 "Sheet1!A1"）查询单元格，找不到时返回 None。"""
        sheet_name, _, addr = full_address.partition("!")
        for s in self.sheets:
            if s.name == sheet_name:
                return s.cells.get(addr)
        return None

    def _sheet_by_name(self, name: str) -> SheetInfo | None:
        """按名称查找工作表索引，找不到时返回 None。"""
        for s in self.sheets:
            if s.name == name:
                return s
        return None

    def formula_cells(self):
        """遍历所有公式单元格（生成器），供统计与检测使用。"""
        for s in self.sheets:
            for info in s.cells.values():
                if info.is_formula:
                    yield info

    def formula_count(self) -> int:
        """统计整个工作簿中公式单元格的总数。"""
        return sum(1 for _ in self.formula_cells())

    def cross_sheet_ref_count(self) -> int:
        """粗略统计跨工作表引用的公式数。

        注意：这里只是检查公式文本中是否出现 "!"，是近似估算；
        精确的跨表依赖统计应使用 DependencyGraph。
        """
        count = 0
        for info in self.formula_cells():
            if info.formula and "!" in info.formula:
                count += 1
        return count

    def to_dict(self) -> dict:
        """把整个索引序列化为 dict（JSON 友好的结构）。"""
        return {
            "path": self.path,
            "sheets": [s.to_dict() for s in self.sheets],
            "named_ranges": dict(self.named_ranges),
        }

    @classmethod
    def from_dict(cls, d: dict) -> "WorkbookIndex":
        """从 dict 还原 WorkbookIndex（to_dict 的逆操作）。"""
        return cls(
            path=d["path"],
            sheets=[SheetInfo.from_dict(s) for s in d.get("sheets", [])],
            named_ranges=d.get("named_ranges", {}),
        )
