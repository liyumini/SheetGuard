"""定义一条合成测评案例需要保存的中文语义数据结构。"""
from __future__ import annotations

from dataclasses import asdict, dataclass


@dataclass
class CaseMetadata:
    """记录一份 broken 工作簿与其 gold 正确答案之间的对应关系。

    LangSmith 上传案例时，会根据这些字段知道：错误在哪里、正确公式是什么，
    以及应读取哪一份 broken/gold 工作簿。
    """

    case_id: str       # 案例唯一编号，例如 case_000。
    error_type: str    # 人为注入的错误类型，例如 wrong_range。
    target_cell: str   # 被故意改错的单元格地址，例如 Revenue!B2。
    old_formula: str   # broken 工作簿中当前的错误公式或硬编码值。
    gold_formula: str  # gold 工作簿中该单元格应有的正确公式。
    gold_path: str     # 正确工作簿的文件路径。
    broken_path: str   # 含错误工作簿的文件路径。

    def to_dict(self) -> dict:
        """转换为可直接写入 cases.json 的普通字典。"""
        return asdict(self)
