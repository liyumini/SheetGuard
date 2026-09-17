"""读取已经生成的 LangSmith 合成案例索引文件。"""
from __future__ import annotations

import json
from pathlib import Path

from evaluation.generation.models import CaseMetadata


def load_cases(cases_json_path: str | Path) -> list[CaseMetadata]:
    """从 cases.json 读取案例，并转成 CaseMetadata 对象列表。

    该函数不生成或修改 Excel；它只负责把保存到磁盘的案例说明读回内存。
    """
    with open(cases_json_path, encoding="utf-8") as file:
        return [CaseMetadata(**item) for item in json.load(file)]
