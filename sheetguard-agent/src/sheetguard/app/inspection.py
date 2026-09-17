"""离线静态检测负载（inspect 命令与 web 检测端点共用）。"""
from __future__ import annotations

from pathlib import Path

from sheetguard.spreadsheet.anomaly_detector import StaticInspector
from sheetguard.spreadsheet.dependency_graph import DependencyGraph
from sheetguard.spreadsheet.parser import parse_workbook


def inspection_payload(path: Path) -> dict:
    index = parse_workbook(str(path))
    graph = DependencyGraph(index)
    graph.build()
    candidates = StaticInspector(index).detect()
    cross_sheet_edges = sum(
        1
        for source, target in graph._graph.edges()
        if source.split("!", 1)[0] != target.split("!", 1)[0]
    )
    return {
        "workbook": str(path),
        "sheets": len(index.sheets),
        "formula_cells": index.formula_count(),
        "cross_sheet_dependencies": cross_sheet_edges,
        "candidates": candidates,
    }
