"""SheetGuard LangGraph 工作流使用的显式状态定义。

每个字段都是 LangGraph 各节点共享的状态，字段名保持代码中的英文标识符。

multi-cell 批处理状态采用"待处理队列 + 当前候选 + 每候选结果"模型
（设计 6），同时明确区分两个工作簿索引：
- ``source_index``：原始工作簿索引，整个运行期间不变，用于原始值/
  公式基线、回归检查和审计；
- ``working_index``：当前工作副本索引，只反映已经成功提交的修复，
  供后续候选调查、诊断、修复和依赖分析使用。
"""
from __future__ import annotations
from typing import TypedDict


class SheetGuardState(TypedDict):
    workbook_path: str  # 工作簿文件路径
    source_index: dict  # 原始工作簿索引（运行期间不变，事实基线）
    working_index: dict  # 当前工作副本索引（只含已成功提交的修复）
    dependency_graph: dict  # working 依赖图摘要
    anomalies: list[dict]  # 完整 static candidate universe（不受 top_k 截断）
    static_candidate_count: int  # 完整静态候选全集大小
    seed_candidates: list[str]  # 满足 min_anomaly_score 的高置信度候选
    eligible_candidates: list[str]  # 通过预算准入的目标及其 hard prerequisites（设计 5.6）
    active_candidates: list[str]  # 运行开始时按 max_candidates 冻结的 active batch
    pending_candidates: list[str]  # active 中仍待处理的候选（可重排）
    current_candidate: str | None  # 当前正在调查/修复的候选
    current_attempt: int  # 当前候选已消耗的 repair attempts 数
    working_revision: int  # 稳定工作副本每次成功 commit 递增
    candidate_results: dict  # {target: 结果记录}，批处理结果唯一事实源
    evidence: list[dict] | list[str]  # 当前候选调查证据
    hypothesis: dict | None  # 当前候选诊断假设
    proposed_patch: dict | None  # 当前候选本轮提议的修复补丁
    verification: dict | None  # 当前候选本轮验证结果
    attempt_path: str | None  # 当前 attempt 副本路径
    repaired_workbook: str | None  # 导出的修复工作簿路径
    min_anomaly_score: float  # seed candidates 静态可疑分阈值
    max_candidates: int  # 一次运行进入 active batch 的候选数量上限
    max_attempts: int  # 单个候选含首次尝试在内的最大 repair attempts
    review_certified: list  # 上一轮用户认证的格子 [{"target","formula"}]
    review_rejected: dict  # 本轮被否格 {cell: {"rejected_formulas": [...], "remark": str|None}}
    status: str  # 当前状态
    error: str | None  # 错误信息
