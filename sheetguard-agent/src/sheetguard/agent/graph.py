"""结构化 SheetGuard multi-cell 批处理工作流。

流程（设计 4/12）：
    parse → build_index → inspect → plan_batch
        → advance ⇄ process_candidate → finalize

- ``plan_batch``：min_anomaly_score 选 seed，补齐 prerequisite closure，
  依赖拓扑排序后按 max_candidates 冻结 active batch；working graph 中
  既有 dependency cycle 的 SCC 及其候选级下游直接 skipped。
- ``process_candidate``：单个候选的完整事务
  investigate → diagnose → repair → attempt(copy-on-write) →
  proposal syntax precheck → tentative graph → pre-commit dependency
  check → verify → commit/discard，含 max_attempts 有界重试。
- ``advance``：把下一 pending 候选设为 current；队列空则 finalize。
- ``finalize``：计算最终批处理状态、导出修复工作簿。

关键约束：source_index/source 工作簿整个运行期间不变；失败 attempt 直接
丢弃，稳定 working 副本只累积成功 commit 的修复；每个候选使用独立模型
对话上下文；纯依赖重排不消耗 max_attempts（working_revision 防重复）。
"""
from __future__ import annotations

import json
import logging
import re
from contextlib import contextmanager, nullcontext
from pathlib import Path
from typing import Any, Callable
from uuid import uuid4

from langchain_core.language_models import BaseChatModel
from langchain_core.messages import AIMessage, HumanMessage, ToolMessage
from langgraph.checkpoint.memory import MemorySaver
from langgraph.graph import END, START, StateGraph
from langfuse import get_client, propagate_attributes
from langfuse.langchain import CallbackHandler

from sheetguard.agent.prompts import (
    DIAGNOSE_PROMPT,
    INVESTIGATE_PROMPT,
    REPAIR_PROMPT,
)
from sheetguard.agent.state import SheetGuardState
from sheetguard.config import (
    DEFAULT_MAX_CANDIDATES,
    DEFAULT_MIN_ANOMALY_SCORE,
    MAX_ATTEMPTS,
    MAX_LOCALIZER_ITERATIONS,
    MAX_LOCALIZER_TOOL_CALLS,
    MAX_RETRIES,
    SheetGuardConfig,
)
from sheetguard.spreadsheet.anomaly_detector import StaticInspector
from sheetguard.spreadsheet.formula_expectation import dependency_judgement
from sheetguard.spreadsheet.dependency_graph import DependencyGraph
from sheetguard.spreadsheet.model import WorkbookIndex
from sheetguard.spreadsheet.parser import parse_workbook
from sheetguard.spreadsheet.patcher import SandboxPatcher
from sheetguard.spreadsheet.recalc import validate_formula
from sheetguard.spreadsheet.value_hints import aggregate_value_hints, format_value_hints
from sheetguard.spreadsheet.verifier import Verifier
from sheetguard.tools.investigation import make_investigation_tools

# 行为常量已收拢至 sheetguard.config（支持 YAML 外置）；这里保留模块级
# 名称供既有导入方使用（tests/test_graph.py 等直接 import 这些名字）。

# 候选最终状态（设计 5.4/6）
STATUS_FIXED = "fixed"
STATUS_FAILED = "failed"
STATUS_DISMISSED = "dismissed"
STATUS_UNRESOLVED = "unresolved"
STATUS_SKIPPED = "skipped_due_to_dependency"
STATUS_DEFERRED_BUDGET = "deferred_due_to_budget"
STATUS_DEFERRED_BUDGET_DEP = "deferred_due_to_budget_dependency"
STATUS_DEFERRED_INACTIVE = "deferred_due_to_inactive_prerequisite"
DEFERRED_STATUSES = {
    STATUS_DEFERRED_BUDGET,
    STATUS_DEFERRED_BUDGET_DEP,
    STATUS_DEFERRED_INACTIVE,
}
# deferred 风险分级（v1.6 设计 6.2）：快速区分"低价值延期"与
# "高置信真实错误被预算挡掉"。
DEFERRED_CONFIDENCE_HIGH = "high"
DEFERRED_CONFIDENCE_MEDIUM = "medium"
DEFERRED_CONFIDENCE_LOW = "low"
CONFIRMED = "confirmed"

# 信号类型强度（设计 5.2）：直接结构矛盾（公式缺失）与结构性 range/模板
# 错误通常比单一 dependency_anomaly 更强；单一弱信号的高总分可能掩盖证据
# 质量，因此弱信号只用类型强度与互证加成计分。
_SIGNAL_TYPE_WEIGHT = {
    "missing_formula": 1.0,
    "range_boundary": 0.9,
    "pattern_anomaly": 0.8,
    "dependency_anomaly": 0.4,
    "neighbor_mismatch": 0.3,
}
# 强证据信号集合：命中即视为 hard prerequisite（设计 5.5）。
_STRONG_PREREQ_SIGNALS = frozenset(
    {"missing_formula", "range_boundary", "pattern_anomaly"}
)

logger = logging.getLogger(__name__)  # 留痕 scheduler stall 等异常路径


def _parse_json_response(text: str) -> dict:
    """从 LLM 响应中提取 JSON 对象，且不抛出异常。

    模型可能返回纯 JSON、Markdown JSON 代码块，或夹杂解释文字的 JSON。
    该方法按多种候选格式依次尝试解析，失败时返回错误字典而不是抛异常。
    """
    if not isinstance(text, str) or not text.strip():
        return {"error": "Could not parse JSON", "raw": text}
    fenced = re.search(r"```json\s*(.*?)\s*```", text, re.DOTALL | re.IGNORECASE)
    candidates = [fenced.group(1)] if fenced else []
    candidates.append(text.strip())
    object_match = re.search(r"\{.*\}", text, re.DOTALL)
    if object_match:
        candidates.append(object_match.group(0))
    for candidate in candidates:
        try:
            parsed = json.loads(candidate)
            if isinstance(parsed, dict):
                return parsed
        except (TypeError, json.JSONDecodeError):
            continue
    return {"error": "Could not parse JSON", "raw": text}


def _message_text(message: Any) -> str:
    """统一提取 LangChain 消息或普通对象中的文本内容。"""
    content = getattr(message, "content", message)
    return content if isinstance(content, str) else str(content)


def _clone_record(record: dict) -> dict:
    """深拷贝一条候选结果记录（纯 JSON 结构，round-trip 安全）。"""
    return json.loads(json.dumps(record, ensure_ascii=False))


def _merge_blockers(*groups: list[dict]) -> list[dict]:
    """合并多组 blocker 根因，去重后保留全部根因（设计 5.4）。"""
    merged: list[dict] = []
    seen: set[tuple] = set()
    for group in groups:
        for entry in group or []:
            if not isinstance(entry, dict):
                continue
            key = (
                tuple(sorted(entry.get("targets", [])))
                if "targets" in entry
                else (entry.get("target", ""),),
                entry.get("cause", ""),
            )
            if key not in seen:
                seen.add(key)
                merged.append(dict(entry))
    return merged


def _deferred_confidence(record: dict) -> str:
    """deferred 记录的风险分级（v1.6 设计 6.2）。

    高置信延期 = 静态证据足够强（单独 0.6+ 的信号，或 0.5 分叠加两个
    以上信号）却被预算/依赖挡掉的候选；低分延期标记为 medium/low。
    """
    score = float(record.get("score") or 0.0)
    signal_count = len(record.get("signals") or [])
    if score >= 0.6 or (score >= 0.5 and signal_count >= 2):
        return DEFERRED_CONFIDENCE_HIGH
    if score >= 0.45:
        return DEFERRED_CONFIDENCE_MEDIUM
    return DEFERRED_CONFIDENCE_LOW


class SheetGuardGraph:
    """围绕 LLM 决策设置确定性边界的有状态批处理工作流。"""

    def __init__(
        self,
        model: BaseChatModel,
        max_attempts: int | None = None,
        *,
        max_retries: int | None = None,  # 兼容旧参数名，内部映射为 max_attempts
        max_localizer_iterations: int | None = None,
        max_localizer_tool_calls: int | None = None,
        min_anomaly_score: float | None = None,
        max_candidates: int | None = None,
        config: SheetGuardConfig | None = None,
    ):
        """初始化结构化工作流及其运行时依赖。

        ``config`` 提供一次性传入全部行为参数（通常来自 YAML，
        见 ``sheetguard.config.load_config``）；未指定时使用内置默认值。
        其余显式参数优先级高于 ``config``（CLI 覆盖 > YAML > 默认），
        未提供的参数以 None 传入即可。

        ``max_attempts`` 是单个候选含首次尝试在内的最大 repair attempts
        （syntax-invalid、proposal 新引入 cycle、Verifier 失败都消耗）；
        ``max_retries`` 作为兼容别名。``min_anomaly_score`` 决定 seed
        candidates；``max_candidates`` 冻结一次运行的 active batch。
        """
        self.model = model
        self.config = config if config is not None else SheetGuardConfig()
        effective_attempts = (
            max_attempts
            if max_attempts is not None
            else (max_retries if max_retries is not None else self.config.max_attempts)
        )
        self.max_attempts = max(1, int(effective_attempts))
        self.max_localizer_iterations = (
            max_localizer_iterations
            if max_localizer_iterations is not None
            else self.config.max_localizer_iterations
        )
        tool_calls = (
            max_localizer_tool_calls
            if max_localizer_tool_calls is not None
            else self.config.max_localizer_tool_calls
        )
        # 下限钳 0 防御直接传参；上限不做硬钳——配置即真相，
        # 调查成本由 config 的值自行约束（配置指纹随审计记录）。
        self.max_localizer_tool_calls = max(0, int(tool_calls))
        self.min_anomaly_score = float(
            min_anomaly_score
            if min_anomaly_score is not None
            else self.config.min_anomaly_score
        )
        self.max_candidates = max(
            1, int(max_candidates if max_candidates is not None else self.config.max_candidates)
        )
        self._verifier = Verifier()
        self._patcher: SandboxPatcher | None = None  # 当前工作簿事务管理器
        self._source_index: WorkbookIndex | None = None  # 原始索引（事实基线，不变）
        self._working_index: WorkbookIndex | None = None  # 工作副本索引（随 commit 演进）
        self._working_graph: DependencyGraph | None = None  # working 依赖图
        self._tentative_graph: DependencyGraph | None = None  # 当前 attempt 的依赖图
        self._tools: list[Any] = []  # 供 LLM 调用的调查工具列表
        self._tool_map: dict[str, Any] = {}  # 工具名称映射
        self._anomaly_map: dict[str, dict] = {}  # universe 的分数/信号查询表
        # (candidate, upstream) → working_revision：识别同一依赖条件下
        # 重复重排而没有图进展的 scheduler stall（设计 4）。
        self._reschedule_log: dict[tuple[str, str], int] = {}
        # 重排过程中发现的"候选级上游边"（dependent → upstreams）累积表：
        # 硬编码常量候选在源图里没有向上依赖边，提案才会揭示真实依赖；
        # 这些边必须跨候选累积，重排时对整个 pending 做拓扑排序，
        # 否则链式候选（A→B→C）中后段上游的重排会把先排好的依赖者
        # 重新甩到自己前面（09-12 hard_002 Dashboard!D2 stall 根因）。
        self._scheduler_extra_edges: dict[str, set[str]] = {}
        self._refreshed_working_index: WorkbookIndex | None = None  # commit 后待写回状态的新索引
        self._export_target: Path | None = None  # 修复工作簿导出目标
        # 审查反馈循环（v1.9）：上一轮用户认证的格子与本轮被否格。
        self._review_certified: list = []
        self._review_rejected: dict = {}
        self._graph = None  # 编译后的 LangGraph 工作流

    def build(self):
        """注册节点和边，并编译 LangGraph 工作流。

        这里只负责组装流程，不会读取工作簿；真正执行发生在 invoke() 中。
        """
        builder = StateGraph(SheetGuardState)
        builder.add_node("parse", self._parse_node)  # 解析源工作簿 → source_index
        builder.add_node("build_index", self._build_index_node)  # working 副本 + 依赖图 + 调查工具
        builder.add_node("inspect", self._inspect_node)  # 完整 static candidate universe
        builder.add_node("plan_batch", self._plan_batch_node)  # seed/闭包/排序/冻结 active batch
        builder.add_node("process_candidate", self._process_candidate_node)  # 单候选事务
        builder.add_node("advance", self._advance_node)  # 推进到下一候选
        builder.add_node("finalize", self._finalize_node)  # 最终状态/导出

        builder.add_edge(START, "parse")
        builder.add_conditional_edges(
            "parse", self._route_after_parse, {"continue": "build_index", "failed": END}
        )
        builder.add_conditional_edges(
            "build_index",
            self._route_after_build_index,
            {"continue": "inspect", "failed": END},
        )
        builder.add_edge("inspect", "plan_batch")
        builder.add_conditional_edges(
            "plan_batch",
            self._route_next_candidate,
            {"process": "process_candidate", "finalize": "finalize"},
        )
        builder.add_edge("process_candidate", "advance")
        builder.add_conditional_edges(
            "advance",
            self._route_next_candidate,
            {"process": "process_candidate", "finalize": "finalize"},
        )
        builder.add_edge("finalize", END)
        self._graph = builder.compile(checkpointer=MemorySaver())
        return self._graph

    def invoke(
        self,
        workbook_path: str,
        repaired_workbook: str | Path | None = None,
        review_context: dict | None = None,
        progress_callback: Callable[[dict], None] | None = None,
    ) -> dict:
        """运行一次完整批处理工作流，返回最终状态和批量审计记录。

        每次运行都会重新初始化索引、依赖图、工具和工作簿事务，避免复用
        上一个工作簿的运行时数据。``repaired_workbook`` 指定修复副本的
        导出路径；不指定时只输出审计报告，临时副本运行结束即清理。

        ``review_context`` 携带上一轮用户审查反馈：
        ``{"certified": [{"target", "formula"}, ...],
        "rejected": {cell: {"rejected_formulas": [...], "remark": ...}}}``。
        认证格本轮直接排除；被否格强制入选种子并注入用户备注；
        被否公式不得重复提案。不传时行为与旧版完全一致。

        ``progress_callback`` 可选：每收到一个 LangGraph 节点更新即被调用一次，
        负载为 ``{"stage": str, "done": int, "total": int,
        "current": {"cell", "attempt", "max_attempts"} | None}``，供 web 层
        /api/jobs 直接透传展示管线进度。传 None 时行为与旧版完全一致
        （仍走 ``_graph.invoke``）。
        """
        # 开始新任务前，先删除上一次运行遗留的副本文件，然后重置运行时数据。
        self.cleanup()
        self._source_index = None
        self._working_index = None
        self._working_graph = None
        self._tentative_graph = None
        self._tools = []
        self._tool_map = {}
        self._anomaly_map = {}
        self._reschedule_log = {}
        self._scheduler_extra_edges = {}
        self._export_target = Path(repaired_workbook) if repaired_workbook else None
        ctx = review_context or {}
        self._review_certified = list(ctx.get("certified") or [])
        self._review_rejected = dict(ctx.get("rejected") or {})
        if self._graph is None:
            self.build()

        initial: SheetGuardState = {
            "workbook_path": str(workbook_path),
            "source_index": {},
            "working_index": {},
            "dependency_graph": {},
            "anomalies": [],
            "static_candidate_count": 0,
            "seed_candidates": [],
            "eligible_candidates": [],
            "active_candidates": [],
            "pending_candidates": [],
            "current_candidate": None,
            "current_attempt": 0,
            "working_revision": 0,
            "candidate_results": {},
            "evidence": [],
            "hypothesis": None,
            "proposed_patch": None,
            "verification": None,
            "attempt_path": None,
            "repaired_workbook": None,
            "min_anomaly_score": self.min_anomaly_score,
            "max_candidates": self.max_candidates,
            "max_attempts": self.max_attempts,
            "review_certified": self._review_certified,
            "review_rejected": self._review_rejected,
            "status": "start",
            "error": None,
        }
        session_id = f"{Path(workbook_path).stem}-{uuid4().hex}"
        langfuse = get_client()
        langfuse_handler = CallbackHandler()
        # 每次调用使用唯一 thread_id，避免 MemorySaver 混淆不同运行实例。
        config = {
            "configurable": {"thread_id": session_id},
            "callbacks": [langfuse_handler],
            "run_name": "execute-sheetguard-graph",
        }
        model_name = getattr(self.model, "model_name", None)
        if model_name:
            # Langfuse uses this LangChain metadata key when an OpenAI-compatible
            # provider does not expose its model name in the callback payload.
            config["metadata"] = {"ls_model_name": str(model_name)}
        try:
            with langfuse.start_as_current_observation(
                name="sheetguard-repair",
                as_type="agent",
                input={"workbook_name": Path(workbook_path).name},
            ) as root_observation:
                with propagate_attributes(
                    trace_name="sheetguard-repair",
                    session_id=session_id,
                    tags=["sheetguard", "structured-agent", "multi-cell"],
                    metadata={"workbook_name": Path(workbook_path).name},
                ):
                    if progress_callback is None:
                        final_state = self._graph.invoke(initial, config)
                    else:
                        final_state = self._stream_with_progress(
                            initial, config, progress_callback
                        )
                root_observation.update(output=self._batch_summary(final_state))
        finally:
            langfuse.flush()
        # 从完整状态中提取适合日志和 CLI 输出的批量审计信息。
        final_state["audit"] = self._build_audit(final_state)
        # 临时 working/attempt 副本不是稳定交付物；导出文件在 cleanup 后保留。
        self.cleanup()
        return final_state

    # ------------------------------------------------------------------
    # 运行时索引/图恢复
    # ------------------------------------------------------------------

    def _runtime_source_index(self, state: SheetGuardState) -> WorkbookIndex:
        """获取运行时 source_index，并在首次使用时从状态恢复。"""
        if self._source_index is None:
            self._source_index = WorkbookIndex.from_dict(state["source_index"])
        return self._source_index

    def _runtime_working_index(self, state: SheetGuardState) -> WorkbookIndex:
        """获取运行时 working_index，并在首次使用时从状态恢复。"""
        if self._working_index is None:
            self._working_index = WorkbookIndex.from_dict(state["working_index"])
        return self._working_index

    def _runtime_working_graph(self, state: SheetGuardState) -> DependencyGraph:
        """获取 working 依赖图，并在首次使用时构建。"""
        if self._working_graph is None:
            graph = DependencyGraph(self._runtime_working_index(state))
            graph.build()
            self._working_graph = graph
        return self._working_graph

    def _rebuild_investigation_tools(self, state: SheetGuardState) -> None:
        """基于新的 working_index 和依赖图重建调查工具。

        工具闭包绑定创建时的索引和依赖图；commit 后不重建会导致
        后续候选读到修复前的公式和依赖关系（设计 7.2）。
        """
        index = self._runtime_working_index(state)
        graph = self._runtime_working_graph(state)
        self._tools = make_investigation_tools(index, graph)
        self._tool_map = {tool.name: tool for tool in self._tools}

    # ------------------------------------------------------------------
    # Langfuse observation 辅助
    # ------------------------------------------------------------------

    @contextmanager
    def _observation(self, name: str, as_type: str = "span", **attributes: Any):
        """打开一个 Langfuse 观测 span；离线或异常时退化为空上下文。"""
        try:
            langfuse = get_client()
            manager = langfuse.start_as_current_observation(
                name=name, as_type=as_type, **attributes
            )
        except Exception:
            manager = nullcontext()
        with manager as observation:
            yield observation

    @staticmethod
    def _safe_update(observation: Any, **attributes: Any) -> None:
        """span 更新失败不影响主流程。"""
        try:
            if observation is not None:
                observation.update(**attributes)
        except Exception:
            pass

    # ------------------------------------------------------------------
    # 管线进度回调（stream 模式）
    # ------------------------------------------------------------------

    # 进度阶段映射：LangGraph 状态 status → 用户可读阶段。
    # 键 = 图节点实际写入的 status 值（planned/candidate_advanced/finalizing
    # 来自调度与逐格循环，末值为批处理终态），不识别的值原样透传。
    _STAGE_NAMES = {
        "start": "启动",
        "parsed": "解析",
        "indexed": "索引",
        "inspected": "静态检测",
        "planned": "批次规划",
        "no_eligible_candidates": "无入选候选",
        "candidate_advanced": "逐格修复",
        "candidate_processed": "逐格修复",
        "finalizing": "汇总导出",
        "no_candidates": "无候选",
        "success": "修复成功",
        "partial_success": "部分成功",
        "completed_without_repairs": "完成（无需修复）",
        "failed": "修复失败",
        "error": "运行错误",
    }

    def _stream_with_progress(self, initial: dict, config: dict, progress_callback) -> dict:
        """stream 模式运行图：逐节点合并状态并回调进度（与 invoke 最终状态等价）。"""
        final_state = dict(initial)
        for chunk in self._graph.stream(initial, config, stream_mode="updates"):
            for _node, update in chunk.items():
                final_state.update(update or {})
            progress_callback(self._progress_payload(final_state))
        return final_state

    def _progress_payload(self, state: dict) -> dict:
        """从图状态提取进度负载：done 只统计 active 内已完成候选。"""
        active = state.get("active_candidates") or []
        results = state.get("candidate_results") or {}
        done = sum(1 for c in active if c in results)
        current = state.get("current_candidate")
        return {
            "stage": self._STAGE_NAMES.get(
                state.get("status") or "", state.get("status") or ""
            ),
            "done": done,
            "total": len(active),
            "current": None if not current else {
                "cell": current,
                "attempt": state.get("current_attempt") or 0,
                "max_attempts": state.get("max_attempts") or 0,
            },
        }

    # ------------------------------------------------------------------
    # 图节点：parse / build_index / inspect
    # ------------------------------------------------------------------

    def _parse_node(self, state: SheetGuardState) -> dict:
        """解析原始 Excel 工作簿，写入 source_index（整个运行期间不变）。"""
        try:
            index = parse_workbook(state["workbook_path"])
        except Exception as exc:
            return {
                "status": "failed",
                "error": f"parse failed: {type(exc).__name__}: {exc}",
            }
        # 保存对象供当前进程后续节点复用，同时把可序列化字典写回 state。
        self._source_index = index
        return {"source_index": index.to_dict(), "status": "parsed", "error": None}

    def _build_index_node(self, state: SheetGuardState) -> dict:
        """校验公式语法、构建 working 副本/依赖图，并创建调查工具。"""
        try:
            index = self._runtime_source_index(state)
            # 先验证公式语法，避免非法公式进入后续依赖图和验证流程。
            for info in index.formula_cells():
                validate_formula(info.formula, info.ref)
            # working 副本初始等于 source（尚无已提交修复）。
            working = WorkbookIndex.from_dict(index.to_dict())
            self._working_index = working
            graph = DependencyGraph(working)
            graph.build()
            self._working_graph = graph
            # 工作簿事务管理器：working 副本从源文件复制，源文件保持只读。
            if self._patcher is None:
                self._patcher = SandboxPatcher(state["workbook_path"])
                self._patcher.create_working_copy()
            self._rebuild_investigation_tools(state)
            return {
                "source_index": index.to_dict(),
                "working_index": working.to_dict(),
                "dependency_graph": {
                    "nodes": len(graph._graph.nodes),
                    "edges": len(graph._graph.edges),
                },
                "status": "indexed",
                "error": None,
            }
        except Exception as exc:
            return {
                "status": "failed",
                "error": f"index failed: {type(exc).__name__}: {exc}",
            }

    @staticmethod
    def _route_after_parse(state: SheetGuardState) -> str:
        """根据解析状态决定进入建图节点还是直接结束。"""
        return (
            "continue"
            if state.get("status") == "parsed" and state.get("source_index")
            else "failed"
        )

    @staticmethod
    def _route_after_build_index(state: SheetGuardState) -> str:
        """根据索引和依赖图构建结果决定是否进入静态检查。"""
        return "continue" if state.get("status") == "indexed" else "failed"

    def _inspect_node(self, state: SheetGuardState) -> dict:
        """产出完整 static candidate universe（不受旧 top_k 截断，验收 31）。"""
        universe = StaticInspector(
            self._runtime_working_index(state),
            weights=self.config.signal_weights,
        ).detect_all()
        self._anomaly_map = {
            item["cell"]: item
            for item in universe
            if isinstance(item, dict) and item.get("cell")
        }
        return {
            "anomalies": universe,
            "static_candidate_count": len(universe),
            # 兼容旧展示层字段：完整 universe 地址列表（不受 top_k 截断）。
            "candidates": [item["cell"] for item in universe],
            "status": "inspected",
        }

    # ------------------------------------------------------------------
    # 图节点：plan_batch（seed 过滤 + 闭包 + 排序 + 冻结 active batch）
    # ------------------------------------------------------------------

    def _plan_batch_node(self, state: SheetGuardState) -> dict:
        """seed filter → prerequisite closure → 排序 → 冻结 active batch。

        - static candidate universe 来自完整 StaticInspector 集合；
        - 循环依赖：SCC 成员及其候选级下游直接 skipped（cause=dependency_cycle）；
        - 排序后前 max_candidates 个冻结为 active，其余 deferred_due_to_budget。
        """
        universe = state.get("anomalies") or []
        # 认证格先排除、被否格后合成，二者都不在空 universe 早退之前：
        # 下一轮输入可能是上一轮修复后的工作簿（静态 0 候选），但用户
        # 否了其中一格的提案——被否格必须由 review_context 合成入选，
        # 否则空早退会静默吞掉用户反馈（设计 5）。
        # 优先级：同一格既 certified 又 rejected 的矛盾输入下，认证排除
        # 在前、被否合成在后，最终行为是入选；本任务不做输入校验。
        # 两个集合皆空时循环体不执行，universe 与原逻辑完全一致。
        for cell in (state.get("review_certified") or []):
            target = cell.get("target") if isinstance(cell, dict) else cell
            universe = [
                item for item in universe
                if not (isinstance(item, dict) and item.get("cell") == target)
            ]
        for cell in state.get("review_rejected") or {}:
            if not any(
                isinstance(item, dict) and item.get("cell") == cell
                for item in universe
            ):
                # 被否格用户已担保"这里有问题"：静态全集没有时补合成条目，
                # 使其必然可入选种子（跳过 min_anomaly_score，设计 5）。
                universe = universe + [
                    {"cell": cell, "score": 0.0, "signals": ["user_rejected"]},
                ]
        if not universe:
            return {
                "candidate_results": {},
                "pending_candidates": [],
                "current_candidate": None,
                "status": "no_candidates",
            }
        self._anomaly_map = {
            item["cell"]: item
            for item in universe
            if isinstance(item, dict) and item.get("cell")
        }
        # 认证格已在上面的 universe 过滤中移除，_anomaly_map 不会包含它们。
        scores = {
            cell: float(item.get("score") or 0.0)
            for cell, item in self._anomaly_map.items()
        }
        universe_cells = set(scores)
        graph = self._runtime_working_graph(state)

        # seed candidates：满足 min_anomaly_score 的高置信度候选；
        # 被否格无条件入选（用户已担保此处有问题，跳过分数阈值）。
        review_rejected = set(state.get("review_rejected") or {})
        seeds = [
            cell
            for cell in sorted(scores, key=lambda c: (-scores[c], c))
            if cell in review_rejected or scores[cell] >= self.min_anomaly_score
        ]
        if not seeds:
            return {
                "seed_candidates": [],
                "eligible_candidates": [],
                "active_candidates": [],
                "pending_candidates": [],
                "current_candidate": None,
                "candidate_results": {},
                "status": "no_eligible_candidates",
            }

        # working graph 既有循环：SCC 成员及依赖该 SCC 的候选级下游 → skipped，
        # 循环安全语义先于预算准入、完全保留（设计 8.2 cycle 行）。
        cycle_blocked = self._cycle_blocked_candidates(graph, universe_cells)
        available = universe_cells - set(cycle_blocked)

        # 风险序（设计 5.3.1）：静态分数优先，分数并列时候选置信度加成，地址稳定；
        # cycle-blocked 候选已被 skipped，不参与准入。
        risk_order = sorted(
            (cell for cell in seeds if cell not in cycle_blocked),
            key=lambda c: (-scores[c], -self._candidate_confidence(c, scores), c),
        )

        # Budget Admission + Lazy Prerequisite Resolution（设计 5.3/5.4/5.6）：
        # 公式依赖不再自动等于修复前置——准备接纳一个目标时才解析它的候选
        # 上游：强证据上游升级为 hard prerequisite 并与目标一起计入预算；
        # 弱/单一证据上游记为 inactive prerequisite，不阻塞也不占预算；
        # 最后只对 admitted 集合做拓扑排序。
        admitted: set[str] = set()
        inactive_prereqs: set[str] = set()
        deferred_dep: dict[str, list[str]] = {}
        deferred_budget: list[str] = []
        for target in risk_order:
            hard, weak = self._resolve_prerequisites(graph, target, available)
            required = hard | {target}
            if len(admitted | required) <= self.max_candidates:
                admitted |= required
                inactive_prereqs.update(weak)
            elif len(admitted | {target}) <= self.max_candidates:
                # hard 前置装不下：目标延后并如实记录前置，不假装已处理（设计 8.2）。
                deferred_dep[target] = sorted(hard)
            else:
                deferred_budget.append(target)

        ordered = graph.order_candidates(sorted(admitted), scores)
        active = ordered

        results: dict[str, dict] = {}
        for cell, blockers in cycle_blocked.items():
            record = self._new_result(cell)
            record["status"] = STATUS_SKIPPED
            record["blocked_by"] = blockers
            record["reason"] = "调度器无法建立安全处理顺序（依赖循环），本次禁止自动修复"
            results[cell] = record
        for cell in deferred_budget:
            record = self._new_result(cell)
            record["status"] = STATUS_DEFERRED_BUDGET
            record["reason"] = "超过本次 max_candidates 预算，未进入 Agent 调查阶段"
            results[cell] = record
        for cell, prereqs in deferred_dep.items():
            record = self._new_result(cell)
            record["status"] = STATUS_DEFERRED_BUDGET_DEP
            record["reason"] = "目标及其 hard prerequisites 超过本次 max_candidates 预算"
            record["required_prerequisites"] = [
                self._prerequisite_entry(prereq) for prereq in prereqs
            ]
            results[cell] = record
        for cell in sorted(inactive_prereqs):
            record = self._new_result(cell)
            record["status"] = STATUS_DEFERRED_INACTIVE
            record["reason"] = (
                "低于 seed 阈值且未确认为修复前置（non-blocking），"
                "不阻塞下游，未进入 Agent 调查阶段"
            )
            results[cell] = record

        eligible = sorted(admitted)
        return {
            "seed_candidates": seeds,
            "eligible_candidates": eligible,
            "active_candidates": active,
            "pending_candidates": list(active),
            "current_candidate": active[0] if active else None,
            "candidate_results": results,
            "status": "planned" if active else "no_eligible_candidates",
        }

    def _route_next_candidate(self, state: SheetGuardState) -> str:
        """pending 队列非空则处理下一候选，否则进入 finalize。"""
        return "process" if state.get("current_candidate") else "finalize"

    def _cycle_blocked_candidates(
        self, graph: DependencyGraph, universe_cells: set[str]
    ) -> dict[str, list[dict]]:
        """识别 working graph 中因依赖循环被阻塞的候选及其 blocker 记录。

        返回 {candidate: [blocker entries]}。SCC 成员记录包含自身在内的
        循环候选集合；依赖该 SCC 的候选级下游记录同一集合（设计 5.1/5.4）。
        """
        blocked: dict[str, list[dict]] = {}
        for scc in graph.cycle_sccs():
            members = set(scc) & universe_cells
            if members:
                blocker = {"targets": sorted(members), "cause": "dependency_cycle"}
                affected = set(members)
            else:
                # 循环完全由非候选公式构成：记下成员地址供审计。
                blocker = {"targets": sorted(scc), "cause": "dependency_cycle"}
                affected = set()
            affected |= {
                cell
                for member in scc
                for cell in graph.descendants_of(member)
                if cell in universe_cells
            }
            for cell in sorted(affected):
                blocked.setdefault(cell, []).append(dict(blocker))
        return blocked

    def _resolve_prerequisites(
        self,
        graph: DependencyGraph,
        target: str,
        universe: set[str],
    ) -> tuple[set[str], set[str]]:
        """懒前置解析（设计 5.3/5.4）：返回 (hard, weak)。

        只在准备接纳目标时才沿其候选上游逐个分类：HARD 计入前置闭包并
        继续向上传递解析（传递 hard closure）；VERIFY 未通过廉价验证的与
        NON_BLOCKING 记入 weak——弱证据上游不阻塞高置信下游。
        """
        hard: set[str] = set()
        weak: set[str] = set()
        seen: set[str] = set()
        queue = [target]
        index = graph.index
        while queue:
            cell = queue.pop()
            for upstream in sorted(graph.ancestors_of(cell) & universe):
                if upstream in seen:
                    continue
                seen.add(upstream)
                decision = self._classify_repair_prerequisite(upstream)
                if decision == "HARD":
                    hard.add(upstream)
                    queue.append(upstream)
                elif decision == "VERIFY":
                    if self._cheap_verify(upstream, index):
                        hard.add(upstream)
                        queue.append(upstream)
                    else:
                        weak.add(upstream)
                else:
                    weak.add(upstream)
        return hard, weak

    def _new_result(self, cell: str) -> dict:
        """创建一条候选结果记录（设计 6 的 single source of truth）。"""
        source = self._anomaly_map.get(cell) or {}
        return {
            "target": cell,
            "status": None,
            "score": float(source.get("score") or 0.0),
            "signals": list(source.get("signals") or []),
            "evidence": [],
            "hypothesis": None,
            "attempts": [],
            "scheduler_events": [],
            "verification": None,
            "blocked_by": [],
            "required_prerequisites": [],
            "reason": None,
        }

    # ------------------------------------------------------------------
    # candidate confidence 与修复前置分类（设计 5.2/5.5）
    # ------------------------------------------------------------------

    def _candidate_confidence(self, cell: str, scores: dict[str, float]) -> float:
        """candidate confidence（设计 5.2）。

        依据至少包括：异常总分、signals 类型强度、独立信号数量、多视角
        共识置信度。总分可能掩盖证据质量——单一弱信号的高分不高于多个
        独立信号共同支持的中分（设计 5.2 反例），因此仅弱信号时不用总分。
        """
        info = self._anomaly_map.get(cell) or {}
        signals = list(dict.fromkeys(info.get("signals") or []))
        score = float(scores.get(cell) or 0.0)
        type_strength = max(
            (_SIGNAL_TYPE_WEIGHT.get(s, 0.3) for s in signals), default=0.0
        )
        consensus = info.get("consensus_confidence")
        consensus_strength = (
            type_strength if consensus is None else max(type_strength, float(consensus))
        )
        if any(s in _STRONG_PREREQ_SIGNALS for s in signals):
            base = max(score, consensus_strength)
        else:
            base = consensus_strength
        independent = len(signals)
        mutual = 0.05 * (independent - 1) if independent > 1 else 0.0
        return min(1.0, base + mutual)

    def _classify_repair_prerequisite(self, upstream: str) -> str:
        """修复前置分类（设计 5.5）：HARD / NON_BLOCKING / VERIFY。

        公式依赖 ≠ 修复前置：只有强证据（直接结构矛盾、结构性错误、
        多检测器互证）才升级为 hard prerequisite；单一弱依赖信号待廉价
        结构验证确认；仅弱信号不阻塞下游。
        """
        info = self._anomaly_map.get(upstream) or {}
        signals = set(info.get("signals") or [])
        if signals & _STRONG_PREREQ_SIGNALS:
            return "HARD"
        if {"dependency_anomaly", "neighbor_mismatch"} <= signals:
            return "HARD"
        if "dependency_anomaly" in signals:
            return "VERIFY"
        return "NON_BLOCKING"

    def _cheap_verify(self, upstream: str, index: WorkbookIndex) -> bool:
        """廉价结构验证（verifier-lite，无 LLM，设计 5.5）。

        单一 dependency_anomaly 只有在"多视角共识置信度足够强且模板同时
        偏离（severity 1.0）"时才确认为真错误；否则不阻塞高置信下游。
        """
        info = index.cell(upstream)
        if info is None or not info.is_formula:
            return False
        deviation, expectation = dependency_judgement(index, info)
        if deviation is None or not deviation.is_dependency_outlier:
            return False
        if deviation.severity < 1.0:
            return False
        if expectation is None or not expectation.has_consensus:
            return False
        return expectation.confidence >= 0.8

    # ------------------------------------------------------------------
    # 图节点：process_candidate（单候选事务）
    # ------------------------------------------------------------------

    def _process_candidate_node(self, state: SheetGuardState) -> dict:
        """处理当前候选：调查确认 → 诊断 → 修复 → 验证 → 提交/丢弃。

        每个候选的模型对话相互隔离（设计 8）；修复尝试受 max_attempts
        有界约束；纯依赖重排不消耗 attempts。
        """
        cand = state.get("current_candidate")
        results = {
            cell: _clone_record(record)
            for cell, record in (state.get("candidate_results") or {}).items()
        }
        pending = list(state.get("pending_candidates") or [])
        rec = results.get(cand) or self._new_result(cand)

        with self._observation(
            f"candidate {cand}",
            as_type="agent",
            input={
                "target": cand,
                "score": rec.get("score"),
                "signals": rec.get("signals"),
            },
        ) as candidate_span:
            pending = self._run_candidate_transaction(state, cand, rec, pending, results)
            self._safe_update(candidate_span, output=self._candidate_summary(rec))
        results[cand] = rec
        update: dict[str, Any] = {
            "candidate_results": results,
            "pending_candidates": pending,
            "current_attempt": len(rec.get("attempts") or []),
            "evidence": rec.get("evidence") or [],
            "hypothesis": rec.get("hypothesis"),
            "proposed_patch": rec.get("proposed_patch"),
            "verification": rec.get("verification"),
            "attempt_path": None,
            "status": "candidate_processed",
        }
        if self._refreshed_working_index is not None:
            # 成功 commit 后把新 working 索引写回状态，供审计与后续节点使用。
            update["working_index"] = self._refreshed_working_index.to_dict()
            update["working_revision"] = self._patcher.working_revision
            self._refreshed_working_index = None
        return update

    def _run_candidate_transaction(
        self,
        state: SheetGuardState,
        cand: str,
        rec: dict,
        pending: list[str],
        results: dict,
    ) -> list[str]:
        """执行单个候选的调查/诊断/修复/验证事务，返回更新后的 pending。"""
        # —— investigate / confirm ——（每候选独立对话，不读前一候选消息）
        verdict, evidence, raw = self._investigate(state, cand)
        rec["evidence"] = evidence
        rec["investigated"] = True
        pending = [p for p in pending if p != cand]
        if verdict == STATUS_DISMISSED:
            rec["status"] = STATUS_DISMISSED
            rec["reason"] = "调查确认该静态候选实际不存在公式错误"
            return pending
        if verdict != CONFIRMED:
            rec["status"] = STATUS_UNRESOLVED
            rec["reason"] = (
                f"调查证据不足或模型输出无效，无法确认该候选是否异常；"
                f"模型原始回复: {raw[:200]!r}"
                if raw
                else "调查证据不足，无法确认该候选是否异常"
            )
            return self._propagate_block(pending, rec, results)

        # —— diagnose ——
        hypothesis = self._diagnose(evidence, cand)
        rec["hypothesis"] = hypothesis
        if hypothesis is None:
            rec["status"] = STATUS_UNRESOLVED
            rec["reason"] = "诊断输出无效（模型输出异常），不修改工作副本"
            return self._propagate_block(pending, rec, results)

        # —— repair / attempt 有界循环（max_attempts 含首次尝试）——
        # 被否公式集合 cand 在循环内不变，提到循环外只算一次。
        rejected_formulas = set(
            (self._review_rejected.get(cand) or {}).get("rejected_formulas") or []
        )
        # 当前公式并入已知集合：原样复投当前公式是 no-op"修复"，验证器
        # 对无修改提案全绿，不早停会把空转提交记成 fixed（09-16 探针）。
        working_info = self._runtime_working_index(state).cell(cand)
        current_formula = (
            working_info.formula
            if working_info is not None and working_info.is_formula
            else None
        )
        while len(rec["attempts"]) < self.max_attempts:
            known_formulas = (
                {a.get("formula") for a in rec["attempts"]}
                | rejected_formulas
                | {current_formula}
            ) - {None}
            patch = self._repair(state, cand, rec)
            if patch is None:
                # proposal 无效（模型输出异常/非公式）→ 消耗 1 次 attempt。
                rec["attempts"].append({
                    "formula": None,
                    "outcome": "invalid_proposal",
                    "verification": {"passed": False, "notes": {}},
                })
                continue
            if patch["new_formula"] in known_formulas:
                # 早停：验证器是确定性的，重复提交此前失败的公式必败，
                # 原样复投当前公式则是不产生任何修改的空转。
                resubmitted_current = patch["new_formula"] == current_formula
                rec["status"] = STATUS_FAILED
                rec["reason"] = (
                    f"模型原样复投当前公式 {patch['new_formula']}，无新修复"
                    if resubmitted_current
                    else f"模型重复提交此前未通过的公式 {patch['new_formula']}"
                )
                return self._propagate_block(pending, rec, results)

            attempt_path = self._patcher.create_attempt()
            self._patcher.apply(cand, patch["new_formula"])
            patch["attempt_path"] = str(attempt_path)
            rec["proposed_patch"] = patch

            # —— proposal syntax precheck（验收 7）——
            syntax_ok, syntax_note = self._verifier.check_proposal_syntax(
                patch["new_formula"], cand
            )
            if not syntax_ok:
                self._patcher.discard_attempt()
                rec["attempts"].append({
                    "formula": patch["new_formula"],
                    "outcome": "syntax_invalid",
                    "verification": {"passed": False, "notes": {"syntax": syntax_note}},
                })
                continue

            # —— tentative graph + pre-commit dependency check（设计 5.2）——
            outcome, payload = self._precommit_dependency_check(state, cand, attempt_path)
            if outcome == "proposal_cycle":
                self._patcher.discard_attempt()
                rec["attempts"].append({
                    "formula": patch["new_formula"],
                    "outcome": "proposal_introduced_cycle",
                    "verification": {"passed": False, "notes": {"dependency": payload}},
                })
                continue
            if outcome == "reschedule":
                # 纯调度重排：不消耗 max_attempts，记录 scheduler event（设计 4）。
                self._patcher.discard_attempt()
                upstreams = payload["upstreams"]
                rec["scheduler_events"].append({
                    "event": "reschedule_pending_upstream",
                    "upstreams": upstreams,
                    "working_revision": self._patcher.working_revision,
                })
                if self._reschedule_stalled(cand, upstreams):
                    # 同一依赖条件下重复重排且 working graph 无进展 → 无进展。
                    rec["status"] = STATUS_UNRESOLVED
                    rec["reason"] = "依赖重排反复触发且工作副本无进展，视为 scheduler 无进展"
                    return self._propagate_block(pending, rec, results)
                self._log_reschedule(cand, upstreams)
                # 累积本次发现的候选级上游边，并对整个 pending 做依赖拓扑
                # 重排。不能只把 cand 插到直接上游之后：链式依赖（A→B→C）
                # 中后段上游（B→C）的重排会把先排好的依赖者（A）重新甩到
                # 自己前面，同版本重复触发 stall 保险丝。
                # order_candidates 的 extra_edges 即为此设计（设计 5.1 规则 4）。
                self._scheduler_extra_edges.setdefault(cand, set()).update(upstreams)
                pending = self._tentative_graph.order_candidates(
                    [p for p in pending if p != cand] + [cand],
                    {
                        cell: float(item.get("score") or 0.0)
                        for cell, item in self._anomaly_map.items()
                    },
                    extra_edges=self._scheduler_extra_edges,
                )
                return pending
            if outcome == "skip":
                # failed/unresolved/cycle 上游 → skipped，记录根 blocker（设计 5.4）。
                self._patcher.discard_attempt()
                rec["status"] = STATUS_SKIPPED
                rec["blocked_by"] = payload
                rec["reason"] = "依赖的上游候选状态不安全，本次禁止自动修复"
                return self._propagate_block(pending, rec, results)
            if outcome == "defer_budget":
                self._patcher.discard_attempt()
                rec["status"] = STATUS_DEFERRED_BUDGET_DEP
                rec["required_prerequisites"] = payload
                rec["reason"] = "repair proposal 新发现预算外静态候选上游，本次不能提交"
                return pending
            if outcome == "defer_inactive":
                self._patcher.discard_attempt()
                rec["status"] = STATUS_DEFERRED_INACTIVE
                rec["required_prerequisites"] = payload
                rec["reason"] = (
                    "repair proposal 新发现静态候选级上游，但该上游在 batch 冻结时"
                    "未进入 eligible/active；下一次运行需先纳入该 prerequisite"
                )
                return pending

            # —— verify ——（依赖顺序安全，attempt 进入完整 Verifier）
            verification = self._verify_attempt(state, cand)
            rec["verification"] = verification.to_dict()
            if verification.passed:
                self._patcher.commit_attempt()
                rec["attempts"].append({
                    "formula": patch["new_formula"],
                    "outcome": "verify_passed",
                    "verification": verification.to_dict(),
                })
                rec["status"] = STATUS_FIXED
                rec["reason"] = None
                # 成功 commit 后刷新 working_index/依赖图/调查工具（设计 7.2）。
                self._refresh_after_commit(state)
                pending = [p for p in pending if p != cand]
                graph = self._working_graph
                if graph is not None and pending:
                    pending = graph.order_candidates(
                        pending,
                        {
                            cell: float(item.get("score") or 0.0)
                            for cell, item in self._anomaly_map.items()
                        },
                    )
                return pending
            rec["attempts"].append({
                "formula": patch["new_formula"],
                "outcome": "verify_failed",
                "verification": verification.to_dict(),
            })
            # attempt 未通过 → 继续下一轮 repair（上限由 while 条件保证）。

        # attempts 用尽仍未通过 → failed（有界，验收 34）。
        rec["status"] = STATUS_FAILED
        rec["reason"] = f"修复尝试已达上限 {self.max_attempts} 次，最终验证未通过"
        return self._propagate_block(pending, rec, results)

    # ------------------------------------------------------------------
    # 调查 / 诊断 / 修复提议
    # ------------------------------------------------------------------

    def _investigate(self, state: SheetGuardState, cand: str):
        """调查并确认当前候选是否真正异常，返回 (verdict, evidence, raw)。

        批处理模式取消"回复不可用时回退第一个候选"的策略：模型没有
        成功完成确认时统一进入 unresolved（设计 8），不能把它自动当作
        真实错误继续修复。
        """
        index = self._runtime_working_index(state)
        anomaly = self._anomaly_map.get(cand) or {}
        signals = ", ".join(anomaly.get("signals") or []) or "无"
        remark = (self._review_rejected.get(cand) or {}).get("remark")
        prompt = INVESTIGATE_PROMPT.format(
            target=cand,
            score=f"{float(anomaly.get('score') or 0.0):.2f}",
            signals=signals,
            sheets=", ".join(sheet.name for sheet in index.sheets),
            formula_count=index.formula_count(),
            cross_sheet_count=index.cross_sheet_ref_count(),
            max_rounds=self.max_localizer_iterations,
            max_tool_calls=self.max_localizer_tool_calls,
            user_remark=remark or "无",
        )
        # conversation 是本候选独立的完整上下文；候选之间不共享消息。
        conversation: list[Any] = [HumanMessage(content=prompt)]
        final_text = ""
        tool_calls = 0
        try:
            # 只有调查阶段需要工具调用；诊断和修复阶段使用普通模型调用。
            bound_model = self.model.bind_tools(self._tools)
            for _ in range(self.max_localizer_iterations):
                response = bound_model.invoke(conversation)
                conversation.append(response)
                if not isinstance(response, AIMessage):
                    final_text = _message_text(response)
                    break
                if not response.tool_calls:
                    final_text = _message_text(response)
                    break
                # 一个 AIMessage 可能包含多个工具调用，逐个执行并累计预算。
                for call in response.tool_calls:
                    if tool_calls >= self.max_localizer_tool_calls:
                        # 预算耗尽时不执行调用，但必须补 ToolMessage 应答：
                        # AIMessage 携带的 tool_calls 若无应答，消息序列非法，
                        # 后续强制结论请求会被 API 以 400 拒绝。
                        conversation.append(ToolMessage(
                            content=(
                                "调查预算已用尽，此调用未执行；"
                                "请基于已有信息直接输出 JSON 结论。"
                            ),
                            tool_call_id=call["id"],
                        ))
                        continue
                    tool_calls += 1
                    tool = self._tool_map.get(call.get("name"))
                    if tool is None:
                        observation = f"Unknown tool: {call.get('name')}"
                    else:
                        try:
                            observation = str(tool.invoke(call.get("args", {})))
                        except Exception as exc:
                            observation = f"Tool error: {type(exc).__name__}: {exc}"
                    conversation.append(
                        ToolMessage(content=observation, tool_call_id=call["id"])
                    )
                if tool_calls >= self.max_localizer_tool_calls:
                    break
        except Exception:
            final_text = ""

        parsed = _parse_json_response(final_text)
        if "error" in parsed:
            # 模型可能把预算花完仍没有文本结论；去掉工具追问一次。
            try:
                response = self.model.invoke([
                    *conversation,
                    HumanMessage(
                        content="调查结束，不要再调用工具，直接按要求输出 JSON 结论。"
                    ),
                ])
                final_text = _message_text(response)
                parsed = _parse_json_response(final_text)
            except Exception:
                pass
        evidence = self._as_evidence(parsed)
        verdict = parsed.get("verdict")
        if verdict == CONFIRMED:
            return CONFIRMED, evidence, final_text
        if verdict == STATUS_DISMISSED:
            return STATUS_DISMISSED, evidence, final_text
        # verdict 无效或缺失（格式解析失败、空回复、预算耗尽）→ unresolved。
        return STATUS_UNRESOLVED, evidence, final_text

    @staticmethod
    def _as_evidence(parsed: dict) -> list:
        """把模型结论里的 evidence 统一为列表。"""
        evidence = parsed.get("evidence", [])
        return evidence if isinstance(evidence, list) else []

    def _diagnose(self, evidence: list, cand: str) -> dict | None:
        """根据候选调查证据生成错误假设；输出无效时返回 None。"""
        remark = (self._review_rejected.get(cand) or {}).get("remark")
        prompt = DIAGNOSE_PROMPT.format(
            target=cand,
            evidence=json.dumps(evidence),
            user_remark=remark or "无",
        )
        try:
            response = self.model.invoke(prompt)
            hypothesis = _parse_json_response(_message_text(response))
        except Exception:
            return None
        if "error" in hypothesis:
            return None
        hypothesis["target"] = cand
        return hypothesis

    def _repair(self, state: SheetGuardState, cand: str, rec: dict) -> dict | None:
        """根据诊断假设提出新公式；不修改任何工作簿。

        历史尝试与验证反馈来自候选自身 attempts（candidate_results 是
        唯一事实源，不再使用全局 repair_history）。
        """
        info = self._runtime_working_index(state).cell(cand)
        if info is None:
            return None
        current_formula = info.formula if info.is_formula else str(info.value)
        # 值反推线索（P1-3）：missing_formula 目标是数值常量时，确定性
        # 枚举可平移聚合范围上的聚合函数值。每候选只算一次（attempts
        # 之间工作簿不变，重复求值纯浪费）。
        if "value_hints" not in rec:
            rec["value_hints"] = aggregate_value_hints(
                self._runtime_working_index(state),
                self._runtime_working_graph(state),
                cand,
            )
        constant = (
            float(info.value)
            if isinstance(info.value, (int, float)) and not isinstance(info.value, bool)
            else None
        )
        value_hints = format_value_hints(rec["value_hints"], constant)
        attempts = rec.get("attempts") or []
        if attempts:
            feedback = json.dumps(
                [
                    {
                        "上次公式": attempt.get("formula"),
                        "未通过的检查": (attempt.get("verification") or {}).get("notes", {}),
                    }
                    for attempt in attempts
                ],
                ensure_ascii=False,
            )
        else:
            feedback = "无（首次修复）"
        family_context = StaticInspector(
            self._runtime_working_index(state)
        ).family_context(cand)
        prompt = REPAIR_PROMPT.format(
            target=cand,
            current_formula=current_formula,
            hypothesis=json.dumps(rec.get("hypothesis") or {}),
            feedback=feedback,
            family_context=json.dumps(family_context, ensure_ascii=False),
            value_hints=value_hints,
        )
        try:
            response = self.model.invoke(prompt)
            patch = _parse_json_response(_message_text(response))
        except Exception:
            return None
        new_formula = patch.get("new_formula")
        if not isinstance(new_formula, str) or not new_formula.startswith("="):
            return None
        patch["target"] = cand
        patch["old_formula"] = current_formula
        return patch

    # ------------------------------------------------------------------
    # 事务：tentative graph / pre-commit 检查 / verify / commit
    # ------------------------------------------------------------------

    def _precommit_dependency_check(
        self, state: SheetGuardState, cand: str, attempt_path: str | Path
    ) -> tuple[str, Any]:
        """commit 前基于 tentative dependency graph 检查候选级上游状态。

        先区分 working graph 既有 cycle 与 proposal 新引入 cycle（设计 5.2），
        再按上游候选当前状态路由：safe / reschedule / skip / defer。
        """
        graph = self._runtime_working_graph(state)
        # attempt 工作簿先解析并构建 tentative graph，再分析候选级上游。
        tentative_index = parse_workbook(str(attempt_path))
        tentative_graph = DependencyGraph(tentative_index)
        tentative_graph.build()
        self._tentative_graph = tentative_graph
        # 1) proposal 新引入 cycle：working graph 中不存在 → 可重试的 proposal 失败。
        working_sccs = set(graph.cycle_sccs())
        for scc in tentative_graph.cycle_sccs():
            if scc not in working_sccs:
                return "proposal_cycle", (
                    f"修复 proposal 新引入循环依赖: {', '.join(sorted(scc)[:3])}"
                )
        # 2) tentative graph 中能到达当前候选的静态候选级上游。
        upstream = tentative_graph.ancestors_of(cand) & set(self._anomaly_map)
        pending = list(state.get("pending_candidates") or [])
        results = state.get("candidate_results") or {}
        bad: list[dict] = []
        budget_prereq: list[dict] = []
        inactive_prereq: list[dict] = []
        pending_upstream: list[str] = []
        for up in sorted(upstream):
            if up == cand:
                continue
            if up in pending:
                # 仍在 active pending 的上游：纯调度重排，不消耗 attempts。
                pending_upstream.append(up)
                continue
            status = (results.get(up) or {}).get("status")
            if status in {STATUS_FIXED, STATUS_DISMISSED}:
                continue  # 已修复/已确认正常：允许继续验证。
            if status == STATUS_FAILED:
                bad.append({"target": up, "cause": "failed"})
            elif status == STATUS_UNRESOLVED:
                bad.append({"target": up, "cause": "unresolved"})
            elif status == STATUS_SKIPPED:
                # 记录根 blocker，不记录中间已被跳过的候选（设计 5.4）。
                bad.extend((results.get(up) or {}).get("blocked_by") or [])
            elif status in {STATUS_DEFERRED_BUDGET, STATUS_DEFERRED_BUDGET_DEP}:
                budget_prereq.append(self._prerequisite_entry(up))
            else:
                # status 为 None：完整 universe 中被 threshold 过滤、未进入
                # eligible 的静态候选（依赖在初始图未知，验收 32）。
                inactive_prereq.append(self._prerequisite_entry(up))
        if bad:
            return "skip", bad
        if budget_prereq:
            return "defer_budget", budget_prereq
        if inactive_prereq:
            return "defer_inactive", inactive_prereq
        if pending_upstream:
            return "reschedule", {"upstreams": pending_upstream}
        return "safe", None

    def _prerequisite_entry(self, cell: str) -> dict:
        """构造 required_prerequisites 记录（地址 + 静态分 + 信号）。"""
        source = self._anomaly_map.get(cell) or {}
        return {
            "target": cell,
            "score": float(source.get("score") or 0.0),
            "signals": list(source.get("signals") or []),
        }

    def _verify_attempt(self, state: SheetGuardState, cand: str):
        """对当前 attempt 执行完整六项验证（三态检查，设计 7.2）。

        回归检查的授权修改范围 = 已成功提交目标 ∪ 当前候选（验收 19）；
        值保持检查在 tentative graph 中存在已修复上游路径时记为
        not_applicable（验收 20/21）。
        """
        fixed = set(self._fixed_targets(state))
        return self._verifier.verify(
            self._runtime_source_index(state),
            str(self._patcher.attempt_path),
            cand,
            working_index=self._runtime_working_index(state),
            fixed_targets=fixed,
            tentative_graph=self._tentative_graph,
        )

    def _fixed_targets(self, state: SheetGuardState) -> list[str]:
        """从 candidate_results 派生已成功提交的目标集合。"""
        results = state.get("candidate_results") or {}
        return [
            cell
            for cell, record in results.items()
            if record.get("status") == STATUS_FIXED
        ]

    def _refresh_after_commit(self, state: SheetGuardState) -> None:
        """成功 commit 后刷新 working_index / 依赖图 / 调查工具（设计 7.2）。"""
        working = parse_workbook(str(self._patcher.working_path))
        self._working_index = working
        self._refreshed_working_index = working  # 由 _process_candidate_node 写回状态
        graph = DependencyGraph(working)
        graph.build()
        self._working_graph = graph
        self._rebuild_investigation_tools(state)

    def _reschedule_stalled(self, cand: str, upstreams: list[str]) -> bool:
        """同一 (candidate, upstream, working_revision) 组合是否已触发过重排。"""
        revision = self._patcher.working_revision
        return any(
            self._reschedule_log.get((cand, up)) == revision for up in upstreams
        )

    def _log_reschedule(self, cand: str, upstreams: list[str]) -> None:
        for up in upstreams:
            self._reschedule_log[(cand, up)] = self._patcher.working_revision

    # ------------------------------------------------------------------
    # 图节点：advance / 失败传播 / finalize
    # ------------------------------------------------------------------

    def _advance_node(self, state: SheetGuardState) -> dict:
        """把下一 pending 候选设为 current；队列空时 current 置空 → finalize。

        重排（reschedule）后的当前候选仍在队列中且排在被提升的上游之后，
        advance 自然先处理上游候选。
        """
        pending = list(state.get("pending_candidates") or [])
        next_candidate = pending[0] if pending else None
        attempt_count = 0
        if next_candidate:
            attempt_count = len(
                (state.get("candidate_results") or {}).get(next_candidate, {}).get("attempts")
                or []
            )
        return {
            "current_candidate": next_candidate,
            "current_attempt": attempt_count,
            "status": "candidate_advanced" if next_candidate else "finalizing",
        }

    def _propagate_block(
        self, pending: list[str], origin: dict, results: dict
    ) -> list[str]:
        """把 origin 的失败/未解决/跳过状态传播给仍待处理的直接与间接下游。

        blocked_by 记录根 blocker（origin 自身或其根因），不记录中间
        已经被跳过的候选（设计 5.4）。
        """
        graph = self._working_graph
        if graph is None:
            return pending
        origin_target = origin["target"]
        status = origin["status"]
        if status == STATUS_SKIPPED:
            cause_entries = list(origin.get("blocked_by") or [])
            reason = "上游候选被依赖阻塞，本次禁止自动修复"
        else:
            cause_entries = [{"target": origin_target, "cause": status}]
            reason = (
                "上游单元格修复失败，下游验证结果不可靠"
                if status == STATUS_FAILED
                else "上游候选状态未解决，下游自动修复被阻塞"
            )
        for p in list(pending):
            if graph.reachable(origin_target, p):
                record = results.get(p) or self._new_result(p)
                record["status"] = STATUS_SKIPPED
                record["blocked_by"] = _merge_blockers(record["blocked_by"], cause_entries)
                record["reason"] = reason
                results[p] = record
                pending = [x for x in pending if x != p]
        return pending

    def _finalize_node(self, state: SheetGuardState) -> dict:
        """计算最终批处理状态并导出修复工作簿（如指定路径）。"""
        export_path: str | None = None
        if self._patcher is not None and self._export_target is not None:
            try:
                exported = self._patcher.export(self._export_target)
                export_path = str(exported)
            except Exception as exc:
                logger.warning("repaired workbook export failed: %s", exc)
        return {
            "status": self._final_status(state),
            "repaired_workbook": export_path,
        }

    def _final_status(self, state: SheetGuardState) -> str:
        """从 candidate_results 派生最终批处理状态（设计 6）。"""
        if state.get("status") in {"no_candidates", "no_eligible_candidates"}:
            return state["status"]
        statuses = [
            record.get("status")
            for record in (state.get("candidate_results") or {}).values()
        ]
        fixed_n = statuses.count(STATUS_FIXED)
        failed_n = statuses.count(STATUS_FAILED)
        unresolved_n = statuses.count(STATUS_UNRESOLVED)
        skipped_n = statuses.count(STATUS_SKIPPED)
        deferred_n = sum(1 for status in statuses if status in DEFERRED_STATUSES)
        if fixed_n == 0 and failed_n == 0 and unresolved_n == 0 and skipped_n == 0:
            if deferred_n:
                # 仍有 deferred：已知 eligible candidate 未完成，不能算成功。
                return "incomplete_due_to_budget"
            return "completed_without_repairs"
        if fixed_n == 0:
            return "failed"
        if failed_n or unresolved_n or skipped_n or deferred_n:
            return "partial_success"
        return "success"

    def cleanup(self) -> None:
        """删除当前活动的工作簿事务副本（如果存在）；可重复调用。"""
        if self._patcher is not None:
            self._patcher.cleanup()
            self._patcher = None

    def _candidate_summary(self, rec: dict) -> dict:
        """候选级 Langfuse observation 的输出摘要（设计 11）。"""
        attempts = rec.get("attempts") or []
        last = attempts[-1] if attempts else {}
        return {
            "target": rec.get("target"),
            "status": rec.get("status"),
            "attempts": len(attempts),
            "scheduler_events": rec.get("scheduler_events") or [],
            "final_formula": last.get("formula") if rec.get("status") == STATUS_FIXED else None,
            "verification_passed": bool((rec.get("verification") or {}).get("passed")),
            "blocked_by": rec.get("blocked_by") or [],
            "required_prerequisites": rec.get("required_prerequisites") or [],
            "reason": rec.get("reason"),
        }

    def _batch_summary(self, state: SheetGuardState) -> dict:
        """根 observation 的批次摘要输出（设计 11）。"""
        results = state.get("candidate_results") or {}
        statuses = [record.get("status") for record in results.values()]
        deferred_n = sum(1 for status in statuses if status in DEFERRED_STATUSES)
        inactive_n = sum(1 for status in statuses if status == STATUS_DEFERRED_INACTIVE)
        # 高置信 deferred 逐项明细进根 observation（设计 6.2/验收 6）。
        high_confidence_deferred = [
            {
                "target": record.get("target"),
                "status": record.get("status"),
                "score": record.get("score"),
                "signals": record.get("signals") or [],
                "confidence": _deferred_confidence(record),
                "required_prerequisites": record.get("required_prerequisites") or [],
                "reason": record.get("reason"),
            }
            for record in results.values()
            if record.get("status") in DEFERRED_STATUSES
            and _deferred_confidence(record) == DEFERRED_CONFIDENCE_HIGH
        ]
        return {
            "status": state.get("status"),
            "static_candidate_count": state.get("static_candidate_count", 0),
            "seed_candidate_count": len(state.get("seed_candidates") or []),
            "eligible_candidate_count": len(state.get("eligible_candidates") or []),
            "initial_active_candidate_count": len(state.get("active_candidates") or []),
            "processed_candidate_count": self._processed_count(results),
            "fixed_count": statuses.count(STATUS_FIXED),
            "failed_count": statuses.count(STATUS_FAILED),
            "dismissed_count": statuses.count(STATUS_DISMISSED),
            "unresolved_count": statuses.count(STATUS_UNRESOLVED),
            "skipped_count": statuses.count(STATUS_SKIPPED),
            "deferred_count": deferred_n,
            "inactive_prerequisite_deferred_count": inactive_n,
            "high_confidence_deferred": high_confidence_deferred,
            "high_confidence_deferred_count": len(high_confidence_deferred),
            "max_deferred_confidence": max(
                (
                    float(record.get("score") or 0.0)
                    for record in results.values()
                    if record.get("status") in DEFERRED_STATUSES
                ),
                default=0.0,
            ),
            "error": state.get("error"),
        }

    def _processed_count(self, results: dict) -> int:
        """实际进入 investigate/confirm 且未落入 deferred 的候选数（验收 35）。"""
        return sum(
            1
            for record in results.values()
            if record.get("investigated")
            and record.get("status") not in DEFERRED_STATUSES
        )

    def _build_audit(self, result: dict) -> dict:
        """批量审计记录（设计 9）：候选结果派生 + 兼容单错误字段。"""
        results = result.get("candidate_results") or {}
        fixed: list[dict] = []
        failed: list[dict] = []
        dismissed: list[dict] = []
        unresolved: list[dict] = []
        skipped: list[dict] = []
        deferred: list[dict] = []
        repair_history: list[dict] = []
        first_fixed_patch: dict | None = None
        for record in results.values():
            attempts = record.get("attempts") or []
            for attempt in attempts:
                repair_history.append({
                    "target": record.get("target"),
                    "formula": attempt.get("formula"),
                    "outcome": attempt.get("outcome"),
                    "verification": attempt.get("verification"),
                })
            status = record.get("status")
            if status == STATUS_FIXED:
                last_formula = attempts[-1].get("formula") if attempts else None
                patch = record.get("proposed_patch") or {}
                fixed.append({
                    "target": record.get("target"),
                    # 最后一次提案的 old_formula 即原公式：失败 attempt 均
                    # 被丢弃，working 在首次尝试前从未变更。
                    "old_formula": patch.get("old_formula"),
                    "new_formula": last_formula,
                    "score": record.get("score"),
                    "signals": record.get("signals") or [],
                    "hypothesis": record.get("hypothesis"),
                    "verification": record.get("verification"),
                })
                if first_fixed_patch is None:
                    first_fixed_patch = {
                        "target": record.get("target"),
                        "old_formula": patch.get("old_formula"),
                        "new_formula": last_formula,
                    }
            elif status == STATUS_FAILED:
                patch = record.get("proposed_patch") or {}
                failed.append({
                    "target": record.get("target"),
                    "old_formula": patch.get("old_formula"),
                    "last_formula": attempts[-1].get("formula") if attempts else None,
                    "score": record.get("score"),
                    "signals": record.get("signals") or [],
                    "hypothesis": record.get("hypothesis"),
                    "reason": record.get("reason"),
                    "verification": record.get("verification"),
                    "attempts": len(attempts),
                })
            elif status == STATUS_DISMISSED:
                # 审查需要上下文（v1.12）：带上调查时点的当前公式；
                # missing_formula 场景该格原本无公式 → None。
                dismissed_target = record.get("target")
                dismissed_info = self._runtime_working_index(result).cell(
                    dismissed_target
                )
                dismissed.append({
                    "target": dismissed_target,
                    "reason": record.get("reason"),
                    "formula": (
                        dismissed_info.formula
                        if dismissed_info is not None and dismissed_info.is_formula
                        else None
                    ),
                })
            elif status == STATUS_UNRESOLVED:
                unresolved.append(
                    {"target": record.get("target"), "reason": record.get("reason")}
                )
            elif status == STATUS_SKIPPED:
                skipped.append({
                    "target": record.get("target"),
                    "status": status,
                    "blocked_by": record.get("blocked_by") or [],
                    "reason": record.get("reason"),
                })
            elif status in DEFERRED_STATUSES:
                deferred.append({
                    "target": record.get("target"),
                    "status": status,
                    "score": record.get("score"),
                    "signals": record.get("signals") or [],
                    "confidence": _deferred_confidence(record),
                    "required_prerequisites": record.get("required_prerequisites") or [],
                    "reason": record.get("reason"),
                })
        blocking = failed or unresolved or skipped or deferred
        high_confidence_deferred = [
            entry for entry in deferred if entry["confidence"] == DEFERRED_CONFIDENCE_HIGH
        ]
        warning = (
            "部分候选失败、未解决、被阻塞或延期，请人工检查 failed、unresolved、"
            "skipped 和 deferred。"
            if blocking
            else None
        )
        return {
            "status": result.get("status"),
            # 最终生效配置随审计记录输出：跨实验对比时配置指纹完整。
            "config": self.config.to_dict(),
            "static_candidate_count": result.get("static_candidate_count", 0),
            "seed_candidate_count": len(result.get("seed_candidates") or []),
            "eligible_candidate_count": len(result.get("eligible_candidates") or []),
            "initial_active_candidate_count": len(result.get("active_candidates") or []),
            "processed_candidate_count": self._processed_count(results),
            "deferred_count": len(deferred),
            "inactive_prerequisite_deferred_count": sum(
                1 for entry in deferred if entry["status"] == STATUS_DEFERRED_INACTIVE
            ),
            "high_confidence_deferred_count": len(high_confidence_deferred),
            "max_deferred_confidence": max(
                (float(entry["score"] or 0.0) for entry in deferred), default=0.0,
            ),
            "fixed": fixed,
            "failed": failed,
            "dismissed": dismissed,
            "unresolved": unresolved,
            "skipped": skipped,
            "deferred": deferred,
            # 上一轮用户认证的格子原样透传：供报告渲染与下一轮累积。
            "certified": list(self._review_certified),
            "high_confidence_deferred": high_confidence_deferred,
            "repaired_workbook": result.get("repaired_workbook"),
            "warning": warning,
            # 兼容单错误字段：第一个成功修复候选的补丁信息（设计 9）。
            "proposed_patch": first_fixed_patch,
            "verification": fixed[0]["verification"] if fixed else None,
            "repair_history": repair_history,
            "anomalies": result.get("anomalies") or [],
            "localization_warning": None,
            "error": result.get("error"),
        }
