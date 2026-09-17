"""仅用于提案的有界 ReAct 基线，用于工作簿诊断。"""
from __future__ import annotations
import json
import re
import time
from typing import Any
from langchain_core.language_models import BaseChatModel
from langchain_core.messages import AIMessage, HumanMessage, ToolMessage
from sheetguard.spreadsheet.model import WorkbookIndex
from sheetguard.spreadsheet.dependency_graph import DependencyGraph
from sheetguard.tools.investigation import make_investigation_tools


SYSTEM_PROMPT = """You are a spreadsheet debugging agent. Find exactly one formula error.
Use the investigation tools to inspect formulas and dependencies. Do not modify
any workbook. Finish with JSON containing target, error_type, old_formula,
new_formula, confidence, and explanation. The new_formula must be the proposed
repair, not an applied mutation.
"""

# 即使单次模型响应包含大量工具调用，也要限制工具执行的总次数，
# 这与基线的默认模型轮次预算保持一致。
MAX_TOOL_CALLS = 15


class BaselineAgent:
    """单个有界工具调用 Agent，从不写入工作簿。"""

    def __init__(self, index: WorkbookIndex, graph: DependencyGraph, model: BaseChatModel):
        self.index = index
        self.graph = graph
        self.model = model
        self.tools = make_investigation_tools(index, graph)
        self._tool_map = {tool.name: tool for tool in self.tools}

    def run(
        self,
        workbook_path: str,
        max_iterations: int = 15,
        max_tool_calls: int = MAX_TOOL_CALLS,
    ) -> dict:
        """调查并提出修复方案，不修改 ``workbook_path`` 文件。

        ``max_tool_calls`` 是所有模型响应共用的总预算。它会被限制在 V1
        上限以内，防止单次响应让模型突破 Agent 的安全边界扩大工作量。
        """
        started = time.perf_counter()
        messages: list[Any] = [
            HumanMessage(
                content=(
                    f"{SYSTEM_PROMPT}\nWorkbook path: {workbook_path}\n"
                    "Analyze the workbook and return the repair proposal."
                )
            )
        ]
        bound_model = self.model.bind_tools(self.tools)
        tool_calls = 0
        tool_budget = max(0, min(max_tool_calls, MAX_TOOL_CALLS))
        tokens = 0
        final_text = ""
        budget_exhausted = False

        for _ in range(max_iterations):
            response = bound_model.invoke(messages)
            if not isinstance(response, AIMessage):
                final_text = getattr(response, "content", str(response))
                break
            messages.append(response)
            usage = response.usage_metadata or {}
            tokens += int(usage.get("total_tokens", 0) or 0)
            if not response.tool_calls:
                final_text = response.content if isinstance(response.content, str) else str(response.content)
                break

            for call in response.tool_calls:
                if tool_calls >= tool_budget:
                    budget_exhausted = True
                    final_text = ""
                    break
                tool_calls += 1
                tool = self._tool_map.get(call["name"])
                if tool is None:
                    content = f"Unknown tool: {call['name']}"
                else:
                    try:
                        content = str(tool.invoke(call.get("args", {})))
                    except Exception as exc:  # keep the agent loop bounded and observable
                        content = f"Tool error: {type(exc).__name__}: {exc}"
                messages.append(ToolMessage(content=content, tool_call_id=call["id"]))
            if budget_exhausted or tool_calls >= tool_budget:
                budget_exhausted = True
                final_text = ""
                break
        else:
            final_text = ""

        elapsed_ms = (time.perf_counter() - started) * 1000
        parsed = self._parse_json_response(final_text)
        parsed.update({"tool_calls": tool_calls, "tokens": tokens, "latency_ms": elapsed_ms})
        if budget_exhausted:
            # Never return a partial proposal as if it were successful.
            parsed.pop("target", None)
            parsed.pop("new_formula", None)
            parsed["error"] = "tool_call_budget_exceeded"
        elif not final_text:
            parsed.setdefault("error", "max_iterations_exceeded")
        return parsed

    @staticmethod
    def _parse_json_response(text: str) -> dict:
        """从带围栏或纯文本的模型响应中提取 JSON 提案。"""
        if not text:
            return {}
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
            except json.JSONDecodeError:
                continue
        return {"error": "Could not parse JSON response", "raw_response": text}
