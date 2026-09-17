"""SheetGuard 结构化 Agent 各节点的提示词模板。

multi-cell 批处理模式下，候选已由调度器选定，原"从多个候选定位唯一
target"的提示改为当前候选的调查/确认：模型返回 confirmed / dismissed /
unresolved 三类结构化结论（设计 8）。
"""

INVESTIGATE_PROMPT = """你是一名电子表格调试专家。
静态检测已把候选单元格 {target} 交给本候选独立调查；请用调查工具核实
它是否真正存在公式错误，而不是继续寻找其他错误。

当前候选单元格：{target}
静态可疑分：{score}
静态信号：{signals}
用户反馈备注：{user_remark}

信号含义：
- pattern_anomaly：公式的相对引用模板偏离同家族主流写法；
- neighbor_mismatch：两侧邻居写法一致而本格不同；
- missing_formula：数字常量打断了原本连续的公式段（该格本该是公式）；
- range_boundary：SUM 范围尺寸与同家族不一致；
- dependency_anomaly：跨表引用集合与同家族主流不同。

工作簿结构：
- 工作表列表：{sheets}
- 公式单元格数量：{formula_count}
- 跨表公式数量：{cross_sheet_count}

调查预算：最多 {max_rounds} 轮对话、{max_tool_calls} 次工具调用。证据不足或无法判断时
必须如实在结论中说明，不要臆断。
调查结束后，最后一条消息必须直接输出 JSON 结论，不要以空消息或仅工具调用结尾：
```json
{{"target":"{target}","verdict":"confirmed","confidence":0.0,"evidence":["observation"],"reasoning":"why"}}
```
verdict 只能取以下三个值之一：
- "confirmed"：核实该候选确实存在公式错误，需要修复；
- "dismissed"：核实该候选实际正常，静态检测是误报；
- "unresolved"：证据不足或无法判断该候选是否异常（不修改工作簿）。
结论中的 evidence 与 reasoning 一律用中文表述；单元格坐标（如 Sheet!A1）保持原样。
"""

DIAGNOSE_PROMPT = """你负责诊断一个电子表格公式错误。

目标单元格：{target}
证据：{evidence}
用户反馈备注：{user_remark}

请从以下类型中选择一个 error_type：wrong_range、wrong_cell_reference、wrong_operator、
missing_formula、cross_sheet_error。
hypothesis 字段一律用中文表述（单元格坐标保持原样）。
只回复 JSON：
```json
{{"target":"Sheet!A1","error_type":"wrong_range","hypothesis":"why","confidence":0.0}}
```
"""

REPAIR_PROMPT = """你负责提出一个电子表格公式修复方案；不要修改任何文件。

目标单元格：{target}
当前公式/值：{current_formula}
诊断结论：{hypothesis}
历史尝试与验证反馈：{feedback}
同家族邻居公式：{family_context}
值反推线索：{value_hints}

如果历史尝试不为空，请避开其中未通过的公式及同类错误。
同家族邻居公式不为空时，新公式应与兄弟单元格保持一致的引用模式和跨表引用集合。
值反推线索不为"无"时，与原常量一致的聚合公式优先考虑；全部不匹配时按线索
末尾的指引考虑条件聚合。
reason 字段一律用中文表述（单元格坐标保持原样）。
只回复 JSON：
```json
{{"target":"Sheet!A1","new_formula":"=...","reason":"why","confidence":0.0}}
```
"""
