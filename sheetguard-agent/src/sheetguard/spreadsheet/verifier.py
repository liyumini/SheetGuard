"""沙箱公式补丁的确定性验证。

本模块负责验证 Agent 提出的公式修复是否安全、合理且能够成功重算。
验证流程不依赖 LLM，而是执行六项固定检查：
1. 公式语法检查；2. 公式模式检查；3. 值保持检查；
4. 依赖循环检查；5. 未授权修改回归检查；6. 公式重新计算检查。

multi-cell 批处理模式（设计 7.2）明确区分三种基线语义：
- ``source_index``：原始工作簿索引，用于原始值/公式基线、未授权修改
  检测和值保持比较，整个运行期间不变；
- ``working_index``：当前工作副本索引（此前成功提交的修复），用于
  公式家族与模式分析；
- ``tentative_graph``：attempt 工作簿的依赖图，用于条件值保持判断
  和新引入循环检测。

每项检查输出三态结果：passed / failed / not_applicable；
``not_applicable`` 不参与 pass/fail 判定（第一版用于 multi-cell
missing_formula 的条件值保持）。
"""
from __future__ import annotations
from dataclasses import dataclass, field
import math
from pathlib import Path
from openpyxl.utils import get_column_letter
from sheetguard.spreadsheet.model import CellRef, WorkbookIndex
from sheetguard.spreadsheet.parser import parse_workbook
from sheetguard.spreadsheet.aggregates import (
    is_wellformed_aggregate,
    family_context_data,
    has_aggregate_function,
    normalize_template,
    structural_template,
)
from sheetguard.spreadsheet.dependency_graph import DependencyGraph
from sheetguard.spreadsheet.formula_parser import to_relative_template
from sheetguard.spreadsheet.recalc import (
    RecalcEngine,
    unsupported_functions,
    validate_formula,
)

# verification checks 的三种合法状态
CHECK_PASSED = "passed"
CHECK_FAILED = "failed"
CHECK_NOT_APPLICABLE = "not_applicable"


@dataclass
class VerificationResult:
    """保存一次验证的总体结果、各项检查结果和补充说明。

    checks 每项为三态结构：``{"status": "passed"}``、
    ``{"status": "failed", "reason": "..."}`` 或
    ``{"status": "not_applicable", "reason": "upstream_changed"}``。

    示例：
        VerificationResult(
            passed=True,
            checks={"syntax": {"status": "passed"}, "recalc": {"status": "passed"}},
        )
    """

    passed: bool = False  # 没有 failed 检查项（not_applicable 不算失败）
    checks: dict[str, dict] = field(default_factory=dict)  # 各项检查的三态结果
    notes: dict[str, str] = field(default_factory=dict)  # 未通过检查的补充说明

    def to_dict(self) -> dict:
        """将验证结果转换为普通字典，方便生成 JSON 报告。"""
        return {"passed": self.passed, "checks": self.checks, "notes": self.notes}


class Verifier:
    """对沙箱中的公式修复执行六层确定性验证。

    验证流程：
        1. syntax：所有公式是否符合支持的语法；
        2. pattern：目标公式是否与同排/列家族保持一致的相对模板；
        3. value：缺失公式修复后是否复现原数值常量（multi-cell 条件语义）；
        4. dependency：是否出现 working graph 之外新引入的循环依赖；
        5. regression：是否修改了不允许修改的单元格；
        6. recalc：工作簿是否能够成功重新计算。

    示例：
        result = Verifier().verify(
            source_index,
            "sheetguard_attempt.xlsx",
            "P&L!C6",
        )
        if result.passed:
            print("修复验证通过")
    """

    def verify(
        self,
        source_index: WorkbookIndex,
        attempt_path: str | Path,
        target_cell: str,
        *,
        working_index: WorkbookIndex | None = None,
        allowed_modifications: set[str] | list[str] | None = None,
        fixed_targets: set[str] | None = None,
        tentative_graph: DependencyGraph | None = None,
    ) -> VerificationResult:
        """执行全部验证并返回汇总结果。

        单目标兼容模式：只传前三个参数时，source_index 同时承担原始
        基线和家族分析职责，只允许修改 ``target_cell``，任何循环都失败。

        multi-cell 模式：``working_index`` 提供当前工作副本语义（家族
        分析），``allowed_modifications`` 缺省时由 ``fixed_targets ∪
        {target_cell}`` 构造授权修改范围，``tentative_graph`` 用于
        条件值保持与新引入循环检测。

        示例：目标单元格为 ``P&L!C6`` 且已成功修复 ``Revenue!B20`` 时，
        allowed_modifications = {"Revenue!B20", "P&L!C6"}。
        """
        # attempt 工作簿只解析一次，供各检查复用；解析失败时所有检查失败。
        parse_error: str | None = None
        attempt_index: WorkbookIndex | None = None
        attempt_graph: DependencyGraph | None = None
        try:
            attempt_index = parse_workbook(str(attempt_path))
            attempt_graph = DependencyGraph(attempt_index)
            attempt_graph.build()
        except Exception as exc:
            parse_error = f"{type(exc).__name__}: {exc}"[:200]

        family_index = working_index if working_index is not None else source_index
        # working graph：multi-cell 模式下用于区分既有循环与 proposal 新引入循环。
        working_graph: DependencyGraph | None = None
        if working_index is not None:
            working_graph = DependencyGraph(working_index)
            working_graph.build()
        # working 基线的 NaN 集合：multi-cell 批处理中其他尚未修复的错误格
        # 重算失败是工作簿既有事实，不该判死当前候选（complete_018 的
        # Revenue!L4 坏 VLOOKUP 曾把 P&L!J2 / Dashboard!N4 的正确提案全部
        # 毒死）。目标自身不入基线——提案换汤不换药仍算不出时必须失败。
        baseline_nan: set[str] = set()
        if working_index is not None:
            baseline_nan = {
                addr
                for addr, value in RecalcEngine(
                    working_index, working_graph
                ).evaluate_all().items()
                if isinstance(value, float) and math.isnan(value)
            }
            baseline_nan.discard(target_cell)
        allowed = self._resolve_allowed(allowed_modifications, fixed_targets, target_cell)
        # multi-cell 回归/值保持需要已提交修复的集合；排除当前目标自身
        # （它在本轮才被修改，不构成"上游已修复"语义）。
        prior_fixed = {addr for addr in (fixed_targets or set()) if addr != target_cell}
        # working graph 中已经存在的循环成员：它们的重算失败是既有事实，
        # 不归因于当前 attempt（设计约束 8/11：无关候选不因既有循环停止）。
        exempt_cells: set[str] = set()
        if working_index is not None:
            exempt_cells = self._working_cycle_members(working_index)

        outcomes: list[tuple[str, str, str]] = []
        if parse_error is not None:
            outcomes.append(("syntax", CHECK_FAILED, f"存在语法无效的公式: {parse_error}"))
            for name in ("pattern", "value", "dependency", "regression", "recalc"):
                outcomes.append((name, CHECK_FAILED, f"attempt 工作簿无法解析: {parse_error}"))
        else:
            assert attempt_index is not None and attempt_graph is not None
            outcomes.append(("syntax", *self._check_syntax(attempt_index)))
            outcomes.append((
                "pattern",
                *self._check_pattern(source_index, family_index, attempt_index, target_cell),
            ))
            outcomes.append((
                "value",
                *self._check_value_preserved(
                    source_index,
                    attempt_index,
                    attempt_graph,
                    target_cell,
                    tentative_graph,
                    prior_fixed_targets=prior_fixed,
                ),
            ))
            outcomes.append((
                "dependency",
                *self._check_dependency(attempt_graph, working_graph=working_graph),
            ))
            outcomes.append((
                "regression",
                *self._check_regression(source_index, attempt_index, allowed),
            ))
            outcomes.append((
                "recalc",
                *self._check_recalc(
                    attempt_index, attempt_graph,
                    exempt_cells=exempt_cells, baseline_nan=baseline_nan,
                ),
            ))

        checks = {name: {"status": status} for name, status, _ in outcomes}
        for name, status, reason in outcomes:
            if reason and status in {CHECK_FAILED, CHECK_NOT_APPLICABLE}:
                checks[name]["reason"] = reason
        # notes 只收录未通过的检查，随验证结果交给修复重试参考。
        notes = {
            name: reason
            for name, status, reason in outcomes
            if status == CHECK_FAILED and reason
        }
        passed = all(result["status"] != CHECK_FAILED for result in checks.values())
        return VerificationResult(passed=passed, checks=checks, notes=notes)

    def check_proposal_syntax(self, formula: str, target_cell: str) -> tuple[bool, str]:
        """proposal 语法预检（设计 5.2/验收 7）：进入 tentative graph 之前先验证。

        返回（是否通过，失败说明）。非法 proposal 直接丢弃并消耗一次
        max_attempts，不进入 attempt 副本和依赖图构建。函数白名单同样
        在这里把关：重算引擎不支持的函数（如 ROUND）到 recalc 才会变成
        NaN，反馈里只剩"无法重算"，模型无从知道错在哪类。
        """
        unsupported = unsupported_functions(formula)
        if unsupported:
            return False, (
                f"proposal 使用了重算引擎不支持的函数: {', '.join(unsupported)}"
            )
        try:
            validate_formula(formula, CellRef.parse(target_cell))
            return True, ""
        except Exception as exc:
            return False, f"proposal 语法无效: {type(exc).__name__}: {exc}"[:200]

    @staticmethod
    def _working_cycle_members(working_index: WorkbookIndex) -> set[str]:
        """计算 working graph 中参与循环的单元格集合。"""
        graph = DependencyGraph(working_index)
        graph.build()
        return graph.cycle_members()

    @staticmethod
    def _resolve_allowed(
        allowed_modifications: set[str] | list[str] | None,
        fixed_targets: set[str] | None,
        target_cell: str,
    ) -> set[str]:
        """解析授权修改范围。

        显式传入 ``allowed_modifications`` 时完全尊重调用方（旧单目标
        语义）；否则由"已成功提交的目标 ∪ 当前候选"动态构造（设计 7.2：
        回归检查允许累计的成功修改，同时拒绝修改其他单元格）。
        """
        if allowed_modifications is not None:
            return set(allowed_modifications) | {target_cell}
        allowed = set(fixed_targets or set())
        allowed.add(target_cell)
        return allowed

    def _check_syntax(self, attempt_index: WorkbookIndex) -> tuple[str, str]:
        """检查 attempt 中所有公式的语法是否有效，返回（状态，失败说明）。

        示例：``=A1+B1`` 可以通过，``=A1+`` 会失败。
        """
        try:
            for info in attempt_index.formula_cells():
                validate_formula(info.formula, info.ref)
            return CHECK_PASSED, ""
        except Exception as exc:
            return CHECK_FAILED, f"存在语法无效的公式: {type(exc).__name__}: {exc}"[:200]

    def _check_pattern(
        self,
        source_index: WorkbookIndex,
        family_index: WorkbookIndex,
        attempt_index: WorkbookIndex,
        target_cell: str,
    ) -> tuple[str, str]:
        """检查目标公式是否与同排/列家族的主流相对模板一致。

        ``family_index`` 承担家族分析语义（multi-cell 模式使用 working_index，
        让此前成功提交的修复参与家族一致性判断）；汇总格角色守卫以
        ``source_index`` 的原始目标为准（角色不随此前修复改变）。格式
        良好的孤立汇总（如行尾合计 =SUM(B2:M2)）豁免；没有家族时才退回
        直接邻居的多数规则。
        示例：目标公式与同排兄弟都属于 ``=S(Revenue)R(0,0)-S(Cost)R(0,0)``
        模式时通过。
        """
        try:
            target = attempt_index.cell(target_cell)
            if target is None or not target.is_formula or not target.formula:
                return CHECK_FAILED, f"目标单元格 {target_cell} 缺少有效公式"
            # 格式良好的孤立汇总（行/列合计）与明细模板不同是正常的，直接通过。
            if is_wellformed_aggregate(attempt_index, target):
                return CHECK_PASSED, ""
            # 汇总格角色守卫：目标原本是汇总公式时，修复不能改成明细公式
            # （即使它恰好匹配明细行家族的模板）。case_015 的
            # =RawData!N2*RawData!N4+RawData!N5 就是这种"看起来像兄弟、
            # 实际不是合计"的错误修复。
            original_target = source_index.cell(target_cell)
            original_is_aggregate = (
                original_target is not None
                and original_target.is_formula
                and original_target.formula
                and has_aggregate_function(original_target.formula)
            )
            if original_is_aggregate and not has_aggregate_function(target.formula):
                return CHECK_FAILED, "目标单元格原本是汇总格，修复后不应改成明细公式"
            # 家族主流模板：候选应与同排/列家族的主流写法一致。
            family = family_context_data(family_index, target.ref)
            candidate_template = normalize_template(
                to_relative_template(target.formula, target.ref)
            )
            candidate_structural = structural_template(candidate_template)
            structural_dominant = family.get("dominant_structural_template")
            if family["dominant_template"] is not None:
                # 有真实相对多数派时保持严格：偏离家族位移语义的提案
                # 不得因"形状相同"放行（否则同形不同位的范围错误可以
                # 全绿提交，wrong_range 型公式目标没有 value 检查兜底）。
                if candidate_template == family["dominant_template"]:
                    return CHECK_PASSED, ""
                return CHECK_FAILED, "目标公式与同排/列家族的主流模板不一致"
            # 相对模板层失明（无 count>=2 相对主流）：结构主流兜底。
            # 同文复制的固定查找区（如 12 格相同的 Rate!A2:B13）逐宿主
            # 相对化后模板互不相同，结构主流一致时视为家族一致
            # （complete_018 的 Revenue!L4）。
            if structural_dominant and candidate_structural == structural_dominant:
                return CHECK_PASSED, ""
            # 兜底：目标没有 >=3 成员的家族时，退回直接邻居多数规则。
            sheet = attempt_index._sheet_by_name(target.ref.sheet)
            if sheet is None:
                return CHECK_FAILED, f"目标单元格 {target_cell} 所在工作表不存在"
            target_template = normalize_template(
                to_relative_template(target.formula, target.ref)
            )
            matches = 0
            total = 0
            for row_delta, col_delta in ((0, -1), (0, 1), (-1, 0), (1, 0)):
                row = target.ref.row + row_delta
                col = target.ref.col + col_delta
                if row < 1 or col < 1:
                    continue
                neighbor = sheet.cells.get(f"{get_column_letter(col)}{row}")
                if neighbor is not None and neighbor.is_formula and neighbor.formula:
                    total += 1
                    if normalize_template(
                        to_relative_template(neighbor.formula, neighbor.ref)
                    ) == target_template:
                        matches += 1
            if total == 0 or matches >= total / 2:
                return CHECK_PASSED, ""
            return CHECK_FAILED, f"目标公式与全部 {total} 个公式邻居的模板均不一致"
        except Exception as exc:
            return CHECK_FAILED, f"模式检查执行失败: {type(exc).__name__}: {exc}"[:200]

    def _check_value_preserved(
        self,
        source_index: WorkbookIndex,
        attempt_index: WorkbookIndex,
        attempt_graph: DependencyGraph,
        target_cell: str,
        tentative_graph: DependencyGraph | None,
        prior_fixed_targets: set[str],
        rel_tol: float = 1e-6,
        abs_tol: float = 1e-3,
    ) -> tuple[str, str]:
        """检查缺失公式修复后，目标单元格能否复现原来的数值常量。

        仅当目标原值为数值常量（missing_formula 场景）时启用：把候选
        公式在 attempt 中重算，与原常量在容差内一致才通过。原值本身是
        公式或非数值时跳过。

        multi-cell 条件语义（设计 7.2）：如果 tentative dependency graph
        中存在"已 fixed 单元格 → 当前候选"的依赖路径，上游修复可能
        合法地改变本格的正确计算结果，value preservation 记为
        ``not_applicable``（reason=upstream_changed），不再把 source 中
        的旧常量作为硬失败条件；syntax/pattern/dependency/regression/
        recalc 仍必须通过。

        示例：原常量 56039.55，候选重算 385266.91 且无已修复上游时拒绝。
        """
        original = source_index.cell(target_cell)
        if original is None or original.is_formula:
            return CHECK_PASSED, ""
        stored = original.value
        if not isinstance(stored, (int, float)) or isinstance(stored, bool):
            return CHECK_PASSED, ""
        try:
            if tentative_graph is not None and prior_fixed_targets:
                upstream_changed = any(
                    tentative_graph.reachable(addr, target_cell)
                    for addr in prior_fixed_targets
                )
                if upstream_changed:
                    return CHECK_NOT_APPLICABLE, "upstream_changed"
            engine = RecalcEngine(attempt_index, attempt_graph)
            values = engine.evaluate_all()
            recomputed = values.get(target_cell, float("nan"))
            if isinstance(recomputed, float) and math.isnan(recomputed):
                root = engine.errors.get(target_cell, "")
                detail = f"（{root}）" if root else ""
                return (
                    CHECK_FAILED,
                    f"修复后目标单元格无法重算{detail}，无法与原值 {stored} 比对",
                )
            if not isinstance(recomputed, (int, float)):
                return CHECK_FAILED, f"修复后目标单元格结果为文本，与原值 {stored} 不符"
            if math.isclose(float(recomputed), float(stored), rel_tol=rel_tol, abs_tol=abs_tol):
                return CHECK_PASSED, ""
            return CHECK_FAILED, f"修复后结果 {recomputed} 与原值 {stored} 不符"
        except Exception as exc:
            return CHECK_FAILED, f"值保持检查执行失败: {type(exc).__name__}: {exc}"[:200]

    def _check_dependency(
        self,
        attempt_graph: DependencyGraph,
        working_graph: DependencyGraph | None = None,
    ) -> tuple[str, str]:
        """检查是否产生 working graph 之外新引入的循环依赖。

        单目标模式（working_graph 为 None）：attempt 中任何循环都失败。
        multi-cell 模式（设计 5.2）：先区分稳定 working graph 中已经存在
        的 cycle 与当前 repair proposal 新引入的 cycle——只有新出现的
        SCC 才算失败；既有循环属于调度器层面的事实，由
        ``blocked_by[].cause=dependency_cycle`` 处理，不该在此拦截。

        示例：``A1=B1+1`` 且 ``B1=A1+1`` 形成循环，检查失败。
        """
        try:
            if working_graph is None:
                if attempt_graph.has_cycle():
                    return CHECK_FAILED, "修复后依赖图出现循环引用"
                return CHECK_PASSED, ""
            working_sccs = set(working_graph.cycle_sccs())
            for scc in attempt_graph.cycle_sccs():
                if scc not in working_sccs:
                    preview = ", ".join(sorted(scc)[:3])
                    return (
                        CHECK_FAILED,
                        f"修复 proposal 新引入循环依赖: {preview}",
                    )
            return CHECK_PASSED, ""
        except Exception as exc:
            return CHECK_FAILED, f"依赖检查执行失败: {type(exc).__name__}: {exc}"[:200]

    def _check_regression(
        self,
        source_index: WorkbookIndex,
        attempt_index: WorkbookIndex,
        allowed: set[str],
    ) -> tuple[str, str]:
        """比较原始 Workbook 和 attempt Workbook，检测未授权修改。

        以 ``source_index`` 为基线（设计 7.2：未授权修改检测用原始事实）。
        返回（状态，失败说明）；失败说明里列出被改动但不在允许
        清单中的单元格地址（最多 5 个）。
        示例：如果只允许修改 ``P&L!C6``，而 ``Revenue!B20`` 也发生变化，
        则回归检查失败。
        """
        try:
            original_cells = self._cell_states(source_index)
            attempt_cells = self._cell_states(attempt_index)
            changed = sorted(
                address
                for address in original_cells.keys() | attempt_cells.keys()
                if address not in allowed
                and original_cells.get(address) != attempt_cells.get(address)
            )
            if changed:
                preview = ", ".join(changed[:5])
                suffix = "..." if len(changed) > 5 else ""
                return CHECK_FAILED, f"修改了目标之外的单元格: {preview}{suffix}"
            return CHECK_PASSED, ""
        except Exception as exc:
            return CHECK_FAILED, f"回归检查执行失败: {type(exc).__name__}: {exc}"[:200]

    @staticmethod
    def _cell_states(index: WorkbookIndex) -> dict[str, tuple[object, ...]]:
        """提取每个单元格的可比较状态。

        公式记录为 ``("formula", 公式文本)``，常量记录为
        ``("value", Python 类型, 值)``，从而区分数字 1 和布尔值 True。
        """
        states: dict[str, tuple[object, ...]] = {}
        for sheet in index.sheets:
            for info in sheet.cells.values():
                if info.is_formula:
                    states[info.ref.full_address] = ("formula", info.formula)
                else:
                    # Include the Python value type: in Python, 1 == True,
                    # but Excel numeric and boolean cells are distinct content.
                    states[info.ref.full_address] = ("value", type(info.value), info.value)
        return states

    def _check_recalc(
        self,
        attempt_index: WorkbookIndex,
        attempt_graph: DependencyGraph,
        exempt_cells: set[str] | None = None,
        baseline_nan: set[str] | None = None,
    ) -> tuple[str, str]:
        """使用 DependencyGraph 和 RecalcEngine 检查 attempt 能否成功重算。

        检查内容包括：所有公式都有计算结果且结果不是 NaN。multi-cell
        模式按 working 基线差分判定：``exempt_cells``（working graph 中
        既有循环的成员）与 ``baseline_nan``（working 基线里就重算失败的
        未修复错误格）豁免——它们无法计算是工作簿既有事实，不属于本次
        尝试的责任；目标单元格不入基线，修复后仍算不出必失败。示例：
        ``=A1/0`` 会导致重算检查失败。
        """
        try:
            exempt = set(exempt_cells or set()) | set(baseline_nan or set())
            engine = RecalcEngine(attempt_index, attempt_graph)
            values = engine.evaluate_all()
            formula_addresses = {info.ref.full_address for info in attempt_index.formula_cells()}
            expected = formula_addresses - exempt
            if not expected.issubset(values):
                return CHECK_FAILED, "部分公式没有计算出结果"
            new_nan = sorted(
                address
                for address, value in values.items()
                if address not in exempt
                and isinstance(value, float)
                and math.isnan(value)
            )
            if new_nan:
                preview = ", ".join(new_nan[:3])
                suffix = "..." if len(new_nan) > 3 else ""
                # 把引擎记录的根因带进 reason（如 Unsupported function:
                # ROUND），让修复重试知道错在哪类，而不是只看到 NaN。
                root = engine.errors.get(new_nan[0], "")
                detail = f"（{new_nan[0]}: {root}）" if root else ""
                return (
                    CHECK_FAILED,
                    f"存在新增 NaN 等异常计算结果: {preview}{suffix}{detail}",
                )
            return CHECK_PASSED, ""
        except Exception as exc:
            return CHECK_FAILED, f"重算执行失败: {type(exc).__name__}: {exc}"[:200]
