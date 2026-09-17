"""用于候选排序的确定性公式异常检测。

本模块是 SheetGuard 的第一道筛选器：不请 LLM，先用确定性规则给每个
公式单元格打"可疑分"，产出按分数降序的候选名单，交给后续的
Localizer / Diagnoser / Verifier 深入调查。终端聚合格（合法汇总格）
的豁免判据在 spreadsheet/aggregates.py，与 Verifier 共用。

方法导览
========

detect()                           入口：跑全部 6 个信号，按分数降序返回候选
├─ _signal_missing_formula         数字常量打断了公式连续段 → missing_formula (0.6)
│    例：B2=A2*2、C2=42(常量)、D2=C2*2 —— 明明两侧都是公式，
│        C2 却是硬编码的数字，疑似该有公式的地方被填成了常量。
├─ _signal_pattern_anomaly         家族里的模板少数派 → pattern_anomaly (0.5)
│    例：一行 12 个公式都是 =RawData!x2*RawData!x3，唯独 E2 写成
│        =RawData!E3*RawData!E4（引用行错位），模板与众不同的就是
│        它。行尾的汇总格 =SUM(B2:M2) 模板同样不同，但它完整盖住
│        了旁边的公式段（见聚合格识别），属于正常"总计"，不报。
├─ _signal_neighbor_inconsistency  左右/上下邻居一致但本格不同 → neighbor_mismatch (0.3)
│    例：B10=A10*C10、C10=B10+D10、D10=C10*E10 —— B10 和 D10 是
│        同一种"乘法"写法，夹在中间的 C10 却是"加法"。
├─ _signal_range_boundary          同家族 SUM 的范围尺寸不一致 → range_boundary (0.5)
│    例：B1=SUM(B2:E2)、C1=SUM(B3:E3)、D1=SUM(B4:D4) —— 别人都
│        竖着加 4 格，D1 只加 3 格，范围少了一截。
├─ _signal_singleton_sum_boundary  孤立 SUM 没盖住相邻公式段 → range_boundary (0.5)
│    例：一行明细 B2..D2 都是公式，E2 却写 =SUM(B2:C2) —— 范围
│        外还紧贴着公式 D2，说明漏加了。由于"该盖的段"是靠模板
│        推算的，段中间夹杂一个坏格子就会把推算截断、冤枉合法
│        汇总，所以报警前用 _is_wellformed_aggregate 复核：范围
│        两端外侧没有公式（真的盖全了）就不报。
└─ _signal_dependency_anomaly      跨表引用集合与家族主流不同 → dependency_anomaly (0.4)
     例：一行公式都引用 Revenue 表，唯独 D1 引用 Cost 表 —— 引用
         的"数据来源"和大家不一样。行尾的本地汇总 =SUM(B2:M2)
         不引用任何跨表数据也是正常的，盖全即不报。

通用工具（被打分函数调用）
_add_signal              把信号记到候选单元格上（同信号不重复计分）
_template                公式 → 相对引用模板（写法相同 = 模板相同）
_neighbor                取某方向的相邻单元格
_template_run            统计某侧"写法完全相同"的连续公式段
_opposite_neighbors      产出 (左,右)、(上,下) 两组对向公式邻居对

聚合格识别（"终端汇总格"豁免的判据，实现在 spreadsheet/aggregates.py，
与 Verifier 的 pattern 检查共用同一份判据）
aggregate_span           是不是孤立的一维本地聚合公式（如 =SUM(B2:M2)）
is_wellformed_aggregate  聚合范围两端外侧没有公式 = 盖全了 = 合法汇总
"""
from __future__ import annotations

from collections import Counter, defaultdict
from typing import Iterable

from openpyxl.formula.tokenizer import TokenizerError
from openpyxl.utils import get_column_letter, range_boundaries

from sheetguard.config import DEFAULT_SIGNAL_WEIGHTS
from sheetguard.spreadsheet.aggregates import (
    aggregate_span,
    family_context_data,
    is_wellformed_aggregate,
)
from sheetguard.spreadsheet.formula_expectation import dependency_judgement
from sheetguard.spreadsheet.formula_parser import extract_refs, to_relative_template
from sheetguard.spreadsheet.model import CellInfo, WorkbookIndex


class StaticInspector:
    """基于确定性局部模式信号，对公式单元格进行异常候选排序。

    这个类是 SheetGuard 的第一道筛选器：它不直接修复公式，
    而是先扫描 WorkbookIndex，从大量单元格中找出最值得进一步调查的候选。

    主要检测信号包括：
        - missing_formula：数字常量打断了原本连续的公式模式；
        - pattern_anomaly：公式的相对引用模板偏离公式家族；
        - neighbor_mismatch：当前公式与两侧一致的邻居不一致；
        - range_boundary：SUM 范围的宽度或高度与同组公式不一致，或孤立汇总
          没有完整覆盖相邻的公式连续段；
        - dependency_anomaly：公式引用的跨工作表集合与同组公式不一致。

    豁免规则：格式良好的孤立行/列汇总（如 =SUM(B2:M2) 且范围两端外侧
    没有其它公式）是正常的"终端聚合格"，其模板和跨表集合本来就应该与
    明细不同，因此不参与 pattern_anomaly 和 dependency_anomaly 报警；
    只有范围没有盖全（范围外侧还紧贴公式）时才记 range_boundary。

    每个信号都会给候选单元格增加分数，多个信号可以叠加，
    最终由 detect() 按分数从高到低返回候选列表。

    示例：
        inspector = StaticInspector(index)
        suspects = inspector.detect(top_k=10)

        返回结果可能是：
        [
            {
                "cell": "P&L!D20",
                "score": 0.8,
                "signals": ["pattern_anomaly", "neighbor_mismatch"],
            }
        ]

    这些候选随后会交给 Localizer、Diagnosis Agent 和 Verifier，
    进行更深入的定位、诊断、修复和验证。
    """

    def __init__(self, index: WorkbookIndex, weights: dict[str, float] | None = None):
        self.index = index
        # 信号权重表：来自运行配置（YAML 可外置）；缺省用内置默认值。
        self._weights: dict[str, float] = dict(
            DEFAULT_SIGNAL_WEIGHTS if weights is None else weights
        )
        # 多视角共识置信度缓存（设计 5.2）：候选的 candidate confidence 证据源。
        self._consensus_confidence: dict[str, float] = {}

    def detect(self, top_k: int | None = 50) -> list[dict]:
        """返回稳定、按分数降序排列的公式异常候选。

        ``top_k`` 仅用于展示层截断（例如旧的单目标 localize 一次最多查看
        50 个候选）；multi-cell 批处理调度器需要完整的 static candidate
        universe，必须传 ``top_k=None`` 或改用 ``detect_all()``。
        """
        scores: dict[str, float] = defaultdict(float)
        signals: dict[str, list[str]] = defaultdict(list)
        self._consensus_confidence = {}

        self._signal_missing_formula(scores, signals)
        self._signal_pattern_anomaly(scores, signals)
        self._signal_neighbor_inconsistency(scores, signals)
        self._signal_range_boundary(scores, signals)
        self._signal_singleton_sum_boundary(scores, signals)
        self._signal_dependency_anomaly(scores, signals)

        suspects = [
            {
                "cell": cell,
                "score": min(score, 1.0),
                "signals": signals[cell],
                **(
                    {"consensus_confidence": self._consensus_confidence[cell]}
                    if cell in self._consensus_confidence
                    else {}
                ),
            }
            for cell, score in scores.items()
        ]
        suspects.sort(key=lambda item: (-item["score"], item["cell"]))
        if top_k is None:
            return suspects
        return suspects[:top_k]

    def detect_all(self) -> list[dict]:
        """返回不受 ``top_k`` 截断的完整 static candidate universe。

        multi-cell scheduler 的 ``static_candidate_count``、seed 过滤和
        prerequisite closure 都必须基于这个完整集合；真正的 Agent 成本
        边界由 ``min_anomaly_score`` 和 ``max_candidates`` 控制，而不是
        这里隐藏的截断。
        """
        return self.detect(top_k=None)

    def family_context(self, cell: str, max_siblings: int = 3) -> dict:
        """返回目标单元格所在行/列家族的修复上下文。

        收集三类信息供修复阶段参考：同家族主流模板的兄弟单元格公式
        原文（最多 max_siblings 个）、家族主流跨表引用集合、目标自身
        当前引用的表集合。修复"该引用哪张表"时，兄弟公式的原文就是
        现成答案——同一个（行, 列）坐标在每张表都存在，闭卷猜表名
        是五选一的赌博，抄邻居则是确定的。

        目标本身是数字常量（missing_formula 场景）时同样支持：按同一
        行/列的反推家族，返回兄弟公式与主流跨表引用（目标当前引用为空），
        让修复模型有模板可抄，而不是闭卷瞎猜。
        """
        empty = {"兄弟单元格公式": [], "家族主流跨表引用": [], "目标当前引用": []}
        info = self.index.cell(cell)
        if info is None:
            return empty
        # 家族主流模板/兄弟公式/主流跨表引用由 aggregates.family_context_data
        # 统一计算（与 Verifier 的 pattern 检查共用同一判据）。
        data = family_context_data(self.index, info.ref, max_siblings=max_siblings)
        if info.is_formula and info.formula:
            target_set = self._cross_sheet_set(info) or frozenset()
        else:
            target_set = frozenset()
        return {
            "兄弟单元格公式": data["siblings"],
            "家族主流跨表引用": data["dominant_sheets"],
            "目标当前引用": sorted(target_set),
        }

    @staticmethod
    def _add_signal(
        scores: dict[str, float],
        signals: dict[str, list[str]],
        cell: str,
        signal: str,
        weight: float,
    ) -> None:
        """把一个异常信号记到候选单元格上。

        同一单元格的同一信号只记一次（避免重复计分）；每记一次，
        该单元格的嫌疑分增加 weight。所有打分函数都通过它写入结果。
        """
        if signal not in signals[cell]:
            signals[cell].append(signal)
            scores[cell] += weight

    @staticmethod
    def _template(info: CellInfo) -> str | None:
        """返回公式的相对引用模板；非公式单元格或解析失败时返回 None。

        模板把公式里的引用替换成相对位移，例如 =RawData!B2*RawData!B3
        会变成 S(RawData)R(0,0)*S(RawData)R(1,0)。因此同行/列中"写法
        相同"的公式模板一致，模板不一致的单元格就是离群者。
        """
        if not info.is_formula or not info.formula:
            return None
        try:
            return to_relative_template(info.formula, info.ref)
        except (IndexError, TokenizerError, TypeError, ValueError):
            return None

    def _family_groups(self) -> list[list[CellInfo]]:
        """返回至少包含 3 个单元格的横向和纵向公式家族。"""
        rows: dict[tuple[str, int], list[CellInfo]] = defaultdict(list)
        columns: dict[tuple[str, int], list[CellInfo]] = defaultdict(list)
        for info in self.index.formula_cells():
            rows[(info.ref.sheet, info.ref.row)].append(info)
            columns[(info.ref.sheet, info.ref.col)].append(info)
        return [
            group
            for group in [*rows.values(), *columns.values()]
            if len(group) >= 3
        ]

    @staticmethod
    def _is_numeric_hardcode(info: CellInfo) -> bool:
        """返回一个非公式单元格是否持有数字值（而非标签）。"""
        return (
            not info.is_formula
            and isinstance(info.value, (int, float))
            and not isinstance(info.value, bool)
        )

    def _template_run(
        self, info: CellInfo, row_delta: int, col_delta: int
    ) -> tuple[str | None, int]:
        """统计 info 某一侧"写法完全相同"的连续公式段。

        从 info 沿 (row_delta, col_delta) 方向的邻居开始，逐格检查模板
        是否一致，返回（该段的统一模板, 段内公式个数）。邻居不是公式
        时返回 (None, 0)。用于推算汇总格"本该覆盖"的范围，以及判断
        数字常量旁边是否存在真实的公式连续段。
        """
        neighbor = self._neighbor(info, row_delta, col_delta)
        template = self._template(neighbor) if neighbor else None
        if template is None:
            return None, 0

        length = 0
        while neighbor and neighbor.is_formula and self._template(neighbor) == template:
            length += 1
            neighbor = self._neighbor(neighbor, row_delta, col_delta)
        return template, length

    def _has_formula_hole_support(self, info: CellInfo) -> bool:
        """判断一个数字常量旁边是否真的存在"本该有公式"的连续段。

        沿右、下两个方向检查：只要某一侧的连续同模板公式段长度 >= 2，
        或常量两侧的模板互相衔接，就认为这个常量打断了原本连续的公式
        模式（疑似 missing_formula）；孤立的普通常量返回 False。
        """
        for row_delta, col_delta in ((0, 1), (1, 0)):
            before_template, before_length = self._template_run(
                info, -row_delta, -col_delta
            )
            after_template, after_length = self._template_run(info, row_delta, col_delta)
            if (
                before_template is not None
                and before_template == after_template
            ):
                return True
            if before_length >= 2 or after_length >= 2:
                return True
        return False

    def _signal_missing_formula(
        self,
        scores: dict[str, float],
        signals: dict[str, list[str]],
    ) -> None:
        """标记打断或紧邻公式模板连续段的数字硬编码。"""
        for sheet in self.index.sheets:
            for info in sheet.cells.values():
                if self._is_numeric_hardcode(info) and self._has_formula_hole_support(info):
                    self._add_signal(
                        scores,
                        signals,
                        info.ref.full_address,
                        "missing_formula",
                        self._weights["missing_formula"],
                    )

    def _signal_pattern_anomaly(
        self,
        scores: dict[str, float],
        signals: dict[str, list[str]],
    ) -> None:
        """标记横向或纵向公式家族中的非主流模板。"""
        for family in self._family_groups():
            templates = [
                (info, template)
                for info in family
                if (template := self._template(info)) is not None
            ]
            if len(templates) < 3:
                continue
            counts = Counter(template for _, template in templates)
            dominant_template, dominant_count = counts.most_common(1)[0]
            if dominant_count < 2:
                continue
            for info, template in templates:
                if template == dominant_template:
                    continue
                # 格式良好的孤立汇总是行/列的正常"总计"格，不作为少数派报警。
                if is_wellformed_aggregate(self.index, info):
                    continue
                self._add_signal(
                    scores,
                    signals,
                    info.ref.full_address,
                        "pattern_anomaly",
                        self._weights["pattern_anomaly"],
                )

    def _neighbor(self, info: CellInfo, row_delta: int, col_delta: int) -> CellInfo | None:
        """返回 info 沿 (row_delta, col_delta) 方向相邻的单元格；越界返回 None。"""
        row = info.ref.row + row_delta
        col = info.ref.col + col_delta
        if row < 1 or col < 1:
            return None
        address = f"{get_column_letter(col)}{row}"
        return self.index.cell(f"{info.ref.sheet}!{address}")

    def _opposite_neighbors(self, info: CellInfo) -> Iterable[tuple[CellInfo, CellInfo]]:
        """产出候选单元格的两组对向公式邻居对：(左, 右) 和 (上, 下)。

        只有两侧都是公式时才产出，供 _signal_neighbor_inconsistency
        判断"两侧写法一致、唯独本格不一致"的情形。
        """
        for before_delta, after_delta in [((0, -1), (0, 1)), ((-1, 0), (1, 0))]:
            before = self._neighbor(info, *before_delta)
            after = self._neighbor(info, *after_delta)
            if before and after and before.is_formula and after.is_formula:
                yield before, after

    def _signal_neighbor_inconsistency(
        self,
        scores: dict[str, float],
        signals: dict[str, list[str]],
    ) -> None:
        """标记与两侧一致的邻居不同的公式。"""
        for info in self.index.formula_cells():
            host_template = self._template(info)
            if host_template is None:
                continue
            for before, after in self._opposite_neighbors(info):
                before_template = self._template(before)
                after_template = self._template(after)
                if (
                    before_template is not None
                    and before_template == after_template
                    and host_template != before_template
                ):
                    self._add_signal(
                        scores,
                        signals,
                        info.ref.full_address,
                        "neighbor_mismatch",
                        self._weights["neighbor_mismatch"],
                    )
                    break

    @staticmethod
    def _sum_range_dimensions(info: CellInfo) -> tuple[int, int] | None:
        """返回第一个有效 SUM 范围的（宽度，高度），如果存在。"""
        if not info.formula or "SUM(" not in info.formula.upper():
            return None
        try:
            refs = extract_refs(info.formula)
        except (IndexError, TokenizerError, TypeError, ValueError):
            return None
        for _, reference in refs:
            if ":" not in reference:
                continue
            try:
                min_col, min_row, max_col, max_row = range_boundaries(reference)
            except (IndexError, TokenizerError, TypeError, ValueError):
                continue
            boundaries = (min_col, min_row, max_col, max_row)
            if not all(type(value) is int and value > 0 for value in boundaries):
                continue
            return max_col - min_col + 1, max_row - min_row + 1
        return None

    def _signal_range_boundary(
        self,
        scores: dict[str, float],
        signals: dict[str, list[str]],
    ) -> None:
        """标记范围尺寸与其公式家族不一致的 SUM 公式。"""
        for family in self._family_groups():
            ranged = [
                (info, dimensions)
                for info in family
                if (dimensions := self._sum_range_dimensions(info)) is not None
            ]
            if len(ranged) < 3:
                continue
            counts = Counter(dimensions for _, dimensions in ranged)
            dominant_dimensions, dominant_count = counts.most_common(1)[0]
            if dominant_count < 2:
                continue
            for info, dimensions in ranged:
                if dimensions != dominant_dimensions:
                    self._add_signal(
                        scores,
                        signals,
                        info.ref.full_address,
                        "range_boundary",
                        self._weights["range_boundary"],
                    )

    @staticmethod
    def _single_local_sum_range(info: CellInfo) -> tuple[int, int, int, int] | None:
        """返回一个有效的本地 SUM 范围，忽略畸形或复杂的公式。"""
        if not info.formula or "SUM(" not in info.formula.upper():
            return None
        try:
            refs = extract_refs(info.formula)
        except (IndexError, TokenizerError, TypeError, ValueError):
            return None
        if len(refs) != 1:
            return None
        sheet, reference = refs[0]
        if sheet is not None or ":" not in reference:
            return None
        try:
            boundaries = range_boundaries(reference)
        except (IndexError, TokenizerError, TypeError, ValueError):
            return None
        if not all(type(value) is int and value > 0 for value in boundaries):
            return None
        return boundaries

    def _adjacent_template_ranges(self, info: CellInfo) -> list[tuple[int, int, int, int]]:
        """推算孤立 SUM "本该覆盖"的范围列表。

        沿右、下两个轴的正反方向找长度 >= 2 的同模板公式段，把每段的
        行列范围换算成 (min_col, min_row, max_col, max_row)。孤立 SUM
        的实际范围若都不在其中，就是"没盖住该盖的段"（但注意：段的
        推算可能被一个写法不同的坏格子截断，所以报警前还要叠加
        _is_wellformed_aggregate 的盖全检查）。
        """
        expected: list[tuple[int, int, int, int]] = []
        for row_delta, col_delta in ((0, 1), (1, 0)):
            for direction in (-1, 1):
                _, length = self._template_run(
                    info, row_delta * direction, col_delta * direction
                )
                if length < 2:
                    continue
                if row_delta:
                    first_row = info.ref.row + direction
                    last_row = info.ref.row + direction * length
                    expected.append(
                        (info.ref.col, min(first_row, last_row), info.ref.col, max(first_row, last_row))
                    )
                else:
                    first_col = info.ref.col + direction
                    last_col = info.ref.col + direction * length
                    expected.append(
                        (min(first_col, last_col), info.ref.row, max(first_col, last_col), info.ref.row)
                    )
        return expected

    def _signal_singleton_sum_boundary(
        self,
        scores: dict[str, float],
        signals: dict[str, list[str]],
    ) -> None:
        """标记未能精确覆盖相邻公式连续段的孤立 SUM。

        原始的模板段推断可能因为范围中夹杂一个异常单元格（正是待查的
        错误）而把合法汇总误判为"没盖全"，因此这里叠加范围覆盖检查：
        只要聚合范围两端外侧没有公式，就视为完整覆盖，不再报警。
        """
        for info in self.index.formula_cells():
            expected_ranges = self._adjacent_template_ranges(info)
            if not expected_ranges:
                continue
            actual_range = self._single_local_sum_range(info)
            if (
                actual_range is not None
                and actual_range not in expected_ranges
                and not is_wellformed_aggregate(self.index, info)
            ):
                self._add_signal(
                    scores,
                    signals,
                    info.ref.full_address,
                    "range_boundary",
                    self._weights["range_boundary"],
                )

    @staticmethod
    def _cross_sheet_set(info: CellInfo) -> frozenset[str] | None:
        """返回公式引用的跨表名称，格式错误时返回 None。"""
        if not info.formula:
            return None
        try:
            return frozenset(
                sheet
                for sheet, _ in extract_refs(info.formula)
                if sheet is not None
            )
        except (IndexError, TokenizerError, TypeError, ValueError):
            return None

    def _signal_dependency_anomaly(
        self,
        scores: dict[str, float],
        signals: dict[str, list[str]],
    ) -> None:
        """标记引用集合偏离多视角结构共识的公式（v1.6 设计 4.5）。

        旧实现把"目标格属于哪个 family"当成唯一入口：目标格公式本身写错
        时会被 family membership 排除，且行/列二选一在混合业务层级的表上
        产生 false positive。新流程把目标格遮住，由横向/纵向/依赖视角独立
        预测期望引用集合，实际集合与所有高置信共识都不一致才报警；聚合格
        （终端汇总）走确定性分支，豁免语义与 v1.5 完全一致。
        """
        for info in self.index.formula_cells():
            deviation, expectation = dependency_judgement(self.index, info)
            if (
                expectation is not None
                and expectation.has_consensus
            ):
                # 共识置信度随候选透出（设计 5.2），供调度器 candidate confidence 使用。
                self._consensus_confidence[info.ref.full_address] = (
                    expectation.confidence
                )
            if deviation is None or not deviation.is_dependency_outlier:
                continue
            self._add_signal(
                scores,
                signals,
                info.ref.full_address,
                "dependency_anomaly",
                self._weights["dependency_anomaly"],
            )
