"""基于 NetworkX 实现的公式依赖图。

本模块定义 DependencyGraph 类：
  - 从 WorkbookIndex 构建单元格级别的有向依赖图
  - 边方向：被依赖单元格（前驱/上游） → 依赖单元格（后继/下游）
  - 提供多种查询接口：前驱/后继搜索、最短依赖路径、循环检测、表级依赖聚合、拓扑排序
  - multi-cell 批处理调度专用查询：传递可达性、静态候选 prerequisite closure、
    候选依赖排序、SCC 循环识别与循环成员查询
  - 供依赖查询、静态异常检测、重算引擎、验证器共同使用
"""
from __future__ import annotations
import heapq
from typing import Iterable
import networkx as nx
from sheetguard.spreadsheet.model import WorkbookIndex
from sheetguard.spreadsheet.formula_parser import extract_refs, expand_range


class DependencyGraph:
    """单元格级公式依赖图。

    用 NetworkX 有向图存储 Excel 单元格之间的依赖关系：
        - 边方向：precedent（被依赖单元格） → dependent（依赖单元格）
        - 如果 C = A + B，则边为 A → C，B → C
    支持查询前驱/后继、最短路径、循环检测、表级聚合、拓扑排序。
    """

    def __init__(self, index: WorkbookIndex):
        self.index = index  # 工作簿索引，从中读取单元格信息
        self._graph: nx.DiGraph = nx.DiGraph()  # 内部存储的 NetworkX 有向图

    def build(self) -> None:
        """从工作簿索引构建依赖图。

        遍历所有单元格，对每个公式单元格提取它引用的所有单元格，
        然后添加边：被引用单元格 → 当前公式单元格。

        例子：
            公式单元格 P&L!C6 = SUM(P&L!B2:B5) + P&L!C4
            → 提取引用得到 (P&L, "B2:B5") 和 (None, "C4")
            → 展开 B2:B5 → B2, B3, B4, B5
            → 添加边：P&L!B2 → P&L!C6, P&L!B3 → P&L!C6, ..., P&L!C4 → P&L!C6
        """
        self._graph = nx.DiGraph()
        for sheet in self.index.sheets:
            for addr, info in sheet.cells.items():
                node = info.ref.full_address
                self._graph.add_node(node)
                if info.is_formula and info.formula:
                    for ref_sheet, ref_range in extract_refs(info.formula):
                        sheet_name = ref_sheet or sheet.name
                        cells = expand_range(sheet_name, ref_range)
                        for (s, a) in cells:
                            prec = f"{s}!{a}"
                            self._graph.add_node(prec)
                            self._graph.add_edge(prec, node)

    def precedents_of(self, cell: str, depth: int = 1) -> list[str]:
        """返回当前单元格依赖的所有单元格（前驱，即上游单元格）。

        参数：
            cell: 目标单元格完整地址，例如 "P&L!C6"
            depth: 向上搜索多少层，depth=1 只找直接依赖，depth>1 继续递归找上游

        例子：
            C6 = B2 + C4， B2 = A1 + A2
            precedents_of("P&L!C6", depth=1) → ["P&L!B2", "P&L!C4"]
            precedents_of("P&L!C6", depth=2) → ["P&L!A1", "P&L!A2", "P&L!B2", "P&L!C4"]
        """
        if depth <= 1:
            return list(self._graph.predecessors(cell))
        result: set[str] = set()
        frontier = {cell}
        for _ in range(depth):
            next_frontier: set[str] = set()
            for c in frontier:
                for p in self._graph.predecessors(c):
                    if p not in result:
                        result.add(p)
                        next_frontier.add(p)
            frontier = next_frontier
        return sorted(result)

    def dependents_of(self, cell: str, depth: int = 1) -> list[str]:
        """返回所有依赖当前单元格的单元格（后继，即下游单元格）。

        参数：
            cell: 目标单元格完整地址，例如 "Revenue!B20"
            depth: 向下搜索多少层，depth=1 只找直接依赖，depth>1 继续递归找下游

        例子：
            Revenue!B20 → P&L!B30 → Dashboard!C8
            dependents_of("Revenue!B20", depth=1) → ["P&L!B30"]
            dependents_of("Revenue!B20", depth=2) → ["P&L!B30", "Dashboard!C8"]
        """
        if depth <= 1:
            return list(self._graph.successors(cell))
        result: set[str] = set()
        frontier = {cell}
        for _ in range(depth):
            next_frontier: set[str] = set()
            for c in frontier:
                for d in self._graph.successors(c):
                    if d not in result:
                        result.add(d)
                        next_frontier.add(d)
            frontier = next_frontier
        return sorted(result)

    def path(self, src: str, dst: str) -> list[str] | None:
        """查找从 src 到 dst 的最短依赖路径。

        用来判断"上游单元格是否会影响某个下游单元格"，并给出具体路径。

        参数：
            src: 起点单元格完整地址，例如 "Orders!B20"
            dst: 终点单元格完整地址，例如 "Dashboard!C8"

        返回：
            如果存在路径，返回路径上的单元格列表；否则返回 None。

        例子：
            path("Orders!B20", "Dashboard!C8")
            → ["Orders!B20", "Revenue!B10", "P&L!B30", "Dashboard!C8"]
        """
        try:
            return nx.shortest_path(self._graph, source=src, target=dst)
        except (nx.NetworkXNoPath, nx.NodeNotFound):
            return None

    def has_cycle(self) -> bool:
        """检查整个依赖图中是否存在有向循环（循环依赖）。

        循环依赖例如：A1 = B1 + 1，B1 = A1 + 1，这样无法计算。

        返回：
            True → 存在循环，False → 无循环。

        例子：
            A1 = B1+1, B1 = A1+1 → has_cycle() → True
        """
        try:
            nx.find_cycle(self._graph, orientation="original")
            return True
        except nx.NetworkXNoCycle:
            return False

    def sheet_level_graph(self) -> dict[str, set[str]]:
        """将单元格级依赖聚合为表级依赖。

        返回：
            {上游表名: {下游表名集合}}，只保留跨表依赖。

        例子：
            单元格级存在 Orders!B20 → Revenue!B10，Revenue!B10 → P&L!B30
            → sheet_level_graph() → {"Orders": {"Revenue"}, "Revenue": {"P&L"}}
        """
        sg: dict[str, set[str]] = {}
        for u, v in self._graph.edges():
            u_sheet = u.split("!")[0]
            v_sheet = v.split("!")[0]
            if u_sheet != v_sheet:
                sg.setdefault(u_sheet, set()).add(v_sheet)
        return sg

    def topological_order(self) -> list[str]:
        """返回公式单元格的拓扑排序（依赖优先顺序）。

        拓扑排序保证：如果 A 依赖 B，则排序中 B 一定出现在 A 前面。
        重算引擎按这个顺序计算，可以保证计算 A 时它依赖的单元格已经算完。

        返回：
            排序后的公式单元格地址列表；如果存在循环依赖，返回空列表。

        例子：
            Orders!B20 → Revenue!B10 → P&L!B30 → Dashboard!C8
            → topological_order() → ["Orders!B20", "Revenue!B10", "P&L!B30", "Dashboard!C8"]
        """
        formula_nodes = {
            info.ref.full_address
            for info in self.index.formula_cells()
        }
        subgraph = self._graph.subgraph(
            n for n in self._graph.nodes if n in formula_nodes
        )
        try:
            return list(nx.topological_sort(subgraph))
        except nx.NetworkXUnfeasible:
            return []

    # ------------------------------------------------------------------
    # multi-cell 批处理调度专用查询
    # ------------------------------------------------------------------

    def ancestors_of(self, cell: str) -> set[str]:
        """返回 cell 的全部传递上游（能沿依赖边到达 cell 的单元格）。

        与 depth 参数版的 precedents_of 不同，这里不受层数限制，覆盖
        直接与间接依赖。cell 不在图中（例如尚未出现在图里的新地址）
        时返回空集合。
        """
        if cell not in self._graph:
            return set()
        return nx.ancestors(self._graph, cell)

    def descendants_of(self, cell: str) -> set[str]:
        """返回 cell 的全部传递下游（从 cell 沿依赖边可达的单元格）。"""
        if cell not in self._graph:
            return set()
        return nx.descendants(self._graph, cell)

    def reachable(self, src: str, dst: str) -> bool:
        """判断 dst 是否依赖 src（图中存在 src → dst 的路径）。"""
        if src not in self._graph or dst not in self._graph:
            return False
        try:
            return nx.has_path(self._graph, src, dst)
        except nx.NetworkXError:
            return False

    def strongly_connected_components(self) -> list[frozenset[str]]:
        """返回图中全部强连通分量（SCC）。

        循环依赖的成员恰好构成一个 size>1 的 SCC；单格自环通过
        selfloop_edges 单独识别。供 working-cycle 与
        proposal-introduced-cycle 的区分查询使用。
        """
        return [frozenset(scc) for scc in nx.strongly_connected_components(self._graph)]

    def cycle_members(self) -> set[str]:
        """返回参与任何有向循环的单元格集合（SCC size>1 或自环）。"""
        members: set[str] = set()
        for scc in self.strongly_connected_components():
            if len(scc) > 1:
                members |= scc
        members.update(u for u, _ in nx.selfloop_edges(self._graph))
        return members

    def cycle_sccs(self) -> list[frozenset[str]]:
        """只返回真正构成循环的 SCC（size>1 或含自环的单格组）。"""
        sccs = [scc for scc in self.strongly_connected_components() if len(scc) > 1]
        # 单格自环不在 size>1 的 SCC 里，单独成组返回。
        for loop in nx.selfloop_edges(self._graph):
            loop_node = loop[0]
            if not any(loop_node in scc for scc in sccs):
                sccs.append(frozenset({loop_node}))
        return sccs

    def evaluation_order(self) -> tuple[list[str], set[str]]:
        """返回（可安全计算的单元格顺序, 参与循环的单元格集合）。

        图无循环时等价于 topological_order()，循环成员为空集；
        存在循环时按 SCC 缩点图的拓扑序计算无循环部分，循环成员
        无法安全求值，由调用方按 NaN 处理（multi-cell 验证中这些
        成员属于工作簿既有事实，可被豁免）。
        """
        cyclic = self.cycle_members()
        if not cyclic:
            return self.topological_order(), set()
        condensation = nx.condensation(self._graph)
        order: list[str] = []
        for scc_id in nx.topological_sort(condensation):
            members = condensation.nodes[scc_id]["members"]
            if len(members) == 1:
                member = next(iter(members))
                if member not in cyclic:
                    order.append(member)
        return order, cyclic

    def prerequisite_closure(
        self, candidates: Iterable[str], universe: set[str]
    ) -> set[str]:
        """计算候选集合的静态候选上游 prerequisite closure。

        对每个 seed candidate，把图中能到达它的全部单元格与 static
        candidate universe 求交集，并集后返回。closure 覆盖直接与
        间接（经过非候选中间单元格）的依赖路径；候选自身也在结果中。
        """
        universe = set(universe)
        closure: set[str] = set()
        for cand in candidates:
            closure.add(cand)
            closure |= self.ancestors_of(cand) & universe
        return closure

    def order_candidates(
        self,
        candidates: Iterable[str],
        scores: dict[str, float],
        extra_edges: dict[str, set[str]] | None = None,
    ) -> list[str]:
        """对候选集合做依赖拓扑排序，返回处理顺序。

        规则（设计 5.1）：
        1. 候选之间的直接或传递依赖构成先后顺序：上游候选永远在前；
        2. 无依赖关系的候选按静态分数降序排列；
        3. 分数相同时按完整地址稳定排序；
        4. ``extra_edges`` 允许调度器把 tentative graph 新发现的
           "候选级上游" 边（dependent → {upstream, ...}）并入排序；
        5. 候选之间若因 extra_edges 形成循环，不伪造顺序，剩余候选
           按分数/地址规则附加在末尾（真正的图内循环候选在排序前就
           已被 SCC 识别并跳过，不会进入本接口）。

        覆盖 missing_formula 常量候选：它们同样以完整地址形式出现在
        图的节点中，与公式候选使用同一套规则。
        """
        cand_list = list(dict.fromkeys(candidates))
        cand_set = set(cand_list)

        # 用整图传递可达性构造候选间依赖边：候选 A 经中间单元格到达
        # 候选 B 时，A 也必须排在 B 前。
        up_map = {cand: self.ancestors_of(cand) & cand_set for cand in cand_list}
        if extra_edges:
            for dependent, upstreams in extra_edges.items():
                if dependent in cand_set:
                    up_map.setdefault(dependent, set()).update(
                        upstream for upstream in upstreams if upstream in cand_set
                    )
        edges: set[tuple[str, str]] = set()
        for dependent, upstreams in up_map.items():
            edges.update((upstream, dependent) for upstream in upstreams)

        # Kahn 算法 + 二叉堆：每次从"已就绪"（无未处理上游）的候选中
        # 取 (-score, address) 最小者，保证确定性与分数并列优先级。
        successors: dict[str, list[str]] = {cand: [] for cand in cand_list}
        in_degree: dict[str, int] = {cand: 0 for cand in cand_list}
        for upstream, dependent in edges:
            successors[upstream].append(dependent)
            in_degree[dependent] += 1

        heap = [
            (-scores.get(cand, 0.0), cand)
            for cand in cand_list
            if in_degree[cand] == 0
        ]
        heapq.heapify(heap)
        ordered: list[str] = []
        while heap:
            _, cand = heapq.heappop(heap)
            ordered.append(cand)
            for dependent in successors[cand]:
                in_degree[dependent] -= 1
                if in_degree[dependent] == 0:
                    heapq.heappush(heap, (-scores.get(dependent, 0.0), dependent))
        if len(ordered) < len(cand_list):
            # 防御分支：extra_edges 引入的循环导致部分候选无法排序。
            ordered_set = set(ordered)
            ordered.extend(
                sorted(
                    (cand for cand in cand_list if cand not in ordered_set),
                    key=lambda cand: (-scores.get(cand, 0.0), cand),
                )
            )
        return ordered
