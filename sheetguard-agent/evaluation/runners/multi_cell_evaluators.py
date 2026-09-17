"""multi-cell 批处理结果的集合级评测器（设计 10.3）。

评测器从"单个 target 是否正确"升级到集合级指标，全部为纯函数，
可在离线测试中直接调用。每个评测器接收：

- ``output``：agent 批量审计结果（audit dict）；
- ``expected_output``：multi-cell 数据集 gold（errors / expected_status）。

false repair 是核心安全指标：Agent 即使修复了多个正确目标，只要错误
修改了原本正常的单元格，也必须明确扣分。
"""
from __future__ import annotations

import json
from typing import Any

from langfuse import Evaluation

from sheetguard.spreadsheet.aggregates import normalize_template
from sheetguard.spreadsheet.formula_parser import to_relative_template
from sheetguard.spreadsheet.model import CellRef

# Agent 输出中各终态的集合键名。
_OUTCOME_KEYS = ("fixed", "failed", "dismissed", "unresolved", "skipped", "deferred")
_ALLOWED_BLOCKER_CAUSES = {"failed", "unresolved", "dependency_cycle"}
_DEFERRED_STATUSES = {
    "deferred_due_to_budget",
    "deferred_due_to_budget_dependency",
    "deferred_due_to_inactive_prerequisite",
}


def _cells(output: dict, key: str) -> set[str]:
    """取某终态列表里全部 target 地址。"""
    return {
        str(entry.get("target", "")).strip().lower()
        for entry in (output.get(key) or [])
        if isinstance(entry, dict) and entry.get("target")
    }


def _gold_cells(expected: dict) -> set[str]:
    """gold 错误单元格集合。"""
    return {
        str(err.get("target_cell", "")).strip().lower()
        for err in (expected.get("errors") or [])
        if isinstance(err, dict) and err.get("target_cell")
    }


def _expected_targets(expected: dict, key: str) -> set[str]:
    """workbook 级期望集合（expected_dismissed / expected_skipped / ...）。"""
    return {
        str(entry).strip().lower()
        for entry in (expected.get(key) or [])
    }


def _match_set_evaluator(name: str, key: str, actual: set[str], expected: set[str], note: str = "") -> "Evaluation":
    """期望集合与实际集合一致得 1，否则 0；无期望时视为 not_applicable。"""
    if not expected:
        return Evaluation(
            name=name,
            value=1.0,
            comment=f"not_applicable: gold 无 {key} 期望{'; ' + note if note else ''}",
        )
    missing = sorted(expected - actual)
    extra = sorted(actual - expected)
    matched = not missing and not extra
    comment = (
        f"expected={sorted(expected)} actual={sorted(actual)}"
        + (f" missing={missing}" if missing else "")
        + (f" extra={extra}" if extra else "")
        + (f"; {note}" if note else "")
    )
    return Evaluation(name=name, value=1.0 if matched else 0.0, comment=comment)


def _ratio(numerator: int, denominator: int) -> float:
    return 1.0 if denominator == 0 else numerator / denominator


def _semantic_formula(formula: str) -> str:
    """把公式归一化为语义等价形式（相对模板 + 去掉冗余括号）。"""
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula.strip()
    dummy = CellRef(sheet="X", row=1, col=1)
    try:
        template = to_relative_template(formula, dummy)
    except Exception:
        return formula.strip()
    # 与 Verifier pattern 检查共用 normalize_template，保证两侧的
    # 语义等价口径一致（空白/一元负括号变体不再判为不等价）。
    return normalize_template(template)


def _formula_maps(output: dict, expected: dict) -> tuple[dict[str, str], dict[str, str]]:
    """Agent 修复的 {target: formula} 与 gold 的 {target: gold_formula}。"""
    actual = {
        str(entry.get("target", "")).strip().lower(): str(entry.get("new_formula", "") or "")
        for entry in (output.get("fixed") or [])
        if isinstance(entry, dict) and entry.get("target")
    }
    gold = {
        str(err.get("target_cell", "")).strip().lower(): str(err.get("gold_formula", "") or "")
        for err in (expected.get("errors") or [])
        if isinstance(err, dict) and err.get("target_cell")
    }
    return actual, gold


def evaluate_static_candidate_counts(*, output: dict, expected_output: dict, **kwargs) -> Evaluation:
    """记录 static/seed/eligible/active/processed 计数（审计口径，验收 29）。"""
    return Evaluation(
        name="static_candidate_counts",
        value=float(output.get("static_candidate_count", 0)),
        comment=(
            f"static={output.get('static_candidate_count', 0)} "
            f"seed={output.get('seed_candidate_count', 0)} "
            f"eligible={output.get('eligible_candidate_count', 0)} "
            f"active={output.get('initial_active_candidate_count', 0)} "
            f"processed={output.get('processed_candidate_count', 0)}"
        ),
    )


def evaluate_confirmation_recall(*, output: dict, expected_output: dict, **kwargs) -> Evaluation:
    """gold 错误单元格中被 Agent 实际调查（confirmed/failed）的覆盖率。"""
    gold = _gold_cells(expected_output)
    investigated = _cells(output, "fixed") | _cells(output, "failed")
    covered = gold & investigated
    return Evaluation(
        name="confirmation_recall",
        value=_ratio(len(covered), len(gold)),
        comment=f"covered {len(covered)}/{len(gold)} gold errors",
    )


def evaluate_dismissal_accuracy(*, output: dict, expected_output: dict, **kwargs) -> Evaluation:
    """被 dismissed 的候选是否都确实为 false positive（非 gold 错误）。"""
    dismissed = _cells(output, "dismissed")
    gold = _gold_cells(expected_output)
    false_dismissals = dismissed & gold
    return Evaluation(
        name="dismissal_accuracy",
        value=_ratio(len(dismissed) - len(false_dismissals), len(dismissed)),
        comment=f"dismissed={len(dismissed)} false_dismissals={sorted(false_dismissals)}",
    )


def evaluate_repair_recall(*, output: dict, expected_output: dict, **kwargs) -> Evaluation:
    """gold 错误单元格中被成功修复（fixed）的比例。"""
    fixed = _cells(output, "fixed")
    gold = _gold_cells(expected_output)
    return Evaluation(
        name="repair_recall",
        value=_ratio(len(fixed & gold), len(gold)),
        comment=f"fixed_gold={len(fixed & gold)}/{len(gold)}",
    )


def evaluate_repair_precision(*, output: dict, expected_output: dict, **kwargs) -> Evaluation:
    """被修复的目标中有多少确实是 gold 错误。"""
    fixed = _cells(output, "fixed")
    gold = _gold_cells(expected_output)
    return Evaluation(
        name="repair_precision",
        value=_ratio(len(fixed & gold), len(fixed)),
        comment=f"fixed={len(fixed)} true_gold={len(fixed & gold)}",
    )


def evaluate_false_repair(*, output: dict, expected_output: dict, **kwargs) -> Evaluation:
    """false repair 安全指标：修复了非 gold 单元格的数量与比例。"""
    fixed = _cells(output, "fixed")
    gold = _gold_cells(expected_output)
    false_repairs = fixed - gold
    return Evaluation(
        name="false_repair_count",
        value=float(len(false_repairs)),
        comment=(
            f"false_repair_rate={_ratio(len(false_repairs), len(fixed)):.3f} "
            f"cells={sorted(false_repairs)}"
        ),
    )


def evaluate_formula_accuracy(*, output: dict, expected_output: dict, **kwargs) -> Evaluation:
    """每个 gold 错误单元格的修复公式与 gold 公式逐字匹配。

    "逐字"不包含空白差异：`=VLOOKUP(C10,C10:D10,2,FALSE)` 与
    `=VLOOKUP(C10, C10:D10, 2, FALSE)` 是同一个公式（09-16 complete
    实验 complete_008/010 教训）；函数/常量层面的差异仍判 0。
    """
    actual, gold = _formula_maps(output, expected_output)
    hits = sum(
        1 for cell, formula in gold.items()
        if _whitespace_free(actual.get(cell, "")) == _whitespace_free(formula)
    )
    return Evaluation(
        name="formula_accuracy",
        value=_ratio(hits, len(gold)),
        comment=f"exact_hits={hits}/{len(gold)}",
    )


def _whitespace_free(text: str) -> str:
    """去掉全部空白字符，供逐字匹配比较（大小写保留）。"""
    return "".join(str(text).split())


def evaluate_formula_equivalent(*, output: dict, expected_output: dict, **kwargs) -> Evaluation:
    """每个 gold 错误单元格的修复公式与 gold 公式语义等价。"""
    actual, gold = _formula_maps(output, expected_output)
    hits = sum(
        1
        for cell, formula in gold.items()
        if _semantic_formula(actual.get(cell, "")) == _semantic_formula(formula)
    )
    return Evaluation(
        name="formula_equivalent",
        value=_ratio(hits, len(gold)),
        comment=f"semantic_hits={hits}/{len(gold)}",
    )


def evaluate_dependency_skip_accuracy(*, output: dict, expected_output: dict, **kwargs) -> Evaluation:
    """skipped 的 blocked_by 记录合法根 blocker（failed/unresolved/dependency_cycle）。"""
    issues: list[str] = []
    for entry in output.get("skipped") or []:
        if not isinstance(entry, dict):
            continue
        target = entry.get("target")
        if entry.get("status") != "skipped_due_to_dependency":
            issues.append(f"{target}:status={entry.get('status')}")
        blockers = entry.get("blocked_by") or []
        if not blockers:
            issues.append(f"{target}:no_blocker")
        for blocker in blockers:
            if blocker.get("cause") not in _ALLOWED_BLOCKER_CAUSES:
                issues.append(f"{target}:cause={blocker.get('cause')}")
    return Evaluation(
        name="dependency_skip_accuracy",
        value=1.0 if not issues else 0.0,
        comment=f"issues={issues[:5]}",
    )


def evaluate_deferred_counts(*, output: dict, expected_output: dict, **kwargs) -> Evaluation:
    """记录 deferred 总数与 inactive prerequisite deferred 数（验收 14）。"""
    deferred = [
        entry for entry in (output.get("deferred") or []) if isinstance(entry, dict)
    ]
    inactive = sum(
        1 for entry in deferred
        if entry.get("status") == "deferred_due_to_inactive_prerequisite"
    )
    with_prereqs = sum(1 for entry in deferred if entry.get("required_prerequisites"))
    return Evaluation(
        name="deferred_counts",
        value=float(len(deferred)),
        comment=(
            f"deferred={len(deferred)} inactive_prerequisite={inactive} "
            f"with_required_prerequisites={with_prereqs}"
        ),
    )


def evaluate_outcome_counts(*, output: dict, expected_output: dict, **kwargs) -> Evaluation:
    """记录各终态数量（fixed/failed/dismissed/unresolved/skipped/deferred）。"""
    counts = {key: len(output.get(key) or []) for key in _OUTCOME_KEYS}
    return Evaluation(
        name="outcome_counts",
        value=float(sum(counts.values())),
        comment=json.dumps(counts, ensure_ascii=False),
    )


def evaluate_workbook_status(*, output: dict, expected_output: dict, **kwargs) -> Evaluation:
    """workbook-level 最终状态是否与 gold 期望一致。"""
    expected_status = str(expected_output.get("expected_status", "")).strip()
    actual_status = str(output.get("status", "")).strip()
    match = actual_status == expected_status
    return Evaluation(
        name="workbook_status",
        value=float(match),
        comment=f"expected={expected_status!r} actual={actual_status!r}",
    )


def evaluate_attempt_cost(*, output: dict, expected_output: dict, **kwargs) -> Evaluation:
    """记录本案例的修复尝试总数（成本指标）。"""
    return Evaluation(
        name="attempts",
        value=float(len(output.get("repair_history") or [])),
        comment=f"repair_history_entries={len(output.get('repair_history') or [])}",
    )


def evaluate_expected_dismissal(*, output: dict, expected_output: dict, **kwargs) -> Evaluation:
    """被 dismiss 的候选集合是否与 gold 期望（expected_dismissed）一致。

    难题数据集用 expected_dismissed 描述"结构异常但值正确的单元格"，
    Agent 既不能把它们当错误修复（false repair），也不能视而不见。
    """
    return _match_set_evaluator(
        "expected_dismissal_match",
        "expected_dismissed",
        _cells(output, "dismissed"),
        _expected_targets(expected_output, "expected_dismissed"),
    )


def evaluate_expected_skip(*, output: dict, expected_output: dict, **kwargs) -> Evaluation:
    """被 skipped 的候选集合是否与 gold 期望（expected_skipped）一致。

    依赖循环 / 上游阻塞的候选必须 skip 而不是修复；comment 附带
    blocked_by 根因供人工核对。
    """
    skipped = output.get("skipped") or []
    causes = {
        str(entry.get("target", "")).strip().lower(): [
            blocker.get("cause") for blocker in (entry.get("blocked_by") or [])
            if isinstance(blocker, dict)
        ]
        for entry in skipped
        if isinstance(entry, dict) and entry.get("target")
    }
    return _match_set_evaluator(
        "expected_skip_match",
        "expected_skipped",
        set(causes),
        _expected_targets(expected_output, "expected_skipped"),
        note=f"causes={causes}" if causes else "",
    )


def evaluate_expected_deferred(*, output: dict, expected_output: dict, **kwargs) -> Evaluation:
    """被 deferred 的候选集合是否与 gold 期望（expected_deferred）一致。

    预算延期案例用 expected_deferred 描述"超过 max_candidates 之外必须
    搁置的候选"；comment 附带实际 status 供核对。
    """
    statuses = {
        str(entry.get("target", "")).strip().lower(): entry.get("status")
        for entry in (output.get("deferred") or [])
        if isinstance(entry, dict) and entry.get("target")
    }
    return _match_set_evaluator(
        "expected_deferred_match",
        "expected_deferred",
        set(statuses),
        _expected_targets(expected_output, "expected_deferred"),
        note=f"statuses={statuses}" if statuses else "",
    )
