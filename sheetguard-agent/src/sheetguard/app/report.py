"""审计字典 → 面向人的 Markdown 报告渲染器。

纯函数：输入 ``run``/``repair`` 产出的 audit 字典，输出 Markdown 文本。
渲染是确定性的（不含时间戳等易变内容），便于 golden-file 测试；
配置指纹随 ``audit["config"]`` 完整呈现在报告头部。

设计目标：财务人员不读 JSON 也能逐格复核——每个候选给出终态、
修复前后公式、调查证据/根因、验证结果与卡点原因。
"""
from __future__ import annotations

_CHECK_MARKS = {"passed": "✓", "failed": "✗", "not_applicable": "—"}


def _fmt_formula(formula: object) -> str:
    """公式包进反引号；清洗换行与反引号，避免破坏 Markdown 结构。"""
    if formula is None:
        return "—"
    text = str(formula).replace("`", "'").replace("\n", " ").strip()
    return f"`{text}`"


def _checks_line(verification: dict | None) -> str | None:
    """把六项检查的三态结果压成一行：`syntax ✓ · pattern ✗（原因）· …`。"""
    checks = (verification or {}).get("checks") or {}
    if not checks:
        return None
    parts: list[str] = []
    for name, info in checks.items():
        status = info.get("status")
        mark = _CHECK_MARKS.get(status, "?")
        reason = info.get("reason")
        suffix = f"（{reason}）" if reason and status != "passed" else ""
        parts.append(f"{name} {mark}{suffix}")
    return " · ".join(parts)


def _orig_formula_line(old_formula: object) -> str:
    """渲染「原公式」行；硬编码常量场景额外说明（missing_formula）。"""
    if old_formula is None:
        return "- 原公式：—（该格原本没有公式，missing_formula 场景）"
    text = str(old_formula)
    if text.startswith("="):
        return f"- 原公式：{_fmt_formula(text)}"
    return f"- 原公式：{_fmt_formula(text)}（原为硬编码常量，missing_formula 场景）"


def _error_lines(entry: dict) -> list[str]:
    """渲染「存在的错误」：error_type + 可疑分/信号 + 诊断文字。"""
    hypothesis = entry.get("hypothesis")
    hypothesis = hypothesis if isinstance(hypothesis, dict) else {}
    bits: list[str] = []
    if hypothesis.get("error_type"):
        bits.append(f"`{hypothesis['error_type']}`")
    if entry.get("score") is not None:
        bits.append(f"可疑分 {entry['score']}")
    if entry.get("signals"):
        bits.append("信号 " + ", ".join(entry["signals"]))
    if not bits:
        return []
    lines = ["- 存在的错误：" + " · ".join(bits)]
    if hypothesis.get("hypothesis"):
        lines.append(f"  - 诊断：{hypothesis['hypothesis']}")
    return lines


def _section(lines: list[str], title: str, items: list, render_item) -> None:
    """有内容才出节：`## 标题（N）` + 逐项渲染。"""
    if not items:
        return
    lines.append("")
    lines.append(f"## {title}（{len(items)}）")
    for item in items:
        rendered = render_item(item)
        if rendered:
            lines.extend(rendered)


def _blocked_by_line(blocked_by: list | None) -> str | None:
    """根因链渲染：`P&L!D2 (failed) → …`；空列表返回 None。"""
    entries = blocked_by or []
    if not entries:
        return None
    parts = []
    for entry in entries:
        if isinstance(entry, dict):
            target = entry.get("target", "?")
            cause = entry.get("cause")
            parts.append(f"{target} ({cause})" if cause else str(target))
        else:
            parts.append(str(entry))
    return " → ".join(parts)


def render_report(audit: dict) -> str:
    """把审计字典渲染成 Markdown；缺键一律容错（baseline 审计也能出报告）。"""
    lines: list[str] = ["# SheetGuard 修复报告", ""]
    lines.append(f"- 工作簿：{audit.get('workbook', '—')}")
    lines.append(f"- 状态：`{audit.get('status') or '—'}`")
    if audit.get("repaired_workbook"):
        lines.append(f"- 修复副本：`{audit['repaired_workbook']}`")

    counts = [
        ("静态", audit.get("static_candidate_count")),
        ("seed", audit.get("seed_candidate_count")),
        ("冻结批次", audit.get("initial_active_candidate_count")),
        ("实际处理", audit.get("processed_candidate_count")),
    ]
    if any(value is not None for _, value in counts):
        parts = [
            f"{label} {value}" for label, value in counts if value is not None
        ]
        lines.append(f"- 候选：{' · '.join(parts)}")
    terminal = (
        f"✅ fixed {len(audit.get('fixed') or [])} · "
        f"✅ certified {len(audit.get('certified') or [])} · "
        f"❌ failed {len(audit.get('failed') or [])} · "
        f"👀 dismissed {len(audit.get('dismissed') or [])} · "
        f"❓ unresolved {len(audit.get('unresolved') or [])} · "
        f"🚧 skipped {len(audit.get('skipped') or [])} · "
        f"⏸ deferred {len(audit.get('deferred') or [])}"
    )
    lines.append(f"- 终态：{terminal}")
    if audit.get("warning"):
        lines.append("")
        lines.append(f"> ⚠️ {audit['warning']}")

    config = audit.get("config")
    if config:
        lines.append("")
        lines.append("## 配置指纹")
        for key in (
            "min_anomaly_score",
            "max_candidates",
            "max_attempts",
            "max_localizer_iterations",
            "max_localizer_tool_calls",
            "max_review_rounds",
        ):
            if key in config:
                lines.append(f"- {key}: {config[key]}")
        weights = config.get("signal_weights")
        if weights:
            rendered = ", ".join(f"{name}={weight}" for name, weight in weights.items())
            lines.append(f"- signal_weights: {rendered}")

    def _render_fixed(entry: dict) -> list[str]:
        out = [f"### {entry.get('target', '—')}", ""]
        out.append(_orig_formula_line(entry.get("old_formula")))
        out.append(f"- 新公式：{_fmt_formula(entry.get('new_formula'))}")
        error_lines = _error_lines(entry)
        if error_lines:
            out.extend(error_lines)
        checks = _checks_line(entry.get("verification"))
        if checks:
            out.append(f"- 验证：{checks}")
        return out

    def _render_failed(entry: dict) -> list[str]:
        out = [f"### {entry.get('target', '—')}", ""]
        out.append(_orig_formula_line(entry.get("old_formula")))
        if entry.get("last_formula") is not None:
            out.append(f"- 最后尝试：{_fmt_formula(entry['last_formula'])}")
        error_lines = _error_lines(entry)
        if error_lines:
            out.extend(error_lines)
        if entry.get("reason"):
            out.append(f"- 原因：{entry['reason']}")
        if entry.get("attempts") is not None:
            out.append(f"- 尝试次数：{entry['attempts']}")
        checks = _checks_line(entry.get("verification"))
        if checks:
            out.append(f"- 验证：{checks}")
        return out

    def _render_certified(entry: dict) -> list[str]:
        formula = _fmt_formula(entry.get("formula"))
        round_no = entry.get("round")
        suffix = f"（round {round_no} 认证）" if round_no is not None else ""
        return [f"- **{entry.get('target', '—')}**：{formula}{suffix}"]

    _section(lines, "✅ 用户已认证", audit.get("certified") or [], _render_certified)
    _section(lines, "✅ 修复成功", audit.get("fixed") or [], _render_fixed)
    _section(lines, "❌ 修复失败", audit.get("failed") or [], _render_failed)

    def _render_simple(entry: dict) -> list[str]:
        target = entry.get("target", "—")
        reason = entry.get("reason")
        return [f"- **{target}**{('：' + reason) if reason else ''}"]

    _section(lines, "👀 已排除（调查确认不是真错误）", audit.get("dismissed") or [], _render_simple)
    _section(lines, "❓ 未解决", audit.get("unresolved") or [], _render_simple)

    def _render_skipped(entry: dict) -> list[str]:
        out = [f"- **{entry.get('target', '—')}**"]
        if entry.get("reason"):
            out.append(f"  - 原因：{entry['reason']}")
        root = _blocked_by_line(entry.get("blocked_by"))
        if root:
            out.append(f"  - 根因：{root}")
        return out

    _section(lines, "🚧 已跳过", audit.get("skipped") or [], _render_skipped)

    def _render_deferred(entry: dict) -> list[str]:
        out = [f"- **{entry.get('target', '—')}**（{entry.get('status', '—')}）"]
        if entry.get("score") is not None:
            out.append(f"  - 可疑分：{entry['score']}")
        if entry.get("signals"):
            out.append(f"  - 信号：{', '.join(entry['signals'])}")
        if entry.get("required_prerequisites"):
            prereq = ", ".join(
                item.get("target", str(item))
                if isinstance(item, dict) else str(item)
                for item in entry["required_prerequisites"]
            )
            out.append(f"  - 欠条：下次运行先修 {prereq}")
        if entry.get("reason"):
            out.append(f"  - 原因：{entry['reason']}")
        return out

    _section(lines, "⏸ 延期（下次运行再修）", audit.get("deferred") or [], _render_deferred)

    history = audit.get("repair_history") or []
    if history:
        lines.append("")
        lines.append(f"## 附录：修复尝试流水（{len(history)}）")
        for item in history:
            formula = _fmt_formula(item.get("formula"))
            lines.append(f"- {item.get('target', '—')} · {formula} · {item.get('outcome', '—')}")

    if audit.get("localization_warning"):
        lines.append("")
        lines.append(f"> ℹ️ 定位警告：{audit['localization_warning']}")
    if audit.get("error"):
        lines.append("")
        lines.append(f"> ❌ 运行错误：{audit['error']}")
    return "\n".join(lines) + "\n"
