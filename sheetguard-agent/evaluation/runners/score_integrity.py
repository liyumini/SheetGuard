"""实验结果完整性校验：服务端评分对账回填 + 实验记录分裂检测。

Langfuse SDK 的评分经后台队列批量上报：网络抖动时整批重试 3 次
（约 7 秒窗口）后静默丢弃，只写文件日志；trace/observation 走另一条
OTel 批量管线，基本不丢。结果是 run_experiment 的内存结果完整、
服务端评分残缺，而终端 results.format() 用的是内存结果——丢失不可见
（2026-09-11 06:49 实验就丢了最后一个案例的全部 13 条评分）。

另外，服务端按 run_name upsert 实验记录时偶发竞态，会把一次运行拆成
多条同名记录、item 分布分裂（同日两次 multi-cell 运行均复现，8+2）。

本模块在实验结束后用服务端数据对账：

1. ``verify_item_results``：按 traceId 回查已入库的评分名（scores_v3
   分页接口，限流/网络错误退避重试），与内存评测结果对账，缺失项用
   同步 API 直写补录；复查最多 ``rounds`` 轮；仍有缺失则抛
   ``ScoreIntegrityError``，而不是留一份静默残缺的服务端数据。
2. ``verify_experiment_records``：按精确 run_name 回查实验记录，发现
   拆分或条目数异常时打印告警（只提示不失败，数据本身仍可查）。
3. ``evaluations_from_observations``：对已经跑完的历史实验，从 trace 的
   EVALUATOR observation 输出里重建评测值（评测器当时已算出并写入
   observation，值完整），再交给 ``ensure_trace_scores`` 补录。
"""
from __future__ import annotations

import json
import time
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone

EVALUATION_SPAN_NAME = "experiment-item-task"


class ScoreIntegrityError(RuntimeError):
    """评分回填重试耗尽后仍有缺失。"""


@dataclass(frozen=True)
class ScoreExpectation:
    """一条预期评分：与 ``Evaluation`` 字段一一对应，用于对账与重建。"""

    name: str
    value: float
    data_type: str
    comment: str | None = None


def _expectation_from_evaluation(evaluation) -> ScoreExpectation:
    """把 SDK 的 Evaluation 对象/字典归一化为 ScoreExpectation。"""
    name = evaluation.get("name") if isinstance(evaluation, dict) else evaluation.name
    value = evaluation.get("value") if isinstance(evaluation, dict) else evaluation.value
    comment = (
        evaluation.get("comment") if isinstance(evaluation, dict) else evaluation.comment
    )
    data_type = (
        evaluation.get("data_type") if isinstance(evaluation, dict) else evaluation.data_type
    )
    if not data_type:
        # create_score 不传 dataType 时服务端按值类型推断；这里显式化，
        # 避免 bool 被当成 0/1 数字入库（bool 必须先于 int 判断）。
        if isinstance(value, bool):
            data_type = "BOOLEAN"
        elif isinstance(value, (int, float)):
            data_type = "NUMERIC"
        else:
            data_type = "TEXT"
    return ScoreExpectation(
        name=name,
        value=float(value),
        data_type=data_type,
        comment=comment,
    )


def _with_retry(fn, *, attempts: int = 5, delay: float = 3.0, description: str = ""):
    """限流/网络错误退避重试。429 时尊重服务端 retryAfter。"""
    last_error: Exception | None = None
    for attempt in range(attempts):
        try:
            return fn()
        except Exception as exc:  # noqa: BLE001 - Fern 客户端异常类型随版本变化
            last_error = exc
            status = getattr(exc, "status", None)
            wait = delay * (attempt + 1)
            if status == 429:
                retry_after = None
                details = getattr(exc, "body", None)
                if isinstance(details, dict):
                    retry_after = (details.get("details") or {}).get("retryAfterSeconds")
                wait = float(retry_after or wait) + 0.5
            if attempt + 1 < attempts:
                time.sleep(wait)
    assert last_error is not None
    raise last_error


def _iter_trace_scores(langfuse, trace_id: str) -> list:
    """分页拉取某 trace 的全部评分（scores_v3）。"""
    api = langfuse.api
    from_timestamp = datetime(2026, 1, 1, tzinfo=timezone.utc) - timedelta(days=365)
    scores: list = []
    cursor: str | None = None
    while True:
        page = _with_retry(
            lambda: api.scores_v3.get_many_v3(
                trace_id=trace_id,
                from_timestamp=from_timestamp,
                limit=100,
                cursor=cursor,
            ),
            description=f"fetch scores trace={trace_id[:8]}",
        )
        data = getattr(page, "data", None) or []
        scores.extend(data)
        meta = getattr(page, "meta", None)
        cursor = getattr(meta, "cursor", None)
        if not cursor or not data:
            break
    return scores


def _fetch_observations(langfuse, trace_id: str) -> list:
    """分页拉取某 trace 的 observations（含 io 组，EVALUATOR 输出在其中）。

    observations v2 接口默认只返回 core+basic，不请求 ``io`` 组时
    ``output`` 恒为 None；v1 路径限流紧，这里统一带退避重试。
    """
    api = langfuse.api
    from datetime import datetime as _dt

    observations: list = []
    cursor: str | None = None
    while True:
        page = _with_retry(
            lambda: api.observations.get_many(
                trace_id=trace_id,
                from_start_time=_dt(2026, 1, 1, tzinfo=timezone.utc),
                limit=100,
                fields="core,basic,io,metadata",
                cursor=cursor,
            ),
            description=f"fetch observations {trace_id[:8]}",
        )
        data = getattr(page, "data", None) or []
        observations.extend(data)
        meta = getattr(page, "meta", None)
        cursor = getattr(meta, "cursor", None)
        if not cursor or not data:
            break
    return observations


def evaluations_from_observations(langfuse, trace_id: str) -> list[ScoreExpectation]:
    """从 trace 的 EVALUATOR observation 输出重建评测值。

    run_experiment 把每个评测器的返回值序列化进名为评测器函数名的
    EVALUATOR observation ``output``（形如 ``[{"name", "value", "comment"}]``）；
    评分入库失败时这些 observation 仍在，可用于补录。
    """
    observations = _fetch_observations(langfuse, trace_id)
    expectations: list[ScoreExpectation] = []
    for obs in observations:
        if getattr(obs, "type", None) != "EVALUATOR":
            continue
        output = getattr(obs, "output", None)
        # observations v2 的 input/output 是原始 JSON 字符串；
        # run_experiment 序列化评测器结果为数组，容忍单个对象的写法。
        if isinstance(output, str):
            try:
                output = json.loads(output)
            except ValueError:
                output = None
        if isinstance(output, dict):
            output = [output]
        entries = output if isinstance(output, list) else []
        for entry in entries:
            if isinstance(entry, dict) and entry.get("name") is not None:
                expectations.append(_expectation_from_evaluation(entry))
    return expectations


def _task_observation_id(langfuse, trace_id: str) -> str | None:
    """找 experiment-item-task observation 的 id，评分跟随原位入库。"""
    try:
        observations = _fetch_observations(langfuse, trace_id)
    except Exception:
        return None
    for obs in observations:
        if getattr(obs, "name", None) == EVALUATION_SPAN_NAME:
            return getattr(obs, "id", None)
    return None


def _create_score_direct(langfuse, trace_id: str, exp: ScoreExpectation, observation_id: str | None) -> None:
    """同步直写一条评分（绕过后台队列）。

    ``create_score`` 入后台队列后由消费线程批量上报，失败重试耗尽即
    静默丢弃（见模块注释）；回填必须同步逐条 POST 并让错误显式抛出，
    由 ``_with_retry`` 统一退避重试。
    """
    api = langfuse.api
    _with_retry(
        lambda: api.scores.create(
            name=exp.name,
            value=exp.value,
            trace_id=trace_id,
            observation_id=observation_id,
            data_type=exp.data_type,
            comment=exp.comment,
            source="EVAL",
        ),
        description=f"create score {exp.name}",
    )


def ensure_trace_scores(
    langfuse,
    trace_id: str,
    expectations: list[ScoreExpectation],
    *,
    rounds: int = 3,
    delay: float = 3.0,
) -> dict:
    """对账单个 trace 的评分并回填缺失项，返回完整性报告。

    每轮：拉取服务端已有评分名 → 缺失的同步直写（逐条重试）→ 复查。
    ``rounds`` 轮后仍有缺失则返回 ``complete=False`` 的报告。
    """
    report: dict = {"trace_id": trace_id, "expected": len(expectations), "backfilled": 0}
    for attempt in range(rounds):
        existing = {
            score.name
            for score in _iter_trace_scores(langfuse, trace_id)
        }
        missing = [exp for exp in expectations if exp.name not in existing]
        report["recorded"] = len(existing)
        if not missing:
            report["complete"] = True
            report["missing"] = []
            return report
        observation_id = _task_observation_id(langfuse, trace_id)
        for exp in missing:
            try:
                _create_score_direct(langfuse, trace_id, exp, observation_id)
                report["backfilled"] += 1
            except Exception:
                # 单条直写失败留给下一轮重试；全部轮次耗尽后统一以
                # ScoreIntegrityError 显式失败。
                continue
        time.sleep(delay)
    existing = {
        score.name for score in _iter_trace_scores(langfuse, trace_id)
    }
    report["recorded"] = len(existing)
    report["complete"] = not any(exp.name not in existing for exp in expectations)
    report["missing"] = sorted(exp.name for exp in expectations if exp.name not in existing)
    return report


def verify_experiment_records(
    langfuse,
    *,
    run_name: str,
    expected_items: int,
    console=None,
) -> list:
    """检查一次 run_experiment 是否被服务端拆成多条同名实验记录。

    Langfuse 服务端按 run_name upsert 实验记录，偶发竞态会把一次运行
    拆成多条同名记录、item 分布分裂（2026-09-11 两次 multi-cell 运行
    均复现，8+2 拆分；升级 SDK 4.15.2 无相关修复，属服务端行为）。
    此处按精确 run_name 回查：多于一条记录或条目总数不足时打印告警，
    让使用者知道该次运行在 UI 里是割裂的，需按 id 分别查看。
    """
    api = langfuse.api
    from_start_time = datetime.now(timezone.utc) - timedelta(days=1)
    page = _with_retry(
        lambda: api.experiments.list(
            from_start_time=from_start_time, name=run_name, fields="core", limit=100
        ),
        description=f"list experiments {run_name[:40]}",
    )
    records = [e for e in (getattr(page, "data", None) or []) if getattr(e, "name", None) == run_name]
    total_items = sum(int(getattr(e, "item_count", 0) or 0) for e in records)
    if len(records) > 1:
        detail = "; ".join(
            f"{getattr(e, 'id', '?')[:16]} items={getattr(e, 'item_count', '?')} "
            f"({getattr(e, 'start_time', '?')}..{getattr(e, 'end_time', '?')})"
            for e in records
        )
        if console is not None:
            console.print(
                f"[yellow]警告：本次实验被服务端拆成 {len(records)} 条同名记录，"
                f"条目合计 {total_items}/{expected_items}：{detail}[/yellow]"
            )
    elif records and total_items != expected_items and console is not None:
        console.print(
            f"[yellow]警告：实验记录条目数 {total_items} 与数据集 {expected_items} 不一致。[/yellow]"
        )
    return records


def verify_item_results(
    langfuse,
    item_results,
    *,
    rounds: int = 3,
    delay: float = 3.0,
    console=None,
) -> list[dict]:
    """对整个实验的 ``item_results`` 做服务端评分对账 + 回填。

    回填值来自内存评测结果（评测器当时算出的值），与丢失前一致。
    任何 trace 经 ``rounds`` 轮回填后仍有缺失时抛 ``ScoreIntegrityError``，
    避免残缺数据静默流入后续对比。
    """
    reports: list[dict] = []
    incomplete: list[dict] = []
    for item_result in item_results:
        trace_id = getattr(item_result, "trace_id", None)
        evaluations = getattr(item_result, "evaluations", None) or []
        if not trace_id or not evaluations:
            continue
        expectations = [_expectation_from_evaluation(e) for e in evaluations]
        report = ensure_trace_scores(
            langfuse, trace_id, expectations, rounds=rounds, delay=delay
        )
        reports.append(report)
        if console is not None and not report["complete"]:
            console.print(
                f"[red]评分完整性异常 {trace_id[:12]}：缺失 "
                f"{report['missing']}[/red]"
            )
        if not report["complete"]:
            incomplete.append(report)
    if incomplete:
        raise ScoreIntegrityError(
            f"{len(incomplete)}/{len(reports)} 个 trace 的评分经 {rounds} 轮回填仍不完整："
            + "; ".join(f"{r['trace_id'][:8]} 缺 {r['missing']}" for r in incomplete)
        )
    return reports
