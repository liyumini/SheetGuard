"""SheetGuard 运行配置。

把散在代码里的行为常量收拢为一处，并支持从 YAML 读取：

    - 调度参数：min_anomaly_score / max_candidates
    - 重试与调查预算：max_attempts / max_localizer_iterations / max_localizer_tool_calls
    - 六类异常信号的嫌疑分权重（signal_weights）

加载优先级：显式覆盖（CLI 参数） > YAML 配置文件 > 内置默认值。

最终生效配置通过 ``SheetGuardConfig.to_dict()`` 随审计记录与 Langfuse
根 observation 输出——每次运行自带完整配置指纹，跨实验对比时才能
确认"指标差异来自代码/提示词，而不是悄悄改过的配置"。

YAML 结构（所有键可选；未知键直接报错，不做静默忽略）：

    min_anomaly_score: 0.5
    max_candidates: 10
    max_attempts: 3
    max_localizer_iterations: 8
    max_localizer_tool_calls: 20
    max_review_rounds: 3
    signal_weights:
      missing_formula: 0.6
      pattern_anomaly: 0.5
      neighbor_mismatch: 0.3
      range_boundary: 0.5
      dependency_anomaly: 0.4
"""
from __future__ import annotations

import os
from dataclasses import asdict, dataclass, field, fields
from pathlib import Path
from typing import Any

import yaml

MAX_ATTEMPTS = 3  # 单个候选含首次尝试在内的最大 repair attempts
MAX_RETRIES = MAX_ATTEMPTS  # 兼容旧名：--max-retries 内部统一映射为 max_attempts
MAX_LOCALIZER_ITERATIONS = 8  # 调查阶段最多与模型交互 8 轮
MAX_LOCALIZER_TOOL_CALLS = 20  # 调查阶段工具执行次数的默认值（无硬钳上限）
DEFAULT_MIN_ANOMALY_SCORE = 0.0  # seed candidates 静态可疑分阈值默认值
DEFAULT_MAX_CANDIDATES = 10  # 一次运行进入 active batch 的候选数量默认上限
DEFAULT_MAX_REVIEW_ROUNDS = 3  # 用户审查反馈循环的最大轮数；超限 → rejected_exhausted

# 信号名与默认权重：与 anomaly_detector 的 _add_signal 调用点一一对应。
# singleton_sum_boundary 记账时使用 "range_boundary" 名，共用同一权重。
DEFAULT_SIGNAL_WEIGHTS: dict[str, float] = {
    "missing_formula": 0.6,
    "pattern_anomaly": 0.5,
    "neighbor_mismatch": 0.3,
    "range_boundary": 0.5,
    "dependency_anomaly": 0.4,
}


class ConfigError(ValueError):
    """配置文件或覆盖值非法。"""


@dataclass
class SheetGuardConfig:
    """一次运行的全部行为参数。"""

    min_anomaly_score: float = DEFAULT_MIN_ANOMALY_SCORE
    max_candidates: int = DEFAULT_MAX_CANDIDATES
    max_attempts: int = MAX_ATTEMPTS
    max_localizer_iterations: int = MAX_LOCALIZER_ITERATIONS
    max_localizer_tool_calls: int = MAX_LOCALIZER_TOOL_CALLS
    max_review_rounds: int = DEFAULT_MAX_REVIEW_ROUNDS
    signal_weights: dict[str, float] = field(
        default_factory=lambda: dict(DEFAULT_SIGNAL_WEIGHTS)
    )

    def to_dict(self) -> dict[str, Any]:
        """转换为普通字典，供审计 JSON 与 Langfuse metadata 记录。"""
        return asdict(self)


_CONFIG_FIELDS = {f.name for f in fields(SheetGuardConfig)}
_INT_KEYS = (
    "max_candidates",
    "max_attempts",
    "max_localizer_iterations",
    "max_localizer_tool_calls",
    "max_review_rounds",
)


def _check_int(name: str, value: Any, minimum: int) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        raise ConfigError(f"{name} must be an integer, got: {value!r}")
    if value < minimum:
        raise ConfigError(f"{name} must be >= {minimum}, got: {value}")
    return int(value)


def _check_float(name: str, value: Any, low: float, high: float) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ConfigError(f"{name} must be a number, got: {value!r}")
    number = float(value)
    if not low <= number <= high:
        raise ConfigError(f"{name} must be within [{low}, {high}], got: {number}")
    return number


def load_config(path: str | Path | None = None, **overrides: Any) -> SheetGuardConfig:
    """按「覆盖值 > YAML > 默认值」的优先级解析出最终配置。

    ``overrides`` 中的 None 值表示"未提供"，跳过（让 YAML 或默认值生效），
    因此 CLI 可以放心把未填写的选项以 None 传入。
    """
    config = SheetGuardConfig()
    if path is not None:
        file_path = Path(path)
        if not file_path.is_file():
            raise ConfigError(f"config file not found: {file_path}")
        try:
            data = yaml.safe_load(file_path.read_text(encoding="utf-8"))
        except yaml.YAMLError as exc:
            raise ConfigError(f"invalid YAML in {file_path}: {exc}") from exc
        if data is None:
            data = {}
        if not isinstance(data, dict):
            raise ConfigError(f"config file must contain a mapping: {file_path}")
        unknown = sorted(set(data) - _CONFIG_FIELDS)
        if unknown:
            raise ConfigError(
                f"unknown config keys: {unknown}; valid keys: {sorted(_CONFIG_FIELDS)}"
            )
        raw_weights = data.get("signal_weights")
        if raw_weights is not None:
            if not isinstance(raw_weights, dict):
                raise ConfigError("signal_weights must be a mapping of signal -> weight")
            unknown_weights = sorted(set(raw_weights) - set(DEFAULT_SIGNAL_WEIGHTS))
            if unknown_weights:
                raise ConfigError(
                    f"unknown signal names: {unknown_weights}; "
                    f"valid: {sorted(DEFAULT_SIGNAL_WEIGHTS)}"
                )
            for name, weight in raw_weights.items():
                config.signal_weights[name] = _check_float(
                    f"signal_weights.{name}", weight, 0.0, 1.0
                )
        if data.get("min_anomaly_score") is not None:
            config.min_anomaly_score = _check_float(
                "min_anomaly_score", data["min_anomaly_score"], 0.0, 1.0
            )
        for key in _INT_KEYS:
            if data.get(key) is not None:
                setattr(config, key, _check_int(key, data[key], 0 if key == "max_localizer_tool_calls" else 1))
    for key, value in overrides.items():
        if value is None or key not in _CONFIG_FIELDS:
            continue
        if key == "signal_weights":
            if not isinstance(value, dict):
                raise ConfigError("signal_weights override must be a mapping")
            for name, weight in value.items():
                if name not in DEFAULT_SIGNAL_WEIGHTS:
                    raise ConfigError(f"unknown signal name: {name!r}")
                config.signal_weights[name] = _check_float(
                    f"signal_weights.{name}", weight, 0.0, 1.0
                )
        elif key == "min_anomaly_score":
            config.min_anomaly_score = _check_float(key, value, 0.0, 1.0)
        elif key in _INT_KEYS:
            setattr(
                config,
                key,
                _check_int(key, value, 0 if key == "max_localizer_tool_calls" else 1),
            )
    return config


def load_config_from_env(env_var: str = "SHEETGUARD_CONFIG") -> SheetGuardConfig | None:
    """从环境变量读取 YAML 配置路径；未设置或为空返回 None（走默认值）。

    评测 runner 借此获得与 CLI 同源的配置能力：跑实验前
    ``set SHEETGUARD_CONFIG=path/to/sheetguard.yaml`` 即可让整轮
    实验使用同一份行为配置（指纹随审计 JSON 入库）。
    """
    value = (os.getenv(env_var) or "").strip()
    if not value:
        return None
    return load_config(value)
