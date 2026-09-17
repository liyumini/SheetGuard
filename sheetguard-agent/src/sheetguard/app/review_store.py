"""用户审查反馈的落盘与真实用户数据集生成。

反馈衍生数据（全部本地 JSON）：

- feedback.json（每轮一份）——用户原始审查记录（每条含 kind 来源类别），
  审计追溯用；
- certified_cases.json——金标准：错误公式 → 用户认可公式（含轨迹）；
- failed_cases.json——失败难题：rejected_exhausted 的格子（无 gold）；
- false_positive_cases.json——误修样本：原公式即金标准（v1.11）；
- missed_cases.json——漏检样本：被推翻的 dismissed 争议格认证时沉淀
  （v1.12，recall 分析档案，不上传）。

认证与失败样本分开存储：用途与置信度语义不同，failed 日后人工补标
正确公式后可直接挪入 certified（两文件 schema 对齐，纯数据搬运）。
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

FEEDBACK_NAME = "feedback.json"
DRAFT_NAME = "feedback_draft.json"
CERTIFIED_NAME = "certified_cases.json"
FAILED_NAME = "failed_cases.json"
FALSE_POSITIVE_NAME = "false_positive_cases.json"
MISSED_NAME = "missed_cases.json"
WORKBOOKS_DIRNAME = "workbooks"
INDEX_NAME = "workbook_index.json"

VERDICTS = frozenset({"correct", "incorrect", "false_positive", "skipped"})
# 审查条目类型（v1.12）：feedback 每条 review 记录来源类别，供历史重现
# 过滤（dismissed 的 skipped=默认通过，不重现）与漏检来源识别使用。
KINDS = frozenset({"fixed", "failed", "skipped", "dismissed"})
# 数据集默认根：依赖 src 布局（src/sheetguard/app/review_store.py →
# parents[3] 即 sheetguard-agent/）。非 editable 安装到 site-packages 时
# 该路径不成立，dataset_dir() 需另传 root 参数。
_PACKAGE_ROOT = Path(__file__).resolve().parents[3]
_DATASET_DIRNAME = Path("evaluation") / "data" / "user_feedback"


class FeedbackError(ValueError):
    """反馈文件缺失、损坏或不符合 schema。"""


def _read_json(path: Path) -> object:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError as exc:
        raise FeedbackError(f"file not found: {path}") from exc
    except json.JSONDecodeError as exc:
        raise FeedbackError(f"invalid JSON in {path}: {exc}") from exc
    except UnicodeDecodeError as exc:
        raise FeedbackError(f"invalid encoding in {path}: {exc}") from exc


def _validate_feedback(payload: object) -> dict:
    if not isinstance(payload, dict):
        raise FeedbackError("feedback must be a JSON object")
    reviews = payload.get("reviews")
    if not isinstance(reviews, list):
        raise FeedbackError("feedback.reviews must be a list")
    for item in reviews:
        if not isinstance(item, dict):
            raise FeedbackError("feedback.reviews items must be objects")
        if not isinstance(item.get("target"), str) or not item["target"]:
            raise FeedbackError("review item needs a non-empty 'target'")
        verdict = item.get("verdict")
        # 先做 isinstance(str) 判断：不可哈希的 verdict（如 list）不能进集合成员运算。
        if not isinstance(verdict, str) or verdict not in VERDICTS:
            raise FeedbackError(
                f"verdict must be one of {sorted(VERDICTS)}, got: {verdict!r}"
            )
        proposal = item.get("proposal")
        # proposal 可为 null：failed 条目可能没有有效提案（graph 早停或
        # 全部尝试非法 → last_formula=None）。但 correct 判定必须有非空
        # 提案——认证样本的 gold_formula 来自它，缺失会出现空金标准。
        if proposal is not None and not isinstance(proposal, str):
            raise FeedbackError("review item 'proposal' must be a string or null")
        if verdict == "correct" and (not isinstance(proposal, str) or not proposal):
            raise FeedbackError(
                "correct verdict requires a non-empty 'proposal'"
            )
        kind = item.get("kind", "fixed")
        if not isinstance(kind, str) or kind not in KINDS:
            raise FeedbackError(
                f"kind must be one of {sorted(KINDS)}, got: {item.get('kind')!r}"
            )
    return payload


def save_feedback(path: Path, payload: dict) -> None:
    """校验并写入一份已提交的反馈（提交即触发重修，调用方负责后续动作）。"""
    _validate_feedback(payload)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


def load_feedback(path: Path) -> dict:
    return _validate_feedback(_read_json(path))


def save_draft(path: Path, payload: dict) -> None:
    """草稿不做 schema 校验：q 退出时保存用户已按键的任意部分状态。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


def load_draft(path: Path) -> dict | None:
    if not path.exists():
        return None
    data = _read_json(path)
    return data if isinstance(data, dict) else None


def round_dirs(base: Path) -> list[Path]:
    """base 下按轮号升序的 round-N 目录列表。"""
    if not base.is_dir():
        return []
    dirs = [d for d in base.glob("round-*") if d.name.split("-")[1].isdigit()]
    return sorted(dirs, key=lambda d: int(d.name.split("-")[1]))


def latest_reviewable_round(base: Path) -> Path | None:
    """最新一个有 audit.json 且尚无 feedback.json 的轮次目录。"""
    for d in reversed(round_dirs(base)):
        if (d / "audit.json").is_file() and not (d / FEEDBACK_NAME).exists():
            return d
    return None


def agent_version() -> str:
    """运行代理版本指纹：环境变量 → git 短 hash → 包版本 → unknown。

    数据集样本据此区分"哪个版本的 agent 产生的提案"，跨实验对比时
    数据集质量可以按 agent 版本切片。
    """
    env = (os.getenv("SHEETGUARD_AGENT_VERSION") or "").strip()
    if env:
        return env
    try:
        result = subprocess.run(
            ["git", "rev-parse", "--short", "HEAD"],
            capture_output=True,
            text=True,
            timeout=5,
            cwd=str(_PACKAGE_ROOT),
            check=False,
        )
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout.strip()
    except (OSError, subprocess.SubprocessError):
        pass
    try:
        from importlib.metadata import version as _version
        return _version("sheetguard")
    except Exception:
        return "unknown"


def ensure_workbook_copy(
    broken_source: Path | None, stem: str, root: Path | None = None
) -> Path | None:
    """把 broken_source 副本放进数据集 workbooks/ 目录（按工作簿一份）。

    副本让数据集自包含（out/ 随时可清，评测路径永远有效）。已存在不
    覆盖——同表多轮样本共享同一副本，内容一致性由 broken_source 自身
    不被覆盖保证；内容意外不同时保留现有副本并警告。broken_source
    缺失返回 None，调用方以 broken_path=null 入库，上传脚本会跳过。
    """
    if broken_source is None or not Path(broken_source).is_file():
        return None
    # root 即数据集根（显式 --dataset-dir 时与 certified_cases.json 同级）；
    # 缺省时落到包内默认数据集目录。
    base = dataset_dir() if root is None else Path(root)
    target = base / WORKBOOKS_DIRNAME / f"{Path(stem).stem}.xlsx"
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        if target.read_bytes() != Path(broken_source).read_bytes():
            print(
                f"warning: workbook copy differs from broken_source, "
                f"keeping existing: {target}",
                file=sys.stderr,
            )
        return target
    shutil.copy2(broken_source, target)
    return target


def load_samples(path: Path) -> list[dict]:
    """读取数据集文件为 dict 列表（缺失/非 dict 条目容错跳过）。"""
    if not path.exists():
        return []
    data = _read_json(path)
    if not isinstance(data, list):
        return []
    return [item for item in data if isinstance(item, dict)]


def load_workbook_index(root: Path | None = None) -> dict:
    """读 workbook_index.json；缺失或损坏返回 {}（不抛错，保守降级）。"""
    base = dataset_dir() if root is None else Path(root)
    path = base / INDEX_NAME
    if not path.exists():
        return {}
    try:
        data = _read_json(path)
    except FeedbackError:
        return {}
    return data if isinstance(data, dict) else {}


def save_workbook_index(root: Path | None, index: dict) -> None:
    base = dataset_dir() if root is None else Path(root)
    path = base / INDEX_NAME
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(index, indent=2, ensure_ascii=False), encoding="utf-8")


def decide_workbook_status(
    rounds: list[dict], resolved_targets: set[str], rounds_readable: bool
) -> tuple[str, list[str], list[str]]:
    """全认证门（v1.11 扩展）：返回 (status, pending_targets, contested_targets)。

    rounds 每项 {"round": int, "fixed": [targets], "rejected": [targets],
    "failed": [targets]}。
    - unknown：轮次数据不全（out/ 被清理）无法完整重算，保守不上传；
    - fully_certified：所有 agent 处理过的格子（修过的 + 失败的）都有终局
      ——终局 = certified 或 false_positive（resolved_targets），且无争议；
    - contested：被判「错」且其后没有轮次再修复它（agent dismiss）——
      该格 gold 处于争议状态，新会话重新修复并认证后可解封。
    注：未判/跳过的格子天然落进 pending（处理过且未 resolved）。
    """
    if not rounds_readable:
        return ("unknown", [], [])
    processed_ever: set[str] = set()
    rejected_at: dict[str, int] = {}
    for entry in rounds:
        no = int(entry.get("round") or 0)
        for target in entry.get("fixed") or []:
            processed_ever.add(target)
        for target in entry.get("failed") or []:
            processed_ever.add(target)
        for target in entry.get("rejected") or []:
            # 直接赋值取最后一次否决轮：多次否决语义对应最后一次。
            rejected_at[target] = no
    resolved = set(resolved_targets)
    pending = sorted(processed_ever - resolved)
    contested = sorted(
        target
        for target, rejected_no in rejected_at.items()
        if target not in resolved
        and not any(
            target in (entry.get("fixed") or [])
            for entry in rounds
            if int(entry.get("round") or 0) > rejected_no
        )
    )
    if pending or contested:
        # contested 非空（如 rejected 从未 fixed）也保守判 partial：
        # fully_certified 必须是"全部认证且无争议"。
        return ("partial", pending, contested)
    return ("fully_certified", [], [])


def record_workbook_entry(
    root: Path | None,
    stem: str,
    *,
    status: str,
    pending_targets: list[str],
    contested_targets: list[str],
    sample_case_ids: list[str],
    false_positive_case_ids: list[str] | None = None,
) -> None:
    """在 workbook_index.json 中更新一个工作簿的认证状态条目。"""
    index = load_workbook_index(root)
    index[Path(stem).stem] = {
        "status": status,
        "sample_case_ids": sorted(sample_case_ids),
        "pending_targets": sorted(pending_targets),
        "contested_targets": sorted(contested_targets),
        "false_positive_case_ids": sorted(false_positive_case_ids or []),
        "updated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
    }
    save_workbook_index(root, index)


def build_trajectory(feedbacks: list[dict], target: str) -> list[dict]:
    """按轮次顺序收集某格的（提案, 判定, 备注）历史。"""
    trajectory: list[dict] = []
    for feedback in feedbacks:
        for item in feedback.get("reviews") or []:
            if item.get("target") == target:
                trajectory.append({
                    "round": feedback.get("round"),
                    "proposal": item.get("proposal"),
                    "verdict": item.get("verdict"),
                    "remark": item.get("remark"),
                })
    return trajectory


def _stem_with_hash(workbook_name: str) -> str:
    """清洗工作簿名 stem；清洗改写或原名含大写时附加原名 sha1 前 8 位防碰撞。

    旧正则把连续非 ASCII 段折叠成单个 `_`，不同中文工作簿名会撞出相同
    case_id（append_samples 去重会静默丢样本）。改用 UNICODE \\W 后 CJK
    保留原名；对仍被清洗改写的名字补 hash 后缀，保证一一对应。另外最终
    会 lower() 折叠大小写，"Budget.xlsx" 与 "budget.xlsx" 清洗后同为
    "budget"，二者之一（含大写的原名）也必须补 hash 才不碰撞。纯小写
    且清洗无损的名字（如 "budget"）不加 hash，保持既有 case_id 前缀稳定。
    """
    raw = Path(workbook_name).stem
    cleaned = re.sub(r"[\W]+", "_", raw)
    if cleaned != raw or raw != raw.lower():
        digest = hashlib.sha1(raw.encode("utf-8")).hexdigest()[:8]
        cleaned = f"{cleaned}_{digest}"
    return cleaned.lower()


def build_certified_sample(
    workbook_name: str,
    entry: dict,
    review: dict,
    round_no: int,
    broken_source: str,
    trajectory: list[dict],
    certified_at: str | None = None,
    agent_version: str | None = None,
    config: dict | None = None,
) -> dict:
    """从本轮 audit 的 fixed 条目 + 用户判定构造金标准样本。"""
    hypothesis = entry.get("hypothesis") or {}
    stem = _stem_with_hash(workbook_name)
    return {
        "case_id": f"user_{stem}_r{round_no}_{review['target'].replace('!', '_').lower()}",
        "workbook": workbook_name,
        "target_cell": review["target"],
        # old_formula 取 audit 条目里的原公式（本轮运行时格子上的错误公式）。
        "old_formula": entry.get("old_formula"),
        "gold_formula": review["proposal"],
        "broken_path": broken_source,
        "error_type": hypothesis.get("error_type"),
        "source": "user_review",
        "round_certified": round_no,
        "certified_at": certified_at,
        "agent_version": agent_version,
        "config": config,
        "trajectory": trajectory,
    }


def build_failed_sample(
    workbook_name: str,
    entry: dict,
    trajectory: list[dict],
    max_rounds: int,
    broken_source: str | None = None,
    recorded_at: str | None = None,
    agent_version: str | None = None,
    config: dict | None = None,
) -> dict:
    """rejected_exhausted 样本：无 gold，只记轨迹（agent 攻不动的难题）。"""
    stem = _stem_with_hash(workbook_name)
    target = entry.get("target") or entry.get("target_cell") or "?"
    return {
        "case_id": f"user_{stem}_failed_{target.replace('!', '_').lower()}",
        "workbook": workbook_name,
        "target_cell": target,
        "old_formula": entry.get("old_formula"),
        "gold_formula": None,
        "broken_path": broken_source,
        "error_type": (entry.get("hypothesis") or {}).get("error_type"),
        "source": "user_review",
        "outcome": "rejected_exhausted",
        "max_rounds": max_rounds,
        "recorded_at": recorded_at,
        "agent_version": agent_version,
        "config": config,
        "trajectory": trajectory,
    }


def find_dismissal_dispute(feedbacks: list[dict], target: str) -> int | None:
    """返回 target 最早的 dismissed 争议轮号（kind=dismissed 且 verdict=incorrect）。

    无争议史返回 None。feedbacks 按轮升序传入；无 kind 字段的旧 review 按
    fixed 处理，不构成 dismissed 争议。
    """
    for feedback in feedbacks:
        for item in feedback.get("reviews") or []:
            if (
                item.get("target") == target
                and item.get("kind", "fixed") == "dismissed"
                and item.get("verdict") == "incorrect"
            ):
                return feedback.get("round")
    return None


def build_missed_sample(
    workbook_name: str,
    entry: dict,
    review: dict,
    round_no: int,
    broken_source: str,
    trajectory: list[dict],
    *,
    dismissed_round: int | None,
    dismissed_reason: str | None,
    certified_at: str | None = None,
    agent_version: str | None = None,
    config: dict | None = None,
) -> dict:
    """漏检样本：agent 曾 dismiss、用户推翻后最终认证的格子（recall 档案）。

    与 certified 同构（old_formula=认证轮 fixed 条目的原公式——dismissed
    后从未被修改过的原始错误公式；gold=用户认可公式），追加争议来源字段；
    不上传 Langfuse（certified 集已天然覆盖该格参与整表评测）。
    """
    stem = _stem_with_hash(workbook_name)
    hypothesis = entry.get("hypothesis") or {}
    return {
        "case_id": f"usermiss_{stem}_r{round_no}_{review['target'].replace('!', '_').lower()}",
        "workbook": workbook_name,
        "target_cell": review["target"],
        "old_formula": entry.get("old_formula"),
        "gold_formula": review["proposal"],
        "broken_path": broken_source,
        "error_type": hypothesis.get("error_type"),
        "source": "user_review",
        "outcome": "missed_by_agent",
        "round_certified": round_no,
        "dismissed_round": dismissed_round,
        "dismissed_reason": dismissed_reason,
        "certified_at": certified_at,
        "agent_version": agent_version,
        "config": config,
        "trajectory": trajectory,
    }


def build_false_positive_sample(
    workbook_name: str,
    *,
    target: str,
    old_formula: str | None,
    error_type: str | None,
    round_no: int,
    broken_source: str | None,
    trajectory: list[dict],
    decided_at: str | None = None,
    agent_version: str | None = None,
    config: dict | None = None,
) -> dict:
    """误修样本：用户裁决"这格本来就没坏，agent 不该动它"。

    gold_formula = old_formula——原公式即金标准（"这格的正确公式就是
    它现在的公式"），schema 与 certified 对齐；outcome 标 false_positive
    与 certified/failed 区分。来源无关（failed 条目或历史跳过格），由
    调用方拼装显式参数。
    """
    stem = _stem_with_hash(workbook_name)
    return {
        "case_id": f"userfp_{stem}_r{round_no}_{target.replace('!', '_').lower()}",
        "workbook": workbook_name,
        "target_cell": target,
        "old_formula": old_formula,
        "gold_formula": old_formula,
        "broken_path": broken_source,
        "error_type": error_type,
        "source": "user_review",
        "outcome": "false_positive",
        "round_decided": round_no,
        "decided_at": decided_at,
        "agent_version": agent_version,
        "config": config,
        "trajectory": trajectory,
    }


def build_promoted_sample(
    sample: dict,
    gold_formula: str,
    promoted_at: str | None = None,
) -> dict:
    """把 rejected_exhausted 样本人工晋升为 certified（source=manual）。

    case_id 保留 stem/target、`_failed_` 段换为 `_manual_`（前缀语义跟随
    当前归属文件，且不会与未来自动认证撞名——该格晋升后已是终局）；
    轨迹完整保留。
    """
    stem = _stem_with_hash(sample.get("workbook") or "")
    target = sample.get("target_cell") or "?"
    return {
        "case_id": f"user_{stem}_manual_{target.replace('!', '_').lower()}",
        "workbook": sample.get("workbook"),
        "target_cell": target,
        "old_formula": sample.get("old_formula"),
        "gold_formula": gold_formula,
        "broken_path": sample.get("broken_path"),
        "error_type": sample.get("error_type"),
        "source": "manual",
        "promoted_at": promoted_at,
        "trajectory": sample.get("trajectory") or [],
    }


def remove_sample(path: Path, case_id: str) -> bool:
    """从数据集文件移除一条样本（重写文件）；返回是否发生了移除。"""
    samples = load_samples(path)
    kept = [s for s in samples if s.get("case_id") != case_id]
    if len(kept) == len(samples):
        return False
    # 回写会顺带丢弃 load_samples 过滤掉的非 dict 损坏条目——与
    # append_samples 对称打警告，防静默丢样本。
    corrupt = []
    if path.exists():
        data = _read_json(path)
        if isinstance(data, list):
            corrupt = [item for item in data if not isinstance(item, dict)]
    if corrupt:
        n = len(corrupt)
        unit = "entry" if n == 1 else "entries"
        print(
            f"warning: skipped {n} corrupt {unit} in {path}",
            file=sys.stderr,
        )
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(kept, indent=2, ensure_ascii=False), encoding="utf-8")
    return True


def append_samples(path: Path, samples: list[dict]) -> None:
    """按 case_id 去重合并写入数据集文件（存在即跳过，不覆盖旧样本）。"""
    existing: list[dict] = []
    if path.exists():
        data = _read_json(path)
        if not isinstance(data, list):
            raise FeedbackError(f"dataset file must be a JSON array: {path}")
        # 跳过损坏条目（非 dict）：不参与去重，也不写回文件；打警告防静默丢样本。
        corrupt = [item for item in data if not isinstance(item, dict)]
        if corrupt:
            n = len(corrupt)
            unit = "entry" if n == 1 else "entries"
            print(
                f"warning: skipped {n} corrupt {unit} in {path}",
                file=sys.stderr,
            )
        existing = [item for item in data if isinstance(item, dict)]
    seen = {item.get("case_id") for item in existing}
    existing.extend(s for s in samples if s.get("case_id") not in seen)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(existing, indent=2, ensure_ascii=False), encoding="utf-8")


def dataset_dir(root: Path | None = None) -> Path:
    """用户反馈数据集目录；测试传 root 参数重定向。"""
    return (root if root is not None else _PACKAGE_ROOT) / _DATASET_DIRNAME
