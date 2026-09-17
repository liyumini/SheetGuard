"""将用户认证数据集上传为 Langfuse 数据集（sheetguard-user-feedback-v1）。

上传资格 = workbook_index.json 中 status == fully_certified 且工作簿
副本存在（评测器是整表语义，期望集合必须是该表全部 agent 修复格）。
item 结构与 create_multi_cell_dataset.py 同构（input={case_id,
broken_path 绝对路径}, expected_output={errors, expected_status}），
因此 run_multi_cell_evaluation.py 的 --dataset 参数可直接消费。
failed_cases.json 不上传（无 gold，本地难例分析用）。

手动运行；重复执行幂等（按 input.case_id 查重跳过已上传工作簿）。
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

from dotenv import load_dotenv
from langfuse import get_client

from sheetguard.app import review_store as rs

load_dotenv()

DATASET_NAME = "sheetguard-user-feedback-v1"
_GET_DATASET_ATTEMPTS = 3
_GET_DATASET_DELAY = 1.0

CERTIFIED_NAME = rs.CERTIFIED_NAME
FALSE_POSITIVE_NAME = rs.FALSE_POSITIVE_NAME
WORKBOOKS_DIRNAME = rs.WORKBOOKS_DIRNAME


class LangfuseUnconfiguredError(RuntimeError):
    """Langfuse 客户端未初始化（缺少凭据或 SDK 以 disabled 状态运行）。"""


class DatasetFetchError(RuntimeError):
    """数据集拉取在退避重试后仍失败（网络等瞬时错误耗尽）。"""


def _is_dataset_missing(exc: Exception) -> bool:
    """判定 get_dataset 的异常是否为「数据集不存在」语义。

    SDK 4.15.1 实测：不存在的数据集经 handle_fern_exception 原样重抛
    langfuse.api.NotFoundError（ApiError 子类，status_code=404，非
    LookupError）。为同时兼容测试桩（LookupError）与 SDK 演进，用
    LookupError / status_code==404 / 消息含 "not found" 三重判定。
    """
    if isinstance(exc, LookupError):
        return True
    if getattr(exc, "status_code", None) == 404:
        return True
    return "not found" in str(exc).lower()


def _get_dataset_with_retry(
    langfuse,
    name: str,
    attempts: int = _GET_DATASET_ATTEMPTS,
    delay: float = _GET_DATASET_DELAY,
):
    """拉取数据集，瞬时错误按固定退避重试（设计 7）。

    返回值三态：
    - DatasetClient（或任意带 .items 的对象）：数据集存在，用于查重；
    - None：数据集不存在（404 语义），调用方稍后创建；
    - 抛 LangfuseUnconfiguredError：客户端未初始化（未配置凭据）；
    - 抛 DatasetFetchError：重试耗尽——调用方应中止上传（此时拿不到
      查重集合，继续跑会重复上传 item）。
    """
    last: Exception | None = None
    for attempt in range(1, attempts + 1):
        try:
            return langfuse.get_dataset(name)
        except Exception as exc:
            if _is_dataset_missing(exc):
                return None
            if isinstance(exc, AttributeError) and "not initialized" in str(exc):
                # SDK 4.15.1：未配置凭据时 get_client() 不抛错而是返回
                # disabled 客户端，首次 API 调用才抛
                # AttributeError("Langfuse client is not initialized")。
                raise LangfuseUnconfiguredError(
                    f"{exc}（请检查 .env 的 LANGFUSE_PUBLIC_KEY/SECRET_KEY）"
                ) from exc
            last = exc
            if attempt < attempts:
                time.sleep(delay)
    raise DatasetFetchError(f"{attempts} 次尝试后仍失败：{last}") from last


def _exit_unconfigured(exc: Exception) -> None:
    print(
        f"错误：Langfuse 未配置或不可用（{exc}）。"
        "请检查 .env 的 LANGFUSE_PUBLIC_KEY/SECRET_KEY。",
        file=sys.stderr,
    )


def build_workbook_item(
    stem: str, samples: list[dict], copy_path: Path,
    fp_samples: list[dict] | None = None,
) -> dict:
    """把一张全认证工作簿的格子级样本聚合为评测 item。

    errors 数组与 multi_cases.json 同构（评测器消费 target_cell +
    gold_formula）；error_type 缺失以 "unknown" 占位（仅记录，不参与
    类型匹配）。metadata 携带每格 certified_at/agent_version 与 config
    指纹，供 Langfuse 端按时间/版本切片对比。

    fp_samples 为该表被判误修（false_positive）的样本：混合进同一
    item——expected_output.expected_dismissed 列出其 target_cell
    （评测器不再要求这些格被修复），metadata.false_positives 记录
    判定溯源（case_id/decided_at/agent_version/round_decided）。
    """
    first = samples[0] if samples else {}
    fps = fp_samples or []
    return {
        "input": {
            "case_id": f"userfb_{stem}",
            "broken_path": str(copy_path.resolve()),
        },
        "expected_output": {
            "errors": [
                {
                    "target_cell": s.get("target_cell"),
                    "error_type": s.get("error_type") or "unknown",
                    "old_formula": s.get("old_formula"),
                    "gold_formula": s.get("gold_formula"),
                }
                for s in samples
            ],
            # 保序去重：同格跨轮多条 fp 样本只列一次。
            "expected_dismissed": list(dict.fromkeys(
                s.get("target_cell") for s in fps if s.get("target_cell")
            )),
            "expected_status": "success",
        },
        "metadata": {
            "source": "user_review",
            "samples": [
                {
                    "case_id": s.get("case_id"),
                    "certified_at": s.get("certified_at"),
                    "agent_version": s.get("agent_version"),
                    "round_certified": s.get("round_certified"),
                }
                for s in samples
            ],
            "false_positives": [
                {
                    "case_id": s.get("case_id"),
                    "decided_at": s.get("decided_at"),
                    "agent_version": s.get("agent_version"),
                    "round_decided": s.get("round_decided"),
                }
                for s in fps
            ],
            "config": first.get("config"),
            "workbook_file": first.get("workbook"),
        },
    }


def main(client=None, root: Path | None = None) -> None:
    """读取本地 user_feedback 数据并上传；查重幂等。root 供测试注入。

    root 语义与 review_store 一致：None → 包内默认数据集目录
    （evaluation/data/user_feedback/）；显式路径 → 直接作为数据集根
    （load_workbook_index 与 data_root 均直用，不叠加子路径）。
    """
    try:
        langfuse = client if client is not None else get_client()
    except Exception as exc:
        # 未配置/不可用 → 友好报错退出（设计 7），不裸 traceback。
        _exit_unconfigured(exc)
        return
    index = rs.load_workbook_index(root)
    # root 直用语义（Task 2/3 确立）：显式 root 即数据集根；
    # root=None → 包内默认数据集目录（dataset_dir 的默认分支）。
    data_root = rs.dataset_dir() if root is None else Path(root)
    samples_by_workbook: dict[str, list[dict]] = {}
    cert_path = data_root / rs.CERTIFIED_NAME
    for sample in rs.load_samples(cert_path):
        samples_by_workbook.setdefault(
            sample.get("workbook") or "?", []
        ).append(sample)
    # 误修样本按表聚合，混合进对应工作簿 item（Task 4）。
    fp_by_workbook: dict[str, list[dict]] = {}
    for sample in rs.load_samples(data_root / rs.FALSE_POSITIVE_NAME):
        fp_by_workbook.setdefault(sample.get("workbook") or "?", []).append(sample)

    # 已上传工作簿查重：Langfuse dataset item 无业务幂等键，重跑靠
    # input.case_id 集合过滤。网络瞬时错误 → 退避重试（设计 7）；
    # 数据集不存在 → None（稍后创建）；重试耗尽 → 中止（查重集合拿
    # 不到就不能安全上传，重复 item 比晚传更糟）。
    try:
        dataset = _get_dataset_with_retry(langfuse, DATASET_NAME)
    except LangfuseUnconfiguredError as exc:
        _exit_unconfigured(exc)
        return
    except DatasetFetchError as exc:
        print(
            f"错误：拉取数据集 {DATASET_NAME} 失败（{exc}），"
            "已中止上传以避免重复 item。",
            file=sys.stderr,
        )
        return
    existing_case_ids: set[str] = set()
    if dataset is not None:
        existing_case_ids = {
            item.input.get("case_id")
            for item in dataset.items
            if isinstance(item.input, dict)
        }
    uploaded = 0
    for stem, info in sorted(index.items()):
        if info.get("status") != "fully_certified":
            continue
        case_id = f"userfb_{stem}"
        if case_id in existing_case_ids:
            print(f"已上传跳过：{case_id}")
            continue
        copy_path = data_root / WORKBOOKS_DIRNAME / f"{stem}.xlsx"
        if not copy_path.is_file():
            print(f"跳过（副本缺失）：{stem}")
            continue
        workbook_samples = samples_by_workbook.get(stem) or []
        fp_samples_for_wb = fp_by_workbook.get(stem) or []
        # v1.11 门扩展后误修格即终局：一张全部判 3（无 certified 样本）的
        # 表也可 fully_certified——只有两个来源都为空才跳过，否则 fp 数据
        # 会静默不上传。
        if not workbook_samples and not fp_samples_for_wb:
            print(
                f"跳过（certified 与 false_positive_cases.json 均无该表样本）：{stem}"
            )
            continue
        if dataset is None:
            langfuse.create_dataset(
                name=DATASET_NAME,
                description=(
                    "SheetGuard user-certified repair benchmark. "
                    "Broken workbooks come from real review sessions; "
                    "gold formulas are user-certified proposals."
                ),
            )
            dataset = True  # 只创建一次
        item = build_workbook_item(
            stem, workbook_samples, copy_path, fp_samples_for_wb,
        )
        langfuse.create_dataset_item(
            dataset_name=DATASET_NAME,
            input=item["input"],
            expected_output=item["expected_output"],
            metadata=item["metadata"],
        )
        uploaded += 1
    langfuse.flush()
    print(f"dataset={DATASET_NAME} uploaded={uploaded}")


if __name__ == "__main__":
    main()
