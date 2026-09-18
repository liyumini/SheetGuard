"""上传 coverage 数据集到 Langfuse（sheetguard-repair-coverage-v1）。

覆盖矩阵数据集：9 个重算引擎函数/四则表达式形态 × 5 类错误类型中所有
可注入组合（33 个单对案例 + 1 个五故障混合）。item 结构与 complete 集
一致（input=case_id/broken_path，expected_output=errors/expected_status）；
数据集已存在且非空时直接中止，防止重复上传（与 complete/hard 集守卫一致）。
"""
import json
from pathlib import Path

from dotenv import load_dotenv
from langfuse import get_client

load_dotenv()

ROOT = Path(__file__).resolve().parents[2]
NAME = "sheetguard-repair-coverage-v1"
CASES_PATH = ROOT / "evaluation/data/coverage_cases.json"


def build_item(case: dict) -> tuple[dict, dict]:
    """组装 (item.input, item.expected_output)。"""
    return (
        {
            "case_id": case["case_id"],
            "broken_path": str(Path(case["broken_path"]).resolve()),
        },
        {
            "errors": case["errors"],
            "expected_status": case["expected_status"],
        },
    )


def main():
    langfuse = get_client()

    existing = None
    try:
        existing = langfuse.get_dataset(NAME)
    except Exception:
        existing = None
    if existing is not None:
        items = list(existing.items or [])
        if items:
            raise SystemExit(
                f"数据集 {NAME} 已存在且含 {len(items)} 个 item，"
                "为避免重复上传请先在 Langfuse UI 清空或改名。"
            )

    cases = json.loads(CASES_PATH.read_text(encoding="utf-8"))
    if not existing:
        langfuse.create_dataset(
            name=NAME,
            description=(
                "SheetGuard full coverage matrix: 9 recalc-engine functions / "
                "expression forms x 5 error types (injectable pairs) + "
                "five-fault mix. Errors carry carrier_function/expression_form."
            ),
        )

    for case in cases:
        item_input, expected_output = build_item(case)
        langfuse.create_dataset_item(
            dataset_name=NAME,
            input=item_input,
            expected_output=expected_output,
        )

    langfuse.flush()
    print(f"dataset={NAME} items={len(cases)}")


if __name__ == "__main__":
    main()
