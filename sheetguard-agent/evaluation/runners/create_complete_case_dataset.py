"""上传 complete 数据集到 Langfuse（sheetguard-repair-complete-v1）。

expected_output = errors / expected_status；complete 集不做 workbook 级
dismissed/skipped/deferred 期望（与 hard 集的可选键不同，本集全部
expected_status=success）。数据集不存在时创建并逐条上传；已存在且非空时
直接中止，防止重复上传（与 hard 集守卫一致）。
"""
import json
from pathlib import Path

from dotenv import load_dotenv
from langfuse import get_client

load_dotenv()

ROOT = Path(__file__).resolve().parents[2]
NAME = "sheetguard-repair-complete-v1"
CASES_PATH = ROOT / "evaluation/data/complete_cases.json"


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
                "SheetGuard complete detection coverage: 5 anomaly signals / "
                "8 formula-subset function carriers / fault combos"
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
