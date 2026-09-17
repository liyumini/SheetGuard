"""上传 multi-cell 难题数据集到 Langfuse（sheetguard-repair-multi-cell-hard-v1）。

expected_output 在 errors / expected_status 之外可选携带
expected_dismissed / expected_skipped / expected_deferred 三个
workbook 级期望键，配套评测器据此判分（见 multi_cell_evaluators.py）。

数据集不存在时创建并逐条上传；已存在且非空时直接中止，防止重复
运行把同一批 item 再传一遍（multi-cell 常规集曾因重建出现新旧 item
并存的问题）。
"""
import json
from pathlib import Path

from dotenv import load_dotenv
from langfuse import get_client

load_dotenv()

ROOT = Path(__file__).resolve().parents[2]
NAME = "sheetguard-repair-multi-cell-hard-v1"
CASES_PATH = ROOT / "evaluation/data/hard_cases.json"


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
                "SheetGuard multi-cell hard cases: false-repair trap / "
                "dependency cycle / budget deferral / clean workbook / "
                "five-fault combo / double chain break"
            ),
        )

    for case in cases:
        broken = Path(case["broken_path"])
        case_input = {
            "case_id": case["case_id"],
            "broken_path": str(broken.resolve()),
        }
        expected_output = {
            "errors": case["errors"],                     # target_cell + gold_formula
            "expected_status": case["expected_status"],   # workbook-level gold
        }
        for key in ("expected_dismissed", "expected_skipped", "expected_deferred"):
            if case.get(key):
                expected_output[key] = case[key]
        langfuse.create_dataset_item(
            dataset_name=NAME,
            input=case_input,
            expected_output=expected_output,
        )

    langfuse.flush()
    print(f"dataset={NAME} items={len(cases)}")


if __name__ == "__main__":
    main()
