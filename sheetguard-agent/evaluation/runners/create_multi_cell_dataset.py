import json
from pathlib import Path
from dotenv import load_dotenv
from langfuse import get_client

load_dotenv()
ROOT = Path(__file__).resolve().parents[2]
NAME = "sheetguard-repair-multi-cell-v1"

langfuse = get_client()
langfuse.create_dataset(name=NAME, description="SheetGuard multi-cell batch repair benchmark")

cases = json.loads((ROOT / "evaluation/data/multi_cases.json").read_text(encoding="utf-8"))
for case in cases:
    broken = Path(case["broken_path"])
    case_input = {
        "case_id": case["case_id"],
        "broken_path": str((ROOT / broken).resolve() if not broken.is_absolute() else broken.resolve()),
    }
    langfuse.create_dataset_item(
        dataset_name=NAME,
        input=case_input,
        expected_output={
            "errors": case["errors"],                 # target_cell + gold_formula，错误级 gold
            "expected_status": case["expected_status"],  # workbook-level gold
        },
    )
langfuse.flush()
print(f"dataset={NAME} items={len(cases)}")