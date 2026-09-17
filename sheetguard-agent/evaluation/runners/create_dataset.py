"""将合成 Excel 修复案例上传为 Langfuse 数据集。"""

# 标准库：读取 JSON 案例文件，以及处理跨平台文件路径。
import json
from pathlib import Path

# Langfuse 客户端：创建数据集并逐条写入样本。
from langfuse import get_client

# 从项目根目录的 .env 文件加载 Langfuse 配置。
from dotenv import load_dotenv

# 允许脚本从项目根目录直接运行时读取 .env 中的环境变量。
load_dotenv()

# 当前模块位于 evaluation/runners/，parents[2] 即项目根目录。
ROOT = Path(__file__).resolve().parents[2]

# 合成案例清单；每条案例包含损坏工作簿、错误类型和标准答案路径。
CASES_PATH = ROOT / "evaluation/data/cases.json"

# Langfuse 中创建的数据集名称。
DATASET_NAME = "sheetguard-repair-v1"

def main():
    """读取本地案例并创建 Langfuse 数据集。"""

    # 客户端从环境变量中读取 Langfuse 地址和 API Key。
    langfuse = get_client()

    # 创建一个新的数据集，用于保存 Excel 公式修复评测案例。
    langfuse.create_dataset(
        name=DATASET_NAME,
        description=(
            "Excel formula repair benchmark. "
            "Input is a broken workbook; expected output is the gold formula."
        ),
    )

    # cases.json 是 JSON 数组，读取后每个元素对应一个评测案例。
    cases = json.loads(CASES_PATH.read_text(encoding="utf-8"))

    for case in cases:
        # 将相对路径转换为绝对路径，避免实验运行时找不到本地工作簿。
        broken_path = Path(case["broken_path"])
        gold_path = Path(case["gold_path"])
        case_input = {
            "case_id": case["case_id"],
            "broken_path": str(
                (ROOT / broken_path).resolve()
                if not broken_path.is_absolute()
                else broken_path.resolve()
            ),
            "error_type": case["error_type"],
        }
        expected_output = {
            "case_id": case["case_id"],
            "target_cell": case["target_cell"],
            "gold_formula": case["gold_formula"],
            "gold_path": str(
                (ROOT / gold_path).resolve()
                if not gold_path.is_absolute()
                else gold_path.resolve()
            ),
        }
        langfuse.create_dataset_item(
            dataset_name=DATASET_NAME,
            input=case_input,
            expected_output=expected_output,
        )

    # 短生命周期脚本退出前刷新，确保数据已发送。
    langfuse.flush()

    # 输出上传结果，便于确认数据集名称和样本数量。
    print(f"Create dataset: {DATASET_NAME}")
    print(f"Examples: {len(cases)}")


if __name__ == "__main__":
    # 只有直接执行本文件时才创建数据集；被其他模块导入时不会自动上传。
    main()
