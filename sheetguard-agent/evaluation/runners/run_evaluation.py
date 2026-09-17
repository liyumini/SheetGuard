"""运行 Langfuse 实验，检查 SheetGuard 的定位和公式修复结果。"""

# 加载 .env 中的模型、API 和 Langfuse 配置。
from dotenv import load_dotenv

# SheetGuard 的工作流和模型构造函数。
from sheetguard.agent.graph import SheetGuardGraph
from sheetguard.agent.llm_config import build_model
from sheetguard.config import load_config_from_env

# Langfuse 客户端、评测结果类型和共享数据集名称。
from langfuse import Evaluation, get_client

# 公式语义等价比较：把公式归一化为相对模板，屏蔽 $ 锚定、表名引号风格
# 和空白差异；再去除包裹单个引用的冗余括号，如 =('P&L'!B2)*0.25 与
# ='P&L'!B2*0.25 视为等价。
from sheetguard.spreadsheet.aggregates import normalize_template
from sheetguard.spreadsheet.formula_parser import to_relative_template
from sheetguard.spreadsheet.model import CellRef

# 进度更新用的锁以及 UTF-8 输出所需的标准库。
import threading
import io
import sys

from rich.console import Console
from rich.progress import (
    BarColumn,
    Progress,
    SpinnerColumn,
    TaskProgressColumn,
    TextColumn,
    TimeElapsedColumn,
)

from evaluation.runners.create_dataset import DATASET_NAME
from evaluation.runners.score_integrity import verify_experiment_records, verify_item_results


# 模块级评估进度条。run_experiment 会在事件循环内并发调用 target()
# （max_concurrency=1 时实际串行执行），所以可以在每次 target 完成后更新进度。
# 用锁保证不同并发下进度更新线程安全。
_console = Console(
    # 进度条含 Unicode 字符，在 Windows 中文（GBK）终端直接写 stdout 会报错，
    # 这里强制用 UTF-8 包装输出流，保证进度条能在任何终端正常渲染。
    file=io.TextIOWrapper(
        sys.stdout.buffer if hasattr(sys.stdout, "buffer") else sys.stdout,
        encoding="utf-8",
        errors="replace",
        line_buffering=True,
    )
)
_progress = Progress(
    SpinnerColumn(),
    TextColumn("[progress.description]{task.description}"),
    BarColumn(),
    TaskProgressColumn(),
    TextColumn("("),
    TimeElapsedColumn(),
    TextColumn(")"),
    console=_console,
)
_progress_lock = threading.Lock()
_progress_task_id: int | None = None

# 从项目根目录的 .env 读取配置变量。
load_dotenv()

# Langfuse task 接收一个数据集 item，调用 Agent，并返回待评测结果。
def target(*, item, **kwargs) -> dict:
    """运行一次 SheetGuard，并返回本案例的修复结果。"""

    inputs = item.input
    # 输入数据集中保存的是待修复工作簿路径。
    broken_path = inputs["broken_path"]

    # 每个评测案例创建一个独立 Agent，避免运行时索引和沙箱相互复用。
    # 行为参数跟 sheetguard.config 的默认值走：评测测的就是"当前项目"，
    # 实验需要临时换配置时用 SHEETGUARD_CONFIG 环境变量指定 YAML。
    agent = SheetGuardGraph(model=build_model(), config=load_config_from_env())

    try:
        # 执行解析、检测、定位、诊断、修复和验证完整流程。
        result = agent.invoke(broken_path)

        # 从 Agent 最终状态中提取模型提出的补丁。
        patch = result.get("proposed_patch") or {}

        # 返回 LangSmith 评测函数需要的预测结果。
        return {
            "case_id": inputs["case_id"],
            "target": patch.get("target", ""),
            "new_formula": patch.get("new_formula", ""),
            "status": result.get("status", ""),
            "verification": result.get("verification"),
            "repair_history": result.get("repair_history", []),
            "attempts": len(result.get("repair_history", [])),
        }
    finally:
        # 无论成功或失败，都删除本次运行创建的沙箱副本。
        agent.cleanup()

        # 当前案例的 Agent 已跑完，把进度条推进一格。
        with _progress_lock:
            if _progress_task_id is not None:
                _progress.update(_progress_task_id, advance=1)


def evaluate_proposal(
    *,
    input: dict,
    output: dict,
    expected_output: dict,
    metadata: dict,
    **kwargs,
) -> Evaluation:
    """同时评估目标单元格和新公式是否与标准答案一致。"""

    # 标准答案来自 Langfuse 数据集的 expected_output。
    expected_target = expected_output["target_cell"].strip().lower()
    expected_formula = expected_output["gold_formula"].strip()

    # 实际答案来自 task 函数的 output。
    actual_target = str(output.get("target", "")).strip().lower()
    actual_formula = str(output.get("new_formula", "")).strip()

    # 分别判断目标单元格和公式文本是否完全匹配。
    localization_correct = actual_target == expected_target
    formula_correct = actual_formula == expected_formula

    # 只有两项都正确时，整体修复才得分 1。
    return Evaluation(
        name="repair_success",
        value=float(localization_correct and formula_correct),
        comment=(
            f"target_correct={localization_correct}, "
            f"formula_correct={formula_correct}"
        ),
    )

def evaluate_localization(
    *,
    input: dict,
    output: dict,
    expected_output: dict,
    metadata: dict,
    **kwargs,
) -> Evaluation:
    """只评估 Agent 是否定位到了正确的目标单元格。"""

    # 统一大小写和首尾空白，减少格式差异带来的误判。
    actual = str(output.get("target", "")).strip().lower()
    expected = expected_output["target_cell"].strip().lower()

    return Evaluation(name="localization_accuracy", value=float(actual == expected))


def evaluate_formula(
    *,
    input: dict,
    output: dict,
    expected_output: dict,
    metadata: dict,
    **kwargs,
) -> Evaluation:
    """只评估 Agent 生成的新公式是否与标准公式一致。"""

    # 公式本身通常区分大小写以外的文本格式；这里仅去除首尾空白。
    actual = str(output.get("new_formula", "")).strip()
    expected = expected_output["gold_formula"].strip()

    return Evaluation(name="formula_accuracy", value=float(actual == expected))


def _semantic_formula(formula: str) -> str:
    """把公式归一化为语义等价形式（相对模板 + 去掉冗余括号）。"""
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula.strip()
    # 双方用同一个占位宿主，相对偏移相同，比较结果不受宿主影响。
    dummy = CellRef(sheet="X", row=1, col=1)
    try:
        template = to_relative_template(formula, dummy)
    except Exception:
        return formula.strip()
    # 与 Verifier pattern 检查共用 normalize_template，保证两侧的
    # 语义等价口径一致（空白/一元负括号变体不再判为不等价）。
    return normalize_template(template)


def evaluate_formula_equivalent(
    *,
    input: dict,
    output: dict,
    expected_output: dict,
    metadata: dict,
    **kwargs,
) -> Evaluation:
    """按语义等价（忽略 $/引号/空白/冗余括号）比较新公式与标准公式。"""
    actual = _semantic_formula(str(output.get("new_formula", "")))
    expected = _semantic_formula(str(expected_output["gold_formula"]))
    return Evaluation(name="formula_equivalent", value=float(actual == expected))


def evaluate_first_attempt(
    *,
    input: dict,
    output: dict,
    expected_output: dict,
    metadata: dict,
    **kwargs,
) -> Evaluation:
    """首次修复尝试的公式是否就是标准公式。"""
    history = output.get("repair_history", []) or []
    first = history[0].get("new_formula", "") if history else output.get("new_formula", "")
    expected = str(expected_output["gold_formula"]).strip()
    return Evaluation(
        name="first_attempt_correct",
        value=float(str(first).strip() == expected),
    )


def evaluate_any_attempt(
    *,
    input: dict,
    output: dict,
    expected_output: dict,
    metadata: dict,
    **kwargs,
) -> Evaluation:
    """历史上任意一次尝试的公式是否等于标准公式。"""
    history = output.get("repair_history", []) or []
    expected = str(expected_output["gold_formula"]).strip()
    formulas = [a.get("new_formula", "") for a in history]
    formulas.append(str(output.get("new_formula", "")))
    return Evaluation(
        name="any_attempt_matches_gold",
        value=float(any(str(f).strip() == expected for f in formulas)),
    )


def evaluate_attempts(
    *,
    input: dict,
    output: dict,
    expected_output: dict,
    metadata: dict,
    **kwargs,
) -> Evaluation:
    """记录本次案例的修复尝试次数（重试成本指标）。"""
    history = output.get("repair_history", []) or []
    return Evaluation(name="attempts", value=float(len(history)))

       
def main():
    """启动 Langfuse 评测实验。"""

    # 客户端读取 Langfuse API 配置，并用于提交实验结果。
    langfuse = get_client()
    dataset = langfuse.get_dataset(DATASET_NAME)

    # 数据集条目数即进度条的总进度。
    total_items = len(dataset.items)

    # 在模块级进度条上注册一个任务，供 target() 内部逐条更新。
    global _progress_task_id
    _progress_task_id = _progress.add_task(
        description=f"正在评测 {DATASET_NAME}（{total_items} 个案例）",
        total=total_items,
    )

    # 使用已经上传的数据集运行 Agent，并执行三个评测器。
    # 进度条需要在事件循环运行期间持续刷新，因此先手动 start，结束后 stop。
    _progress.start()
    try:
        results = dataset.run_experiment(
            name=DATASET_NAME,
            description=(
                "SheetGuard structured repair evaluation: localization / formula / "
                "repair / formula-equivalent / first-attempt / any-attempt / attempts"
            ),
            task=target,
            evaluators=[
                evaluate_localization,
                evaluate_formula,
                evaluate_proposal,
                evaluate_formula_equivalent,
                evaluate_first_attempt,
                evaluate_any_attempt,
                evaluate_attempts,
            ],
            max_concurrency=1,
        )
    finally:
        # 停止进度条渲染，避免与后续 print 输出串行。
        _progress.stop()
        # 短生命周期脚本退出前刷新，确保 Trace 和评分已发送。
        langfuse.flush()

    # 输出评测对象或摘要，便于在命令行查看实验结果。
    # results.format() 含 emoji 等非 GBK 字符，必须走 UTF-8 包装的 console。
    # format 用内存结果，服务端丢分在此不可见（见 score_integrity.py）。
    # 先打印结果再对账：回填重试耗尽时 verify_item_results 显式抛错，
    # 但结果仍留在终端可见。
    _console.print(results.format())
    _console.print("[dim]正在对账服务端评分完整性……[/dim]")
    verify_item_results(langfuse, results.item_results, console=_console)
    _console.print("[green]评分完整性校验通过。[/green]")
    # 服务端偶发把一次运行拆成多条同名实验记录（见 verify_experiment_records）。
    verify_experiment_records(
        langfuse,
        run_name=results.run_name,
        expected_items=total_items,
        console=_console,
    )


if __name__ == "__main__":
    # 直接运行脚本时启动评测；被导入时不自动发起远程评测。
    main()
