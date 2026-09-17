"""沙箱工作簿修复器 —— copy-on-write 候选级事务管理。

本模块维护三层文件语义（设计 7.1），原始工作簿永远不被原地修改：

    source workbook     用户原始文件，只读
        ↓ 初始复制
    working workbook    当前已提交全部成功修复的稳定工作副本
        ↓ 每次候选/尝试复制
    attempt workbook    某个候选某次修复尝试的临时副本

单次修复尝试流程：
    working → copy → attempt → apply(patch) → 验证
    ├─ 验证通过：commit_attempt()，attempt 提升为新的 working
    └─ 验证失败：discard_attempt()，只删除临时副本，working 完全不变

因此失败尝试不需要任何"逆向回滚"——丢弃 attempt 文件即可，
下一次重试仍从当前稳定的 working workbook 复制新的 attempt。
"""
from __future__ import annotations
import os
import shutil
import tempfile
from pathlib import Path
from openpyxl import load_workbook


class SandboxPatcher:
    """copy-on-write 候选级工作簿事务管理器。

    职责：隔离源文件、管理 working/attempt 两级副本、应用补丁、
    提交或丢弃尝试、导出最终修复副本。类名保留 SandboxPatcher
    （设计允许第一阶段不改名），职责已扩展为工作簿事务管理。
    """

    def __init__(self, source_path: str | Path):
        self.source_path = Path(source_path)  # 用户原始文件，只读
        self.working_path: Path | None = None  # 稳定工作副本
        self.attempt_path: Path | None = None  # 当前修复尝试副本
        self.working_revision: int = 0  # 每次成功 commit 递增
        self.modified_cells: list[str] = []  # 已写入过补丁的单元格（兼容旧审计）

    @property
    def sandbox_path(self) -> Path | None:
        """兼容旧接口：单目标流程中的沙箱即当前工作副本。"""
        return self.working_path

    def create_working_copy(self) -> Path:
        """从源工作簿复制出稳定工作副本；已存在时复用。"""
        if self.working_path is None:
            self.working_path = self._copy_source("sheetguard_working_")
        return self.working_path

    def create_attempt(self) -> Path:
        """从当前稳定工作副本复制出独立的 attempt 副本。

        每次尝试都从最新 working 复制，保证此前成功提交的修复
        全部可见，而失败尝试永远不会污染 working。
        """
        if self.attempt_path is not None:
            self.discard_attempt()
        working = self.create_working_copy()
        fd, temp_path = tempfile.mkstemp(
            suffix=working.suffix, prefix="sheetguard_attempt_"
        )
        os.close(fd)
        shutil.copy2(working, temp_path)
        self.attempt_path = Path(temp_path)
        return self.attempt_path

    def apply(self, cell: str, new_formula: str, path: str | Path | None = None) -> Path:
        """在工作簿副本上应用补丁并返回其路径。

        默认写入当前 attempt（事务流程）；没有 attempt 时写入 working；
        连 working 也没有时先从源文件创建 working（兼容旧的单补丁用法）。
        """
        target = Path(path) if path is not None else self.attempt_path or self.create_working_copy()
        sheet_name, separator, coordinate = cell.partition("!")
        if not separator or not sheet_name or not coordinate:
            raise ValueError("cell must use 'Sheet!A1' format")

        wb = load_workbook(target)
        try:
            wb[sheet_name][coordinate] = new_formula
            wb.save(target)
        finally:
            wb.close()
        self.modified_cells.append(cell)
        return target

    def commit_attempt(self) -> Path:
        """把通过验证的 attempt 提升为新的稳定工作副本。

        提交后 working_revision 递增，后续候选从这个新的 working
        复制 attempt，能看到本次修复。
        """
        if self.attempt_path is None:
            raise ValueError("no attempt to commit; call create_attempt() first")
        if self.working_path is not None and self.working_path.exists():
            self.working_path.unlink()
        working = self.create_working_copy()
        os.replace(self.attempt_path, working)
        self.attempt_path = None
        self.working_revision += 1
        return working

    def discard_attempt(self) -> None:
        """丢弃当前 attempt 副本；稳定工作副本不受影响。"""
        if self.attempt_path is not None and self.attempt_path.exists():
            self.attempt_path.unlink()
        self.attempt_path = None

    def export(self, destination: str | Path) -> Path:
        """把最终稳定工作副本复制到用户指定的修复工作簿路径。"""
        if self.working_path is None or not self.working_path.exists():
            raise ValueError("no working workbook to export; run the workflow first")
        dest = Path(destination)
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(self.working_path, dest)
        return dest

    def cleanup(self) -> None:
        """删除当前 attempt 与 working 临时副本；可重复调用。"""
        self.discard_attempt()
        if self.working_path and self.working_path.exists():
            self.working_path.unlink()
        self.working_path = None

    def _copy_source(self, prefix: str) -> Path:
        """从源文件创建临时副本。"""
        fd, temp_path = tempfile.mkstemp(
            suffix=self.source_path.suffix, prefix=prefix
        )
        os.close(fd)
        shutil.copy2(self.source_path, temp_path)
        return Path(temp_path)
