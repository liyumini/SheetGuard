"""修复任务管理器：本地单机单并发 + 线程安全进度状态。"""
from __future__ import annotations

import threading
import uuid
from typing import Callable


class JobRunningError(RuntimeError):
    """已有修复任务在跑。"""


class Job:
    def __init__(self, job_id: str, kind: str, workbook: str):
        self.id = job_id
        self.kind = kind          # "repair" | "review_next_round"
        self.workbook = workbook  # 工作簿 stem
        self.status = "running"   # running | done | error
        self.progress: dict = {}
        self.result: dict | None = None
        self.error: str | None = None
        self._lock = threading.Lock()

    def set_progress(self, payload: dict) -> None:
        with self._lock:
            self.progress = dict(payload)

    def finish(self, result: dict) -> None:
        with self._lock:
            self.status = "done"
            self.result = result

    def fail(self, message: str) -> None:
        with self._lock:
            self.status = "error"
            self.error = message

    def snapshot(self) -> dict:
        with self._lock:
            return {
                "job_id": self.id,
                "kind": self.kind,
                "workbook": self.workbook,
                "status": self.status,
                "progress": dict(self.progress),
                "result": dict(self.result) if self.result else None,
                "error": self.error,
            }


class JobManager:
    def __init__(self) -> None:
        self._jobs: dict[str, Job] = {}
        self._current: str | None = None
        self._lock = threading.Lock()

    def busy(self) -> bool:
        with self._lock:
            return self._current is not None

    def start(self, kind: str, workbook: str, fn: Callable[[Job], dict]) -> Job:
        with self._lock:
            if self._current is not None:
                raise JobRunningError("已有修复任务在运行")
            job = Job(uuid.uuid4().hex[:12], kind, workbook)
            self._jobs[job.id] = job
            self._current = job.id
        thread = threading.Thread(target=self._run, args=(job, fn), daemon=True)
        thread.start()
        return job

    def _run(self, job: Job, fn: Callable[[Job], dict]) -> None:
        try:
            job.finish(fn(job))
        except Exception as exc:  # noqa: BLE001 —— 任务线程兜底，错误进快照
            job.fail(f"{type(exc).__name__}: {exc}")
        finally:
            with self._lock:
                if self._current == job.id:
                    self._current = None

    def get(self, job_id: str) -> Job | None:
        return self._jobs.get(job_id)

    def current_for(self, workbook: str) -> Job | None:
        with self._lock:
            current_id = self._current
        if current_id is None:
            return None
        job = self._jobs.get(current_id)
        return job if job is not None and job.workbook == workbook else None
