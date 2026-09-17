"""SheetGuard 本地网页入口（FastAPI）。

路由约定：/api/* 为 JSON API，/ 与 /static/* 服务前端静态文件。
目录约定（CWD 相对，与 CLI 一致）：产物 out/<工作簿名>/round-N/；
上传工作簿 web_data/workbooks/<stem>.xlsx。
"""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Callable

from fastapi import FastAPI, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from sheetguard.agent.llm_config import model_configured
from sheetguard.app import review_store as rs
from sheetguard.app.inspection import inspection_payload
from sheetguard.app.review_service import (
    ReviewStateError,
    load_review_state,
    run_next_round,
    save_draft_entry,
    submit_review,
)
from sheetguard.app.rounds import (
    BATCH_SUCCESS_STATUSES,
    AuditReadError,
    load_audit,
    round_dir,
    write_round_artifacts,
)
from sheetguard.config import load_config
from sheetguard.web.jobs import Job, JobManager, JobRunningError

WEB_DATA_DIR = Path("web_data")
WORKBOOKS_DIR = WEB_DATA_DIR / "workbooks"

app = FastAPI(title="SheetGuard", docs_url=None, redoc_url=None)
JOB_MANAGER = JobManager()


@app.get("/api/health")
def health() -> dict:
    return {"status": "ok"}


def _workbook_path(name: str) -> Path:
    path = WORKBOOKS_DIR / f"{name}.xlsx"
    if not path.exists():
        raise HTTPException(status_code=404, detail=f"工作簿不存在：{name}")
    return path


@app.post("/api/workbooks")
async def upload_workbook(file: UploadFile) -> dict:
    """上传 .xlsx 工作簿并立即返回离线检测载荷（同名即覆盖续用会话）。"""
    name = Path(file.filename or "").stem
    if not name:
        raise HTTPException(status_code=422, detail="缺少文件名")
    if Path(file.filename or "").suffix.lower() != ".xlsx":
        raise HTTPException(status_code=422, detail="仅支持 .xlsx 文件")
    WORKBOOKS_DIR.mkdir(parents=True, exist_ok=True)
    dest = WORKBOOKS_DIR / f"{name}.xlsx"
    dest.write_bytes(await file.read())
    try:
        payload = inspection_payload(dest)
    except Exception as exc:  # openpyxl 损坏文件
        dest.unlink(missing_ok=True)
        raise HTTPException(status_code=422, detail=f"工作簿无法解析：{exc}") from exc
    return {"name": name, "inspect": payload}


@app.get("/api/workbooks")
def list_workbooks() -> dict:
    """列出已上传工作簿：repairing（修复中）> reviewable（有待审轮次）> idle。"""
    index = rs.load_workbook_index()
    items = []
    if WORKBOOKS_DIR.exists():
        for path in sorted(WORKBOOKS_DIR.glob("*.xlsx")):
            name = path.stem
            base = Path("out") / name
            rnd = rs.latest_reviewable_round(base)
            status = "repairing" if JOB_MANAGER.current_for(name) is not None else None
            if status is None:
                status = "reviewable" if rnd is not None else "idle"
            items.append({
                "name": name,
                "status": status,
                "latest_round": int(rnd.name.split("-")[1]) if rnd else None,
                "index_status": (index.get(name) or {}).get("status"),
            })
    return {"workbooks": items}


@app.post("/api/workbooks/{name}/inspect")
def inspect_workbook(name: str) -> dict:
    """对已上传工作簿重跑离线检测。"""
    path = _workbook_path(name)
    try:
        return inspection_payload(path)
    except Exception as exc:
        raise HTTPException(status_code=422, detail=f"工作簿无法解析：{exc}") from exc


@app.get("/api/model-status")
def model_status() -> dict:
    """模型配置与审查轮数上限（会话内常量，顶栏状态灯与轮次徽章共用）。"""
    return {
        "configured": model_configured(),
        "model": os.getenv("OPENAI_MODEL") or None,
        "max_review_rounds": load_config(None).max_review_rounds,
    }


def build_repair_fn(workbook_path: Path, cfg) -> Callable:
    """构造一次修复调用：真实管线（可被测试 monkeypatch 整体替换）。"""
    from sheetguard.agent.graph import SheetGuardGraph
    from sheetguard.agent.llm_config import build_model

    def repair_fn(path, repaired_path, review_context, progress_callback=None):
        agent = SheetGuardGraph(build_model(), config=cfg)
        try:
            kwargs = {"review_context": review_context}
            if progress_callback is not None:
                kwargs["progress_callback"] = progress_callback
            return agent.invoke(str(path), repaired_workbook=repaired_path, **kwargs)
        finally:
            agent.cleanup()
    return repair_fn


def _start_job(kind: str, name: str, fn_builder: Callable) -> dict:
    if JOB_MANAGER.busy():
        raise HTTPException(
            status_code=409,
            detail={"error_code": "job_running", "message": "已有修复任务在运行"},
        )
    try:
        cfg = load_config(None)
    except Exception as exc:
        raise HTTPException(status_code=422, detail=f"配置非法：{exc}") from exc
    fn = fn_builder(cfg)
    # busy() 预检与 start() 持锁重检之间存在窗口：竞态失败方由这里统一转 409
    try:
        job = JOB_MANAGER.start(kind, name, fn)
    except JobRunningError:
        raise HTTPException(
            status_code=409,
            detail={"error_code": "job_running", "message": "已有修复任务在运行"},
        ) from None
    return {"job_id": job.id, "kind": kind, "workbook": name}


@app.post("/api/repair/{name}")
def start_repair(name: str) -> dict:
    if not model_configured():
        raise HTTPException(
            status_code=409,
            detail={"error_code": "model_not_configured",
                    "message": "请先在 .env 配置 OPENAI_API_KEY / OPENAI_MODEL"},
        )
    path = _workbook_path(name)

    def fn_builder(cfg):
        repair_fn = build_repair_fn(path, cfg)
        # 轮次目录在请求时确定（round-N = 已有最大轮次 +1），任务线程内不重算
        artifacts = round_dir(path)

        def job_fn(job: Job) -> dict:
            result = repair_fn(
                str(path), artifacts / "repaired.xlsx", None, job.set_progress
            )
            audit = {"workbook": str(path), **result.get("audit", {})}
            write_round_artifacts(artifacts, audit)
            return {
                "round": int(artifacts.name.split("-")[1]),
                "dir": str(artifacts),
                "status": result.get("status", "unknown"),
                "success": result.get("status") in BATCH_SUCCESS_STATUSES,
            }
        return job_fn

    return _start_job("repair", name, fn_builder)


@app.get("/api/jobs/{job_id}")
def job_status(job_id: str) -> dict:
    job = JOB_MANAGER.get(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="任务不存在")
    return job.snapshot()


class DraftBody(BaseModel):
    target: str
    verdict: str
    remark: str | None = None


class SubmitBody(BaseModel):
    verdicts: list[DraftBody]


def _review_state_or_404(name: str) -> dict:
    base = Path("out") / name
    try:
        return load_review_state(base)
    except ReviewStateError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc


@app.get("/api/workbooks/{name}/review-items")
def review_items(name: str) -> dict:
    """最新待审轮次的审查条目（含已存草稿裁决与计数）。"""
    state = _review_state_or_404(name)
    counts = {"correct": 0, "incorrect": 0, "false_positive": 0, "skipped": 0}
    for v, _ in state["verdicts"].values():
        counts[v] += 1
    return {
        "round": state["round"],
        "items": [
            {"kind": kind,
             "target": e["target"],
             # dismissed 条目形状 {target, reason, formula}：old_formula 回退 formula
             "old_formula": e.get("old_formula") or e.get("formula"),
             "new_formula": e.get("new_formula") or e.get("proposal")
                            or e.get("last_formula"),
             "hypothesis": (e.get("hypothesis") or {}).get("hypothesis"),
             "error_type": (e.get("hypothesis") or {}).get("error_type"),
             "reason": e.get("reason"),
             "src_round": e.get("src_round")}
            for kind, e in state["review_items"]
        ],
        "verdicts": {t: [v, m] for t, (v, m) in state["verdicts"].items()},
        "counts": counts,
        "max_review_rounds": load_config(None).max_review_rounds,
    }


@app.post("/api/workbooks/{name}/review-draft")
def review_draft(name: str, body: DraftBody) -> dict:
    """单条裁决实时存草稿（同名 target 覆盖，续审不丢）。"""
    if body.verdict not in rs.VERDICTS:
        raise HTTPException(status_code=422, detail=f"invalid verdict: {body.verdict}")
    try:
        save_draft_entry(Path("out") / name, body.target, body.verdict, body.remark)
    except ReviewStateError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    return {"ok": True}


@app.post("/api/workbooks/{name}/review-submit")
def review_submit(name: str, body: SubmitBody) -> dict:
    """提交一轮裁决：feedback 落盘 + 样本沉淀；错项触发下一轮后台重修。

    只提交已判条目；未判条目由服务层按 skipped 落盘（轨迹完整）。
    load_config(None) 走内置默认——CLI 的 --config 语义在 web 第一版不支持。
    """
    # busy 预检必须在落盘之前：submit_review 会持久化 feedback/清草稿/沉淀
    # 样本，若落盘后才 409，本轮已被消费且无法从 web 触发下一轮（不可恢复）。
    if JOB_MANAGER.busy():
        raise HTTPException(
            status_code=409,
            detail={"error_code": "job_running", "message": "已有修复任务在运行"},
        )
    state = _review_state_or_404(name)
    verdicts = {v.target: (v.verdict, v.remark) for v in body.verdicts}
    for target, (verdict, _) in verdicts.items():
        if verdict not in {"correct", "incorrect", "false_positive", "skipped"}:
            raise HTTPException(status_code=422, detail=f"invalid verdict: {verdict}")
    result = submit_review(state, verdicts, load_config(None))
    payload = {
        "round": result["round_no"],
        "counts": result["counts"],
        "logs": result["logs"],
        "warnings": result["warnings"],
        "exhausted": result["exhausted"],
        "next_round": None,
    }
    if result["review_context"] is None:
        return payload

    def fn_builder(cfg):
        base = Path("out") / name

        def job_fn(job: Job) -> dict:
            outcome = run_next_round(
                base, result["round_no"], result["review_context"],
                build_repair_fn(_workbook_path(name), cfg), job.set_progress,
            )
            return {"round": outcome["round"], "dir": str(outcome["dir"]),
                    "status": outcome["status"]}
        return job_fn

    started = _start_job("review_next_round", name, fn_builder)
    payload["next_round"] = {"job_id": started["job_id"],
                             "round": result["round_no"] + 1}
    return payload


@app.get("/api/workbooks/{name}/rounds")
def rounds_listing(name: str) -> dict:
    _workbook_path(name)  # 404 若工作簿未上传
    base = Path("out") / name
    items = []
    for d in sorted(rs.round_dirs(base), key=lambda d: int(d.name.split("-")[1])):
        n = int(d.name.split("-")[1])
        audit = {}
        audit_path = d / "audit.json"
        if audit_path.exists():
            try:
                audit = json.loads(audit_path.read_text(encoding="utf-8"))
            except (OSError, ValueError):
                audit = {}
        items.append({
            "round": n,
            "status": audit.get("status"),
            "has_audit": audit_path.exists(),
            "has_repaired": (d / "repaired.xlsx").exists(),
            "has_feedback": (d / rs.FEEDBACK_NAME).exists(),
        })
    return {"rounds": items}


def _round_dir_or_404(name: str, n: int) -> Path:
    _workbook_path(name)
    d = Path("out") / name / f"round-{n}"
    if not d.is_dir():
        raise HTTPException(status_code=404, detail=f"round-{n} 不存在")
    return d


@app.get("/api/workbooks/{name}/rounds/{n}/audit")
def round_audit(name: str, n: int) -> dict:
    d = _round_dir_or_404(name, n)
    try:
        return load_audit(d / "audit.json")
    except AuditReadError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc


@app.get("/api/workbooks/{name}/rounds/{n}/repaired")
def round_repaired(name: str, n: int) -> FileResponse:
    d = _round_dir_or_404(name, n)
    p = d / "repaired.xlsx"
    if not p.exists():
        raise HTTPException(status_code=404, detail="该轮无修复副本")
    return FileResponse(p, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        filename=f"{name}-round{n}-repaired.xlsx")


@app.get("/api/workbooks/{name}/rounds/{n}/report")
def round_report(name: str, n: int) -> FileResponse:
    d = _round_dir_or_404(name, n)
    p = d / "report.md"
    if not p.exists():
        raise HTTPException(status_code=404, detail="该轮无报告")
    return FileResponse(p, media_type="text/markdown; charset=utf-8", filename=f"round-{n}-report.md")


_STATIC = Path(__file__).parent / "static"
app.mount("/static", StaticFiles(directory=_STATIC), name="static")


@app.get("/")
def index() -> FileResponse:
    return FileResponse(_STATIC / "index.html")
