"""
Paper Downloader web application.

Upload an Excel/CSV list -> download legally available OA PDFs -> review status & Excel report.
"""

from __future__ import annotations

import csv
import io
import json
import shutil
import threading
import uuid
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Optional

from fastapi import BackgroundTasks, FastAPI, File, Form, HTTPException, Query, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

import download_papers as dl

BASE_DIR = Path(__file__).resolve().parent
JOBS_DIR = BASE_DIR / "jobs"
WEB_DIR = BASE_DIR / "web"
MAX_UPLOAD_BYTES = 40 * 1024 * 1024
ALLOWED_EXTENSIONS = (".xlsx", ".xlsm", ".xls", ".csv")

JOBS_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="PaperFetch", version="1.1.0")
app.mount("/static", StaticFiles(directory=WEB_DIR / "static"), name="static")

_lock = threading.Lock()
_download_lock = threading.Lock()
_jobs: dict[str, dict[str, Any]] = {}


def _utc_now() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _job_dir(job_id: str) -> Path:
    return JOBS_DIR / job_id


def _state_path(job_id: str) -> Path:
    return _job_dir(job_id) / "state.json"


def _public_job(job: dict[str, Any], since: int = 0) -> dict[str, Any]:
    """Return a JSON-safe view of a job (supports incremental results via `since`)."""
    all_results = job.get("results") or []
    offset = max(0, min(since, len(all_results)))
    return {
        "id": job["id"],
        "status": job["status"],
        "created_at": job["created_at"],
        "updated_at": job["updated_at"],
        "filename": job["filename"],
        "email": job["email"],
        "dry_run": job["dry_run"],
        "limit": job["limit"],
        "cancel_requested": bool(job.get("cancel_requested")),
        "error": job.get("error") or "",
        "progress": job.get("progress") or {},
        "summary": job.get("summary") or {},
        "results": all_results[offset:],
        "results_offset": offset,
        "results_total": len(all_results),
        "current": job.get("current") or None,
    }


def _save_job(job: dict[str, Any]) -> None:
    path = _state_path(job["id"])
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {k: v for k, v in job.items() if k != "workspace"}
    payload["workspace"] = str(job["workspace"])
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def _load_job(job_id: str) -> Optional[dict[str, Any]]:
    with _lock:
        if job_id in _jobs:
            return _jobs[job_id]
    path = _state_path(job_id)
    if not path.exists():
        return None
    data = json.loads(path.read_text(encoding="utf-8"))
    data["workspace"] = Path(data["workspace"])
    with _lock:
        _jobs[job_id] = data
    return data


def _update_job(job_id: str, *, persist: bool = True, **fields: Any) -> None:
    with _lock:
        job = _jobs[job_id]
        job.update(fields)
        job["updated_at"] = _utc_now()
        if persist:
            _save_job(job)


def _write_report_csv(job: dict[str, Any]) -> Path:
    report_path = Path(job["workspace"]) / "report.csv"
    fields = [
        "paper_number",
        "title",
        "doi",
        "source",
        "year",
        "status",
        "publisher",
        "destination_folder",
        "filename",
        "pdf_url",
        "oa_source",
        "error",
    ]
    with report_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        for row in job.get("results") or []:
            writer.writerow(row)
    return report_path


def _write_report_xlsx(job: dict[str, Any]) -> Path:
    report_path = Path(job["workspace"]) / "report.xlsx"
    return dl.generate_excel_report(
        list(job.get("results") or []),
        dict(job.get("summary") or {}),
        report_path,
        source_filename=str(job.get("filename") or ""),
        email=str(job.get("email") or ""),
        dry_run=bool(job.get("dry_run")),
    )


def _run_job(job_id: str) -> None:
    job = _load_job(job_id)
    if job is None:
        return

    workspace = Path(job["workspace"])
    input_name = job.get("input_filename") or "input.xlsx"
    excel_path = workspace / input_name

    def cancel_requested() -> bool:
        with _lock:
            return bool(_jobs.get(job_id, {}).get("cancel_requested"))

    def on_progress(event: dict[str, Any]) -> None:
        kind = event.get("event")
        if kind == "started":
            _update_job(
                job_id,
                persist=True,
                status="running",
                progress={
                    "index": 0,
                    "of": event.get("to_process", 0),
                    "percent": 0,
                    "total_papers": event.get("total_papers", 0),
                    "skipped": event.get("skipped", 0),
                },
            )
        elif kind == "paper_start":
            of = max(1, int(event.get("of") or 1))
            index = int(event.get("index") or 0)
            with _lock:
                cur = _jobs[job_id]
                prev_progress = cur.get("progress") or {}
                cur["current"] = {
                    "index": index,
                    "of": of,
                    "paper_number": event.get("paper_number"),
                    "title": event.get("title"),
                    "doi": event.get("doi"),
                }
                cur["progress"] = {
                    "index": index - 1,
                    "of": of,
                    "percent": round(((index - 1) / of) * 100, 1),
                    "total_papers": prev_progress.get("total_papers")
                    or (cur.get("summary") or {}).get("total")
                    or of,
                }
                cur["updated_at"] = _utc_now()
        elif kind == "paper_done":
            result = event.get("result") or {}
            of = max(1, int(event.get("of") or 1))
            index = int(event.get("index") or 0)
            should_persist = (index % 5 == 0) or (index == of)
            with _lock:
                current = _jobs[job_id]
                results = list(current.get("results") or [])
                results.append(result)
                prev_progress = current.get("progress") or {}
                current["results"] = results
                current["summary"] = event.get("summary") or current.get("summary")
                current["progress"] = {
                    "index": index,
                    "of": of,
                    "percent": round((index / of) * 100, 1),
                    "total_papers": prev_progress.get("total_papers")
                    or (current["summary"] or {}).get("total")
                    or of,
                }
                current["updated_at"] = _utc_now()
                if should_persist:
                    _save_job(current)
        elif kind in {"finished", "cancelled"}:
            final_status = "cancelled" if kind == "cancelled" else "completed"
            with _lock:
                current = _jobs[job_id]
                if event.get("summary"):
                    current["summary"] = event["summary"]
                if event.get("results") is not None:
                    current["results"] = event["results"]
                current["status"] = final_status
                current["current"] = None
                prev_progress = current.get("progress") or {}
                done_count = len(current.get("results") or [])
                current["progress"] = {
                    **prev_progress,
                    "index": done_count,
                    "of": done_count if final_status == "cancelled" else prev_progress.get("of", done_count),
                    "percent": 100 if final_status == "completed" else prev_progress.get("percent", 0),
                }
                current["updated_at"] = _utc_now()
                _save_job(current)
            _write_report_csv(_jobs[job_id])
            _write_report_xlsx(_jobs[job_id])

    try:
        _update_job(job_id, persist=True, status="running", error="")
        # Module-level paths are shared inside download_papers, so serialize active downloads.
        with _download_lock:
            summary, results = dl.run_download(
                excel_path,
                job["email"],
                dry_run=bool(job.get("dry_run")),
                limit=int(job.get("limit") or 0),
                retry_failed=bool(job.get("retry_failed")),
                workspace=workspace,
                on_progress=on_progress,
                cancel_check=cancel_requested,
                verbose=False,
            )
        final_status = "cancelled" if cancel_requested() else "completed"
        with _lock:
            current = _jobs[job_id]
            current["summary"] = summary.__dict__
            current["results"] = results
            current["status"] = final_status
            current["current"] = None
            current["progress"] = {
                "index": len(results),
                "of": len(results),
                "percent": 100 if final_status == "completed" else current.get("progress", {}).get("percent", 0),
                "total_papers": summary.total,
            }
            current["updated_at"] = _utc_now()
            _save_job(current)
        _write_report_csv(_jobs[job_id])
        _write_report_xlsx(_jobs[job_id])
    except Exception as exc:  # noqa: BLE001
        _update_job(job_id, persist=True, status="failed", error=str(exc), current=None)


@app.get("/", response_class=HTMLResponse)
async def index() -> HTMLResponse:
    html_path = WEB_DIR / "index.html"
    return HTMLResponse(html_path.read_text(encoding="utf-8"))


def _as_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    return str(value or "").strip().lower() in {"1", "true", "yes", "on"}


@app.post("/api/jobs")
async def create_job(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    email: str = Form(...),
    dry_run: str = Form("false"),
    limit: str = Form("0"),
) -> dict[str, Any]:
    clean_email = email.strip()
    if not clean_email or "@" not in clean_email or "." not in clean_email.split("@")[-1]:
        raise HTTPException(status_code=400, detail="Please enter a valid contact email address.")

    filename = file.filename or "upload.xlsx"
    suffix = Path(filename).suffix.lower()
    if suffix not in ALLOWED_EXTENSIONS:
        raise HTTPException(
            status_code=400,
            detail="Please upload a spreadsheet file (.xlsx, .xlsm, .xls, or .csv).",
        )

    try:
        limit_n = max(0, int(limit or 0))
    except ValueError as exc:
        raise HTTPException(status_code=400, detail="Limit must be a number.") from exc

    dry_run_flag = _as_bool(dry_run)

    data = await file.read()
    if not data:
        raise HTTPException(status_code=400, detail="Uploaded file is empty.")
    if len(data) > MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=400, detail="File exceeds the 40 MB upload limit.")

    job_id = uuid.uuid4().hex[:12]
    workspace = _job_dir(job_id)
    workspace.mkdir(parents=True, exist_ok=True)
    (workspace / "Downloaded_Papers").mkdir(exist_ok=True)
    (workspace / "Logs").mkdir(exist_ok=True)
    (workspace / "temp").mkdir(exist_ok=True)

    input_filename = f"input{suffix}"
    excel_path = workspace / input_filename
    excel_path.write_bytes(data)

    # Pure read-only validation: does not touch module globals or block on active downloads.
    try:
        df, columns = dl.load_excel(excel_path, verbose=False)
        papers = dl.dataframe_to_papers(df, columns)
    except Exception as exc:  # noqa: BLE001
        shutil.rmtree(workspace, ignore_errors=True)
        raise HTTPException(
            status_code=400,
            detail=f"Could not read the spreadsheet: {exc}",
        ) from exc

    if not papers:
        shutil.rmtree(workspace, ignore_errors=True)
        raise HTTPException(
            status_code=400,
            detail="No paper rows were found. Include Title or DOI columns.",
        )

    effective_total = min(len(papers), limit_n) if limit_n else len(papers)
    job: dict[str, Any] = {
        "id": job_id,
        "status": "queued",
        "created_at": _utc_now(),
        "updated_at": _utc_now(),
        "filename": filename,
        "input_filename": input_filename,
        "email": clean_email,
        "dry_run": dry_run_flag,
        "limit": limit_n,
        "retry_failed": False,
        "cancel_requested": False,
        "workspace": workspace,
        "error": "",
        "progress": {"index": 0, "of": effective_total, "percent": 0, "total_papers": effective_total},
        "summary": {
            "total": effective_total,
            "downloaded": 0,
            "already_exists": 0,
            "no_oa": 0,
            "access_required": 0,
            "invalid_doi": 0,
            "failed": 0,
            "not_found": 0,
            "skipped": 0,
        },
        "results": [],
        "current": None,
        "detected_columns": {
            "title": columns.title,
            "doi": columns.doi,
            "doi_url": columns.doi_url,
            "pmid": columns.pmid,
            "source": columns.source,
            "publisher": columns.publisher,
            "authors": columns.authors,
            "year": columns.year,
            "paper_number": columns.paper_number,
        },
    }

    with _lock:
        _jobs[job_id] = job
        _save_job(job)

    background_tasks.add_task(_run_job, job_id)
    return _public_job(job)


@app.get("/api/jobs/{job_id}")
async def get_job(job_id: str, since: int = Query(0, ge=0)) -> dict[str, Any]:
    job = _load_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="Job not found.")
    return _public_job(job, since=since)


@app.post("/api/jobs/{job_id}/cancel")
async def cancel_job(job_id: str) -> dict[str, Any]:
    job = _load_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="Job not found.")
    with _lock:
        current = _jobs[job_id]
        if current["status"] in {"queued", "running"}:
            current["cancel_requested"] = True
            current["updated_at"] = _utc_now()
            _save_job(current)
    return _public_job(job, since=len(job.get("results") or []))


@app.get("/api/jobs/{job_id}/report.xlsx")
async def download_report_excel(job_id: str) -> FileResponse:
    job = _load_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="Job not found.")
    path = _write_report_xlsx(job)
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"paper_report_{job_id}.xlsx",
    )


@app.get("/api/jobs/{job_id}/report.csv")
async def download_report(job_id: str) -> FileResponse:
    job = _load_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="Job not found.")
    path = _write_report_csv(job)
    return FileResponse(
        path,
        media_type="text/csv",
        filename=f"paper_report_{job_id}.csv",
    )


@app.get("/api/jobs/{job_id}/papers.zip")
async def download_papers_zip(job_id: str) -> StreamingResponse:
    job = _load_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="Job not found.")

    papers_dir = Path(job["workspace"]) / "Downloaded_Papers"
    if not papers_dir.exists():
        raise HTTPException(status_code=404, detail="No papers folder for this job.")

    pdfs = list(papers_dir.rglob("*.pdf"))
    if not pdfs:
        raise HTTPException(
            status_code=404,
            detail="No open-access PDFs were downloaded in this job.",
        )

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for path in pdfs:
            archive.write(path, arcname=str(path.relative_to(papers_dir)))
    buffer.seek(0)

    return StreamingResponse(
        buffer,
        media_type="application/zip",
        headers={
            "Content-Disposition": f'attachment; filename="papers_{job_id}.zip"'
        },
    )


@app.get("/api/health")
async def health() -> dict[str, str]:
    return {"status": "ok"}
