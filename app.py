from __future__ import annotations

import asyncio
import hashlib
import hmac
import os
import re
import secrets
import shlex
import sys
import uuid
from collections import deque
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from fastapi import Cookie, FastAPI, HTTPException, Request, Response
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel


ROOT_DIR = Path(__file__).resolve().parent.parent
STATIC_DIR = ROOT_DIR / "hb_web_panel" / "static"
SCRAPER_PATH = Path(os.getenv("SCRAPER_PATH", ROOT_DIR / "hepsiburada_price.py"))
OUTPUT_DIR = Path(os.getenv("REPORT_OUTPUT_DIR", ROOT_DIR / "generated_reports"))
APP_PASSWORD = os.getenv("APP_PASSWORD", "").strip()
SESSION_SECRET = os.getenv("APP_SESSION_SECRET") or secrets.token_hex(32)
COOKIE_NAME = "hb_panel_session"
MAX_LOG_LINES = 350

app = FastAPI(title="Hepsiburada Fiyat Takip Merkezi", docs_url=None, redoc_url=None)
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


class LoginPayload(BaseModel):
    password: str


@dataclass
class Job:
    id: str
    status: str = "queued"
    created_at: str = field(default_factory=lambda: datetime.now(timezone.utc).isoformat())
    started_at: str | None = None
    finished_at: str | None = None
    total: int = 114
    completed: int = 0
    found: int = 0
    not_found: int = 0
    errors: int = 0
    current_model: str = ""
    message: str = "Sıraya alındı"
    output_path: str = ""
    return_code: int | None = None
    logs: deque[str] = field(default_factory=lambda: deque(maxlen=MAX_LOG_LINES))
    completed_models: set[str] = field(default_factory=set)
    process: asyncio.subprocess.Process | None = field(default=None, repr=False)

    def public(self) -> dict[str, Any]:
        data = {
            "id": self.id,
            "status": self.status,
            "created_at": self.created_at,
            "started_at": self.started_at,
            "finished_at": self.finished_at,
            "total": self.total,
            "completed": self.completed,
            "found": self.found,
            "not_found": self.not_found,
            "errors": self.errors,
            "current_model": self.current_model,
            "message": self.message,
            "return_code": self.return_code,
            "logs": list(self.logs),
        }
        data["download_ready"] = (
            self.status == "completed"
            and bool(self.output_path)
            and Path(self.output_path).is_file()
        )
        data["progress"] = round(min(100, (self.completed / max(self.total, 1)) * 100), 1)
        return data


jobs: dict[str, Job] = {}
latest_job_id: str | None = None
job_lock = asyncio.Lock()

TOTAL_RE = re.compile(r"(\d+)\s+benzersiz model", re.IGNORECASE)
MODEL_START_RE = re.compile(r"^\s*→\s*(\S+)\s+aranıyor", re.IGNORECASE)
MODEL_DONE_RE = re.compile(r"^\s*✓\s*(\S+)\s+taraması tamamlandı", re.IGNORECASE)
RESULT_RE = re.compile(r"^\s*\[(\d+)/(\d+)\]\s+(\S+)\s+(.+)$")
SUMMARY_RE = re.compile(
    r"Bulunan:\s*(\d+)\s*\|\s*Bulunamayan:\s*(\d+)\s*\|\s*Hata:\s*(\d+)",
    re.IGNORECASE,
)


def session_token() -> str:
    return hmac.new(
        SESSION_SECRET.encode(),
        APP_PASSWORD.encode(),
        hashlib.sha256,
    ).hexdigest()


def is_authenticated(cookie: str | None) -> bool:
    if not APP_PASSWORD:
        return True
    return bool(cookie and hmac.compare_digest(cookie, session_token()))


def require_auth(cookie: str | None) -> None:
    if not is_authenticated(cookie):
        raise HTTPException(status_code=401, detail="Oturum açmanız gerekiyor.")


def get_job(job_id: str) -> Job:
    job = jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Rapor görevi bulunamadı.")
    return job


def parse_log_line(job: Job, line: str) -> None:
    clean = line.strip()
    if not clean:
        return
    job.logs.append(clean)

    if match := TOTAL_RE.search(clean):
        job.total = int(match.group(1))
    if match := MODEL_START_RE.match(clean):
        job.current_model = match.group(1)
        job.message = f"{job.current_model} aranıyor"
    if match := MODEL_DONE_RE.match(clean):
        model = match.group(1)
        job.completed_models.add(model)
        job.completed = len(job.completed_models)
        job.current_model = model
        job.message = f"{model} tamamlandı"
    if match := RESULT_RE.match(clean):
        job.total = int(match.group(2))
        job.current_model = match.group(3)
        outcome = match.group(4).lower()
        if "bulunamadı" in outcome:
            job.not_found += 1
        elif "hata" in outcome:
            job.errors += 1
        else:
            job.found += 1
    if match := SUMMARY_RE.search(clean):
        job.found = int(match.group(1))
        job.not_found = int(match.group(2))
        job.errors = int(match.group(3))


async def run_job(job: Job) -> None:
    job.status = "running"
    job.started_at = datetime.now(timezone.utc).isoformat()
    job.message = "Tarama başlatılıyor"
    output_path = Path(job.output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    command = [
        sys.executable,
        "-u",
        str(SCRAPER_PATH),
        "--workers",
        os.getenv("SCRAPER_WORKERS", "4"),
        "-o",
        str(output_path),
    ]
    extra_args = os.getenv("SCRAPER_EXTRA_ARGS", "").strip()
    if extra_args:
        command.extend(shlex.split(extra_args))

    try:
        job.process = await asyncio.create_subprocess_exec(
            *command,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.STDOUT,
            cwd=ROOT_DIR,
            env={**os.environ, "PYTHONUNBUFFERED": "1"},
        )
        assert job.process.stdout is not None
        while True:
            raw_line = await job.process.stdout.readline()
            if not raw_line:
                break
            parse_log_line(job, raw_line.decode("utf-8", errors="replace"))

        job.return_code = await job.process.wait()
        if job.status == "cancelled":
            return
        if job.return_code == 0 and output_path.is_file():
            job.status = "completed"
            job.completed = max(job.completed, job.total)
            job.message = "Rapor hazır"
        else:
            job.status = "failed"
            job.message = "Rapor oluşturulamadı; ayrıntılar çalışma günlüğünde."
    except asyncio.CancelledError:
        job.status = "cancelled"
        job.message = "Tarama iptal edildi"
        raise
    except Exception as exc:
        job.logs.append(f"Web paneli hatası: {exc}")
        job.status = "failed"
        job.message = "Beklenmeyen bir sunucu hatası oluştu."
    finally:
        job.finished_at = datetime.now(timezone.utc).isoformat()
        job.process = None


@app.on_event("startup")
async def startup() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


@app.get("/")
async def index() -> FileResponse:
    return FileResponse(STATIC_DIR / "index.html")


@app.get("/health")
async def health() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/api/config")
async def config(hb_panel_session: str | None = Cookie(default=None)) -> dict[str, Any]:
    return {
        "requires_auth": bool(APP_PASSWORD),
        "authenticated": is_authenticated(hb_panel_session),
        "product_count": 114,
    }


@app.post("/api/login")
async def login(payload: LoginPayload, request: Request, response: Response) -> dict[str, bool]:
    if not APP_PASSWORD or not hmac.compare_digest(payload.password, APP_PASSWORD):
        raise HTTPException(status_code=401, detail="Parola hatalı.")
    response.set_cookie(
        COOKIE_NAME,
        session_token(),
        httponly=True,
        secure=request.headers.get("x-forwarded-proto") == "https",
        samesite="lax",
        max_age=60 * 60 * 24 * 7,
    )
    return {"ok": True}


@app.post("/api/logout")
async def logout(response: Response) -> dict[str, bool]:
    response.delete_cookie(COOKIE_NAME)
    return {"ok": True}


@app.post("/api/jobs", status_code=202)
async def create_job(hb_panel_session: str | None = Cookie(default=None)) -> dict[str, Any]:
    global latest_job_id
    require_auth(hb_panel_session)
    if not SCRAPER_PATH.is_file():
        raise HTTPException(status_code=500, detail="Scraper dosyası sunucuda bulunamadı.")

    async with job_lock:
        for existing in jobs.values():
            if existing.status in {"queued", "running"}:
                raise HTTPException(
                    status_code=409,
                    detail="Zaten devam eden bir tarama var.",
                )
        job_id = uuid.uuid4().hex[:12]
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        output_path = OUTPUT_DIR / f"HB-Fiyat-Raporu-{stamp}-{job_id}.xlsx"
        job = Job(id=job_id, output_path=str(output_path))
        jobs[job_id] = job
        latest_job_id = job_id
        asyncio.create_task(run_job(job))
        return job.public()


@app.get("/api/jobs/current")
async def current_job(hb_panel_session: str | None = Cookie(default=None)) -> dict[str, Any]:
    require_auth(hb_panel_session)
    if not latest_job_id:
        return {"job": None}
    return {"job": jobs[latest_job_id].public()}


@app.get("/api/jobs/{job_id}")
async def job_status(
    job_id: str,
    hb_panel_session: str | None = Cookie(default=None),
) -> dict[str, Any]:
    require_auth(hb_panel_session)
    return get_job(job_id).public()


@app.post("/api/jobs/{job_id}/cancel")
async def cancel_job(
    job_id: str,
    hb_panel_session: str | None = Cookie(default=None),
) -> dict[str, Any]:
    require_auth(hb_panel_session)
    job = get_job(job_id)
    if job.status not in {"queued", "running"}:
        return job.public()

    job.status = "cancelled"
    job.message = "Tarama iptal ediliyor"
    if job.process and job.process.returncode is None:
        job.process.terminate()
        try:
            await asyncio.wait_for(job.process.wait(), timeout=8)
        except asyncio.TimeoutError:
            job.process.kill()
            await job.process.wait()
    job.finished_at = datetime.now(timezone.utc).isoformat()
    job.message = "Tarama iptal edildi"
    return job.public()


@app.get("/api/jobs/{job_id}/download")
async def download_report(
    job_id: str,
    hb_panel_session: str | None = Cookie(default=None),
) -> FileResponse:
    require_auth(hb_panel_session)
    job = get_job(job_id)
    output_path = Path(job.output_path)
    if job.status != "completed" or not output_path.is_file():
        raise HTTPException(status_code=409, detail="Rapor henüz indirilmeye hazır değil.")
    return FileResponse(
        output_path,
        filename=output_path.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
