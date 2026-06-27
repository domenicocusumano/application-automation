"""
Application Automation — local web UI
Run: uvicorn app:app --reload --port 8000
Then open: http://localhost:8000
"""

import asyncio
import json
import os
from pathlib import Path
from typing import List
from docx import Document
from pypdf import PdfReader
from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse
from pydantic import BaseModel

app = FastAPI()
SCRIPT_DIR = Path(__file__).parent
PROMPT_FILES = {
    "background": "background_prompt.txt",
}
def _resume_stem():
    try:
        cfg = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        prefix = cfg.get("resume_prefix", "My_Resume").strip()
        return f"{prefix}_General_Resume" if prefix else "My_Resume_General_Resume"
    except Exception:
        return "My_Resume_General_Resume"

RESUME_STEM = _resume_stem()
RESUME_EXTENSIONS = (".pdf", ".docx")
CONFIG_FILE = SCRIPT_DIR / "config.json"
_run_lock = asyncio.Lock()


def find_resume_file():
    for ext in RESUME_EXTENSIONS:
        path = SCRIPT_DIR / f"{RESUME_STEM}{ext}"
        if path.exists():
            return path
    return None


def extract_docx_text(path):
    doc = Document(path)
    return "\n".join(p.text for p in doc.paragraphs)


def extract_pdf_text(path):
    reader = PdfReader(path)
    return "\n".join(page.extract_text(extraction_mode="layout") or "" for page in reader.pages)


@app.get("/", response_class=HTMLResponse)
async def index():
    return (SCRIPT_DIR / "ui.html").read_text(encoding="utf-8")


@app.get("/prompts/{key}")
async def get_prompt(key: str):
    if key not in PROMPT_FILES:
        raise HTTPException(status_code=404, detail="Unknown prompt key")
    content = (SCRIPT_DIR / PROMPT_FILES[key]).read_text(encoding="utf-8")
    return JSONResponse({"content": content})


class PromptBody(BaseModel):
    content: str


@app.post("/prompts/{key}")
async def save_prompt(key: str, body: PromptBody):
    if key not in PROMPT_FILES:
        raise HTTPException(status_code=404, detail="Unknown prompt key")
    (SCRIPT_DIR / PROMPT_FILES[key]).write_text(body.content, encoding="utf-8")
    return JSONResponse({"ok": True})


@app.get("/resume")
async def get_resume():
    path = find_resume_file()
    if path is None:
        return JSONResponse({"filename": None, "text": ""})
    if path.suffix.lower() == ".docx":
        text = extract_docx_text(path)
    else:
        text = extract_pdf_text(path)
    return JSONResponse({"filename": path.name, "text": text})


@app.post("/resume")
async def upload_resume(file: UploadFile = File(...)):
    ext = Path(file.filename).suffix.lower()
    if ext not in RESUME_EXTENSIONS:
        raise HTTPException(status_code=400, detail="Resume must be a .pdf or .docx file")
    existing = find_resume_file()
    if existing is not None:
        existing.unlink()
    content = await file.read()
    (SCRIPT_DIR / f"{RESUME_STEM}{ext}").write_bytes(content)
    return JSONResponse({"ok": True})


def _read_config():
    try:
        return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except (FileNotFoundError, ValueError):
        return {}


_DEFAULT_SENIORITY_TIERS = [
    "vp", "vice president", "cpo", "head of", "director",
    "principal", "staff", "lead", "senior", "group",
]
_DEFAULT_TITLE_KEYWORDS = [
    "product manager", "product management", "product lead", "product director",
    "vp product", "head of product", "chief product", "product owner",
]
_DEFAULT_EXCLUDED_TITLES = ["program manager", "project manager", "program management"]
_DEFAULT_EXCLUDED_TITLE_WORDS = [
    "engineer", "engineering", "architect", "developer", "scientist",
    "consultant", "devops", "sre",
]


@app.get("/config")
async def get_config():
    config = _read_config()
    return JSONResponse({
        "score_threshold":        config.get("score_threshold", 5.0),
        "top_applicant":          config.get("top_applicant", False),
        "linkedin_enabled":       config.get("linkedin_enabled", True),
        "linkedin_search_term":   config.get("linkedin_search_term", "Product Manager"),
        "builtin_enabled":        config.get("builtin_enabled", False),
        "builtin_url":            config.get("builtin_url", ""),
        "title_keywords":         config.get("title_keywords",        _DEFAULT_TITLE_KEYWORDS),
        "excluded_titles":        config.get("excluded_titles",        _DEFAULT_EXCLUDED_TITLES),
        "excluded_title_words":   config.get("excluded_title_words",   _DEFAULT_EXCLUDED_TITLE_WORDS),
        "seniority_tiers":        config.get("seniority_tiers",        _DEFAULT_SENIORITY_TIERS),
        "preferred_locations":    config.get("preferred_locations",    ["remote", "miami"]),
        "salary_minimum":         config.get("salary_minimum", 0),
        "google_sheet_url":       config.get("google_sheet_url", ""),
    })


class ConfigBody(BaseModel):
    score_threshold: float
    top_applicant: bool = False
    linkedin_enabled: bool = True
    linkedin_search_term: str = "Product Manager"
    builtin_enabled: bool = False
    builtin_url: str = ""
    title_keywords: List[str] = []
    excluded_titles: List[str] = []
    excluded_title_words: List[str] = []
    seniority_tiers: List[str] = []
    preferred_locations: List[str] = ["remote", "miami"]
    salary_minimum: float = 0
    google_sheet_url: str = ""


@app.post("/config")
async def save_config(body: ConfigBody):
    config = _read_config()
    config["score_threshold"]      = body.score_threshold
    config["top_applicant"]        = body.top_applicant
    config["linkedin_enabled"]     = body.linkedin_enabled
    config["linkedin_search_term"] = body.linkedin_search_term
    config["builtin_enabled"]      = body.builtin_enabled
    config["builtin_url"]          = body.builtin_url
    config["title_keywords"]       = body.title_keywords
    config["excluded_titles"]      = body.excluded_titles
    config["excluded_title_words"] = body.excluded_title_words
    config["seniority_tiers"]      = body.seniority_tiers
    config["preferred_locations"]  = body.preferred_locations
    config["salary_minimum"]       = body.salary_minimum
    config["google_sheet_url"]     = body.google_sheet_url
    CONFIG_FILE.write_text(json.dumps(config, indent=2) + "\n", encoding="utf-8")
    return JSONResponse({"ok": True})


@app.get("/check-session")
async def check_session():
    from job_scraper import check_linkedin_session
    valid, reason = check_linkedin_session()
    return JSONResponse({"valid": valid, "reason": reason})


_relogin_lock = asyncio.Lock()


@app.get("/relogin")
async def relogin():
    async def stream():
        if _relogin_lock.locked():
            yield "data: Re-login already in progress.\n\n"
            return
        async with _relogin_lock:
            proc = await asyncio.create_subprocess_exec(
                "python3", "-u", "relogin.py",
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.STDOUT,
                cwd=str(SCRIPT_DIR),
            )
            async for line in proc.stdout:
                text = line.decode(errors="replace").rstrip()
                yield f"data: {text}\n\n"
            await proc.wait()
            yield "data: [DONE]\n\n"

    return StreamingResponse(
        stream(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


_pipeline_proc = None


@app.post("/cancel")
async def cancel_pipeline():
    global _pipeline_proc
    if _pipeline_proc and _pipeline_proc.returncode is None:
        _pipeline_proc.terminate()
        return JSONResponse({"ok": True})
    return JSONResponse({"ok": False, "reason": "No pipeline running"})


@app.get("/run")
async def run_pipeline():
    global _pipeline_proc

    async def stream():
        global _pipeline_proc
        if _run_lock.locked():
            yield "data: Pipeline is already running.\n\n"
            return
        async with _run_lock:
            config = _read_config()
            if config.get("builtin_enabled", False):
                cmd = ["python3", "-u", "builtin_scraper.py"]
            else:
                cmd = ["python3", "-u", "job_scraper.py", "--resume"]
            _pipeline_proc = await asyncio.create_subprocess_exec(
                *cmd,
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.STDOUT,
                cwd=str(SCRIPT_DIR),
            )
            async for line in _pipeline_proc.stdout:
                text = line.decode(errors="replace").rstrip()
                yield f"data: {text}\n\n"
            await _pipeline_proc.wait()
            code = _pipeline_proc.returncode
            _pipeline_proc = None
            yield f"data: \n\n"
            yield f"data: ── Process exited with code {code} ──\n\n"
            yield "data: [DONE]\n\n"

    return StreamingResponse(
        stream(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )
