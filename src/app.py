"""Orchestrator — ties Excel extraction + Word editing together."""

import asyncio
import json
import logging
import sys
import os
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pathlib import Path

from fastapi import FastAPI, File, Form, UploadFile, HTTPException
from fastapi.responses import HTMLResponse, PlainTextResponse, Response

from src.session import Session, SessionManager
from src.excel_reader import extract_report_data
from src.data_formatter import format_report_data_as_text, report_data_to_dict, report_data_from_dict
from src.template_scanner import scan_template, suggest_mapping
from src.word_client import WordToolClient

logging.basicConfig(level=logging.DEBUG, format="%(asctime)s [%(levelname)s] %(name)s — %(message)s")
logger = logging.getLogger("orchestrator")

# Keep noisy libraries quiet
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("httpcore").setLevel(logging.WARNING)
logging.getLogger("uvicorn.access").setLevel(logging.INFO)

app = FastAPI(title="Document Orchestrator")
session_manager = SessionManager()
word_client = WordToolClient()

TEMPLATES_DIR = Path(__file__).parent / "templates"


# ── Pages ──────────────────────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse)
async def index():
    return (TEMPLATES_DIR / "orchestrator.html").read_text()


# ── Upload & Fill (single step) ───────────────────────────────────────

@app.post("/api/start")
async def start(
    template: UploadFile = File(...),
    data_source: UploadFile = File(...),
    instruction: str = Form(...),
):
    """Upload template + data source, extract data, and immediately start filling."""

    # Validate template
    if not template.filename or not template.filename.lower().endswith(".docx"):
        raise HTTPException(400, "Template must be a .docx file.")
    template_bytes = await template.read()
    if len(template_bytes) > 50 * 1024 * 1024:
        raise HTTPException(400, "Template too large (max 50 MB).")

    # Validate data source
    if not data_source.filename or not data_source.filename.lower().endswith(".xlsx"):
        raise HTTPException(400, "Data source must be an .xlsx file.")
    xlsx_bytes = await data_source.read()
    if len(xlsx_bytes) > 50 * 1024 * 1024:
        raise HTTPException(400, "Data source too large (max 50 MB).")

    logger.info(f"Start — template: '{template.filename}', data: '{data_source.filename}', "
                f"instruction: '{instruction[:100]}'")

    # Create session
    session = session_manager.create(template.filename, template_bytes)
    session.instruction = instruction
    session.xlsx_path = session.temp_dir / data_source.filename
    session.xlsx_path.write_bytes(xlsx_bytes)

    # Extract data from Excel
    try:
        report = extract_report_data(xlsx_bytes, filename=data_source.filename)
    except ValueError as e:
        session_manager.remove(session.session_id)
        raise HTTPException(400, f"Failed to read Excel file: {e}")

    report_dict = report_data_to_dict(report)
    session.report_data = report_dict

    # ── Detailed extraction log ──
    sid = session.session_id
    logger.info(f"[{sid}] ── Excel Extraction Results ──")
    logger.info(f"[{sid}]   Source: {data_source.filename}")
    logger.info(f"[{sid}]   Sheets: {report.metadata.get('sheet_names', [])}")
    logger.info(f"[{sid}]   Total rows read: {report.metadata.get('total_rows', 0)}")
    logger.info(f"[{sid}]   Fields extracted ({len(report.fields)}):")
    for k, v in report.fields.items():
        logger.info(f"[{sid}]     {k}: {v!r}")
    logger.info(f"[{sid}]   Tables extracted ({len(report.tables)}):")
    for t in report.tables:
        logger.info(f"[{sid}]     Table '{t.name}': {len(t.headers)} cols, {len(t.rows)} rows")

    # Scan template for placeholders (for logging only)
    try:
        placeholders = scan_template(template_bytes)
        logger.info(f"[{sid}] ── Template Placeholders ({len(placeholders)}) ──")
        for p in placeholders:
            logger.info(f"[{sid}]   '{p.field_name}' [{p.location}] current='{p.current_value}'")
    except Exception as e:
        logger.warning(f"[{sid}] Template scan failed: {e}")

    # Check word-tool is available
    if not await word_client.health_check():
        session_manager.remove(session.session_id)
        raise HTTPException(503, "Word editing tool is not available. Make sure it's running on localhost:8000.")

    # Go straight to filling — no confirmation step
    session.confirmed_data = report_dict
    session.step = "filling"
    session.fill_started_at = time.time()

    logger.info(f"[{sid}] ── Starting Fill (auto-confirmed) ──")
    logger.info(f"[{sid}]   Fields: {len(report_dict.get('fields', {}))}, "
                f"Tables: {len(report_dict.get('tables', []))}")
    logger.info(f"[{sid}]   Instruction: {instruction!r}")

    asyncio.create_task(_run_fill_in_background(session, report_dict))

    return {
        "session_id": session.session_id,
        "status": "filling",
        "report_data": report_dict,
        "message": "Data extracted and fill started. Poll /api/fill-status for progress.",
    }


async def _run_fill_in_background(session: Session, report_data: dict):
    """Background task: call word-tool and update session when done."""
    sid = session.session_id
    try:
        start_time = time.time()
        template_bytes = session.template_path.read_bytes()

        result = await word_client.upload_with_data(
            template_bytes=template_bytes,
            template_filename=session.template_path.name,
            instruction=session.instruction,
            report_data=report_data,
        )
        elapsed = time.time() - start_time

        # ── Detailed word-tool response log ──
        logger.info(f"[{sid}] ── Word Tool Response (took {elapsed:.1f}s) ──")
        logger.info(f"[{sid}]   Word session ID: {result.get('session_id')}")
        logger.info(f"[{sid}]   Has modified file: {result.get('has_modified_file')}")
        logger.info(f"[{sid}]   Summary: {result.get('summary', '(none)')}")
        verification = result.get("verification", {})
        if verification:
            logger.info(f"[{sid}]   Verification pass: {verification.get('pass')}")
            logger.info(f"[{sid}]   Verification score: {verification.get('score')}")
            issues = verification.get("issues", [])
            if issues:
                logger.info(f"[{sid}]   Issues ({len(issues)}):")
                for iss in issues:
                    logger.info(f"[{sid}]     {iss}")
            questions = verification.get("questions", [])
            if questions:
                logger.info(f"[{sid}]   Questions ({len(questions)}):")
                for q in questions:
                    logger.info(f"[{sid}]     {q}")

        session.word_session_id = result.get("session_id")
        session.verification_result = result.get("verification")
        session.fill_result = result
        session.step = "verified"

        # Fetch and log word-tool's agent logs if available
        try:
            word_logs = await word_client.get_logs(result.get("session_id"))
            if word_logs:
                logger.info(f"[{sid}] ── Word Tool Agent Logs ──")
                for line in word_logs.splitlines():
                    logger.info(f"[{sid}] [word-tool] {line}")
        except Exception:
            pass  # Non-critical

    except Exception as e:
        logger.error(f"[{sid}] Word-tool error: {e}", exc_info=True)
        session.fill_error = str(e)
        session.step = "fill_failed"


@app.get("/api/fill-status/{session_id}")
async def fill_status(session_id: str):
    """Poll this endpoint while the word-tool is filling the document."""
    session = session_manager.get(session_id)
    if session is None:
        raise HTTPException(404, "Session not found or expired.")

    elapsed = time.time() - getattr(session, "fill_started_at", time.time())

    if session.step == "filling":
        return {
            "status": "filling",
            "elapsed_seconds": round(elapsed, 1),
            "message": f"Still working... ({round(elapsed)}s elapsed)",
        }

    if session.step == "fill_failed":
        error = getattr(session, "fill_error", "Unknown error")
        session.step = "uploaded"  # allow retry
        raise HTTPException(500, f"Word editing tool error: {error}")

    if session.step == "verified":
        result = getattr(session, "fill_result", {})
        return {
            "status": "done",
            "elapsed_seconds": round(elapsed, 1),
            "session_id": session.session_id,
            "word_session_id": session.word_session_id,
            "summary": result.get("summary"),
            "has_modified_file": result.get("has_modified_file"),
            "verification": result.get("verification"),
        }

    return {"status": session.step, "elapsed_seconds": round(elapsed, 1)}


# ── Revise ────────────────────────────────────────────────────────────

@app.post("/api/revise")
async def revise(
    session_id: str = Form(...),
    answers: str = Form(...),
):
    """User answers verification questions, re-run via word-tool."""

    session = session_manager.get(session_id)
    if session is None:
        raise HTTPException(404, "Session not found or expired.")
    if session.step != "verified":
        raise HTTPException(400, f"Cannot revise in step '{session.step}'.")
    if not session.word_session_id:
        raise HTTPException(400, "No word-tool session to revise.")

    try:
        qa_pairs = json.loads(answers)
    except json.JSONDecodeError:
        raise HTTPException(400, "Invalid answers JSON.")

    session.step = "revising"
    session.revision_count += 1

    logger.info(f"[{session.session_id}] Revise #{session.revision_count} — {len(qa_pairs)} answers")

    try:
        result = await word_client.revise(session.word_session_id, qa_pairs)
    except Exception as e:
        logger.error(f"[{session.session_id}] Word-tool revise error: {e}", exc_info=True)
        session.step = "verified"
        raise HTTPException(500, f"Word editing tool error: {e}")

    session.verification_result = result.get("verification")
    session.step = "verified"

    return {
        "session_id": session.session_id,
        "word_session_id": session.word_session_id,
        "summary": result.get("summary"),
        "has_modified_file": result.get("has_modified_file"),
        "verification": result.get("verification"),
    }


# ── Download ───────────────────────────────────────────────────────────

@app.get("/api/download/{session_id}")
async def download(session_id: str):
    """Download the edited document from the word-tool."""

    session = session_manager.get(session_id)
    if session is None:
        raise HTTPException(404, "Session not found or expired.")
    if not session.word_session_id:
        raise HTTPException(400, "No edited document available.")

    try:
        file_bytes = await word_client.download(session.word_session_id)
    except Exception as e:
        logger.error(f"[{session.session_id}] Download error: {e}", exc_info=True)
        raise HTTPException(500, f"Download failed: {e}")

    filename = f"{session.template_path.stem}_filled.docx"
    return Response(
        content=file_bytes,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# ── Status ─────────────────────────────────────────────────────────────

@app.get("/api/status/{session_id}")
async def status(session_id: str):
    session = session_manager.get(session_id)
    if session is None:
        raise HTTPException(404, "Session not found or expired.")

    return {
        "session_id": session.session_id,
        "step": session.step,
        "has_report_data": session.report_data is not None,
        "word_session_id": session.word_session_id,
        "verification": session.verification_result,
        "revision_count": session.revision_count,
    }


# ── Logs ───────────────────────────────────────────────────────────────

@app.get("/api/logs/{session_id}", response_class=PlainTextResponse)
async def get_logs(session_id: str):
    """Return the full session log file as plain text."""
    session = session_manager.get(session_id)
    if session is None:
        raise HTTPException(404, "Session not found or expired.")
    if not session.log_path or not session.log_path.exists():
        raise HTTPException(404, "No log file for this session.")

    # Flush all handlers to ensure latest logs are written
    for handler in logging.getLogger().handlers:
        handler.flush()

    return session.log_path.read_text(errors="replace")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8001)
