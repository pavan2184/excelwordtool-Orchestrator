"""Session management for the orchestrator."""

import logging
import shutil
import tempfile
import time
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from threading import Thread
from typing import Any


@dataclass
class Session:
    session_id: str
    temp_dir: Path
    template_path: Path
    created_at: float = field(default_factory=time.time)
    log_path: Path | None = None

    # User input
    instruction: str = ""
    xlsx_path: Path | None = None

    # Pipeline state
    step: str = "uploaded"  # uploaded | extracted | confirmed | filling | verified | revising
    report_data: dict | None = None
    template_placeholders: list[dict] | None = None
    suggested_mapping: dict[str, Any] | None = None
    confirmed_data: dict | None = None
    skip_fields: list[str] = field(default_factory=list)

    # Word-tool state
    word_session_id: str | None = None
    verification_result: dict | None = None
    revision_count: int = 0

    # Background fill state
    fill_started_at: float = 0.0
    fill_result: dict | None = None
    fill_error: str | None = None


class SessionManager:
    def __init__(self, ttl_seconds: int = 3600):
        self.sessions: dict[str, Session] = {}
        self.ttl = ttl_seconds
        t = Thread(target=self._cleanup_loop, daemon=True)
        t.start()

    def create(self, template_filename: str, template_bytes: bytes) -> Session:
        session_id = uuid.uuid4().hex[:12]
        temp_dir = Path(tempfile.mkdtemp(prefix=f"orch_{session_id}_"))
        template_path = temp_dir / template_filename
        template_path.write_bytes(template_bytes)

        # Create per-session log file
        log_path = temp_dir / "session.log"

        session = Session(
            session_id=session_id,
            temp_dir=temp_dir,
            template_path=template_path,
            log_path=log_path,
        )

        # Add a file handler for this session to the orchestrator logger
        file_handler = logging.FileHandler(str(log_path), mode="w")
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(
            logging.Formatter("%(asctime)s [%(levelname)s] %(name)s — %(message)s")
        )
        # Tag the handler so we can find it later
        file_handler.session_id = session_id  # type: ignore[attr-defined]

        # Attach to the root logger so ALL log output (orchestrator, excel_reader,
        # template_scanner, word_client, etc.) goes to this file
        logging.getLogger().addHandler(file_handler)

        self.sessions[session_id] = session
        return session

    def get(self, session_id: str) -> Session | None:
        return self.sessions.get(session_id)

    def remove(self, session_id: str):
        session = self.sessions.pop(session_id, None)
        if session:
            shutil.rmtree(session.temp_dir, ignore_errors=True)

    def _cleanup_loop(self):
        while True:
            time.sleep(300)
            now = time.time()
            expired = [
                sid for sid, s in self.sessions.items()
                if now - s.created_at > self.ttl
            ]
            for sid in expired:
                self.remove(sid)
