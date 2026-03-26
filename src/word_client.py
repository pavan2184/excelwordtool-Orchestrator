"""HTTP client for the Word editing tool (docx-editor)."""

from __future__ import annotations

import json
import logging
from typing import Any

import httpx

logger = logging.getLogger("word_client")

DEFAULT_WORD_TOOL_URL = "http://localhost:8000"
TIMEOUT = 300.0  # 5 minutes — agent loop can be slow


class WordToolClient:
    """Client that calls the docx-editor API over HTTP."""

    def __init__(self, base_url: str = DEFAULT_WORD_TOOL_URL):
        self.base_url = base_url.rstrip("/")

    async def upload_with_data(
        self,
        template_bytes: bytes,
        template_filename: str,
        instruction: str,
        report_data: dict,
    ) -> dict:
        """Call POST /upload-with-data on the word-tool.

        Returns the JSON response with session_id, summary, verification, etc.
        """
        async with httpx.AsyncClient(timeout=TIMEOUT) as client:
            files = {"file": (template_filename, template_bytes, "application/octet-stream")}
            data = {
                "instruction": instruction,
                "report_data_json": json.dumps(report_data),
            }
            logger.info(f"Calling word-tool /upload-with-data — template: {template_filename}")
            logger.debug(f"  Payload instruction: {instruction!r}")
            logger.debug(f"  Payload report_data_json ({len(data['report_data_json'])} chars):")
            logger.debug(f"  {data['report_data_json'][:2000]}")
            if len(data['report_data_json']) > 2000:
                logger.debug(f"  ... ({len(data['report_data_json']) - 2000} more chars)")
            resp = await client.post(f"{self.base_url}/upload-with-data", files=files, data=data)
            resp.raise_for_status()
            result = resp.json()
            logger.info(f"Word-tool response — session: {result.get('session_id')}, "
                        f"modified: {result.get('has_modified_file')}")
            logger.debug(f"  Full response: {json.dumps(result, indent=2, default=str)[:3000]}")
            return result

    async def revise(
        self,
        session_id: str,
        answers: list[dict[str, str]],
    ) -> dict:
        """Call POST /revise on the word-tool."""
        async with httpx.AsyncClient(timeout=TIMEOUT) as client:
            data = {
                "session_id": session_id,
                "answers": json.dumps(answers),
            }
            logger.info(f"Calling word-tool /revise — session: {session_id}, answers: {len(answers)}")
            logger.debug(f"  Answers payload: {json.dumps(answers, indent=2, default=str)}")
            resp = await client.post(f"{self.base_url}/revise", data=data)
            resp.raise_for_status()
            result = resp.json()
            logger.debug(f"  Revise response: {json.dumps(result, indent=2, default=str)[:3000]}")
            return result

    async def download(self, session_id: str) -> bytes:
        """Call GET /download/{session_id} on the word-tool and return file bytes."""
        async with httpx.AsyncClient(timeout=TIMEOUT) as client:
            logger.info(f"Calling word-tool /download/{session_id}")
            resp = await client.get(f"{self.base_url}/download/{session_id}")
            resp.raise_for_status()
            return resp.content

    async def get_logs(self, session_id: str) -> str | None:
        """Fetch agent logs from the word-tool if the endpoint exists."""
        try:
            async with httpx.AsyncClient(timeout=10.0) as client:
                resp = await client.get(f"{self.base_url}/logs/{session_id}")
                if resp.status_code == 200:
                    return resp.text
        except Exception:
            pass
        return None

    async def health_check(self) -> bool:
        """Check if the word-tool is reachable."""
        try:
            async with httpx.AsyncClient(timeout=5.0) as client:
                resp = await client.get(f"{self.base_url}/")
                return resp.status_code == 200
        except Exception:
            return False
