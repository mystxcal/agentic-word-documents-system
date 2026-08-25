from __future__ import annotations

import ctypes
import json
import os
import uuid
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterator

from .errors import DocumentSystemError


def _process_exists(pid: int) -> bool:
    if pid <= 0:
        return False
    if os.name == "nt":
        process_query_limited_information = 0x1000
        handle = ctypes.windll.kernel32.OpenProcess(  # type: ignore[attr-defined]
            process_query_limited_information, False, pid
        )
        if handle:
            ctypes.windll.kernel32.CloseHandle(handle)  # type: ignore[attr-defined]
            return True
        return False
    try:
        os.kill(pid, 0)
        return True
    except OSError:
        return False


def _read_lock(path: Path) -> dict:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
        return value if isinstance(value, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}


@contextmanager
def document_lock(system_root: Path, document_id: str, operation: str) -> Iterator[Path]:
    """Prevent two mutating operator commands from racing on one document."""

    lock_root = Path(system_root) / ".locks"
    lock_root.mkdir(parents=True, exist_ok=True)
    lock_path = lock_root / f"{document_id}.json"
    token = uuid.uuid4().hex
    record = {
        "schema": "agentic-document-lock/v2",
        "document_id": document_id,
        "operation": operation,
        "pid": os.getpid(),
        "started_utc": datetime.now(timezone.utc).isoformat(),
        "token": token,
    }
    encoded = (json.dumps(record, indent=2) + "\n").encode("utf-8")
    while True:
        try:
            descriptor = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            with os.fdopen(descriptor, "wb") as handle:
                handle.write(encoded)
                handle.flush()
                os.fsync(handle.fileno())
            break
        except FileExistsError:
            existing = _read_lock(lock_path)
            pid = existing.get("pid")
            if isinstance(pid, int) and _process_exists(pid):
                raise DocumentSystemError(
                    f"Document {document_id!r} is already being changed by "
                    f"{existing.get('operation', 'another operation')} (process {pid})."
                )
            stale_root = lock_root / "stale"
            stale_root.mkdir(parents=True, exist_ok=True)
            stale_name = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S_%fZ")
            try:
                os.replace(lock_path, stale_root / f"{document_id}-{stale_name}.json")
            except FileNotFoundError:
                continue
    try:
        yield lock_path
    finally:
        current = _read_lock(lock_path)
        if current.get("token") == token:
            try:
                lock_path.unlink()
            except FileNotFoundError:
                pass
