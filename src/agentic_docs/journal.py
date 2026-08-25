from __future__ import annotations

import json
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


def append_activity(
    system_root: Path,
    document_id: str,
    operation: str,
    status: str,
    *,
    details: dict[str, Any] | None = None,
) -> Path:
    """Append one compact machine-readable operator event."""

    root = Path(system_root) / "operations" / document_id
    root.mkdir(parents=True, exist_ok=True)
    path = root / "activity.jsonl"
    record = {
        "schema": "agentic-activity/v2",
        "time_utc": datetime.now(timezone.utc).isoformat(),
        "operation": operation,
        "status": status,
        "details": details or {},
    }
    with path.open("a", encoding="utf-8", newline="\n") as handle:
        handle.write(json.dumps(record, default=str, separators=(",", ":")) + "\n")
        handle.flush()
        os.fsync(handle.fileno())
    return path
