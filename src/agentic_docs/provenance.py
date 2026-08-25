from __future__ import annotations

import importlib.metadata
import json
import platform
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

from .rendering import poppler_status
from .resolver import file_hash


def _program(path_value: str | None) -> dict | None:
    if not path_value:
        return None
    path = Path(path_value).resolve()
    if not path.is_file():
        return {"path": str(path), "available": False}
    return {
        "path": str(path),
        "available": True,
        "sha256": file_hash(path),
        "size_bytes": path.stat().st_size,
        "modified_ns": path.stat().st_mtime_ns,
    }


def _word_provenance() -> dict:
    script = r"""
$ErrorActionPreference='Stop'
$word=$null
try {
  $word=New-Object -ComObject Word.Application
  [pscustomobject]@{
    available=$true
    name=$word.Name
    version=$word.Version
    build=$word.Build
    path=$word.Path
  } | ConvertTo-Json -Compress
} catch {
  [pscustomobject]@{available=$false; error=$_.Exception.Message} | ConvertTo-Json -Compress
} finally {
  if($null -ne $word){$word.Quit()}
}
"""
    try:
        completed = subprocess.run(
            ["powershell.exe", "-NoLogo", "-NoProfile", "-Command", script],
            capture_output=True,
            text=True,
            timeout=30,
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        output = (completed.stdout or completed.stderr).strip().splitlines()
        if output:
            value = json.loads(output[-1])
            if isinstance(value, dict):
                value["exit_code"] = completed.returncode
                return value
        return {"available": False, "exit_code": completed.returncode, "error": "No structured result"}
    except Exception as exc:
        return {"available": False, "error": str(exc)}


def collect_runtime_provenance() -> dict:
    packages = {}
    for name in ("lxml", "openpyxl", "Pillow", "python-docx", "pydantic"):
        try:
            packages[name] = importlib.metadata.version(name)
        except importlib.metadata.PackageNotFoundError:
            packages[name] = None
    poppler = poppler_status()
    return {
        "schema": "agentic-runtime-provenance/v1",
        "captured_utc": datetime.now(timezone.utc).isoformat(),
        "platform": {
            "system": platform.system(),
            "release": platform.release(),
            "version": platform.version(),
            "machine": platform.machine(),
        },
        "python": {
            "version": sys.version,
            "implementation": platform.python_implementation(),
            "executable": str(Path(sys.executable).resolve()),
            "packages": packages,
        },
        "microsoft_word": _word_provenance(),
        "poppler": {
            "available": poppler.get("available", False),
            "pdftoppm": _program(poppler.get("pdftoppm")),
            "pdfinfo": _program(poppler.get("pdfinfo")),
        },
        "reconstruction_note": (
            "This records the authoring environment and locked inputs. Office and PDF packages may contain "
            "timestamps or application-generated identifiers, so reconstruction is evidence-backed rather "
            "than promised to be byte-for-byte identical."
        ),
    }
