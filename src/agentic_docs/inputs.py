from __future__ import annotations

import hashlib
import json
from functools import lru_cache
from pathlib import Path

from .model import ResolvedDocument
from .resolver import file_hash


def _stable_hash(value: object) -> str:
    encoded = json.dumps(
        value,
        sort_keys=True,
        separators=(",", ":"),
        ensure_ascii=False,
        default=str,
    ).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest().upper()


def _record(key: str, category: str, path: Path) -> dict:
    value = Path(path).resolve()
    return {
        "key": key,
        "category": category,
        "path": str(value),
        "sha256": file_hash(value),
        "size_bytes": value.stat().st_size,
    }


@lru_cache(maxsize=4)
def engine_signature(system_root_text: str) -> dict:
    root = Path(system_root_text).resolve()
    files = sorted((root / "src" / "agentic_docs").rglob("*.py")) + sorted((root / "scripts").glob("*.ps1"))
    records = [
        {
            "path": path.relative_to(root).as_posix(),
            "sha256": file_hash(path),
        }
        for path in files
        if path.is_file()
    ]
    return {
        "sha256": _stable_hash(records),
        "file_count": len(records),
        "files": records,
    }


def resolved_input_snapshot(resolved: ResolvedDocument) -> dict:
    records = [
        _record("manifest.document", "configuration", resolved.manifest_path),
        _record("manifest.kit", "configuration", resolved.kit_path),
        _record("manifest.profile", "configuration", resolved.profile_path),
        _record("manifest.project", "configuration", resolved.project_path),
        _record("profile.shell", "presentation", resolved.shell_path),
    ]
    for key, path in sorted(resolved.presentation_paths.items()):
        if path is not None:
            records.append(_record(f"presentation.{key}", "presentation", path))
    for component_id, component in sorted(resolved.components.items()):
        if component.source_path is not None:
            records.append(_record(f"component.{component_id}", "content", component.source_path))
        for name, path in sorted(getattr(component, "related_paths", {}).items()):
            records.append(_record(f"component.{component_id}.{name}", "content", path))

    engine = engine_signature(str(resolved.system_root))
    identity = {
        "records": [
            {
                "key": item["key"],
                "category": item["category"],
                "sha256": item["sha256"],
            }
            for item in records
        ],
        "engine_sha256": engine["sha256"],
    }
    return {
        "schema": "agentic-input-snapshot/v1",
        "fingerprint": _stable_hash(identity),
        "records": records,
        "engine": engine,
    }


def compare_input_snapshots(current: dict, previous: dict | None) -> dict:
    if previous is None:
        return {
            "schema": "agentic-input-change-summary/v1",
            "baseline_available": False,
            "changed": True,
            "reason": "first_recorded_build",
            "items": [],
            "counts": {"added": 0, "removed": 0, "changed": 0},
        }
    before = {item["key"]: item for item in previous.get("records") or []}
    after = {item["key"]: item for item in current.get("records") or []}
    items = []
    for key in sorted(set(before) | set(after)):
        old = before.get(key)
        new = after.get(key)
        if old is None:
            change = "added"
        elif new is None:
            change = "removed"
        elif old.get("sha256") != new.get("sha256"):
            change = "changed"
        else:
            continue
        sample = new or old
        items.append(
            {
                "key": key,
                "category": sample.get("category"),
                "change": change,
                "before_sha256": old.get("sha256") if old else None,
                "after_sha256": new.get("sha256") if new else None,
                "path": sample.get("path"),
            }
        )
    if previous.get("engine", {}).get("sha256") != current.get("engine", {}).get("sha256"):
        items.append(
            {
                "key": "engine",
                "category": "engine",
                "change": "changed",
                "before_sha256": previous.get("engine", {}).get("sha256"),
                "after_sha256": current.get("engine", {}).get("sha256"),
                "path": None,
            }
        )
    counts = {
        name: sum(item["change"] == name for item in items)
        for name in ("added", "removed", "changed")
    }
    return {
        "schema": "agentic-input-change-summary/v1",
        "baseline_available": True,
        "baseline_fingerprint": previous.get("fingerprint"),
        "current_fingerprint": current.get("fingerprint"),
        "changed": bool(items),
        "reason": "inputs_changed" if items else "identical_inputs",
        "items": items,
        "counts": counts,
    }


def load_current_input_snapshot(resolved: ResolvedDocument) -> dict | None:
    report_path = resolved.system_root / "current" / resolved.manifest.id / "build-report.json"
    if not report_path.is_file():
        return None
    try:
        report = json.loads(report_path.read_text(encoding="utf-8"))
        if report.get("resolved", {}).get("document", {}).get("id") != resolved.manifest.id:
            return None
        run_directory = Path(report["run_directory"])
        payload = json.loads((run_directory / "resolved-inputs.json").read_text(encoding="utf-8"))
        snapshot = payload.get("input_snapshot")
        return snapshot if isinstance(snapshot, dict) else None
    except (OSError, KeyError, TypeError, json.JSONDecodeError):
        return None
