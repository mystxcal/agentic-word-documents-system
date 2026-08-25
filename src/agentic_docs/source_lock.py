from __future__ import annotations

import hashlib
import json
import os
import shutil
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .errors import IntegrityError
from .resolver import file_hash


LOCK_SCHEMA = "agentic-source-lock/v1"
VAULT_RELATIVE = Path(".objects") / "source-lock" / "v1" / "sha256"


def _stable_hash(value: object) -> str:
    encoded = json.dumps(
        value,
        sort_keys=True,
        separators=(",", ":"),
        ensure_ascii=False,
    ).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest().upper()


def _load_object(path: Path, label: str) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
        raise IntegrityError(f"Could not read {label} {path}: {exc}") from exc
    if not isinstance(value, dict):
        raise IntegrityError(f"{label.capitalize()} is not a JSON object: {path}")
    return value


def _atomic_json(path: Path, value: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.{uuid.uuid4().hex}.tmp")
    try:
        temporary.write_text(json.dumps(value, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def _verified_source(path: Path, expected_hash: str, key: str) -> tuple[str, int]:
    if not path.is_file():
        raise IntegrityError(f"Release source-lock input {key!r} is missing: {path}")
    actual = file_hash(path)
    if actual != expected_hash.upper():
        raise IntegrityError(
            f"Release source-lock input {key!r} changed after the candidate build: {path}"
        )
    return actual, path.stat().st_size


def _store_object(system_root: Path, source: Path, sha256: str) -> tuple[Path, bool]:
    vault = system_root / VAULT_RELATIVE
    target = vault / sha256[:2].lower() / sha256.lower()
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.is_file():
        if file_hash(target) == sha256:
            return target, False
        quarantine = target.with_name(
            f"{target.name}.corrupt-{datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%SZ')}-{uuid.uuid4().hex[:8]}"
        )
        os.replace(target, quarantine)
    elif target.exists():
        raise IntegrityError(f"Source-lock object path is not a file: {target}")

    temporary = target.with_name(f".{target.name}.{uuid.uuid4().hex}.tmp")
    try:
        shutil.copy2(source, temporary)
        if file_hash(temporary) != sha256:
            raise IntegrityError(f"Source-lock object failed staging verification: {source}")
        os.replace(temporary, target)
        if file_hash(target) != sha256:
            raise IntegrityError(f"Source-lock object failed commit verification: {target}")
    finally:
        temporary.unlink(missing_ok=True)
    return target, True


def capture_source_lock(
    system_root: Path,
    resolved_inputs_path: Path,
    lock_manifest_path: Path,
) -> dict[str, Any]:
    """Freeze every build input into the deduplicated release source vault."""

    system_root = Path(system_root).resolve()
    resolved_inputs = _load_object(Path(resolved_inputs_path), "resolved input record")
    snapshot = resolved_inputs.get("input_snapshot")
    if not isinstance(snapshot, dict):
        raise IntegrityError(f"Resolved input record has no input_snapshot: {resolved_inputs_path}")

    candidates: list[dict[str, Any]] = []
    for record in snapshot.get("records") or []:
        if not isinstance(record, dict) or not record.get("key") or not record.get("path") or not record.get("sha256"):
            raise IntegrityError(f"Resolved input record contains an incomplete source-lock item: {record!r}")
        candidates.append(
            {
                "key": str(record["key"]),
                "category": str(record.get("category") or "input"),
                "path": Path(record["path"]).resolve(),
                "sha256": str(record["sha256"]).upper(),
            }
        )
    engine = snapshot.get("engine") if isinstance(snapshot.get("engine"), dict) else {}
    for record in engine.get("files") or []:
        if not isinstance(record, dict) or not record.get("path") or not record.get("sha256"):
            raise IntegrityError(f"Engine signature contains an incomplete source-lock item: {record!r}")
        relative = Path(record["path"])
        candidate = (system_root / relative).resolve()
        try:
            candidate.relative_to(system_root)
        except ValueError as exc:
            raise IntegrityError(f"Engine source-lock path escapes the system root: {relative}") from exc
        candidates.append(
            {
                "key": f"engine.{relative.as_posix()}",
                "category": "compiler",
                "path": candidate,
                "sha256": str(record["sha256"]).upper(),
            }
        )

    entries = []
    stored_objects: set[str] = set()
    new_objects: set[str] = set()
    stored_bytes = 0
    for candidate in sorted(candidates, key=lambda item: item["key"]):
        actual, size = _verified_source(candidate["path"], candidate["sha256"], candidate["key"])
        object_path, created = _store_object(system_root, candidate["path"], actual)
        object_key = object_path.relative_to(system_root).as_posix()
        stored_objects.add(actual)
        if created:
            new_objects.add(actual)
            stored_bytes += size
        entries.append(
            {
                "key": candidate["key"],
                "category": candidate["category"],
                "original_path": str(candidate["path"]),
                "sha256": actual,
                "size_bytes": size,
                "object": object_key,
            }
        )

    identity = [
        {key: item[key] for key in ("key", "category", "sha256", "size_bytes", "object")}
        for item in entries
    ]
    manifest = {
        "schema": LOCK_SCHEMA,
        "created_utc": datetime.now(timezone.utc).isoformat(),
        "input_fingerprint": snapshot.get("fingerprint"),
        "engine_sha256": engine.get("sha256"),
        "lock_fingerprint": _stable_hash(identity),
        "vault": VAULT_RELATIVE.as_posix(),
        "entries": entries,
    }
    _atomic_json(Path(lock_manifest_path), manifest)
    verification = verify_source_lock(system_root, lock_manifest_path)
    if not verification["valid"]:
        raise IntegrityError(
            f"New source lock failed verification: {', '.join(verification['issues'])}"
        )
    return {
        "schema": LOCK_SCHEMA,
        "manifest": str(Path(lock_manifest_path)),
        "lock_fingerprint": manifest["lock_fingerprint"],
        "input_fingerprint": manifest["input_fingerprint"],
        "entry_count": len(entries),
        "object_count": len(stored_objects),
        "new_object_count": len(new_objects),
        "new_object_bytes": stored_bytes,
        "verified": True,
    }


def verify_source_lock(system_root: Path, lock_manifest_path: Path) -> dict[str, Any]:
    system_root = Path(system_root).resolve()
    path = Path(lock_manifest_path).resolve()
    manifest = _load_object(path, "source lock")
    issues: list[str] = []
    if manifest.get("schema") != LOCK_SCHEMA:
        issues.append(f"unsupported schema {manifest.get('schema')!r}")
    entries = manifest.get("entries")
    if not isinstance(entries, list):
        entries = []
        issues.append("entries is not an array")
    identity = []
    verified_objects: set[str] = set()
    for item in entries:
        if not isinstance(item, dict):
            issues.append("entry is not an object")
            continue
        required = ("key", "category", "sha256", "size_bytes", "object")
        if any(key not in item for key in required):
            issues.append(f"entry is incomplete: {item!r}")
            continue
        identity.append({key: item[key] for key in required})
        object_path = (system_root / str(item["object"])).resolve()
        try:
            object_path.relative_to(system_root / VAULT_RELATIVE)
        except ValueError:
            issues.append(f"object escapes the source vault: {item['object']}")
            continue
        if str(item["sha256"]) in verified_objects:
            continue
        if not object_path.is_file():
            issues.append(f"object is missing: {item['object']}")
            continue
        if object_path.stat().st_size != int(item["size_bytes"]):
            issues.append(f"object size differs: {item['object']}")
            continue
        if file_hash(object_path) != str(item["sha256"]).upper():
            issues.append(f"object hash differs: {item['object']}")
            continue
        verified_objects.add(str(item["sha256"]))
    expected_fingerprint = _stable_hash(identity)
    if manifest.get("lock_fingerprint") != expected_fingerprint:
        issues.append("lock fingerprint differs")
    return {
        "schema": "agentic-source-lock-verification/v1",
        "manifest": str(path),
        "valid": not issues,
        "entry_count": len(entries),
        "verified_object_count": len(verified_objects),
        "lock_fingerprint": manifest.get("lock_fingerprint"),
        "issues": issues,
    }
