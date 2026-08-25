from __future__ import annotations

import json
import os
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .errors import IntegrityError


def _load_json(path: Path) -> dict[str, Any] | None:
    try:
        value = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError):
        return None
    return value if isinstance(value, dict) else None


def _directory_size(path: Path) -> int:
    return sum(item.stat().st_size for item in path.rglob("*") if item.is_file())


def _current_build_id(system_root: Path, document_id: str) -> str | None:
    report = _load_json(system_root / "current" / document_id / "build-report.json")
    return str(report.get("build_id")) if report and report.get("build_id") else None


def _released_build_ids(system_root: Path, document_id: str) -> set[str]:
    result: set[str] = set()
    root = system_root / "releases" / document_id
    if not root.is_dir():
        return result
    for path in root.rglob("release-report.json"):
        report = _load_json(path)
        if report and report.get("build_id"):
            result.add(str(report["build_id"]))
    return result


def _atomic_json(path: Path, value: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.{uuid.uuid4().hex}.tmp")
    try:
        temporary.write_text(json.dumps(value, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def manage_retention(
    system_root: Path,
    document_id: str,
    *,
    keep_drafts: int = 10,
    keep_previews: int = 3,
    apply: bool = False,
) -> dict[str, Any]:
    """Plan or recoverably retire old draft builds; releases are never candidates."""

    if keep_drafts < 1:
        raise IntegrityError("Retention must keep at least one complete draft build")
    if keep_previews < 0:
        raise IntegrityError("Retention preview count cannot be negative")
    system_root = Path(system_root).resolve()
    build_root = system_root / "builds" / document_id
    current_id = _current_build_id(system_root, document_id)
    released_ids = _released_build_ids(system_root, document_id)
    records: list[dict[str, Any]] = []
    if build_root.is_dir():
        for directory in sorted(
            (item for item in build_root.iterdir() if item.is_dir() and not item.name.startswith(".")),
            key=lambda item: item.name,
        ):
            report_path = directory / "build-report.json"
            report = _load_json(report_path)
            record = {
                "build_id": directory.name,
                "path": str(directory),
                "size_bytes": _directory_size(directory),
                "mode": report.get("mode") if report else None,
                "report_valid": bool(report and str(report.get("build_id")) == directory.name),
                "protected_reasons": [],
            }
            if not record["report_valid"]:
                record["protected_reasons"].append("missing_or_invalid_report")
            if directory.name == current_id:
                record["protected_reasons"].append("current")
            if directory.name in released_ids:
                record["protected_reasons"].append("released")
            if record["mode"] == "release":
                record["protected_reasons"].append("release_candidate")
            records.append(record)

    preview_modes = {"component-preview", "page-furniture-preview", "lightweight-preview"}
    complete = [
        item
        for item in records
        if item["report_valid"]
        and item["mode"] not in preview_modes | {"release"}
        and "released" not in item["protected_reasons"]
    ]
    previews = [item for item in records if item["report_valid"] and item["mode"] in preview_modes]
    for item in complete[-keep_drafts:]:
        item["protected_reasons"].append("draft_retention_window")
    if keep_previews:
        for item in previews[-keep_previews:]:
            item["protected_reasons"].append("preview_retention_window")

    candidates = [item for item in records if not item["protected_reasons"]]
    protected = [item for item in records if item["protected_reasons"]]
    operation_id = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S_%fZ")
    archive_root = system_root / "archive" / "retired-builds" / document_id / operation_id
    moved: list[tuple[Path, Path]] = []
    if apply and candidates:
        archive_root.mkdir(parents=True, exist_ok=False)
        try:
            for item in candidates:
                source = Path(item["path"]).resolve()
                try:
                    source.relative_to(build_root.resolve())
                except ValueError as exc:
                    raise IntegrityError(f"Retention candidate escapes the build root: {source}") from exc
                destination = archive_root / source.name
                if destination.exists():
                    raise IntegrityError(f"Retention archive target already exists: {destination}")
                os.replace(source, destination)
                moved.append((source, destination))
                item["archived_path"] = str(destination)
        except Exception:
            for source, destination in reversed(moved):
                if destination.exists() and not source.exists():
                    os.replace(destination, source)
            raise

    result = {
        "schema": "agentic-retention-plan/v1",
        "operation_id": operation_id,
        "document_id": document_id,
        "applied": apply,
        "recoverable": True,
        "policy": {"keep_drafts": keep_drafts, "keep_previews": keep_previews},
        "current_build_id": current_id,
        "released_build_ids": sorted(released_ids),
        "build_count": len(records),
        "protected_count": len(protected),
        "candidate_count": len(candidates),
        "candidate_bytes": sum(item["size_bytes"] for item in candidates),
        "archive": str(archive_root) if apply and candidates else None,
        "candidates": candidates,
        "protected": protected,
        "notes": [
            "Controlled releases, release-mode candidates, current builds, and invalid or unknown histories are never retired.",
            "Apply moves candidates into a recoverable archive; it does not delete them or garbage-collect the source vault.",
        ],
    }
    if apply and candidates:
        _atomic_json(archive_root / "retention-receipt.json", result)
    return result
