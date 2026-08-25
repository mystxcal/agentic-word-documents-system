from __future__ import annotations

import json
import os
import shutil
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .errors import IntegrityError
from .publishing import verified_build_artifacts
from .reporting import write_json
from .resolver import file_hash


def _load_previous(path: Path) -> dict[str, Any] | None:
    if not path.is_file():
        return None
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    return value if isinstance(value, dict) else None


def _previous_artifact_hashes(report: dict[str, Any] | None) -> dict[str, str]:
    if not report:
        return {}
    artifacts = report.get("artifacts") if isinstance(report.get("artifacts"), dict) else {}
    result = {}
    for kind in ("docx", "pdf"):
        path_value = artifacts.get(kind)
        hash_value = artifacts.get(f"{kind}_sha256")
        if path_value and hash_value:
            result[Path(path_value).name] = str(hash_value)
        current = report.get("current") if isinstance(report.get("current"), dict) else {}
        current_value = current.get(kind)
        if current_value and hash_value:
            result[Path(current_value).name] = str(hash_value)
    return result


def update_current_from_build(
    report: dict[str, Any],
    current_root: Path,
    *,
    basename: str,
) -> dict[str, Any]:
    """Commit a build pair to current without losing a manually edited workcopy."""

    current_root = Path(current_root)
    current_root.mkdir(parents=True, exist_ok=True)
    artifacts = verified_build_artifacts(report, require_pdf=False)
    if not any(item["kind"] == "docx" for item in artifacts):
        raise IntegrityError("A current update requires a verified DOCX artifact")

    build_id = str(report.get("build_id") or "unknown")
    operation_id = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S_%fZ")
    report_path = current_root / "build-report.json"
    previous = _load_previous(report_path)
    previous_hashes = _previous_artifact_hashes(previous)
    target_by_kind = {
        "docx": current_root / f"{basename}.docx",
        "pdf": current_root / f"{basename}.pdf",
    }
    stage = current_root / f".update-{operation_id}"
    rollback = stage / "rollback"
    stage.mkdir(parents=True, exist_ok=False)
    staged: dict[str, Path] = {}
    try:
        for artifact in artifacts:
            candidate = stage / target_by_kind[artifact["kind"]].name
            shutil.copy2(artifact["path"], candidate)
            if file_hash(candidate) != artifact["sha256"]:
                raise IntegrityError(f"Current-update staging hash mismatch: {artifact['path']}")
            staged[artifact["kind"]] = candidate
    except Exception:
        shutil.rmtree(stage, ignore_errors=True)
        raise

    previous_named_files = {
        current_root / name
        for name in previous_hashes
        if (current_root / name).is_file()
    }
    affected = {
        *previous_named_files,
        *(path for path in target_by_kind.values() if path.is_file()),
    }
    manual_files = []
    generated_files = []
    for path in sorted(affected, key=lambda item: item.name.lower()):
        expected = previous_hashes.get(path.name)
        actual = file_hash(path)
        if expected and actual.upper() == expected.upper():
            generated_files.append(path)
        else:
            manual_files.append(path)

    history = (
        current_root
        / ".history"
        / "build-displaced"
        / f"{operation_id}-{build_id}"
    )
    displaced: list[dict[str, str]] = []
    moved_for_rollback: list[tuple[Path, Path]] = []
    committed: list[Path] = []
    try:
        if manual_files:
            history.mkdir(parents=True, exist_ok=False)
            for path in manual_files:
                retained = history / path.name
                os.replace(path, retained)
                displaced.append({"from": str(path), "to": str(retained)})

        rollback.mkdir(parents=True, exist_ok=True)
        for path in [*generated_files, *([report_path] if report_path.is_file() else [])]:
            target = rollback / path.name
            os.replace(path, target)
            moved_for_rollback.append((path, target))

        for kind, candidate in staged.items():
            target = target_by_kind[kind]
            os.replace(candidate, target)
            committed.append(target)
            expected = next(item["sha256"] for item in artifacts if item["kind"] == kind)
            if file_hash(target) != expected:
                raise IntegrityError(f"Current {kind.upper()} failed final hash verification")

        current_report = report | {
            "current": {
                "docx": str(target_by_kind["docx"]),
                "pdf": str(target_by_kind["pdf"]) if "pdf" in staged else None,
            },
            "current_update": {
                "operation_id": operation_id,
                "manual_files_preserved": displaced,
                "history": str(history) if displaced else None,
            },
        }
        staged_report = stage / report_path.name
        write_json(staged_report, current_report)
        os.replace(staged_report, report_path)
        committed.append(report_path)
    except Exception:
        for path in reversed(committed):
            try:
                path.unlink()
            except OSError:
                pass
        for original, retained in reversed(moved_for_rollback):
            if retained.exists() and not original.exists():
                try:
                    os.replace(retained, original)
                except OSError:
                    pass
        for item in reversed(displaced):
            original = Path(item["from"])
            retained = Path(item["to"])
            if retained.exists() and not original.exists():
                try:
                    os.replace(retained, original)
                except OSError:
                    pass
        raise
    finally:
        shutil.rmtree(stage, ignore_errors=True)

    return {
        "operation_id": operation_id,
        "current_docx": str(target_by_kind["docx"]),
        "current_pdf": str(target_by_kind["pdf"]) if "pdf" in staged else None,
        "manual_files_preserved": displaced,
        "history": str(history) if displaced else None,
    }
