from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from .model import ResolvedDocument
from .resolver import file_hash
from .source_lock import verify_source_lock


def _issue(
    code: str,
    severity: str,
    message: str,
    *,
    path: Path | str | None = None,
    hint: str | None = None,
) -> dict[str, Any]:
    return {
        "code": code,
        "severity": severity,
        "message": message,
        "path": str(path) if path is not None else None,
        "hint": hint,
    }


def _hash_record(path: Path | None, expected: str | None = None) -> dict[str, Any]:
    if path is None:
        return {"path": None, "exists": False, "sha256": None, "expected_sha256": expected, "matches": False}
    exists = path.is_file()
    actual = None
    error = None
    if exists:
        try:
            actual = file_hash(path)
        except OSError as exc:
            error = str(exc)
    return {
        "path": str(path),
        "exists": exists,
        "sha256": actual,
        "expected_sha256": expected,
        "matches": bool(expected and actual and expected.upper() == actual.upper()),
        "error": error,
    }


def _load_report(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"Could not read build report {path}: {exc}") from exc
    if not isinstance(value, dict):
        raise ValueError(f"Build report is not a JSON object: {path}")
    return value


def _audit_releases(system_root: Path, document_id: str, issues: list[dict[str, Any]]) -> dict[str, Any]:
    release_root = system_root / "releases" / document_id
    records = []
    if not release_root.is_dir():
        return {"count": 0, "valid": True, "releases": []}
    for report_path in sorted(release_root.rglob("release-report.json")):
        try:
            report = _load_report(report_path)
        except ValueError as exc:
            issues.append(_issue("RELEASE_REPORT_INVALID", "error", str(exc), path=report_path))
            records.append({"report": str(report_path), "valid": False, "issues": [str(exc)]})
            continue
        release_artifacts = report.get("release_artifacts") if isinstance(report.get("release_artifacts"), dict) else {}
        build_artifacts = report.get("artifacts") if isinstance(report.get("artifacts"), dict) else {}
        release_issues = []
        artifact_records = {}
        for kind in ("docx", "pdf"):
            path_value = release_artifacts.get(kind)
            expected = build_artifacts.get(f"{kind}_sha256")
            record = _hash_record(Path(path_value) if path_value else None, expected)
            artifact_records[kind] = record
            if not record["exists"]:
                message = f"Controlled release {kind.upper()} is missing"
            elif not record["matches"]:
                message = f"Controlled release {kind.upper()} no longer matches its recorded hash"
            else:
                continue
            release_issues.append(message)
            issues.append(
                _issue(
                    "RELEASE_ARTIFACT_DRIFT",
                    "error",
                    message + f" for build {report.get('build_id')}",
                    path=record["path"],
                )
            )

        lock_path = report_path.parent / "source-lock.json"
        source_lock_report = report.get("source_lock")
        if source_lock_report and lock_path.is_file():
            try:
                lock_verification = verify_source_lock(system_root, lock_path)
            except Exception as exc:
                lock_verification = {"valid": False, "issues": [str(exc)], "manifest": str(lock_path)}
            if not lock_verification["valid"]:
                message = "Controlled release source lock is invalid"
                release_issues.append(message)
                issues.append(
                    _issue(
                        "RELEASE_SOURCE_LOCK_INVALID",
                        "error",
                        message + f" for build {report.get('build_id')}",
                        path=lock_path,
                        hint="Restore the recorded source-vault object from backup; do not recreate history from current sources.",
                    )
                )
        else:
            lock_verification = {"valid": False, "legacy": True, "issues": ["source lock is absent"]}
            issues.append(
                _issue(
                    "RELEASE_SOURCE_LOCK_MISSING",
                    "warning",
                    f"Controlled release {report.get('build_id')} predates source locking or has no lock manifest.",
                    path=report_path.parent,
                )
            )
        runtime = report.get("runtime_provenance") if isinstance(report.get("runtime_provenance"), dict) else {}
        runtime_path = report_path.parent / "runtime-provenance.json"
        if runtime:
            runtime_record = _hash_record(runtime_path, runtime.get("sha256"))
            if not runtime_record["matches"]:
                message = "Controlled release runtime provenance is missing or has changed"
                release_issues.append(message)
                issues.append(
                    _issue(
                        "RELEASE_RUNTIME_PROVENANCE_INVALID",
                        "error",
                        message + f" for build {report.get('build_id')}",
                        path=runtime_path,
                    )
                )
        else:
            runtime_record = {"path": str(runtime_path), "legacy": True, "matches": False}
        records.append(
            {
                "build_id": report.get("build_id"),
                "revision": report_path.parent.parent.name if report_path.parent.parent != release_root else None,
                "directory": str(report_path.parent),
                "artifacts": artifact_records,
                "source_lock": lock_verification,
                "runtime_provenance": runtime_record,
                "valid": not release_issues,
                "issues": release_issues,
            }
        )
    return {
        "count": len(records),
        "valid": all(item["valid"] for item in records),
        "releases": records,
    }


def audit_document(resolved: ResolvedDocument) -> dict[str, Any]:
    """Verify canonical inputs, immutable build artifacts, and the mutable current mirror."""

    current_root = resolved.system_root / "current" / resolved.manifest.id
    report_path = current_root / "build-report.json"
    issues: list[dict[str, Any]] = []
    if not report_path.is_file():
        issues.append(
            _issue(
                "CURRENT_BUILD_REPORT_MISSING",
                "error",
                "No current build report exists for this document.",
                path=report_path,
                hint="Run the build command before publishing or treating current artifacts as authoritative.",
            )
        )
        return {
            "schema": "agentic-integrity-report/v2",
            "document_id": resolved.manifest.id,
            "build_id": None,
            "ready": False,
            "artifact_integrity": False,
            "sources_current": False,
            "issues": issues,
            "current_root": str(current_root),
        }

    try:
        report = _load_report(report_path)
    except ValueError as exc:
        issues.append(_issue("CURRENT_BUILD_REPORT_INVALID", "error", str(exc), path=report_path))
        return {
            "schema": "agentic-integrity-report/v2",
            "document_id": resolved.manifest.id,
            "build_id": None,
            "ready": False,
            "artifact_integrity": False,
            "sources_current": False,
            "issues": issues,
            "current_root": str(current_root),
        }

    build_id = report.get("build_id")
    artifacts = report.get("artifacts") if isinstance(report.get("artifacts"), dict) else {}
    expected_docx_hash = artifacts.get("docx_sha256")
    expected_pdf_hash = artifacts.get("pdf_sha256")
    immutable_docx_path = Path(artifacts["docx"]) if artifacts.get("docx") else None
    immutable_pdf_path = Path(artifacts["pdf"]) if artifacts.get("pdf") else None
    immutable_docx = _hash_record(immutable_docx_path, expected_docx_hash)
    immutable_pdf = _hash_record(immutable_pdf_path, expected_pdf_hash)

    expected_docx_path = current_root / f"{resolved.manifest.outputs.basename}.docx"
    expected_pdf_path = current_root / f"{resolved.manifest.outputs.basename}.pdf"
    current_docx = _hash_record(expected_docx_path, expected_docx_hash)
    current_pdf = _hash_record(expected_pdf_path, expected_pdf_hash)

    for label, record in (
        ("immutable DOCX", immutable_docx),
        ("immutable PDF", immutable_pdf),
        ("current DOCX", current_docx),
        ("current PDF", current_pdf),
    ):
        if record["expected_sha256"] is None and label.endswith("PDF"):
            continue
        if not record["exists"]:
            issues.append(
                _issue(
                    "ARTIFACT_MISSING",
                    "error",
                    f"The {label} is missing.",
                    path=record["path"],
                )
            )
        elif record.get("error"):
            issues.append(
                _issue(
                    "ARTIFACT_UNREADABLE",
                    "error",
                    f"The {label} could not be hashed: {record['error']}",
                    path=record["path"],
                )
            )
        elif not record["matches"]:
            issues.append(
                _issue(
                    "ARTIFACT_DRIFT",
                    "error",
                    f"The {label} no longer matches build {build_id}.",
                    path=record["path"],
                    hint="Adopt the edited workcopy or restore current from the immutable build; do not mix it with the other artifact.",
                )
            )

    built_components = {
        item.get("id"): item
        for item in report.get("resolved", {}).get("components", [])
        if isinstance(item, dict) and item.get("id")
    }
    built_resolved = report.get("resolved") if isinstance(report.get("resolved"), dict) else {}
    input_records: list[dict[str, Any]] = []

    def record_input(category: str, name: str, path: Path | None, built_hash: str | None) -> None:
        current_hash = None
        if path is not None and path.is_file():
            try:
                current_hash = file_hash(path)
            except OSError:
                current_hash = None
        matches = current_hash == built_hash if built_hash is not None else current_hash is None
        record = {
            "category": category,
            "name": name,
            "path": str(path) if path is not None else None,
            "current_sha256": current_hash,
            "built_sha256": built_hash,
            "matches": matches,
        }
        input_records.append(record)
        if not matches:
            issues.append(
                _issue(
                    "CANONICAL_INPUT_CHANGED",
                    "warning",
                    f"Canonical {category} input {name!r} has changed since build {build_id}.",
                    path=path,
                    hint="Run build to make the current Word/PDF pair reflect all canonical inputs.",
                )
            )

    built_layers = built_resolved.get("layers") if isinstance(built_resolved.get("layers"), dict) else {}
    layer_paths = {
        "kit": getattr(resolved, "kit_path", None),
        "profile": getattr(resolved, "profile_path", None),
        "project": getattr(resolved, "project_path", None),
        "document": getattr(resolved, "manifest_path", None),
        "shell": getattr(resolved, "shell_path", None),
    }
    for name, path in layer_paths.items():
        built = built_layers.get(name) if isinstance(built_layers.get(name), dict) else {}
        if path is not None or built:
            record_input("layer", name, Path(path) if path is not None else None, built.get("sha256"))

    built_presentation = (
        built_resolved.get("presentation")
        if isinstance(built_resolved.get("presentation"), dict)
        else {}
    )
    presentation_paths = getattr(resolved, "presentation_paths", {}) or {}
    for name in sorted(set(built_presentation) | set(presentation_paths)):
        built = built_presentation.get(name) if isinstance(built_presentation.get(name), dict) else {}
        path = presentation_paths.get(name)
        if path is not None or built:
            record_input(
                "presentation",
                name,
                Path(path) if path is not None else None,
                built.get("sha256"),
            )

    source_records: list[dict[str, Any]] = []
    for component_id, component in resolved.components.items():
        built = built_components.get(component_id, {})
        built_hash = built.get("source_hash")
        record = {
            "component": component_id,
            "path": str(component.source_path) if component.source_path else None,
            "current_sha256": component.source_hash,
            "built_sha256": built_hash,
            "matches": component.source_hash == built_hash,
        }
        source_records.append(record)
        if built_hash is not None and component.source_hash != built_hash:
            issues.append(
                _issue(
                    "CANONICAL_SOURCE_CHANGED",
                    "warning",
                    f"Canonical component {component_id!r} has changed since build {build_id}.",
                    path=component.source_path,
                    hint="Run build to make the current Word/PDF pair reflect the canonical source.",
                )
            )

    expected_names = {
        "build-report.json",
        expected_docx_path.name,
        expected_pdf_path.name,
    }
    stale_files = []
    if current_root.is_dir():
        stale_files = sorted(
            str(path)
            for path in current_root.iterdir()
            if path.is_file() and path.name not in expected_names
        )
    if stale_files:
        issues.append(
            _issue(
                "CURRENT_FOLDER_STALE_FILES",
                "warning",
                f"The current folder contains {len(stale_files)} stale or manually created file(s).",
                path=current_root,
                hint="Use restore-current --clean to retain them in history and restore one verified pair.",
            )
        )

    release_integrity = _audit_releases(resolved.system_root, resolved.manifest.id, issues)

    immutable_integrity = immutable_docx["matches"] and all(
        record["expected_sha256"] is None or record["matches"]
        for record in (immutable_pdf,)
    )
    current_integrity = current_docx["matches"] and all(
        record["expected_sha256"] is None or record["matches"]
        for record in (current_pdf,)
    )
    artifact_integrity = immutable_integrity and current_integrity
    sources_current = all(record["matches"] for record in [*input_records, *source_records])
    quality = report.get("quality") if isinstance(report.get("quality"), dict) else {}
    build_quality_passed = quality.get("passed") is True
    if not build_quality_passed:
        verification_mode = str(report.get("verification_mode") or "unknown")
        issues.append(
            _issue(
                "BUILD_QUALITY_NOT_FULLY_PROVED",
                "warning",
                f"Build {build_id} does not have a complete passing quality proof (mode: {verification_mode}).",
                path=report_path,
                hint="Run a normal full build before formal review or release.",
            )
        )
    ready = artifact_integrity and sources_current and release_integrity["valid"] and not any(
        item["severity"] == "error" for item in issues
    )
    return {
        "schema": "agentic-integrity-report/v2",
        "document_id": resolved.manifest.id,
        "build_id": build_id,
        "ready": ready,
        "artifact_integrity": artifact_integrity,
        "immutable_integrity": immutable_integrity,
        "current_integrity": current_integrity,
        "sources_current": sources_current,
        "build_quality_passed": build_quality_passed,
        "verification_mode": report.get("verification_mode") or "unknown",
        "release_integrity": release_integrity,
        "current_root": str(current_root),
        "report": str(report_path),
        "immutable": {"docx": immutable_docx, "pdf": immutable_pdf},
        "current": {"docx": current_docx, "pdf": current_pdf},
        "sources": source_records,
        "inputs": input_records,
        "stale_files": stale_files,
        "issues": issues,
    }
