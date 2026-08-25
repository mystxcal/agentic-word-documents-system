from __future__ import annotations

import json
import os
import shutil
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Literal

from .errors import IntegrityError, PublishError
from .resolver import file_hash


PublishPolicy = Literal["fail", "versioned", "replace"]


def require_delivery_alignment(audit: dict[str, Any]) -> None:
    """Refuse an ambiguous default delivery while preserving an explicit override."""

    problems = []
    if not audit.get("immutable_integrity", audit.get("artifact_integrity", False)):
        problems.append("the recorded immutable pair is not verified")
    if not audit.get("current_integrity", audit.get("artifact_integrity", False)):
        problems.append("the visible current Word/PDF pair differs from the recorded build")
    if not audit.get("sources_current", False):
        problems.append("canonical inputs have changed since the recorded build")
    if problems:
        raise IntegrityError(
            "Delivery refused because the document is out of sync: "
            + "; ".join(problems)
            + ". Run 'check' and build, adopt, or restore the intended state first. "
            "Use --allow-out-of-sync only when you deliberately want the older immutable build referenced by current."
        )


def load_immutable_build_report(
    system_root: Path,
    document_id: str,
    build_id: str,
) -> tuple[dict[str, Any], Path]:
    """Load one explicitly selected build report without allowing path escape."""

    if (
        not build_id
        or build_id in {".", ".."}
        or Path(build_id).name != build_id
        or "/" in build_id
        or "\\" in build_id
    ):
        raise IntegrityError(f"Invalid build id: {build_id!r}")
    builds_root = (Path(system_root) / "builds" / document_id).resolve()
    report_path = (builds_root / build_id / "build-report.json").resolve()
    try:
        report_path.relative_to(builds_root)
    except ValueError as exc:
        raise IntegrityError(f"Build id escapes the document build history: {build_id!r}") from exc
    if not report_path.is_file():
        raise IntegrityError(f"Selected immutable build report does not exist: {report_path}")
    try:
        report = json.loads(report_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise IntegrityError(f"Could not read selected build report {report_path}: {exc}") from exc
    if not isinstance(report, dict) or str(report.get("build_id")) != build_id:
        raise IntegrityError(f"Selected build report identity does not match {build_id!r}: {report_path}")
    return report, report_path


def _matches_hash(path: Path, expected_hash: str) -> bool:
    if not path.is_file():
        return False
    try:
        return file_hash(path).upper() == expected_hash.upper()
    except OSError as exc:
        raise PublishError(f"Could not read existing publish target {path}: {exc}") from exc


def verified_build_artifacts(report: dict[str, Any], *, require_pdf: bool = True) -> list[dict[str, str]]:
    if report.get("mode") in {
        "component-preview",
        "page-furniture-preview",
        "lightweight-preview",
    } or report.get("content_scope") not in {None, "complete"}:
        raise IntegrityError(
            "A scoped preview is diagnostic and cannot be delivered as a complete document"
        )
    quality = report.get("quality")
    if isinstance(quality, dict) and not quality.get(
        "release_ready",
        quality.get("passed", False),
    ):
        raise IntegrityError(
            "The selected build does not have a complete passing quality proof and cannot be delivered"
        )
    artifacts = report.get("artifacts")
    if not isinstance(artifacts, dict):
        raise IntegrityError("Build report does not contain an artifacts object")
    result: list[dict[str, str]] = []
    for kind in ("docx", "pdf"):
        path_value = artifacts.get(kind)
        expected = artifacts.get(f"{kind}_sha256")
        if not path_value:
            if kind == "pdf" and require_pdf:
                raise IntegrityError("The selected build has no PDF artifact")
            continue
        path = Path(path_value)
        if not path.is_file():
            raise IntegrityError(f"Immutable build {kind.upper()} is missing: {path}")
        try:
            actual = file_hash(path)
        except OSError as exc:
            raise IntegrityError(f"Immutable build {kind.upper()} cannot be read: {path}: {exc}") from exc
        if not expected or actual.upper() != str(expected).upper():
            raise IntegrityError(
                f"Immutable build {kind.upper()} does not match its recorded SHA-256: {path}"
            )
        result.append({"kind": kind, "path": str(path), "sha256": actual})
    if not result:
        raise IntegrityError("The build report contains no publishable artifacts")
    return result


def _versioned_target(destination: Path, build_id: str, expected_hash: str) -> tuple[Path, bool]:
    if not destination.exists():
        return destination, False
    if _matches_hash(destination, expected_hash):
        return destination, True
    base = destination.with_name(f"{destination.stem} [build {build_id}]{destination.suffix}")
    if not base.exists():
        return base, False
    if _matches_hash(base, expected_hash):
        return base, True
    counter = 2
    while True:
        candidate = destination.with_name(
            f"{destination.stem} [build {build_id}-{counter}]{destination.suffix}"
        )
        if not candidate.exists():
            return candidate, False
        if _matches_hash(candidate, expected_hash):
            return candidate, True
        counter += 1


def publish_build(
    report: dict[str, Any],
    destination: Path,
    *,
    policy: PublishPolicy = "versioned",
) -> dict[str, Any]:
    """Stage, hash-check, and commit one immutable DOCX/PDF build pair."""

    if policy not in {"fail", "versioned", "replace"}:
        raise PublishError(f"Unsupported publish policy: {policy}")
    build_id = str(report.get("build_id") or "unknown")
    artifacts = verified_build_artifacts(report)
    destination = Path(destination).expanduser().resolve()
    anchor = Path(destination.anchor) if destination.anchor else destination.parent
    if destination.anchor and not anchor.exists():
        raise PublishError(f"Destination drive or root is unavailable: {anchor}")
    try:
        destination.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise PublishError(f"Could not create or access publish destination {destination}: {exc}") from exc
    if not destination.is_dir():
        raise PublishError(f"Publish destination is not a directory: {destination}")

    planned: list[dict[str, Any]] = []
    for artifact in artifacts:
        source = Path(artifact["path"])
        nominal = destination / source.name
        if policy == "versioned":
            target, already_present = _versioned_target(nominal, build_id, artifact["sha256"])
        else:
            target = nominal
            already_present = _matches_hash(target, artifact["sha256"])
            if target.exists() and not target.is_file():
                raise PublishError(f"Publish target exists but is not a file: {target}")
            if policy == "fail" and target.exists() and not already_present:
                raise PublishError(
                    f"Publish target already exists with different content: {target}. "
                    "Use --policy versioned or --policy replace."
                )
        planned.append(
            {
                **artifact,
                "source_path": source,
                "target_path": target,
                "already_present": already_present,
            }
        )

    if all(item["already_present"] for item in planned):
        return {
            "schema": "agentic-publish-report/v2",
            "build_id": build_id,
            "destination": str(destination),
            "policy": policy,
            "changed": False,
            "artifacts": [
                {
                    "kind": item["kind"],
                    "path": str(item["target_path"]),
                    "sha256": item["sha256"],
                    "status": "already-present",
                }
                for item in planned
            ],
        }

    stage_directory = Path(tempfile.mkdtemp(prefix=".agentic-publish-", dir=destination))
    history_directory = destination / ".agentic-docs-history" / (
        datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ") + f"-{build_id}"
    )
    staged: dict[Path, Path] = {}
    backups: dict[Path, Path] = {}
    committed: list[Path] = []
    try:
        for item in planned:
            if item["already_present"]:
                continue
            stage = stage_directory / item["target_path"].name
            shutil.copy2(item["source_path"], stage)
            staged_hash = file_hash(stage)
            if staged_hash != item["sha256"]:
                raise PublishError(
                    f"Hash verification failed while staging {item['source_path']} to {stage}"
                )
            staged[item["target_path"]] = stage

        if policy == "replace":
            existing = [target for target in staged if target.exists()]
            if existing:
                history_directory.mkdir(parents=True, exist_ok=False)
                for target in existing:
                    backup = history_directory / target.name
                    os.replace(target, backup)
                    backups[target] = backup

        for target, stage in staged.items():
            os.replace(stage, target)
            committed.append(target)
            expected = next(item["sha256"] for item in planned if item["target_path"] == target)
            if file_hash(target) != expected:
                raise PublishError(f"Destination hash verification failed after commit: {target}")
    except Exception as exc:
        for target in reversed(committed):
            try:
                target.unlink()
            except OSError:
                pass
        for target, backup in backups.items():
            if backup.exists():
                try:
                    os.replace(backup, target)
                except OSError:
                    pass
        if isinstance(exc, PublishError):
            raise
        raise PublishError(f"Publish transaction failed: {exc}") from exc
    finally:
        shutil.rmtree(stage_directory, ignore_errors=True)

    return {
        "schema": "agentic-publish-report/v2",
        "build_id": build_id,
        "destination": str(destination),
        "policy": policy,
        "changed": True,
        "history": str(history_directory) if backups else None,
        "artifacts": [
            {
                "kind": item["kind"],
                "path": str(item["target_path"]),
                "sha256": item["sha256"],
                "status": "already-present" if item["already_present"] else "published",
            }
            for item in planned
        ],
    }
