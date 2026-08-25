from __future__ import annotations

import json
import os
import shutil
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .build import build_document
from .diagnostics import DiagnosticBag
from .docx_text import TextReplacement, replace_docx_text, replace_plain_text
from .errors import OperationPartialError, RevisionError
from .model import ComponentType, Ownership, ResolvedComponent, ResolvedDocument
from .publishing import PublishPolicy, publish_build
from .publishing import verified_build_artifacts
from .reporting import write_json
from .resolver import file_hash, resolve_document


def _editable_component(
    resolved: ResolvedDocument,
    component_id: str | None,
) -> ResolvedComponent:
    if component_id is not None:
        item = resolved.components.get(component_id)
        if item is None:
            raise RevisionError(f"Unknown component {component_id!r}")
        candidates = [item]
    else:
        candidates = [
            item
            for item in resolved.components.values()
            if item.declaration.type == ComponentType.DOCUMENT
            and item.source_path is not None
            and (
                (
                    item.declaration.ownership == Ownership.WORD_FRAGMENT
                    and item.source_path.suffix.lower() == ".docx"
                )
                or (
                    item.declaration.ownership == Ownership.SOURCE
                    and item.source_path.suffix.lower() in {".md", ".markdown"}
                )
            )
        ]
        if len(candidates) != 1:
            names = ", ".join(item.id for item in candidates) or "none"
            raise RevisionError(
                "The document does not have exactly one editable Word/Markdown component. "
                f"Use --component. Candidates: {names}"
            )
    item = candidates[0]
    is_word = (
        item.declaration.ownership == Ownership.WORD_FRAGMENT
        and item.source_path is not None
        and item.source_path.suffix.lower() == ".docx"
    )
    is_markdown = (
        item.declaration.ownership == Ownership.SOURCE
        and item.source_path is not None
        and item.source_path.suffix.lower() in {".md", ".markdown"}
    )
    if not (is_word or is_markdown):
        raise RevisionError(
            f"Component {item.id!r} is owned as {item.declaration.ownership}; "
            "revise edits canonical Word fragments or Markdown source components."
        )
    return item


def revise_document(
    resolved: ResolvedDocument,
    diagnostics: DiagnosticBag,
    replacements: list[TextReplacement],
    *,
    component_id: str | None = None,
    expected_total: int | None = None,
    allow_zero: bool = False,
    all_stories: bool = False,
    build: bool = True,
    render_pages: bool = True,
    retain_intermediates: bool = False,
    output_root: Path | None = None,
    publish_destination: Path | None = None,
    publish_policy: PublishPolicy = "versioned",
) -> dict[str, Any]:
    """Revise one canonical Word or Markdown component and rebuild as one transaction."""

    component = _editable_component(resolved, component_id)
    source = component.source_path
    assert source is not None
    original_hash = file_hash(source)
    temporary_root = resolved.manifest_path.parent / ".agentic-tmp"
    temporary_root.mkdir(parents=True, exist_ok=True)
    descriptor, candidate_name = tempfile.mkstemp(
        prefix=f"{component.id}.",
        suffix=f".candidate{source.suffix.lower()}",
        dir=temporary_root,
    )
    os.close(descriptor)
    candidate = Path(candidate_name)
    candidate.unlink()
    history: Path | None = None
    backup: Path | None = None
    try:
        if source.suffix.lower() == ".docx":
            replacement_summary = replace_docx_text(
                source,
                candidate,
                replacements,
                all_stories=all_stories,
            )
        else:
            if all_stories:
                raise RevisionError("--all-stories applies only to canonical Word components")
            replacement_summary = replace_plain_text(source, candidate, replacements)
        total = replacement_summary.total_replacements
        if total == 0 and not allow_zero:
            raise RevisionError(
                f"No text was replaced in canonical component {component.id!r}; source was not changed."
            )
        if expected_total is not None and total != expected_total:
            raise RevisionError(
                f"Expected {expected_total} replacement(s), found {total}; source was not changed."
            )
        if total == 0:
            if publish_destination is not None:
                raise RevisionError("A zero-change revision cannot publish a new build")
            return {
                "schema": "agentic-revision-report/v2",
                "operation_id": None,
                "document_id": resolved.manifest.id,
                "component": component.id,
                "canonical_source": str(source),
                "source_before_sha256": original_hash,
                "source_after_sha256": original_hash,
                "backup": None,
                "changed": False,
                "replacement": replacement_summary.to_dict(),
                "build": {"performed": False},
                "publish": None,
            }
        if file_hash(source) != original_hash:
            raise RevisionError(
                f"Canonical component changed while the revision was being prepared: {source}"
            )

        operation_id = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S_%fZ")
        history = resolved.manifest_path.parent / ".history" / "revisions" / operation_id
        history.mkdir(parents=True, exist_ok=False)
        backup = history / source.name
        shutil.copy2(source, backup)
        os.replace(candidate, source)

        build_report = None
        publish_report = None
        publish_error = None
        if build:
            try:
                refreshed_diagnostics = DiagnosticBag()
                refreshed = resolve_document(resolved.manifest_path, refreshed_diagnostics)
                build_report = build_document(
                    refreshed,
                    refreshed_diagnostics,
                    output_root=output_root,
                    render_pages=render_pages,
                    retain_intermediates=retain_intermediates,
                )
            except Exception as exc:
                restore_candidate = source.with_name(source.name + ".restore.tmp")
                shutil.copy2(backup, restore_candidate)
                os.replace(restore_candidate, source)
                raise RevisionError(
                    f"Build failed after canonical revision; the original source was restored. Cause: {exc}"
                ) from exc
            if publish_destination is not None:
                try:
                    publish_report = publish_build(
                        build_report,
                        publish_destination,
                        policy=publish_policy,
                    )
                except Exception as exc:
                    publish_error = str(exc)
        elif publish_destination is not None:
            raise RevisionError("--publish requires the automatic build; remove --no-build")

        report = {
            "schema": "agentic-revision-report/v2",
            "operation_id": operation_id,
            "document_id": resolved.manifest.id,
            "component": component.id,
            "canonical_source": str(source),
            "source_before_sha256": original_hash,
            "source_after_sha256": file_hash(source),
            "backup": str(backup),
            "changed": True,
            "replacement": replacement_summary.to_dict(),
            "build": {
                "performed": build,
                "build_id": build_report.get("build_id") if build_report else None,
                "run_directory": build_report.get("run_directory") if build_report else None,
                "docx": (
                    build_report.get("current_docx")
                    or (build_report.get("artifacts") or {}).get("docx")
                ) if build_report else None,
                "pdf": (
                    build_report.get("current_pdf")
                    or (build_report.get("artifacts") or {}).get("pdf")
                ) if build_report else None,
                "current_updated": build_report.get("current_updated") if build_report else None,
                "current_update_error": build_report.get("current_update_error") if build_report else None,
                "quality": build_report.get("quality") if build_report else None,
            },
            "publish": publish_report or (
                {"succeeded": False, "destination": str(publish_destination), "error": publish_error}
                if publish_error
                else None
            ),
        }
        write_json(history / "revision-report.json", report)
        if publish_error:
            raise OperationPartialError(
                "The canonical change and rebuild succeeded, but delivery failed. "
                f"The verified current files remain available. Cause: {publish_error}",
                report,
            )
        return report
    finally:
        if candidate.exists():
            candidate.unlink()
        try:
            temporary_root.rmdir()
        except OSError:
            pass


def restore_current(resolved: ResolvedDocument, *, clean: bool = False) -> dict[str, Any]:
    """Restore current from its immutable build, retaining displaced files in history."""

    current_root = resolved.system_root / "current" / resolved.manifest.id
    report_path = current_root / "build-report.json"
    if not report_path.is_file():
        raise RevisionError(f"No current build report exists: {report_path}")
    try:
        report = json.loads(report_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise RevisionError(f"Could not read current build report {report_path}: {exc}") from exc
    artifacts = verified_build_artifacts(report)
    build_id = str(report.get("build_id") or "unknown")
    operation_id = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S_%fZ")
    history = current_root / ".history" / "restores" / f"{operation_id}-{build_id}"
    stage = current_root / f".restore-{operation_id}"
    stage.mkdir(parents=True, exist_ok=False)
    target_by_kind = {
        "docx": current_root / f"{resolved.manifest.outputs.basename}.docx",
        "pdf": current_root / f"{resolved.manifest.outputs.basename}.pdf",
    }
    staged: dict[str, Path] = {}
    moved: list[dict[str, str]] = []
    committed: list[Path] = []
    try:
        for artifact in artifacts:
            source = Path(artifact["path"])
            candidate = stage / target_by_kind[artifact["kind"]].name
            shutil.copy2(source, candidate)
            if file_hash(candidate) != artifact["sha256"]:
                raise RevisionError(f"Restore staging hash mismatch: {source}")
            staged[artifact["kind"]] = candidate

        to_retain: list[Path] = []
        for target in target_by_kind.values():
            if target.exists():
                to_retain.append(target)
        if clean and current_root.is_dir():
            expected = {report_path.name, *(path.name for path in target_by_kind.values())}
            to_retain.extend(
                path
                for path in current_root.iterdir()
                if path.is_file() and path.name not in expected
            )
        unique_to_retain = list(dict.fromkeys(to_retain))
        if unique_to_retain:
            history.mkdir(parents=True, exist_ok=False)
            for path in unique_to_retain:
                retained = history / path.name
                os.replace(path, retained)
                moved.append({"from": str(path), "to": str(retained)})

        for kind, candidate in staged.items():
            os.replace(candidate, target_by_kind[kind])
            committed.append(target_by_kind[kind])
            expected_hash = next(item["sha256"] for item in artifacts if item["kind"] == kind)
            if file_hash(target_by_kind[kind]) != expected_hash:
                raise RevisionError(f"Restored current {kind.upper()} failed final hash verification")
    except Exception:
        for path in committed:
            try:
                path.unlink()
            except OSError:
                pass
        for item in reversed(moved):
            retained = Path(item["to"])
            original = Path(item["from"])
            if retained.exists() and not original.exists():
                try:
                    os.replace(retained, original)
                except OSError:
                    pass
        raise
    finally:
        shutil.rmtree(stage, ignore_errors=True)

    return {
        "schema": "agentic-current-restore-report/v2",
        "document_id": resolved.manifest.id,
        "build_id": build_id,
        "clean": clean,
        "history": str(history) if moved else None,
        "retained": moved,
        "current": {
            kind: {
                "path": str(target_by_kind[kind]),
                "sha256": file_hash(target_by_kind[kind]),
            }
            for kind in staged
        },
    }
