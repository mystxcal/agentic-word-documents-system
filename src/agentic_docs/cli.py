from __future__ import annotations

import argparse
import json
import os
import sys
import traceback
from pathlib import Path
from typing import Callable

from .diagnostics import DiagnosticBag
from .errors import DocumentSystemError
from .reporting import resolved_summary
from .resolver import resolve_document
from .workspace import discover_documents, discover_system_root, resolve_document_spec


def _document_argument(command: argparse.ArgumentParser) -> None:
    command.add_argument("document", help="Document id or path to a V2 document.jsonc manifest")


def _metadata_items(values: list[str]) -> dict[str, str]:
    result: dict[str, str] = {}
    for item in values:
        key, separator, value = item.partition("=")
        key = key.strip()
        if not separator or not key:
            raise DocumentSystemError(f"Project metadata must use KEY=VALUE; received {item!r}")
        if key in result:
            raise DocumentSystemError(f"Project metadata key {key!r} was supplied more than once")
        result[key] = value.strip()
    return result


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(
        prog="agentic-doc",
        description="Agentic Word Documents",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""Everyday loop:
  agentic-doc workspace DOCUMENT          orient from the live canonical map
  agentic-doc edit DOCUMENT               open the primary canonical source
  agentic-doc build DOCUMENT --quick      render a lightweight Word/PDF proof with heavy placeholders
  agentic-doc build DOCUMENT              create the full rendered-page proof
  agentic-doc open DOCUMENT --pdf         review the current verified PDF

Add --json to any command for the complete machine-readable result.
Run agentic-doc COMMAND --help for command-specific options.""",
    )
    result.add_argument("--json", action="store_true", help="Print the full machine-readable result")
    result.add_argument(
        "--root",
        help="Document-system root; may be written before or after the command",
    )
    commands = result.add_subparsers(dest="command", required=True)

    commands.add_parser("documents", help="List discoverable documents and their manifest paths")

    schema = commands.add_parser("schema", help="Refresh or check editor schemas for all manifest types")
    schema.add_argument("--check", action="store_true", help="Check generated schemas without changing files")

    inspect_source = commands.add_parser(
        "inspect-source",
        help="Inspect a Word, Markdown, Excel, PDF, image, or Visio source and show usable structural facts",
    )
    inspect_source.add_argument("path")

    new = commands.add_parser("new", help="Create a safe project or Markdown-first document starter")
    new_kinds = new.add_subparsers(dest="new_kind", required=True)
    new_project = new_kinds.add_parser("project", help="Create an empty document collection")
    new_project.add_argument("id")
    new_project.add_argument("--name", required=True)
    new_project.add_argument("--description")
    new_project.add_argument(
        "--meta",
        action="append",
        default=[],
        metavar="KEY=VALUE",
        help="Optional open project metadata; repeat for additional values",
    )

    new_document = new_kinds.add_parser("document", help="Create a Markdown-first document starter")
    new_document.add_argument("id")
    new_document.add_argument("--project", required=True)
    new_document.add_argument("--title", required=True)
    new_document.add_argument("--type", default="Document", dest="document_type")
    new_document.add_argument("--revision", default="draft")
    new_document.add_argument("--profile", default="plain")
    new_document.add_argument("--kit", default="studio")
    new_document.add_argument("--styles", default="studio")
    new_document.add_argument("--short-title")
    new_document.add_argument("--number")
    new_document.add_argument("--revision-display")
    new_document.add_argument("--date")
    new_document.add_argument("--prepared-by")
    new_document.add_argument("--basename")
    new_document.add_argument("--word-source", help="Existing DOCX to keep as the canonical Word-owned body")
    new_document.add_argument("--source-tag", help="Root content-control tag to extract from --word-source")
    new_document.add_argument("--allow-untagged", action="store_true", help="Explicitly import the whole Word body")
    new_document.add_argument("--preserve-sections", action="store_true")
    new_document.add_argument(
        "--use-source-styles",
        action="store_true",
        help="Use the canonical Word source itself as this document's style donor",
    )
    new_document.add_argument(
        "--preserve-source-layout",
        action="store_true",
        help="Keep the Word source's sections, page numbering, headers, footers, and page furniture",
    )

    add = commands.add_parser("add", help="Register and place one canonical component without hand-editing JSONC")
    _document_argument(add)
    add_kinds = add.add_subparsers(dest="component_kind", required=True)

    def placement_options(command: argparse.ArgumentParser) -> None:
        placement = command.add_mutually_exclusive_group(required=True)
        placement.add_argument("--into", dest="parent_component", help="Nest in this Markdown or Word parent")
        placement.add_argument("--region", help="Place as a top-level item in this page region")
        command.add_argument("--slot", help="Insertion marker/tag; defaults to the new component id")
        command.add_argument("--after", dest="after_component", help="For --region, insert after this component")
        command.add_argument(
            "--append-marker",
            action="store_true",
            help="For Markdown --into only, deliberately append a missing :::insert marker at the end",
        )

    prose = add_kinds.add_parser("prose", help="Add a Markdown+ prose component")
    prose.add_argument("id")
    prose.add_argument("--source", help="Existing Markdown source; omit to create a new file")
    prose.add_argument("--title", help="Required when creating a new Markdown source")
    placement_options(prose)

    word_fragment = add_kinds.add_parser("word-fragment", help="Add a native Word-owned component")
    word_fragment.add_argument("id")
    word_fragment.add_argument("--source", required=True)
    word_fragment.add_argument("--source-tag")
    word_fragment.add_argument("--allow-untagged", action="store_true")
    word_fragment.add_argument("--preserve-sections", action="store_true")
    placement_options(word_fragment)

    table = add_kinds.add_parser("table", help="Add an arbitrary Excel Table or explicit range")
    table.add_argument("id")
    table.add_argument("--source", required=True)
    locator = table.add_mutually_exclusive_group(required=True)
    locator.add_argument("--table", dest="excel_table", help="Excel Table name")
    locator.add_argument("--range", dest="excel_range", help="Explicit Excel cell range such as B4:H20")
    table.add_argument("--sheet", dest="excel_sheet", help="Required with --range")
    table.add_argument("--style-role", default="technical")
    table.add_argument(
        "--formula-policy",
        choices=("cached_values", "require_no_formulas", "require_cached_results"),
        default="cached_values",
    )
    placement_options(table)

    pdf_pages = add_kinds.add_parser("pdf-pages", help="Add explicit pages from an original PDF")
    pdf_pages.add_argument("id")
    pdf_pages.add_argument("--source", required=True)
    pdf_pages.add_argument("--pages", required=True, help="One-based pages such as 1,3-5,8")
    pdf_pages.add_argument("--title")
    pdf_pages.add_argument("--caption", nargs="?", const=True)
    pdf_pages.add_argument("--alt-text")
    pdf_pages.add_argument("--dpi", type=int, default=150)
    pdf_pages.add_argument("--width", type=float, dest="width_inches")
    pdf_pages.add_argument("--alignment", choices=("left", "center", "right"), default="center")
    pdf_pages.add_argument("--no-page-break-between", action="store_true")
    placement_options(pdf_pages)

    figure = add_kinds.add_parser("figure", help="Add a reviewed raster figure")
    figure.add_argument("id")
    figure.add_argument("--source", required=True)
    figure.add_argument("--title")
    figure.add_argument("--caption", nargs="?", const=True)
    figure.add_argument("--alt-text", required=True)
    figure.add_argument("--width", type=float, dest="width_inches")
    figure.add_argument("--alignment", choices=("left", "center", "right"), default="center")
    placement_options(figure)

    diagram = add_kinds.add_parser(
        "diagram",
        help="Add an editable native drawing plus its explicitly reviewed publication rendition",
    )
    diagram.add_argument("id")
    diagram.add_argument("--source", required=True, help="Native editable drawing, such as VSDX or draw.io")
    diagram.add_argument("--rendition", required=True, help="Reviewed PNG/JPEG/TIFF export used in Word")
    diagram.add_argument("--title")
    diagram.add_argument("--caption", nargs="?", const=True)
    diagram.add_argument("--alt-text", required=True)
    diagram.add_argument("--width", type=float, dest="width_inches")
    diagram.add_argument("--alignment", choices=("left", "center", "right"), default="center")
    placement_options(diagram)

    page_break = add_kinds.add_parser("page-break", help="Add an explicit page break")
    page_break.add_argument("id")
    placement_options(page_break)

    accept_rendition = commands.add_parser(
        "accept-rendition",
        help="Bind a reviewed diagram export to the current revision of its native drawing",
    )
    _document_argument(accept_rendition)
    accept_rendition.add_argument("--component", required=True)
    accept_rendition.add_argument("--file", required=True, help="Reviewed raster export")

    validate = commands.add_parser("validate", help="Validate and resolve a document without building it")
    _document_argument(validate)

    status = commands.add_parser("status", help="Show resolved layers, sequence, ownership, sources, and diagnostics")
    _document_argument(status)

    explain = commands.add_parser("explain", help="Explain one resolved component and where to edit it")
    _document_argument(explain)
    explain.add_argument("--component", required=True)

    workspace = commands.add_parser(
        "workspace",
        help="Show the live edit map, primary source, outputs, state, and exact next commands",
    )
    _document_argument(workspace)

    edit = commands.add_parser("edit", help="Open or locate the primary canonical source without hunting through folders")
    _document_argument(edit)
    edit_selection = edit.add_mutually_exclusive_group()
    edit_selection.add_argument("--component", help="Edit this exact canonical component")
    edit_selection.add_argument(
        "--presentation",
        help="Edit cover, header, footer, styles, shell, or an exact page-region presentation role",
    )
    edit.add_argument("--show", action="store_true", help="Print the selected path without opening an application")

    open_item = commands.add_parser("open", help="Open the exact current output or one canonical source")
    _document_argument(open_item)
    open_selection = open_item.add_mutually_exclusive_group()
    open_selection.add_argument("--component", help="Open this canonical component source")
    open_selection.add_argument("--pdf", action="store_true", help="Open the current verified PDF path")
    open_selection.add_argument("--manifest", action="store_true", help="Open document.jsonc")
    open_selection.add_argument("--folder", action="store_true", help="Open the canonical document folder")

    build = commands.add_parser("build", help="Build an immutable draft run")
    _document_argument(build)
    build.add_argument("--output-root")
    build.add_argument(
        "--quick",
        action="store_true",
        help="Build and render a lightweight proof; PDF-page components become visible placeholders",
    )
    build.add_argument(
        "--include-heavy",
        action="store_true",
        help="With --quick, include heavy components instead of their preview policy",
    )
    build.add_argument(
        "--retain-intermediates",
        action="store_true",
        help="Keep raw Word and source-adapter work files for diagnosis",
    )

    preview = commands.add_parser("preview", help="Build a small component or presentation proof")
    _document_argument(preview)
    preview_target = preview.add_mutually_exclusive_group(required=True)
    preview_target.add_argument("--component")
    preview_target.add_argument(
        "--presentation",
        choices=("page-furniture",),
        help="Build a compact cover/header/footer/page-region proof with synthetic content",
    )
    preview.add_argument(
        "--quick",
        action="store_true",
        help="Defer heavy nested components according to their preview policy",
    )
    preview.add_argument(
        "--include-heavy",
        action="store_true",
        help="Include heavy nested components even in a quick component preview",
    )
    preview.add_argument(
        "--retain-intermediates",
        action="store_true",
        help="Keep raw Word and source-adapter work files for diagnosis",
    )

    release = commands.add_parser("release", help="Create a controlled immutable release after all gates pass")
    _document_argument(release)
    release.add_argument(
        "--retain-intermediates",
        action="store_true",
        help="Keep raw Word and source-adapter work files for diagnosis",
    )

    inspect_word = commands.add_parser("inspect-docx", help="Inspect any Word donor, source, or output without changing it")
    inspect_word.add_argument("path")

    commands.add_parser("doctor", help="Check the local Word, Python, and PDF runtime")

    audit = commands.add_parser("audit", help="Verify canonical-source, immutable-build, and current-pair integrity")
    _document_argument(audit)

    check = commands.add_parser("check", help="Plain-language alias for audit")
    _document_argument(check)

    paths = commands.add_parser("paths", help="Show where to edit, review, and find builds")
    _document_argument(paths)

    history = commands.add_parser("history", help="Show recent builds, changes, deliveries, and recoveries")
    _document_argument(history)
    history.add_argument("--limit", type=int, default=20)

    builds = commands.add_parser("builds", help="List immutable builds, proof level, and storage use")
    _document_argument(builds)
    builds.add_argument("--limit", type=int, default=20)

    retention = commands.add_parser(
        "retention",
        help="Plan or recoverably archive old drafts without touching releases",
    )
    _document_argument(retention)
    retention.add_argument("--keep-drafts", type=int, default=10)
    retention.add_argument("--keep-previews", type=int, default=3)
    retention.add_argument(
        "--apply",
        action="store_true",
        help="Move planned candidates to the recoverable archive; default is read-only",
    )

    revise = commands.add_parser(
        "revise",
        help="Replace text in canonical Word or Markdown, rebuild, and optionally publish",
    )
    _document_argument(revise)
    revise.add_argument(
        "--component",
        help="Canonical Word-fragment or Markdown component id; auto-selected when unique",
    )
    revise.add_argument(
        "--replace",
        nargs=2,
        action="append",
        required=True,
        metavar=("FIND", "REPLACE"),
        dest="replacements",
        help="Exact text replacement; repeat the option to apply multiple replacements",
    )
    revise.add_argument("--ignore-case", action="store_true")
    revise.add_argument("--whole-word", action="store_true")
    revise.add_argument("--expect", type=int, help="Abort unless the total replacement count matches")
    revise.add_argument("--allow-zero", action="store_true")
    revise.add_argument("--all-stories", action="store_true", help="Also search canonical headers, footers, and notes")
    revise.add_argument("--no-build", action="store_true", help="Revise the canonical source without rebuilding")
    revise.add_argument("--quick", action="store_true", help="Create Word/PDF without rendering every PDF page")
    revise.add_argument("--output-root", help="Alternative immutable build root")
    revise.add_argument(
        "--retain-intermediates",
        action="store_true",
        help="Keep raw Word and source-adapter work files for diagnosis",
    )
    revise.add_argument("--publish", help="Publish the verified DOCX/PDF pair to this directory")
    revise.add_argument(
        "--policy",
        choices=("fail", "versioned", "replace"),
        default="versioned",
        help="Collision policy for --publish (default: versioned)",
    )

    publish = commands.add_parser("publish", help="Publish the latest verified immutable DOCX/PDF build pair")
    _document_argument(publish)
    publish.add_argument("--to", required=True, help="Destination directory")
    publish.add_argument(
        "--policy",
        choices=("fail", "versioned", "replace"),
        default="versioned",
    )
    publish.add_argument(
        "--allow-out-of-sync",
        action="store_true",
        help="Deliberately send the recorded immutable build even when current/canonical files differ",
    )
    publish.add_argument("--build-id", help="Publish this exact immutable build instead of the current recorded build")

    replace = commands.add_parser(
        "replace",
        help="Find/replace canonical Word or Markdown text and regenerate the document in one command",
    )
    _document_argument(replace)
    replace.add_argument("find", help="Exact text to find")
    replace.add_argument("replacement", help="Replacement text")
    replace.add_argument("--component", help="Canonical Word component; auto-selected when unique")
    replace.add_argument("--ignore-case", action="store_true")
    replace.add_argument("--whole-word", action="store_true")
    replace.add_argument("--expect", type=int, help="Abort unless this many replacements are found")
    replace.add_argument("--allow-zero", action="store_true")
    replace.add_argument("--all-stories", action="store_true")
    replace.add_argument("--no-build", action="store_true")
    replace.add_argument("--quick", action="store_true", help="Create Word/PDF without rendering every PDF page")
    replace.add_argument("--output-root")
    replace.add_argument(
        "--retain-intermediates",
        action="store_true",
        help="Keep raw Word and source-adapter work files for diagnosis",
    )
    replace.add_argument("--publish")
    replace.add_argument("--policy", choices=("fail", "versioned", "replace"), default="versioned")

    deliver = commands.add_parser("deliver", help="Plain-language alias for publish")
    _document_argument(deliver)
    deliver.add_argument("destination", help="Destination directory")
    deliver.add_argument("--policy", choices=("fail", "versioned", "replace"), default="versioned")
    deliver.add_argument(
        "--allow-out-of-sync",
        action="store_true",
        help="Deliberately send the recorded immutable build even when current/canonical files differ",
    )
    deliver.add_argument("--build-id", help="Publish this exact immutable build instead of the current recorded build")

    restore = commands.add_parser(
        "restore-current",
        help="Restore the current DOCX/PDF from the immutable build and retain displaced files",
    )
    _document_argument(restore)
    restore.add_argument(
        "--clean",
        action="store_true",
        help="Also move stale current-folder files into the restore history",
    )

    workcopy_diff = commands.add_parser(
        "diff-workcopy",
        help="Compare a coworker-edited compiled Word file with its canonical component",
    )
    _document_argument(workcopy_diff)
    workcopy_diff.add_argument("workcopy", help="Edited compiled DOCX received from a coworker")
    workcopy_diff.add_argument("--component", help="Tagged word_fragment component id; auto-selected when unique")
    workcopy_diff.add_argument("--max-diff-lines", type=int, default=400)

    compare = commands.add_parser("compare", help="Plain-language alias for diff-workcopy")
    _document_argument(compare)
    compare.add_argument("workcopy", help="Edited compiled DOCX received from a coworker")
    compare.add_argument("--component", help="Tagged Word component; auto-selected when unique")
    compare.add_argument("--max-diff-lines", type=int, default=400)

    adopt = commands.add_parser(
        "adopt",
        help="Adopt a tagged component from a coworker Word workcopy, rebuild, and optionally publish",
    )
    _document_argument(adopt)
    adopt.add_argument("workcopy", help="Edited compiled DOCX received from a coworker")
    adopt.add_argument("--component", help="Tagged word_fragment component id; auto-selected when unique")
    adopt.add_argument("--allow-identical", action="store_true")
    adopt.add_argument(
        "--accept-conflict",
        action="store_true",
        help="Intentionally replace canonical changes when both sides changed since the build baseline",
    )
    adopt.add_argument("--no-build", action="store_true")
    adopt.add_argument("--quick", action="store_true", help="Create Word/PDF without rendering every PDF page")
    adopt.add_argument("--output-root")
    adopt.add_argument(
        "--retain-intermediates",
        action="store_true",
        help="Keep raw Word and source-adapter work files for diagnosis",
    )
    adopt.add_argument("--publish")
    adopt.add_argument(
        "--policy",
        choices=("fail", "versioned", "replace"),
        default="versioned",
    )

    return result


def _explain(resolved, component_id: str) -> dict:
    if component_id not in resolved.components:
        raise DocumentSystemError(f"Unknown component {component_id!r}")
    item = resolved.components[component_id]
    declaration = item.declaration
    if declaration.ownership == "word_fragment":
        edit_instruction = "Open the canonical DOCX fragment in Microsoft Word."
    elif item.source_path and item.source_path.suffix.lower() in {".md", ".markdown"}:
        edit_instruction = (
            "Edit the canonical Markdown+ source and rebuild. Compiled Word edits are not reverse-synced."
        )
    elif declaration.ownership == "source":
        edit_instruction = "Edit the declared canonical source; conflicting edits in a compiled workcopy will not be overwritten silently."
    else:
        edit_instruction = "Replace the registered snapshot explicitly when a new reviewed asset is ready."
    return {
        "component": component_id,
        "type": declaration.type,
        "ownership": declaration.ownership,
        "canonical_source": str(item.source_path) if item.source_path else None,
        "source_hash": item.source_hash,
        "edit_instruction": edit_instruction,
    }


def _paths(resolved) -> dict:
    current_root = resolved.system_root / "current" / resolved.manifest.id
    report_path = current_root / "build-report.json"
    current_docx = current_root / f"{resolved.manifest.outputs.basename}.docx"
    current_pdf = current_root / f"{resolved.manifest.outputs.basename}.pdf"
    return {
        "schema": "agentic-document-paths/v2",
        "document_id": resolved.manifest.id,
        "manifest": str(resolved.manifest_path),
        "document_root": str(resolved.manifest_path.parent),
        "current_root": str(current_root),
        "current_report": str(report_path) if report_path.is_file() else None,
        "current_docx": str(current_docx) if current_docx.is_file() else None,
        "current_pdf": str(current_pdf) if current_pdf.is_file() else None,
        "builds_root": str(resolved.system_root / "builds" / resolved.manifest.id),
        "activity_log": str(resolved.system_root / "operations" / resolved.manifest.id / "activity.jsonl"),
        "canonical_components": [
            {
                "id": component_id,
                "type": item.declaration.type,
                "ownership": item.declaration.ownership,
                "path": str(item.source_path) if item.source_path else None,
            }
            for component_id, item in resolved.components.items()
        ],
    }


def _history(resolved, limit: int) -> dict:
    if limit < 1 or limit > 500:
        raise DocumentSystemError("--limit must be between 1 and 500")
    path = resolved.system_root / "operations" / resolved.manifest.id / "activity.jsonl"
    records = []
    if path.is_file():
        try:
            for line in path.read_text(encoding="utf-8").splitlines():
                if line.strip():
                    value = json.loads(line)
                    if isinstance(value, dict):
                        records.append(value)
        except (OSError, json.JSONDecodeError) as exc:
            raise DocumentSystemError(f"Could not read activity log {path}: {exc}") from exc
    return {
        "schema": "agentic-activity-history/v2",
        "document_id": resolved.manifest.id,
        "path": str(path),
        "events": records[-limit:],
        "total_events": len(records),
    }


def _build_history(resolved, limit: int) -> dict:
    if limit < 1 or limit > 500:
        raise DocumentSystemError("--limit must be between 1 and 500")
    root = resolved.system_root / "builds" / resolved.manifest.id
    current_report_path = resolved.system_root / "current" / resolved.manifest.id / "build-report.json"
    current_build_id = None
    if current_report_path.is_file():
        try:
            current_build_id = json.loads(current_report_path.read_text(encoding="utf-8")).get("build_id")
        except (OSError, json.JSONDecodeError, AttributeError):
            pass
    records = []
    total_bytes = 0
    if root.is_dir():
        for directory in sorted(
            (path for path in root.iterdir() if path.is_dir() and not path.name.startswith(".")),
            key=lambda path: path.name,
        ):
            size = sum(path.stat().st_size for path in directory.rglob("*") if path.is_file())
            total_bytes += size
            report_path = directory / "build-report.json"
            report = {}
            if report_path.is_file():
                try:
                    report = json.loads(report_path.read_text(encoding="utf-8"))
                except (OSError, json.JSONDecodeError):
                    report = {}
            records.append(
                {
                    "build_id": report.get("build_id") or directory.name,
                    "directory": str(directory),
                    "size_bytes": size,
                    "mode": report.get("mode"),
                    "verification_mode": report.get("verification_mode") or "unknown",
                    "quality_passed": (report.get("quality") or {}).get("passed") is True,
                    "word": (report.get("artifacts") or {}).get("docx"),
                    "pdf": (report.get("artifacts") or {}).get("pdf"),
                    "current": (report.get("build_id") or directory.name) == current_build_id,
                }
            )
    return {
        "schema": "agentic-build-history/v2",
        "document_id": resolved.manifest.id,
        "root": str(root),
        "current_build_id": current_build_id,
        "total_builds": len(records),
        "total_size_bytes": total_bytes,
        "builds": records[-limit:][::-1],
    }


def _activity_details(value: dict) -> dict:
    return {
        key: value.get(key)
        for key in (
            "operation_id",
            "build_id",
            "run_directory",
            "canonical_source",
            "backup",
            "destination",
        )
        if value.get(key) is not None
    }


def _mutate(resolved, operation: str, action: Callable[[], dict]) -> dict:
    from .errors import OperationPartialError
    from .journal import append_activity
    from .locking import document_lock

    with document_lock(resolved.system_root, resolved.manifest.id, operation):
        try:
            value = action()
        except OperationPartialError as exc:
            append_activity(
                resolved.system_root,
                resolved.manifest.id,
                operation,
                "partial",
                details={**_activity_details(exc.report), "error": str(exc)},
            )
            raise
        except Exception as exc:
            append_activity(
                resolved.system_root,
                resolved.manifest.id,
                operation,
                "failed",
                details={"error": str(exc)},
            )
            raise
        append_activity(
            resolved.system_root,
            resolved.manifest.id,
            operation,
            "succeeded",
            details=_activity_details(value),
        )
        return value


def main(argv: list[str] | None = None) -> int:
    raw_arguments = list(sys.argv[1:] if argv is None else argv)
    json_output = "--json" in raw_arguments
    raw_arguments = [item for item in raw_arguments if item != "--json"]
    root_override = None
    for index, item in enumerate(list(raw_arguments)):
        if item.startswith("--root="):
            if root_override is not None:
                print("ERROR: --root may be specified only once", file=sys.stderr)
                return 2
            root_override = item.split("=", 1)[1]
            raw_arguments.remove(item)
        elif item == "--root":
            if root_override is not None:
                print("ERROR: --root may be specified only once", file=sys.stderr)
                return 2
            try:
                root_override = raw_arguments[index + 1]
            except IndexError:
                print("ERROR: --root requires a folder path", file=sys.stderr)
                return 2
            del raw_arguments[index:index + 2]
            break
    arguments = parser().parse_args(raw_arguments)
    arguments.root = root_override or arguments.root
    diagnostics = DiagnosticBag()
    try:
        if arguments.command == "documents":
            root = discover_system_root(
                Path.cwd(),
                explicit=Path(arguments.root) if arguments.root else None,
            )
            value = {
                "schema": "agentic-document-catalog/v2",
                "system_root": str(root),
                "documents": [
                    {
                        "id": item.id,
                        "project": item.project,
                        "title": item.title,
                        "manifest": str(item.manifest_path),
                        "error": item.error,
                    }
                    for item in discover_documents(root)
                ],
            }
        elif arguments.command == "doctor":
            from .inspection import doctor_status

            value = doctor_status()
        elif arguments.command == "schema":
            from .schemas import refresh_schemas

            root = discover_system_root(
                Path.cwd(),
                explicit=Path(arguments.root) if arguments.root else None,
            )
            value = refresh_schemas(root, check_only=arguments.check)
        elif arguments.command == "inspect-source":
            from .inspection import inspect_source

            value = inspect_source(Path(arguments.path))
        elif arguments.command == "inspect-docx":
            from .inspection import inspect_docx

            value = inspect_docx(Path(arguments.path))
        elif arguments.command == "new":
            from .authoring import create_document, create_project

            root = discover_system_root(
                Path.cwd(),
                explicit=Path(arguments.root) if arguments.root else None,
            )
            if arguments.new_kind == "project":
                value = create_project(
                    root,
                    arguments.id,
                    name=arguments.name,
                    description=arguments.description,
                    metadata=_metadata_items(arguments.meta),
                )
            else:
                value = create_document(
                    root,
                    arguments.id,
                    project_id=arguments.project,
                    title=arguments.title,
                    document_type=arguments.document_type,
                    revision=arguments.revision,
                    profile_id=arguments.profile,
                    kit_id=arguments.kit,
                    short_title=arguments.short_title,
                    number=arguments.number,
                    revision_display=arguments.revision_display,
                    document_date=arguments.date,
                    prepared_by=arguments.prepared_by,
                    basename=arguments.basename,
                    styles=arguments.styles,
                    word_source=Path(arguments.word_source) if arguments.word_source else None,
                    source_tag=arguments.source_tag,
                    allow_untagged=arguments.allow_untagged,
                    preserve_sections=arguments.preserve_sections,
                    use_source_styles=arguments.use_source_styles,
                    preserve_source_layout=arguments.preserve_source_layout,
                )
        else:
            manifest_path = resolve_document_spec(
                arguments.document,
                start=Path.cwd(),
                system_root=Path(arguments.root) if arguments.root else None,
            )
            resolved = resolve_document(
                manifest_path,
                diagnostics,
                allow_stale_diagrams=arguments.command == "accept-rendition",
            )

        if arguments.command in {"validate", "status"}:
            value = resolved_summary(resolved, diagnostics)
            if arguments.command == "status":
                from .integrity import audit_document

                value["integrity"] = audit_document(resolved)
        elif arguments.command == "explain":
            value = _explain(resolved, arguments.component)
        elif arguments.command == "workspace":
            from .guidance import document_workspace
            from .integrity import audit_document

            value = document_workspace(resolved, audit_document(resolved))
        elif arguments.command == "edit":
            from .guidance import edit_target

            value = edit_target(
                resolved,
                component_id=arguments.component,
                presentation=arguments.presentation,
                open_file=not arguments.show,
            )
        elif arguments.command == "open":
            from .authoring import open_document

            value = open_document(
                resolved,
                component_id=arguments.component,
                pdf=arguments.pdf,
                manifest=arguments.manifest,
                folder=arguments.folder,
            )
        elif arguments.command in {"build", "preview"}:
            from .build import build_document

            value = _mutate(
                resolved,
                arguments.command,
                lambda: build_document(
                    resolved,
                    diagnostics,
                    output_root=Path(arguments.output_root).resolve() if getattr(arguments, "output_root", None) else None,
                    component_id=arguments.component if arguments.command == "preview" else None,
                    preview_presentation=getattr(arguments, "presentation", None)
                    if arguments.command == "preview"
                    else None,
                    render_pages=True,
                    lightweight=bool(arguments.quick),
                    include_heavy=bool(getattr(arguments, "include_heavy", False)),
                    update_current=arguments.command == "build" and not arguments.quick,
                    retain_intermediates=arguments.retain_intermediates,
                ),
            )
        elif arguments.command == "add":
            from .authoring import add_document_component, parse_pages

            value = _mutate(
                resolved,
                "add",
                lambda: add_document_component(
                    resolved,
                    kind=arguments.component_kind,
                    component_id=arguments.id,
                    source=Path(arguments.source) if getattr(arguments, "source", None) else None,
                    rendition=Path(arguments.rendition) if getattr(arguments, "rendition", None) else None,
                    title=getattr(arguments, "title", None),
                    caption=getattr(arguments, "caption", None),
                    alt_text=getattr(arguments, "alt_text", None),
                    parent_component=arguments.parent_component,
                    slot_name=arguments.slot,
                    region=arguments.region,
                    after_component=arguments.after_component,
                    append_marker=arguments.append_marker,
                    source_tag=getattr(arguments, "source_tag", None),
                    allow_untagged=getattr(arguments, "allow_untagged", False),
                    preserve_sections=getattr(arguments, "preserve_sections", False),
                    excel_table=getattr(arguments, "excel_table", None),
                    excel_sheet=getattr(arguments, "excel_sheet", None),
                    excel_range=getattr(arguments, "excel_range", None),
                    table_style_role=getattr(arguments, "style_role", "technical"),
                    formula_policy=getattr(arguments, "formula_policy", "cached_values"),
                    pages=parse_pages(arguments.pages) if getattr(arguments, "pages", None) else None,
                    dpi=getattr(arguments, "dpi", 150),
                    width_inches=getattr(arguments, "width_inches", None),
                    alignment=getattr(arguments, "alignment", "center"),
                    page_break_between=not getattr(arguments, "no_page_break_between", False),
                ),
            )
        elif arguments.command == "accept-rendition":
            from .authoring import accept_diagram_rendition

            value = _mutate(
                resolved,
                "accept-rendition",
                lambda: accept_diagram_rendition(
                    resolved,
                    component_id=arguments.component,
                    rendition=Path(arguments.file),
                ),
            )
        elif arguments.command == "release":
            from .release import release_document

            value = _mutate(
                resolved,
                "release",
                lambda: release_document(
                    resolved,
                    diagnostics,
                    retain_intermediates=arguments.retain_intermediates,
                ),
            )
        elif arguments.command in {"audit", "check"}:
            from .integrity import audit_document

            value = audit_document(resolved)
        elif arguments.command == "paths":
            value = _paths(resolved)
        elif arguments.command == "history":
            value = _history(resolved, arguments.limit)
        elif arguments.command == "builds":
            value = _build_history(resolved, arguments.limit)
        elif arguments.command == "retention":
            from .retention import manage_retention

            action = lambda: manage_retention(
                resolved.system_root,
                resolved.manifest.id,
                keep_drafts=arguments.keep_drafts,
                keep_previews=arguments.keep_previews,
                apply=arguments.apply,
            )
            value = _mutate(resolved, "retention", action) if arguments.apply else action()
        elif arguments.command in {"revise", "replace"}:
            from .docx_text import TextReplacement
            from .operations import revise_document

            pairs = arguments.replacements if arguments.command == "revise" else [(arguments.find, arguments.replacement)]
            replacements = [
                TextReplacement(find=find, replace=replace, case_sensitive=not arguments.ignore_case, whole_word=arguments.whole_word)
                for find, replace in pairs
            ]
            value = _mutate(
                resolved,
                arguments.command,
                lambda: revise_document(
                    resolved,
                    diagnostics,
                    replacements,
                    component_id=arguments.component,
                    expected_total=arguments.expect,
                    allow_zero=arguments.allow_zero,
                    all_stories=arguments.all_stories,
                    build=not arguments.no_build,
                    render_pages=not arguments.quick,
                    retain_intermediates=arguments.retain_intermediates,
                    output_root=Path(arguments.output_root).resolve() if arguments.output_root else None,
                    publish_destination=Path(arguments.publish) if arguments.publish else None,
                    publish_policy=arguments.policy,
                ),
            )
        elif arguments.command in {"publish", "deliver"}:
            from .errors import IntegrityError
            from .publishing import (
                load_immutable_build_report,
                publish_build,
                require_delivery_alignment,
            )

            if arguments.build_id:
                report, report_path = load_immutable_build_report(
                    resolved.system_root,
                    resolved.manifest.id,
                    arguments.build_id,
                )
            else:
                if not arguments.allow_out_of_sync:
                    from .integrity import audit_document

                    require_delivery_alignment(audit_document(resolved))
                report_path = resolved.system_root / "current" / resolved.manifest.id / "build-report.json"
                if not report_path.is_file():
                    raise IntegrityError(f"No current build report exists: {report_path}")
                try:
                    report = json.loads(report_path.read_text(encoding="utf-8"))
                except (OSError, json.JSONDecodeError) as exc:
                    raise IntegrityError(f"Could not read current build report {report_path}: {exc}") from exc
            destination = arguments.to if arguments.command == "publish" else arguments.destination
            value = _mutate(
                resolved,
                arguments.command,
                lambda: publish_build(report, Path(destination), policy=arguments.policy),
            )
        elif arguments.command == "restore-current":
            from .operations import restore_current

            value = _mutate(
                resolved,
                "restore-current",
                lambda: restore_current(resolved, clean=arguments.clean),
            )
        elif arguments.command in {"diff-workcopy", "compare"}:
            from .workcopy import diff_workcopy

            value = diff_workcopy(
                resolved,
                Path(arguments.workcopy),
                component_id=arguments.component,
                max_diff_lines=arguments.max_diff_lines,
            )
        elif arguments.command == "adopt":
            from .workcopy import adopt_workcopy

            value = _mutate(
                resolved,
                "adopt",
                lambda: adopt_workcopy(
                    resolved,
                    diagnostics,
                    Path(arguments.workcopy),
                    component_id=arguments.component,
                    allow_identical=arguments.allow_identical,
                    accept_conflict=arguments.accept_conflict,
                    build=not arguments.no_build,
                    render_pages=not arguments.quick,
                    retain_intermediates=arguments.retain_intermediates,
                    output_root=Path(arguments.output_root).resolve() if arguments.output_root else None,
                    publish_destination=Path(arguments.publish) if arguments.publish else None,
                    publish_policy=arguments.policy,
                ),
            )
        elif arguments.command in {"documents", "doctor", "schema", "inspect-source", "inspect-docx", "new"}:
            pass
        else:
            raise DocumentSystemError(f"Unsupported command {arguments.command}")
        if json_output:
            print(json.dumps(value, indent=2, default=str))
        else:
            from .console import render_result

            print(render_result(arguments.command, value))
        if arguments.command in {"audit", "check"} and not value.get("ready", False):
            return 3
        if arguments.command == "schema" and arguments.check and not value.get("current", False):
            return 3
        return 0
    except DocumentSystemError as exc:
        from .errors import OperationPartialError

        if isinstance(exc, OperationPartialError):
            if json_output:
                print(json.dumps(exc.report, indent=2, default=str))
            else:
                from .console import render_result

                print(render_result(arguments.command, exc.report))
                print()
            print(f"PARTIAL: {exc}", file=sys.stderr)
            return 4
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:
        print(f"ERROR: Unexpected internal failure ({type(exc).__name__}): {exc}", file=sys.stderr)
        if os.environ.get("AGENTIC_DOCS_DEBUG") == "1":
            traceback.print_exc()
        else:
            print("Set AGENTIC_DOCS_DEBUG=1 to print the Python traceback.", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
