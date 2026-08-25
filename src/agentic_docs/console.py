from __future__ import annotations

from pathlib import Path
from typing import Any


def _yes(value: Any) -> str:
    return "YES" if value else "NO"


def _line(label: str, value: Any) -> str:
    return f"{label:<20} {value}"


def _artifact_lines(value: dict[str, Any]) -> list[str]:
    artifacts = value.get("artifacts") or {}
    return [
        _line("Word", artifacts.get("docx") or value.get("current_docx") or "not created"),
        _line("PDF", artifacts.get("pdf") or value.get("current_pdf") or "not created"),
    ]


def _documents(value: dict[str, Any]) -> str:
    items = value.get("documents", [])
    lines = ["AVAILABLE DOCUMENTS", ""]
    if not items:
        lines.append("No document manifests were found.")
    for item in items:
        if item.get("error"):
            lines.append(f"  INVALID: {item['manifest']}")
            lines.append(f"    {item['error']}")
            continue
        lines.append(f"  {item['id']}")
        lines.append(f"    {item.get('title') or '(untitled)'}")
        lines.append(f"    {item['manifest']}")
    lines.extend(["", "Use the short ID in every command; the full manifest path is optional."])
    return "\n".join(lines)


def _doctor(value: dict[str, Any]) -> str:
    checks = value.get("checks", {})
    lines = [
        "SYSTEM CHECK: " + ("READY" if value.get("ready") else "ATTENTION NEEDED"),
        "",
        _line("Version", value.get("system_version") or "unknown"),
    ]
    labels = {
        "python_packages": "Python libraries",
        "word_automation": "Microsoft Word",
        "pdf_rendering": "PDF renderer",
    }
    for key, label in labels.items():
        lines.append(_line(label, "OK" if checks.get(key) else "MISSING"))
    return "\n".join(lines)


def _audit(value: dict[str, Any]) -> str:
    lines = [
        "DOCUMENT CHECK: " + ("IN SYNC" if value.get("ready") else "ATTENTION NEEDED"),
        "",
        _line("Document", value.get("document_id")),
        _line("Build", value.get("build_id") or "none"),
        _line("Immutable pair", "VERIFIED" if value.get("immutable_integrity", value.get("artifact_integrity")) else "PROBLEM"),
        _line("Current pair", "VERIFIED" if value.get("current_integrity", value.get("artifact_integrity")) else "PROBLEM"),
        _line("Canonical inputs", "MATCH BUILD" if value.get("sources_current") else "CHANGED"),
        _line("Full build proof", "PASS" if value.get("build_quality_passed") else "NOT CONFIRMED"),
    ]
    issues = value.get("issues") or []
    if issues:
        lines.extend(["", "Findings:"])
        for issue in issues:
            severity = str(issue.get("severity", "info")).upper()
            lines.append(f"  [{severity}] {issue.get('message')}")
            if issue.get("hint"):
                lines.append(f"          {issue['hint']}")
    else:
        lines.extend(["", "No integrity or synchronization problems found."])
    return "\n".join(lines)


def _build(value: dict[str, Any]) -> str:
    quality = value.get("quality", {})
    mode = value.get("verification_mode", "full")
    artifacts = value.get("artifacts") or {}
    lines = [
        "BUILD COMPLETE",
        "",
        _line("Document", value.get("resolved", {}).get("document", {}).get("id")),
        _line("Build", value.get("build_id")),
        _line("Verification", str(mode).upper()),
        _line("Content scope", str(value.get("content_scope") or "complete").upper()),
        _line("Scoped proof", "PASS" if quality.get("scope_passed", quality.get("passed")) else "INCOMPLETE"),
        _line("Release ready", "YES" if quality.get("release_ready") else "NO"),
        _line("Immutable run", value.get("run_directory")),
        _line("Word", value.get("current_docx") or artifacts.get("docx") or "not created"),
        _line("PDF", value.get("current_pdf") or artifacts.get("pdf") or "not created"),
    ]
    inputs = value.get("inputs", {}).get("changes_since_current") or {}
    if inputs.get("baseline_available"):
        counts = inputs.get("counts") or {}
        change_total = sum(int(counts.get(name, 0)) for name in ("added", "removed", "changed"))
        lines.append(_line("Canonical inputs", "UNCHANGED" if not inputs.get("changed") else f"{change_total} CHANGE(S)"))
    cache = value.get("component_cache") or {}
    if cache.get("events"):
        lines.append(_line("Adapter cache", f"{cache.get('hits', 0)} HIT / {cache.get('misses', 0)} MISS"))
    comparison = value.get("comparison") or {}
    if comparison.get("baseline_available"):
        visual = comparison.get("visual") or {}
        if visual.get("available"):
            comparison_text = (
                "UNCHANGED"
                if not visual.get("changed") and not comparison.get("structure", {}).get("changed")
                else f"REVIEW {visual.get('changed_page_count', 0)} PAGE(S)"
            )
        else:
            comparison_text = "STRUCTURE CHANGED" if comparison.get("structure", {}).get("changed") else "STRUCTURE UNCHANGED"
        lines.append(_line("Against prior build", comparison_text))
    total_ms = value.get("timings", {}).get("total_ms")
    if total_ms is not None:
        lines.append(_line("Elapsed", f"{float(total_ms) / 1000:.1f} s"))
    if mode == "quick":
        lines.extend(["", "The Word and PDF were regenerated; page-image verification was intentionally skipped."])
    if value.get("release_directory"):
        lines.append(_line("Release directory", value.get("release_directory")))
    if value.get("current_update_error"):
        lines.append(_line("Current refresh", f"FAILED: {value['current_update_error']}"))
    return "\n".join(lines)


def _revision(value: dict[str, Any]) -> str:
    build = value.get("build") or {}
    replacement = value.get("replacement") or {}
    lines = [
        "CHANGE APPLIED" if value.get("changed", True) else "NO CHANGE MADE",
        "",
        _line("Document", value.get("document_id")),
        _line("Component", value.get("component")),
        _line("Replacements", replacement.get("total_replacements", 0)),
        _line("Canonical source", value.get("canonical_source")),
    ]
    if value.get("backup"):
        lines.append(_line("Recovery copy", value.get("backup")))
    if build.get("performed"):
        lines.extend(
            [
                _line("Build", build.get("build_id")),
                _line("Current Word", build.get("docx")),
                _line("Current PDF", build.get("pdf") or "not created"),
            ]
        )
        if build.get("current_update_error"):
            lines.append(_line("Current refresh", f"FAILED: {build['current_update_error']}"))
    else:
        lines.append(_line("Build", "not requested"))
    if value.get("publish"):
        publish = value["publish"]
        label = "Delivery FAILED" if publish.get("succeeded") is False else "Delivered to"
        lines.append(_line(label, publish.get("destination")))
    return "\n".join(lines)


def _publish(value: dict[str, Any]) -> str:
    lines = [
        "DELIVERY " + ("COMPLETE" if value.get("changed") else "ALREADY PRESENT"),
        "",
        _line("Build", value.get("build_id")),
        _line("Destination", value.get("destination")),
        _line("Collision policy", value.get("policy")),
    ]
    for artifact in value.get("artifacts", []):
        lines.append(_line(artifact.get("kind", "file").upper(), artifact.get("path")))
    if value.get("history"):
        lines.append(_line("Replaced-file history", value.get("history")))
    return "\n".join(lines)


def _restore(value: dict[str, Any]) -> str:
    lines = [
        "CURRENT OUTPUT RESTORED",
        "",
        _line("Document", value.get("document_id")),
        _line("Build", value.get("build_id")),
    ]
    for kind, item in (value.get("current") or {}).items():
        lines.append(_line(kind.upper(), item.get("path")))
    if value.get("history"):
        lines.append(_line("Displaced files kept", value.get("history")))
    return "\n".join(lines)


def _workcopy_diff(value: dict[str, Any]) -> str:
    formatting = value.get("formatting_same")
    same = value.get("same")
    lines = [
        "WORKCOPY COMPARISON",
        "",
        _line("Document", value.get("document_id")),
        _line("Component", value.get("component")),
        _line("Same as canonical", "YES" if same is True else "NO" if same is False else "UNKNOWN"),
        _line("Text/tables/images", "SAME" if value.get("content_same") else "CHANGED"),
        _line("Formatting", "SAME" if formatting is True else "CHANGED" if formatting is False else "UNKNOWN (older workcopy)"),
        _line("Comments/revisions", "SAME" if value.get("annotations_same") else "CHANGED"),
        _line("Change groups", value.get("change_group_count", 0)),
        _line("Canonical", value.get("canonical", {}).get("path")),
        _line("Workcopy", value.get("workcopy", {}).get("path")),
    ]
    if value.get("baseline_available"):
        lines.insert(8, _line("Canonical changed", _yes((value.get("canonical_changed_since_build") or {}).get("content"))))
        lines.insert(9, _line("Returned changed", _yes((value.get("workcopy_changed_since_build") or {}).get("content"))))
        lines.insert(10, _line("Content conflict", _yes(value.get("content_conflict"))))
    changes = value.get("changes") or []
    if changes:
        lines.extend(["", "Changed text groups:"])
        for index, change in enumerate(changes, 1):
            before = " | ".join(change.get("canonical") or []) or "(nothing)"
            after = " | ".join(change.get("workcopy") or []) or "(nothing)"
            lines.append(f"  {index}. {change.get('kind')}: {before}")
            lines.append(f"     -> {after}")
    return "\n".join(lines)


def _adoption(value: dict[str, Any]) -> str:
    lines = [
        "WORKCOPY ADOPTED",
        "",
        _line("Document", value.get("document_id")),
        _line("Component", value.get("component")),
        _line("Source workcopy", value.get("workcopy")),
        _line("Canonical source", value.get("canonical_source")),
        _line("Recovery copy", value.get("backup")),
    ]
    build = value.get("build") or {}
    if build.get("performed"):
        lines.extend(
            [
                _line("Build", build.get("build_id")),
                _line("Current Word", build.get("docx")),
                _line("Current PDF", build.get("pdf") or "not created"),
            ]
        )
        if build.get("current_update_error"):
            lines.append(_line("Current refresh", f"FAILED: {build['current_update_error']}"))
    if value.get("publish"):
        publish = value["publish"]
        label = "Delivery FAILED" if publish.get("succeeded") is False else "Delivered to"
        lines.append(_line(label, publish.get("destination")))
    return "\n".join(lines)


def _paths(value: dict[str, Any]) -> str:
    lines = ["DOCUMENT PATHS", ""]
    for label, key in (
        ("Manifest", "manifest"),
        ("Project folder", "document_root"),
        ("Current folder", "current_root"),
        ("Latest Word", "current_docx"),
        ("Latest PDF", "current_pdf"),
        ("Build history", "builds_root"),
        ("Activity log", "activity_log"),
    ):
        lines.append(_line(label, value.get(key) or "not available"))
    components = value.get("canonical_components") or []
    if components:
        lines.extend(["", "Canonical components:"])
        for item in components:
            lines.append(f"  {item['id']}: {item.get('path') or '(generated)'}")
    return "\n".join(lines)


def _status(value: dict[str, Any]) -> str:
    document = value.get("document") or {}
    release = value.get("release") or {}
    lines = [
        "DOCUMENT STATUS",
        "",
        _line("Document", document.get("id")),
        _line("Title", document.get("title")),
        _line("Revision", document.get("revision")),
        _line("Components", len(value.get("components") or [])),
        _line("Release gates ready", _yes(release.get("ready"))),
    ]
    integrity = value.get("integrity")
    if isinstance(integrity, dict):
        lines.extend(
            [
                _line("Current in sync", _yes(integrity.get("ready"))),
                _line("Full build proof", "PASS" if integrity.get("build_quality_passed") else "NOT CONFIRMED"),
            ]
        )
    lines.extend(["", "Assembly order:"])
    lines.append("  " + " -> ".join(value.get("sequence") or []))
    components = value.get("components") or []
    if components:
        lines.extend(["", "Canonical owners:"])
        for item in components:
            owner = item.get("ownership") or "generated"
            lines.append(f"  {item.get('id')}: {owner} - {item.get('source') or '(generated)'}")
    diagnostics = value.get("diagnostics") or []
    if diagnostics:
        lines.extend(["", f"Diagnostics: {len(diagnostics)} (use --json for full detail)"])
    return "\n".join(lines)


def _history(value: dict[str, Any]) -> str:
    lines = [
        "DOCUMENT ACTIVITY",
        "",
        _line("Document", value.get("document_id")),
        _line("Recorded events", value.get("total_events", 0)),
    ]
    events = value.get("events") or []
    if not events:
        lines.extend(["", "No operator activity has been recorded yet."])
        return "\n".join(lines)
    lines.append("")
    for item in events:
        timestamp = str(item.get("time_utc") or "")
        operation = str(item.get("operation") or "unknown")
        status = str(item.get("status") or "unknown").upper()
        details = item.get("details") or {}
        identity = details.get("build_id") or details.get("operation_id") or ""
        lines.append(f"  {timestamp}  {operation:<18} {status:<10} {identity}")
        if details.get("destination"):
            lines.append(f"    {details['destination']}")
        if details.get("error"):
            lines.append(f"    {details['error']}")
    return "\n".join(lines)


def _human_size(value: int) -> str:
    size = float(value)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if size < 1024 or unit == "TB":
            return f"{size:.1f} {unit}" if unit != "B" else f"{int(size)} B"
        size /= 1024
    return f"{size:.1f} TB"


def _builds(value: dict[str, Any]) -> str:
    lines = [
        "IMMUTABLE BUILDS",
        "",
        _line("Document", value.get("document_id")),
        _line("Build count", value.get("total_builds", 0)),
        _line("Storage", _human_size(int(value.get("total_size_bytes") or 0))),
        _line("Current build", value.get("current_build_id") or "none"),
    ]
    records = value.get("builds") or []
    if not records:
        lines.extend(["", "No immutable builds found."])
        return "\n".join(lines)
    lines.append("")
    for item in records:
        marker = "CURRENT" if item.get("current") else ""
        if item.get("mode") in {"component-preview", "page-furniture-preview", "lightweight-preview"}:
            proof = "PREVIEW " + str(item.get("verification_mode") or "unknown").upper()
        else:
            proof = "FULL PASS" if item.get("quality_passed") else str(item.get("verification_mode") or "unknown").upper()
        lines.append(
            f"  {item.get('build_id')}  {proof:<14} {_human_size(int(item.get('size_bytes') or 0)):>10}  {marker}"
        )
        lines.append(f"    {item.get('directory')}")
    return "\n".join(lines)


def _explain(value: dict[str, Any]) -> str:
    return "\n".join(
        [
            "COMPONENT OWNER",
            "",
            _line("Component", value.get("component")),
            _line("Type", value.get("type")),
            _line("Ownership", value.get("ownership")),
            _line("Edit this file", value.get("canonical_source") or "generated by the engine"),
            "",
            str(value.get("edit_instruction") or ""),
        ]
    )


def _workspace(value: dict[str, Any]) -> str:
    document = value.get("document") or {}
    action = value.get("next_action") or {}
    primary = value.get("primary_edit") or {}
    lines = [
        "DOCUMENT WORKSPACE",
        "",
        _line("Document", document.get("id")),
        _line("Title", document.get("title")),
        _line("Revision", document.get("revision")),
        _line("State", str(action.get("state") or "unknown").upper()),
    ]
    if primary:
        lines.extend(
            [
                "",
                "Edit first:",
                _line("  Component", primary.get("id")),
                _line("  Source type", primary.get("source_kind")),
                _line("  File", primary.get("source")),
                f"  {primary.get('instruction')}",
            ]
        )
    lines.extend(
        [
            "",
            "Next:",
            f"  {action.get('label')}: {action.get('command')}",
            f"  {action.get('reason')}",
            "",
            "Core commands:",
        ]
    )
    commands = value.get("commands") or {}
    for key in ("edit_primary", "quick_build", "full_build", "check", "review_pdf"):
        if commands.get(key):
            lines.append(f"  {key.replace('_', ' ').title()}: {commands[key]}")
    presentation = value.get("presentation") or []
    presentation_paths = {item.get("path") for item in presentation if item.get("path")}
    assembly = value.get("assembly") or []
    other = [
        item
        for item in assembly
        if item.get("editable")
        and item.get("id") != primary.get("id")
        and item.get("source") not in presentation_paths
    ]
    if other:
        lines.extend(["", "Other canonical sources:"])
        for item in other:
            indent = "  " + "  " * int(item.get("depth") or 0)
            lines.append(f"{indent}{item.get('id')} [{item.get('source_kind')}]: {item.get('source')}")
    if presentation:
        lines.extend(["", "Presentation templates:"])
        for item in presentation:
            shared = " (shared)" if item.get("shared") else ""
            lines.append(f"  {item.get('role')}{shared}: {item.get('path')}")
    return "\n".join(lines)


def _edit(value: dict[str, Any]) -> str:
    lines = [
        "EDIT TARGET" if not value.get("opened") else "EDIT TARGET OPENED",
        "",
        _line("Document", value.get("document_id")),
        _line("Selection", value.get("selection")),
        _line("Scope", value.get("scope")),
        _line("Path", value.get("path")),
        _line("Why", value.get("reason")),
    ]
    if value.get("shared_warning"):
        lines.extend(["", "WARNING: " + str(value["shared_warning"])])
    return "\n".join(lines)


def _inspect(value: dict[str, Any]) -> str:
    validation = value.get("package_validation") or {}
    body = value.get("body") or {}
    lines = [
        "WORD FILE INSPECTION",
        "",
        _line("File", value.get("path")),
        _line("Package valid", _yes(validation.get("valid"))),
        _line("Paragraphs", body.get("paragraphs", "unknown")),
        _line("Tables", body.get("tables", "unknown")),
        _line("Drawings", body.get("drawings", "unknown")),
        _line("Sections", len(value.get("sections") or [])),
    ]
    issues = validation.get("issues") or []
    if issues:
        lines.extend(["", "Package issues:", *(f"  - {item}" for item in issues)])
    return "\n".join(lines)


def _inspect_source(value: dict[str, Any]) -> str:
    lines = [
        "SOURCE INSPECTION",
        "",
        _line("File", value.get("path")),
        _line("Type", value.get("kind")),
        _line("Size", _human_size(int(value.get("size_bytes") or 0))),
        _line("SHA-256", value.get("sha256")),
    ]
    kind = value.get("kind")
    if kind == "excel":
        lines.extend(
            [
                _line("Worksheets", len(value.get("sheets") or [])),
                _line("Excel Tables", len(value.get("table_suggestions") or [])),
                _line("Formula cells", value.get("formula_count", 0)),
                "",
                "Worksheets:",
            ]
        )
        for sheet in value.get("sheets") or []:
            tables = ", ".join(item["name"] for item in sheet.get("tables") or []) or "no Excel Tables"
            lines.append(
                f"  {sheet['name']}: {sheet.get('max_row')} row(s) x {sheet.get('max_column')} column(s); {tables}"
            )
    elif kind == "pdf":
        lines.append(_line("Pages", value.get("page_count")))
        lines.extend(["", "Select explicit one-based source pages before adding this PDF to a document."])
    elif kind == "markdown_plus":
        lines.extend(
            [
                _line("Lines", value.get("line_count")),
                _line("Blocks", value.get("block_count")),
                _line("Headings", len(value.get("headings") or [])),
                _line("Insert slots", ", ".join(value.get("insert_slots") or []) or "none"),
            ]
        )
    elif kind == "word":
        body = value.get("body") or {}
        lines.extend(
            [
                _line("Paragraphs", body.get("paragraphs")),
                _line("Tables", body.get("tables")),
                _line("Drawings", body.get("drawings")),
                _line("Sections", len(value.get("sections") or [])),
            ]
        )
    elif kind == "visio":
        lines.extend([_line("Pages", value.get("page_count")), "", "Visio pages:"])
        for page in value.get("pages") or []:
            lines.append(f"  {page.get('name') or page.get('universal_name') or page.get('id')}")
    elif kind == "figure":
        pixels = value.get("pixels") or {}
        lines.append(_line("Pixels", f"{pixels.get('width')} x {pixels.get('height')}"))
    return "\n".join(lines)


def _schemas(value: dict[str, Any]) -> str:
    if value.get("current"):
        status = "CURRENT"
    elif value.get("check_only"):
        status = "OUT OF DATE"
    else:
        status = "REFRESHED"
    lines = [
        "MANIFEST SCHEMAS " + status,
        "",
        _line("Folder", value.get("root")),
    ]
    for item in value.get("files") or []:
        state = "updated" if item.get("updated") else "current" if item.get("current") else "out of date"
        lines.append(f"  {item.get('name')}: {state}")
    return "\n".join(lines)


def _new(value: dict[str, Any]) -> str:
    label = "PROJECT CREATED" if value.get("operation") == "new-project" else "DOCUMENT STARTER CREATED"
    lines = [label, ""]
    if value.get("project_id"):
        lines.append(_line("Project", value.get("project_id")))
    if value.get("document_id"):
        lines.append(_line("Document", value.get("document_id")))
    lines.extend(
        [
            _line("Folder", value.get("root")),
            _line("Manifest", value.get("manifest")),
        ]
    )
    if value.get("body"):
        lines.append(_line("Write prose here", value.get("body")))
        lines.append(_line("Build ready", _yes(value.get("build_ready"))))
    steps = value.get("next_steps") or []
    if steps:
        lines.extend(["", "Next:", *(f"  {index}. {step}" for index, step in enumerate(steps, 1))])
    return "\n".join(lines)


def _added(value: dict[str, Any]) -> str:
    placement = value.get("placement") or {}
    where = (
        f"{placement.get('parent')} at {placement.get('slot')}"
        if placement.get("parent")
        else f"region {placement.get('region')}"
    )
    lines = [
        "COMPONENT ADDED",
        "",
        _line("Document", value.get("document_id")),
        _line("Component", value.get("component_id")),
        _line("Type", value.get("component_type")),
        _line("Placement", where),
        _line("Source", value.get("canonical_source") or "generated"),
        _line("Recovery", value.get("recovery")),
    ]
    if value.get("marker_appended"):
        lines.append(_line("Markdown marker", "appended deliberately"))
    steps = value.get("next_steps") or []
    if steps:
        lines.extend(["", "Next:", *(f"  {index}. {step}" for index, step in enumerate(steps, 1))])
    return "\n".join(lines)


def _opened(value: dict[str, Any]) -> str:
    return "\n".join(["OPENED", "", _line("Type", value.get("kind")), _line("Path", value.get("path"))])


def _generic(value: Any) -> str:
    if isinstance(value, dict):
        lines = []
        for key, item in value.items():
            if isinstance(item, (str, int, float, bool)) or item is None:
                lines.append(_line(key.replace("_", " ").title(), item))
        if lines:
            return "\n".join(lines)
    return str(value)


def render_result(command: str, value: dict[str, Any]) -> str:
    renderers = {
        "documents": _documents,
        "schema": _schemas,
        "inspect-source": _inspect_source,
        "new": _new,
        "add": _added,
        "open": _opened,
        "doctor": _doctor,
        "audit": _audit,
        "check": _audit,
        "build": _build,
        "preview": _build,
        "revise": _revision,
        "replace": _revision,
        "publish": _publish,
        "deliver": _publish,
        "restore-current": _restore,
        "diff-workcopy": _workcopy_diff,
        "compare": _workcopy_diff,
        "adopt": _adoption,
        "paths": _paths,
        "history": _history,
        "builds": _builds,
        "validate": _status,
        "status": _status,
        "explain": _explain,
        "workspace": _workspace,
        "edit": _edit,
        "inspect-docx": _inspect,
        "release": _build,
    }
    return renderers.get(command, _generic)(value)
