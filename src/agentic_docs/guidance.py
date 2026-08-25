from __future__ import annotations

import os
from pathlib import Path

from .errors import DocumentSystemError


def _component_instruction(item) -> str:
    path = item.source_path
    if path is None:
        return "Generated during composition; change its declaration or selected template instead."
    suffix = path.suffix.lower()
    component_type = item.declaration.type.value
    ownership = item.declaration.ownership.value if item.declaration.ownership else None
    if suffix in {".md", ".markdown"}:
        return "Edit the Markdown+ source; keep any :::insert markers, then rebuild."
    if suffix == ".docx":
        if component_type == "cover":
            return "Edit this native Word cover template; keep its tagged fields, then rebuild."
        return "Edit this canonical Word source in Microsoft Word, save it, then rebuild."
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        return "Edit the declared Excel Table or range without changing its selector, then rebuild."
    if component_type == "diagram":
        return "Edit the native drawing, export and review a new rendition, then accept that rendition."
    if component_type in {"figure", "pdf_pages"} or ownership == "snapshot":
        return "Replace this reviewed source explicitly when a new approved asset is ready."
    return "Edit the declared canonical source, then rebuild."


def _source_kind(path: Path | None) -> str:
    if path is None:
        return "generated"
    suffix = path.suffix.lower()
    return {
        ".md": "markdown-plus",
        ".markdown": "markdown-plus",
        ".docx": "word",
        ".xlsx": "excel",
        ".xlsm": "excel",
        ".xls": "excel",
        ".pdf": "pdf",
        ".png": "image",
        ".jpg": "image",
        ".jpeg": "image",
        ".svg": "drawing",
        ".vsdx": "drawing",
        ".drawio": "drawing",
    }.get(suffix, suffix.lstrip(".") or "file")


def _assembly_entries(resolved) -> list[dict]:
    entries: list[dict] = []

    def visit(component_id: str, *, region: str, parent: str | None, slot: str | None, depth: int) -> None:
        item = resolved.components[component_id]
        entries.append(
            {
                "id": component_id,
                "type": item.declaration.type.value,
                "ownership": item.declaration.ownership.value if item.declaration.ownership else None,
                "region": region,
                "parent": parent,
                "slot": slot,
                "depth": depth,
                "source": str(item.source_path) if item.source_path else None,
                "source_kind": _source_kind(item.source_path),
                "editable": item.source_path is not None,
                "instruction": _component_instruction(item),
            }
        )
        for slot_name, children in item.declaration.slots.items():
            for child_id in children:
                visit(child_id, region=region, parent=component_id, slot=slot_name, depth=depth + 1)

    for group in resolved.manifest.sequence:
        for component_id in group.items:
            visit(component_id, region=group.region, parent=None, slot=None, depth=0)
    return entries


def _primary_component(entries: list[dict]) -> dict | None:
    candidates = []
    for index, entry in enumerate(entries):
        if not entry["editable"] or entry["type"] in {"cover", "toc", "page_break"}:
            continue
        score = 0
        if entry["type"] == "document":
            score += 100
        if entry["source_kind"] == "markdown-plus":
            score += 40
        elif entry["source_kind"] == "word":
            score += 35
        if entry["parent"] is None:
            score += 10
        candidates.append((score, -index, entry))
    return max(candidates, default=(None, None, None), key=lambda item: (item[0], item[1]))[2]


def _path_scope(resolved, path: Path) -> str:
    checks = (
        (resolved.manifest_path.parent, "document"),
        (resolved.kit_path.parent, "shared-kit"),
        (resolved.profile_path.parent, "shared-profile"),
        (resolved.project_path.parent, "project"),
    )
    for root, label in checks:
        try:
            path.resolve().relative_to(root.resolve())
            return label
        except ValueError:
            continue
    return "external"


def _presentation_entries(resolved) -> list[dict]:
    entries = []
    for role, path in sorted(resolved.presentation_paths.items()):
        if path is None:
            continue
        entries.append(
            {
                "role": role,
                "path": str(path),
                "source_kind": _source_kind(path),
                "scope": _path_scope(resolved, path),
                "shared": _path_scope(resolved, path).startswith("shared-"),
            }
        )
    entries.append(
        {
            "role": "profile.shell",
            "path": str(resolved.shell_path),
            "source_kind": "word",
            "scope": "shared-profile",
            "shared": True,
        }
    )
    return entries


def _next_action(document_id: str, integrity: dict) -> dict:
    wrapper = ".\\Document-System.cmd"
    if not integrity.get("build_id"):
        return {
            "state": "not-built",
            "label": "Create a drafting build",
            "reason": "No current immutable build exists yet.",
            "command": f"{wrapper} build {document_id} --quick",
        }
    if not integrity.get("sources_current"):
        return {
            "state": "source-changed",
            "label": "Rebuild the edited canonicals",
            "reason": "One or more canonical inputs changed after the current build.",
            "command": f"{wrapper} build {document_id} --quick",
        }
    if not integrity.get("current_integrity"):
        return {
            "state": "current-output-drift",
            "label": "Restore the verified current pair",
            "reason": "The convenient current Word/PDF pair no longer matches its immutable build.",
            "command": f"{wrapper} restore-current {document_id}",
        }
    if not integrity.get("build_quality_passed") or integrity.get("verification_mode") != "full":
        return {
            "state": "full-proof-needed",
            "label": "Create the full visual proof",
            "reason": "The current build does not carry a complete full-page verification pass.",
            "command": f"{wrapper} build {document_id}",
        }
    return {
        "state": "ready",
        "label": "Edit or review",
        "reason": "Canonical inputs, immutable artifacts and the current pair are aligned.",
        "command": f"{wrapper} open {document_id} --pdf",
    }


def document_workspace(resolved, integrity: dict) -> dict:
    entries = _assembly_entries(resolved)
    primary = _primary_component(entries)
    document_id = resolved.manifest.id
    current_root = resolved.system_root / "current" / document_id
    current_docx = current_root / f"{resolved.manifest.outputs.basename}.docx"
    current_pdf = current_root / f"{resolved.manifest.outputs.basename}.pdf"
    commands = {
        "workspace": f".\\Document-System.cmd workspace {document_id}",
        "quick_build": f".\\Document-System.cmd build {document_id} --quick",
        "full_build": f".\\Document-System.cmd build {document_id}",
        "check": f".\\Document-System.cmd check {document_id}",
        "review_pdf": f".\\Document-System.cmd open {document_id} --pdf",
    }
    if primary:
        commands["edit_primary"] = f".\\Document-System.cmd edit {document_id}"
        commands["locate_primary"] = f".\\Document-System.cmd edit {document_id} --show"
    return {
        "schema": "agentic-document-workspace/v1",
        "document": {
            "id": document_id,
            "title": resolved.manifest.metadata.title,
            "revision": resolved.manifest.metadata.revision,
            "manifest": str(resolved.manifest_path),
        },
        "next_action": _next_action(document_id, integrity),
        "primary_edit": primary,
        "assembly": entries,
        "presentation": _presentation_entries(resolved),
        "outputs": {
            "current_word": str(current_docx) if current_docx.is_file() else None,
            "current_pdf": str(current_pdf) if current_pdf.is_file() else None,
            "builds": str(resolved.system_root / "builds" / document_id),
        },
        "integrity": {
            key: integrity.get(key)
            for key in (
                "build_id",
                "ready",
                "sources_current",
                "current_integrity",
                "build_quality_passed",
                "verification_mode",
            )
        },
        "release_gates": {
            "required": list(resolved.profile.release_gates),
            "states": {key: value.value for key, value in resolved.manifest.release.gates.items()},
            "open": [
                gate
                for gate in resolved.profile.release_gates
                if resolved.manifest.release.gates.get(gate, "open") == "open"
            ],
        },
        "commands": commands,
    }


def _presentation_target(resolved, selector: str) -> tuple[str, Path]:
    if selector == "shell":
        return "profile.shell", resolved.shell_path
    aliases = {
        "cover": ["cover"],
        "styles": ["styles"],
        "header": [key for key in resolved.presentation_paths if ".header." in key],
        "footer": [key for key in resolved.presentation_paths if ".footer." in key],
    }
    keys = aliases.get(selector, [selector])
    matches = [(key, resolved.presentation_paths.get(key)) for key in keys]
    matches = [(key, path) for key, path in matches if path is not None]
    unique_paths = {path.resolve() for _, path in matches}
    if not matches:
        available = ", ".join(sorted(key for key, value in resolved.presentation_paths.items() if value))
        raise DocumentSystemError(f"Presentation source {selector!r} is not selected. Available: {available}, shell")
    if len(unique_paths) > 1:
        available = ", ".join(key for key, _ in matches)
        raise DocumentSystemError(f"Presentation alias {selector!r} is ambiguous; choose one exact role: {available}")
    return matches[0]


def edit_target(
    resolved,
    *,
    component_id: str | None = None,
    presentation: str | None = None,
    open_file: bool = True,
) -> dict:
    if component_id and presentation:
        raise DocumentSystemError("Choose a component or presentation source, not both")
    if presentation:
        role, target = _presentation_target(resolved, presentation)
        kind = "presentation"
        scope = _path_scope(resolved, target)
        reason = f"Selected presentation role {role}."
        selection = role
    else:
        entries = _assembly_entries(resolved)
        if component_id:
            match = next((entry for entry in entries if entry["id"] == component_id), None)
            if match is None:
                raise DocumentSystemError(f"Unknown component {component_id!r}")
        else:
            match = _primary_component(entries)
            if match is None:
                raise DocumentSystemError("This document has no editable canonical content component")
        if not match["source"]:
            raise DocumentSystemError(f"Component {match['id']!r} is generated and has no canonical source")
        target = Path(match["source"])
        kind = "canonical-component"
        scope = _path_scope(resolved, target)
        reason = "Selected as the primary editable content component." if not component_id else "Selected explicitly."
        selection = match["id"]
    if not target.is_file():
        raise DocumentSystemError(f"Selected edit target does not exist: {target}")
    if open_file:
        os.startfile(str(target))
    return {
        "schema": "agentic-edit-target/v1",
        "document_id": resolved.manifest.id,
        "selection": selection,
        "kind": kind,
        "scope": scope,
        "path": str(target),
        "reason": reason,
        "opened": bool(open_file),
        "shared_warning": (
            "This is a shared template; changes can affect multiple documents."
            if scope.startswith("shared-")
            else None
        ),
    }
