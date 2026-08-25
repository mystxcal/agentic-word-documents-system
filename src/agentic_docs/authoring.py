from __future__ import annotations

import json
import os
import re
import shutil
import time
import uuid
from datetime import date
from datetime import datetime, timezone
from pathlib import Path

from .diagnostics import DiagnosticBag
from .errors import DocumentSystemError
from .jsonc import load_jsonc
from .jsonc_edit import add_component as add_component_to_jsonc, replace_values
from .model import DocumentManifest, KitManifest, ProfileManifest, ProjectManifest
from .resolver import file_hash
from .sources.markdown_ast import parse_document
from .word.ooxml import find_sdts
from .word.package import DOCUMENT_PART, DocxPackage


SAFE_IDENTIFIER = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]*$")
WINDOWS_UNSAFE = re.compile(r'[<>:"/\\|?*]+')


def _safe_identifier(value: str, label: str) -> str:
    value = value.strip()
    if not SAFE_IDENTIFIER.fullmatch(value):
        raise DocumentSystemError(
            f"{label} must contain only letters, numbers, dot, underscore, or hyphen and cannot start with punctuation"
        )
    return value


def _write_json(path: Path, value: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        raise DocumentSystemError(f"Refusing to replace an existing file: {path}")
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(value, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    os.replace(temporary, path)


def _atomic_write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.{uuid.uuid4().hex}.tmp")
    try:
        temporary.write_text(text, encoding="utf-8")
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def _staging_directory(target: Path) -> Path:
    target.parent.mkdir(parents=True, exist_ok=True)
    staging = target.parent / f".{target.name}.creating-{uuid.uuid4().hex}"
    staging.mkdir()
    return staging


def _promote_directory(staging: Path, target: Path) -> None:
    for attempt in range(20):
        try:
            os.replace(staging, target)
            return
        except PermissionError:
            if attempt == 19:
                raise
            time.sleep(0.05)


def create_project(
    system_root: Path,
    project_id: str,
    *,
    name: str,
    description: str | None = None,
    metadata: dict[str, str] | None = None,
) -> dict:
    project_id = _safe_identifier(project_id, "Project id")
    target_root = Path(system_root).resolve() / "projects" / project_id
    if target_root.exists():
        raise DocumentSystemError(f"Project folder already exists: {target_root}")
    payload = {
        "$schema": "../../schemas/agentic-project-v2.schema.json",
        "schema": "agentic-project/v2",
        "id": project_id,
        "name": name.strip(),
        "description": description.strip() if description else None,
        "metadata": dict(metadata or {}),
        "sources": {},
    }
    ProjectManifest.model_validate(payload)
    staging = _staging_directory(target_root)
    promoted = False
    try:
        _write_json(staging / "project.jsonc", payload)
        (staging / "documents").mkdir()
        (staging / "sources").mkdir()
        _promote_directory(staging, target_root)
        promoted = True
    except Exception:
        shutil.rmtree(staging, ignore_errors=True)
        if promoted and target_root.exists():
            shutil.rmtree(target_root, ignore_errors=True)
        raise
    manifest_path = target_root / "project.jsonc"
    return {
        "schema": "agentic-authoring-result/v1",
        "operation": "new-project",
        "project_id": project_id,
        "root": str(target_root),
        "manifest": str(manifest_path),
        "created": [str(manifest_path), str(target_root / "documents"), str(target_root / "sources")],
    }


def _safe_basename(value: str) -> str:
    result = WINDOWS_UNSAFE.sub(" - ", value).strip().rstrip(".")
    result = re.sub(r"\s+", " ", result)
    if not result:
        raise DocumentSystemError("Could not derive a Windows-safe output filename")
    return result


def create_document(
    system_root: Path,
    document_id: str,
    *,
    project_id: str,
    title: str,
    document_type: str = "Document",
    revision: str = "draft",
    profile_id: str = "plain",
    kit_id: str = "studio",
    short_title: str | None = None,
    number: str | None = None,
    revision_display: str | None = None,
    document_date: str | None = None,
    prepared_by: str | None = None,
    basename: str | None = None,
    styles: str = "studio",
    word_source: Path | None = None,
    source_tag: str | None = None,
    allow_untagged: bool = False,
    preserve_sections: bool = False,
    use_source_styles: bool = False,
    preserve_source_layout: bool = False,
) -> dict:
    root = Path(system_root).resolve()
    document_id = _safe_identifier(document_id, "Document id")
    project_id = _safe_identifier(project_id, "Project id")
    profile_id = _safe_identifier(profile_id, "Profile id")
    kit_id = _safe_identifier(kit_id, "Kit id")
    project_path = root / "projects" / project_id / "project.jsonc"
    profile_path = root / "profiles" / profile_id / "profile.jsonc"
    kit_path = root / "kits" / kit_id / "kit.jsonc"
    for label, path in (("project", project_path), ("profile", profile_path), ("kit", kit_path)):
        if not path.is_file():
            raise DocumentSystemError(f"Selected {label} does not exist: {path}")
    ProjectManifest.model_validate(load_jsonc(project_path))
    profile = ProfileManifest.model_validate(load_jsonc(profile_path))
    kit = KitManifest.model_validate(load_jsonc(kit_path))
    if preserve_source_layout:
        preserve_sections = True
        use_source_styles = True
    if not use_source_styles and styles not in kit.components:
        available = ", ".join(sorted(kit.components)) or "none"
        raise DocumentSystemError(f"Kit {kit_id!r} has no style component {styles!r}; available: {available}")
    selected_word_source = None
    if word_source is not None:
        selected_word_source = Path(word_source).expanduser().resolve()
        if not selected_word_source.is_file() or selected_word_source.suffix.lower() != ".docx":
            raise DocumentSystemError(f"--word-source must select an existing DOCX file: {selected_word_source}")
        if not source_tag and not allow_untagged:
            raise DocumentSystemError(
                "A Word source requires --source-tag for precise extraction or explicit --allow-untagged for whole-body intake"
            )
    elif source_tag or allow_untagged or preserve_sections or use_source_styles or preserve_source_layout:
        raise DocumentSystemError(
            "--source-tag, --allow-untagged, --preserve-sections, --use-source-styles, and "
            "--preserve-source-layout require --word-source"
        )
    target_root = root / "projects" / project_id / "documents" / document_id
    if target_root.exists():
        raise DocumentSystemError(f"Document folder already exists: {target_root}")
    selected_date = document_date or date.today().isoformat()
    output_name = basename.strip() if basename else _safe_basename(title)
    payload = {
        "$schema": "../../../../schemas/agentic-document-v2.schema.json",
        "schema": "agentic-document/v2",
        "id": document_id,
        "project": project_id,
        "profile": profile_id,
        "kit": kit_id,
        "metadata": {
            "type": document_type.strip(),
            "title": title.strip(),
            "short_title": (short_title or title).strip(),
            "number": number.strip() if number else None,
            "revision": revision.strip(),
            "revision_display": revision_display.strip() if revision_display else None,
            "date": selected_date,
            "prepared_by": prepared_by.strip() if prepared_by else None,
        },
        "presentation": {
            "styles": "document:body" if use_source_styles else f"kit:{styles}",
            "cover": None,
            "page_regions": {
                "main": {
                    "layout_mode": "preserve" if preserve_source_layout else "managed",
                    "header": None,
                    "footer": None,
                    "numbering": (
                        None
                        if preserve_source_layout
                        else {"style": "arabic", "start": 1, "page_count_scope": "region"}
                    ),
                }
            },
        },
        "sequence": [{"region": "main", "items": ["body"]}],
        "components": {
            "body": (
                {
                    "type": "document",
                    "ownership": "word_fragment",
                    "source": "content/body.docx",
                    "source_tag": source_tag,
                    "allow_untagged": allow_untagged,
                    "options": {
                        "preserve_sections": preserve_sections,
                        "whole_document": preserve_source_layout,
                    },
                }
                if selected_word_source
                else {
                    "type": "document",
                    "ownership": "source",
                    "source": "content/body.md",
                }
            )
        },
        "outputs": {"basename": output_name},
        "release": {
            "gates": {gate: "open" for gate in profile.release_gates},
            "note": None,
        },
    }
    DocumentManifest.model_validate(payload)
    staging = _staging_directory(target_root)
    promoted = False
    try:
        _write_json(staging / "document.jsonc", payload)
        body_stage = staging / "content" / ("body.docx" if selected_word_source else "body.md")
        body_stage.parent.mkdir(parents=True, exist_ok=True)
        if selected_word_source:
            shutil.copy2(selected_word_source, body_stage)
            if file_hash(body_stage) != file_hash(selected_word_source):
                raise DocumentSystemError("Word source failed canonical copy verification")
        else:
            body_stage.write_text(f"# {title.strip()}\n", encoding="utf-8")
        (staging / "presentation").mkdir()
        _promote_directory(staging, target_root)
        promoted = True
        from .resolver import resolve_document

        resolve_document(target_root / "document.jsonc", DiagnosticBag())
    except Exception:
        shutil.rmtree(staging, ignore_errors=True)
        if promoted and target_root.exists():
            shutil.rmtree(target_root, ignore_errors=True)
        raise
    manifest_path = target_root / "document.jsonc"
    body_path = target_root / "content" / ("body.docx" if selected_word_source else "body.md")
    required_bindings = [
        tag
        for tag, binding in profile.field_bindings.items()
        if getattr(binding, "required", False)
    ]
    build_ready = not required_bindings and not profile.region_starts and profile.main_start_tag is None
    next_steps = [
        f"Edit content/{body_path.name}",
        f"Run: Document-System.cmd workspace {document_id}",
        f"Run: Document-System.cmd validate {document_id}",
    ]
    if not build_ready:
        next_steps.append(
            "The selected profile requires presentation donors or layout markers; attach the appropriate cover/header/footer template before building."
        )
    else:
        next_steps.append(f"Run: Document-System.cmd build {document_id} --quick")
    return {
        "schema": "agentic-authoring-result/v1",
        "operation": "new-document",
        "document_id": document_id,
        "root": str(target_root),
        "manifest": str(manifest_path),
        "body": str(body_path),
        "body_ownership": "word_fragment" if selected_word_source else "source",
        "word_source_sha256": file_hash(body_path) if selected_word_source else None,
        "source_layout_preserved": bool(selected_word_source and preserve_source_layout),
        "profile": profile_id,
        "build_ready": build_ready,
        "required_profile_bindings": required_bindings,
        "next_steps": next_steps,
        "created": [str(manifest_path), str(body_path), str(target_root / "presentation")],
    }


def open_document(
    resolved,
    *,
    component_id: str | None = None,
    pdf: bool = False,
    manifest: bool = False,
    folder: bool = False,
) -> dict:
    selections = sum(bool(item) for item in (component_id, pdf, manifest, folder))
    if selections > 1:
        raise DocumentSystemError("Choose only one of --component, --pdf, --manifest, or --folder")
    current_root = resolved.system_root / "current" / resolved.manifest.id
    if component_id:
        if component_id not in resolved.components:
            raise DocumentSystemError(f"Unknown component {component_id!r}")
        target = resolved.components[component_id].source_path
        if target is None:
            raise DocumentSystemError(f"Component {component_id!r} is generated and has no canonical file to open")
        kind = "canonical-component"
    elif manifest:
        target = resolved.manifest_path
        kind = "manifest"
    elif folder:
        target = resolved.manifest_path.parent
        kind = "document-folder"
    elif pdf:
        target = current_root / f"{resolved.manifest.outputs.basename}.pdf"
        kind = "current-pdf"
    else:
        target = current_root / f"{resolved.manifest.outputs.basename}.docx"
        kind = "current-word"
    if not target.exists():
        raise DocumentSystemError(f"Requested {kind} does not exist: {target}")
    os.startfile(str(target))
    return {
        "schema": "agentic-open-result/v1",
        "document_id": resolved.manifest.id,
        "kind": kind,
        "path": str(target),
        "opened": True,
    }


def parse_pages(value: str) -> list[int | list[int]]:
    result: list[int | list[int]] = []
    for raw in value.split(","):
        item = raw.strip()
        if not item:
            continue
        if "-" in item:
            parts = [part.strip() for part in item.split("-", 1)]
            if len(parts) != 2 or not all(part.isdigit() for part in parts):
                raise DocumentSystemError(f"Invalid PDF page selector {item!r}; use values such as 1,3-5,8")
            start, end = (int(part) for part in parts)
            if start < 1 or end < start:
                raise DocumentSystemError(f"Invalid PDF page range {item!r}")
            result.append([start, end])
        elif item.isdigit() and int(item) >= 1:
            result.append(int(item))
        else:
            raise DocumentSystemError(f"Invalid PDF page selector {item!r}; use values such as 1,3-5,8")
    if not result:
        raise DocumentSystemError("At least one PDF page must be selected")
    return result


def _relative_selector(resolved, source: Path, component_id: str) -> tuple[str, Path | None, Path | None]:
    """Return manifest selector, optional copy source, and optional new target."""

    source = source.resolve()
    document_root = resolved.manifest_path.parent.resolve()
    project_root = resolved.project_path.parent.resolve()
    project_sources = project_root / "sources"
    try:
        return source.relative_to(document_root).as_posix(), None, None
    except ValueError:
        pass
    try:
        relative = source.relative_to(project_sources)
        return f"project-file:sources/{relative.as_posix()}", None, None
    except ValueError:
        pass

    destination_root = document_root / "content" / "sources"
    candidate = destination_root / source.name
    if candidate.exists():
        if file_hash(candidate) == file_hash(source):
            return candidate.relative_to(document_root).as_posix(), None, None
        candidate = destination_root / f"{source.stem}-{file_hash(source)[:8].lower()}{source.suffix.lower()}"
        if candidate.exists() and file_hash(candidate) != file_hash(source):
            raise DocumentSystemError(f"Could not choose an unambiguous canonical target for {source}")
        if candidate.exists():
            return candidate.relative_to(document_root).as_posix(), None, None
    return candidate.relative_to(document_root).as_posix(), source, candidate


def _source_plan(resolved, kind: str, component_id: str, source: Path | None, title: str | None):
    expected = {
        "prose": {".md", ".markdown"},
        "word-fragment": {".docx"},
        "table": {".xlsx", ".xlsm", ".xltx", ".xltm"},
        "pdf-pages": {".pdf"},
        "figure": {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff"},
        # A diagram's canonical source is intentionally format-neutral. Visio,
        # draw.io, SVG, CAD exports, and future native drawing formats all remain
        # valid; Word receives the separately reviewed raster rendition.
        "diagram": None,
    }
    if kind == "page-break":
        if source is not None:
            raise DocumentSystemError("A page-break component does not use a source file")
        return None, None, None, None
    if kind == "prose" and source is None:
        if not title or not title.strip():
            raise DocumentSystemError("A new prose component requires --title when no existing Markdown source is supplied")
        target = resolved.manifest_path.parent / "content" / f"{component_id}.md"
        if target.exists():
            raise DocumentSystemError(f"Canonical prose source already exists: {target}")
        return target.relative_to(resolved.manifest_path.parent).as_posix(), None, target, f"# {title.strip()}\n"
    if source is None:
        raise DocumentSystemError(f"An existing source file is required for a {kind} component")
    source = Path(source).expanduser().resolve()
    if not source.is_file():
        raise DocumentSystemError(f"Component source does not exist: {source}")
    allowed_extensions = expected[kind]
    if allowed_extensions is not None and source.suffix.lower() not in allowed_extensions:
        allowed = ", ".join(sorted(allowed_extensions))
        raise DocumentSystemError(f"A {kind} component requires one of: {allowed}")
    selector, copy_source, target = _relative_selector(resolved, source, component_id)
    return selector, copy_source, target, None


def _placement_ready(resolved, parent_id: str, slot_name: str, *, append_marker: bool):
    if parent_id not in resolved.components:
        raise DocumentSystemError(f"Parent component {parent_id!r} does not exist")
    parent = resolved.components[parent_id]
    if parent.declaration.type.value != "document" or parent.source_path is None:
        raise DocumentSystemError("Nested components require a Markdown or Word document parent")
    suffix = parent.source_path.suffix.lower()
    if suffix in {".md", ".markdown"}:
        original = parent.source_path.read_text(encoding="utf-8-sig")
        parsed = parse_document(original, DiagnosticBag(), str(parent.source_path), strict=False)
        count = parsed.slot_names.count(slot_name)
        if count > 1:
            raise DocumentSystemError(f"Markdown insertion marker {slot_name!r} occurs {count} times; exactly one is required")
        if count == 1:
            return parent.source_path, original, original, False
        if not append_marker:
            raise DocumentSystemError(
                f"Markdown parent has no ':::insert {slot_name}' marker. Place it where the component belongs, "
                "or rerun with --append-marker to deliberately place it at the end."
            )
        updated = original.rstrip() + f"\n\n:::insert {slot_name}\n"
        return parent.source_path, original, updated, True
    if suffix == ".docx":
        package = DocxPackage(parent.source_path)
        matches = find_sdts(package.xml(DOCUMENT_PART), slot_name)
        if len(matches) != 1:
            raise DocumentSystemError(
                f"Word parent must already contain exactly one content control tagged {slot_name!r}; found {len(matches)}"
            )
        return None, None, None, False
    raise DocumentSystemError("Nested components require a Markdown or DOCX parent")


def add_document_component(
    resolved,
    *,
    kind: str,
    component_id: str,
    source: Path | None = None,
    rendition: Path | None = None,
    title: str | None = None,
    caption: str | bool | None = None,
    alt_text: str | None = None,
    parent_component: str | None = None,
    slot_name: str | None = None,
    region: str | None = None,
    after_component: str | None = None,
    append_marker: bool = False,
    source_tag: str | None = None,
    allow_untagged: bool = False,
    preserve_sections: bool = False,
    excel_table: str | None = None,
    excel_sheet: str | None = None,
    excel_range: str | None = None,
    table_style_role: str = "technical",
    formula_policy: str = "cached_values",
    pages: list[int | list[int]] | None = None,
    dpi: int = 150,
    width_inches: float | None = None,
    alignment: str = "center",
    page_break_between: bool = True,
) -> dict:
    component_id = _safe_identifier(component_id, "Component id")
    if bool(parent_component) == bool(region):
        raise DocumentSystemError("Choose exactly one placement: --into a parent component or --region for top-level placement")
    if after_component and not region:
        raise DocumentSystemError("--after can only be used with top-level --region placement")
    if append_marker and not parent_component:
        raise DocumentSystemError("--append-marker can only be used with --into")
    selector, copy_source, new_source_target, generated_text = _source_plan(
        resolved, kind, component_id, source, title
    )
    rendition_selector = None
    rendition_copy_source = None
    new_rendition_target = None
    if kind == "diagram":
        if rendition is None:
            raise DocumentSystemError("A diagram component requires --rendition with a reviewed raster export")
        rendition = Path(rendition).expanduser().resolve()
        if not rendition.is_file():
            raise DocumentSystemError(f"Diagram rendition does not exist: {rendition}")
        if rendition.suffix.lower() not in {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff"}:
            raise DocumentSystemError("A diagram rendition must be a reviewed PNG, JPEG, BMP, GIF, or TIFF image")
        rendition_selector, rendition_copy_source, new_rendition_target = _relative_selector(
            resolved, rendition, component_id
        )
    selected_slot = slot_name or component_id
    marker_path = marker_original = marker_updated = None
    marker_changed = False
    if parent_component:
        marker_path, marker_original, marker_updated, marker_changed = _placement_ready(
            resolved, parent_component, selected_slot, append_marker=append_marker
        )

    declaration: dict = {"type": kind.replace("-", "_")}
    if kind != "page-break":
        declaration.update({"ownership": "source", "source": selector})
    if kind == "prose":
        declaration["type"] = "document"
    elif kind == "word-fragment":
        declaration.update(
            {
                "type": "document",
                "ownership": "word_fragment",
                "source_tag": source_tag,
                "allow_untagged": allow_untagged,
                "options": {"preserve_sections": preserve_sections},
            }
        )
    elif kind == "table":
        if bool(excel_table) == bool(excel_range):
            raise DocumentSystemError("A table requires exactly one locator: --table, or --sheet together with --range")
        if excel_range and not excel_sheet:
            raise DocumentSystemError("--range requires --sheet")
        locator = {"table": excel_table} if excel_table else {"sheet": excel_sheet, "range": excel_range}
        declaration["options"] = {
            "locator": locator,
            "view": {"columns": "*", "style_role": table_style_role},
            "formula_policy": formula_policy,
        }
    elif kind == "pdf-pages":
        if not pages:
            raise DocumentSystemError("A pdf-pages component requires explicit --pages")
        declaration.update(
            {
                "type": "pdf_pages",
                "options": {
                    "pages": pages,
                    "dpi": dpi,
                    "image_width_inches": width_inches,
                    "alignment": alignment,
                    "page_break_between": page_break_between,
                },
            }
        )
    elif kind == "figure":
        if not alt_text:
            raise DocumentSystemError("A figure component requires concise --alt-text")
        declaration.update(
            {
                "ownership": "snapshot",
                "alignment": alignment,
                "alt_text": alt_text,
                "options": {"width_inches": width_inches},
            }
        )
    elif kind == "diagram":
        if not alt_text:
            raise DocumentSystemError("A diagram component requires concise --alt-text")
        declaration.update(
            {
                "ownership": "source",
                "alignment": alignment,
                "alt_text": alt_text,
                "options": {
                    "rendition": rendition_selector,
                    "rendition_of_sha256": file_hash(Path(source).expanduser().resolve()),
                    "width_inches": width_inches,
                },
            }
        )
    elif kind == "page-break":
        declaration = {"type": "page_break"}
    if title:
        declaration["title"] = title
    if caption is not None:
        declaration["caption"] = caption

    original_manifest = resolved.manifest_path.read_text(encoding="utf-8-sig")
    candidate_manifest = add_component_to_jsonc(
        original_manifest,
        component_id,
        declaration,
        parent_component=parent_component,
        slot_name=selected_slot if parent_component else None,
        region=region,
        after_component=after_component,
    )
    DocumentManifest.model_validate(load_jsonc_text(candidate_manifest, resolved.manifest_path))

    operation_id = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S_%fZ")
    history = resolved.manifest_path.parent / ".history" / "authoring" / f"{operation_id}-add-{component_id}"
    history.mkdir(parents=True, exist_ok=False)
    manifest_backup = history / "document.jsonc"
    manifest_backup.write_text(original_manifest, encoding="utf-8")
    if marker_changed and marker_path is not None and marker_original is not None:
        (history / marker_path.name).write_text(marker_original, encoding="utf-8")

    created_targets: list[Path] = []
    updated_resolved = None
    try:
        copy_plans = [
            (copy_source, new_source_target, generated_text),
            (rendition_copy_source, new_rendition_target, None),
        ]
        for planned_source, planned_target, planned_text in copy_plans:
            if planned_target is None:
                continue
            planned_target.parent.mkdir(parents=True, exist_ok=True)
            temporary_source = planned_target.with_name(f".{planned_target.name}.{uuid.uuid4().hex}.tmp")
            try:
                if planned_text is not None:
                    temporary_source.write_text(planned_text, encoding="utf-8")
                else:
                    shutil.copy2(planned_source, temporary_source)
                os.replace(temporary_source, planned_target)
                created_targets.append(planned_target)
            finally:
                temporary_source.unlink(missing_ok=True)
        if marker_changed and marker_path is not None and marker_updated is not None:
            temporary_marker = marker_path.with_suffix(marker_path.suffix + ".tmp")
            temporary_marker.write_text(marker_updated, encoding="utf-8")
            os.replace(temporary_marker, marker_path)
        temporary_manifest = resolved.manifest_path.with_suffix(".jsonc.tmp")
        temporary_manifest.write_text(candidate_manifest, encoding="utf-8")
        os.replace(temporary_manifest, resolved.manifest_path)
        from .resolver import resolve_document

        updated_resolved = resolve_document(resolved.manifest_path, DiagnosticBag())

        canonical_path = (
            updated_resolved.components[component_id].source_path
            if selector and component_id in updated_resolved.components
            else None
        )
        canonical_related = (
            updated_resolved.components[component_id].related_paths
            if component_id in updated_resolved.components
            else {}
        )
        receipt = {
            "schema": "agentic-authoring-result/v1",
            "operation": "add-component",
            "operation_id": operation_id,
            "document_id": resolved.manifest.id,
            "component_id": component_id,
            "component_type": declaration["type"],
            "canonical_source": str(canonical_path) if canonical_path else None,
            "canonical_source_sha256": file_hash(canonical_path) if canonical_path else None,
            "source_selector": selector,
            "related_sources": {
                name: {"path": str(path), "sha256": file_hash(path)}
                for name, path in sorted(canonical_related.items())
            },
            "placement": {
                "parent": parent_component,
                "slot": selected_slot if parent_component else None,
                "region": region,
                "after": after_component,
            },
            "marker_appended": marker_changed,
            "manifest": str(resolved.manifest_path),
            "manifest_sha256": file_hash(resolved.manifest_path),
            "recovery": str(history),
            "build_performed": False,
            "next_steps": [
                f"Run: Document-System.cmd validate {resolved.manifest.id}",
                f"Run: Document-System.cmd preview {resolved.manifest.id} --component {component_id}",
            ],
        }
        _atomic_write_text(history / "receipt.json", json.dumps(receipt, indent=2) + "\n")
    except Exception as exc:
        _atomic_write_text(resolved.manifest_path, original_manifest)
        if marker_changed and marker_path is not None and marker_original is not None:
            _atomic_write_text(marker_path, marker_original)
        for created_target in reversed(created_targets):
            created_target.unlink(missing_ok=True)
        try:
            _atomic_write_text(
                history / "rolled-back.json",
                json.dumps(
                    {
                        "schema": "agentic-authoring-rollback/v1",
                        "operation_id": operation_id,
                        "component_id": component_id,
                        "error": str(exc),
                        "canonical_restored": True,
                    },
                    indent=2,
                )
                + "\n",
            )
        except Exception:
            pass
        raise
    return receipt


def accept_diagram_rendition(resolved, *, component_id: str, rendition: Path) -> dict:
    """Bind an explicitly reviewed raster rendition to the current native diagram revision."""

    component_id = _safe_identifier(component_id, "Component id")
    if component_id not in resolved.components:
        raise DocumentSystemError(f"Component {component_id!r} does not exist")
    component = resolved.components[component_id]
    if component.declaration.type.value != "diagram" or component.source_path is None:
        raise DocumentSystemError(f"Component {component_id!r} is not a native diagram")
    rendition = Path(rendition).expanduser().resolve()
    if not rendition.is_file():
        raise DocumentSystemError(f"Reviewed diagram rendition does not exist: {rendition}")
    if rendition.suffix.lower() not in {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff"}:
        raise DocumentSystemError("A reviewed diagram rendition must be a PNG, JPEG, BMP, GIF, or TIFF image")

    selector, copy_source, new_target = _relative_selector(resolved, rendition, component_id)
    native_hash = file_hash(component.source_path)
    original_manifest = resolved.manifest_path.read_text(encoding="utf-8-sig")
    candidate_manifest = replace_values(
        original_manifest,
        {
            ("components", component_id, "options", "rendition"): selector,
            ("components", component_id, "options", "rendition_of_sha256"): native_hash,
        },
    )
    DocumentManifest.model_validate(load_jsonc_text(candidate_manifest, resolved.manifest_path))

    operation_id = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S_%fZ")
    history = resolved.manifest_path.parent / ".history" / "authoring" / f"{operation_id}-rendition-{component_id}"
    history.mkdir(parents=True, exist_ok=False)
    (history / "document.jsonc").write_text(original_manifest, encoding="utf-8")
    created_target = False
    try:
        if new_target is not None:
            new_target.parent.mkdir(parents=True, exist_ok=True)
            temporary = new_target.with_name(f".{new_target.name}.{uuid.uuid4().hex}.tmp")
            try:
                shutil.copy2(copy_source, temporary)
                os.replace(temporary, new_target)
                created_target = True
            finally:
                temporary.unlink(missing_ok=True)
        _atomic_write_text(resolved.manifest_path, candidate_manifest)
        from .resolver import resolve_document

        updated = resolve_document(resolved.manifest_path, DiagnosticBag())
        canonical = updated.components[component_id].related_paths["rendition"]
        receipt = {
            "schema": "agentic-authoring-result/v1",
            "operation": "accept-diagram-rendition",
            "operation_id": operation_id,
            "document_id": resolved.manifest.id,
            "component_id": component_id,
            "native_source": str(component.source_path),
            "native_source_sha256": native_hash,
            "reviewed_rendition": str(canonical),
            "reviewed_rendition_sha256": file_hash(canonical),
            "manifest": str(resolved.manifest_path),
            "manifest_sha256": file_hash(resolved.manifest_path),
            "recovery": str(history),
            "build_performed": False,
            "next_steps": [
                f"Run: Document-System.cmd preview {resolved.manifest.id} --component {component_id}",
                f"Run: Document-System.cmd build {resolved.manifest.id}",
            ],
        }
        _atomic_write_text(history / "receipt.json", json.dumps(receipt, indent=2) + "\n")
    except Exception as exc:
        _atomic_write_text(resolved.manifest_path, original_manifest)
        if created_target and new_target is not None:
            new_target.unlink(missing_ok=True)
        _atomic_write_text(
            history / "rolled-back.json",
            json.dumps(
                {
                    "schema": "agentic-authoring-rollback/v1",
                    "operation_id": operation_id,
                    "component_id": component_id,
                    "error": str(exc),
                    "canonical_restored": True,
                },
                indent=2,
            )
            + "\n",
        )
        raise
    return receipt


def load_jsonc_text(text: str, path: Path) -> object:
    from .jsonc import loads_jsonc

    return loads_jsonc(text, source=str(path))
