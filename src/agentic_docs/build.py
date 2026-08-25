from __future__ import annotations

import json
import os
import shutil
import subprocess
import time
from datetime import datetime, timezone
from pathlib import Path

from lxml import etree

from .current import update_current_from_build
from .component_cache import ComponentAdapterCache, summarize_cache_events
from .build_comparison import compare_build_candidate, load_current_build_report, structure_diff, word_structure
from .diagnostics import DiagnosticBag
from .errors import DocumentSystemError, PackageError
from .inputs import compare_input_snapshots, load_current_input_snapshot, resolved_input_snapshot
from .model import BuildMode, ResolvedDocument
from .rendering import render_pdf, word_field_error_inventory
from .reporting import resolved_summary, write_json
from .resolver import file_hash
from .timing import StageTimings
from .visual_quality import analyze_page_furniture_preview, analyze_rendered_pages
from .word.bindings import apply_field_bindings
from .word.components import compile_component_wrapper
from .sources.markdown import markdown_slot_tag
from .word.fragments import replace_nested_slot, replace_sequence_slot
from .word.fragments import component_wrapper
from .word.ooxml import qn
from .word.package import (
    DocxPackage,
    apply_page_furniture,
    apply_styles,
    clear_page_furniture,
    core_property_inventory,
    extract_document_body_state,
    extract_word_block_state,
    page_furniture_inventory,
    normalize_footer_page_total_fields,
    prune_unreferenced_page_furniture,
    replace_sdt,
    set_different_first_page,
    set_even_and_odd_headers,
    set_core_property,
    set_do_not_compress_pictures,
    save_component_states,
    validate_docx,
)
from .word.regions import apply_region_layout


def _utc_id() -> str:
    return datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S_%fZ")


def _tree_inventory(root: Path) -> dict[str, dict[str, int | str]]:
    return {
        path.relative_to(root).as_posix(): {
            "size_bytes": path.stat().st_size,
            "sha256": file_hash(path),
        }
        for path in sorted(root.rglob("*"))
        if path.is_file()
    }


def _promote_run_directory(source: Path, destination: Path) -> dict:
    """Publish a completed run, with a verified-copy fallback for protected folders.

    Some Windows protected folders permit file creation but temporarily or
    permanently refuse a directory rename. The normal path remains an atomic
    rename. The fallback creates the final directory only after the run is
    complete, copies every file, compares a full relative-path/size/hash
    inventory, and records the non-atomic promotion method explicitly.
    """

    if destination.exists():
        raise DocumentSystemError(f"Build destination already exists: {destination}")
    last_error: PermissionError | None = None
    rename_attempts = 3
    for attempt in range(rename_attempts):
        try:
            source.replace(destination)
            return {"method": "atomic-rename", "attempts": attempt + 1, "verified": True}
        except PermissionError as exc:
            last_error = exc
            time.sleep(0.5)

    expected = _tree_inventory(source)
    try:
        shutil.copytree(source, destination)
        actual = _tree_inventory(destination)
        if actual != expected:
            raise DocumentSystemError("Verified-copy build promotion produced a mismatched file inventory")
    except Exception:
        shutil.rmtree(destination, ignore_errors=True)
        raise
    source_removed = False
    try:
        shutil.rmtree(source)
        source_removed = True
    except OSError:
        # A scanner may still hold the temporary directory. It remains hidden
        # and is not authoritative; the verified final directory is complete.
        source_removed = False
    return {
        "method": "verified-copy",
        "attempts": rename_attempts,
        "verified": True,
        "file_count": len(actual),
        "temporary_removed": source_removed,
        "rename_error": str(last_error) if last_error else None,
    }


def _flatten_sequence(resolved: ResolvedDocument) -> list[str]:
    return [item for group in resolved.manifest.sequence for item in group.items]


def _assembly_order(resolved: ResolvedDocument) -> list[str]:
    result: list[str] = []

    def append(component_id: str) -> None:
        result.append(component_id)
        for children in getattr(resolved.components[component_id].declaration, "slots", {}).values():
            for child in children:
                append(child)

    for component_id in _flatten_sequence(resolved):
        append(component_id)
    return result


def _component_subtree_order(resolved: ResolvedDocument, component_id: str) -> list[str]:
    result = [component_id]
    for children in getattr(resolved.components[component_id].declaration, "slots", {}).values():
        for child in children:
            result.extend(_component_subtree_order(resolved, child))
    return result


def _region_containing(resolved: ResolvedDocument, component_type: str) -> str | None:
    for group in resolved.manifest.sequence:
        if any(resolved.components[item].declaration.type.value == component_type for item in group.items):
            return group.region
    return None


def _selector_variants(resolved: ResolvedDocument, region: str, kind: str) -> dict[str, Path]:
    result = {}
    for variant in ("default", "first", "even"):
        path = resolved.presentation_paths.get(f"region.{region}.{kind}.{variant}")
        if path:
            result[variant] = path
    return result


def _effective_furniture_variants(
    resolved: ResolvedDocument,
    region: str,
) -> tuple[dict[str, dict[str, tuple[Path, str, str]]], bool, bool]:
    """Normalize section-wide Word header/footer switches before copying donors.

    Word's different-first and odd/even settings apply to the whole section,
    not independently to headers and footers. A donor DOCX may contain several
    distinct reference types, so two selectors pointing at the same donor file
    are still distinct selections. A variant in one kind requires the other
    kind's default donor to be inherited into that variant.
    """

    selected: dict[str, dict[str, tuple[Path, str, str]]] = {}
    for kind in ("header", "footer"):
        raw = _selector_variants(resolved, region, kind)
        normalized: dict[str, tuple[Path, str, str]] = {}
        for variant, path in raw.items():
            normalized[variant] = (path, "explicit", variant)
        selected[kind] = normalized

    first_used = any("first" in values for values in selected.values())
    for kind, values in selected.items():
        default = values.get("default")
        if default is None:
            continue
        if first_used and "first" not in values:
            values["first"] = (default[0], "inherited-default", "default")

    even_selected = any("even" in values for values in selected.values())
    if even_selected:
        for values in selected.values():
            default = values.get("default")
            if default is not None and "even" not in values:
                values["even"] = (default[0], "inherited-default", "default")

    def selected_hash(path: Path, kind: str, source_variant: str) -> str | None:
        try:
            inventory = page_furniture_inventory(DocxPackage(path))
        except Exception:
            return None
        sections = inventory.get("sections", [])
        if not sections:
            return None
        for reference in sections[0].get("references", []):
            if reference.get("kind") == kind and reference.get("type") == source_variant:
                return reference.get("sha256")
        return None

    even_used = False
    redundant_even_kinds: list[str] = []
    for kind, values in selected.items():
        even = values.get("even")
        default = values.get("default")
        if even is None:
            continue
        if default is None:
            even_used = True
            break
        even_hash = selected_hash(even[0], kind, even[2])
        default_hash = selected_hash(default[0], kind, default[2])
        if even_hash is not None and default_hash is not None and even_hash == default_hash:
            redundant_even_kinds.append(kind)
        else:
            even_used = True
    if not even_used:
        for kind in redundant_even_kinds:
            selected[kind].pop("even", None)
    return selected, first_used, even_used


def _has_furniture(resolved: ResolvedDocument, region: str) -> bool:
    return bool(_selector_variants(resolved, region, "header") or _selector_variants(resolved, region, "footer"))


def _apply_furniture_for_region(
    package: DocxPackage,
    resolved: ResolvedDocument,
    region: str,
    sections: list[int],
    *,
    preserve_existing: bool = False,
) -> list[dict]:
    results = []
    variants, first_used, even_used = _effective_furniture_variants(resolved, region)
    for kind in ("header", "footer"):
        for variant, (path, source, source_variant) in variants[kind].items():
            result = apply_page_furniture(
                package,
                DocxPackage(path),
                section_numbers=sections,
                kinds={kind},
                reference_types={source_variant},
                target_reference_type=variant,
                require_reference_types=True,
            )
            result.update(
                {
                    "region": region,
                    "kind": kind,
                    "variant": variant,
                    "source_variant": source_variant,
                    "selection": source,
                }
            )
            results.append(result)
    if first_used or not preserve_existing:
        set_different_first_page(package, sections, first_used)
    if even_used:
        set_even_and_odd_headers(package, True)
    return results


def _preview_action(component, *, lightweight: bool, include_heavy: bool) -> str:
    if not lightweight or include_heavy:
        return "include"
    configured = component.declaration.preview
    if configured is not None:
        return configured.mode
    return "placeholder" if component.declaration.type.value == "pdf_pages" else "include"


def _fixture_paragraph(text: str, *, style: str | None = None, page_break_before: bool = False):
    paragraph = etree.Element(qn("w:p"))
    properties = etree.SubElement(paragraph, qn("w:pPr"))
    if style:
        style_node = etree.SubElement(properties, qn("w:pStyle"))
        style_node.set(qn("w:val"), style)
    if page_break_before:
        etree.SubElement(properties, qn("w:pageBreakBefore"))
    run = etree.SubElement(paragraph, qn("w:r"))
    value = etree.SubElement(run, qn("w:t"))
    value.text = text
    return paragraph


def _fixture_page_break():
    paragraph = etree.Element(qn("w:p"))
    run = etree.SubElement(paragraph, qn("w:r"))
    page_break = etree.SubElement(run, qn("w:br"))
    page_break.set(qn("w:type"), "page")
    return paragraph


def _page_furniture_fixture_wrappers() -> list[etree._Element]:
    """Small stable body for inspecting cover, margins, headers, and footers."""

    pages = [
        (
            "Page furniture preview",
            "This compact proof isolates the selected cover, page regions, margins, header, footer, and page numbering.",
            False,
        ),
        (
            "Even-page check",
            "This page verifies the even-page presentation without compiling the document's full canonical content.",
            True,
        ),
        (
            "Odd-page check",
            "This page verifies the following odd-page presentation and makes parity mistakes visible immediately.",
            True,
        ),
    ]
    wrappers = []
    for index, (heading, body, page_break) in enumerate(pages, 1):
        blocks = ([] if not page_break else [_fixture_page_break()]) + [
            _fixture_paragraph(heading, style="Heading1"),
            _fixture_paragraph(body),
            _fixture_paragraph("Top-of-page clearance sample"),
            _fixture_paragraph("Bottom-of-page and folio alignment sample"),
        ]
        wrappers.append(component_wrapper(f"preview-page-{index}", blocks))
    return wrappers


def _compose_raw(
    resolved: ResolvedDocument,
    diagnostics: DiagnosticBag,
    *,
    build_work: Path,
    output_path: Path,
    component_id: str | None,
    lightweight: bool,
    include_heavy: bool,
    preview_presentation: str | None,
    cache: ComponentAdapterCache,
    cache_events: list[dict],
    timings: StageTimings,
) -> dict:
    whole_documents = [
        item
        for item in resolved.components.values()
        if item.declaration.type.value == "document"
        and bool(item.declaration.options.get("whole_document", False))
    ]
    whole_document = whole_documents[0] if whole_documents else None
    use_whole_document = bool(
        whole_document
        and whole_document.source_path is not None
        and preview_presentation is None
        and (component_id is None or component_id == whole_document.id)
    )
    package = DocxPackage(whole_document.source_path if use_whole_document else resolved.shell_path)
    style_path = resolved.presentation_paths.get("styles")
    if style_path is None:
        raise PackageError("A V2 build requires a selected Word style donor")
    if style_path.resolve() != package.path.resolve():
        apply_styles(package, DocxPackage(style_path))

    cover_path = resolved.presentation_paths.get("cover")
    if cover_path:
        cover_components = [
            item
            for item in resolved.components.values()
            if item.declaration.type.value == "cover"
        ]
        source_tag = cover_components[0].declaration.source_tag if cover_components else None
        replace_sdt(package, DocxPackage(cover_path), "AGDOC.COVER", source_tag or "AGDOC.COVER")

    wrappers = []
    compiled_ids: list[str] = []
    deferred_components: list[dict] = []
    selected_sequence = _flatten_sequence(resolved)
    if component_id:
        if component_id not in resolved.components:
            raise DocumentSystemError(f"Unknown preview component {component_id!r}")
        selected_sequence = [component_id]
    def compile_with_slots(item_id: str):
        declaration = resolved.components[item_id].declaration
        with timings.measure(
            "compile_component",
            component_id=item_id,
            component_type=declaration.type.value,
        ):
            wrapper = compile_component_wrapper(
                package,
                resolved,
                resolved.components[item_id],
                build_work=build_work,
                diagnostics=diagnostics,
                cache=cache,
                cache_events=cache_events,
                preview_action=_preview_action(
                    resolved.components[item_id],
                    lightweight=lightweight,
                    include_heavy=include_heavy,
                ),
            )
        action = _preview_action(
            resolved.components[item_id],
            lightweight=lightweight,
            include_heavy=include_heavy,
        )
        if action != "include":
            deferred_components.append(
                {
                    "id": item_id,
                    "type": declaration.type.value,
                    "action": action,
                    "source": str(resolved.components[item_id].source_path)
                    if resolved.components[item_id].source_path
                    else None,
                }
            )
        if wrapper is None:
            return None
        compiled_ids.append(item_id)
        if action != "include":
            return wrapper
        for slot_name, child_ids in resolved.components[item_id].declaration.slots.items():
            children = []
            for child_id in child_ids:
                child = compile_with_slots(child_id)
                if child is None:
                    raise PackageError(
                        f"Nested component {child_id!r} cannot be inserted because it produces no Word block"
                    )
                children.append(child)
            source_path = resolved.components[item_id].source_path
            slot_tag = (
                markdown_slot_tag(item_id, slot_name)
                if source_path is not None and source_path.suffix.lower() in {".md", ".markdown"}
                else slot_name
            )
            replace_nested_slot(wrapper, slot_tag, children)
        return wrapper

    if preview_presentation == "page-furniture":
        wrappers = _page_furniture_fixture_wrappers()
        compiled_ids.extend(["preview-page-1", "preview-page-2", "preview-page-3"])
        replace_sequence_slot(package, resolved.profile.body_slot, wrappers)
    elif use_whole_document:
        declaration = whole_document.declaration
        compiled_ids.append(whole_document.id)
        for slot_name, child_ids in declaration.slots.items():
            children = []
            for child_id in child_ids:
                child = compile_with_slots(child_id)
                if child is None:
                    raise PackageError(
                        f"Nested component {child_id!r} cannot be inserted because it produces no Word block"
                    )
                children.append(child)
            root = package.xml("word/document.xml")
            replace_nested_slot(root, slot_name, children)
            package.set_xml("word/document.xml", root)
    else:
        for item_id in selected_sequence:
            wrapper = compile_with_slots(item_id)
            if wrapper is not None:
                wrappers.append(wrapper)
        replace_sequence_slot(package, resolved.profile.body_slot, wrappers)

    region_order = [group.region for group in resolved.manifest.sequence]
    region_starts = dict(resolved.profile.region_starts)
    if resolved.profile.main_start_tag and "main" not in region_starts:
        region_starts["main"] = {
            "tag": resolved.profile.main_start_tag,
            "boundary": resolved.profile.layout_boundary,
        }
    region_configuration = []
    for index, region_id in enumerate(region_order):
        entry = {
            "id": region_id,
            "config": resolved.manifest.presentation.page_regions[region_id],
        }
        if index:
            start = region_starts.get(region_id)
            if start is None:
                raise PackageError(
                    f"Profile {resolved.profile.id!r} has no region_starts entry for non-first page region {region_id!r}"
                )
            if isinstance(start, dict):
                entry.update({"start_tag": start["tag"], "boundary": start.get("boundary", "next_page")})
            else:
                entry.update({"start_tag": start.tag, "boundary": start.boundary})
        region_configuration.append(entry)

    layout_result = apply_region_layout(package, region_configuration)
    preserve_regions = {
        region_id
        for region_id in region_order
        if resolved.manifest.presentation.page_regions[region_id].layout_mode == "preserve"
    }
    if not preserve_regions:
        set_even_and_odd_headers(package, False)
    furniture_results = []
    page_total_results = []
    for region_id in region_order:
        sections = layout_result["region_sections"][region_id]
        region = resolved.manifest.presentation.page_regions[region_id]
        preserve_existing = region.layout_mode == "preserve"
        if preserve_existing:
            furniture_results.append(
                {
                    "region": region_id,
                    "layout_mode": "preserve",
                    "preserved_existing": True,
                    "sections": sections,
                }
            )
        else:
            clear_page_furniture(package, section_numbers=sections)
        furniture_results.extend(
            _apply_furniture_for_region(
                package,
                resolved,
                region_id,
                sections,
                preserve_existing=preserve_existing,
            )
        )
        if region.numbering is not None:
            if region.numbering.page_count_scope == "region" and len(sections) > 1:
                diagnostics.warn(
                    "REGION_PAGE_TOTAL_SPANS_SECTIONS",
                    f"Page region {region_id!r} spans {len(sections)} Word sections; SECTIONPAGES totals each section separately",
                    location=str(resolved.manifest_path),
                    hint="Use one Word section for a region-scoped total, or select document page-count scope.",
                )
            page_total_results.append(
                {
                    "region": region_id,
                    **normalize_footer_page_total_fields(
                        package,
                        sections,
                        "section" if region.numbering.page_count_scope == "region" else "document",
                    ),
                }
            )
    layout_result["page_total_fields"] = page_total_results
    prune_result = prune_unreferenced_page_furniture(package)

    bindings = apply_field_bindings(package, resolved, diagnostics)
    metadata = resolved.manifest.metadata
    set_core_property(package, "title", metadata.title)
    set_core_property(package, "subject", metadata.subject or metadata.type)
    set_core_property(package, "creator", metadata.author or metadata.prepared_by or "")
    set_core_property(package, "description", metadata.description or "")
    set_core_property(package, "keywords", ", ".join(metadata.keywords))
    core_properties = core_property_inventory(package)
    picture_compression = set_do_not_compress_pictures(package, True)
    package.write(output_path)
    validation = validate_docx(output_path)
    if not validation["valid"]:
        raise PackageError("Compiled DOCX failed package validation: " + "; ".join(validation["issues"]))
    return {
        "output": str(output_path),
        "sha256": file_hash(output_path),
        "component_count": len(compiled_ids),
        "compiled_components": compiled_ids,
        "deferred_components": deferred_components,
        "base_document": {
            "mode": "canonical-whole-document" if use_whole_document else "profile-shell",
            "component": whole_document.id if use_whole_document else None,
            "path": str(package.path),
        },
        "layout": layout_result,
        "page_furniture": furniture_results,
        "page_furniture_inventory": page_furniture_inventory(DocxPackage(output_path)),
        "media_normalizations": package.media_normalizations,
        "picture_compression": picture_compression,
        "pruned_page_furniture": prune_result,
        "field_bindings": bindings,
        "core_properties": core_properties,
        "package_validation": validation,
    }


def _preservation_signature(structure: dict) -> dict:
    """Keep layout/content-shape evidence while ignoring Word's relationship and XML rewrites."""

    furniture = structure.get("page_furniture") or {}
    sections = []
    for section in furniture.get("sections") or []:
        references = []
        for reference in section.get("references") or []:
            references.append(
                {
                    "kind": reference.get("kind"),
                    "type": reference.get("type"),
                    "visible_text": reference.get("visible_text"),
                    "fields": reference.get("fields") or [],
                }
            )
        sections.append(
            {
                "section": section.get("section"),
                "different_first_page": bool(section.get("different_first_page")),
                "references": references,
            }
        )
    return {
        key: structure.get(key)
        for key in (
            "package_valid",
            "tables",
            "table_rows",
            "drawings",
            "explicit_page_breaks",
            "headings_by_style",
            "section_count",
            "sections",
            "comments",
            "tracked_revisions",
        )
    } | {
        "page_furniture": {
            "section_count": furniture.get("section_count"),
            "even_and_odd_headers": bool(furniture.get("even_and_odd_headers")),
            "sections": sections,
        }
    }


def _verify_preserved_source_layout(resolved: ResolvedDocument, candidate: Path) -> dict:
    whole_documents = [
        item
        for item in resolved.components.values()
        if item.declaration.type.value == "document"
        and bool(item.declaration.options.get("whole_document", False))
    ]
    preserve_requested = any(
        region.layout_mode == "preserve"
        for region in resolved.manifest.presentation.page_regions.values()
    )
    if not whole_documents or not preserve_requested:
        return {
            "applicable": False,
            "passed": True,
            "reason": "no canonical whole-document source is governed by preserved page regions",
        }
    source = whole_documents[0].source_path
    if source is None:
        return {"applicable": True, "passed": False, "reason": "whole-document source is unavailable"}
    before = _preservation_signature(word_structure(source))
    after = _preservation_signature(word_structure(candidate))
    comparison = structure_diff(before, after)
    return {
        "applicable": True,
        "passed": not comparison["changed"],
        "source": str(source),
        "source_sha256": file_hash(source),
        "candidate": str(candidate),
        "change_count": comparison["change_count"],
        "changes": comparison["changes"],
    }


def _run_powershell(script: Path, parameters: list[str]) -> dict:
    command = [
        "powershell.exe",
        "-NoLogo",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script),
        *parameters,
    ]
    completed = subprocess.run(command, capture_output=True, text=True, timeout=180, check=False)
    if completed.returncode != 0:
        raise DocumentSystemError(
            f"Word automation failed in {script.name}: {(completed.stderr or completed.stdout).strip()}"
        )
    output = completed.stdout.strip().splitlines()
    if output:
        try:
            return json.loads(output[-1])
        except json.JSONDecodeError:
            pass
    return {"ok": True, "stdout": completed.stdout.strip()}


def _atomic_copy(source: Path, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_suffix(target.suffix + ".tmp")
    shutil.copy2(source, temporary)
    os.replace(temporary, target)


def _relocate_paths(value, old_root: Path, new_root: Path):
    """Translate paths recorded during staging into immutable final-run paths."""
    if isinstance(value, dict):
        return {key: _relocate_paths(item, old_root, new_root) for key, item in value.items()}
    if isinstance(value, list):
        return [_relocate_paths(item, old_root, new_root) for item in value]
    if isinstance(value, str):
        old_text = str(old_root)
        if value.lower().startswith(old_text.lower()):
            return str(new_root) + value[len(old_text):]
    return value


def _embed_component_baselines(
    path: Path,
    resolved: ResolvedDocument,
    *,
    component_id: str | None,
) -> dict:
    """Record post-Word state so a returned workcopy can be compared precisely."""

    package = DocxPackage(path)
    selected = (
        _component_subtree_order(resolved, component_id)
        if component_id
        else _assembly_order(resolved)
    )
    states = {}
    for item_id in selected:
        if item_id not in resolved.components:
            continue
        try:
            component = resolved.components[item_id]
            if bool(getattr(getattr(component.declaration, "options", None), "whole_document", False)):
                states[item_id] = extract_document_body_state(package)
            else:
                states[item_id] = extract_word_block_state(package, f"AGDOC.COMPONENT.{item_id}")
        except ValueError:
            continue
    if not states:
        return {"embedded": False, "component_count": 0, "components": []}
    save_component_states(package, states)
    candidate = path.with_suffix(path.suffix + ".state.tmp")
    if candidate.exists():
        candidate.unlink()
    package.write(candidate)
    os.replace(candidate, path)
    return {
        "embedded": True,
        "component_count": len(states),
        "components": sorted(states),
    }


def _finalize_intermediates(
    raw_docx: Path,
    work_directory: Path,
    *,
    retain: bool,
    raw_sha256: str,
) -> dict:
    """Inventory diagnostic artifacts and remove them unless explicitly kept."""

    raw_docx_size = raw_docx.stat().st_size if raw_docx.is_file() else 0
    work_size = (
        sum(item.stat().st_size for item in work_directory.rglob("*") if item.is_file())
        if work_directory.is_dir()
        else 0
    )
    result = {
        "retained": retain,
        "raw_docx": str(raw_docx),
        "raw_docx_sha256": raw_sha256,
        "work_directory": str(work_directory),
        "raw_docx_size_bytes": raw_docx_size,
        "work_size_bytes": work_size,
        "removed_size_bytes": 0 if retain else raw_docx_size + work_size,
    }
    if not retain:
        try:
            raw_docx.unlink()
        except FileNotFoundError:
            pass
        shutil.rmtree(work_directory, ignore_errors=True)
    return result


def _refresh_current(
    report: dict,
    resolved: ResolvedDocument,
    *,
    update_current: bool,
) -> dict:
    """Refresh the convenience mirror without invalidating a completed build."""

    report["current_update_requested"] = update_current
    if not update_current:
        report["current_updated"] = False
        report["current_docx"] = None
        report["current_pdf"] = None
        return report

    report["current_updated"] = True
    try:
        current_update = update_current_from_build(
            report,
            resolved.system_root / "current" / resolved.manifest.id,
            basename=resolved.manifest.outputs.basename,
        )
    except Exception as exc:
        report["current_updated"] = False
        report["current_docx"] = None
        report["current_pdf"] = None
        report["current_update_error"] = str(exc)
        return report

    report["current_updated"] = True
    report["current_docx"] = current_update["current_docx"]
    report["current_pdf"] = current_update["current_pdf"]
    report["current_update"] = current_update
    return report


def build_document(
    resolved: ResolvedDocument,
    diagnostics: DiagnosticBag,
    *,
    output_root: Path | None = None,
    component_id: str | None = None,
    preview_presentation: str | None = None,
    mode: BuildMode = BuildMode.DRAFT,
    render_pages: bool = True,
    lightweight: bool = False,
    include_heavy: bool = False,
    update_current: bool = True,
    retain_intermediates: bool = False,
) -> dict:
    if preview_presentation not in {None, "page-furniture"}:
        raise DocumentSystemError(f"Unknown presentation preview {preview_presentation!r}")
    if preview_presentation and component_id:
        raise DocumentSystemError("Choose either a component preview or a presentation preview")
    if preview_presentation:
        lightweight = True
    content_scope = (
        "page-furniture"
        if preview_presentation
        else "component"
        if component_id
        else "lightweight"
        if lightweight
        else "complete"
    )
    preview_run = content_scope != "complete"
    if preview_run:
        update_current = False
    timings = StageTimings()
    component_cache = ComponentAdapterCache(resolved.system_root)
    cache_events: list[dict] = []
    build_id = _utc_id()
    builds_root = (output_root or (resolved.system_root / "builds")).resolve()
    document_root = builds_root / resolved.manifest.id
    temporary_run = document_root / f".{build_id}.tmp"
    final_run = document_root / build_id
    if temporary_run.exists() or final_run.exists():
        raise DocumentSystemError(f"Build run already exists: {build_id}")
    temporary_run.mkdir(parents=True)

    try:
        with timings.measure("snapshot_inputs"):
            input_snapshot = resolved_input_snapshot(resolved)
            baseline_report = load_current_build_report(resolved.system_root, resolved.manifest.id)
            input_changes = compare_input_snapshots(
                input_snapshot,
                load_current_input_snapshot(resolved),
            )
        raw_docx = temporary_run / "document.raw.docx"
        with timings.measure("compose"):
            compose_result = _compose_raw(
                resolved,
                diagnostics,
                build_work=temporary_run / "work",
                output_path=raw_docx,
                component_id=component_id,
                lightweight=lightweight,
                include_heavy=include_heavy,
                preview_presentation=preview_presentation,
                cache=component_cache,
                cache_events=cache_events,
                timings=timings,
            )
        scripts = resolved.system_root / "scripts"
        initial_word_docx = temporary_run / "document.word-initial.docx"
        final_docx = temporary_run / f"{resolved.manifest.outputs.basename}.docx"
        final_pdf = temporary_run / f"{resolved.manifest.outputs.basename}.pdf"
        finalized = False
        exported_pdf = False
        field_error_report = {
            "schema": "agentic-word-field-errors/v1",
            "source": None,
            "page_count": 0,
            "error_count": 0,
            "errors": [],
        }
        render_result = None
        if preview_run:
            component_baselines = {
                "embedded": False,
                "component_count": 0,
                "components": [],
                "reason": "preview builds do not embed coworker-edit baselines",
            }
            try:
                with timings.measure("word_preview_finalize_and_export"):
                    _run_powershell(
                        scripts / "Finalize-And-Export.ps1",
                        [
                            "-InputPath",
                            str(raw_docx),
                            "-OutputDocxPath",
                            str(final_docx),
                            "-OutputPdfPath",
                            str(final_pdf),
                        ],
                    )
                finalized = True
                exported_pdf = True
            except DocumentSystemError as exc:
                diagnostics.warn(
                    "WORD_PREVIEW_FINALIZATION_UNAVAILABLE",
                    str(exc),
                    hint="The raw preview DOCX remains usable; Word is required for a matching preview PDF.",
                )
                _atomic_copy(raw_docx, final_docx)
        else:
            try:
                with timings.measure("word_finalize_initial"):
                    _run_powershell(
                        scripts / "Update-WordFields.ps1",
                        ["-InputPath", str(raw_docx), "-OutputPath", str(initial_word_docx), "-Passes", "1"],
                    )
            except DocumentSystemError as exc:
                diagnostics.warn(
                    "WORD_FINALIZATION_UNAVAILABLE",
                    str(exc),
                    hint="The raw DOCX remains usable; controlled release requires successful native Word finalization.",
                )
                _atomic_copy(raw_docx, initial_word_docx)

            try:
                with timings.measure("embed_component_baselines"):
                    component_baselines = _embed_component_baselines(
                        initial_word_docx,
                        resolved,
                        component_id=component_id,
                    )
            except Exception as exc:
                component_baselines = {
                    "embedded": False,
                    "component_count": 0,
                    "components": [],
                    "error": str(exc),
                }
                diagnostics.warn(
                    "COMPONENT_BASELINE_EMBED_FAILED",
                    f"Could not embed coworker-edit baselines: {exc}",
                    hint="The DOCX remains usable, but compare/adopt reporting will be less precise.",
                )

            try:
                with timings.measure("word_finalize_after_baselines_and_export"):
                    _run_powershell(
                        scripts / "Finalize-And-Export.ps1",
                        [
                            "-InputPath",
                            str(initial_word_docx),
                            "-OutputDocxPath",
                            str(final_docx),
                            "-OutputPdfPath",
                            str(final_pdf),
                        ],
                    )
                finalized = True
                exported_pdf = True
            except DocumentSystemError as exc:
                diagnostics.warn(
                    "WORD_POST_BASELINE_FINALIZATION_UNAVAILABLE",
                    str(exc),
                    hint="The DOCX remains usable, but controlled release requires the final post-baseline field refresh.",
                )
                _atomic_copy(initial_word_docx, final_docx)
            finally:
                initial_word_docx.unlink(missing_ok=True)

        with timings.measure("verify_preserved_source_layout"):
            source_layout_verification = _verify_preserved_source_layout(resolved, final_docx)
        if not preview_run and not source_layout_verification["passed"]:
            diagnostics.error(
                "PRESERVED_SOURCE_LAYOUT_CHANGED",
                "The final Word document changed structure governed by source-layout preservation",
                location=str(final_docx),
                hint="Review the reported section, page-furniture, or body-shape changes before accepting this build.",
            )

        if exported_pdf:
            with timings.measure("inspect_word_field_results"):
                field_error_report = word_field_error_inventory(final_pdf)
            if field_error_report["error_count"]:
                diagnostics.error(
                    "WORD_FIELD_ERROR_IN_PDF",
                    f"The exported PDF contains {field_error_report['error_count']} broken Word field result(s)",
                    location=str(final_pdf),
                    hint="Rebuild after correcting the TOC, cross-reference, or PDF export field behavior.",
                )
        else:
            try:
                with timings.measure("pdf_export"):
                    _run_powershell(
                        scripts / "Export-WordPdf.ps1",
                        ["-InputPath", str(final_docx), "-OutputPath", str(final_pdf)],
                    )
                exported_pdf = True
                with timings.measure("inspect_word_field_results"):
                    field_error_report = word_field_error_inventory(final_pdf)
                if field_error_report["error_count"]:
                    diagnostics.error(
                        "WORD_FIELD_ERROR_IN_PDF",
                        f"The exported PDF contains {field_error_report['error_count']} broken Word field result(s)",
                        location=str(final_pdf),
                        hint="Rebuild after correcting the TOC, cross-reference, or PDF export field behavior.",
                    )
            except DocumentSystemError as exc:
                diagnostics.warn(
                    "PDF_EXPORT_UNAVAILABLE",
                    str(exc),
                    hint="The DOCX remains usable; controlled release requires a successful PDF export.",
                )

        if exported_pdf and render_pages:
            try:
                with timings.measure("pdf_render"):
                    render_result = render_pdf(final_pdf, temporary_run / "pages")
            except DocumentSystemError as exc:
                diagnostics.warn(
                    "PDF_RENDER_UNAVAILABLE",
                    str(exc),
                    hint="The PDF remains usable; controlled release requires a complete page-image render for inspection.",
                )

        visual_quality = analyze_rendered_pages(
            render_result,
            allow_blank_pages=resolved.manifest.quality.allow_blank_pages,
        )
        if render_result and not visual_quality["no_unexpected_blank_pages"]:
            diagnostics.error(
                "UNEXPECTED_BLANK_PDF_PAGE",
                "The rendered PDF contains unexpected blank page(s): "
                + ", ".join(str(page) for page in visual_quality["blank_pages"]),
                location=str(final_pdf),
                hint="Check section starts and odd/even header settings, or explicitly allow intentional blank pages.",
            )
        page_furniture_visual = {
            "schema": "agentic-page-furniture-visual-proof/v1",
            "applicable": False,
            "passed": True,
        }
        if preview_presentation == "page-furniture":
            page_furniture_visual = analyze_page_furniture_preview(
                visual_quality,
                expect_header=bool(_selector_variants(resolved, "main", "header")),
                expect_footer=bool(_selector_variants(resolved, "main", "footer")),
            )
            if not page_furniture_visual["passed"]:
                missing = sorted(
                    set(page_furniture_visual["missing_header_pages"])
                    | set(page_furniture_visual["missing_footer_pages"])
                )
                diagnostics.error(
                    "PAGE_FURNITURE_NOT_VISIBLE",
                    "The full-page preview could not verify selected page furniture on sample page(s): "
                    + ", ".join(str(page) for page in missing),
                    location=str(final_pdf),
                    hint="Inspect the original full-page renders; do not accept a cropped image-viewer preview as proof.",
                )

        if preview_run:
            build_comparison = {
                "schema": "agentic-build-comparison/v1",
                "baseline_available": False,
                "reason": "scoped previews are not compared with a complete current document",
                "review_required": False,
                "unexplained_change": False,
            }
        else:
            with timings.measure("compare_with_current"):
                build_comparison = compare_build_candidate(
                    candidate_docx=final_docx,
                    candidate_render=render_result,
                    baseline_report=baseline_report,
                    input_changes=input_changes,
                    output_directory=temporary_run / "comparison" / "page-differences",
                )

        with timings.measure("compact_intermediates"):
            intermediate_result = _finalize_intermediates(
                raw_docx,
                temporary_run / "work",
                retain=retain_intermediates,
                raw_sha256=compose_result["sha256"],
            )
        compose_result["output_retained"] = retain_intermediates

        summary = resolved_summary(resolved, diagnostics)
        scope_checks = {
            "docx_package_valid": bool(compose_result["package_validation"]["valid"]),
            "word_fields_updated": finalized,
            "pdf_exported": exported_pdf,
            "all_pdf_pages_rendered": bool(
                render_result
                and render_result.get("page_count", 0) > 0
                and render_result.get("page_count") == render_result.get("pdf_info", {}).get("pages")
            ),
            "metadata_title_matches": bool(
                render_result
                and render_result.get("pdf_info", {}).get("title") == resolved.manifest.metadata.title
            ),
            "no_word_field_errors": bool(exported_pdf and field_error_report["error_count"] == 0),
            "no_unexpected_blank_pages": bool(visual_quality["no_unexpected_blank_pages"]),
            "no_error_diagnostics": not diagnostics.has_errors,
        }
        if preview_presentation == "page-furniture":
            scope_checks["page_furniture_visible_on_sample_pages"] = bool(page_furniture_visual["passed"])
        complete_checks = scope_checks | {
            "component_baselines_embedded": bool(component_baselines.get("embedded")),
            "preserved_source_layout_verified": bool(source_layout_verification.get("passed")),
            "no_unexplained_output_change": not build_comparison.get("unexplained_change", False),
        }
        quality_checks = complete_checks if content_scope == "complete" else scope_checks
        scope_passed = all(quality_checks.values())
        release_ready = bool(content_scope == "complete" and scope_passed)
        report_mode = (
            "page-furniture-preview"
            if preview_presentation
            else "component-preview"
            if component_id
            else "lightweight-preview"
            if lightweight
            else mode.value
        )
        report = {
            "schema": "agentic-build-report/v2",
            "build_id": build_id,
            "mode": report_mode,
            "content_scope": content_scope,
            "complete_content": content_scope == "complete",
            "verification_mode": (
                "full"
                if content_scope == "complete" and render_pages
                else "scoped"
                if render_pages
                else "unrendered"
            ),
            "preview_component": component_id,
            "preview_presentation": preview_presentation,
            "resolved": summary,
            "compose": compose_result,
            "artifacts": {
                "docx": str(final_docx),
                "docx_sha256": file_hash(final_docx),
                "pdf": str(final_pdf) if exported_pdf else None,
                "pdf_sha256": file_hash(final_pdf) if exported_pdf else None,
                "word_fields_updated": finalized,
            },
            "render": render_result,
            "word_field_errors": field_error_report,
            "visual_quality": visual_quality,
            "page_furniture_visual": page_furniture_visual,
            "component_baselines": component_baselines,
            "source_layout_verification": source_layout_verification,
            "component_cache": summarize_cache_events(cache_events),
            "inputs": {
                "fingerprint": input_snapshot["fingerprint"],
                "engine_sha256": input_snapshot["engine"]["sha256"],
                "changes_since_current": input_changes,
            },
            "comparison": build_comparison,
            "intermediates": intermediate_result,
            "timings": timings.snapshot(),
            "quality": {
                "checks": quality_checks,
                "scope_passed": scope_passed,
                "release_ready": release_ready,
                "passed": scope_passed,
            },
            "diagnostics": diagnostics.to_list(),
        }
        report = _relocate_paths(report, temporary_run, final_run)
        report["run_directory"] = str(final_run)
        report["current_update_requested"] = update_current
        report["current_updated"] = False if not update_current else None
        write_json(
            temporary_run / "resolved-inputs.json",
            {
                "schema": "agentic-resolved-inputs/v2",
                "document": resolved.manifest.model_dump(mode="json", by_alias=True),
                "kit": resolved.kit.model_dump(mode="json", by_alias=True),
                "profile": resolved.profile.model_dump(mode="json", by_alias=True),
                "project": resolved.project.model_dump(mode="json", by_alias=True),
                "resolved": summary,
                "input_snapshot": input_snapshot,
            },
        )
        write_json(temporary_run / "build-report.json", report)
        with timings.measure("promote_immutable_build"):
            report["promotion"] = _promote_run_directory(temporary_run, final_run)

        with timings.measure("refresh_current"):
            report = _refresh_current(report, resolved, update_current=update_current)
        report["timings"] = timings.snapshot()
        write_json(final_run / "build-report.json", report)
        return report
    except Exception:
        if temporary_run.exists():
            shutil.rmtree(temporary_run, ignore_errors=True)
        raise
