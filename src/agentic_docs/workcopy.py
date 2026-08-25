from __future__ import annotations

import copy
import difflib
import os
import shutil
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from lxml import etree

from .build import build_document
from .diagnostics import DiagnosticBag
from .errors import OperationPartialError, RevisionError
from .model import ComponentType, Ownership, ResolvedComponent, ResolvedDocument
from .publishing import PublishPolicy, publish_build
from .reporting import write_json
from .resolver import file_hash, resolve_document
from .word.ooxml import qn
from .word.package import (
    DOCUMENT_PART,
    CONTENT_TYPES,
    DocxPackage,
    XNS,
    extract_word_block_state,
    json_hash,
    load_component_states,
    rels_part_for,
    require_one_sdt,
    resolve_target,
    validate_docx,
)


ALWAYS_KEEP_DOCUMENT_RELATIONSHIPS = {
    "styles",
    "stylesWithEffects",
    "settings",
    "fontTable",
    "theme",
    "numbering",
    "webSettings",
}


def _component(resolved: ResolvedDocument, component_id: str | None) -> ResolvedComponent:
    if component_id:
        component = resolved.components.get(component_id)
        if component is None:
            raise RevisionError(f"Unknown component {component_id!r}")
        candidates = [component]
    else:
        candidates = [
            item
            for item in resolved.components.values()
            if item.declaration.type == ComponentType.DOCUMENT
            and item.declaration.ownership == Ownership.WORD_FRAGMENT
            and item.source_path is not None
        ]
        if len(candidates) != 1:
            names = ", ".join(item.id for item in candidates) or "none"
            raise RevisionError(
                "The document does not have exactly one adoptable Word component. "
                f"Use --component. Candidates: {names}"
            )
    component = candidates[0]
    if component.source_path is None or component.source_path.suffix.lower() != ".docx":
        raise RevisionError(f"Component {component.id!r} is not backed by a canonical DOCX")
    if component.declaration.ownership != Ownership.WORD_FRAGMENT:
        raise RevisionError(f"Component {component.id!r} is not owned as word_fragment")
    return component


def _tags(component: ResolvedComponent) -> tuple[str, str]:
    target_tag = component.declaration.source_tag
    if not target_tag:
        raise RevisionError(
            f"Canonical component {component.id!r} has no source_tag; precise workcopy adoption is unavailable"
        )
    source_tag = f"AGDOC.COMPONENT.{component.id}"
    return target_tag, source_tag


def _paragraph_records(block: etree._Element) -> list[tuple[str, str, bool, str]]:
    records = []
    for paragraph in block.xpath(".//w:p", namespaces=XNS):
        text = "".join(paragraph.xpath(".//w:t/text()", namespaces=XNS))
        style = paragraph.xpath("./w:pPr/w:pStyle/@w:val", namespaces=XNS)
        num_id = paragraph.xpath("./w:pPr/w:numPr/w:numId/@w:val", namespaces=XNS)
        level = paragraph.xpath("./w:pPr/w:numPr/w:ilvl/@w:val", namespaces=XNS)
        records.append(
            (
                text,
                style[0] if style else "",
                bool(num_id and num_id[0] != "0"),
                level[0] if level else "",
            )
        )
    return records


def _block(package: DocxPackage, tag: str) -> etree._Element:
    return require_one_sdt(package.xml(DOCUMENT_PART), tag)


def diff_workcopy(
    resolved: ResolvedDocument,
    workcopy_path: Path,
    *,
    component_id: str | None = None,
    max_diff_lines: int = 400,
) -> dict[str, Any]:
    component = _component(resolved, component_id)
    target_tag, source_tag = _tags(component)
    workcopy_path = Path(workcopy_path).expanduser().resolve()
    if not workcopy_path.is_file():
        raise RevisionError(f"Workcopy does not exist: {workcopy_path}")
    canonical_package = DocxPackage(component.source_path)
    workcopy_package = DocxPackage(workcopy_path)
    canonical_block = _block(canonical_package, target_tag)
    workcopy_block = _block(workcopy_package, source_tag)
    canonical_records = _paragraph_records(canonical_block)
    workcopy_records = _paragraph_records(workcopy_block)
    canonical_state = extract_word_block_state(canonical_package, target_tag)
    workcopy_state = extract_word_block_state(workcopy_package, source_tag)
    canonical_content_state = {
        key: canonical_state.get(key)
        for key in ("visible_text", "tables", "images", "alt_text")
    }
    workcopy_content_state = {
        key: workcopy_state.get(key)
        for key in ("visible_text", "tables", "images", "alt_text")
    }
    canonical_formatting = canonical_state.get("formatting")
    workcopy_formatting = workcopy_state.get("formatting")
    canonical_annotations = canonical_state.get("annotations")
    workcopy_annotations = workcopy_state.get("annotations")
    baselines = load_component_states(workcopy_package)
    baseline = baselines.get(component.id) if isinstance(baselines, dict) else None
    baseline_formatting = baseline.get("formatting") if isinstance(baseline, dict) else None
    baseline_content_state = (
        {
            key: baseline.get(key)
            for key in ("visible_text", "tables", "images", "alt_text")
        }
        if isinstance(baseline, dict)
        else None
    )
    canonical_lines = [record[0] for record in canonical_records]
    workcopy_lines = [record[0] for record in workcopy_records]
    unified = list(
        difflib.unified_diff(
            canonical_lines,
            workcopy_lines,
            fromfile="canonical",
            tofile="workcopy",
            lineterm="",
        )
    )
    matcher = difflib.SequenceMatcher(a=canonical_lines, b=workcopy_lines, autojunk=False)
    change_groups = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag != "equal":
            change_groups.append(
                {
                    "kind": tag,
                    "canonical": canonical_lines[i1:i2],
                    "workcopy": workcopy_lines[j1:j2],
                }
            )
    canonical_shape = {
        "paragraphs": len(canonical_records),
        "tables": len(canonical_block.xpath(".//w:tbl", namespaces=XNS)),
        "drawings": len(canonical_block.xpath(".//w:drawing | .//w:pict", namespaces=XNS)),
    }
    workcopy_shape = {
        "paragraphs": len(workcopy_records),
        "tables": len(workcopy_block.xpath(".//w:tbl", namespaces=XNS)),
        "drawings": len(workcopy_block.xpath(".//w:drawing | .//w:pict", namespaces=XNS)),
    }
    content_same = canonical_content_state == workcopy_content_state
    formatting_same = (
        workcopy_formatting == baseline_formatting
        if baseline_formatting is not None
        else None
    )
    annotations_same = canonical_annotations == workcopy_annotations
    known_same = (
        canonical_records == workcopy_records
        and canonical_shape == workcopy_shape
        and content_same
        and annotations_same
    )
    same = (
        False
        if not known_same or formatting_same is False
        else True
        if formatting_same is True
        else None
    )
    canonical_content_changed = (
        baseline_content_state != canonical_content_state
        if baseline_content_state is not None
        else None
    )
    workcopy_content_changed = (
        baseline_content_state != workcopy_content_state
        if baseline_content_state is not None
        else None
    )
    content_conflict = bool(
        canonical_content_changed
        and workcopy_content_changed
        and canonical_content_state != workcopy_content_state
    )
    return {
        "schema": "agentic-workcopy-diff/v2",
        "document_id": resolved.manifest.id,
        "component": component.id,
        "canonical": {
            "path": str(component.source_path),
            "sha256": file_hash(component.source_path),
            "tag": target_tag,
            "content_sha256": json_hash(canonical_content_state),
            "formatting_sha256": json_hash(canonical_formatting),
            "annotations_sha256": json_hash(canonical_annotations),
            **canonical_shape,
        },
        "workcopy": {
            "path": str(workcopy_path),
            "sha256": file_hash(workcopy_path),
            "tag": source_tag,
            "content_sha256": json_hash(workcopy_content_state),
            "formatting_sha256": json_hash(workcopy_formatting),
            "annotations_sha256": json_hash(workcopy_annotations),
            **workcopy_shape,
        },
        "same": same,
        "content_same": content_same,
        "formatting_same": formatting_same,
        "annotations_same": annotations_same,
        "baseline_available": isinstance(baseline, dict),
        "canonical_changed_since_build": {
            "content": canonical_content_changed,
        },
        "workcopy_changed_since_build": {
            "content": workcopy_content_changed,
            "formatting": (
                baseline_formatting != workcopy_formatting
                if baseline_formatting is not None
                else None
            ),
            "annotations": (
                baseline.get("annotations") != workcopy_annotations
                if isinstance(baseline, dict) and "annotations" in baseline
                else None
            ),
        },
        "content_conflict": content_conflict,
        "change_group_count": len(change_groups),
        "changes": change_groups,
        "unified_diff": unified[:max_diff_lines],
        "diff_truncated": len(unified) > max_diff_lines,
    }


def _candidate_from_workcopy(
    workcopy_path: Path,
    output_path: Path,
    *,
    canonical_tag: str,
    workcopy_tag: str,
) -> None:
    # The workcopy package is the most faithful owner of any newly introduced
    # numbering, styles, media, hyperlinks, comments, or tracked changes. Keep
    # that package, isolate the selected component, and retag it as canonical.
    package = DocxPackage(workcopy_path)
    root = package.xml(DOCUMENT_PART)
    selected = copy.deepcopy(require_one_sdt(root, workcopy_tag))
    tag_node = selected.find("./w:sdtPr/w:tag", namespaces=XNS)
    if tag_node is None:
        raise RevisionError(f"Workcopy component {workcopy_tag!r} has no Word tag node")
    tag_node.set(qn("w:val"), canonical_tag)
    alias = selected.find("./w:sdtPr/w:alias", namespaces=XNS)
    if alias is not None:
        alias.set(qn("w:val"), canonical_tag)
    body = root.find(qn("w:body"))
    if body is None:
        raise RevisionError("Workcopy has no Word document body")
    final_section = body.find(qn("w:sectPr"))
    final_section = copy.deepcopy(final_section) if final_section is not None else None
    for child in list(body):
        body.remove(child)
    body.append(selected)
    if final_section is not None:
        for reference in final_section.xpath("./w:headerReference | ./w:footerReference", namespaces=XNS):
            final_section.remove(reference)
        body.append(final_section)
    package.set_xml(DOCUMENT_PART, root)
    _prune_candidate_package(package)
    package.write(output_path)
    validation = validate_docx(output_path)
    if not validation["valid"]:
        raise RevisionError(
            "Adopted canonical candidate is not a valid DOCX: " + "; ".join(validation["issues"])
        )


def _prune_candidate_package(package: DocxPackage) -> None:
    """Remove unrelated compiled-document parts after isolating one adopted block."""

    document = package.xml(DOCUMENT_PART)
    active_rids = {
        value
        for node in document.iter()
        for attribute in (f"{{{XNS['r']}}}id", f"{{{XNS['r']}}}embed", f"{{{XNS['r']}}}link")
        for value in (node.get(attribute),)
        if value
    }
    has_comments = bool(document.xpath(".//w:commentReference", namespaces=XNS))
    has_footnotes = bool(document.xpath(".//w:footnoteReference", namespaces=XNS))
    has_endnotes = bool(document.xpath(".//w:endnoteReference", namespaces=XNS))
    rels_name, relationships = package.relationship_root(DOCUMENT_PART)
    for relationship in list(relationships.xpath("./pr:Relationship", namespaces=XNS)):
        rid = relationship.get("Id")
        kind = (relationship.get("Type") or "").rsplit("/", 1)[-1]
        keep = (
            rid in active_rids
            or kind in ALWAYS_KEEP_DOCUMENT_RELATIONSHIPS
            or (kind == "comments" and has_comments)
            or (kind == "footnotes" and has_footnotes)
            or (kind == "endnotes" and has_endnotes)
        )
        if not keep:
            relationships.remove(relationship)
    package.set_xml(rels_name, relationships)

    reachable = {CONTENT_TYPES, "_rels/.rels"}
    pending = [""]
    visited = set()
    while pending:
        owner = pending.pop()
        if owner in visited:
            continue
        visited.add(owner)
        relationship_part = "_rels/.rels" if owner == "" else rels_part_for(owner)
        if relationship_part not in package.parts:
            continue
        reachable.add(relationship_part)
        root = package.xml(relationship_part)
        for relationship in root.xpath("./pr:Relationship", namespaces=XNS):
            if relationship.get("TargetMode") == "External":
                continue
            target = resolve_target(owner, relationship.get("Target"))
            if target in package.parts and target not in reachable:
                reachable.add(target)
                pending.append(target)

    removed = set(package.parts) - reachable
    for name in removed:
        del package.parts[name]
    content_types = package.xml(CONTENT_TYPES)
    for override in list(content_types.xpath("./ct:Override", namespaces=XNS)):
        part_name = (override.get("PartName") or "").lstrip("/")
        if part_name in removed:
            content_types.remove(override)
    package.set_xml(CONTENT_TYPES, content_types)


def adopt_workcopy(
    resolved: ResolvedDocument,
    diagnostics: DiagnosticBag,
    workcopy_path: Path,
    *,
    component_id: str | None = None,
    allow_identical: bool = False,
    accept_conflict: bool = False,
    build: bool = True,
    render_pages: bool = True,
    retain_intermediates: bool = False,
    output_root: Path | None = None,
    publish_destination: Path | None = None,
    publish_policy: PublishPolicy = "versioned",
) -> dict[str, Any]:
    component = _component(resolved, component_id)
    target_tag, source_tag = _tags(component)
    workcopy_path = Path(workcopy_path).expanduser().resolve()
    difference = diff_workcopy(resolved, workcopy_path, component_id=component.id)
    if difference["content_conflict"] and not accept_conflict:
        raise RevisionError(
            "Adoption refused because canonical content and the returned workcopy both changed "
            "since the embedded build baseline. Review 'compare' and use --accept-conflict only "
            "when the returned component should intentionally replace the canonical changes."
        )
    if difference["same"] is True and not allow_identical:
        raise RevisionError("The selected workcopy component is already identical to canonical content")
    source = component.source_path
    assert source is not None
    original_hash = file_hash(source)
    temporary_root = resolved.manifest_path.parent / ".agentic-tmp"
    temporary_root.mkdir(parents=True, exist_ok=True)
    descriptor, candidate_name = tempfile.mkstemp(
        prefix=f"{component.id}.", suffix=".adopt.docx", dir=temporary_root
    )
    os.close(descriptor)
    candidate = Path(candidate_name)
    candidate.unlink()
    try:
        _candidate_from_workcopy(
            workcopy_path,
            candidate,
            canonical_tag=target_tag,
            workcopy_tag=source_tag,
        )
        if file_hash(source) != original_hash:
            raise RevisionError(f"Canonical source changed while adoption was being prepared: {source}")
        operation_id = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S_%fZ")
        history = resolved.manifest_path.parent / ".history" / "adoptions" / operation_id
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
                    f"Build failed after workcopy adoption; canonical source was restored. Cause: {exc}"
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
            "schema": "agentic-adoption-report/v2",
            "operation_id": operation_id,
            "document_id": resolved.manifest.id,
            "component": component.id,
            "workcopy": str(workcopy_path),
            "canonical_source": str(source),
            "source_before_sha256": original_hash,
            "source_after_sha256": file_hash(source),
            "backup": str(backup),
            "changed": True,
            "difference": difference,
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
        write_json(history / "adoption-report.json", report)
        if publish_error:
            raise OperationPartialError(
                "The workcopy adoption and rebuild succeeded, but delivery failed. "
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
