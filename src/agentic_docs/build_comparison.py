from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from PIL import Image, ImageChops, ImageOps

from .word.ooxml import content_control_inventory
from .word.package import DOCUMENT_PART, DocxPackage, XNS, page_furniture_inventory, validate_docx


def comparison_page_furniture(inventory: dict) -> dict:
    """Keep page-furniture meaning while excluding Word's volatile package IDs.

    Desktop Word rewrites relationship IDs, part names and harmless XML bytes on
    save.  Those values remain available in the build audit, but they are not
    structural changes when the furniture type, visible text and fields are the
    same.
    """

    sections = []
    for section in inventory.get("sections") or []:
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
        "section_count": inventory.get("section_count"),
        "even_and_odd_headers": bool(inventory.get("even_and_odd_headers")),
        "sections": sections,
    }


def word_structure(path: Path) -> dict:
    validation = validate_docx(path)
    if not validation.get("valid"):
        return {"package_valid": False, "issues": validation.get("issues") or []}
    package = DocxPackage(path)
    root = package.xml(DOCUMENT_PART)
    controls = content_control_inventory(root)
    paragraphs = root.xpath(".//w:body//w:p", namespaces=XNS)
    heading_styles: dict[str, int] = {}
    for paragraph in paragraphs:
        styles = paragraph.xpath("./w:pPr/w:pStyle/@w:val", namespaces=XNS)
        if styles and str(styles[0]).lower().startswith("heading"):
            heading_styles[styles[0]] = heading_styles.get(styles[0], 0) + 1
    sections = []
    for section in root.xpath(".//w:sectPr", namespaces=XNS):
        page_size = section.find("w:pgSz", namespaces=XNS)
        margins = section.find("w:pgMar", namespaces=XNS)

        def attribute(node, name):
            return node.get(f"{{{XNS['w']}}}{name}") if node is not None else None

        sections.append(
            {
                "page_width_twips": attribute(page_size, "w"),
                "page_height_twips": attribute(page_size, "h"),
                "orientation": attribute(page_size, "orient"),
                "margins_twips": {
                    name: attribute(margins, name)
                    for name in ("top", "right", "bottom", "left", "header", "footer", "gutter")
                },
            }
        )
    return {
        "package_valid": True,
        "paragraphs": len(paragraphs),
        "tables": len(root.xpath(".//w:body//w:tbl", namespaces=XNS)),
        "table_rows": len(root.xpath(".//w:body//w:tr", namespaces=XNS)),
        "drawings": len(root.xpath(".//w:body//w:drawing | .//w:body//w:pict", namespaces=XNS)),
        "explicit_page_breaks": len(root.xpath(".//w:br[@w:type='page']", namespaces=XNS)),
        "headings_by_style": dict(sorted(heading_styles.items())),
        "section_count": len(sections),
        "sections": sections,
        "content_control_count": len(controls),
        "component_tags": sorted(
            item["tag"] for item in controls if str(item.get("tag") or "").startswith("AGDOC.COMPONENT.")
        ),
        "slot_tags": sorted(
            item["tag"] for item in controls if str(item.get("tag") or "").startswith("AGDOC.MD.SLOT.")
        ),
        "comments": len(root.xpath(".//w:commentReference", namespaces=XNS)),
        "tracked_revisions": len(root.xpath(".//w:ins | .//w:del | .//w:moveFrom | .//w:moveTo", namespaces=XNS)),
        "page_furniture": comparison_page_furniture(page_furniture_inventory(package)),
    }


def _flatten(value: Any, prefix: str = "") -> dict[str, Any]:
    if isinstance(value, dict):
        result = {}
        for key in sorted(value):
            child = f"{prefix}.{key}" if prefix else str(key)
            result.update(_flatten(value[key], child))
        return result
    return {prefix: value}


def structure_diff(before: dict, after: dict) -> dict:
    old = _flatten(before)
    new = _flatten(after)
    changes = [
        {"field": key, "before": old.get(key), "after": new.get(key)}
        for key in sorted(set(old) | set(new))
        if old.get(key) != new.get(key)
    ]
    return {
        "available": True,
        "changed": bool(changes),
        "change_count": len(changes),
        "changes": changes,
        "before": before,
        "after": after,
    }


def _image_difference(before_path: Path, after_path: Path, diff_path: Path | None) -> dict:
    with Image.open(before_path) as before_image, Image.open(after_path) as after_image:
        before = before_image.convert("RGB")
        after = after_image.convert("RGB")
        if before.size != after.size:
            return {
                "changed": True,
                "reason": "dimensions_changed",
                "before_pixels": list(before.size),
                "after_pixels": list(after.size),
                "changed_fraction": 1.0,
                "mean_absolute_delta": 1.0,
                "diff_image": None,
            }
        difference = ImageChops.difference(before, after).convert("L")
        histogram = difference.histogram()
        pixels = max(1, before.width * before.height)
        changed_pixels = sum(histogram[9:])
        mean_delta = sum(level * count for level, count in enumerate(histogram)) / (pixels * 255)
        changed_fraction = changed_pixels / pixels
        changed = changed_fraction > 0.00001 or mean_delta > 0.000001
        artifact = None
        if changed and diff_path is not None:
            diff_path.parent.mkdir(parents=True, exist_ok=True)
            emphasized = ImageOps.autocontrast(difference)
            heatmap = ImageOps.colorize(emphasized, black="white", white="#D7191C")
            heatmap.save(diff_path)
            artifact = str(diff_path)
        return {
            "changed": changed,
            "reason": "pixels_changed" if changed else "identical_pixels",
            "before_pixels": list(before.size),
            "after_pixels": list(after.size),
            "changed_fraction": round(changed_fraction, 8),
            "mean_absolute_delta": round(mean_delta, 8),
            "diff_image": artifact,
        }


def _visual_diff(before_render: dict | None, after_render: dict | None, output_directory: Path) -> dict:
    if not before_render or not after_render:
        return {
            "available": False,
            "reason": "both builds require full page rendering",
            "changed": None,
        }
    before_pages = [Path(path) for path in before_render.get("page_images") or []]
    after_pages = [Path(path) for path in after_render.get("page_images") or []]
    if not before_pages or not after_pages or not all(path.is_file() for path in [*before_pages, *after_pages]):
        return {
            "available": False,
            "reason": "one rendered page set is unavailable",
            "changed": None,
        }
    page_results = []
    artifact_limit = 25
    artifact_count = 0
    for index in range(max(len(before_pages), len(after_pages))):
        page_number = index + 1
        if index >= len(before_pages):
            page_results.append({"page": page_number, "changed": True, "reason": "page_added"})
            continue
        if index >= len(after_pages):
            page_results.append({"page": page_number, "changed": True, "reason": "page_removed"})
            continue
        diff_path = (
            output_directory / f"page-{page_number:04d}-difference.png"
            if artifact_count < artifact_limit
            else None
        )
        result = _image_difference(before_pages[index], after_pages[index], diff_path)
        if result["changed"] and result.get("diff_image"):
            artifact_count += 1
        page_results.append({"page": page_number, **result})
    changed_pages = [item["page"] for item in page_results if item["changed"]]
    return {
        "available": True,
        "changed": bool(changed_pages),
        "before_page_count": len(before_pages),
        "after_page_count": len(after_pages),
        "changed_page_count": len(changed_pages),
        "changed_pages": changed_pages,
        "pages": page_results,
        "difference_artifacts_created": artifact_count,
        "difference_artifact_limit": artifact_limit,
        "difference_artifacts_capped": len(changed_pages) > artifact_count,
    }


def load_current_build_report(system_root: Path, document_id: str) -> dict | None:
    path = Path(system_root) / "current" / document_id / "build-report.json"
    if not path.is_file():
        return None
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    if value.get("resolved", {}).get("document", {}).get("id") != document_id:
        return None
    if value.get("mode") in {
        "component-preview",
        "page-furniture-preview",
        "lightweight-preview",
    } or value.get("content_scope") not in {None, "complete"}:
        return None
    return value


def compare_build_candidate(
    *,
    candidate_docx: Path,
    candidate_render: dict | None,
    baseline_report: dict | None,
    input_changes: dict,
    output_directory: Path,
) -> dict:
    if baseline_report is None:
        return {
            "schema": "agentic-build-comparison/v1",
            "baseline_available": False,
            "reason": "no current full build is available",
            "review_required": False,
            "unexplained_change": False,
        }
    baseline_docx_value = baseline_report.get("artifacts", {}).get("docx")
    baseline_docx = Path(baseline_docx_value) if baseline_docx_value else None
    if baseline_docx is None or not baseline_docx.is_file():
        return {
            "schema": "agentic-build-comparison/v1",
            "baseline_available": False,
            "reason": "the current build's immutable Word artifact is unavailable",
            "review_required": False,
            "unexplained_change": False,
        }
    structure = structure_diff(word_structure(baseline_docx), word_structure(candidate_docx))
    visual = _visual_diff(baseline_report.get("render"), candidate_render, output_directory)
    output_changed = structure["changed"] or visual.get("changed") is True
    inputs_changed = bool(input_changes.get("changed"))
    unexplained = bool(output_changed and not inputs_changed and input_changes.get("baseline_available"))
    pagination_changed = bool(
        visual.get("available")
        and visual.get("before_page_count") != visual.get("after_page_count")
    )
    presentation_input_changed = any(
        item.get("category") in {"presentation", "configuration", "engine"}
        for item in input_changes.get("items") or []
    )
    layout_ripple = bool(
        structure["changed"]
        and inputs_changed
        and not presentation_input_changed
        and any(
            change["field"].startswith(("sections", "section_count", "page_furniture"))
            for change in structure["changes"]
        )
    )
    return {
        "schema": "agentic-build-comparison/v1",
        "baseline_available": True,
        "baseline_build_id": baseline_report.get("build_id"),
        "baseline_run_directory": baseline_report.get("run_directory"),
        "inputs_changed": inputs_changed,
        "structure": structure,
        "visual": visual,
        "pagination_changed": pagination_changed,
        "possible_layout_ripple": layout_ripple,
        "unexplained_change": unexplained,
        "review_required": bool(output_changed),
    }
