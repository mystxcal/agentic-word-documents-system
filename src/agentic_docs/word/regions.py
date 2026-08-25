from __future__ import annotations

import copy
from pathlib import Path

from lxml import etree

from .ooxml import find_sdts
from .package import DOCUMENT_PART, W_NS, XNS, DocxPackage, insert_section_property


def _last_flow_paragraph(block):
    local_name = etree.QName(block).localname
    if local_name == "p":
        return block
    if local_name != "sdt":
        return None
    content = block.find(f"{{{W_NS}}}sdtContent")
    if content is None:
        return None
    for child in reversed(content):
        if etree.QName(child).localname in {"bookmarkStart", "bookmarkEnd", "proofErr", "permStart", "permEnd"}:
            continue
        return _last_flow_paragraph(child)
    return None


def _remove(section, names: set[str]) -> int:
    removed = 0
    for child in list(section):
        if etree.QName(child).localname in names:
            section.remove(child)
            removed += 1
    return removed


def _body_child(body, marker, tag: str):
    value = marker
    while value.getparent() is not body:
        value = value.getparent()
        if value is None:
            raise ValueError(f"Region-start tag {tag!r} is not in the Word document body")
    return value


def _configuration_dict(value) -> dict:
    if hasattr(value, "model_dump"):
        return value.model_dump(mode="json", exclude_none=False)
    if isinstance(value, dict):
        return dict(value)
    raise ValueError("Region configuration must be a mapping or typed page-region model")


def apply_region_layout(package: DocxPackage, regions: list[dict]) -> dict:
    """Create and configure any ordered set of contiguous Word page regions."""

    if not regions:
        raise ValueError("At least one page region is required")
    region_ids = [str(item.get("id") or "").strip() for item in regions]
    if any(not region_id for region_id in region_ids) or len(set(region_ids)) != len(region_ids):
        raise ValueError("Page-region IDs must be non-empty and unique")

    root = package.xml(DOCUMENT_PART)
    body = root.find(f"{{{W_NS}}}body")
    if body is None:
        raise ValueError("Word document has no body")
    sections = root.xpath(".//w:sectPr", namespaces=XNS)
    if not sections:
        raise ValueError("Word document has no section properties")
    prototype = copy.deepcopy(sections[-1])

    boundaries = []
    prior_body_index = -1
    markers = {}
    for index, region in enumerate(regions):
        if index == 0:
            continue
        tag = str(region.get("start_tag") or "").strip()
        if not tag:
            raise ValueError(f"Page region {region_ids[index]!r} requires a profile region-start tag")
        matches = find_sdts(root, tag)
        if len(matches) != 1:
            raise ValueError(f"Region-start tag {tag!r} must appear exactly once; found {len(matches)}")
        marker = matches[0]
        top_level = _body_child(body, marker, tag)
        body_index = body.index(top_level)
        if body_index <= prior_body_index:
            raise ValueError("Profile region-start tags do not follow the document's region order")
        if body_index == 0:
            raise ValueError(f"Region-start tag {tag!r} has no preceding content")
        prior_body_index = body_index
        markers[region_ids[index]] = marker

        preceding = body[body_index - 1]
        # Section properties belong to a body-level paragraph. Descending into
        # a preceding cover/content-control and attaching sectPr there produces
        # unstable Word pagination: page geometry and furniture can alternate
        # between the old and new sections. Reuse only a direct body paragraph;
        # otherwise insert a compact boundary paragraph between the components.
        paragraph = preceding if etree.QName(preceding).localname == "p" else None
        inserted = paragraph is None
        if paragraph is None:
            paragraph = etree.Element(f"{{{W_NS}}}p")
            properties = etree.SubElement(paragraph, f"{{{W_NS}}}pPr")
            spacing = etree.SubElement(properties, f"{{{W_NS}}}spacing")
            spacing.set(f"{{{W_NS}}}before", "0")
            spacing.set(f"{{{W_NS}}}after", "0")
            spacing.set(f"{{{W_NS}}}line", "1")
            spacing.set(f"{{{W_NS}}}lineRule", "exact")
            body.insert(body_index, paragraph)

        removed_breaks = 0
        for page_break in list(paragraph.xpath(".//w:br[@w:type='page']", namespaces=XNS)):
            page_break.getparent().remove(page_break)
            removed_breaks += 1
        properties = paragraph.find(f"{{{W_NS}}}pPr")
        if properties is None:
            properties = etree.Element(f"{{{W_NS}}}pPr")
            paragraph.insert(0, properties)
        existing = properties.findall(f"{{{W_NS}}}sectPr")
        if len(existing) > 1:
            raise ValueError(f"Paragraph before region {region_ids[index]!r} has multiple section properties")
        section = existing[0] if existing else copy.deepcopy(prototype)
        if not existing:
            properties.append(section)
        _remove(section, {"type"})
        boundary = str(region.get("boundary") or "next_page")
        if boundary not in {"next_page", "continuous"}:
            raise ValueError("Region boundary must be next_page or continuous")
        section_type = etree.Element(f"{{{W_NS}}}type")
        section_type.set(f"{{{W_NS}}}val", "nextPage" if boundary == "next_page" else "continuous")
        insert_section_property(section, section_type)
        boundaries.append(
            {
                "region": region_ids[index],
                "start_tag": tag,
                "boundary": boundary,
                "inserted_boundary_paragraph": inserted,
                "removed_preceding_page_breaks": removed_breaks,
            }
        )

    sections = root.xpath(".//w:sectPr", namespaces=XNS)
    starts = [1]
    for region_id in region_ids[1:]:
        marker = markers[region_id]
        start = int(marker.xpath("count(preceding::w:sectPr)", namespaces=XNS)) + 1
        starts.append(start)
    if starts != sorted(set(starts)):
        raise ValueError(f"Could not establish strictly ordered page-region sections: {starts}")

    region_sections = {}
    for index, region_id in enumerate(region_ids):
        end = starts[index + 1] - 1 if index + 1 < len(starts) else len(sections)
        region_sections[region_id] = list(range(starts[index], end + 1))
        if not region_sections[region_id]:
            raise ValueError(f"Page region {region_id!r} contains no Word sections")

    margin_names = {
        "top_margin_twips": "top",
        "bottom_margin_twips": "bottom",
        "left_margin_twips": "left",
        "right_margin_twips": "right",
    }
    numbering_results = []
    for region, region_id in zip(regions, region_ids):
        config = _configuration_dict(region.get("config") or {})
        layout_mode = config.get("layout_mode", "managed")
        if layout_mode not in {"managed", "preserve"}:
            raise ValueError(f"Page region {region_id!r} has unsupported layout_mode {layout_mode!r}")
        selected_sections = [sections[number - 1] for number in region_sections[region_id]]
        for section in selected_sections:
            margins = section.find(f"{{{W_NS}}}pgMar")
            supplied_margins = {key: config.get(key) for key in margin_names if config.get(key) is not None}
            if supplied_margins and margins is None:
                margins = etree.Element(f"{{{W_NS}}}pgMar")
                insert_section_property(section, margins)
            for key, value in supplied_margins.items():
                if not isinstance(value, int) or isinstance(value, bool) or value < 0:
                    raise ValueError(f"Page region {region_id!r} has invalid {key}")
                margins.set(f"{{{W_NS}}}{margin_names[key]}", str(value))

        numbering = config.get("numbering")
        if numbering is not None or layout_mode == "managed":
            for section in selected_sections:
                _remove(section, {"pgNumType"})
        if numbering is not None:
            numbering = _configuration_dict(numbering)
            style = numbering.get("style", "arabic")
            formats = {"arabic": "decimal", "roman_lower": "lowerRoman", "roman_upper": "upperRoman"}
            if style not in formats:
                raise ValueError(f"Page region {region_id!r} has unsupported numbering style {style!r}")
            start = numbering.get("start", 1)
            if start is not None and (not isinstance(start, int) or isinstance(start, bool) or start < 0):
                raise ValueError(f"Page region {region_id!r} has invalid numbering start")
            page_numbering = etree.Element(f"{{{W_NS}}}pgNumType")
            page_numbering.set(f"{{{W_NS}}}fmt", formats[style])
            if start is not None:
                page_numbering.set(f"{{{W_NS}}}start", str(start))
            insert_section_property(selected_sections[0], page_numbering)
            numbering_results.append(
                {
                    "region": region_id,
                    "layout_mode": layout_mode,
                    "style": style,
                    "start": start,
                    "page_count_scope": numbering.get("page_count_scope", "region"),
                }
            )

    package.set_xml(DOCUMENT_PART, root)
    result = {
        "applied": True,
        "region_order": region_ids,
        "section_count": len(sections),
        "region_sections": region_sections,
        "boundaries": boundaries,
        "numbering": numbering_results,
        "layout_modes": {
            region_id: _configuration_dict(region.get("config") or {}).get("layout_mode", "managed")
            for region, region_id in zip(regions, region_ids)
        },
    }
    if "front" in region_sections:
        result["front_sections"] = len(region_sections["front"])
    if "main" in region_sections:
        result["main_sections"] = len(region_sections["main"])
    return result
