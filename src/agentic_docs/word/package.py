"""Low-level Word-native operations with no project or document-type assumptions."""

from __future__ import annotations

import base64
import copy
import hashlib
import io
import json
import posixpath
import re
import zipfile
import zlib
from datetime import date, datetime, timezone
from pathlib import Path

from lxml import etree
from PIL import Image

from .ooxml import CT_NS, NS, PKG_REL_NS, R_NS, W_NS, find_sdts


DOCUMENT_PART = "word/document.xml"
DOCUMENT_RELS = "word/_rels/document.xml.rels"
CONTENT_TYPES = "[Content_Types].xml"
REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
REL_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
DC_NS = "http://purl.org/dc/elements/1.1/"
DCTERMS_NS = "http://purl.org/dc/terms/"
CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
XNS = {
    **NS,
    "a": A_NS,
    "wp": WP_NS,
    "pic": PIC_NS,
    "dc": DC_NS,
    "dcterms": DCTERMS_NS,
    "cp": CP_NS,
}
ROW_MARKER = re.compile(r"\[\[AGDOCROW:([A-Za-z0-9_-]+)\]\]")
STATE_PART = "customXml/agenticManagedComponents.xml"
STATE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
STATE_NS = "urn:agentic-documents:managed-components:v1"
STATE_DOCVAR_PREFIX = "AgenticDocs.ManagedComponents."
STATE_DOCVAR_CHUNK = 20000


def utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def hash_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest().upper()


def normalize_imported_media(source_part: str, data: bytes) -> tuple[bytes, dict | None]:
    """Normalize an indexed PNG before it enters the composed OOXML package.

    Expanding an indexed palette removes one source of ambiguity while preserving
    visible pixels.  It does not promise deterministic Word/PDF rendering by
    itself: repeated page furniture should still use an asset sized sensibly for
    its rendered dimensions.  Other formats and already-expanded PNGs pass
    through byte-for-byte.
    """

    if posixpath.splitext(source_part)[1].lower() != ".png":
        return data, None
    try:
        with Image.open(io.BytesIO(data)) as image:
            if image.mode != "P":
                return data, None
            target_mode = "RGBA" if "transparency" in image.info else "RGB"
            expanded = image.convert(target_mode)
            output = io.BytesIO()
            save_options: dict[str, object] = {
                "format": "PNG",
                "compress_level": 9,
                "optimize": False,
            }
            if image.info.get("icc_profile"):
                save_options["icc_profile"] = image.info["icc_profile"]
            if image.info.get("dpi"):
                save_options["dpi"] = image.info["dpi"]
            expanded.save(output, **save_options)
    except (OSError, ValueError):
        return data, None
    normalized = output.getvalue()
    return normalized, {
        "source_part": source_part,
        "reason": "indexed-png-expanded-for-stable-word-rendering",
        "source_mode": "P",
        "target_mode": target_mode,
        "source_sha256": hash_bytes(data),
        "normalized_sha256": hash_bytes(normalized),
    }


def rels_part_for(owner_part: str) -> str:
    directory = posixpath.dirname(owner_part)
    filename = posixpath.basename(owner_part)
    return posixpath.join(directory, "_rels", filename + ".rels")


def resolve_target(owner_part: str, target: str) -> str:
    return posixpath.normpath(posixpath.join(posixpath.dirname(owner_part), target))


def relative_target(owner_part: str, target_part: str) -> str:
    return posixpath.relpath(target_part, posixpath.dirname(owner_part))


def text_value(element) -> str:
    values = etree.XPath(".//w:t/text()", namespaces=XNS)(element)
    return re.sub(r"\s+", " ", "".join(values)).strip()


def visible_text(element) -> str:
    return re.sub(r"\s+", " ", ROW_MARKER.sub("", text_value(element))).strip()


def _drawing_identity_nodes(root):
    """Return OOXML drawing identity nodes whose numeric IDs must not collide."""

    return etree.XPath(".//wp:docPr|.//pic:cNvPr", namespaces=XNS)(root)


class DocxPackage:
    def __init__(self, path: Path):
        self.path = Path(path).resolve()
        with zipfile.ZipFile(self.path) as archive:
            self.parts = {name: archive.read(name) for name in archive.namelist()}
        self._copy_cache: dict[tuple[str, str], str] = {}
        self.media_normalizations: list[dict] = []

    def xml(self, name: str):
        if name not in self.parts:
            raise KeyError(f"Package part missing: {name}")
        return etree.fromstring(self.parts[name])

    def set_xml(self, name: str, root) -> None:
        self.parts[name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")

    def write(self, output: Path) -> None:
        output = Path(output).resolve()
        output.parent.mkdir(parents=True, exist_ok=True)
        if output.exists():
            raise FileExistsError(f"Refusing to overwrite {output}")
        temporary = output.with_suffix(output.suffix + ".tmp")
        with zipfile.ZipFile(temporary, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for name in sorted(self.parts):
                archive.writestr(name, self.parts[name])
        temporary.replace(output)

    def relationship_root(self, owner_part: str):
        rels_name = rels_part_for(owner_part)
        if rels_name in self.parts:
            return rels_name, self.xml(rels_name)
        return rels_name, etree.Element(f"{{{PKG_REL_NS}}}Relationships", nsmap={None: PKG_REL_NS})

    @staticmethod
    def next_rid(root) -> str:
        used = set(root.xpath("./pr:Relationship/@Id", namespaces=XNS))
        index = 1
        while f"rId{index}" in used:
            index += 1
        return f"rId{index}"

    def add_relationship(self, owner_part: str, *, rel_type: str, target: str, target_mode: str | None = None) -> str:
        rels_name, root = self.relationship_root(owner_part)
        rid = self.next_rid(root)
        attributes = {"Id": rid, "Type": rel_type, "Target": target}
        if target_mode:
            attributes["TargetMode"] = target_mode
        etree.SubElement(root, f"{{{PKG_REL_NS}}}Relationship", **attributes)
        self.set_xml(rels_name, root)
        return rid

    def unique_part(self, desired: str, data: bytes) -> str:
        if desired not in self.parts or self.parts[desired] == data:
            return desired
        stem, extension = posixpath.splitext(desired)
        digest = hash_bytes(data)[:12].lower()
        candidate = f"{stem}_agdoc_{digest}{extension}"
        index = 2
        while candidate in self.parts and self.parts[candidate] != data:
            candidate = f"{stem}_agdoc_{digest}_{index}{extension}"
            index += 1
        return candidate

    def copy_content_type(self, source: "DocxPackage", source_part: str, target_part: str) -> None:
        source_types = source.xml(CONTENT_TYPES)
        target_types = self.xml(CONTENT_TYPES)
        overrides = source_types.xpath("./ct:Override[@PartName=$p]", namespaces=XNS, p="/" + source_part)
        if overrides:
            existing = target_types.xpath("./ct:Override[@PartName=$p]", namespaces=XNS, p="/" + target_part)
            if not existing:
                etree.SubElement(target_types, f"{{{CT_NS}}}Override", PartName="/" + target_part, ContentType=overrides[0].get("ContentType"))
                self.set_xml(CONTENT_TYPES, target_types)
            return
        extension = posixpath.splitext(source_part)[1].lstrip(".").lower()
        defaults = source_types.xpath(
            "./ct:Default[translate(@Extension,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')=$e]",
            namespaces=XNS,
            e=extension,
        )
        existing = target_types.xpath(
            "./ct:Default[translate(@Extension,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')=$e]",
            namespaces=XNS,
            e=extension,
        )
        if defaults and not existing:
            etree.SubElement(target_types, f"{{{CT_NS}}}Default", Extension=extension, ContentType=defaults[0].get("ContentType"))
            self.set_xml(CONTENT_TYPES, target_types)

    def find_identical_media(self, data: bytes, extension: str) -> str | None:
        digest = hash_bytes(data)
        for name, payload in self.parts.items():
            if name.startswith("word/media/") and posixpath.splitext(name)[1].lower() == extension.lower() and hash_bytes(payload) == digest:
                return name
        return None

    def remap_cloned_drawing_ids(self, data: bytes) -> bytes:
        """Give drawings in a cloned XML part fresh package-wide identities.

        A donor may point its default and even stories at the same physical
        header/footer part. Materialising independent parts is necessary, but
        copying their drawing IDs verbatim still leaves Word with two distinct
        stories containing the same drawing identities. Fresh IDs make the
        clone independent without changing its layout or visible design.
        """

        try:
            clone = etree.fromstring(data)
        except etree.XMLSyntaxError:
            return data
        clone_nodes = _drawing_identity_nodes(clone)
        if not clone_nodes:
            return data
        used: set[int] = set()
        for part_name, payload in self.parts.items():
            if not part_name.endswith(".xml"):
                continue
            try:
                root = etree.fromstring(payload)
            except etree.XMLSyntaxError:
                continue
            for node in _drawing_identity_nodes(root):
                try:
                    used.add(int(node.get("id", "")))
                except ValueError:
                    continue
        next_id = max(used, default=0) + 1
        for node in clone_nodes:
            while next_id in used:
                next_id += 1
            node.set("id", str(next_id))
            used.add(next_id)
            next_id += 1
        return etree.tostring(clone, xml_declaration=True, encoding="UTF-8", standalone="yes")

    def copy_part_from(self, source: "DocxPackage", source_part: str, *, force_unique: bool = False) -> str:
        cache_key = (str(source.path), source_part)
        if not force_unique and cache_key in self._copy_cache:
            return self._copy_cache[cache_key]
        if source_part not in source.parts:
            raise KeyError(f"Relationship target missing in donor: {source_part}")
        data = source.parts[source_part]
        extension = posixpath.splitext(source_part)[1]
        if source_part.startswith("word/media/"):
            data, normalization = normalize_imported_media(source_part, data)
            if normalization is not None:
                self.media_normalizations.append(
                    {
                        **normalization,
                        "source_document": str(source.path),
                    }
                )
            identical = self.find_identical_media(data, extension)
            if identical:
                self._copy_cache[cache_key] = identical
                return identical
            desired = posixpath.join("word/media", posixpath.basename(source_part))
        elif source_part.startswith("word/header"):
            desired = f"word/header_agdoc_{hash_bytes(data)[:12].lower()}.xml"
        elif source_part.startswith("word/footer"):
            desired = f"word/footer_agdoc_{hash_bytes(data)[:12].lower()}.xml"
        else:
            desired = source_part
        if force_unique and source_part.startswith(("word/header", "word/footer")):
            data = self.remap_cloned_drawing_ids(data)
        if force_unique and desired in self.parts:
            stem, suffix = posixpath.splitext(desired)
            index = 2
            target_part = f"{stem}_{index}{suffix}"
            while target_part in self.parts:
                index += 1
                target_part = f"{stem}_{index}{suffix}"
        else:
            target_part = self.unique_part(desired, data)
        if not force_unique:
            self._copy_cache[cache_key] = target_part
        self.parts[target_part] = data
        self.copy_content_type(source, source_part, target_part)
        source_rels_name = rels_part_for(source_part)
        if source_rels_name in source.parts:
            source_rels = source.xml(source_rels_name)
            target_rels = etree.Element(f"{{{PKG_REL_NS}}}Relationships", nsmap={None: PKG_REL_NS})
            for relationship in source_rels.xpath("./pr:Relationship", namespaces=XNS):
                attributes = {"Id": relationship.get("Id"), "Type": relationship.get("Type")}
                if relationship.get("TargetMode") == "External":
                    attributes.update({"Target": relationship.get("Target"), "TargetMode": "External"})
                else:
                    nested_source = resolve_target(source_part, relationship.get("Target"))
                    nested_target = self.copy_part_from(source, nested_source)
                    attributes["Target"] = relative_target(target_part, nested_target)
                etree.SubElement(target_rels, f"{{{PKG_REL_NS}}}Relationship", **attributes)
            self.set_xml(rels_part_for(target_part), target_rels)
        return target_part

    def clone_relationship(self, source: "DocxPackage", *, source_owner: str, source_rid: str, target_owner: str) -> str:
        _, source_rels = source.relationship_root(source_owner)
        matches = source_rels.xpath("./pr:Relationship[@Id=$rid]", namespaces=XNS, rid=source_rid)
        if len(matches) != 1:
            raise ValueError(f"Donor relationship {source_owner}#{source_rid} not found exactly once")
        relationship = matches[0]
        if relationship.get("TargetMode") == "External":
            return self.add_relationship(target_owner, rel_type=relationship.get("Type"), target=relationship.get("Target"), target_mode="External")
        source_part = resolve_target(source_owner, relationship.get("Target"))
        target_part = self.copy_part_from(source, source_part)
        return self.add_relationship(target_owner, rel_type=relationship.get("Type"), target=relative_target(target_owner, target_part))


def validate_docx(path: Path) -> dict:
    issues = []
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        if len(names) != len(set(names)):
            issues.append("Duplicate ZIP entries")
        for name in names:
            if name.endswith((".xml", ".rels")):
                try:
                    etree.fromstring(archive.read(name))
                except Exception as exc:
                    issues.append(f"XML parse failure {name}: {exc}")
        for rels_name in [name for name in names if name.endswith(".rels")]:
            root = etree.fromstring(archive.read(rels_name))
            owner = "" if rels_name == "_rels/.rels" else posixpath.join(posixpath.dirname(posixpath.dirname(rels_name)), posixpath.basename(rels_name)[:-5])
            for relationship in root.xpath("./pr:Relationship", namespaces=XNS):
                if relationship.get("TargetMode") == "External":
                    continue
                target = posixpath.normpath(posixpath.join(posixpath.dirname(owner), relationship.get("Target")))
                if target not in names:
                    issues.append(f"Missing relationship target: {rels_name}#{relationship.get('Id')} -> {target}")
    return {"valid": not issues, "issues": issues}


def require_one_sdt(root, tag: str):
    matches = find_sdts(root, tag)
    if len(matches) != 1:
        raise ValueError(f"Expected one content control tagged {tag}; found {len(matches)}")
    return matches[0]


def replace_sdt(target: DocxPackage, source: DocxPackage, target_tag: str, source_tag: str | None = None) -> None:
    source_tag = source_tag or target_tag
    target_root = target.xml(DOCUMENT_PART)
    source_root = source.xml(DOCUMENT_PART)
    target_sdt = require_one_sdt(target_root, target_tag)
    replacement = copy.deepcopy(require_one_sdt(source_root, source_tag))
    for node in replacement.iter():
        for attribute in (f"{{{R_NS}}}id", f"{{{R_NS}}}embed", f"{{{R_NS}}}link"):
            old_rid = node.get(attribute)
            if old_rid:
                node.set(attribute, target.clone_relationship(source, source_owner=DOCUMENT_PART, source_rid=old_rid, target_owner=DOCUMENT_PART))
    target_sdt.getparent().replace(target_sdt, replacement)
    target.set_xml(DOCUMENT_PART, target_root)


def remove_sdt(package: DocxPackage, tag: str) -> None:
    root = package.xml(DOCUMENT_PART)
    sdt = require_one_sdt(root, tag)
    sdt.getparent().remove(sdt)
    package.set_xml(DOCUMENT_PART, root)


def set_sdt_text(root, tag: str, value: str) -> None:
    sdt = require_one_sdt(root, tag)
    texts = etree.XPath(".//w:t", namespaces=XNS)(sdt)
    if not texts:
        raise ValueError(f"Field content control {tag} has no text node")
    texts[0].text = value
    for extra in texts[1:]:
        extra.text = ""


def set_cell_text(cell, value: str) -> None:
    for child in list(cell):
        if etree.QName(child).localname != "tcPr":
            cell.remove(child)
    paragraph = etree.SubElement(cell, f"{{{W_NS}}}p")
    run = etree.SubElement(paragraph, f"{{{W_NS}}}r")
    text = etree.SubElement(run, f"{{{W_NS}}}t")
    text.text = value


def set_labeled_table_value(block, label: str, value: str) -> None:
    matches = []
    for row in etree.XPath(".//w:tr", namespaces=XNS)(block):
        cells = etree.XPath("./w:tc", namespaces=XNS)(row)
        if len(cells) >= 2 and visible_text(cells[0]) == label:
            matches.append(cells[1])
    if len(matches) != 1:
        raise ValueError(f"Expected one row labelled {label!r}; found {len(matches)}")
    set_cell_text(matches[0], value)


def set_table_cell_value(block, row_index: int, column_index: int, value: str) -> None:
    tables = etree.XPath(".//w:tbl", namespaces=XNS)(block)
    if len(tables) != 1:
        raise ValueError(f"Expected one table in bound block; found {len(tables)}")
    rows = etree.XPath("./w:tr", namespaces=XNS)(tables[0])
    if row_index < 0 or row_index >= len(rows):
        raise ValueError(f"Bound table row {row_index} is outside 0..{len(rows) - 1}")
    cells = etree.XPath("./w:tc", namespaces=XNS)(rows[row_index])
    if column_index < 0 or column_index >= len(cells):
        raise ValueError(f"Bound table column {column_index} is outside 0..{len(cells) - 1} on row {row_index}")
    set_cell_text(cells[column_index], value)


PAGE_FURNITURE_KINDS = {"header", "footer"}
PAGE_FURNITURE_TYPES = {"default", "first", "even"}


def _normalized_page_furniture_kinds(kinds) -> set[str]:
    result = set(kinds or PAGE_FURNITURE_KINDS)
    unknown = result - PAGE_FURNITURE_KINDS
    if unknown:
        raise ValueError(f"Unsupported page-furniture kinds: {', '.join(sorted(unknown))}")
    return result


def _normalized_page_furniture_types(reference_types) -> set[str] | None:
    if reference_types is None:
        return None
    result = set(reference_types)
    unknown = result - PAGE_FURNITURE_TYPES
    if unknown:
        raise ValueError(f"Unsupported Word header/footer reference types: {', '.join(sorted(unknown))}")
    return result


def _selected_sections(root, section_numbers: list[int] | None):
    sections = etree.XPath(".//w:sectPr", namespaces=XNS)(root)
    if section_numbers is None:
        return list(enumerate(sections, start=1))
    requested = []
    seen = set()
    for raw in section_numbers:
        if not isinstance(raw, int) or isinstance(raw, bool) or raw < 1 or raw > len(sections):
            raise ValueError(f"Section number {raw!r} is outside 1..{len(sections)}")
        if raw not in seen:
            requested.append((raw, sections[raw - 1]))
            seen.add(raw)
    return requested


def page_furniture_inventory(package: DocxPackage) -> dict:
    """Return section/type/part evidence for every active header/footer reference."""
    root = package.xml(DOCUMENT_PART)
    _, relationships = package.relationship_root(DOCUMENT_PART)
    sections = []
    for section_number, section in _selected_sections(root, None):
        references = []
        for reference in section.xpath("./w:headerReference|./w:footerReference", namespaces=XNS):
            kind = "header" if etree.QName(reference).localname == "headerReference" else "footer"
            reference_type = reference.get(f"{{{W_NS}}}type", "default")
            rid = reference.get(f"{{{R_NS}}}id")
            matches = relationships.xpath("./pr:Relationship[@Id=$rid]", namespaces=XNS, rid=rid)
            if len(matches) != 1:
                references.append({"kind": kind, "type": reference_type, "relationship": rid, "part": None, "sha256": None})
                continue
            relationship = matches[0]
            part = None if relationship.get("TargetMode") == "External" else resolve_target(DOCUMENT_PART, relationship.get("Target"))
            payload = package.parts.get(part) if part else None
            visible = ""
            fields = []
            if part and payload is not None:
                part_root = package.xml(part)
                visible = visible_text(part_root)
                fields = [value.strip() for value in part_root.xpath(".//w:instrText/text()", namespaces=XNS) if value.strip()]
            references.append(
                {
                    "kind": kind,
                    "type": reference_type,
                    "relationship": rid,
                    "part": part,
                    "sha256": hash_bytes(payload) if payload is not None else None,
                    "visible_text": visible,
                    "fields": fields,
                }
            )
        sections.append(
            {
                "section": section_number,
                "different_first_page": bool(section.xpath("./w:titlePg", namespaces=XNS)),
                "references": references,
            }
        )
    settings_even_odd = False
    if "word/settings.xml" in package.parts:
        settings_even_odd = bool(package.xml("word/settings.xml").xpath("./w:evenAndOddHeaders", namespaces=XNS))
    return {"section_count": len(sections), "even_and_odd_headers": settings_even_odd, "sections": sections}


def clear_page_furniture(
    target: DocxPackage,
    *,
    section_numbers: list[int] | None = None,
    kinds=None,
    reference_types=None,
) -> dict:
    selected_kinds = _normalized_page_furniture_kinds(kinds)
    selected_types = _normalized_page_furniture_types(reference_types)
    root = target.xml(DOCUMENT_PART)
    removed = []
    for section_number, section in _selected_sections(root, section_numbers):
        for reference in list(section.xpath("./w:headerReference|./w:footerReference", namespaces=XNS)):
            kind = "header" if etree.QName(reference).localname == "headerReference" else "footer"
            reference_type = reference.get(f"{{{W_NS}}}type", "default")
            if kind in selected_kinds and (selected_types is None or reference_type in selected_types):
                section.remove(reference)
                removed.append({"section": section_number, "kind": kind, "type": reference_type})
    target.set_xml(DOCUMENT_PART, root)
    return {"removed": removed, "removed_count": len(removed)}


def prune_unreferenced_page_furniture(package: DocxPackage) -> dict:
    """Remove inactive document relationships and orphan header/footer parts."""
    document = package.xml(DOCUMENT_PART)
    active_rids = set(document.xpath(".//w:headerReference/@r:id|.//w:footerReference/@r:id", namespaces=XNS))
    rels_name, relationships = package.relationship_root(DOCUMENT_PART)
    removed_relationships = []
    candidate_parts = set()
    for relationship in list(relationships.xpath("./pr:Relationship", namespaces=XNS)):
        rel_type = relationship.get("Type", "")
        rid = relationship.get("Id")
        if rel_type.rsplit("/", 1)[-1] not in PAGE_FURNITURE_KINDS or rid in active_rids:
            continue
        if relationship.get("TargetMode") != "External":
            candidate_parts.add(resolve_target(DOCUMENT_PART, relationship.get("Target")))
        relationships.remove(relationship)
        removed_relationships.append(rid)
    package.set_xml(rels_name, relationships)

    remaining_targets = {
        resolve_target(DOCUMENT_PART, relationship.get("Target"))
        for relationship in relationships.xpath("./pr:Relationship", namespaces=XNS)
        if relationship.get("TargetMode") != "External"
    }
    removed_parts = []
    for part in sorted(candidate_parts - remaining_targets):
        if not (part.startswith("word/header") or part.startswith("word/footer")):
            continue
        if part in package.parts:
            del package.parts[part]
            removed_parts.append(part)
        part_rels = rels_part_for(part)
        if part_rels in package.parts:
            del package.parts[part_rels]
            removed_parts.append(part_rels)
    if removed_parts and CONTENT_TYPES in package.parts:
        content_types = package.xml(CONTENT_TYPES)
        removed_set = {"/" + name for name in removed_parts}
        for override in list(content_types.xpath("./ct:Override", namespaces=XNS)):
            if override.get("PartName") in removed_set:
                content_types.remove(override)
        package.set_xml(CONTENT_TYPES, content_types)
    return {
        "removed_relationships": removed_relationships,
        "removed_relationship_count": len(removed_relationships),
        "removed_parts": removed_parts,
        "removed_part_count": len(removed_parts),
    }


def apply_page_furniture(
    target: DocxPackage,
    source: DocxPackage,
    *,
    section_numbers: list[int] | None = None,
    kinds=None,
    reference_types=None,
    target_reference_type: str | None = None,
    donor_section: int = 1,
    require_reference_types: bool = False,
) -> dict:
    """Apply selected native header/footer references from one donor section.

    `section_numbers` and `donor_section` are one-based. Only matching kind/type
    references are replaced; callers may use `clear_page_furniture` first when
    they want replace-all rather than merge behavior.
    """
    selected_kinds = _normalized_page_furniture_kinds(kinds)
    selected_types = _normalized_page_furniture_types(reference_types)
    if target_reference_type is not None:
        if target_reference_type not in PAGE_FURNITURE_TYPES:
            raise ValueError(f"Unsupported target page-furniture type: {target_reference_type!r}")
        if selected_types is None or len(selected_types) != 1:
            raise ValueError("target_reference_type requires exactly one selected donor reference type")
    source_doc = source.xml(DOCUMENT_PART)
    source_rels = source.xml(DOCUMENT_RELS)
    sections = etree.XPath(".//w:sectPr", namespaces=XNS)(source_doc)
    if not sections:
        raise ValueError("Header/footer donor has no section properties")
    if not isinstance(donor_section, int) or isinstance(donor_section, bool) or donor_section < 1 or donor_section > len(sections):
        raise ValueError(f"Donor section {donor_section!r} is outside 1..{len(sections)} in {source.path}")
    references = []
    for reference in etree.XPath("./w:headerReference|./w:footerReference", namespaces=XNS)(sections[donor_section - 1]):
        kind = "header" if etree.QName(reference).localname == "headerReference" else "footer"
        reference_type = reference.get(f"{{{W_NS}}}type", "default")
        if kind in selected_kinds and (selected_types is None or reference_type in selected_types):
            references.append(reference)
    if require_reference_types and selected_types is not None:
        available = {
            ("header" if etree.QName(reference).localname == "headerReference" else "footer", reference.get(f"{{{W_NS}}}type", "default"))
            for reference in references
        }
        missing = sorted((kind, reference_type) for kind in selected_kinds for reference_type in selected_types if (kind, reference_type) not in available)
        if missing:
            display = ", ".join(f"{kind}:{reference_type}" for kind, reference_type in missing)
            raise ValueError(f"Donor {source.path} section {donor_section} has no requested references: {display}")
    if not references:
        raise ValueError(
            f"Donor {source.path} section {donor_section} has no selected "
            f"{','.join(sorted(selected_kinds))} references"
        )
    prepared = []
    for reference in references:
        kind = "header" if etree.QName(reference).localname == "headerReference" else "footer"
        reference_type = reference.get(f"{{{W_NS}}}type", "default")
        old_rid = reference.get(f"{{{R_NS}}}id")
        relationships = source_rels.xpath("./pr:Relationship[@Id=$rid]", namespaces=XNS, rid=old_rid)
        if len(relationships) != 1:
            raise ValueError(f"Header/footer donor relationship missing: {old_rid}")
        relationship = relationships[0]
        source_part = resolve_target(DOCUMENT_PART, relationship.get("Target"))
        # Header/footer reference types must own distinct physical parts even
        # when their XML is identical. Word can suppress an even/first story
        # when multiple section references share one part at document open.
        target_part = target.copy_part_from(source, source_part, force_unique=True)
        new_rid = target.add_relationship(DOCUMENT_PART, rel_type=relationship.get("Type"), target=relative_target(DOCUMENT_PART, target_part))
        prepared.append((kind, etree.QName(reference).localname, reference_type, target_reference_type or reference_type, new_rid, target_part))
    target_doc = target.xml(DOCUMENT_PART)
    applied = []
    for section_number, section in _selected_sections(target_doc, section_numbers):
        for kind, local_name, source_ref_type, target_ref_type, new_rid, target_part in prepared:
            for old in list(section):
                if etree.QName(old).localname == local_name and old.get(f"{{{W_NS}}}type", "default") == target_ref_type:
                    section.remove(old)
            reference = etree.Element(f"{{{W_NS}}}{local_name}")
            reference.set(f"{{{W_NS}}}type", target_ref_type)
            reference.set(f"{{{R_NS}}}id", new_rid)
            _insert_sectpr_child(section, reference)
            applied.append({"section": section_number, "kind": kind, "source_type": source_ref_type, "type": target_ref_type, "part": target_part})
    target.set_xml(DOCUMENT_PART, target_doc)
    return {
        "source": str(source.path),
        "donor_section": donor_section,
        "applied": applied,
        "applied_count": len(applied),
    }


def apply_header_footer(target: DocxPackage, source: DocxPackage) -> None:
    """Backward-compatible combined donor application to every section."""
    apply_page_furniture(target, source)


def section_number_for_tag(package: DocxPackage, tag: str) -> int:
    """Return the one-based Word section containing one tagged content control."""
    root = package.xml(DOCUMENT_PART)
    matches = find_sdts(root, tag)
    if len(matches) != 1:
        raise ValueError(f"Section selector tag {tag!r} must appear exactly once; found {len(matches)}")
    preceding = int(matches[0].xpath("count(preceding::w:sectPr)", namespaces=XNS))
    section_count = len(root.xpath(".//w:sectPr", namespaces=XNS))
    result = preceding + 1
    if result > section_count:
        raise ValueError(f"Section selector tag {tag!r} could not be mapped to a Word section")
    return result


def resolve_section_scope(package: DocxPackage, scope, *, layout: dict | None = None) -> list[int]:
    """Resolve a friendly section selector to one-based Word section numbers.

    Supported selectors are ``all``, ``front_matter``, ``main``, a section
    number, a list of section numbers, or an object using ``sections``,
    ``containing_tag``, ``from_tag`` and/or ``through_tag``.  This keeps the
    engine generic while profiles and recipes provide document semantics.
    """
    root = package.xml(DOCUMENT_PART)
    section_count = len(root.xpath(".//w:sectPr", namespaces=XNS))
    if section_count < 1:
        raise ValueError("Document has no Word sections")

    if scope is None or scope == "all":
        return list(range(1, section_count + 1))
    if scope == "front_matter":
        front_count = int((layout or {}).get("front_sections", 0))
        if front_count < 1:
            raise ValueError("Section scope 'front_matter' requires an applied page layout with front sections")
        return list(range(1, front_count + 1))
    if scope == "main":
        front_count = int((layout or {}).get("front_sections", 0))
        if not (layout or {}).get("applied"):
            return list(range(1, section_count + 1))
        return list(range(front_count + 1, section_count + 1))
    if isinstance(scope, int) and not isinstance(scope, bool):
        return [number for number, _ in _selected_sections(root, [scope])]
    if isinstance(scope, list):
        return [number for number, _ in _selected_sections(root, scope)]
    if not isinstance(scope, dict):
        raise ValueError("Section scope must be all, front_matter, main, a section number/list, or a selector object")

    allowed = {"sections", "containing_tag", "from_tag", "through_tag"}
    unknown = sorted(set(scope) - allowed)
    if unknown:
        raise ValueError(f"Unsupported section selector keys: {', '.join(unknown)}")
    if "sections" in scope:
        if len(scope) != 1:
            raise ValueError("Section selector 'sections' cannot be combined with tag range keys")
        return resolve_section_scope(package, scope["sections"], layout=layout)
    if "containing_tag" in scope:
        if len(scope) != 1:
            raise ValueError("Section selector 'containing_tag' cannot be combined with range keys")
        return [section_number_for_tag(package, str(scope["containing_tag"]))]
    start = section_number_for_tag(package, str(scope["from_tag"])) if scope.get("from_tag") else 1
    end = section_number_for_tag(package, str(scope["through_tag"])) if scope.get("through_tag") else section_count
    if start > end:
        raise ValueError(f"Section selector range starts at section {start} after ending section {end}")
    return list(range(start, end + 1))


def set_different_first_page(package: DocxPackage, section_numbers: list[int], enabled: bool) -> dict:
    root = package.xml(DOCUMENT_PART)
    changed = []
    for section_number, section in _selected_sections(root, section_numbers):
        existing = section.xpath("./w:titlePg", namespaces=XNS)
        before = bool(existing)
        for node in existing:
            section.remove(node)
        if enabled:
            _insert_sectpr_child(section, etree.Element(f"{{{W_NS}}}titlePg"))
        if before != bool(enabled):
            changed.append(section_number)
    package.set_xml(DOCUMENT_PART, root)
    return {"enabled": bool(enabled), "changed_sections": changed}


def set_even_and_odd_headers(package: DocxPackage, enabled: bool) -> dict:
    part = "word/settings.xml"
    if part not in package.parts:
        raise ValueError("Document has no word/settings.xml part")
    root = package.xml(part)
    existing = root.xpath("./w:evenAndOddHeaders", namespaces=XNS)
    before = bool(existing)
    for node in existing:
        root.remove(node)
    if enabled:
        root.insert(0, etree.Element(f"{{{W_NS}}}evenAndOddHeaders"))
    package.set_xml(part, root)
    return {"enabled": bool(enabled), "changed": before != bool(enabled)}


def set_do_not_compress_pictures(package: DocxPackage, enabled: bool = True) -> dict:
    """Control Word's package-level picture-compression policy explicitly."""

    part = "word/settings.xml"
    if part not in package.parts:
        raise ValueError("Document has no word/settings.xml part")
    root = package.xml(part)
    existing = root.xpath("./w:doNotCompressPictures", namespaces=XNS)
    before = bool(existing)
    for node in existing:
        root.remove(node)
    if enabled:
        node = etree.Element(f"{{{W_NS}}}doNotCompressPictures")
        node.set(f"{{{W_NS}}}val", "true")
        insertion_index = 0
        for index, child in enumerate(root):
            if etree.QName(child).localname in {"view", "zoom", "removePersonalInformation", "removeDateAndTime"}:
                insertion_index = index + 1
        root.insert(insertion_index, node)
    package.set_xml(part, root)
    return {"enabled": bool(enabled), "changed": before != bool(enabled)}


def normalize_footer_page_total_fields(package: DocxPackage, section_numbers: list[int], scope: str) -> dict:
    """Use NUMPAGES or SECTIONPAGES in footer parts referenced by selected sections."""
    if scope not in {"document", "section"}:
        raise ValueError("Footer page-count scope must be 'document' or 'section'")
    root = package.xml(DOCUMENT_PART)
    _, relationships = package.relationship_root(DOCUMENT_PART)
    parts = set()
    for _, section in _selected_sections(root, section_numbers):
        for reference in section.xpath("./w:footerReference", namespaces=XNS):
            rid = reference.get(f"{{{R_NS}}}id")
            rels = relationships.xpath("./pr:Relationship[@Id=$rid]", namespaces=XNS, rid=rid)
            if len(rels) != 1:
                raise ValueError(f"Footer relationship missing while normalizing page totals: {rid}")
            parts.add(resolve_target(DOCUMENT_PART, rels[0].get("Target")))
    desired = "SECTIONPAGES" if scope == "section" else "NUMPAGES"
    changes = 0
    for part in sorted(parts):
        footer = package.xml(part)
        for instruction in footer.xpath(".//w:instrText", namespaces=XNS):
            original = instruction.text or ""
            replacement = re.sub(r"\b(?:NUMPAGES|SECTIONPAGES)\b", desired, original, flags=re.IGNORECASE)
            if replacement != original:
                instruction.text = replacement
                changes += 1
        package.set_xml(part, footer)
    return {"scope": scope, "field": desired, "footer_parts": sorted(parts), "changes": changes}


def _insert_sectpr_child(section, child) -> None:
    """Insert one section-property child in WordprocessingML schema order."""
    order = {
        "headerReference": 10,
        "footerReference": 20,
        "footnotePr": 30,
        "endnotePr": 40,
        "type": 50,
        "pgSz": 60,
        "pgMar": 70,
        "paperSrc": 80,
        "pgBorders": 90,
        "lnNumType": 100,
        "pgNumType": 110,
        "cols": 120,
        "formProt": 130,
        "vAlign": 140,
        "noEndnote": 150,
        "titlePg": 160,
        "textDirection": 170,
        "bidi": 180,
        "rtlGutter": 190,
        "docGrid": 200,
        "printerSettings": 210,
        "sectPrChange": 220,
    }
    child_order = order.get(etree.QName(child).localname, 1000)
    insert_at = len(section)
    for index, existing in enumerate(section):
        if order.get(etree.QName(existing).localname, 1000) > child_order:
            insert_at = index
            break
    section.insert(insert_at, child)


def insert_section_property(section, child) -> None:
    """Public schema-ordering helper for the generic page-region adapter."""

    _insert_sectpr_child(section, child)


def apply_styles(target: DocxPackage, source: DocxPackage) -> None:
    for part in ("word/styles.xml", "word/theme/theme1.xml", "word/fontTable.xml", "word/numbering.xml"):
        if part in source.parts:
            target.parts[part] = source.parts[part]


def ensure_default_content_type(package: DocxPackage, extension: str, content_type: str) -> None:
    root = package.xml(CONTENT_TYPES)
    extension = extension.lstrip(".").lower()
    existing = root.xpath("./ct:Default[translate(@Extension,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')=$e]", namespaces=XNS, e=extension)
    if not existing:
        etree.SubElement(root, f"{{{CT_NS}}}Default", Extension=extension, ContentType=content_type)
        package.set_xml(CONTENT_TYPES, root)


def apply_figure(package: DocxPackage, slot_tag: str, image_path: Path, *, component_id: str, title: str, caption: str, alt_text: str) -> None:
    root = package.xml(DOCUMENT_PART)
    slot = require_one_sdt(root, slot_tag)
    blips = etree.XPath(".//a:blip[@r:embed]", namespaces=XNS)(slot)
    if len(blips) != 1:
        raise ValueError(f"Figure slot {slot_tag} must contain one image prototype; found {len(blips)}")
    data = Path(image_path).read_bytes()
    extension = Path(image_path).suffix.lower()
    safe_component = re.sub(r"[^A-Za-z0-9_.-]+", "_", component_id).strip("._") or "figure"
    media_part = package.unique_part(f"word/media/{safe_component}_{hash_bytes(data)[:12].lower()}{extension}", data)
    package.parts[media_part] = data
    content_type = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".gif": "image/gif", ".bmp": "image/bmp", ".svg": "image/svg+xml"}.get(extension, "application/octet-stream")
    ensure_default_content_type(package, extension, content_type)
    rid = package.add_relationship(DOCUMENT_PART, rel_type=REL_IMAGE, target=relative_target(DOCUMENT_PART, media_part))
    blips[0].set(f"{{{R_NS}}}embed", rid)
    for node in etree.XPath(".//wp:docPr|.//pic:cNvPr", namespaces=XNS)(slot):
        node.set("name", title)
        node.set("descr", alt_text)
    caption_paragraphs = [paragraph for paragraph in etree.XPath(".//w:p", namespaces=XNS)(slot) if paragraph.xpath("boolean(w:pPr/w:pStyle[@w:val='Caption'])", namespaces=XNS)]
    if caption_paragraphs:
        texts = etree.XPath(".//w:t", namespaces=XNS)(caption_paragraphs[0])
        suffix = next((node for node in reversed(texts) if (node.text or "").lstrip().startswith(("-", "–", "—"))), None)
        if suffix is not None:
            suffix.text = " - " + caption
        else:
            run = etree.SubElement(caption_paragraphs[0], f"{{{W_NS}}}r")
            text = etree.SubElement(run, f"{{{W_NS}}}t")
            text.text = " - " + caption
    package.set_xml(DOCUMENT_PART, root)


def document_usable_width(root) -> int:
    sections = etree.XPath(".//w:sectPr", namespaces=XNS)(root)
    if not sections:
        return 10440
    section = sections[-1]
    page = section.find(f"{{{W_NS}}}pgSz")
    margins = section.find(f"{{{W_NS}}}pgMar")
    if page is None or margins is None:
        return 10440
    width = int(page.get(f"{{{W_NS}}}w", "12240"))
    left = int(margins.get(f"{{{W_NS}}}left", "900"))
    right = int(margins.get(f"{{{W_NS}}}right", "900"))
    return max(3600, width - left - right)


def display_value(value, column: dict) -> str:
    if value is None:
        return ""
    format_name = column.get("format")
    if format_name == "integer" and isinstance(value, (int, float)):
        return f"{value:,.0f}"
    if isinstance(format_name, str) and format_name.startswith("decimal:") and isinstance(value, (int, float)):
        return f"{value:,.{int(format_name.split(':', 1)[1])}f}"
    if format_name == "yes_no" and isinstance(value, bool):
        return "Yes" if value else "No"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def calculate_widths(columns: list[dict], records: list[dict], total_width: int) -> list[int]:
    weights = []
    for index, column in enumerate(columns):
        if column.get("width") is not None:
            weights.append(max(1.0, float(column["width"])))
            continue
        samples = [str(column.get("heading", column.get("source", "")))]
        samples.extend(display_value(record["values"][index], column) for record in records[:100])
        longest = max((len(sample) for sample in samples), default=8)
        weights.append(float(min(40, max(7, longest))))
    count = len(weights)
    minimum = min(420, max(120, total_width // max(1, count * 2)))
    if minimum * count >= total_width:
        widths = [total_width // count for _ in weights]
    else:
        distributable = total_width - minimum * count
        widths = [minimum + int(distributable * weight / sum(weights)) for weight in weights]
    widths[-1] += total_width - sum(widths)
    if any(width <= 0 for width in widths):
        raise ValueError(f"Could not calculate positive widths for {count} table columns")
    return widths


def set_cell_margins(tc_pr, margin: int) -> None:
    margins = etree.SubElement(tc_pr, f"{{{W_NS}}}tcMar")
    for edge in ("top", "left", "bottom", "right"):
        node = etree.SubElement(margins, f"{{{W_NS}}}{edge}")
        node.set(f"{{{W_NS}}}w", str(margin))
        node.set(f"{{{W_NS}}}type", "dxa")


def add_table_cell(row, text: str, width: int, *, fill: str, color: str, bold: bool, align: str, font_name: str, font_size_pt: float, margin: int, span: int | None = None, hidden_marker: str | None = None):
    cell = etree.SubElement(row, f"{{{W_NS}}}tc")
    tc_pr = etree.SubElement(cell, f"{{{W_NS}}}tcPr")
    tc_width = etree.SubElement(tc_pr, f"{{{W_NS}}}tcW")
    tc_width.set(f"{{{W_NS}}}w", str(width))
    tc_width.set(f"{{{W_NS}}}type", "dxa")
    if span:
        grid_span = etree.SubElement(tc_pr, f"{{{W_NS}}}gridSpan")
        grid_span.set(f"{{{W_NS}}}val", str(span))
    shading = etree.SubElement(tc_pr, f"{{{W_NS}}}shd")
    shading.set(f"{{{W_NS}}}fill", fill)
    vertical = etree.SubElement(tc_pr, f"{{{W_NS}}}vAlign")
    vertical.set(f"{{{W_NS}}}val", "center")
    set_cell_margins(tc_pr, margin)
    paragraph = etree.SubElement(cell, f"{{{W_NS}}}p")
    p_pr = etree.SubElement(paragraph, f"{{{W_NS}}}pPr")
    justification = etree.SubElement(p_pr, f"{{{W_NS}}}jc")
    justification.set(f"{{{W_NS}}}val", align)
    spacing = etree.SubElement(p_pr, f"{{{W_NS}}}spacing")
    spacing.set(f"{{{W_NS}}}before", "0")
    spacing.set(f"{{{W_NS}}}after", "0")
    run = etree.SubElement(paragraph, f"{{{W_NS}}}r")
    r_pr = etree.SubElement(run, f"{{{W_NS}}}rPr")
    fonts = etree.SubElement(r_pr, f"{{{W_NS}}}rFonts")
    fonts.set(f"{{{W_NS}}}ascii", font_name)
    fonts.set(f"{{{W_NS}}}hAnsi", font_name)
    size = etree.SubElement(r_pr, f"{{{W_NS}}}sz")
    size.set(f"{{{W_NS}}}val", str(round(font_size_pt * 2)))
    run_color = etree.SubElement(r_pr, f"{{{W_NS}}}color")
    run_color.set(f"{{{W_NS}}}val", color)
    if bold:
        etree.SubElement(r_pr, f"{{{W_NS}}}b")
    value = etree.SubElement(run, f"{{{W_NS}}}t")
    value.text = text
    if hidden_marker:
        hidden_run = etree.SubElement(paragraph, f"{{{W_NS}}}r")
        hidden_pr = etree.SubElement(hidden_run, f"{{{W_NS}}}rPr")
        etree.SubElement(hidden_pr, f"{{{W_NS}}}vanish")
        hidden_text = etree.SubElement(hidden_run, f"{{{W_NS}}}t")
        hidden_text.text = hidden_marker
    return cell


def encoded_row_marker(row_id: str) -> str:
    encoded = base64.urlsafe_b64encode(row_id.encode("utf-8")).decode("ascii").rstrip("=")
    return f"[[AGDOCROW:{encoded}]]"


def decode_row_marker(marker: str) -> str:
    padding = "=" * (-len(marker) % 4)
    return base64.urlsafe_b64decode(marker + padding).decode("utf-8")


def build_native_table(dataset: dict, style: dict, total_width: int, empty_text: str) -> etree._Element:
    columns = dataset["columns"]
    records = dataset["records"]
    if not columns:
        raise ValueError("A managed table must have at least one selected column")
    widths = calculate_widths(columns, records, total_width)
    table = etree.Element(f"{{{W_NS}}}tbl")
    table_pr = etree.SubElement(table, f"{{{W_NS}}}tblPr")
    table_width = etree.SubElement(table_pr, f"{{{W_NS}}}tblW")
    table_width.set(f"{{{W_NS}}}w", str(total_width))
    table_width.set(f"{{{W_NS}}}type", "dxa")
    layout = etree.SubElement(table_pr, f"{{{W_NS}}}tblLayout")
    layout.set(f"{{{W_NS}}}type", "fixed")
    borders = etree.SubElement(table_pr, f"{{{W_NS}}}tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = etree.SubElement(borders, f"{{{W_NS}}}{edge}")
        border.set(f"{{{W_NS}}}val", "single")
        border.set(f"{{{W_NS}}}sz", "4")
        border.set(f"{{{W_NS}}}color", style["border_color"])
    grid = etree.SubElement(table, f"{{{W_NS}}}tblGrid")
    for width in widths:
        column = etree.SubElement(grid, f"{{{W_NS}}}gridCol")
        column.set(f"{{{W_NS}}}w", str(width))

    header = etree.SubElement(table, f"{{{W_NS}}}tr")
    header_pr = etree.SubElement(header, f"{{{W_NS}}}trPr")
    if style.get("repeat_header", True):
        etree.SubElement(header_pr, f"{{{W_NS}}}tblHeader")
    for index, column in enumerate(columns):
        add_table_cell(header, str(column.get("heading", column["source"])), widths[index], fill=style["header_fill"], color=style["header_text"], bold=True, align=column.get("header_align", "center"), font_name=style["font_name"], font_size_pt=style["font_size_pt"], margin=int(style["cell_margin_twips"]))

    if not records:
        row = etree.SubElement(table, f"{{{W_NS}}}tr")
        add_table_cell(row, empty_text, total_width, fill=style["body_fill"], color=style["text_color"], bold=False, align="left", font_name=style["font_name"], font_size_pt=style["font_size_pt"], margin=int(style["cell_margin_twips"]), span=len(columns))
        return table

    previous_group = object()
    for row_index, record in enumerate(records):
        group = record.get("group")
        if group not in (None, "") and group != previous_group:
            group_row = etree.SubElement(table, f"{{{W_NS}}}tr")
            add_table_cell(group_row, str(group), total_width, fill=style["group_fill"], color=style["group_text"], bold=True, align="left", font_name=style["font_name"], font_size_pt=style["font_size_pt"], margin=int(style["cell_margin_twips"]), span=len(columns))
            previous_group = group
        row = etree.SubElement(table, f"{{{W_NS}}}tr")
        if not style.get("allow_row_split", True):
            row_pr = etree.SubElement(row, f"{{{W_NS}}}trPr")
            etree.SubElement(row_pr, f"{{{W_NS}}}cantSplit")
        fill = style["alternate_fill"] if row_index % 2 else style["body_fill"]
        for column_index, column in enumerate(columns):
            value = record["values"][column_index]
            alignment = column.get("align") or ("right" if isinstance(value, (int, float)) and not isinstance(value, bool) else "left")
            marker = encoded_row_marker(record["id"]) if column_index == 0 and record.get("id") else None
            add_table_cell(row, display_value(value, column), widths[column_index], fill=fill, color=style["text_color"], bold=False, align=alignment, font_name=style["font_name"], font_size_pt=style["font_size_pt"], margin=int(style["cell_margin_twips"]), hidden_marker=marker)
    return table


def replace_table(package: DocxPackage, slot_tag: str, dataset: dict, style: dict, *, empty_behavior: str, empty_text: str) -> None:
    root = package.xml(DOCUMENT_PART)
    slot = require_one_sdt(root, slot_tag)
    tables = etree.XPath(".//w:tbl", namespaces=XNS)(slot)
    if len(tables) != 1:
        raise ValueError(f"Managed table slot {slot_tag} must contain one table; found {len(tables)}")
    if not dataset["records"] and empty_behavior == "error":
        raise ValueError(f"Managed table {slot_tag} selected zero rows")
    if not dataset["records"] and empty_behavior == "remove":
        tables[0].getparent().remove(tables[0])
    else:
        replacement = build_native_table(dataset, style, document_usable_width(root), empty_text)
        tables[0].getparent().replace(tables[0], replacement)
    package.set_xml(DOCUMENT_PART, root)


def extract_table_state(package: DocxPackage, slot_tag: str) -> dict:
    root = package.xml(DOCUMENT_PART)
    slot = require_one_sdt(root, slot_tag)
    tables = etree.XPath(".//w:tbl", namespaces=XNS)(slot)
    if len(tables) != 1:
        return {"headers": [], "records": [], "table_count": len(tables)}
    rows = etree.XPath("./w:tr", namespaces=XNS)(tables[0])
    if not rows:
        return {"headers": [], "records": [], "table_count": 1}
    headers = [visible_text(cell) for cell in etree.XPath("./w:tc", namespaces=XNS)(rows[0])]
    records = []
    for row in rows[1:]:
        cells = etree.XPath("./w:tc", namespaces=XNS)(row)
        if len(cells) != len(headers):
            continue
        raw = text_value(row)
        marker = ROW_MARKER.search(raw)
        row_id = decode_row_marker(marker.group(1)) if marker else None
        records.append({"id": row_id, "values": [visible_text(cell) for cell in cells]})
    return {"headers": headers, "records": records, "table_count": 1}


def extract_sdt_state(package: DocxPackage, slot_tag: str | None) -> dict:
    """Return reviewable state for a tagged block, or the complete body when tag is null."""
    root = package.xml(DOCUMENT_PART)
    slot = root.find(f"{{{W_NS}}}body") if slot_tag is None else require_one_sdt(root, slot_tag)
    if slot is None:
        raise ValueError("Word document has no body")
    tables = []
    for table in etree.XPath(".//w:tbl", namespaces=XNS)(slot):
        rows = []
        for row in etree.XPath("./w:tr", namespaces=XNS)(table):
            rows.append([visible_text(cell) for cell in etree.XPath("./w:tc", namespaces=XNS)(row)])
        tables.append(rows)

    _, relationships = package.relationship_root(DOCUMENT_PART)
    image_hashes = []
    for blip in etree.XPath(".//a:blip[@r:embed]", namespaces=XNS)(slot):
        rid = blip.get(f"{{{R_NS}}}embed")
        matches = relationships.xpath("./pr:Relationship[@Id=$rid]", namespaces=XNS, rid=rid)
        if len(matches) != 1:
            image_hashes.append({"relationship": "missing", "sha256": None})
            continue
        relationship = matches[0]
        if relationship.get("TargetMode") == "External":
            image_hashes.append({"relationship": "external", "sha256": None})
            continue
        image_part = resolve_target(DOCUMENT_PART, relationship.get("Target"))
        payload = package.parts.get(image_part)
        image_hashes.append({"relationship": "embedded", "sha256": hash_bytes(payload) if payload is not None else None})

    alt_text = sorted(
        {
            value.strip()
            for node in etree.XPath(".//wp:docPr|.//pic:cNvPr", namespaces=XNS)(slot)
            for value in (node.get("descr"),)
            if value and value.strip()
        }
    )

    def element_state(parent, path: str, attributes: tuple[str, ...] = ("val",)):
        node = parent.find(path, namespaces=XNS) if parent is not None else None
        if node is None:
            return None
        values = {
            name: node.get(f"{{{W_NS}}}{name}")
            for name in attributes
            if node.get(f"{{{W_NS}}}{name}") is not None
        }
        return values or True

    def border_state(parent, path: str):
        node = parent.find(path, namespaces=XNS) if parent is not None else None
        if node is None:
            return None
        return {
            etree.QName(side).localname: {
                name: side.get(f"{{{W_NS}}}{name}")
                for name in ("val", "sz", "space", "color", "themeColor")
                if side.get(f"{{{W_NS}}}{name}") is not None
            }
            for side in node
        }

    formatting_paragraphs = []
    for paragraph in slot.xpath(".//w:p", namespaces=XNS):
        properties = paragraph.find("./w:pPr", namespaces=XNS)
        runs = []
        for run in paragraph.xpath("./w:r | ./w:hyperlink/w:r", namespaces=XNS):
            run_properties = run.find("./w:rPr", namespaces=XNS)
            runs.append(
                {
                    "text": "".join(run.xpath(".//w:t/text() | .//w:delText/text()", namespaces=XNS)),
                    "style": element_state(run_properties, "./w:rStyle"),
                    "bold": element_state(run_properties, "./w:b"),
                    "italic": element_state(run_properties, "./w:i"),
                    "underline": element_state(run_properties, "./w:u", ("val", "color")),
                    "strike": element_state(run_properties, "./w:strike"),
                    "color": element_state(run_properties, "./w:color", ("val", "themeColor")),
                    "highlight": element_state(run_properties, "./w:highlight"),
                    "size": element_state(run_properties, "./w:sz"),
                    "font": element_state(
                        run_properties,
                        "./w:rFonts",
                        ("ascii", "hAnsi", "eastAsia", "cs", "asciiTheme", "hAnsiTheme"),
                    ),
                }
            )
        formatting_paragraphs.append(
            {
                "style": element_state(properties, "./w:pStyle"),
                "alignment": element_state(properties, "./w:jc"),
                "indent": element_state(properties, "./w:ind", ("left", "right", "firstLine", "hanging")),
                "spacing": element_state(
                    properties,
                    "./w:spacing",
                    ("before", "after", "line", "lineRule", "beforeAutospacing", "afterAutospacing"),
                ),
                "outline": element_state(properties, "./w:outlineLvl"),
                "list": bool(properties is not None and properties.xpath("./w:numPr/w:numId", namespaces=XNS)),
                "list_level": element_state(properties, "./w:numPr/w:ilvl"),
                "page_break_before": element_state(properties, "./w:pageBreakBefore"),
                "runs": runs,
            }
        )

    formatting_tables = []
    for table in slot.xpath(".//w:tbl", namespaces=XNS):
        table_properties = table.find("./w:tblPr", namespaces=XNS)
        cells = []
        for cell_properties in table.xpath(".//w:tc/w:tcPr", namespaces=XNS):
            cells.append(
                {
                    "width": element_state(cell_properties, "./w:tcW", ("w", "type")),
                    "shading": element_state(cell_properties, "./w:shd", ("fill",)),
                    "vertical": element_state(cell_properties, "./w:vAlign"),
                    "span": element_state(cell_properties, "./w:gridSpan"),
                    "borders": border_state(cell_properties, "./w:tcBorders"),
                }
            )
        formatting_tables.append(
            {
                "style": element_state(table_properties, "./w:tblStyle"),
                "width": element_state(table_properties, "./w:tblW", ("w", "type")),
                "alignment": element_state(table_properties, "./w:jc"),
                "layout": element_state(table_properties, "./w:tblLayout", ("type",)),
                "shading": element_state(table_properties, "./w:shd", ("fill",)),
                "borders": border_state(table_properties, "./w:tblBorders"),
                "cells": cells,
            }
        )

    tracked_changes = []
    for node in slot.xpath(".//w:ins | .//w:del | .//w:moveFrom | .//w:moveTo", namespaces=XNS):
        tracked_changes.append(
            {
                "kind": etree.QName(node).localname,
                "author": node.get(f"{{{W_NS}}}author"),
                "date": node.get(f"{{{W_NS}}}date"),
                "text": "".join(node.xpath(".//w:t/text() | .//w:delText/text()", namespaces=XNS)),
            }
        )
    comment_ids = sorted(
        set(
            slot.xpath(
                ".//w:commentRangeStart/@w:id | .//w:commentRangeEnd/@w:id | .//w:commentReference/@w:id",
                namespaces=XNS,
            )
        )
    )
    comments = []
    if comment_ids:
        matches = relationships.xpath("./pr:Relationship[@Type=$t]", namespaces=XNS, t=REL_COMMENTS)
        comments_root = None
        if len(matches) == 1 and matches[0].get("TargetMode") != "External":
            comments_part = resolve_target(DOCUMENT_PART, matches[0].get("Target"))
            if comments_part in package.parts:
                comments_root = package.xml(comments_part)
        for comment_id in comment_ids:
            definitions = (
                comments_root.xpath("./w:comment[@w:id=$value]", namespaces=XNS, value=comment_id)
                if comments_root is not None
                else []
            )
            if len(definitions) != 1:
                comments.append({"missing": True})
                continue
            comment = definitions[0]
            comments.append(
                {
                    "author": comment.get(f"{{{W_NS}}}author"),
                    "initials": comment.get(f"{{{W_NS}}}initials"),
                    "date": comment.get(f"{{{W_NS}}}date"),
                    "text": "".join(comment.xpath(".//w:t/text()", namespaces=XNS)),
                }
            )
    return {
        "visible_text": visible_text(slot),
        "tables": tables,
        "images": image_hashes,
        "alt_text": alt_text,
        "formatting": {"paragraphs": formatting_paragraphs, "tables": formatting_tables},
        "annotations": {"tracked_changes": tracked_changes, "comments": comments},
    }


def extract_figure_state(package: DocxPackage, slot_tag: str) -> dict:
    state = extract_sdt_state(package, slot_tag)
    state["figure_count"] = len(state["images"])
    return state


def extract_word_block_state(package: DocxPackage, slot_tag: str) -> dict:
    return extract_sdt_state(package, slot_tag)


def extract_document_body_state(package: DocxPackage) -> dict:
    return extract_sdt_state(package, None)


def set_core_property(package: DocxPackage, name: str, value: str) -> None:
    if "docProps/core.xml" not in package.parts:
        return
    root = package.xml("docProps/core.xml")
    namespace = DC_NS if name in {"title", "subject", "creator", "description"} else None
    if name in {"keywords", "category", "contentStatus"}:
        namespace = CP_NS
    if not namespace:
        raise ValueError(f"Unsupported core property binding: {name}")
    node = root.find(f"{{{namespace}}}{name}")
    if node is None:
        node = etree.SubElement(root, f"{{{namespace}}}{name}")
    node.text = value
    modified = root.find(f"{{{DCTERMS_NS}}}modified")
    if modified is not None:
        modified.text = utc_now()
    package.set_xml("docProps/core.xml", root)


def core_property_inventory(package: DocxPackage) -> dict[str, str]:
    """Return the human-facing core properties used by Word and PDF export."""
    if "docProps/core.xml" not in package.parts:
        return {}
    root = package.xml("docProps/core.xml")
    properties = {}
    for name, namespace in {
        "title": DC_NS,
        "subject": DC_NS,
        "creator": DC_NS,
        "description": DC_NS,
        "keywords": CP_NS,
        "category": CP_NS,
        "contentStatus": CP_NS,
    }.items():
        node = root.find(f"{{{namespace}}}{name}")
        if node is not None and node.text:
            properties[name] = node.text
    return properties


def json_hash(value) -> str:
    payload = json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(payload.encode("utf-8")).hexdigest().upper()


def load_component_states(package: DocxPackage) -> dict:
    """Return the internal refresh baselines; absence is a valid unmanaged file."""
    if "word/settings.xml" in package.parts:
        settings = package.xml("word/settings.xml")
        values = {
            node.get(f"{{{W_NS}}}name"): node.get(f"{{{W_NS}}}val", "")
            for node in settings.xpath("./w:docVars/w:docVar[starts-with(@w:name, $prefix)]", namespaces=XNS, prefix=STATE_DOCVAR_PREFIX)
        }
        numbered = sorted((name, value) for name, value in values.items() if name[len(STATE_DOCVAR_PREFIX):].isdigit())
        if numbered:
            encoded = "".join(value for _, value in numbered)
            try:
                payload = zlib.decompress(base64.urlsafe_b64decode(encoded.encode("ascii"))).decode("utf-8")
                states = json.loads(payload)
            except Exception as exc:
                raise ValueError(f"Invalid managed-component document variables: {exc}") from exc
            if not isinstance(states, dict):
                raise ValueError("Managed-component document variables do not contain an object")
            return states
    if STATE_PART not in package.parts:
        return {}
    root = package.xml(STATE_PART)
    if etree.QName(root).namespace != STATE_NS:
        raise ValueError(f"Unsupported managed-component state namespace: {etree.QName(root).namespace}")
    states = {}
    for element in root.findall(f"{{{STATE_NS}}}component"):
        component_id = element.get("id")
        if not component_id:
            continue
        try:
            states[component_id] = json.loads(element.text or "{}")
        except json.JSONDecodeError as exc:
            raise ValueError(f"Invalid managed-component state for {component_id}: {exc}") from exc
    return states


def save_component_states(package: DocxPackage, states: dict) -> None:
    """Embed refresh baselines in Word-native document variables and custom XML.

    Word document variables are the interoperability anchor because desktop Word
    preserves them during ordinary editing/saving.  The custom-XML copy keeps the
    package human-inspectable for tools that do not round-trip through Word.
    """
    if "word/settings.xml" not in package.parts:
        raise ValueError("The Word master has no settings.xml for managed-component document variables")
    settings = package.xml("word/settings.xml")
    doc_vars = settings.find(f"{{{W_NS}}}docVars")
    if doc_vars is None:
        doc_vars = etree.Element(f"{{{W_NS}}}docVars")
        settings.append(doc_vars)
    for node in list(doc_vars):
        if (node.get(f"{{{W_NS}}}name") or "").startswith(STATE_DOCVAR_PREFIX):
            doc_vars.remove(node)
    encoded = base64.urlsafe_b64encode(zlib.compress(canonical_json(states).encode("utf-8"), level=9)).decode("ascii")
    chunks = [encoded[index:index + STATE_DOCVAR_CHUNK] for index in range(0, len(encoded), STATE_DOCVAR_CHUNK)] or [""]
    for index, chunk in enumerate(chunks):
        node = etree.SubElement(doc_vars, f"{{{W_NS}}}docVar")
        node.set(f"{{{W_NS}}}name", f"{STATE_DOCVAR_PREFIX}{index:03d}")
        node.set(f"{{{W_NS}}}val", chunk)
    package.set_xml("word/settings.xml", settings)

    root = etree.Element(f"{{{STATE_NS}}}managedComponents", nsmap={"agdoc": STATE_NS})
    root.set("schema", "agentic-managed-components/v1")
    for component_id in sorted(states):
        element = etree.SubElement(root, f"{{{STATE_NS}}}component", id=component_id)
        element.text = canonical_json(states[component_id])
    package.set_xml(STATE_PART, root)

    rels_name, rels = package.relationship_root(DOCUMENT_PART)
    target = relative_target(DOCUMENT_PART, STATE_PART)
    matches = rels.xpath("./pr:Relationship[@Type=$t and @Target=$target]", namespaces=XNS, t=STATE_REL, target=target)
    if not matches:
        etree.SubElement(rels, f"{{{PKG_REL_NS}}}Relationship", Id=package.next_rid(rels), Type=STATE_REL, Target=target)
        package.set_xml(rels_name, rels)


def canonical_json(value) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), default=str)
