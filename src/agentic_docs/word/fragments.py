from __future__ import annotations

import copy
import hashlib
from pathlib import Path
from typing import Iterable

from lxml import etree

from ..diagnostics import DiagnosticBag
from ..errors import PackageError
from .ooxml import R_NS, W_NS, find_sdts, qn
from .package import (
    DOCUMENT_PART,
    DocxPackage,
    XNS,
    relative_target,
    require_one_sdt,
    resolve_target,
)


COMMENTS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
COMMENTS_PART = "word/comments.xml"


def _element_hash(element) -> str:
    return hashlib.sha256(etree.tostring(element, method="c14n")).hexdigest().upper()


def _relationship_attributes(element) -> list[tuple[str, str]]:
    result = []
    for node in element.iter():
        for attribute in (f"{{{R_NS}}}id", f"{{{R_NS}}}embed", f"{{{R_NS}}}link"):
            value = node.get(attribute)
            if value:
                result.append((attribute, value))
    return result


def _clone_relationships(target: DocxPackage, source: DocxPackage, element) -> None:
    for node in element.iter():
        for attribute in (f"{{{R_NS}}}id", f"{{{R_NS}}}embed", f"{{{R_NS}}}link"):
            old_rid = node.get(attribute)
            if old_rid:
                node.set(
                    attribute,
                    target.clone_relationship(
                        source,
                        source_owner=DOCUMENT_PART,
                        source_rid=old_rid,
                        target_owner=DOCUMENT_PART,
                    ),
                )


def _source_elements(source: DocxPackage, source_tag: str | None, allow_untagged: bool):
    root = source.xml(DOCUMENT_PART)
    if source_tag:
        matches = find_sdts(root, source_tag)
        if len(matches) == 1:
            content = matches[0].find(qn("w:sdtContent"))
            if content is None:
                raise PackageError(f"Word fragment tag {source_tag!r} has no content")
            return list(content), "tagged"
        if not allow_untagged:
            raise PackageError(
                f"Word fragment expected exactly one content control tagged {source_tag!r}; found {len(matches)}"
            )
    elif not allow_untagged:
        raise PackageError("Word fragment requires source_tag unless allow_untagged is true")

    body = root.find(qn("w:body"))
    if body is None:
        raise PackageError("Word fragment has no document body")
    return [child for child in body if etree.QName(child).localname != "sectPr"], "whole-body"


def _remove_section_properties(elements: Iterable) -> int:
    removed = 0
    for element in elements:
        for node in list(element.xpath(".//w:sectPr", namespaces=XNS)):
            parent = node.getparent()
            if parent is not None:
                parent.remove(node)
                removed += 1
    return removed


def _merge_styles(target: DocxPackage, source: DocxPackage, elements: Iterable, diagnostics: DiagnosticBag, location: str) -> None:
    part = "word/styles.xml"
    if part not in source.parts or part not in target.parts:
        return
    source_root = source.xml(part)
    target_root = target.xml(part)
    referenced = {
        value
        for element in elements
        for value in element.xpath(".//w:pStyle/@w:val | .//w:rStyle/@w:val | .//w:tblStyle/@w:val", namespaces=XNS)
    }
    copied: set[str] = set()
    visiting: set[str] = set()
    collisions: set[str] = set()
    missing: set[str] = set()

    def copy_style(style_id: str) -> None:
        if not style_id or style_id in copied or style_id in visiting:
            return
        visiting.add(style_id)
        source_matches = source_root.xpath("./w:style[@w:styleId=$value]", namespaces=XNS, value=style_id)
        if not source_matches:
            missing.add(style_id)
            visiting.remove(style_id)
            return
        source_style = source_matches[0]
        for dependency in source_style.xpath("./w:basedOn/@w:val | ./w:next/@w:val | ./w:link/@w:val", namespaces=XNS):
            copy_style(dependency)
        target_matches = target_root.xpath("./w:style[@w:styleId=$value]", namespaces=XNS, value=style_id)
        if not target_matches:
            target_root.append(copy.deepcopy(source_style))
        elif _element_hash(target_matches[0]) != _element_hash(source_style):
            collisions.add(style_id)
        copied.add(style_id)
        visiting.remove(style_id)

    for style_id in sorted(referenced):
        copy_style(style_id)
    if missing:
        diagnostics.warn(
            "WORD_STYLE_DEFINITIONS_MISSING",
            "Fragment references styles without source definitions: " + ", ".join(sorted(missing)),
            location=location,
        )
    if collisions:
        diagnostics.warn(
            "WORD_STYLE_COLLISIONS",
            "Fragment styles differ from the selected company/profile styles; target definitions were retained: "
            + ", ".join(sorted(collisions)),
            location=location,
            hint="Review the rendered pages; normalize the fragment only if the retained company styles produce an unintended result.",
        )
    target.set_xml(part, target_root)


def _merge_numbering(target: DocxPackage, source: DocxPackage, elements: Iterable, diagnostics: DiagnosticBag, location: str) -> None:
    part = "word/numbering.xml"
    referenced_nodes = [
        node
        for element in elements
        for node in element.xpath(".//w:numPr/w:numId", namespaces=XNS)
        if node.get(qn("w:val")) not in {None, "0"}
    ]
    if not referenced_nodes:
        return
    if part not in source.parts or part not in target.parts:
        diagnostics.warn(
            "WORD_NUMBERING_PART_MISSING",
            "Fragment contains numbered content but source or target numbering definitions are unavailable",
            location=location,
        )
        return
    source_root = source.xml(part)
    target_root = target.xml(part)
    target_num_ids = {int(value) for value in target_root.xpath("./w:num/@w:numId", namespaces=XNS)}
    target_abstract_ids = {int(value) for value in target_root.xpath("./w:abstractNum/@w:abstractNumId", namespaces=XNS)}
    next_num_id = max(target_num_ids, default=0) + 1
    next_abstract_id = max(target_abstract_ids, default=0) + 1
    mapping: dict[str, str] = {}
    abstract_mapping: dict[str, str] = {}

    for old_num_id in sorted({node.get(qn("w:val")) for node in referenced_nodes}, key=int):
        source_nums = source_root.xpath("./w:num[@w:numId=$value]", namespaces=XNS, value=old_num_id)
        if len(source_nums) != 1:
            diagnostics.warn(
                "WORD_NUMBERING_DEFINITION_MISSING",
                f"Fragment numbering id {old_num_id} has no unique source definition",
                location=location,
            )
            continue
        source_num = source_nums[0]
        abstract_values = source_num.xpath("./w:abstractNumId/@w:val", namespaces=XNS)
        if len(abstract_values) != 1:
            diagnostics.warn(
                "WORD_ABSTRACT_NUMBERING_MISSING",
                f"Fragment numbering id {old_num_id} has no unique abstract numbering reference",
                location=location,
            )
            continue
        old_abstract_id = abstract_values[0]
        if old_abstract_id not in abstract_mapping:
            source_abstracts = source_root.xpath(
                "./w:abstractNum[@w:abstractNumId=$value]", namespaces=XNS, value=old_abstract_id
            )
            if len(source_abstracts) != 1:
                diagnostics.warn(
                    "WORD_ABSTRACT_NUMBERING_DEFINITION_MISSING",
                    f"Fragment abstract numbering id {old_abstract_id} has no unique definition",
                    location=location,
                )
                continue
            new_abstract_id = str(next_abstract_id)
            next_abstract_id += 1
            abstract = copy.deepcopy(source_abstracts[0])
            abstract.set(qn("w:abstractNumId"), new_abstract_id)
            first_num = target_root.find(qn("w:num"))
            if first_num is None:
                target_root.append(abstract)
            else:
                target_root.insert(target_root.index(first_num), abstract)
            abstract_mapping[old_abstract_id] = new_abstract_id
        if old_abstract_id not in abstract_mapping:
            continue
        new_num_id = str(next_num_id)
        next_num_id += 1
        new_num = copy.deepcopy(source_num)
        new_num.set(qn("w:numId"), new_num_id)
        abstract_reference = new_num.find(qn("w:abstractNumId"))
        abstract_reference.set(qn("w:val"), abstract_mapping[old_abstract_id])
        target_root.append(new_num)
        mapping[old_num_id] = new_num_id

    for node in referenced_nodes:
        old_value = node.get(qn("w:val"))
        if old_value in mapping:
            node.set(qn("w:val"), mapping[old_value])
    target.set_xml(part, target_root)


def _comment_ids(elements: Iterable) -> set[str]:
    result: set[str] = set()
    for element in elements:
        result.update(
            value
            for value in element.xpath(
                ".//w:commentRangeStart/@w:id | .//w:commentRangeEnd/@w:id | .//w:commentReference/@w:id",
                namespaces=XNS,
            )
            if value is not None
        )
    return result


def _comment_part(package: DocxPackage) -> str | None:
    _, relationships = package.relationship_root(DOCUMENT_PART)
    matches = relationships.xpath("./pr:Relationship[@Type=$t]", namespaces=XNS, t=COMMENTS_REL)
    if len(matches) > 1:
        raise PackageError("Word document contains more than one comments relationship")
    if not matches:
        return None
    relationship = matches[0]
    if relationship.get("TargetMode") == "External":
        raise PackageError("Word comments relationship cannot be external")
    return resolve_target(DOCUMENT_PART, relationship.get("Target"))


def _merge_comments(
    target: DocxPackage,
    source: DocxPackage,
    elements: Iterable,
    diagnostics: DiagnosticBag,
    location: str,
) -> None:
    """Carry referenced legacy Word comments with a fragment and remap IDs."""

    elements = list(elements)
    referenced = _comment_ids(elements)
    if not referenced:
        return
    source_part = _comment_part(source)
    if source_part is None or source_part not in source.parts:
        diagnostics.warn(
            "WORD_COMMENTS_PART_MISSING",
            "Fragment contains comment markers but its comments part is unavailable",
            location=location,
        )
        return

    target_part = _comment_part(target)
    if target_part is None:
        target_part = COMMENTS_PART
        if target_part in target.parts:
            target_part = target.unique_part(COMMENTS_PART, source.parts[source_part])
        root = etree.Element(qn("w:comments"), nsmap={"w": W_NS})
        target.set_xml(target_part, root)
        target.copy_content_type(source, source_part, target_part)
        target.add_relationship(
            DOCUMENT_PART,
            rel_type=COMMENTS_REL,
            target=relative_target(DOCUMENT_PART, target_part),
        )

    source_root = source.xml(source_part)
    target_root = target.xml(target_part)
    used = {
        int(value)
        for value in target_root.xpath("./w:comment/@w:id", namespaces=XNS)
        if str(value).lstrip("-").isdigit()
    }
    next_id = max(used, default=-1) + 1
    mapping: dict[str, str] = {}
    for old_id in sorted(
        referenced,
        key=lambda item: (0, int(item)) if item.lstrip("-").isdigit() else (1, item),
    ):
        matches = source_root.xpath("./w:comment[@w:id=$value]", namespaces=XNS, value=old_id)
        if len(matches) != 1:
            diagnostics.warn(
                "WORD_COMMENT_DEFINITION_MISSING",
                f"Fragment comment id {old_id!r} has no unique source definition",
                location=location,
            )
            continue
        new_id = str(next_id)
        next_id += 1
        mapping[old_id] = new_id
        comment = copy.deepcopy(matches[0])
        comment.set(qn("w:id"), new_id)
        _clone_relationships_for_owner(
            target,
            source,
            comment,
            source_owner=source_part,
            target_owner=target_part,
        )
        target_root.append(comment)

    for element in elements:
        for node in element.xpath(
            ".//w:commentRangeStart | .//w:commentRangeEnd | .//w:commentReference",
            namespaces=XNS,
        ):
            old_id = node.get(qn("w:id"))
            if old_id in mapping:
                node.set(qn("w:id"), mapping[old_id])
    target.set_xml(target_part, target_root)


def _clone_relationships_for_owner(
    target: DocxPackage,
    source: DocxPackage,
    element,
    *,
    source_owner: str,
    target_owner: str,
) -> None:
    for node in element.iter():
        for attribute in (f"{{{R_NS}}}id", f"{{{R_NS}}}embed", f"{{{R_NS}}}link"):
            old_rid = node.get(attribute)
            if old_rid:
                node.set(
                    attribute,
                    target.clone_relationship(
                        source,
                        source_owner=source_owner,
                        source_rid=old_rid,
                        target_owner=target_owner,
                    ),
                )


def _remap_drawing_ids(target: DocxPackage, elements: Iterable) -> None:
    target_root = target.xml(DOCUMENT_PART)
    used = {
        int(value)
        for value in target_root.xpath(".//wp:docPr/@id", namespaces=XNS)
        if str(value).isdigit()
    }
    next_id = max(used, default=0) + 1
    for element in elements:
        for node in element.xpath(".//wp:docPr", namespaces=XNS):
            node.set("id", str(next_id))
            next_id += 1


def _remap_bookmarks(target: DocxPackage, elements: Iterable, component_id: str) -> None:
    target_root = target.xml(DOCUMENT_PART)
    used_ids = {
        int(value)
        for value in target_root.xpath(".//w:bookmarkStart/@w:id", namespaces=XNS)
        if str(value).isdigit()
    }
    used_names = set(target_root.xpath(".//w:bookmarkStart/@w:name", namespaces=XNS))
    next_id = max(used_ids, default=0) + 1
    id_mapping: dict[str, str] = {}
    prefix = "".join(char if char.isalnum() else "_" for char in component_id)[:24]
    for element in elements:
        for start in element.xpath(".//w:bookmarkStart", namespaces=XNS):
            old_id = start.get(qn("w:id"))
            if old_id is None:
                continue
            new_id = str(next_id)
            next_id += 1
            id_mapping[old_id] = new_id
            start.set(qn("w:id"), new_id)
            name = start.get(qn("w:name"))
            if name and name in used_names and not name.startswith("_"):
                candidate = f"{prefix}_{name}"[:40]
                suffix = 2
                while candidate in used_names:
                    candidate = f"{prefix}_{name}_{suffix}"[:40]
                    suffix += 1
                start.set(qn("w:name"), candidate)
                used_names.add(candidate)
            elif name:
                used_names.add(name)
        for end in element.xpath(".//w:bookmarkEnd", namespaces=XNS):
            old_id = end.get(qn("w:id"))
            if old_id in id_mapping:
                end.set(qn("w:id"), id_mapping[old_id])


def component_wrapper(component_id: str, elements: Iterable) -> etree._Element:
    wrapper = etree.Element(qn("w:sdt"))
    properties = etree.SubElement(wrapper, qn("w:sdtPr"))
    alias = etree.SubElement(properties, qn("w:alias"))
    alias.set(qn("w:val"), component_id)
    tag = etree.SubElement(properties, qn("w:tag"))
    tag.set(qn("w:val"), f"AGDOC.COMPONENT.{component_id}")
    identifier = etree.SubElement(properties, qn("w:id"))
    identifier.set(qn("w:val"), str(int(hashlib.sha256(component_id.encode()).hexdigest()[:7], 16)))
    content = etree.SubElement(wrapper, qn("w:sdtContent"))
    for element in elements:
        content.append(element)
    return wrapper


def import_word_fragment(
    target: DocxPackage,
    source_path: Path,
    *,
    component_id: str,
    source_tag: str | None,
    allow_untagged: bool,
    preserve_sections: bool,
    diagnostics: DiagnosticBag,
) -> etree._Element:
    source = DocxPackage(source_path)
    source_elements, mode = _source_elements(source, source_tag, allow_untagged)
    elements = [copy.deepcopy(element) for element in source_elements]
    if mode == "whole-body":
        diagnostics.warn(
            "WORD_FRAGMENT_UNTAGGED_FALLBACK",
            f"Component {component_id!r} imported the entire Word body because no usable source tag was selected",
            location=str(source_path),
            hint="This is supported for real-world intake; add a root content-control tag when precise component extraction is needed.",
        )
    if not preserve_sections:
        removed = _remove_section_properties(elements)
        if removed:
            diagnostics.info(
                "WORD_FRAGMENT_SECTIONS_NORMALIZED",
                f"Removed {removed} embedded section definition(s) from component {component_id!r} so document-level page regions remain authoritative",
                location=str(source_path),
            )
    _merge_styles(target, source, elements, diagnostics, str(source_path))
    _merge_numbering(target, source, elements, diagnostics, str(source_path))
    _merge_comments(target, source, elements, diagnostics, str(source_path))
    _remap_drawing_ids(target, elements)
    _remap_bookmarks(target, elements, component_id)
    for element in elements:
        _clone_relationships(target, source, element)
    return component_wrapper(component_id, elements)


def replace_sequence_slot(target: DocxPackage, slot_tag: str, wrappers: Iterable[etree._Element]) -> None:
    root = target.xml(DOCUMENT_PART)
    slot = require_one_sdt(root, slot_tag)
    content = slot.find(qn("w:sdtContent"))
    if content is None:
        raise PackageError(f"Sequence slot {slot_tag!r} has no content container")
    for child in list(content):
        content.remove(child)
    for wrapper in wrappers:
        content.append(wrapper)
    target.set_xml(DOCUMENT_PART, root)


def replace_nested_slot(wrapper: etree._Element, slot_tag: str, wrappers: Iterable[etree._Element]) -> None:
    """Replace one exact Word content-control slot inside a component wrapper."""

    matches = find_sdts(wrapper, slot_tag)
    if len(matches) != 1:
        raise PackageError(f"Nested Word slot {slot_tag!r} was expected exactly once; found {len(matches)}")
    content = matches[0].find(qn("w:sdtContent"))
    if content is None:
        raise PackageError(f"Nested Word slot {slot_tag!r} has no content container")
    for child in list(content):
        content.remove(child)
    for child_wrapper in wrappers:
        content.append(child_wrapper)
