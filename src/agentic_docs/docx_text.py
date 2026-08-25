from __future__ import annotations

import os
import re
import tempfile
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path

from lxml import etree

from .errors import PackageError


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
NS = {"w": W_NS}
TEXT_PART = re.compile(
    r"^word/(?:document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$"
)


@dataclass(frozen=True)
class TextReplacement:
    find: str
    replace: str
    case_sensitive: bool = True
    whole_word: bool = False


@dataclass(frozen=True)
class ReplacementSummary:
    source: str
    output: str
    replacements: list[dict]
    total_replacements: int
    paragraphs_changed: int
    parts_changed: list[str]

    def to_dict(self) -> dict:
        return asdict(self)


def _pattern(spec: TextReplacement) -> re.Pattern[str]:
    if not spec.find:
        raise ValueError("Replacement search text cannot be empty")
    expression = re.escape(spec.find)
    if spec.whole_word:
        expression = rf"(?<!\w){expression}(?!\w)"
    flags = 0 if spec.case_sensitive else re.IGNORECASE
    return re.compile(expression, flags)


def _set_text(node: etree._Element, value: str) -> None:
    node.text = value
    space_key = f"{{{XML_NS}}}space"
    if value.startswith((" ", "\t", "\n")) or value.endswith((" ", "\t", "\n")):
        node.set(space_key, "preserve")
    else:
        node.attrib.pop(space_key, None)


def _replace_in_nodes(nodes: list[etree._Element], spec: TextReplacement) -> int:
    values = [node.text or "" for node in nodes]
    starts: list[int] = []
    total = 0
    for value in values:
        starts.append(total)
        total += len(value)
    combined = "".join(values)
    matches = list(_pattern(spec).finditer(combined))
    if not matches:
        return 0

    def node_for(character_index: int) -> int:
        # The documents handled here are small enough that a linear lookup is
        # clearer and safer than relying on edge-sensitive bisect arithmetic.
        for index in range(len(starts) - 1, -1, -1):
            if starts[index] <= character_index:
                return index
        return 0

    for match in reversed(matches):
        start, end = match.span()
        start_node = node_for(start)
        end_node = node_for(end - 1)
        start_offset = start - starts[start_node]
        end_offset = end - starts[end_node]
        current_start = nodes[start_node].text or ""
        current_end = nodes[end_node].text or ""
        prefix = current_start[:start_offset]
        suffix = current_end[end_offset:]
        if start_node == end_node:
            _set_text(nodes[start_node], prefix + spec.replace + suffix)
        else:
            _set_text(nodes[start_node], prefix + spec.replace)
            for index in range(start_node + 1, end_node):
                _set_text(nodes[index], "")
            _set_text(nodes[end_node], suffix)
    return len(matches)


def replace_docx_text(
    source: Path,
    output: Path,
    replacements: list[TextReplacement],
    *,
    all_stories: bool = False,
) -> ReplacementSummary:
    """Apply deterministic native-text replacements without rebuilding the DOCX."""

    source = Path(source).resolve()
    output = Path(output).resolve()
    if source == output:
        raise PackageError("DOCX text replacement requires different source and output paths")
    if not source.is_file() or source.suffix.lower() != ".docx":
        raise PackageError(f"Canonical Word component is not a DOCX: {source}")
    if not replacements:
        raise ValueError("At least one replacement is required")

    output.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=output.stem + ".",
        suffix=".tmp.docx",
        dir=output.parent,
    )
    os.close(descriptor)
    temporary = Path(temporary_name)
    replacement_counts = [0 for _ in replacements]
    changed_paragraphs = 0
    changed_parts: list[str] = []
    try:
        with zipfile.ZipFile(source, "r") as source_zip, zipfile.ZipFile(temporary, "w") as target_zip:
            if source_zip.testzip() is not None:
                raise PackageError(f"Source DOCX contains a corrupt ZIP member: {source}")
            for info in source_zip.infolist():
                payload = source_zip.read(info.filename)
                selected = info.filename == "word/document.xml" or (
                    all_stories and TEXT_PART.fullmatch(info.filename)
                )
                if selected:
                    try:
                        root = etree.fromstring(payload)
                    except etree.XMLSyntaxError as exc:
                        raise PackageError(f"Malformed Word XML part {info.filename}: {exc}") from exc
                    part_changed = False
                    for paragraph in root.xpath(".//w:p", namespaces=NS):
                        nodes = paragraph.xpath(
                            ".//w:t[not(ancestor::w:del)]",
                            namespaces=NS,
                        )
                        if not nodes:
                            continue
                        paragraph_changed = False
                        for index, spec in enumerate(replacements):
                            count = _replace_in_nodes(nodes, spec)
                            replacement_counts[index] += count
                            paragraph_changed = paragraph_changed or count > 0
                        if paragraph_changed:
                            changed_paragraphs += 1
                            part_changed = True
                    if part_changed:
                        changed_parts.append(info.filename)
                        payload = etree.tostring(
                            root,
                            xml_declaration=True,
                            encoding="UTF-8",
                            standalone=True,
                        )
                target_zip.writestr(info, payload)

        with zipfile.ZipFile(temporary, "r") as check_zip:
            bad_member = check_zip.testzip()
            if bad_member is not None:
                raise PackageError(f"Edited DOCX contains a corrupt ZIP member: {bad_member}")
            if "word/document.xml" not in check_zip.namelist():
                raise PackageError("Edited DOCX is missing word/document.xml")
        os.replace(temporary, output)
    finally:
        if temporary.exists():
            temporary.unlink()

    details = []
    for spec, count in zip(replacements, replacement_counts, strict=True):
        details.append(
            {
                "find": spec.find,
                "replace": spec.replace,
                "case_sensitive": spec.case_sensitive,
                "whole_word": spec.whole_word,
                "count": count,
            }
        )
    return ReplacementSummary(
        source=str(source),
        output=str(output),
        replacements=details,
        total_replacements=sum(replacement_counts),
        paragraphs_changed=changed_paragraphs,
        parts_changed=changed_parts,
    )


def replace_plain_text(
    source: Path,
    output: Path,
    replacements: list[TextReplacement],
) -> ReplacementSummary:
    """Apply the same checked replacement semantics to a UTF-8 canonical text source."""

    source = Path(source).resolve()
    output = Path(output).resolve()
    if source == output:
        raise PackageError("Plain-text replacement requires different source and output paths")
    if not source.is_file():
        raise PackageError(f"Canonical text source does not exist: {source}")
    if not replacements:
        raise ValueError("At least one replacement is required")
    original = source.read_text(encoding="utf-8-sig")
    value = original
    details = []
    for spec in replacements:
        value, count = _pattern(spec).subn(lambda _match: spec.replace, value)
        details.append(
            {
                "find": spec.find,
                "replace": spec.replace,
                "case_sensitive": spec.case_sensitive,
                "whole_word": spec.whole_word,
                "count": count,
            }
        )
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = output.with_suffix(output.suffix + ".tmp")
    temporary.write_text(value, encoding="utf-8")
    os.replace(temporary, output)
    before_lines = original.splitlines()
    after_lines = value.splitlines()
    changed_lines = sum(
        before != after
        for before, after in zip(before_lines, after_lines, strict=False)
    ) + abs(len(before_lines) - len(after_lines))
    total = sum(item["count"] for item in details)
    return ReplacementSummary(
        source=str(source),
        output=str(output),
        replacements=details,
        total_replacements=total,
        paragraphs_changed=changed_lines,
        parts_changed=[source.name] if total else [],
    )
