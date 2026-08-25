from __future__ import annotations

import hashlib
import re
import shlex
from dataclasses import dataclass, field
from typing import Any

from ..diagnostics import DiagnosticBag
from ..errors import PackageError


HEADING = re.compile(r"^(#{1,6})[ \t]+(.+?)[ \t]*$")
HEADING_ID = re.compile(r"^(.*?)[ \t]+\{#([A-Za-z0-9][A-Za-z0-9._-]*)\}[ \t]*$")
LIST_ITEM = re.compile(r"^(?P<indent>[ \t]*)(?P<marker>(?:[-+*])|(?:\d+[.)]))[ \t]+(?P<text>.+)$")
TABLE_RULE_CELL = re.compile(r"^:?-{3,}:?$")
SAFE_NAME = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]*$")
MARKDOWN_IMAGE = re.compile(r"!\[[^\]]*\]\([^\n)]*\)")


@dataclass(frozen=True)
class Inline:
    kind: str
    text: str = ""
    target: str | None = None
    children: tuple["Inline", ...] = ()


@dataclass
class Block:
    kind: str
    value: Any
    start_line: int = 1
    end_line: int = 1
    block_id: str | None = None


@dataclass
class MarkdownDocument:
    blocks: list[Block]
    slot_names: list[str]
    source_map: list[dict[str, Any]] = field(default_factory=list)


def _append_text(nodes: list[Inline], value: str) -> None:
    if not value:
        return
    if nodes and nodes[-1].kind == "text":
        previous = nodes[-1]
        nodes[-1] = Inline("text", previous.text + value)
    else:
        nodes.append(Inline("text", value))


def parse_inlines(text: str) -> list[Inline]:
    """Parse the deliberately supported Markdown inline subset into nodes."""

    nodes: list[Inline] = []
    index = 0
    pairs = (("**", "strong"), ("__", "strong"), ("~~", "strike"), ("*", "emphasis"), ("_", "emphasis"))
    while index < len(text):
        if text[index] == "\\" and index + 1 < len(text):
            _append_text(nodes, text[index + 1])
            index += 2
            continue
        if text[index] == "\n":
            nodes.append(Inline("hard_break"))
            index += 1
            continue
        if text[index] == "`":
            width = 1
            while index + width < len(text) and text[index + width] == "`":
                width += 1
            marker = "`" * width
            closing = text.find(marker, index + width)
            if closing >= 0:
                nodes.append(Inline("code", text[index + width : closing]))
                index = closing + width
                continue
        if text[index] == "[":
            close_label = text.find("](", index + 1)
            if close_label >= 0:
                close_target = text.find(")", close_label + 2)
                if close_target >= 0:
                    label = text[index + 1 : close_label]
                    target = text[close_label + 2 : close_target].strip()
                    nodes.append(Inline("link", target=target, children=tuple(parse_inlines(label))))
                    index = close_target + 1
                    continue
        if text[index] == "<":
            close = text.find(">", index + 1)
            if close >= 0:
                target = text[index + 1 : close].strip()
                if target.startswith(("http://", "https://", "mailto:")):
                    nodes.append(Inline("link", target=target, children=(Inline("text", target),)))
                    index = close + 1
                    continue
        matched = False
        for marker, kind in pairs:
            if not text.startswith(marker, index):
                continue
            closing = text.find(marker, index + len(marker))
            if closing <= index + len(marker):
                continue
            content = text[index + len(marker) : closing]
            nodes.append(Inline(kind, children=tuple(parse_inlines(content))))
            index = closing + len(marker)
            matched = True
            break
        if matched:
            continue
        next_special = min(
            (position for character in "\\\n`[*_~<" for position in [text.find(character, index + 1)] if position >= 0),
            default=len(text),
        )
        _append_text(nodes, text[index:next_special])
        index = next_special
    return nodes


def _split_table_row(line: str) -> list[str]:
    value = line.strip()
    if value.startswith("|"):
        value = value[1:]
    if value.endswith("|") and not value.endswith("\\|"):
        value = value[:-1]
    cells: list[str] = []
    current: list[str] = []
    escaped = False
    code = False
    for character in value:
        if escaped:
            current.append(character)
            escaped = False
            continue
        if character == "\\":
            escaped = True
            continue
        if character == "`":
            code = not code
            current.append(character)
            continue
        if character == "|" and not code:
            cells.append("".join(current).strip())
            current = []
        else:
            current.append(character)
    if escaped:
        current.append("\\")
    cells.append("".join(current).strip())
    return cells


def _table_alignment(rule: str) -> str:
    value = rule.strip()
    if value.startswith(":") and value.endswith(":"):
        return "center"
    if value.endswith(":"):
        return "right"
    return "left"


def _is_table_rule(line: str) -> bool:
    cells = _split_table_row(line)
    return bool(cells) and all(TABLE_RULE_CELL.fullmatch(cell.replace(" ", "")) for cell in cells)


def _paragraph_text(lines: list[str]) -> str:
    result = ""
    previous_hard_break = False
    for line in lines:
        value = line.strip()
        if result:
            result += "\n" if previous_hard_break else " "
        result += value
        previous_hard_break = line.endswith("  ")
    return result


def _parse_directive_header(value: str, location: str, line_number: int) -> tuple[str, list[str], dict[str, str]]:
    try:
        tokens = shlex.split(value)
    except ValueError as exc:
        raise PackageError(f"Invalid Markdown+ directive at {location}:{line_number}: {exc}") from exc
    if not tokens:
        return "", [], {}
    positional: list[str] = []
    attributes: dict[str, str] = {}
    for token in tokens[1:]:
        if "=" in token:
            key, item = token.split("=", 1)
            attributes[key.strip()] = item.strip()
        else:
            positional.append(token)
    return tokens[0].lower(), positional, attributes


def _assign_ids(blocks: list[Block], location: str) -> None:
    seen: dict[str, int] = {}
    explicit: set[str] = set()
    for block in blocks:
        requested = None
        if isinstance(block.value, dict):
            requested = block.value.get("id")
        payload = f"{block.kind}\n{block.value!r}".encode("utf-8")
        base = requested or f"{block.kind}-{hashlib.sha256(payload).hexdigest()[:10]}"
        if requested:
            if requested in explicit:
                raise PackageError(f"Duplicate explicit Markdown block id {requested!r} in {location}")
            explicit.add(requested)
        seen[base] = seen.get(base, 0) + 1
        block.block_id = base if seen[base] == 1 else f"{base}-{seen[base]}"
        if block.kind == "callout":
            _assign_ids(block.value["blocks"], location)


def _collect_slots(blocks: list[Block]) -> list[str]:
    result: list[str] = []
    for block in blocks:
        if block.kind == "insert":
            result.append(str(block.value["slot"]))
        elif block.kind == "callout":
            result.extend(_collect_slots(block.value["blocks"]))
    return result


def parse_document(
    text: str,
    diagnostics: DiagnosticBag,
    location: str,
    *,
    strict: bool = True,
    line_offset: int = 0,
) -> MarkdownDocument:
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    blocks: list[Block] = []
    index = 0

    def add(kind: str, value: Any, start: int, end: int) -> None:
        blocks.append(Block(kind, value, line_offset + start + 1, line_offset + end + 1))

    while index < len(lines):
        line = lines[index]
        stripped = line.strip()
        if not stripped:
            index += 1
            continue

        if MARKDOWN_IMAGE.search(line):
            message = (
                f"Markdown images are not supported at {location}:{line_offset + index + 1}; "
                "declare a figure component and place it with :::insert"
            )
            if strict:
                raise PackageError(message)
            diagnostics.warn("MARKDOWN_IMAGE_LITERAL", message, location=location)

        if stripped.startswith(("```", "~~~")):
            start = index
            marker = stripped[:3]
            language = stripped[3:].strip()
            index += 1
            content: list[str] = []
            while index < len(lines) and not lines[index].strip().startswith(marker):
                content.append(lines[index])
                index += 1
            if index >= len(lines):
                message = f"Unclosed fenced code block at {location}:{line_offset + start + 1}"
                if strict:
                    raise PackageError(message)
                diagnostics.warn("MARKDOWN_UNCLOSED_FENCE", message, location=location)
                end = len(lines) - 1
            else:
                end = index
                index += 1
            add("code", {"language": language, "text": "\n".join(content)}, start, end)
            continue

        heading = HEADING.match(line)
        if heading:
            value = heading.group(2)
            explicit_id = None
            identifier = HEADING_ID.match(value)
            if identifier:
                value = identifier.group(1).strip()
                explicit_id = identifier.group(2)
            add("heading", {"level": len(heading.group(1)), "text": value, "id": explicit_id}, index, index)
            index += 1
            continue

        if stripped.startswith(":::"):
            start = index
            directive, positional, attributes = _parse_directive_header(
                stripped[3:].strip(), location, line_offset + index + 1
            )
            if directive == "page-break":
                add("page_break", None, index, index)
                index += 1
                continue
            if directive in {"insert", "component"}:
                slot = attributes.get("slot") or attributes.get("id") or (positional[0] if positional else "")
                if not SAFE_NAME.fullmatch(slot):
                    raise PackageError(
                        f"Markdown+ insert at {location}:{line_offset + index + 1} requires a safe slot name"
                    )
                add("insert", {"slot": slot}, index, index)
                index += 1
                continue
            if directive in {"note", "warning", "important", "callout"}:
                role = attributes.get("role") or (directive if directive != "callout" else "note")
                if not SAFE_NAME.fullmatch(role):
                    raise PackageError(f"Invalid callout role {role!r} at {location}:{line_offset + index + 1}")
                index += 1
                inner: list[str] = []
                while index < len(lines) and lines[index].strip() != ":::":
                    inner.append(lines[index])
                    index += 1
                if index >= len(lines):
                    raise PackageError(f"Unclosed Markdown+ {directive} block at {location}:{line_offset + start + 1}")
                nested = parse_document(
                    "\n".join(inner),
                    diagnostics,
                    location,
                    strict=strict,
                    line_offset=line_offset + start + 1,
                )
                add("callout", {"role": role, "blocks": nested.blocks}, start, index)
                index += 1
                continue
            message = f"Unsupported Markdown+ directive {stripped!r} at {location}:{line_offset + index + 1}"
            if strict:
                raise PackageError(message)
            diagnostics.warn("MARKDOWN_UNKNOWN_DIRECTIVE", message, location=location)
            add("paragraph", stripped, index, index)
            index += 1
            continue

        list_match = LIST_ITEM.match(line)
        if list_match:
            start = index
            base_indent = len(list_match.group("indent").expandtabs(4))
            items: list[dict[str, Any]] = []
            while index < len(lines):
                match = LIST_ITEM.match(lines[index])
                if match:
                    indent = len(match.group("indent").expandtabs(4))
                    if indent < base_indent:
                        break
                    level = min(8, max(0, (indent - base_indent) // 2))
                    marker = match.group("marker")
                    items.append(
                        {
                            "ordered": marker[0].isdigit(),
                            "level": level,
                            "text": match.group("text").strip(),
                        }
                    )
                    index += 1
                    continue
                if items and lines[index].strip() and len(lines[index]) - len(lines[index].lstrip()) > base_indent:
                    items[-1]["text"] += " " + lines[index].strip()
                    index += 1
                    continue
                break
            add("list", {"items": items}, start, index - 1)
            continue

        if index + 1 < len(lines) and "|" in line and _is_table_rule(lines[index + 1]):
            start = index
            headers = _split_table_row(line)
            rule = _split_table_row(lines[index + 1])
            if len(headers) != len(rule):
                raise PackageError(f"Markdown table header/rule width mismatch at {location}:{line_offset + index + 1}")
            alignments = [_table_alignment(item) for item in rule]
            index += 2
            rows: list[list[str]] = []
            while index < len(lines) and lines[index].strip() and "|" in lines[index]:
                row = _split_table_row(lines[index])
                if len(row) > len(headers):
                    raise PackageError(f"Markdown table row is wider than its header at {location}:{line_offset + index + 1}")
                rows.append(row + [""] * (len(headers) - len(row)))
                index += 1
            add("table", {"headers": headers, "rows": rows, "alignments": alignments}, start, index - 1)
            continue

        if stripped.startswith(">"):
            start = index
            quote_lines = []
            while index < len(lines) and lines[index].lstrip().startswith(">"):
                value = lines[index].lstrip()[1:]
                quote_lines.append(value[1:] if value.startswith(" ") else value)
                index += 1
            add("quote", _paragraph_text(quote_lines), start, index - 1)
            continue

        if stripped.startswith("<") and stripped.endswith(">"):
            message = f"Raw HTML is not supported at {location}:{line_offset + index + 1}"
            if strict:
                raise PackageError(message)
            diagnostics.warn("MARKDOWN_HTML_LITERAL", message, location=location)

        start = index
        paragraph_lines = [line]
        index += 1
        while index < len(lines) and lines[index].strip():
            candidate = lines[index]
            candidate_stripped = candidate.strip()
            if (
                HEADING.match(candidate)
                or LIST_ITEM.match(candidate)
                or candidate_stripped.startswith(("```", "~~~", ":::", ">"))
                or (index + 1 < len(lines) and "|" in candidate and _is_table_rule(lines[index + 1]))
            ):
                break
            paragraph_lines.append(candidate)
            index += 1
        add("paragraph", _paragraph_text(paragraph_lines), start, index - 1)

    _assign_ids(blocks, location)
    slots = _collect_slots(blocks)
    duplicates = sorted({name for name in slots if slots.count(name) > 1})
    if duplicates:
        raise PackageError(f"Markdown+ insert slots must occur once; duplicates: {', '.join(duplicates)}")
    return MarkdownDocument(blocks=blocks, slot_names=slots)


def parse_blocks(text: str, diagnostics: DiagnosticBag, location: str) -> list[Block]:
    """Compatibility entry point used by tests and simple callers."""

    return parse_document(text, diagnostics, location, strict=False).blocks
