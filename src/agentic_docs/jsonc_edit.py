from __future__ import annotations

import json
from dataclasses import dataclass, field
from typing import Any

from .errors import ManifestError
from .jsonc import loads_jsonc


@dataclass(frozen=True)
class Token:
    kind: str
    value: Any
    start: int
    end: int


@dataclass
class Node:
    kind: str
    start: int
    end: int
    value: Any = None
    properties: dict[str, "Node"] = field(default_factory=dict)
    items: list["Node"] = field(default_factory=list)


def _tokens(text: str) -> list[Token]:
    result: list[Token] = []
    index = 0
    while index < len(text):
        char = text[index]
        if char.isspace():
            index += 1
            continue
        if char == "/" and index + 1 < len(text) and text[index + 1] == "/":
            index += 2
            while index < len(text) and text[index] not in "\r\n":
                index += 1
            continue
        if char == "/" and index + 1 < len(text) and text[index + 1] == "*":
            end = text.find("*/", index + 2)
            if end < 0:
                raise ManifestError("Unterminated block comment in JSONC manifest")
            index = end + 2
            continue
        if char in "{}[]:,":
            result.append(Token(char, char, index, index + 1))
            index += 1
            continue
        if char == '"':
            start = index
            index += 1
            escaped = False
            while index < len(text):
                current = text[index]
                if escaped:
                    escaped = False
                elif current == "\\":
                    escaped = True
                elif current == '"':
                    index += 1
                    break
                index += 1
            else:
                raise ManifestError("Unterminated string in JSONC manifest")
            raw = text[start:index]
            result.append(Token("string", json.loads(raw), start, index))
            continue
        start = index
        while index < len(text) and not text[index].isspace() and text[index] not in "{}[]:,/":
            index += 1
        raw = text[start:index]
        try:
            value = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise ManifestError(f"Invalid JSON value near offset {start}: {raw!r}") from exc
        result.append(Token("scalar", value, start, index))
    return result


class _Parser:
    def __init__(self, tokens: list[Token]):
        self.tokens = tokens
        self.index = 0

    def _peek(self) -> Token:
        if self.index >= len(self.tokens):
            raise ManifestError("Unexpected end of JSONC manifest")
        return self.tokens[self.index]

    def _take(self, kind: str | None = None) -> Token:
        token = self._peek()
        if kind is not None and token.kind != kind:
            raise ManifestError(f"Expected {kind!r} at offset {token.start}; found {token.kind!r}")
        self.index += 1
        return token

    def value(self) -> Node:
        token = self._peek()
        if token.kind == "{":
            return self.object()
        if token.kind == "[":
            return self.array()
        self._take()
        return Node("value", token.start, token.end, value=token.value)

    def object(self) -> Node:
        opening = self._take("{")
        properties: dict[str, Node] = {}
        if self._peek().kind == "}":
            closing = self._take("}")
            return Node("object", opening.start, closing.end, properties=properties)
        while True:
            key = self._take("string")
            self._take(":")
            value = self.value()
            if key.value in properties:
                raise ManifestError(f"Duplicate JSON object key {key.value!r}")
            properties[key.value] = value
            if self._peek().kind == ",":
                self._take(",")
                if self._peek().kind == "}":
                    closing = self._take("}")
                    break
                continue
            closing = self._take("}")
            break
        return Node("object", opening.start, closing.end, properties=properties)

    def array(self) -> Node:
        opening = self._take("[")
        items: list[Node] = []
        if self._peek().kind == "]":
            closing = self._take("]")
            return Node("array", opening.start, closing.end, items=items)
        while True:
            items.append(self.value())
            if self._peek().kind == ",":
                self._take(",")
                if self._peek().kind == "]":
                    closing = self._take("]")
                    break
                continue
            closing = self._take("]")
            break
        return Node("array", opening.start, closing.end, items=items)


def parse_tree(text: str) -> Node:
    tokens = _tokens(text)
    if not tokens:
        raise ManifestError("JSONC manifest is empty")
    parser = _Parser(tokens)
    result = parser.value()
    if parser.index != len(tokens):
        token = tokens[parser.index]
        raise ManifestError(f"Unexpected token at offset {token.start}")
    return result


def _property(node: Node, key: str) -> Node:
    if node.kind != "object" or key not in node.properties:
        raise ManifestError(f"JSONC object has no property {key!r}")
    return node.properties[key]


def _line_indent(text: str, offset: int) -> str:
    start = text.rfind("\n", 0, offset) + 1
    prefix = text[start:offset]
    return prefix[: len(prefix) - len(prefix.lstrip(" \t"))]


def _line_start(text: str, offset: int) -> int:
    return text.rfind("\n", 0, offset) + 1


def _json_property(key: str, value: Any, indent: str) -> str:
    rendered = json.dumps({key: value}, indent=2, ensure_ascii=False).splitlines()[1:-1]
    return "\n".join(indent + (line[2:] if line.startswith("  ") else line) for line in rendered)


def _add_object_property(text: str, node: Node, key: str, value: Any) -> list[tuple[int, int, str]]:
    if node.kind != "object":
        raise ManifestError(f"Cannot add property {key!r} to a non-object JSON value")
    if key in node.properties:
        raise ManifestError(f"JSON object already contains {key!r}")
    base_indent = _line_indent(text, node.start)
    child_indent = base_indent + "  "
    if not node.properties:
        replacement = "{\n" + _json_property(key, value, child_indent) + "\n" + base_indent + "}"
        return [(node.start, node.end, replacement)]
    closing = node.end - 1
    closing_line_start = _line_start(text, closing)
    last = list(node.properties.values())[-1]
    if closing_line_start <= last.start:
        compact = json.dumps(key, ensure_ascii=False) + ": " + json.dumps(value, ensure_ascii=False)
        return [(closing, closing, ", " + compact)]
    between = text[last.end:closing]
    has_trailing_comma = "," in between
    edits = [(closing_line_start, closing_line_start, _json_property(key, value, child_indent) + "\n")]
    if not has_trailing_comma:
        edits.append((last.end, last.end, ","))
    return edits


def _append_array(text: str, node: Node, value: Any) -> list[tuple[int, int, str]]:
    if node.kind != "array":
        raise ManifestError("Cannot append an item to a non-array JSON value")
    base_indent = _line_indent(text, node.start)
    child_indent = base_indent + "  "
    rendered = json.dumps(value, indent=2, ensure_ascii=False)
    rendered = "\n".join(child_indent + line for line in rendered.splitlines())
    if not node.items:
        return [(node.start, node.end, "[\n" + rendered + "\n" + base_indent + "]")]
    closing = node.end - 1
    closing_line_start = _line_start(text, closing)
    last = node.items[-1]
    if closing_line_start <= last.start:
        return [(closing, closing, ", " + json.dumps(value, ensure_ascii=False))]
    between = text[last.end:closing]
    has_trailing_comma = "," in between
    edits = [(closing_line_start, closing_line_start, rendered + "\n")]
    if not has_trailing_comma:
        edits.append((last.end, last.end, ","))
    return edits


def _insert_array_after(text: str, node: Node, after_value: Any, value: Any) -> list[tuple[int, int, str]]:
    if node.kind != "array":
        raise ManifestError("Cannot insert an item into a non-array JSON value")
    matching = [index for index, item in enumerate(node.items) if item.kind == "value" and item.value == after_value]
    if len(matching) != 1:
        raise ManifestError(f"Expected exactly one sequence item {after_value!r}; found {len(matching)}")
    index = matching[0]
    if index == len(node.items) - 1:
        return _append_array(text, node, value)
    next_item = node.items[index + 1]
    next_line_start = _line_start(text, next_item.start)
    if next_line_start < next_item.start:
        indent = _line_indent(text, next_item.start)
        return [(next_line_start, next_line_start, indent + json.dumps(value, ensure_ascii=False) + ",\n")]
    return [(next_item.start, next_item.start, json.dumps(value, ensure_ascii=False) + ", ")]


def _scalar(node: Node) -> Any:
    if node.kind != "value":
        raise ManifestError("Expected a scalar JSON value")
    return node.value


def add_component(
    text: str,
    component_id: str,
    declaration: dict,
    *,
    parent_component: str | None = None,
    slot_name: str | None = None,
    region: str | None = None,
    after_component: str | None = None,
) -> str:
    """Add one declared component and exactly one placement while preserving JSONC comments."""

    if bool(parent_component) == bool(region):
        raise ManifestError("Component placement requires exactly one parent component or top-level region")
    root = parse_tree(text)
    components = _property(root, "components")
    if component_id in components.properties:
        raise ManifestError(f"Component {component_id!r} already exists")
    edits = _add_object_property(text, components, component_id, declaration)

    if parent_component:
        if parent_component not in components.properties:
            raise ManifestError(f"Parent component {parent_component!r} does not exist")
        parent = components.properties[parent_component]
        if parent.kind != "object":
            raise ManifestError(f"Parent component {parent_component!r} is not an object")
        selected_slot = slot_name or component_id
        slots = parent.properties.get("slots")
        if slots is None:
            edits.extend(_add_object_property(text, parent, "slots", {selected_slot: [component_id]}))
        else:
            edits.extend(_add_object_property(text, slots, selected_slot, [component_id]))
    else:
        sequence = _property(root, "sequence")
        if sequence.kind != "array":
            raise ManifestError("Document sequence is not an array")
        match = None
        for group in sequence.items:
            if group.kind != "object":
                continue
            region_node = group.properties.get("region")
            if region_node is not None and _scalar(region_node) == region:
                match = group
                break
        if match is None:
            raise ManifestError(f"Document sequence has no region {region!r}")
        items = _property(match, "items")
        edits.extend(
            _insert_array_after(text, items, after_component, component_id)
            if after_component
            else _append_array(text, items, component_id)
        )

    candidate = text
    for start, end, replacement in sorted(edits, key=lambda item: (item[0], item[1]), reverse=True):
        candidate = candidate[:start] + replacement + candidate[end:]
    loads_jsonc(candidate)
    return candidate


def replace_values(text: str, values: dict[tuple[str, ...], Any]) -> str:
    """Replace exact JSONC values without reserializing the surrounding manifest.

    This is intentionally narrow: authoring operations may update known scalar
    controls while comments, ordering, whitespace, and unrelated extensions stay
    byte-for-byte untouched.
    """

    root = parse_tree(text)
    edits: list[tuple[int, int, str]] = []
    for path, value in values.items():
        if not path:
            raise ManifestError("A JSONC replacement path cannot be empty")
        node = root
        for key in path:
            node = _property(node, key)
        edits.append((node.start, node.end, json.dumps(value, ensure_ascii=False)))
    candidate = text
    for start, end, replacement in sorted(edits, key=lambda item: (item[0], item[1]), reverse=True):
        candidate = candidate[:start] + replacement + candidate[end:]
    loads_jsonc(candidate)
    return candidate
