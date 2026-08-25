from __future__ import annotations

import json
from pathlib import Path

from .errors import ManifestError


def _strip_comments(text: str) -> str:
    """Remove JSONC comments without changing content inside string literals."""
    result: list[str] = []
    index = 0
    in_string = False
    escaped = False
    while index < len(text):
        char = text[index]
        next_char = text[index + 1] if index + 1 < len(text) else ""
        if in_string:
            result.append(char)
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == '"':
                in_string = False
            index += 1
            continue
        if char == '"':
            in_string = True
            result.append(char)
            index += 1
            continue
        if char == "/" and next_char == "/":
            result.extend("  ")
            index += 2
            while index < len(text) and text[index] not in "\r\n":
                result.append(" ")
                index += 1
            continue
        if char == "/" and next_char == "*":
            result.extend("  ")
            index += 2
            closed = False
            while index < len(text):
                if index + 1 < len(text) and text[index] == "*" and text[index + 1] == "/":
                    result.extend("  ")
                    index += 2
                    closed = True
                    break
                result.append(text[index] if text[index] in "\r\n" else " ")
                index += 1
            if not closed:
                raise ManifestError("Unterminated block comment in JSONC manifest")
            continue
        result.append(char)
        index += 1
    return "".join(result)


def _strip_trailing_commas(text: str) -> str:
    result: list[str] = []
    index = 0
    in_string = False
    escaped = False
    while index < len(text):
        char = text[index]
        if in_string:
            result.append(char)
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == '"':
                in_string = False
            index += 1
            continue
        if char == '"':
            in_string = True
            result.append(char)
            index += 1
            continue
        if char == ",":
            lookahead = index + 1
            while lookahead < len(text) and text[lookahead].isspace():
                lookahead += 1
            if lookahead < len(text) and text[lookahead] in "}]":
                result.append(" ")
                index += 1
                continue
        result.append(char)
        index += 1
    return "".join(result)


def loads_jsonc(text: str, *, source: str = "<memory>") -> object:
    try:
        return json.loads(_strip_trailing_commas(_strip_comments(text)))
    except json.JSONDecodeError as exc:
        raise ManifestError(
            f"Invalid JSONC in {source} at line {exc.lineno}, column {exc.colno}: {exc.msg}"
        ) from exc


def load_jsonc(path: Path) -> object:
    path = Path(path).resolve()
    try:
        text = path.read_text(encoding="utf-8-sig")
    except OSError as exc:
        raise ManifestError(f"Could not read manifest {path}: {exc}") from exc
    return loads_jsonc(text, source=str(path))
