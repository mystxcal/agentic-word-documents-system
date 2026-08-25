from __future__ import annotations

import re
from datetime import date, datetime

from lxml import etree

from ..diagnostics import DiagnosticBag
from ..errors import ResolutionError
from ..model import FieldBinding, ResolvedDocument
from .ooxml import W_NS, find_sdts, qn
from .package import DocxPackage


VALUE_TOKEN = re.compile(r"\{([A-Za-z0-9_.-]+)\}")


def _value_path(resolved: ResolvedDocument, path: str):
    roots = {
        "metadata": resolved.manifest.metadata,
        "project": resolved.project,
        "document": resolved.manifest,
    }
    parts = path.split(".")
    if not parts or parts[0] not in roots:
        raise ResolutionError(f"Binding path must begin with metadata, project, or document: {path!r}")
    value = roots[parts[0]]
    for part in parts[1:]:
        if hasattr(value, part):
            value = getattr(value, part)
        elif isinstance(value, dict) and part in value:
            value = value[part]
        else:
            raise ResolutionError(f"Binding path does not exist: {path!r}")
    return value


def _display(value, binding: FieldBinding) -> str:
    if value is None:
        result = ""
    elif isinstance(value, (date, datetime)):
        result = value.strftime(binding.date_format or "%Y-%m-%d")
    else:
        result = str(value)
    if binding.transform == "upper":
        return result.upper()
    if binding.transform == "lower":
        return result.lower()
    if binding.transform == "title":
        return result.title()
    return result


def _resolve(binding_value: str | FieldBinding, resolved: ResolvedDocument) -> tuple[str, bool]:
    binding = FieldBinding(path=binding_value) if isinstance(binding_value, str) else binding_value
    if binding.path:
        return _display(_value_path(resolved, binding.path), binding), binding.required

    def replace(match: re.Match) -> str:
        return _display(_value_path(resolved, match.group(1)), binding)

    return VALUE_TOKEN.sub(replace, binding.template or ""), binding.required


def _set_text(sdt, value: str) -> None:
    texts = sdt.xpath(".//w:t", namespaces={"w": W_NS})
    if texts:
        texts[0].text = value
        if value.startswith(" ") or value.endswith(" "):
            texts[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        for extra in texts[1:]:
            extra.text = ""
        return
    content = sdt.find(qn("w:sdtContent"))
    if content is None:
        raise ResolutionError("Bound content control has no sdtContent container")
    paragraph = etree.SubElement(content, qn("w:p"))
    run = etree.SubElement(paragraph, qn("w:r"))
    text = etree.SubElement(run, qn("w:t"))
    text.text = value


def apply_field_bindings(package: DocxPackage, resolved: ResolvedDocument, diagnostics: DiagnosticBag) -> list[dict]:
    results = []
    candidate_parts = [
        name
        for name in package.parts
        if name == "word/document.xml" or re.fullmatch(r"word/(?:header|footer)[^/]*\.xml", name)
    ]
    for tag, specification in resolved.profile.field_bindings.items():
        value, required = _resolve(specification, resolved)
        matches = []
        roots = {}
        for part in candidate_parts:
            try:
                root = package.xml(part)
            except Exception:
                continue
            found = find_sdts(root, tag)
            if found:
                roots[part] = root
                matches.extend((part, item) for item in found)
        if not matches:
            message = f"No content control tagged {tag!r} exists in the selected shell or page furniture"
            if required:
                raise ResolutionError(message)
            diagnostics.warn("OPTIONAL_BINDING_TARGET_MISSING", message, location=str(resolved.profile_path))
            continue
        for _, item in matches:
            _set_text(item, value)
        for part, root in roots.items():
            package.set_xml(part, root)
        results.append({"tag": tag, "value": value, "occurrences": len(matches)})
    return results
