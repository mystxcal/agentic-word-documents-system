from __future__ import annotations

import hashlib
import json
import os
import shutil
import uuid
from functools import lru_cache
from pathlib import Path
from typing import Any

from .errors import DocumentSystemError
from .model import ResolvedComponent, ResolvedDocument
from .rendering import poppler_status
from .resolver import file_hash
from .word.package import validate_docx


CACHE_SCHEMA = "agentic-component-adapter-cache/v1"


def _json_hash(value: object) -> str:
    encoded = json.dumps(
        value,
        sort_keys=True,
        separators=(",", ":"),
        ensure_ascii=False,
        default=str,
    ).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest().upper()


@lru_cache(maxsize=8)
def _adapter_code_signature(adapter: str) -> str:
    package = Path(__file__).resolve().parent
    common = [
        package / "component_cache.py",
        package / "model.py",
        package / "word" / "components.py",
        package / "word" / "fragments.py",
        package / "word" / "package.py",
    ]
    specific = {
        "markdown": [package / "sources" / "markdown.py", package / "sources" / "markdown_ast.py"],
        "figure": [],
        "pdf_pages": [package / "sources_pdf.py", package / "rendering.py"],
    }[adapter]
    records = [
        {"file": path.relative_to(package).as_posix(), "sha256": file_hash(path)}
        for path in [*common, *specific]
    ]
    return _json_hash(records)


def _program_record(path_value: str | None) -> dict | None:
    if not path_value:
        return None
    path = Path(path_value)
    if not path.is_file():
        return {"path": str(path), "available": False}
    stat = path.stat()
    return {
        "path": str(path.resolve()),
        "size": stat.st_size,
        "modified_ns": stat.st_mtime_ns,
    }


def adapter_payload(
    resolved: ResolvedDocument,
    component: ResolvedComponent,
    *,
    available_width_twips: int,
) -> dict | None:
    source = component.source_path
    if source is None:
        return None
    suffix = source.suffix.lower()
    component_type = component.declaration.type.value
    if component_type == "document" and suffix in {".md", ".markdown"}:
        adapter = "markdown"
    elif component_type in {"figure", "diagram"}:
        adapter = "figure"
    elif component_type == "pdf_pages" and suffix == ".pdf":
        adapter = "pdf_pages"
    else:
        return None

    kit_dump = resolved.kit.model_dump(mode="json", by_alias=True)
    payload: dict[str, Any] = {
        "schema": CACHE_SCHEMA,
        "adapter": adapter,
        "adapter_code_sha256": _adapter_code_signature(adapter),
        "component_id": component.id,
        "declaration": component.declaration.model_dump(mode="json", by_alias=True),
        "source_sha256": component.source_hash or file_hash(source),
        "related_source_hashes": dict(sorted(component.related_hashes.items())),
        "style_source_sha256": file_hash(resolved.presentation_paths["styles"]),
        "available_width_twips": available_width_twips,
    }
    if adapter == "markdown":
        payload["semantic_styles"] = kit_dump.get("semantic_styles", {})
        payload["table_styles"] = kit_dump.get("table_styles", {})
    if adapter == "pdf_pages":
        status = poppler_status()
        payload["renderer"] = {
            "pdftoppm": _program_record(status.get("pdftoppm")),
            "pdfinfo": _program_record(status.get("pdfinfo")),
        }
    payload["fingerprint"] = _json_hash(payload)
    return payload


class ComponentAdapterCache:
    """Disposable, content-addressed cache for generated DOCX adapter fragments."""

    def __init__(self, system_root: Path):
        self.root = Path(system_root).resolve() / ".cache" / "component-adapters" / "v1"

    def _entry(self, fingerprint: str) -> Path:
        return self.root / fingerprint[:2].lower() / fingerprint.lower()

    def lookup(self, payload: dict) -> tuple[Path | None, dict]:
        fingerprint = payload["fingerprint"]
        entry = self._entry(fingerprint)
        output = entry / "component.docx"
        metadata_path = entry / "metadata.json"
        event = {
            "component_id": payload["component_id"],
            "adapter": payload["adapter"],
            "fingerprint": fingerprint,
            "hit": False,
            "stored": False,
            "path": str(output),
        }
        if not output.is_file() or not metadata_path.is_file():
            event["reason"] = "not_found"
            return None, event
        try:
            metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
            valid = validate_docx(output)
            if (
                metadata.get("schema") != CACHE_SCHEMA
                or metadata.get("fingerprint") != fingerprint
                or metadata.get("output_sha256") != file_hash(output)
                or not valid.get("valid")
            ):
                raise ValueError("cache metadata, hash, or DOCX package validation did not match")
        except Exception as exc:
            quarantine = entry.with_name(f".{entry.name}.invalid-{uuid.uuid4().hex}")
            try:
                os.replace(entry, quarantine)
            except OSError:
                pass
            event["reason"] = "invalid"
            event["detail"] = str(exc)
            return None, event
        event.update({"hit": True, "reason": "verified"})
        return output, event

    def store(self, payload: dict, compiled: Path) -> tuple[Path, dict]:
        fingerprint = payload["fingerprint"]
        entry = self._entry(fingerprint)
        existing, event = self.lookup(payload)
        if existing is not None:
            event["reason"] = "verified_race_or_existing"
            return existing, event

        entry.parent.mkdir(parents=True, exist_ok=True)
        staging = entry.parent / f".{entry.name}.staging-{uuid.uuid4().hex}"
        staging.mkdir()
        try:
            candidate = staging / "component.docx"
            shutil.copy2(compiled, candidate)
            validation = validate_docx(candidate)
            if not validation.get("valid"):
                raise DocumentSystemError(
                    "Generated component cannot enter the cache because its DOCX package is invalid: "
                    + "; ".join(validation.get("issues") or [])
                )
            metadata = {
                **payload,
                "output_sha256": file_hash(candidate),
                "output_size_bytes": candidate.stat().st_size,
            }
            (staging / "metadata.json").write_text(
                json.dumps(metadata, indent=2, ensure_ascii=False) + "\n",
                encoding="utf-8",
            )
            try:
                os.replace(staging, entry)
            except OSError:
                winner, winner_event = self.lookup(payload)
                if winner is None:
                    raise
                return winner, winner_event
        finally:
            shutil.rmtree(staging, ignore_errors=True)
        return entry / "component.docx", {
            "component_id": payload["component_id"],
            "adapter": payload["adapter"],
            "fingerprint": fingerprint,
            "hit": False,
            "stored": True,
            "reason": "compiled_and_stored",
            "path": str(entry / "component.docx"),
        }


def summarize_cache_events(events: list[dict]) -> dict:
    return {
        "schema": "agentic-component-cache-summary/v1",
        "enabled": True,
        "events": events,
        "hits": sum(bool(item.get("hit")) for item in events),
        "misses": sum(not bool(item.get("hit")) for item in events),
        "stored": sum(bool(item.get("stored")) for item in events),
    }
