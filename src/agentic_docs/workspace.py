from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from .errors import ResolutionError
from .jsonc import load_jsonc


@dataclass(frozen=True)
class DocumentRecord:
    id: str | None
    manifest_path: Path
    project: str | None
    title: str | None
    error: str | None = None


def _require_system_root(value: Path, *, label: str) -> Path:
    root = Path(value).expanduser().resolve()
    missing = [name for name in ("kits", "profiles", "projects") if not (root / name).is_dir()]
    if missing:
        raise ResolutionError(
            f"{label} does not identify a valid document system: {root}. "
            f"Missing: {', '.join(missing)}"
        )
    return root


def discover_system_root(start: Path | None = None, *, explicit: Path | None = None) -> Path:
    """Find the nearest Agentic Word Documents root from an operator's current location."""

    if explicit is not None:
        return _require_system_root(explicit, label="Selected --root")
    configured = os.environ.get("AGENTIC_DOCS_ROOT")
    if configured:
        return _require_system_root(Path(configured), label="AGENTIC_DOCS_ROOT")
    origin = Path(start or Path.cwd()).resolve()
    candidates = [origin, *origin.parents]
    for candidate in candidates:
        if all((candidate / name).is_dir() for name in ("kits", "profiles", "projects")):
            return candidate
        nested = candidate / "agentic-word-documents-system"
        if all((nested / name).is_dir() for name in ("kits", "profiles", "projects")):
            return nested.resolve()
    raise ResolutionError(
        f"Could not locate an Agentic Word Documents root from {origin}. "
        "Run the command inside the system folder or pass a manifest path."
    )


def discover_documents(system_root: Path) -> list[DocumentRecord]:
    root = Path(system_root).resolve()
    records: list[DocumentRecord] = []
    for manifest_path in sorted((root / "projects").glob("*/documents/*/document.jsonc")):
        try:
            payload = load_jsonc(manifest_path)
        except Exception as exc:
            records.append(
                DocumentRecord(
                    id=None,
                    manifest_path=manifest_path.resolve(),
                    project=None,
                    title=None,
                    error=f"Could not read manifest: {exc}",
                )
            )
            continue
        identifier = payload.get("id") if isinstance(payload, dict) else None
        if not isinstance(identifier, str) or not identifier:
            records.append(
                DocumentRecord(
                    id=None,
                    manifest_path=manifest_path.resolve(),
                    project=payload.get("project") if isinstance(payload, dict) else None,
                    title=None,
                    error="Manifest has no usable document id",
                )
            )
            continue
        metadata = payload.get("metadata", {}) if isinstance(payload, dict) else {}
        records.append(
            DocumentRecord(
                id=identifier,
                manifest_path=manifest_path.resolve(),
                project=payload.get("project") if isinstance(payload, dict) else None,
                title=metadata.get("title") if isinstance(metadata, dict) else None,
                error=None,
            )
        )
    return records


def resolve_document_spec(
    spec: str | Path,
    *,
    start: Path | None = None,
    system_root: Path | None = None,
) -> Path:
    """Resolve either a manifest path or a unique document id."""

    raw = str(spec)
    candidate = Path(raw).expanduser()
    if candidate.is_file():
        return candidate.resolve()
    if not candidate.is_absolute() and start is not None:
        relative = (Path(start).resolve() / candidate)
        if relative.is_file():
            return relative.resolve()

    looks_like_path = any(char in raw for char in ("/", "\\")) or raw.lower().endswith(".jsonc")
    if looks_like_path:
        raise ResolutionError(f"Document manifest does not exist: {candidate}")

    root = _require_system_root(system_root, label="Selected --root") if system_root else discover_system_root(start)
    matches = [record for record in discover_documents(root) if record.id == raw]
    if not matches:
        available = ", ".join(record.id for record in discover_documents(root) if record.id) or "none"
        raise ResolutionError(f"Unknown document id {raw!r}. Available documents: {available}")
    if len(matches) > 1:
        locations = "; ".join(str(record.manifest_path) for record in matches)
        raise ResolutionError(f"Document id {raw!r} is ambiguous: {locations}")
    return matches[0].manifest_path
