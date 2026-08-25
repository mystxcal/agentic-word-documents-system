from __future__ import annotations

import json
from pathlib import Path

from .diagnostics import DiagnosticBag
from .model import ResolvedDocument
from .resolver import file_hash


def _file_record(path) -> dict | None:
    if path is None:
        return None
    return {"path": str(path), "sha256": file_hash(path)}


def resolved_summary(resolved: ResolvedDocument, diagnostics: DiagnosticBag) -> dict:
    sequence = [item for group in resolved.manifest.sequence for item in group.items]

    def tree(component_id: str) -> dict:
        declaration = resolved.components[component_id].declaration
        return {
            "id": component_id,
            "slots": {
                tag: [tree(child) for child in children]
                for tag, children in declaration.slots.items()
            },
        }

    return {
        "schema": "agentic-resolved-summary/v2",
        "document": {
            "id": resolved.manifest.id,
            "title": resolved.manifest.metadata.title,
            "revision": resolved.manifest.metadata.revision,
        },
        "layers": {
            "kit": {"id": resolved.kit.id, **_file_record(resolved.kit_path)},
            "profile": {"id": resolved.profile.id, **_file_record(resolved.profile_path)},
            "project": {"id": resolved.project.id, **_file_record(resolved.project_path)},
            "document": _file_record(resolved.manifest_path),
            "shell": _file_record(resolved.shell_path),
        },
        "sequence": sequence,
        "assembly": [tree(component_id) for component_id in sequence],
        "components": [
            {
                "id": component_id,
                "type": item.declaration.type,
                "ownership": item.declaration.ownership,
                "source": str(item.source_path) if item.source_path else None,
                "source_hash": item.source_hash,
                "related_sources": {
                    name: {"path": str(path), "sha256": item.related_hashes.get(name)}
                    for name, path in sorted(item.related_paths.items())
                },
            }
            for component_id, item in resolved.components.items()
        ],
        "presentation": {
            key: _file_record(value)
            for key, value in sorted(resolved.presentation_paths.items())
        },
        "release": {
            "required_gates": resolved.profile.release_gates,
            "states": {
                key: value.value
                for key, value in resolved.manifest.release.gates.items()
            },
            "ready": all(
                resolved.manifest.release.gates.get(gate) is not None
                and resolved.manifest.release.gates[gate].value in {"met", "not_applicable"}
                for gate in resolved.profile.release_gates
            ),
        },
        "diagnostics": diagnostics.to_list(),
    }


def write_json(path: Path, value: object) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(value, indent=2, default=str) + "\n", encoding="utf-8")
    temporary.replace(path)
