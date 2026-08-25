from __future__ import annotations

import json
import os
from pathlib import Path

from .model import DocumentManifest, KitManifest, ProfileManifest, ProjectManifest


SCHEMAS = {
    "agentic-document-v2.schema.json": DocumentManifest,
    "agentic-kit-v2.schema.json": KitManifest,
    "agentic-profile-v2.schema.json": ProfileManifest,
    "agentic-project-v2.schema.json": ProjectManifest,
}


def _schema_bytes(model, filename: str) -> bytes:
    value = model.model_json_schema(by_alias=True, mode="validation")
    value["$id"] = filename
    value["title"] = {
        "agentic-document-v2.schema.json": "Agentic document manifest V2",
        "agentic-kit-v2.schema.json": "Agentic style-kit manifest V2",
        "agentic-profile-v2.schema.json": "Agentic document-profile manifest V2",
        "agentic-project-v2.schema.json": "Agentic project manifest V2",
    }[filename]
    return (json.dumps(value, indent=2, ensure_ascii=False) + "\n").encode("utf-8")


def refresh_schemas(system_root: Path, *, check_only: bool = False) -> dict:
    target_root = Path(system_root).resolve() / "schemas"
    results = []
    for filename, model in SCHEMAS.items():
        target = target_root / filename
        expected = _schema_bytes(model, filename)
        current = target.read_bytes() if target.is_file() else None
        same = current == expected
        if not check_only and not same:
            target_root.mkdir(parents=True, exist_ok=True)
            temporary = target.with_suffix(target.suffix + ".tmp")
            temporary.write_bytes(expected)
            os.replace(temporary, target)
        results.append(
            {
                "name": filename,
                "path": str(target),
                "current": same,
                "updated": not check_only and not same,
            }
        )
    return {
        "schema": "agentic-schema-refresh/v1",
        "root": str(target_root),
        "check_only": check_only,
        "current": all(item["current"] for item in results),
        "files": results,
    }
