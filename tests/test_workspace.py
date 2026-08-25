import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from agentic_docs.workspace import discover_documents, discover_system_root, resolve_document_spec


class WorkspaceTests(unittest.TestCase):
    def test_discovers_document_by_id_from_nested_folder(self):
        with TemporaryDirectory() as directory:
            root = Path(directory) / "system"
            for name in ("kits", "profiles", "projects"):
                (root / name).mkdir(parents=True)
            manifest = root / "projects" / "p" / "documents" / "d" / "document.jsonc"
            manifest.parent.mkdir(parents=True)
            manifest.write_text(
                json.dumps(
                    {
                        "id": "example-document",
                        "project": "p",
                        "metadata": {"title": "Example"},
                    }
                ),
                encoding="utf-8",
            )
            nested = root / "projects" / "p" / "documents"
            self.assertEqual(discover_system_root(nested), root.resolve())
            self.assertEqual(discover_documents(root)[0].id, "example-document")
            self.assertEqual(resolve_document_spec("example-document", start=nested), manifest.resolve())

    def test_malformed_manifest_is_reported_without_hiding_valid_documents(self):
        with TemporaryDirectory() as directory:
            root = Path(directory) / "system"
            for name in ("kits", "profiles", "projects"):
                (root / name).mkdir(parents=True)
            valid = root / "projects" / "p" / "documents" / "valid" / "document.jsonc"
            invalid = root / "projects" / "p" / "documents" / "invalid" / "document.jsonc"
            valid.parent.mkdir(parents=True)
            invalid.parent.mkdir(parents=True)
            valid.write_text(json.dumps({"id": "valid-doc", "metadata": {}}), encoding="utf-8")
            invalid.write_text("{ invalid json", encoding="utf-8")

            records = discover_documents(root)
            self.assertEqual(len(records), 2)
            self.assertEqual(len([record for record in records if record.error]), 1)
            self.assertEqual(resolve_document_spec("valid-doc", start=root), valid.resolve())

    def test_explicit_root_overrides_a_stale_environment_root(self):
        with TemporaryDirectory() as directory:
            root = Path(directory) / "system"
            for name in ("kits", "profiles", "projects"):
                (root / name).mkdir(parents=True)
            with patch.dict("os.environ", {"AGENTIC_DOCS_ROOT": str(root / "missing")}, clear=False):
                self.assertEqual(discover_system_root(explicit=root), root.resolve())


if __name__ == "__main__":
    unittest.main()
