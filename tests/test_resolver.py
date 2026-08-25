import tempfile
import unittest
import json
from pathlib import Path

from agentic_docs.diagnostics import DiagnosticBag
from agentic_docs.errors import ResolutionError
from agentic_docs.resolver import _safe_relative, resolve_document


class ResolverTests(unittest.TestCase):
    def test_safe_relative_accepts_path_inside_root(self):
        with tempfile.TemporaryDirectory() as value:
            root = Path(value)
            result = _safe_relative(root, "content/part.docx", allowed_root=root, label="component")
            self.assertEqual(result, (root / "content" / "part.docx").resolve())

    def test_safe_relative_rejects_escape(self):
        with tempfile.TemporaryDirectory() as value:
            root = Path(value)
            with self.assertRaisesRegex(ResolutionError, "escapes"):
                _safe_relative(root, "../outside.docx", allowed_root=root, label="component")

    def test_project_file_component_source_is_confined_to_selected_project(self):
        with tempfile.TemporaryDirectory() as value:
            root = Path(value)
            kit = root / "kits" / "studio"
            profile = root / "profiles" / "guide"
            project = root / "projects" / "demo"
            document = project / "documents" / "doc"
            for path in (kit, profile, document, project / "sources"):
                path.mkdir(parents=True, exist_ok=True)
            (kit / "styles.docx").write_bytes(b"styles")
            (profile / "shell.docx").write_bytes(b"shell")
            (project / "sources" / "guide.pdf").write_bytes(b"pdf")
            (kit / "kit.jsonc").write_text(
                json.dumps(
                    {
                        "schema": "agentic-kit/v2",
                        "id": "studio",
                        "components": {"styles": "styles.docx"},
                    }
                ),
                encoding="utf-8",
            )
            (profile / "profile.jsonc").write_text(
                json.dumps(
                    {
                        "schema": "agentic-profile/v2",
                        "id": "guide",
                        "shell": "shell.docx",
                        "body_slot": "AGDOC.BODY",
                    }
                ),
                encoding="utf-8",
            )
            (project / "project.jsonc").write_text(
                json.dumps(
                    {
                        "schema": "agentic-project/v2",
                        "id": "demo",
                        "name": "Demo",
                        "description": "Neutral resolver fixture",
                        "metadata": {"audience": "test"},
                    }
                ),
                encoding="utf-8",
            )
            manifest = document / "document.jsonc"
            manifest.write_text(
                json.dumps(
                    {
                        "schema": "agentic-document/v2",
                        "id": "demo-doc",
                        "project": "demo",
                        "profile": "guide",
                        "kit": "studio",
                        "metadata": {
                            "type": "Guide",
                            "title": "Demo",
                            "revision": "R00",
                            "date": "2026-08-24",
                        },
                        "presentation": {
                            "styles": "kit:styles",
                            "page_regions": {"main": {"header": None, "footer": None}},
                        },
                        "sequence": [{"region": "main", "items": ["pages"]}],
                        "components": {
                            "pages": {
                                "type": "pdf_pages",
                                "ownership": "source",
                                "source": "project-file:sources/guide.pdf",
                                "options": {"pages": [1]},
                            }
                        },
                        "outputs": {"basename": "Demo"},
                    }
                ),
                encoding="utf-8",
            )
            resolved = resolve_document(manifest, DiagnosticBag())
            self.assertEqual(
                resolved.components["pages"].source_path,
                (project / "sources" / "guide.pdf").resolve(),
            )


if __name__ == "__main__":
    unittest.main()
