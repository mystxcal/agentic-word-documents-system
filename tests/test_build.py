from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from types import SimpleNamespace
from unittest.mock import patch

from docx import Document

from agentic_docs.build import (
    _effective_furniture_variants,
    _finalize_intermediates,
    _promote_run_directory,
    _refresh_current,
)
from agentic_docs.diagnostics import DiagnosticBag
from agentic_docs.model import DocumentManifest, ResolvedComponent
from agentic_docs.rendering import find_word_field_errors
from agentic_docs.word.components import compile_component_wrapper
from agentic_docs.word.package import DocxPackage


class BuildIntermediateTests(unittest.TestCase):
    def test_pdf_pages_placeholder_does_not_open_or_render_source_pdf(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            shell = root / "shell.docx"
            Document().save(shell)
            source = root / "manual.pdf"
            source.write_bytes(b"not opened by placeholder mode")
            manifest = DocumentManifest.model_validate(
                {
                    "schema": "agentic-document/v2",
                    "id": "preview-test",
                    "project": "example",
                    "profile": "plain",
                    "kit": "plain",
                    "metadata": {"title": "Preview", "date": "2026-08-25"},
                    "presentation": {
                        "styles": "kit:plain",
                        "page_regions": {"main": {"header": None, "footer": None}},
                    },
                    "sequence": [{"region": "main", "items": ["manual"]}],
                    "components": {
                        "manual": {
                            "type": "pdf_pages",
                            "ownership": "source",
                            "source": "manual.pdf",
                            "options": {"pages": [1, [3, 5]]},
                        }
                    },
                    "outputs": {"basename": "Preview"},
                }
            )
            component = ResolvedComponent(
                id="manual",
                declaration=manifest.components["manual"],
                source_path=source,
                source_hash="unused",
            )
            with patch(
                "agentic_docs.word.components.compile_pdf_pages",
                side_effect=AssertionError("heavy adapter must not run"),
            ):
                wrapper = compile_component_wrapper(
                    DocxPackage(shell),
                    SimpleNamespace(),
                    component,
                    build_work=root / "work",
                    diagnostics=DiagnosticBag(),
                    preview_action="placeholder",
                )
            self.assertIn("Preview placeholder", "".join(wrapper.itertext()))
    def test_same_donor_path_still_selects_its_even_reference(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            donor = root / "header.docx"
            donor.write_bytes(b"same donor")
            resolved = SimpleNamespace(
                presentation_paths={
                    "region.main.header.default": donor,
                    "region.main.header.even": donor,
                }
            )
            variants, first_used, even_used = _effective_furniture_variants(resolved, "main")
            self.assertFalse(first_used)
            self.assertTrue(even_used)
            self.assertEqual({"default", "even"}, set(variants["header"]))

    def test_real_even_variant_inherits_other_kind_default(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            header_default = root / "header-default.docx"
            header_even = root / "header-even.docx"
            footer_default = root / "footer-default.docx"
            header_default.write_bytes(b"header default")
            header_even.write_bytes(b"header even")
            footer_default.write_bytes(b"footer default")
            resolved = SimpleNamespace(
                presentation_paths={
                    "region.main.header.default": header_default,
                    "region.main.header.even": header_even,
                    "region.main.footer.default": footer_default,
                }
            )
            variants, _first_used, even_used = _effective_furniture_variants(resolved, "main")
            self.assertTrue(even_used)
            self.assertEqual("explicit", variants["header"]["even"][1])
            self.assertEqual("inherited-default", variants["footer"]["even"][1])
            self.assertEqual("default", variants["footer"]["even"][2])
    def test_known_word_field_errors_are_reported_by_page(self):
        result = find_word_field_errors(
            [
                "A clean cover",
                "Topic ........ Error! Bookmark not defined.",
                "Error! Reference source not found. and Error! Reference source not found.",
            ]
        )

        self.assertEqual(
            [
                {"page": 2, "message": "Error! Bookmark not defined.", "count": 1},
                {"page": 3, "message": "Error! Reference source not found.", "count": 2},
            ],
            result,
        )

    def test_run_promotion_falls_back_to_verified_copy_when_directory_rename_is_denied(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / ".run.tmp"
            destination = root / "run"
            source.mkdir()
            (source / "artifact.txt").write_text("verified artifact", encoding="utf-8")
            with patch("pathlib.Path.replace", side_effect=PermissionError("protected folder")):
                with patch("agentic_docs.build.time.sleep"):
                    result = _promote_run_directory(source, destination)

            self.assertEqual("verified-copy", result["method"])
            self.assertTrue(result["verified"])
            self.assertEqual("verified artifact", (destination / "artifact.txt").read_text(encoding="utf-8"))

    def test_compact_default_removes_raw_and_work_but_reports_reclaimed_size(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            raw = root / "document.raw.docx"
            work = root / "work"
            work.mkdir()
            raw.write_bytes(b"raw-docx")
            (work / "adapter.png").write_bytes(b"adapter")

            result = _finalize_intermediates(
                raw,
                work,
                retain=False,
                raw_sha256="abc",
            )

            self.assertFalse(result["retained"])
            self.assertEqual(result["removed_size_bytes"], len(b"raw-docx") + len(b"adapter"))
            self.assertFalse(raw.exists())
            self.assertFalse(work.exists())

    def test_diagnostic_mode_retains_raw_and_work(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            raw = root / "document.raw.docx"
            work = root / "work"
            work.mkdir()
            raw.write_bytes(b"raw-docx")
            (work / "adapter.png").write_bytes(b"adapter")

            result = _finalize_intermediates(
                raw,
                work,
                retain=True,
                raw_sha256="abc",
            )

            self.assertTrue(result["retained"])
            self.assertEqual(result["removed_size_bytes"], 0)
            self.assertTrue(raw.is_file())
            self.assertTrue(work.is_dir())

    def test_failed_current_refresh_does_not_invalidate_completed_build(self):
        resolved = SimpleNamespace(
            system_root=Path("C:/system"),
            manifest=SimpleNamespace(
                id="document-1",
                outputs=SimpleNamespace(basename="Example"),
            ),
        )
        report = {"build_id": "BUILD-1", "artifacts": {}}

        with patch(
            "agentic_docs.build.update_current_from_build",
            side_effect=RuntimeError("current folder locked"),
        ):
            result = _refresh_current(report, resolved, update_current=True)

        self.assertFalse(result["current_updated"])
        self.assertIn("current folder locked", result["current_update_error"])
        self.assertEqual(result["build_id"], "BUILD-1")


if __name__ == "__main__":
    unittest.main()
