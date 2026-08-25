from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from agentic_docs.errors import IntegrityError
from agentic_docs.publishing import (
    load_immutable_build_report,
    publish_build,
    require_delivery_alignment,
)
from agentic_docs.resolver import file_hash


class PublishingTests(unittest.TestCase):
    def _report(self, root: Path) -> dict:
        docx = root / "Example.docx"
        pdf = root / "Example.pdf"
        docx.write_bytes(b"docx-build")
        pdf.write_bytes(b"pdf-build")
        return {
            "build_id": "BUILD-1",
            "artifacts": {
                "docx": str(docx),
                "docx_sha256": file_hash(docx),
                "pdf": str(pdf),
                "pdf_sha256": file_hash(pdf),
            },
        }

    def test_publishes_and_hash_verifies_pair(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            build = root / "build"
            build.mkdir()
            destination = root / "delivery"
            report = self._report(build)
            result = publish_build(report, destination)
            self.assertTrue(result["changed"])
            self.assertEqual(file_hash(destination / "Example.docx"), report["artifacts"]["docx_sha256"])
            self.assertEqual(file_hash(destination / "Example.pdf"), report["artifacts"]["pdf_sha256"])

    def test_replace_policy_retains_displaced_file(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            build = root / "build"
            build.mkdir()
            destination = root / "delivery"
            destination.mkdir()
            report = self._report(build)
            (destination / "Example.docx").write_bytes(b"old-docx")
            (destination / "Example.pdf").write_bytes(b"old-pdf")
            result = publish_build(report, destination, policy="replace")
            self.assertTrue(result["changed"])
            history = Path(result["history"])
            self.assertEqual((history / "Example.docx").read_bytes(), b"old-docx")
            self.assertEqual((history / "Example.pdf").read_bytes(), b"old-pdf")
            self.assertEqual(file_hash(destination / "Example.docx"), report["artifacts"]["docx_sha256"])

    def test_default_delivery_requires_current_and_sources_to_match_build(self):
        with self.assertRaises(IntegrityError) as captured:
            require_delivery_alignment(
                {
                    "immutable_integrity": True,
                    "current_integrity": False,
                    "sources_current": False,
                }
            )
        self.assertIn("out of sync", str(captured.exception))
        self.assertIn("--allow-out-of-sync", str(captured.exception))

    def test_aligned_delivery_state_is_accepted(self):
        require_delivery_alignment(
            {
                "immutable_integrity": True,
                "current_integrity": True,
                "sources_current": True,
            }
        )

    def test_loads_an_explicit_immutable_build_by_identity(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            report_path = root / "builds" / "doc-1" / "BUILD-1" / "build-report.json"
            report_path.parent.mkdir(parents=True)
            report_path.write_text('{"build_id":"BUILD-1"}', encoding="utf-8")

            report, selected_path = load_immutable_build_report(root, "doc-1", "BUILD-1")

            self.assertEqual(report["build_id"], "BUILD-1")
            self.assertEqual(selected_path, report_path.resolve())

    def test_explicit_build_selection_rejects_path_traversal(self):
        with TemporaryDirectory() as directory:
            with self.assertRaises(IntegrityError):
                load_immutable_build_report(Path(directory), "doc-1", "..\\other")

    def test_component_preview_cannot_be_published_as_a_document(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            build = root / "build"
            build.mkdir()
            report = self._report(build)
            report["mode"] = "component-preview"
            with self.assertRaisesRegex(IntegrityError, "cannot be delivered"):
                publish_build(report, root / "delivery")

    def test_lightweight_preview_cannot_be_published_as_a_document(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            build = root / "build"
            build.mkdir()
            report = self._report(build)
            report["mode"] = "lightweight-preview"
            report["content_scope"] = "lightweight"
            with self.assertRaisesRegex(IntegrityError, "cannot be delivered"):
                publish_build(report, root / "delivery")

    def test_failed_complete_quality_proof_cannot_be_published(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            build = root / "build"
            build.mkdir()
            report = self._report(build)
            report["content_scope"] = "complete"
            report["quality"] = {"passed": False, "release_ready": False}
            with self.assertRaisesRegex(IntegrityError, "quality proof"):
                publish_build(report, root / "delivery")


if __name__ == "__main__":
    unittest.main()
