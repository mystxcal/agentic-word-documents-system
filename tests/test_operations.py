from pathlib import Path
from types import SimpleNamespace
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from docx import Document

from agentic_docs.diagnostics import DiagnosticBag
from agentic_docs.docx_text import TextReplacement
from agentic_docs.errors import OperationPartialError, RevisionError
from agentic_docs.model import ComponentType, Ownership
from agentic_docs.operations import revise_document
from agentic_docs.resolver import file_hash
from agentic_docs.word.ooxml import wrap_elements


def tagged_docx(path: Path, text: str) -> None:
    document = Document()
    paragraph = document.add_paragraph(text)
    wrap_elements("AGDOC.BODY.TEST", [paragraph._p])
    document.save(path)


def resolved_fixture(root: Path, source: Path):
    component = SimpleNamespace(
        id="body",
        declaration=SimpleNamespace(
            type=ComponentType.DOCUMENT,
            ownership=Ownership.WORD_FRAGMENT,
        ),
        source_path=source,
    )
    manifest = root / "document.jsonc"
    manifest.write_text("{}", encoding="utf-8")
    return SimpleNamespace(
        manifest_path=manifest,
        manifest=SimpleNamespace(id="doc-1"),
        components={"body": component},
    )


class RevisionOperationTests(unittest.TestCase):
    def test_allowed_zero_match_is_a_true_no_op(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "body.docx"
            tagged_docx(source, "already correct")
            original_hash = file_hash(source)
            resolved = resolved_fixture(root, source)
            with patch("agentic_docs.operations.build_document") as build:
                report = revise_document(
                    resolved,
                    DiagnosticBag(),
                    [TextReplacement("missing", "replacement")],
                    allow_zero=True,
                )
            build.assert_not_called()
            self.assertFalse(report["changed"])
            self.assertEqual(file_hash(source), original_hash)
            self.assertFalse((root / ".history").exists())

    def test_build_failure_restores_original_canonical_source(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "body.docx"
            tagged_docx(source, "fibre link")
            original_hash = file_hash(source)
            resolved = resolved_fixture(root, source)

            with patch("agentic_docs.operations.resolve_document", return_value=resolved), patch(
                "agentic_docs.operations.build_document", side_effect=RuntimeError("simulated build failure")
            ):
                with self.assertRaises(RevisionError):
                    revise_document(
                        resolved,
                        DiagnosticBag(),
                        [TextReplacement("fibre", "fiber")],
                    )

            self.assertEqual(file_hash(source), original_hash)
            backups = list((root / ".history" / "revisions").glob("*/body.docx"))
            self.assertEqual(len(backups), 1)

    def test_publish_failure_reports_partial_success_and_keeps_receipt(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "body.docx"
            tagged_docx(source, "fibre link")
            resolved = resolved_fixture(root, source)
            build_report = {
                "build_id": "BUILD-1",
                "run_directory": str(root / "build"),
                "current_docx": str(root / "current.docx"),
                "current_pdf": str(root / "current.pdf"),
                "quality": {"passed": True},
            }

            with patch("agentic_docs.operations.resolve_document", return_value=resolved), patch(
                "agentic_docs.operations.build_document", return_value=build_report
            ), patch("agentic_docs.operations.publish_build", side_effect=RuntimeError("USB unavailable")):
                with self.assertRaises(OperationPartialError) as captured:
                    revise_document(
                        resolved,
                        DiagnosticBag(),
                        [TextReplacement("fibre", "fiber")],
                        publish_destination=root / "delivery",
                    )

            self.assertEqual(captured.exception.report["build"]["build_id"], "BUILD-1")
            self.assertFalse(captured.exception.report["publish"]["succeeded"])
            receipts = list((root / ".history" / "revisions").glob("*/revision-report.json"))
            self.assertEqual(len(receipts), 1)
            self.assertNotEqual(file_hash(source), captured.exception.report["source_before_sha256"])

    def test_checked_revision_edits_canonical_markdown_without_a_build(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "body.md"
            source.write_text("# Method\n\nInstall fibre links.", encoding="utf-8")
            component = SimpleNamespace(
                id="body",
                declaration=SimpleNamespace(
                    type=ComponentType.DOCUMENT,
                    ownership=Ownership.SOURCE,
                ),
                source_path=source,
            )
            manifest = root / "document.jsonc"
            manifest.write_text("{}", encoding="utf-8")
            resolved = SimpleNamespace(
                manifest_path=manifest,
                manifest=SimpleNamespace(id="doc-1"),
                components={"body": component},
            )

            report = revise_document(
                resolved,
                DiagnosticBag(),
                [TextReplacement("fibre", "fiber", whole_word=True)],
                expected_total=1,
                build=False,
            )

            self.assertTrue(report["changed"])
            self.assertIn("fiber", source.read_text(encoding="utf-8"))
            self.assertTrue(Path(report["backup"]).is_file())


if __name__ == "__main__":
    unittest.main()
