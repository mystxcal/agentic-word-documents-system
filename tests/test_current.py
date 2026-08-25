import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from agentic_docs.current import update_current_from_build
from agentic_docs.resolver import file_hash


def build_report(root: Path, build_id: str, docx_payload: bytes, pdf_payload: bytes) -> dict:
    run = root / build_id
    run.mkdir(parents=True)
    docx = run / "Example.docx"
    pdf = run / "Example.pdf"
    docx.write_bytes(docx_payload)
    pdf.write_bytes(pdf_payload)
    return {
        "schema": "agentic-build-report/v2",
        "build_id": build_id,
        "artifacts": {
            "docx": str(docx),
            "docx_sha256": file_hash(docx),
            "pdf": str(pdf),
            "pdf_sha256": file_hash(pdf),
        },
    }


class CurrentUpdateTests(unittest.TestCase):
    def test_preserves_manual_current_edit_before_refresh(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            current = root / "current"
            first = build_report(root / "builds", "BUILD-1", b"first-docx", b"first-pdf")
            update_current_from_build(first, current, basename="Example")
            current_docx = current / "Example.docx"
            current_docx.write_bytes(b"coworker-edited-docx")

            second = build_report(root / "builds", "BUILD-2", b"second-docx", b"second-pdf")
            result = update_current_from_build(second, current, basename="Example")

            self.assertEqual(current_docx.read_bytes(), b"second-docx")
            self.assertEqual((current / "Example.pdf").read_bytes(), b"second-pdf")
            self.assertEqual(len(result["manual_files_preserved"]), 1)
            retained = Path(result["manual_files_preserved"][0]["to"])
            self.assertEqual(retained.read_bytes(), b"coworker-edited-docx")
            current_report = json.loads((current / "build-report.json").read_text(encoding="utf-8"))
            self.assertEqual(current_report["build_id"], "BUILD-2")

    def test_removes_prior_generated_basename_without_losing_immutable_build(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            current = root / "current"
            first = build_report(root / "builds", "BUILD-1", b"first-docx", b"first-pdf")
            update_current_from_build(first, current, basename="Old Name")
            second = build_report(root / "builds", "BUILD-2", b"second-docx", b"second-pdf")
            update_current_from_build(second, current, basename="New Name")

            self.assertFalse((current / "Old Name.docx").exists())
            self.assertFalse((current / "Old Name.pdf").exists())
            self.assertTrue((current / "New Name.docx").is_file())
            self.assertTrue(Path(first["artifacts"]["docx"]).is_file())


if __name__ == "__main__":
    unittest.main()
