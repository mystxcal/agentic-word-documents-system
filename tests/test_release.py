from pathlib import Path
from types import SimpleNamespace
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch
import json

from agentic_docs.diagnostics import DiagnosticBag
from agentic_docs.errors import ReleaseGateError
from agentic_docs.model import GateState
from agentic_docs.release import release_document
from agentic_docs.resolver import file_hash


def resolved_fixture(root: Path, gate_state: GateState = GateState.MET):
    return SimpleNamespace(
        system_root=root,
        profile=SimpleNamespace(release_gates=["review"]),
        manifest=SimpleNamespace(
            id="doc-1",
            release=SimpleNamespace(gates={"review": gate_state}),
            metadata=SimpleNamespace(revision="R00"),
            outputs=SimpleNamespace(basename="Example"),
        ),
    )


def build_report(root: Path, *, passed: bool) -> dict:
    run = root / "build"
    run.mkdir()
    docx = run / "Example.docx"
    pdf = run / "Example.pdf"
    docx.write_bytes(b"docx")
    pdf.write_bytes(b"pdf")
    source = root / "canonical.md"
    source.write_text("canonical", encoding="utf-8")
    (run / "resolved-inputs.json").write_text(
        json.dumps(
            {
                "input_snapshot": {
                    "fingerprint": "INPUT-1",
                    "records": [
                        {
                            "key": "component.body",
                            "category": "content",
                            "path": str(source),
                            "sha256": file_hash(source),
                            "size_bytes": source.stat().st_size,
                        }
                    ],
                    "engine": {"sha256": "ENGINE-1", "files": []},
                }
            }
        ),
        encoding="utf-8",
    )
    return {
        "build_id": "BUILD-1",
        "run_directory": str(run),
        "quality": {"passed": passed},
        "diagnostics": [],
        "artifacts": {
            "docx": str(docx),
            "docx_sha256": file_hash(docx),
            "pdf": str(pdf),
            "pdf_sha256": file_hash(pdf),
        },
    }


class ReleaseTests(unittest.TestCase):
    def test_open_gate_refuses_release_before_build(self):
        with TemporaryDirectory() as directory:
            resolved = resolved_fixture(Path(directory), GateState.OPEN)
            with patch("agentic_docs.release.build_document") as build:
                with self.assertRaises(ReleaseGateError):
                    release_document(resolved, DiagnosticBag())
            build.assert_not_called()

    def test_failed_candidate_does_not_refresh_current(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            resolved = resolved_fixture(root)
            report = build_report(root, passed=False)
            with patch("agentic_docs.release.build_document", return_value=report), patch(
                "agentic_docs.release.update_current_from_build"
            ) as update_current:
                with self.assertRaises(ReleaseGateError):
                    release_document(resolved, DiagnosticBag())
            update_current.assert_not_called()

    def test_successful_release_is_hash_verified_and_then_refreshes_current(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            resolved = resolved_fixture(root)
            report = build_report(root, passed=True)
            current = {
                "current_docx": str(root / "current" / "Example.docx"),
                "current_pdf": str(root / "current" / "Example.pdf"),
            }
            with patch("agentic_docs.release.build_document", return_value=report), patch(
                "agentic_docs.release.update_current_from_build", return_value=current
            ) as update_current, patch(
                "agentic_docs.release.collect_runtime_provenance",
                return_value={"schema": "agentic-runtime-provenance/v1", "test": True},
            ):
                result = release_document(resolved, DiagnosticBag())

            release_root = Path(result["release_directory"])
            self.assertEqual((release_root / "Example.docx").read_bytes(), b"docx")
            self.assertEqual((release_root / "Example.pdf").read_bytes(), b"pdf")
            self.assertTrue((release_root / "release-report.json").is_file())
            self.assertTrue((release_root / "source-lock.json").is_file())
            self.assertTrue((release_root / "resolved-inputs.json").is_file())
            self.assertTrue((release_root / "runtime-provenance.json").is_file())
            self.assertTrue(result["source_lock"]["verified"])
            self.assertTrue(result["current_updated"])
            update_current.assert_called_once()


if __name__ == "__main__":
    unittest.main()
