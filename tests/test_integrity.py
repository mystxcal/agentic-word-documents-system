import json
from pathlib import Path
from types import SimpleNamespace
from tempfile import TemporaryDirectory
import unittest

from agentic_docs.integrity import audit_document
from agentic_docs.resolver import file_hash
from agentic_docs.source_lock import capture_source_lock


class IntegrityTests(unittest.TestCase):
    def _resolved(self, root: Path):
        source = root / "source.docx"
        source.write_bytes(b"source")
        manifest = root / "document.jsonc"
        manifest.write_bytes(b"manifest")
        component = SimpleNamespace(source_path=source, source_hash=file_hash(source))
        return SimpleNamespace(
            system_root=root,
            manifest_path=manifest,
            manifest=SimpleNamespace(
                id="doc-1",
                outputs=SimpleNamespace(basename="Example"),
            ),
            components={"body": component},
        )

    def _build_state(self, root: Path, resolved) -> None:
        run = root / "builds" / "doc-1" / "BUILD-1"
        run.mkdir(parents=True)
        docx = run / "Example.docx"
        pdf = run / "Example.pdf"
        docx.write_bytes(b"docx")
        pdf.write_bytes(b"pdf")
        current = root / "current" / "doc-1"
        current.mkdir(parents=True)
        (current / "Example.docx").write_bytes(docx.read_bytes())
        (current / "Example.pdf").write_bytes(pdf.read_bytes())
        report = {
            "build_id": "BUILD-1",
            "artifacts": {
                "docx": str(docx),
                "docx_sha256": file_hash(docx),
                "pdf": str(pdf),
                "pdf_sha256": file_hash(pdf),
            },
            "resolved": {
                "layers": {
                    "document": {
                        "path": str(resolved.manifest_path),
                        "sha256": file_hash(resolved.manifest_path),
                    }
                },
                "components": [
                    {
                        "id": "body",
                        "source_hash": resolved.components["body"].source_hash,
                    }
                ]
            },
        }
        (current / "build-report.json").write_text(json.dumps(report), encoding="utf-8")

    def test_detects_current_drift_and_stale_files(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            resolved = self._resolved(root)
            self._build_state(root, resolved)
            clean = audit_document(resolved)
            self.assertTrue(clean["ready"])
            current = root / "current" / "doc-1"
            (current / "Example.docx").write_bytes(b"edited")
            (current / "old.tmp").write_bytes(b"old")
            drifted = audit_document(resolved)
            self.assertFalse(drifted["ready"])
            codes = [issue["code"] for issue in drifted["issues"]]
            self.assertIn("ARTIFACT_DRIFT", codes)
            self.assertIn("CURRENT_FOLDER_STALE_FILES", codes)

    def test_detects_manifest_or_presentation_input_drift(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            resolved = self._resolved(root)
            self._build_state(root, resolved)
            resolved.manifest_path.write_bytes(b"manifest changed")
            drifted = audit_document(resolved)
            self.assertFalse(drifted["sources_current"])
            self.assertIn(
                "CANONICAL_INPUT_CHANGED",
                [issue["code"] for issue in drifted["issues"]],
            )

    def test_rechecks_controlled_release_artifacts_and_source_vault(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            resolved = self._resolved(root)
            self._build_state(root, resolved)
            release = root / "releases" / "doc-1" / "R00" / "REL-1"
            release.mkdir(parents=True)
            docx = release / "Example.docx"
            pdf = release / "Example.pdf"
            docx.write_bytes(b"released docx")
            pdf.write_bytes(b"released pdf")
            inputs = release / "resolved-inputs.json"
            inputs.write_text(
                json.dumps(
                    {
                        "input_snapshot": {
                            "fingerprint": "INPUT-1",
                            "records": [
                                {
                                    "key": "component.body",
                                    "category": "content",
                                    "path": str(resolved.components["body"].source_path),
                                    "sha256": resolved.components["body"].source_hash,
                                    "size_bytes": resolved.components["body"].source_path.stat().st_size,
                                }
                            ],
                            "engine": {"sha256": "ENGINE-1", "files": []},
                        }
                    }
                ),
                encoding="utf-8",
            )
            lock = capture_source_lock(root, inputs, release / "source-lock.json")
            runtime = release / "runtime-provenance.json"
            runtime.write_text(json.dumps({"schema": "agentic-runtime-provenance/v1"}), encoding="utf-8")
            (release / "release-report.json").write_text(
                json.dumps(
                    {
                        "build_id": "REL-1",
                        "artifacts": {
                            "docx_sha256": file_hash(docx),
                            "pdf_sha256": file_hash(pdf),
                        },
                        "release_artifacts": {"docx": str(docx), "pdf": str(pdf)},
                        "source_lock": lock,
                        "runtime_provenance": {
                            "path": str(runtime),
                            "sha256": file_hash(runtime),
                        },
                    }
                ),
                encoding="utf-8",
            )
            clean = audit_document(resolved)
            self.assertTrue(clean["release_integrity"]["valid"])

            lock_payload = json.loads((release / "source-lock.json").read_text(encoding="utf-8"))
            (root / lock_payload["entries"][0]["object"]).write_bytes(b"corrupt")
            drifted = audit_document(resolved)
            self.assertFalse(drifted["ready"])
            self.assertIn(
                "RELEASE_SOURCE_LOCK_INVALID",
                [issue["code"] for issue in drifted["issues"]],
            )


if __name__ == "__main__":
    unittest.main()
