import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from agentic_docs.resolver import file_hash
from agentic_docs.source_lock import capture_source_lock, verify_source_lock


class SourceLockTests(unittest.TestCase):
    def test_lock_survives_canonical_change_and_detects_vault_corruption(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "project" / "content.md"
            source.parent.mkdir(parents=True)
            source.write_text("revision one", encoding="utf-8")
            engine = root / "src" / "compiler.py"
            engine.parent.mkdir(parents=True)
            engine.write_text("VERSION = 1\n", encoding="utf-8")
            resolved_inputs = root / "build" / "resolved-inputs.json"
            resolved_inputs.parent.mkdir()
            resolved_inputs.write_text(
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
                            "engine": {
                                "sha256": "ENGINE-1",
                                "files": [
                                    {"path": "src/compiler.py", "sha256": file_hash(engine)}
                                ],
                            },
                        }
                    }
                ),
                encoding="utf-8",
            )
            lock_path = root / "release" / "source-lock.json"
            result = capture_source_lock(root, resolved_inputs, lock_path)
            self.assertTrue(result["verified"])
            self.assertEqual(result["entry_count"], 2)

            source.write_text("revision two", encoding="utf-8")
            self.assertTrue(verify_source_lock(root, lock_path)["valid"])

            lock = json.loads(lock_path.read_text(encoding="utf-8"))
            object_path = root / lock["entries"][0]["object"]
            object_path.write_bytes(b"corrupt")
            verification = verify_source_lock(root, lock_path)
            self.assertFalse(verification["valid"])
            self.assertTrue(any("differs" in issue for issue in verification["issues"]))


if __name__ == "__main__":
    unittest.main()
