import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from agentic_docs.errors import DocumentSystemError
from agentic_docs.locking import document_lock


class LockingTests(unittest.TestCase):
    def test_refuses_concurrent_mutation_and_releases_cleanly(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            with document_lock(root, "doc-1", "build") as path:
                self.assertTrue(path.is_file())
                with self.assertRaises(DocumentSystemError):
                    with document_lock(root, "doc-1", "replace"):
                        pass
            self.assertFalse((root / ".locks" / "doc-1.json").exists())

    def test_archives_dead_process_lock(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            lock_root = root / ".locks"
            lock_root.mkdir()
            stale = lock_root / "doc-1.json"
            stale.write_text(
                json.dumps({"pid": 2147483647, "operation": "old-build", "token": "old"}),
                encoding="utf-8",
            )
            with document_lock(root, "doc-1", "build"):
                archived = list((lock_root / "stale").glob("doc-1-*.json"))
                self.assertEqual(len(archived), 1)


if __name__ == "__main__":
    unittest.main()
