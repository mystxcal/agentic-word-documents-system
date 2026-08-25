from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from docx import Document

from agentic_docs.component_cache import ComponentAdapterCache


class ComponentCacheTests(unittest.TestCase):
    def test_verified_component_is_reused_and_corruption_is_refused(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "compiled.docx"
            Document().save(source)
            payload = {
                "schema": "agentic-component-adapter-cache/v1",
                "adapter": "markdown",
                "component_id": "body",
                "fingerprint": "A" * 64,
            }
            cache = ComponentAdapterCache(root)
            stored, event = cache.store(payload, source)
            self.assertTrue(stored.is_file())
            self.assertTrue(event["stored"])

            reused, event = cache.lookup(payload)
            self.assertEqual(reused, stored)
            self.assertTrue(event["hit"])

            stored.write_bytes(b"not a DOCX package")
            refused, event = cache.lookup(payload)
            self.assertIsNone(refused)
            self.assertEqual(event["reason"], "invalid")


if __name__ == "__main__":
    unittest.main()
