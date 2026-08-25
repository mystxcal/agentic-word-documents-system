from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest

from agentic_docs.inputs import compare_input_snapshots, resolved_input_snapshot


class InputSnapshotTests(unittest.TestCase):
    def _resolved(self, root: Path):
        for relative, value in {
            "document.jsonc": "document",
            "kit.jsonc": "kit",
            "profile.jsonc": "profile",
            "project.jsonc": "project",
            "shell.docx": "shell",
            "styles.docx": "styles",
            "body.md": "body",
        }.items():
            (root / relative).write_text(value, encoding="utf-8")
        component = SimpleNamespace(source_path=root / "body.md")
        return SimpleNamespace(
            system_root=root,
            manifest_path=root / "document.jsonc",
            kit_path=root / "kit.jsonc",
            profile_path=root / "profile.jsonc",
            project_path=root / "project.jsonc",
            shell_path=root / "shell.docx",
            presentation_paths={"styles": root / "styles.docx", "cover": None},
            components={"body": component},
        )

    def test_snapshot_diff_names_the_exact_changed_input(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            resolved = self._resolved(root)
            before = resolved_input_snapshot(resolved)
            self.assertFalse(compare_input_snapshots(before, before)["changed"])

            (root / "body.md").write_text("revised body", encoding="utf-8")
            after = resolved_input_snapshot(resolved)
            change = compare_input_snapshots(after, before)
            self.assertTrue(change["changed"])
            self.assertEqual([item["key"] for item in change["items"]], ["component.body"])
            self.assertEqual(change["counts"]["changed"], 1)


if __name__ == "__main__":
    unittest.main()
