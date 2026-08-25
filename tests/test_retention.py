import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from agentic_docs.retention import manage_retention


class RetentionTests(unittest.TestCase):
    @staticmethod
    def _build(root: Path, build_id: str, mode: str) -> None:
        directory = root / "builds" / "doc-1" / build_id
        directory.mkdir(parents=True)
        (directory / "artifact.bin").write_bytes(build_id.encode())
        (directory / "build-report.json").write_text(
            json.dumps({"build_id": build_id, "mode": mode}),
            encoding="utf-8",
        )

    def test_plan_and_apply_never_delete_current_or_released_builds(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            for build_id in ("D001", "D002", "D003", "D004", "D005"):
                self._build(root, build_id, "draft")
            for build_id in ("P001", "P002", "P003"):
                self._build(root, build_id, "component-preview")
            self._build(root, "R001", "release")

            current = root / "current" / "doc-1"
            current.mkdir(parents=True)
            (current / "build-report.json").write_text(
                json.dumps({"build_id": "D005"}), encoding="utf-8"
            )
            release = root / "releases" / "doc-1" / "R00" / "release-1"
            release.mkdir(parents=True)
            (release / "release-report.json").write_text(
                json.dumps({"build_id": "D001"}), encoding="utf-8"
            )

            plan = manage_retention(root, "doc-1", keep_drafts=2, keep_previews=1)
            self.assertFalse(plan["applied"])
            self.assertEqual(
                {item["build_id"] for item in plan["candidates"]},
                {"D002", "D003", "P001", "P002"},
            )

            result = manage_retention(
                root,
                "doc-1",
                keep_drafts=2,
                keep_previews=1,
                apply=True,
            )
            archive = Path(result["archive"])
            self.assertTrue((archive / "retention-receipt.json").is_file())
            for build_id in ("D002", "D003", "P001", "P002"):
                self.assertFalse((root / "builds" / "doc-1" / build_id).exists())
                self.assertTrue((archive / build_id).is_dir())
            for build_id in ("D001", "D004", "D005", "P003", "R001"):
                self.assertTrue((root / "builds" / "doc-1" / build_id).is_dir())


if __name__ == "__main__":
    unittest.main()
