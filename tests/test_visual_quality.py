from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from PIL import Image, ImageDraw

from agentic_docs.visual_quality import analyze_page_furniture_preview, analyze_rendered_pages


class VisualQualityTests(unittest.TestCase):
    def test_detects_blank_page_and_reports_edge_band_activity(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            blank = root / "page-1.png"
            marked = root / "page-2.png"
            Image.new("RGB", (600, 800), "white").save(blank)
            image = Image.new("RGB", (600, 800), "white")
            draw = ImageDraw.Draw(image)
            draw.rectangle((30, 20, 570, 65), fill="black")
            draw.rectangle((30, 740, 570, 780), fill="black")
            image.save(marked)

            result = analyze_rendered_pages({"page_images": [str(blank), str(marked)]})

            self.assertEqual([1], result["blank_pages"])
            self.assertFalse(result["no_unexpected_blank_pages"])
            self.assertGreater(result["pages"][1]["top_band_ink_ratio"], 0)
            self.assertGreater(result["pages"][1]["bottom_band_ink_ratio"], 0)
            self.assertGreater(result["pages"][1]["header_band_ink_ratio"], 0)
            self.assertGreater(result["pages"][1]["footer_band_ink_ratio"], 0)

    def test_intentional_blank_pages_can_be_allowed(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "page-1.png"
            Image.new("RGB", (100, 100), "white").save(path)
            result = analyze_rendered_pages(
                {"page_images": [str(path)]},
                allow_blank_pages=True,
            )
            self.assertTrue(result["no_unexpected_blank_pages"])

    def test_page_furniture_proof_checks_three_full_page_edge_samples(self):
        pages = [
            {"page": 1, "header_band_ink_ratio": 0, "footer_band_ink_ratio": 0},
            {"page": 2, "header_band_ink_ratio": 0.02, "footer_band_ink_ratio": 0.01},
            {"page": 3, "header_band_ink_ratio": 0.02, "footer_band_ink_ratio": 0.01},
            {"page": 4, "header_band_ink_ratio": 0.02, "footer_band_ink_ratio": 0.01},
        ]

        passed = analyze_page_furniture_preview(
            {"pages": pages},
            expect_header=True,
            expect_footer=True,
        )
        pages[-1]["footer_band_ink_ratio"] = 0
        failed = analyze_page_furniture_preview(
            {"pages": pages},
            expect_header=True,
            expect_footer=True,
        )

        self.assertTrue(passed["passed"])
        self.assertFalse(failed["passed"])
        self.assertEqual([4], failed["missing_footer_pages"])


if __name__ == "__main__":
    unittest.main()
