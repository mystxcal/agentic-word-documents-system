from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageStat


def _ink_ratio(image: Image.Image) -> float:
    grayscale = image.convert("L")
    histogram = grayscale.histogram()
    ink_pixels = sum(histogram[:248])
    return ink_pixels / max(1, grayscale.width * grayscale.height)


def analyze_rendered_pages(render: dict | None, *, allow_blank_pages: bool = False) -> dict:
    """Measure page ink and page-edge activity without guessing document content."""

    images = [Path(value) for value in (render or {}).get("page_images", [])]
    pages = []
    blank_pages = []
    for number, path in enumerate(images, 1):
        with Image.open(path) as image:
            width, height = image.size
            top = image.crop((0, 0, width, max(1, round(height * 0.16))))
            bottom = image.crop((0, max(0, round(height * 0.84)), width, height))
            header_band = image.crop((0, 0, width, max(1, round(height * 0.12))))
            footer_band = image.crop((0, max(0, round(height * 0.88)), width, height))
            ink = _ink_ratio(image)
            page = {
                "page": number,
                "image": str(path),
                "ink_ratio": round(ink, 6),
                "top_band_ink_ratio": round(_ink_ratio(top), 6),
                "bottom_band_ink_ratio": round(_ink_ratio(bottom), 6),
                "header_band_ink_ratio": round(_ink_ratio(header_band), 6),
                "footer_band_ink_ratio": round(_ink_ratio(footer_band), 6),
                "mean_luminance": round(ImageStat.Stat(image.convert("L")).mean[0], 2),
                "blank": ink < 0.00035,
            }
            if page["blank"]:
                blank_pages.append(number)
            pages.append(page)
    return {
        "schema": "agentic-visual-quality/v1",
        "page_count": len(pages),
        "pages": pages,
        "blank_pages": blank_pages,
        "allow_blank_pages": allow_blank_pages,
        "no_unexpected_blank_pages": allow_blank_pages or not blank_pages,
    }


def analyze_page_furniture_preview(
    visual_quality: dict,
    *,
    expect_header: bool,
    expect_footer: bool,
    sample_count: int = 3,
    minimum_ink_ratio: float = 0.0005,
) -> dict:
    """Prove page-edge furniture on the compact preview's fixture pages.

    The fixture deliberately leaves its top and bottom edge bands empty. Any
    ink there therefore comes from the active header/footer, not body content.
    The last three pages cover consecutive page parity without relying on an
    image viewer that may crop white page edges for display.
    """

    pages = list(visual_quality.get("pages", []))
    samples = pages[-sample_count:] if len(pages) >= sample_count else []
    missing_headers = [page["page"] for page in samples if expect_header and page.get("header_band_ink_ratio", 0) < minimum_ink_ratio]
    missing_footers = [page["page"] for page in samples if expect_footer and page.get("footer_band_ink_ratio", 0) < minimum_ink_ratio]
    passed = len(samples) == sample_count and not missing_headers and not missing_footers
    return {
        "schema": "agentic-page-furniture-visual-proof/v1",
        "applicable": True,
        "sample_pages": [page["page"] for page in samples],
        "expect_header": bool(expect_header),
        "expect_footer": bool(expect_footer),
        "minimum_ink_ratio": minimum_ink_ratio,
        "missing_header_pages": missing_headers,
        "missing_footer_pages": missing_footers,
        "passed": passed,
    }
