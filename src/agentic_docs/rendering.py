from __future__ import annotations

import os
import re
import shutil
import subprocess
from pathlib import Path

from pypdf import PdfReader

from .errors import DocumentSystemError


WORD_FIELD_ERROR_MESSAGES = (
    "Error! Bookmark not defined.",
    "Error! Reference source not found.",
    "Error! No table of contents entries found.",
    "Error! No table of figures entries found.",
    "Error! Not a valid bookmark self-reference.",
)


def _candidate_poppler_directories() -> list[Path]:
    values: list[Path] = []
    configured = os.environ.get("AGENTIC_DOCS_POPPLER_BIN")
    if configured:
        values.append(Path(configured))
    values.append(
        Path.home()
        / ".cache"
        / "codex-runtimes"
        / "codex-primary-runtime"
        / "dependencies"
        / "native"
        / "poppler"
        / "Library"
        / "bin"
    )
    return values


def _find_program(name: str) -> Path | None:
    located = shutil.which(name)
    if located:
        return Path(located).resolve()
    executable = name if name.lower().endswith(".exe") else name + ".exe"
    for directory in _candidate_poppler_directories():
        candidate = directory / executable
        if candidate.is_file():
            return candidate.resolve()
    return None


def poppler_status() -> dict:
    pdftoppm = _find_program("pdftoppm")
    pdfinfo = _find_program("pdfinfo")
    return {
        "available": bool(pdftoppm and pdfinfo),
        "pdftoppm": str(pdftoppm) if pdftoppm else None,
        "pdfinfo": str(pdfinfo) if pdfinfo else None,
    }


def _run(command: list[str], *, timeout: int = 180) -> str:
    completed = subprocess.run(
        command,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
        check=False,
    )
    if completed.returncode != 0:
        detail = (completed.stderr or completed.stdout).strip()
        raise DocumentSystemError(f"PDF inspection command failed: {detail}")
    return completed.stdout


def _parse_pdfinfo(text: str) -> dict:
    values: dict[str, str | int] = {}
    for line in text.splitlines():
        if ":" not in line:
            continue
        key, value = line.split(":", 1)
        normalized = re.sub(r"[^a-z0-9]+", "_", key.strip().lower()).strip("_")
        clean_value = value.strip()
        values[normalized] = int(clean_value) if normalized == "pages" and clean_value.isdigit() else clean_value
    return values


def pdf_information(pdf_path: Path) -> dict:
    """Return Poppler metadata for one PDF without rendering any pages."""

    status = poppler_status()
    if not status["available"]:
        raise DocumentSystemError(
            "Poppler was not found. Set AGENTIC_DOCS_POPPLER_BIN to the directory containing pdftoppm.exe and pdfinfo.exe."
        )
    path = Path(pdf_path).resolve()
    if not path.is_file():
        raise DocumentSystemError(f"PDF source does not exist: {path}")
    return {
        "path": str(path),
        "info": _parse_pdfinfo(_run([status["pdfinfo"], str(path)])),
        "program": status["pdfinfo"],
    }


def find_word_field_errors(page_texts: list[str]) -> list[dict[str, int | str]]:
    """Locate known Word field failures in extracted PDF text."""

    errors: list[dict[str, int | str]] = []
    for page_number, text in enumerate(page_texts, 1):
        for message in WORD_FIELD_ERROR_MESSAGES:
            count = text.count(message)
            if count:
                errors.append({"page": page_number, "message": message, "count": count})
    return errors


def word_field_error_inventory(pdf_path: Path) -> dict:
    """Inspect the delivered PDF, not just the DOCX, for broken Word fields."""

    path = Path(pdf_path).resolve()
    if not path.is_file():
        raise DocumentSystemError(f"PDF source does not exist: {path}")
    try:
        reader = PdfReader(path)
        page_texts = [page.extract_text() or "" for page in reader.pages]
    except Exception as exc:
        raise DocumentSystemError(f"Could not inspect PDF field results: {exc}") from exc
    errors = find_word_field_errors(page_texts)
    return {
        "schema": "agentic-word-field-errors/v1",
        "source": str(path),
        "page_count": len(page_texts),
        "error_count": sum(int(item["count"]) for item in errors),
        "errors": errors,
    }


def render_pdf(pdf_path: Path, output_directory: Path, *, dpi: int = 110) -> dict:
    """Render every PDF page through Poppler and return machine-checkable facts."""
    status = poppler_status()
    if not status["available"]:
        raise DocumentSystemError(
            "Poppler was not found. Set AGENTIC_DOCS_POPPLER_BIN to the directory containing pdftoppm.exe and pdfinfo.exe."
        )
    output_directory.mkdir(parents=True, exist_ok=False)
    prefix = output_directory / "page"
    _run([status["pdftoppm"], "-png", "-r", str(dpi), str(pdf_path), str(prefix)])
    info = _parse_pdfinfo(_run([status["pdfinfo"], str(pdf_path)]))
    pages = sorted(output_directory.glob("page-*.png"))
    expected_pages = info.get("pages")
    if isinstance(expected_pages, int) and expected_pages != len(pages):
        raise DocumentSystemError(
            f"PDF render produced {len(pages)} images for a {expected_pages}-page PDF"
        )
    return {
        "renderer": "poppler",
        "dpi": dpi,
        "directory": str(output_directory),
        "page_images": [str(path) for path in pages],
        "page_count": len(pages),
        "pdf_info": info,
        "programs": status,
    }


def render_pdf_pages(
    pdf_path: Path,
    output_directory: Path,
    pages: list[int],
    *,
    dpi: int = 150,
) -> dict:
    """Render an explicit ordered page selection without assuming document content."""

    status = poppler_status()
    if not status["available"]:
        raise DocumentSystemError(
            "Poppler was not found. Set AGENTIC_DOCS_POPPLER_BIN to the directory containing pdftoppm.exe and pdfinfo.exe."
        )
    if not Path(pdf_path).is_file():
        raise DocumentSystemError(f"PDF source does not exist: {pdf_path}")
    if not pages or any(not isinstance(page, int) or isinstance(page, bool) or page < 1 for page in pages):
        raise DocumentSystemError("PDF page selection must contain positive one-based page numbers")
    info = _parse_pdfinfo(_run([status["pdfinfo"], str(pdf_path)]))
    page_count = info.get("pages")
    if isinstance(page_count, int) and max(pages) > page_count:
        raise DocumentSystemError(
            f"PDF {pdf_path.name} has {page_count} pages; page {max(pages)} was requested"
        )
    output_directory.mkdir(parents=True, exist_ok=False)
    rendered = []
    for index, page in enumerate(pages, 1):
        prefix = output_directory / f"selection-{index:04d}-source-{page:04d}"
        _run(
            [
                status["pdftoppm"],
                "-png",
                "-r",
                str(dpi),
                "-f",
                str(page),
                "-l",
                str(page),
                "-singlefile",
                str(pdf_path),
                str(prefix),
            ]
        )
        output = prefix.with_suffix(".png")
        if not output.is_file():
            raise DocumentSystemError(f"PDF page render did not create {output}")
        rendered.append({"source_page": page, "image": str(output)})
    return {
        "renderer": "poppler",
        "dpi": dpi,
        "source": str(Path(pdf_path).resolve()),
        "pdf_info": info,
        "pages": rendered,
        "programs": status,
    }
