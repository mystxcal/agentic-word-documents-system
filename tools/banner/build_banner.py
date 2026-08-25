#!/usr/bin/env python3
"""Render the project banner from its canonical source.

The artwork lives in ``banner.template.html`` as a deterministic canvas scene:
every frame is a pure function of loop time, so a rebuild reproduces the same
pixels. This script inlines the vendored faces, drives a headless capture, and
encodes the looping GIF plus a static poster frame.

    python tools/banner/build_banner.py                 # both themes
    python tools/banner/build_banner.py --theme light
    python tools/banner/build_banner.py --work build/banner   # keep the frames

Requirements: Node with a puppeteer package, Chrome, ``ffmpeg``, ``gifsicle``.
The banner is a release asset, not part of a document build, so nothing in the
compiler depends on this script.
"""

from __future__ import annotations

import argparse
import base64
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
REPO = HERE.parent.parent
TEMPLATE = HERE / "banner.template.html"
CAPTURE = HERE / "capture.js"
FONTS = HERE / "fonts"

FONT_SLOTS = {
    "__F_DISPLAY__": "SourceSerif4-Display.woff2",
    "__F_TEXT__": "SourceSerif4-Text.woff2",
    "__F_ITALIC__": "SourceSerif4-Italic-Text.woff2",
    "__F_UI__": "Inter-var.woff2",
    "__F_MONO__": "JetBrainsMono-var.woff2",
}

THEME_FILES = {
    "light": ("banner.gif", "banner.png"),
    "dark": ("banner-dark.gif", "banner-dark.png"),
}

LOOP_SECONDS = 12.0


def need(tool: str) -> str:
    found = shutil.which(tool)
    if not found:
        sys.exit(f"required tool not found on PATH: {tool}")
    return found


def inline_fonts(dest: Path) -> Path:
    """Write a self-contained render page.

    A capture over ``file://`` cannot fetch sibling font files, and a face that
    arrived late would silently change metrics part-way through a run. Inlining
    removes both failure modes.
    """
    html = TEMPLATE.read_text(encoding="utf-8")
    for slot, name in FONT_SLOTS.items():
        blob = base64.b64encode((FONTS / name).read_bytes()).decode("ascii")
        html = html.replace(slot, f"data:font/woff2;base64,{blob}")
    if "__F_" in html:
        sys.exit("template still has unfilled font slots")
    page = dest / "banner.page.html"
    page.write_text(html, encoding="utf-8")
    return page


def capture(page: Path, frames: Path, theme: str, fps: int, phase: float, scale: int) -> None:
    frames.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        [need("node"), str(CAPTURE), str(page), str(frames), theme,
         str(fps), str(phase), str(scale)],
        check=True,
    )


def encode(frames: Path, out_gif: Path, fps: int, colors: int) -> None:
    ffmpeg = need("ffmpeg")
    palette = frames.parent / f"palette-{out_gif.stem}.png"
    # One global palette over the whole loop. The art is flat, so a single
    # palette is exact and a per-frame palette would only add churn the
    # encoder then has to store.
    subprocess.run(
        [ffmpeg, "-v", "error", "-y", "-framerate", str(fps),
         "-i", str(frames / "f_%04d.png"),
         "-vf", f"palettegen=max_colors={colors}:stats_mode=full", str(palette)],
        check=True,
    )
    raw = frames.parent / f"raw-{out_gif.name}"
    # dither=none: with a palette that already covers the art, dithering only
    # adds noise that defeats frame differencing.
    subprocess.run(
        [ffmpeg, "-v", "error", "-y", "-framerate", str(fps),
         "-i", str(frames / "f_%04d.png"), "-i", str(palette),
         "-lavfi", "paletteuse=dither=none:diff_mode=rectangle",
         "-loop", "0", str(raw)],
        check=True,
    )
    out_gif.parent.mkdir(parents=True, exist_ok=True)
    subprocess.run([need("gifsicle"), "-O3", str(raw), "-o", str(out_gif)], check=True)


def verify(out_gif: Path, fps: int) -> str:
    """Confirm the encoded loop still runs for exactly one period."""
    try:
        from PIL import Image, ImageSequence
    except ImportError:
        return "duration unverified (Pillow not importable)"
    total, count = 0, 0
    with Image.open(out_gif) as im:
        for f in ImageSequence.Iterator(im):
            total += f.info.get("duration", 0)
            count += 1
    ms = round(LOOP_SECONDS * 1000)
    ok = "ok" if abs(total - ms) <= 1000 // fps else "MISMATCH"
    return f"{count} frames, {total} ms ({ok}, expected {ms} ms)"


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--theme", choices=["light", "dark", "both"], default="both")
    ap.add_argument("--fps", type=int, default=25,
                    help="frames per second; 25 divides the GIF centisecond clock exactly")
    ap.add_argument("--scale", type=int, default=2,
                    help="output scale over the 1280x400 design grid")
    ap.add_argument("--phase", type=float, default=10.8,
                    help="loop time of frame 0, so the poster frame shows a proven build")
    ap.add_argument("--colors", type=int, default=200)
    ap.add_argument("--out", type=Path, default=REPO / "docs" / "assets")
    ap.add_argument("--work", type=Path, default=None,
                    help="keep frames and intermediates here (default: a temporary directory)")
    args = ap.parse_args()

    if 100 % args.fps:
        print(f"note: {args.fps} fps does not divide the GIF centisecond clock; "
              "frame delays will be uneven", file=sys.stderr)

    themes = ["light", "dark"] if args.theme == "both" else [args.theme]
    holder = None
    if args.work:
        work = args.work
        work.mkdir(parents=True, exist_ok=True)
    else:
        holder = tempfile.TemporaryDirectory(prefix="awd-banner-")
        work = Path(holder.name)

    try:
        page = inline_fonts(work)
        for theme in themes:
            gif_name, png_name = THEME_FILES[theme]
            frames = work / f"frames-{theme}"
            print(f"[{theme}] capturing…")
            capture(page, frames, theme, args.fps, args.phase, args.scale)
            print(f"[{theme}] encoding…")
            gif = args.out / gif_name
            encode(frames, gif, args.fps, args.colors)
            shutil.copyfile(frames / "f_0000.png", args.out / png_name)
            print(f"[{theme}] {gif.relative_to(REPO)}  "
                  f"{gif.stat().st_size / 1024 / 1024:.2f} MB  {verify(gif, args.fps)}")
            print(f"[{theme}] {(args.out / png_name).relative_to(REPO)}  "
                  f"{(args.out / png_name).stat().st_size / 1024:.0f} KB")
    finally:
        if holder:
            holder.cleanup()
    return 0


if __name__ == "__main__":
    sys.exit(main())
