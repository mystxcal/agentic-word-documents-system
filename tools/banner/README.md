# Project banner

The animated banner at the top of the repository README. Source of truth is
`banner.template.html`; everything in `docs/assets/` is generated from it.

```powershell
python tools/banner/build_banner.py
```

That writes four files into `docs/assets/`:

| File | What it is |
|---|---|
| `banner.gif` | 2560×800, 12 s loop, light |
| `banner-dark.gif` | the same loop, dark |
| `banner.png` | the loop's first frame, light — a static poster |
| `banner-dark.png` | the loop's first frame, dark |

## What the picture says

The banner draws the compiler doing its job, and it stays honest to the
ownership rules the engine actually enforces.

- **Canonical sources.** Four cards, four owners, four different textures:
  Markdown+ prose, an arbitrary Excel table, a designed Word fragment, a
  reviewed figure.
- **Routing.** Each source runs a route to the one component it owns. The
  routes cross, because list order is not manifest order — the page still
  assembles top down.
- **Word-native assembly.** The page grows a cover band, a meta table, a TOC,
  body prose, a maintained table, and a figure. A coloured tick in the page
  margin marks which source owns each block, the way components are tagged in
  the compiled document.
- **Finalization.** A scan runs down the saved document and the TOC page
  references resolve from placeholders to results, because Word is the
  authority for pagination.
- **Verified output.** The PDF is exported, all twelve pages are rendered as
  images, the proof ledger closes, and the source hash settles.
- **Clear.** The bench wipes and the next build starts. The loop closes exactly.

## How it is built

`banner.template.html` is a single canvas scene with one rule: **every frame is
a pure function of loop time.** There is no accumulated state, no physics, no
frame-rate dependence. That buys three things — a rebuild is reproducible, any
frame can be inspected on its own, and the loop closes without a crossfade.

Two consequences worth knowing before editing it:

- The structure — ground, grid, cards, sheet, empty slots, labels, stage rail —
  is drawn identically on every frame; only the work moves. That is the right
  read (a bench that stays, a job that repeats) and it also lets the encoder
  difference most of the picture away, which is why a 2560×800 twelve-second
  loop lands near a megabyte.
- There are two palettes. `P` is the bench and inverts with the viewer's theme.
  `Q` is paper and does not, because a printed sheet is the same sheet under any
  light. Anything drawn onto a card, a page, or a page image reads from `Q`.

Layout lives in one `L` table and timing in one `B` table. Read `B` to know what
the banner is doing at any second.

`build_banner.py` inlines the vendored faces into a self-contained page, drives
`capture.js` over headless Chrome stepping `window.__step(t)`, then encodes with
a single global palette and `dither=none` — with a palette that already covers
flat art, dithering only adds noise that defeats frame differencing.

## Requirements

- Node, with a puppeteer package resolvable from this directory. `npm install
  puppeteer` is the simplest; with `puppeteer-core`, set `CHROME_PATH`. A module
  path in `BANNER_PUPPETEER` also works.
- `ffmpeg` and `gifsicle` on `PATH`.
- Pillow, for the loop-duration check (already a project dependency).

Useful switches: `--theme light`, `--scale 1` for a fast 1280×400 proof,
`--work build/banner` to keep the frames and intermediates for inspection.

The banner is a release asset. Nothing in the compiler imports this directory,
and a document build never runs it.

Typefaces are vendored and subset under `fonts/`; see `fonts/NOTICE.md`.
