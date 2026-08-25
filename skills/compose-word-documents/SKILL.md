---
name: compose-word-documents
description: Build, revise, preview, and verify native Word documents with the Agentic Word Documents system. Use when a task involves canonical Markdown+, Word donors, Excel tables, figures, PDF-page components, covers, headers, footers, page regions, DOCX/PDF output, or rapid document-layout iteration in a repository powered by agentic-doc.
---

# Compose Word Documents

Use the document engine as the editing and assembly boundary. Change the canonical owner, compile through `agentic-doc`, and judge the rendered result. Do not rewrite a DOCX ad hoc when the engine already owns that concern.

## Work in the smallest valid loop

1. Read the repository `AGENTS.md` and run `agentic-doc workspace DOCUMENT`.
2. Use `agentic-doc edit DOCUMENT --show`, optionally with `--component` or `--presentation`, to locate the canonical source.
3. Make the requested source change. Preserve unrelated components and Word-native structures.
4. Select the narrowest proof below.
5. Inspect every rendered page in that proof. Fix the canonical source or engine, then repeat.
6. Run a complete build only when the complete assembled document must be validated or delivered.

## Select the proof deliberately

- Header, footer, cover, margin, numbering, or page-region change: run `agentic-doc preview DOCUMENT --presentation page-furniture`. This creates a compact cover plus odd/even page proof without compiling the real body.
- One content component: run `agentic-doc preview DOCUMENT --component ID`. Add `--quick` only when heavy nested components should be represented by placeholders.
- Whole-document drafting with large PDF-page or other deferred components: run `agentic-doc build DOCUMENT --quick`. This still creates and renders Word/PDF, but it does not update `current` and cannot be published.
- Final validation: run `agentic-doc build DOCUMENT`. This compiles all components, embeds baselines, exports PDF, renders every page, and produces the only ordinary publishable draft.

Use `--include-heavy` with a quick proof only when the heavy material itself is under review. Never render hundreds of unchanged source pages merely to judge a header, footer, or nearby prose edit.

## Preserve the architecture

- Keep prose in Markdown+ where appropriate, native Word fragments in DOCX, arbitrary structured tables in Excel, and reviewed drawings/images in their declared sources.
- Edit header and footer donors independently. Word's different-first-page and odd/even switches are section-wide; use the page-furniture preview to prove all variants.
- Treat visible preview placeholders as deliberate omissions from a scoped proof. They are never final document content.
- Do not deliver or publish a build whose `content_scope` is not `complete` or whose `quality.release_ready` is false.
- If a requested construct is unsupported, extend the generic engine and its tests instead of creating a one-off document script.

## Verify the result

Confirm the build report states the intended `content_scope`, `quality.scope_passed`, and expected deferred components. For a complete build, also require `quality.release_ready`.

Inspect the actual page PNGs at original/full-page detail, not only the DOCX XML or a viewer's cropped high-detail tile. Check cover placement, header/footer centering, top clearance, page-number sequence, table continuation, image placement, unexpected blank pages, clipping, and the final page. For page-furniture previews, also require `page_furniture_visual.passed`. A scoped proof must be small enough that every page can be inspected immediately.

Read [references/commands.md](references/commands.md) for the command map and failure recovery.
