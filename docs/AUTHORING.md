# Authoring

Use the simplest source that naturally owns the content. Markdown is excellent for prose, but it should not absorb every table, form, drawing, or inherited page.

## Markdown+

Markdown+ supports a deliberate subset that compiles cleanly into semantic Word structures.

### Headings and stable block IDs

```markdown
# Main section {#main-section}

## Supporting section
```

Heading levels map through the active kit's semantic styles. Explicit IDs are recommended for maintained headings because they give source blocks stable identities across wording changes. The TOC level is configured by its component.

### Paragraphs and inline formatting

```markdown
Use **bold**, *emphasis*, ~~strikeout~~, `inline code`, and
[safe links](https://example.com).
```

Supported links use `http`, `https`, or `mailto`. Raw HTML is rejected in strict mode.

### Lists

```markdown
1. First step
2. Second step
  - one nested point

- a bullet
- another bullet
```

Lists compile to native Word numbering. Separate ordered lists restart rather than inheriting numbering accidentally.

### Code blocks

````markdown
```text
One line of literal or command-style text.
```
````

The kit decides the code/command style. The source does not carry font sizes or colors.

### Callouts

```markdown
:::note
A restrained note using the kit's semantic callout style.
:::

:::warning
Something the reader should not miss.
:::

:::callout role=important
A custom semantic role when the kit defines one.
:::
```

Supported directives are `note`, `warning`, `important`, and `callout`. Unknown directives fail in strict mode instead of appearing as unexplained text.

### Markdown tables

```markdown
| Item | Owner | Due |
|:---|:---|---:|
| Draft | Mina | 14 |
```

Use a Markdown table for small prose-adjacent data. Use Excel when the table is maintained independently, has many rows, or needs business users to edit it without touching prose.

### Component insertion

```markdown
:::insert observation-table
```

The slot name must match the manifest and must occur exactly once. Images written as Markdown image syntax are rejected; register a figure component instead so width, caption, alt text, review status, and provenance remain explicit.

### Page breaks

```markdown
:::page-break
```

Use explicit page breaks sparingly. Normal flow, keep-with-next styles, and table row rules should solve most pagination.

## Excel tables

An Excel component can select:

- a named Excel Table, recommended for maintained data; or
- an explicit worksheet range.

Columns are not hardcoded. The workbook defines the data shape; the manifest can select, rename, order, size, and align columns through a view when needed. A style role such as `technical` or `schedule` controls visual treatment.

Example locator:

```jsonc
"options": {
  "locator": {"table": "ObservationsData"},
  "view": {"columns": "*", "style_role": "technical"},
  "formula_policy": "require_no_formulas"
}
```

Formula policy is explicit because a headless reader may see formulas without reliable current results. Prefer maintained values for publication tables, or configure and test a formula policy suited to the workflow.

## Word fragments

Use a Word fragment when Word itself is the natural editor: a form, worksheet, designed matrix, cover, legal block, or layout-heavy page.

A controlled fragment usually contains a tagged content control. The manifest identifies that tag through `source_tag`. This lets the compiler import exactly the intended block without inheriting unrelated margins, empty paragraphs, or document-level settings.

Use `allow_untagged` only as an explicit decision. It imports the whole Word body and is more vulnerable to accidental surrounding content.

For an existing Word source that must retain its sections, headers, footers, or pagination, create the document with the preservation options rather than converting it to Markdown:

```powershell
.\Document-System.cmd new document inherited-guide `
  --project my-project `
  --title 'Inherited Guide' `
  --word-source C:\path\source.docx `
  --allow-untagged `
  --preserve-sections `
  --preserve-source-layout `
  --use-source-styles
```

## Figures

A figure component owns publication behavior around a reviewed raster image:

- source path;
- width;
- alignment;
- caption;
- alt text.

Replacing the image does not require rewriting the prose or recreating its paragraph treatment. Use an appropriate export resolution; the compiler can normalize media but cannot restore detail absent from the source.

## Diagrams

A diagram has two identities:

1. the native editable drawing;
2. the reviewed rendition inserted into Word.

The accepted rendition is bound to the current native-source hash. If the native drawing changes, the old rendition is no longer implicitly trusted. Export and review the new image, then register it:

```powershell
.\Document-System.cmd accept-rendition my-document `
  --component architecture-diagram `
  --file C:\path\architecture-reviewed.png
```

## PDF pages

PDF-page components insert explicit, one-based pages from an original PDF. This is useful for pages that must remain visually exact and are legitimately part of the document. Selection is explicit; the engine does not guess “relevant pages.”

In `build --quick`, PDF-page components become visible placeholders by default. Use `--include-heavy` only when those pages themselves must be reviewed. A complete build always inserts the selected pages.

Use the source inspection command before registering a file:

```powershell
.\Document-System.cmd inspect-source C:\path\reference.pdf
```

## Presentation editing

Presentation parts are independently editable:

```powershell
.\Document-System.cmd edit my-document --presentation cover
.\Document-System.cmd edit my-document --presentation header
.\Document-System.cmd edit my-document --presentation footer
.\Document-System.cmd edit my-document --presentation styles
.\Document-System.cmd edit my-document --presentation shell
```

Use `--show` when you only need the path. A kit-level donor affects every document that selects it; a document-level override affects only that document.

After changing a cover, header, footer, style donor, margin, or page numbering, run:

```powershell
.\Document-System.cmd preview my-document --presentation page-furniture
```

This proof uses real presentation sources with compact synthetic content. Inspect all of its rendered pages before running a complete document build.

## Before a full build

Run:

```powershell
.\Document-System.cmd validate my-document
.\Document-System.cmd status my-document
.\Document-System.cmd build my-document --quick
```

The quick build renders a lightweight Word/PDF proof, does not update `current/`, and cannot be published. Then use a full build for review. The full path is where complete content, PDF text, every page render, metadata, pagination, blank-page detection, and output comparisons are proven.
