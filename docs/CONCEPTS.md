# Concepts

The system is easiest to use when four ideas stay separate: source ownership, composition, presentation, and proof.

## Canonical source

A canonical source is the file that owns future changes to one piece of the document.

- If prose is owned by Markdown, edit the `.md` file.
- If a table is owned by Excel, edit the workbook.
- If a worksheet is owned by Word, edit its fragment `.docx`.
- If a diagram is owned by a native drawing file, edit that file and accept a new reviewed rendition.

The compiled document is not automatically canonical. This prevents an agent, a coworker, and a build script from each maintaining a different version of the same content.

## Component

A component is a typed unit that the compiler knows how to place and verify. Supported roles include:

- Markdown+ prose;
- native Word fragment;
- Excel table or explicit range;
- figure;
- diagram with native source and accepted rendition;
- explicit PDF pages;
- table of contents;
- cover;
- page break.

A component declares its source, ownership, options, and any child slots. The compiler does not guess that every spreadsheet is the same kind of table or that every image is a diagram.

## Slot

A slot is a named insertion point. In Markdown+, it looks like this:

```markdown
The maintained schedule follows.

:::insert schedule

The prose continues here.
```

The document manifest declares which component may occupy `schedule`. The slot must appear exactly once, so content cannot quietly duplicate or vanish.

## Composition layers

The layers answer different questions:

| Layer | Question it answers | Typical contents |
|---|---|---|
| kit | What is the visual language? | semantic styles, reusable header/footer donors, table themes |
| profile | What shape does this kind of document have? | shell, body slot, page-region boundaries, metadata bindings, release gates |
| project | What context is shared by a collection? | name, description, open metadata, shared source aliases |
| document | What exactly are we building? | metadata, sequence, components, presentation selections, output name |

Specificity belongs as low as necessary. A header used by many documents belongs in a kit. A cover unique to one document belongs beside that document. A workbook column list belongs in the workbook, not in the engine.

## Word-native output

“Word-native” means the result is assembled from real Word structures rather than a picture of a document:

- headings use semantic Word styles;
- TOCs are native fields;
- lists use native numbering definitions;
- tables remain tables;
- page regions become Word sections;
- headers and footers remain editable stories;
- page numbers and totals remain fields;
- Word fragments remain editable content.

The compiler may use content controls to track ownership and components. Those controls are infrastructure around editable Word content, not a flattened rendering layer.

## Draft, current, and release

- A **build run** is immutable and includes its proof.
- **Current** is a convenient copy of the latest verified Word/PDF pair.
- A **release** is a controlled package created only when configured gates are closed.
- A **publication** copies a verified immutable pair to a destination using an explicit collision policy.

This distinction means a partially edited file in `current/` cannot silently rewrite build history.

## Markdown+ is deliberately one-way

Markdown+ compiles into Word; arbitrary Word edits are not converted back into Markdown. Bidirectional conversion would have to guess how Word layout, fields, floating objects, and content controls map to source text. The system instead makes the canonical source easy to find and lets genuinely Word-owned material remain Word-owned.

## What the system does not decide

The compiler proves document mechanics. It does not decide whether a claim is true, whether a policy is approved, or whether the content is safe to publish. Release gates represent human decisions; they are not replaced by a green build.
