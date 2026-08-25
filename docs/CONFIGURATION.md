# Configuration

Configuration uses commented JSON (`.jsonc`) validated against generated schemas. Editors can provide completion and diagnostics through each file's `$schema` field.

Refresh or check schemas with:

```powershell
.\Document-System.cmd schema
.\Document-System.cmd schema --check
```

## Kit manifest

A kit defines reusable visual vocabulary, not document content.

```jsonc
{
  "$schema": "../../schemas/agentic-kit-v2.schema.json",
  "schema": "agentic-kit/v2",
  "id": "studio",
  "components": {
    "studio": "donors/Style_Gallery.docx",
    "running-header": "donors/Running_Header.docx",
    "folio-footer": "donors/Folio_Footer.docx"
  },
  "semantic_styles": {
    "heading_1": "Heading 1",
    "body": "Normal",
    "bullet": "List Bullet",
    "numbered_step": "List Number",
    "command": "Studio Code",
    "note": "Studio Callout",
    "caption": "Caption"
  },
  "table_styles": {
    "technical": {
      "font_name": "Aptos",
      "font_size_pt": 8.8,
      "header_fill": "17324D",
      "header_text": "FFFFFF",
      "body_fill": "FFFFFF",
      "alternate_fill": "EDF4F3",
      "border_color": "B8C8CC",
      "text_color": "18242B",
      "repeat_header": true,
      "allow_row_split": false
    }
  }
}
```

Table roles style whatever columns the source supplies. They do not define domain-specific column names.

## Profile manifest

A profile defines reusable document shape:

```jsonc
{
  "$schema": "../../schemas/agentic-profile-v2.schema.json",
  "schema": "agentic-profile/v2",
  "id": "report",
  "shell": "shell.docx",
  "body_slot": "AGDOC.BODY.REPORT",
  "region_starts": {
    "main": {"tag": "AGDOC.LAYOUT.MAIN_START", "boundary": "next_page"}
  },
  "field_bindings": {
    "AGDOC.FIELD.COVER.TITLE": {"path": "metadata.title", "required": true},
    "AGDOC.FIELD.HEADER.TITLE": {"path": "metadata.short_title", "required": true}
  },
  "release_gates": ["content-reviewed", "visual-proof-reviewed"]
}
```

Bindings populate tagged content controls from resolved project/document data. A binding can be required, transformed, or date-formatted. Keep profile bindings generic enough for every document selecting the profile.

## Project manifest

Projects are open collections, not rigid customer records:

```jsonc
{
  "$schema": "../../schemas/agentic-project-v2.schema.json",
  "schema": "agentic-project/v2",
  "id": "learning-notes",
  "name": "Learning Notes",
  "description": "A maintained collection of working guides.",
  "metadata": {
    "audience": "new facilitators",
    "language": "English"
  },
  "sources": {}
}
```

`metadata` is intentionally open. Put shared values there only when more than one document genuinely uses them.

## Document manifest

The document manifest selects the layers and owns assembly.

```jsonc
{
  "$schema": "../../../../schemas/agentic-document-v2.schema.json",
  "schema": "agentic-document/v2",
  "id": "field-journal",
  "project": "learning-notes",
  "profile": "report",
  "kit": "studio",
  "metadata": {
    "type": "Journal",
    "title": "Field Journal",
    "short_title": "Field Journal",
    "revision": "draft-01",
    "date": "2026-08-24",
    "author": "Example Author",
    "description": "A concise maintained journal."
  },
  "sequence": [
    {"region": "front", "items": ["cover"]},
    {"region": "main", "items": ["toc", "body"]}
  ],
  "components": {
    "cover": {
      "type": "cover",
      "ownership": "word_fragment",
      "source": "presentation/cover.docx",
      "source_tag": "AGDOC.COVER"
    },
    "toc": {"type": "toc", "toc_levels": 2},
    "body": {
      "type": "document",
      "ownership": "source",
      "source": "content/body.md",
      "preview": {"mode": "include"}
    }
  },
  "quality": {"allow_blank_pages": false},
  "outputs": {"basename": "Field Journal"},
  "release": {
    "gates": {
      "content-reviewed": "open",
      "visual-proof-reviewed": "open"
    }
  }
}
```

## Presentation and page regions

Presentation selects reusable donors and page behavior independently:

```jsonc
"presentation": {
  "styles": "kit:studio",
  "cover": "document:cover",
  "page_regions": {
    "front": {
      "header": "none",
      "footer": "none",
      "numbering": null,
      "top_margin_twips": 720,
      "bottom_margin_twips": 720,
      "left_margin_twips": 900,
      "right_margin_twips": 900
    },
    "main": {
      "header": "kit:running-header",
      "footer": "kit:folio-footer",
      "numbering": {"style": "arabic", "start": 1, "page_count_scope": "region"},
      "top_margin_twips": 1220,
      "bottom_margin_twips": 1080,
      "left_margin_twips": 1200,
      "right_margin_twips": 1050
    }
  }
}
```

Margins use twips (`1 inch = 1440 twips`). Regions are semantic: `front`, `main`, or another profile-defined region. A region can deliberately have no header, footer, or numbering.

A header or footer may also select `default`, `first`, and `even` donors. Use variants only when their content is genuinely different. Word enables first-page and odd/even behavior for the whole section, so the compiler deduplicates identical donors and carries the other kind's default into a real variant when required.

## Preview and quality policies

Each component may declare a lightweight preview policy:

```jsonc
"preview": {
  "mode": "placeholder",
  "label": "Reference publication pages"
}
```

Modes are `include`, `placeholder`, and `omit`. PDF-page components default to `placeholder`; other types default to `include`. This setting affects only quick/scoped proofs and never changes a complete build.

`quality.allow_blank_pages` defaults to `false`. Set it to `true` only when blank pages are intentional document design, because a blank page commonly indicates a section-boundary or odd/even furniture error.

## References

Presentation references use scoped names:

- `kit:name` — reusable component declared by the selected kit;
- `document:id` — a component declared in this document;
- `none` — explicitly no component.

This prevents a bare path from silently changing meaning when files move.

## Release gates

Profiles define the gate names; documents record `open`, `met`, or `not_applicable`. A successful build proves mechanics but does not close a human review gate. Change a gate to `met` only after the named review has actually happened, or to `not_applicable` only when that decision is defensible, then run:

```powershell
.\Document-System.cmd validate my-document
.\Document-System.cmd release my-document
```

The schema is the authoritative list of allowed states and fields for the installed version. Do not invent a new gate or metadata contract merely because another project used one.

## Prefer commands for routine changes

Use `new`, `add`, `accept-rendition`, `replace`, `revise`, and `adopt` when they fit. Edit JSONC directly when changing composition or presentation semantics. Always validate after a manual manifest change.
