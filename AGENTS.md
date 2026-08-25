# Agent operating rules

- Begin with `workspace`, `status`, and `explain`; edit the canonical source that owns the requested change.
- Keep the repository generic. Never add real client, employer, personal, branded, confidential, or locally identifying examples.
- Do not treat compiled Word output as canonical for Markdown-, Excel-, image-, PDF-, or drawing-owned content.
- Preserve Word-native headings, TOCs, numbering, tables, sections, headers, footers, fields, and editable fragments.
- Do not hardcode business-specific workbook columns, document sections, identifiers, or metadata into the engine.
- Put reusable visual rules in kits, structural contracts in profiles, shared context in projects, and one-document choices in document manifests.
- Do not invent substantive content, close review gates, accept renditions, resolve conflicts, or publish externally without appropriate user authority.
- After source changes, validate and run relevant tests. After layout or compiler changes, fully build both samples and inspect the rendered PDFs/pages.
- A green unit suite is not visual proof. A generated DOCX without successful field finalization, PDF export, page rendering, and field-error checks is not a fully proven build.
- Keep runtime outputs and local paths out of Git. Run a tracked-file privacy scan before any public release.
