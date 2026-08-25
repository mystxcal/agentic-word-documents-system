# AI workflows

The system is designed so an agent can work quickly without turning the compiled Word file into an untraceable source.

## Give the agent a small operating contract

A useful starting instruction is:

> Treat the repository's canonical sources and manifests as authoritative. First run `workspace` and `status`. Edit only the source that owns the requested change. Do not invent missing content or hardcode a source-specific schema in the engine. Validate, build, inspect the final PDF pages, and report the exact artifact paths and remaining human decisions.

Adjust the “do not invent” language to the stakes of the document. The compiler controls mechanics; it cannot determine whether generated content is substantively correct.

## Recommended loop

```powershell
.\Document-System.cmd workspace my-document
.\Document-System.cmd status my-document
.\Document-System.cmd explain my-document --component body
.\Document-System.cmd edit my-document --component body --show
.\Document-System.cmd validate my-document
.\Document-System.cmd build my-document --quick
.\Document-System.cmd build my-document
.\Document-System.cmd audit my-document
```

An agent should use the quick build to catch structural mistakes, then inspect the full build's PDF and page images before claiming completion.

## Good agent tasks

- improve or reorganize a bounded Markdown section;
- add a maintained Excel table at an explicit slot;
- create a new neutral kit while leaving the old kit untouched;
- update document metadata in one manifest;
- compare a coworker Word workcopy to its canonical fragment;
- regenerate and visually inspect both sample documents;
- audit a proposed release for source drift and output integrity.

## Tasks that need human authority

- deciding whether a factual claim is approved;
- closing a review gate;
- choosing whether a source is licensed for redistribution;
- accepting a diagram rendition after visual review;
- resolving a real content conflict between current canonical source and a returned workcopy;
- publishing to an external destination when that action was not requested.

## Keep edits narrow

When asked to change only a cover, use the cover source. When asked to update prose, do not redesign the kit. When asked to change a table's styling, change its style role or kit—not its workbook columns. The layer model gives the agent a concrete way to control blast radius.

## Do not reverse-engineer Markdown from every Word edit

Markdown+ is intentionally one-way. If a coworker edited prose in the compiled Word output but the canonical prose is Markdown, compare the wording and apply the accepted change to Markdown manually. Automatic Word-to-Markdown would discard or guess at structure.

For Word-owned components, use the supported workcopy path:

```powershell
.\Document-System.cmd compare my-document C:\path\returned.docx --component decision-canvas
.\Document-System.cmd adopt my-document C:\path\returned.docx --component decision-canvas
```

Do not use `--accept-conflict` unless the human responsible for the content has chosen which side wins.

## Ask for machine-readable state

Every command accepts `--json` at the root level:

```powershell
.\Document-System.cmd --json workspace my-document
.\Document-System.cmd --json build my-document --quick
```

Agents should parse this output instead of scraping human console alignment. Human-readable output is a summary; JSON contains the full resolved paths, checks, diagnostics, and artifacts.

## Privacy checklist for a public template

Before publishing a repository or template, ask the agent to scan tracked files and binary metadata for:

- personal names and usernames;
- organization or client names;
- private locations, identifiers, reference numbers, and URLs;
- logos and brand marks;
- absolute local paths;
- document core properties and workbook creators;
- comments, tracked revisions, hidden sheets, custom XML, and embedded files;
- build reports or activity logs that contain local paths.

The shipped examples use fictional data, but new derivatives must be audited independently.
