# Quickstart

This walkthrough gets you from a clean checkout to an edited, fully proven Word/PDF build.

## 1. Check the prerequisites

You need:

- Windows 10 or 11;
- Microsoft Word desktop with COM automation available;
- Python 3.12 or newer;
- Poppler tools `pdftoppm.exe` and `pdfinfo.exe`.

If Poppler is not on `PATH`, set its binary directory for the current shell:

```powershell
$env:AGENTIC_DOCS_POPPLER_BIN = 'C:\path\to\poppler\Library\bin'
```

## 2. Install the local command

From the repository root:

```powershell
powershell -ExecutionPolicy Bypass -File .\setup.ps1
```

The installer creates `.venv`, installs the package in editable mode, and runs the runtime doctor. It does not alter Word, install Office, or change machine-wide settings.

You can rerun the check at any time:

```powershell
.\Document-System.cmd doctor
```

Do not continue to a release build until the result says `READY`.

## 3. Build a sample

```powershell
.\Document-System.cmd documents
.\Document-System.cmd workspace shade-study
.\Document-System.cmd build shade-study
.\Document-System.cmd open shade-study --pdf
```

The build creates:

- an immutable run under `builds/shade-study/<build-id>/`;
- an editable Word document;
- a matching PDF;
- one PNG per PDF page;
- `build-report.json` with input hashes, timings, checks, diagnostics, and change comparison;
- a convenient latest pair under `current/shade-study/`.

## 4. Make a real source edit

Open the canonical prose directly:

```powershell
.\Document-System.cmd edit shade-study
```

Change one sentence in `projects/field-study/documents/shade-study/content/body.md`, save it, then iterate:

```powershell
.\Document-System.cmd build shade-study --quick
```

The quick build is an immutable scoped proof. It renders its own small PDF, does not replace `current/`, and cannot be published as a complete document. Its console output links the exact Word/PDF artifacts.

For a cover, header, footer, margin, or page-numbering change, use the smaller presentation proof:

```powershell
.\Document-System.cmd preview shade-study --presentation page-furniture
```

It builds a compact cover plus odd/even page fixture. Inspect every rendered page before continuing.

When the content is ready for review, run the full proof:

```powershell
.\Document-System.cmd build shade-study
.\Document-System.cmd audit shade-study
```

Only a full build has `content_scope: complete` and can become release-ready.

## 5. Start your own document

```powershell
.\Document-System.cmd new project my-notes `
  --name 'My Notes' `
  --description 'A collection of maintained documents'

.\Document-System.cmd new document my-guide `
  --project my-notes `
  --title 'My Working Guide' `
  --type 'Guide' `
  --profile report `
  --kit studio `
  --date 2026-08-24
```

The command creates a Markdown-first starter. Use `workspace my-guide` immediately; it shows the source to edit, the active design layers, the latest outputs, and the next commands.

## 6. Add maintained content without hand-editing JSONC

The `add` command registers and places typed components. Ask for exact options with a subcommand help, for example:

```powershell
.\Document-System.cmd add my-guide table --help
.\Document-System.cmd add my-guide figure --help
.\Document-System.cmd add my-guide word-fragment --help
```

The usual pattern is:

1. place `:::insert component-name` at the desired point in Markdown;
2. add the component with the same slot name;
3. validate;
4. build.

## 7. Know which build to trust

`current/` is convenient; `builds/<document>/<build-id>/` is authoritative. If `current/` is manually changed or partially replaced, the audit reports it and `restore-current` can reconstruct the pair from the immutable build:

```powershell
.\Document-System.cmd audit my-guide
.\Document-System.cmd restore-current my-guide --clean
```

## If something fails

Run these in order:

```powershell
.\Document-System.cmd doctor
.\Document-System.cmd validate my-guide
.\Document-System.cmd status my-guide
.\Document-System.cmd build my-guide --retain-intermediates
```

The retained raw DOCX and adapter work files are diagnostic artifacts, not new canonical sources. See [docs/RELEASES_AND_RECOVERY.md](docs/RELEASES_AND_RECOVERY.md) for recovery and [docs/AUTHORING.md](docs/AUTHORING.md) for source-specific problems.
