# Releases and recovery

The system separates iteration, proof, publication, and recovery so a convenient folder is never mistaken for history.

## Draft builds

```powershell
.\Document-System.cmd build my-document --quick
.\Document-System.cmd build my-document
```

A quick build creates Word and PDF without rendering every page. A full build exports, renders, checks, compares, and records complete proof. Both create immutable run directories; only a full build should be treated as review-ready.

List recent runs:

```powershell
.\Document-System.cmd builds my-document --limit 10
.\Document-System.cmd history my-document --limit 20
```

## Current pair

`current/<document>/` contains the latest convenient DOCX/PDF pair and build report. It can be opened with:

```powershell
.\Document-System.cmd open my-document
.\Document-System.cmd open my-document --pdf
```

Audit detects a missing file, mismatched hash, stale extra file, source drift, or build inconsistency:

```powershell
.\Document-System.cmd audit my-document
```

Restore current from the recorded immutable build without losing displaced files:

```powershell
.\Document-System.cmd restore-current my-document
.\Document-System.cmd restore-current my-document --clean
```

`--clean` also moves stale current-folder files into restore history; it does not silently delete them.

## Controlled release

Profiles define required gates. Documents record each gate as `open`, `met`, or `not_applicable`. A release is refused when:

- a required gate remains open;
- the current full build did not pass quality proof;
- canonical inputs have changed since that build;
- artifact integrity does not match the recorded hashes.

After the responsible reviewer sets the gates appropriately:

```powershell
.\Document-System.cmd validate my-document
.\Document-System.cmd audit my-document
.\Document-System.cmd release my-document
```

The release contains a verified Word/PDF pair, release metadata, and a source lock tying the package to canonical input hashes. Releases are immutable.

## Publish or deliver

Publishing copies an already verified immutable pair to a chosen directory:

```powershell
.\Document-System.cmd publish my-document --to C:\path\delivery
.\Document-System.cmd deliver my-document C:\path\delivery
```

Collision policies are explicit:

- `fail` — stop if the destination names exist;
- `versioned` — preserve existing files and create a distinct versioned pair; this is the default;
- `replace` — replace destination files deliberately.

Use `--build-id` to publish an exact immutable run. `--allow-out-of-sync` is an explicit exception for sending the recorded build when current or canonical state differs; it should never be a routine shortcut.

## Retention

Retention is read-only by default:

```powershell
.\Document-System.cmd retention my-document --keep-drafts 10 --keep-previews 5
```

Review the plan, then apply it:

```powershell
.\Document-System.cmd retention my-document --keep-drafts 10 --keep-previews 5 --apply
```

Candidates move to a recoverable archive. Releases are not draft-retention candidates.

## Coworker Word workcopies

For a component whose canonical owner is a Word fragment:

```powershell
.\Document-System.cmd compare my-document C:\path\returned.docx --component worksheet
.\Document-System.cmd adopt my-document C:\path\returned.docx --component worksheet
```

Comparison uses the embedded build baseline. Adoption refuses a conflict when both the canonical component and the workcopy changed from the baseline. Resolve the content decision first; use `--accept-conflict` only to record an intentional choice.

Markdown-owned prose, Excel-owned tables, and drawing-owned diagrams are returned to their own canonical sources rather than adopted from compiled Word.

## Incident checklist

If an output looks wrong:

1. stop editing the compiled output;
2. run `audit` and record the build ID;
3. run `workspace` and locate the canonical owner;
4. compare the last good immutable build with the current run;
5. restore `current/` if only the convenience copy is damaged;
6. rebuild with `--retain-intermediates` if assembly or rendering needs diagnosis;
7. fix the canonical source or engine, then create a new build—never rewrite an old run.
