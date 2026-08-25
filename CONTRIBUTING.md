# Contributing

Contributions should preserve the central promise: simple canonical sources, Word-native output, and honest proof.

## Set up

```powershell
powershell -ExecutionPolicy Bypass -File .\setup.ps1
$env:PYTHONPATH = (Join-Path (Get-Location) 'src')
python -m unittest discover -s tests -p 'test_*.py'
```

## Change discipline

- Keep domain-specific facts and column schemas out of the compiler core.
- Add configuration only when it exposes a meaningful user lever.
- Prefer typed component adapters over special cases in document assembly.
- Preserve native Word structures and package relationships.
- Treat Markdown as one-way unless a separately designed, loss-aware import path exists.
- Include adapter implementation code in cache invalidation.
- Do not weaken proof checks to make a failing sample look green.
- Do not commit runtime directories, private source documents, local paths, or real organization data.

## Tests

For source changes:

```powershell
$env:PYTHONPATH = (Join-Path (Get-Location) 'src')
python -m unittest discover -s tests -p 'test_*.py'
```

For changes affecting Word output, also run:

```powershell
.\Document-System.cmd schema --check
.\Document-System.cmd doctor
.\Document-System.cmd build shade-study
.\Document-System.cmd build clear-decisions
.\Document-System.cmd audit shade-study
.\Document-System.cmd audit clear-decisions
```

Inspect the actual rendered page PNGs and PDFs. A passing unit suite does not prove typography, page flow, table density, or visual continuity.

## Adding a template

A public template must be fictional or intentionally licensed for redistribution. Include:

- canonical source files;
- a project and document manifest;
- any new reusable kit/profile only when needed;
- a rendered reference PDF if it materially helps users evaluate the template;
- a leakage scan covering text and binary metadata;
- build and visual proof.

Avoid adding a third template that merely recolors an existing one. New examples should demonstrate a genuinely different source mix or document shape.

## Pull request notes

Describe:

- the user problem;
- the ownership or configuration decision;
- the files and layers changed;
- unit-test result;
- sample build IDs and quality status;
- pages visually inspected;
- any compatibility or migration effect.
