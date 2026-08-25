# Security and privacy

## Supported version

Security and privacy fixes target the latest version on the default branch.

## Reporting a vulnerability

Do not publish credentials, private documents, or exploit details in a public issue. Use the repository host's private vulnerability-reporting feature when available, or contact the maintainer through a private channel listed by that host.

## Local execution model

The system runs locally and invokes:

- Python libraries for package assembly and inspection;
- Microsoft Word through Windows COM for native field/layout finalization and PDF export;
- Poppler for PDF metadata and page rendering.

It does not require a hosted service. Canonical documents may nevertheless contain sensitive information, so users must control repository visibility, build locations, backups, and publication destinations.

## Untrusted inputs

Treat unfamiliar Word, Excel, PDF, image, and diagram files as untrusted. Inspect them with standard endpoint protections before opening them in desktop applications. The compiler validates package structure and relationships for correctness; it is not a malware sandbox.

Do not enable macros or external links merely to make a build complete. The system uses `.docx` and `.xlsx` sources and does not require macro-enabled formats.

## Public-repository hygiene

Runtime build reports contain absolute local paths and should remain ignored. Before publishing a template, inspect both visible content and embedded metadata for names, usernames, organizations, locations, identifiers, comments, revisions, custom XML, hidden workbook content, and embedded files.

The examples in this repository are fictional and neutral. That does not make derivatives automatically safe to publish.
