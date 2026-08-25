# Changelog

All notable changes are documented here. The format is based on Keep a Changelog and versions follow semantic versioning.

## [0.2.0] - 2026-08-25

### Added

- true lightweight builds with visible per-component include, placeholder, or omit policies;
- compact page-furniture proofs for cover, page-region, header, footer, margin, and numbering iteration;
- single-session Word finalization and PDF export for scoped previews;
- rendered-page ink analysis and unexpected blank-page detection;
- full-page top/bottom-band proof across consecutive page-furniture samples;
- explicit `content_scope`, scoped-proof, and release-readiness reporting;
- companion `compose-word-documents` Codex skill and focused command reference;
- animated repository banner in `docs/assets/`, light and dark, with its canonical source and rebuild tool in `tools/banner/`.

### Fixed

- identical default/even donors no longer activate Word's section-wide odd/even mode;
- real first/even variants inherit the other page-furniture kind's default, preventing disappearing headers or footers;
- Word PDF export restores the user's machine-level field-update preferences;
- section boundaries are always body-level Word paragraphs, never nested inside a cover/content control;
- cloned first/even header and footer stories receive independent package parts and drawing identities;
- Word repaginates before final PDF export so refreshed fields and section layout share one page model;
- scoped previews cannot update `current/` or pass the publishing gate.

## [0.1.0] - 2026-08-24

### Added

- composable kit, profile, project, and document manifests with generated schemas;
- Markdown+ prose compiler with semantic Word styles and named component slots;
- arbitrary Excel Table/range components;
- native Word fragments, figures, diagrams, explicit PDF pages, covers, TOCs, and page breaks;
- configurable page regions, headers, footers, numbering, core properties, and field bindings;
- immutable builds, current-pair recovery, publishing, release gates, source locks, retention, and activity history;
- component adapter cache with code-aware fingerprints and package verification;
- workcopy comparison and conflict-aware adoption for Word-owned components;
- native Word field finalization and deterministic PDF export behavior;
- PDF-level detection of broken Word bookmark and reference results;
- verified-copy build promotion fallback for protected Windows folders;
- two fictional, visually distinct example projects and reproducible binary asset generator;
- unit tests, runtime doctor, schema checks, build audits, and complete user documentation.
