# Changelog

## 0.3.1 - 2026-08-19

- Fixed Windows frozen-GUI startup when bundled journal-profile files are unavailable in the PyInstaller extraction directory.
- Added an embedded profile fallback so the desktop app and CLI fail safely instead of crashing during profile discovery.
- Hardened release packaging to include bundled profiles explicitly.
- Added frozen-CLI profile smoke tests to the release workflow so this packaging regression is caught before assets are published.

## 0.3.0 - 2026-08-19

- Promoted the project to public beta after release, CI, and Pages workflows were validated.
- Added a combined `analyze` workflow that produces a single structure, integrity, citation, and optional journal-readiness report.
- Added local HTML report generation for researcher-friendly review and archival.
- Added bundled, source-dated journal profiles for PLOS ONE, Scientific Reports, and Frontiers in Microbiology Original Research.
- Added `profiles` and `validate-profile` CLI commands.
- Added profile freshness warnings and section-heading aliases for common declaration names.
- Upgraded the desktop GUI with bundled-profile selection, full-analysis mode, human-readable summaries, and HTML report export.
- Updated release automation to package Windows, macOS, and Linux builds with SHA-256 checksums and attach them directly to GitHub Releases.
- Updated the landing page with direct release downloads and transparent unsigned/notarized binary warnings.
- Expanded the synthetic regression suite to 11 passing tests.

## 0.2.1 - 2026-08-19

- Renamed the project to **Word Journal Manuscript Converter** so its purpose is explicit.
- Renamed the Python package to `word_journal_manuscript_converter`.
- Renamed user commands to `word-journal-converter` and `word-journal-converter-gui`.
- Updated the website, Word add-in starter, release workflow, documentation, tests, and citation metadata.

## 0.2.0

- Added manuscript structure extraction.
- Added citation/reference graph for numbered and conservative author-year patterns.
- Added safe profile-based retargeting for margins, line numbering, Normal style font, size, and line spacing.
- Added strict fail-closed preservation gate after transformation.
- Added plain numbered citation linking for simple `[N]` citations when no live citation-manager fields are present.
- Added local Tk desktop GUI.
- Added Word task-pane starter for local quick inspection.
- Expanded preservation checks for custom XML, tracked changes, comments, notes, equations, and tables.
- Added GitHub Pages landing page and release build workflows.

## 0.1.0

- Initial integrity-first DOCX inspector, verifier, journal-profile checker, tests, privacy policy, and architecture documentation.
