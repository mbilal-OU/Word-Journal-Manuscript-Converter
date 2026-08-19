# Word Journal Manuscript Converter

**Retarget the format. Preserve the science. Make citations traceable.**

[![CI](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/actions/workflows/ci.yml/badge.svg)](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/actions/workflows/ci.yml)
[![Release builds](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/actions/workflows/release.yml/badge.svg)](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/actions/workflows/release.yml)
[![Pages](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/actions/workflows/pages.yml/badge.svg)](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/actions/workflows/pages.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

Word Journal Manuscript Converter is a local-first toolkit for existing Microsoft Word research manuscripts. It now has three independent workflows: journal conversion, citation/reference navigation, and manuscript auditing.

> **Current status: v0.4.0 public beta candidate.** The project is intentionally conservative. It does not claim perfect conversion for every journal or every Word file.

## Download

The easiest route is the prebuilt release package. No Python installation is required for the desktop build.

| Platform | Package | Includes |
|---|---|---|
| Windows x64 | [Download latest ZIP](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/releases/latest/download/Word-Journal-Manuscript-Converter-Windows-x64.zip) | Desktop GUI + CLI |
| macOS | [Download latest ZIP](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/releases/latest/download/Word-Journal-Manuscript-Converter-macOS.zip) | Desktop GUI + CLI |
| Linux x64 | [Download latest tar.gz](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/releases/latest/download/Word-Journal-Manuscript-Converter-Linux-x64.tar.gz) | CLI |

Each release asset has a matching `.sha256` checksum file. Current Windows and macOS binaries are unsigned/not notarized, so the operating system may show a security prompt.

**Website:** https://mbilal-ou.github.io/Word-Journal-Manuscript-Converter/

## Three workflows

### 1. Journal Conversion

Use a source-dated journal profile to check explicit requirements and apply only supported formatting changes to a copy. The transformed DOCX is kept only if the preservation gate passes.

### 2. Citation Navigator

Use this when you do **not** need journal conversion and only want citations and references to be easier to trace.

- Detects EndNote, Zotero, CSL, and Mendeley field signatures.
- Maps numbered citations to bibliography entries.
- Generates a local clickable HTML citation map.
- Creates a separate internally linked DOCX for safely linkable plain numbered citations.
- For live EndNote/Zotero/Mendeley documents, it does not wrap or rewrite citation fields.
- The Word add-in provides live-safe click-to-jump navigation by moving the Word selection between citation and reference locations.

This means a thesis, dissertation, assignment, review article, or manuscript can use Citation Navigator without selecting any journal.

### 3. Manuscript Audit

Run structure, citation, figure, table, equation, field, comment, tracked-change, and other preservation-sensitive checks without retargeting the document.

## Why only a few bundled journals?

Bundled journal profiles are intentionally source-dated and manually verified. The project avoids shipping hundreds of guessed or stale journal rules. You can already use a custom JSON profile for any journal, while verified built-in coverage grows separately.

Current bundled profiles include:

- PLOS ONE, Research Article
- Scientific Reports, Article
- Frontiers in Microbiology, Original Research
- Generic review-copy demonstration profile

Journal profiles include an official source URL and checked date. Always confirm the current journal instructions before submission.

## Citation safety model

A research `.docx` can contain live citation-manager fields, bookmarks, hyperlinks, equations, figures, comments, tracked changes, footnotes, endnotes, custom XML, and internal package relationships.

The converter follows four rules:

1. **Do not rewrite scientific content during formatting operations.**
2. **Do not flatten or rewrite live citation-manager fields to force hyperlinks.**
3. **Use non-mutating Word navigation for live EndNote/Zotero/Mendeley documents.**
4. **Do not keep a transformed file if the preservation audit fails.**

## Desktop app

The GUI is organized around the three workflows above.

**Journal Conversion** includes journal analysis, readiness checks, and safe retargeting.

**Citation Navigator** includes citation mapping, navigation analysis, local HTML navigation reports, and safe clickable-copy export for plain numbered citations.

**Manuscript Audit** provides a no-journal integrity and structure review.

## CLI quick start

Citation navigation with no journal:

```bash
word-journal-converter navigate manuscript.docx \
  --html-out citation_navigation.html
```

Create a clickable copy when the document contains safely linkable plain numbered citations:

```bash
word-journal-converter make-navigable manuscript.docx \
  --output manuscript_navigable.docx
```

For a live EndNote/Zotero/Mendeley manuscript, `make-navigable` refuses to rewrite the citation fields and directs the user to live-safe Word add-in navigation.

Run a full manuscript audit without a journal profile:

```bash
word-journal-converter analyze manuscript.docx \
  --html-out manuscript_report.html
```

Run journal-specific analysis:

```bash
word-journal-converter analyze manuscript.docx \
  --profile scientific-reports-article \
  --html-out manuscript_report.html
```

List bundled profiles:

```bash
word-journal-converter profiles
```

Apply supported journal formatting to a copy:

```bash
word-journal-converter retarget manuscript.docx \
  --profile frontiers-microbiology-original-research \
  --output manuscript_retargeted.docx \
  --report retarget_report.json
```

Verify any before/after pair:

```bash
word-journal-converter verify manuscript.docx manuscript_retargeted.docx
```

## Word add-in Citation Navigator

The task-pane companion in [`integrations/word-addin/`](integrations/word-addin/) now scans the open Word document, builds a numbered citation/reference map, and provides buttons to jump to the first citation occurrence or its matching bibliography entry.

For live citation-manager documents, this is the preferred navigation method because it changes only the Word selection. It does not edit citation-manager field payloads.

## Preservation gate

The before/after verifier currently checks:

- visible text
- scientific numeric tokens
- complete Word field instructions
- live citation-field counts
- embedded-media SHA-256 hashes
- custom XML hashes
- package relationships
- content-type declarations
- tracked changes
- comments
- footnotes and endnotes
- equations
- tables
- bookmark loss
- hyperlink loss

A retargeted or linked file is removed automatically if the preservation gate fails.

## Supported safe formatting changes

The retargeting engine can change, when explicitly requested by a journal profile:

- page margins
- Word line numbering
- Normal-style body font
- body font size
- line spacing

The engine intentionally does not rewrite manuscript prose, renumber live citation-manager fields, or rebuild the entire DOCX from plain text.

## Custom journal profiles

Use either a bundled key or your own JSON file.

```bash
word-journal-converter profiles
word-journal-converter validate-profile scientific-reports-article
word-journal-converter validate-profile path/to/custom-profile.json
```

Bundled profiles are in `src/word_journal_manuscript_converter/bundled_profiles/`. Editable examples also remain under `journal-profiles/`.

## Privacy

The core does not require network access for manuscript inspection, citation navigation analysis, report generation, retargeting, or preservation verification.

```text
local manuscript
    -> local analysis
    -> optional transformed copy
    -> local preservation audit
    -> user review
```

See [PRIVACY.md](PRIVACY.md) and [Research safety](docs/RESEARCH_SAFETY.md).

## Install from source

```bash
git clone https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter.git
cd Word-Journal-Manuscript-Converter
python -m venv .venv
```

Activate the environment, then:

```bash
pip install -e .
word-journal-converter --help
```

The core package uses the Python standard library only.

## Tests

```bash
pytest
```

Regression tests use synthetic DOCX packages so unpublished research is never needed in the repository test corpus.

## Project structure

```text
src/word_journal_manuscript_converter/   core engine, GUI, reports, Citation Navigator, profiles
journal-profiles/                        editable journal profile templates
integrations/word-addin/                 live-safe Word Citation Navigator
packaging/                               executable build entry points
site/                                    GitHub Pages landing page
tests/                                   synthetic OOXML regression tests
docs/                                    architecture, safety, user guide
.github/workflows/                       CI, release assets, Pages deployment
```

## Non-goals

Word Journal Manuscript Converter does not:

- guarantee journal acceptance
- fabricate or invent references
- rewrite scientific claims during formatting
- silently flatten EndNote, Zotero, or Mendeley fields
- force hyperlinks into live citation-manager field payloads
- submit manuscripts to publisher portals
- treat a dated local profile as more authoritative than current official journal instructions

## Documentation

- [User guide](docs/USER_GUIDE.md)
- [Citation Navigator](docs/CITATION_NAVIGATOR.md)
- [Architecture](docs/ARCHITECTURE.md)
- [Preservation contract](docs/PRESERVATION_CONTRACT.md)
- [Research safety](docs/RESEARCH_SAFETY.md)
- [Journal profiles](docs/JOURNAL_PROFILES.md)
- [Word add-in strategy](docs/WORD_ADDIN.md)
- [Roadmap](ROADMAP.md)
- [FAQ](docs/FAQ.md)

## License

MIT. See [LICENSE](LICENSE).
