# Word Journal Manuscript Converter

**Retarget the format. Preserve the science.**

[![CI](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/actions/workflows/ci.yml/badge.svg)](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/actions/workflows/ci.yml)
[![Release builds](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/actions/workflows/release.yml/badge.svg)](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/actions/workflows/release.yml)
[![Pages](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/actions/workflows/pages.yml/badge.svg)](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/actions/workflows/pages.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

Word Journal Manuscript Converter is a local-first toolkit for existing Microsoft Word research manuscripts. It audits preservation-sensitive DOCX structures, maps in-text citations to references, checks explicit journal requirements, applies a limited set of safe formatting changes, and verifies that protected scientific content survived unchanged.

> **Current status: v0.3.2 public beta.** The project is designed to be conservative. It does not claim perfect conversion for every journal or every Word file.

## Download

The easiest route is the prebuilt release package. No Python installation is required for the desktop build.

| Platform | Package | Includes |
|---|---|---|
| Windows x64 | [Download latest ZIP](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/releases/latest/download/Word-Journal-Manuscript-Converter-Windows-x64.zip) | Desktop GUI + CLI |
| macOS | [Download latest ZIP](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/releases/latest/download/Word-Journal-Manuscript-Converter-macOS.zip) | Desktop GUI + CLI |
| Linux x64 | [Download latest tar.gz](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/releases/latest/download/Word-Journal-Manuscript-Converter-Linux-x64.tar.gz) | CLI |

Each release asset has a matching `.sha256` checksum file. Current Windows and macOS binaries are unsigned/not notarized, so the operating system may show a security prompt.

**Website:** https://mbilal-ou.github.io/Word-Journal-Manuscript-Converter/

## Why this project exists

A research `.docx` is more than text plus formatting. It can contain live EndNote or Zotero fields, bookmarks, hyperlinks, equations, figures, comments, tracked changes, footnotes, endnotes, custom XML, and internal package relationships.

The converter follows three rules:

1. **Do not rewrite scientific content during formatting operations.**
2. **Do not flatten live citation-manager fields.**
3. **Do not keep a transformed file if the preservation audit fails.**

## Desktop workflow

Launch the desktop application and choose a manuscript. The current GUI supports:

- **Full analysis**: structure + integrity inventory + citation graph + optional journal readiness
- **Inspect**: preservation-sensitive DOCX inventory
- **Citation map**: in-text citation to bibliography matching
- **Readiness**: journal-profile checks with source/date provenance
- **Safe retarget**: controlled formatting changes followed by a fail-closed audit
- **Link `[N]` citations**: clickable internal links for simple plain numbered citations
- **Save HTML report**: local human-readable integrity/readiness report

The GUI ships with bundled profiles for:

- PLOS ONE, Research Article
- Scientific Reports, Article
- Frontiers in Microbiology, Original Research
- Generic review-copy demonstration profile

Journal profiles include an official source URL and checked date. Always confirm the current journal instructions before submission.

## CLI quick start

Inspect a manuscript:

```bash
word-journal-converter inspect manuscript.docx
```

Run a complete analysis and save an HTML report:

```bash
word-journal-converter analyze manuscript.docx \
  --profile scientific-reports-article \
  --json-out manuscript_report.json \
  --html-out manuscript_report.html
```

List bundled profiles:

```bash
word-journal-converter profiles
```

Check readiness:

```bash
word-journal-converter readiness manuscript.docx \
  --profile plos-one-research-article
```

Apply supported formatting to a copy:

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

## Citation integrity

Current citation graph support includes:

- plain numbered citations such as `[4]`, `[2, 5]`, and `[7-10]`
- conservative author-year matching
- inventory of EndNote, Zotero, CSL, and Mendeley field signatures
- unresolved citation reporting
- uncited-reference reporting
- internal Word links for simple plain numbered citations when no live citation-manager fields are present

Live citation-manager payloads are treated as protected content. The plain-text linking operation refuses to run when live citation-manager fields are detected.

## Preservation gate

The before/after verifier currently checks:

- visible text
- scientific numeric tokens
- complete Word field instructions
- live citation-field counts
- embedded-media SHA-256 hashes
- custom XML hashes
- tracked changes
- comments
- footnotes and endnotes
- equations
- tables
- bookmark loss
- hyperlink loss

A retargeted file is removed automatically if the preservation gate fails.

## Supported safe formatting changes

The current retargeting engine can change, when explicitly requested by a journal profile:

- page margins
- Word line numbering
- Normal-style body font
- body font size
- line spacing

The engine intentionally does not rewrite manuscript prose, renumber live citation-manager fields, or rebuild the entire DOCX from plain text.

## Bundled journal profiles

Use either a bundled key or your own JSON file.

```bash
word-journal-converter profiles
word-journal-converter validate-profile scientific-reports-article
word-journal-converter validate-profile path/to/custom-profile.json
```

Bundled profiles are in `src/word_journal_manuscript_converter/bundled_profiles/`. Editable examples also remain under `journal-profiles/`.

## Privacy

The core does not require network access for manuscript inspection, citation mapping, report generation, retargeting, or preservation verification.

```text
local manuscript
    -> local analysis
    -> transformed copy
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

## Microsoft Word add-in

A task-pane starter lives in [`integrations/word-addin/`](integrations/word-addin/). It performs lightweight in-document checks. Full package-level mutation stays in the local core because that is where the converter can inspect the complete DOCX package and enforce the preservation gate.

## Project structure

```text
src/word_journal_manuscript_converter/   core engine, GUI, reporting, bundled profiles
journal-profiles/                        editable profile templates
integrations/word-addin/                 Word task-pane starter
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
- silently flatten EndNote or Zotero fields
- submit manuscripts to publisher portals
- treat a dated local profile as more authoritative than current official journal instructions

## Documentation

- [User guide](docs/USER_GUIDE.md)
- [Architecture](docs/ARCHITECTURE.md)
- [Preservation contract](docs/PRESERVATION_CONTRACT.md)
- [Research safety](docs/RESEARCH_SAFETY.md)
- [Journal profiles](docs/JOURNAL_PROFILES.md)
- [Word add-in strategy](docs/WORD_ADDIN.md)
- [Roadmap](ROADMAP.md)
- [FAQ](docs/FAQ.md)

## License

MIT. See [LICENSE](LICENSE).
