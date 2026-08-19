# Word Journal Manuscript Converter

**Retarget the format. Preserve the science. Make citations traceable.**

[![CI](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/actions/workflows/ci.yml/badge.svg)](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/actions/workflows/ci.yml)
[![Release builds](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/actions/workflows/release.yml/badge.svg)](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/actions/workflows/release.yml)
[![Pages](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/actions/workflows/pages.yml/badge.svg)](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/actions/workflows/pages.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

Word Journal Manuscript Converter is a local-first toolkit for existing Microsoft Word research manuscripts. It combines journal readiness checks, safe formatting retargeting, Word-template adaptation, citation/reference navigation, and manuscript integrity auditing.

> **Current status: v0.5.0 public beta candidate.** The project is deliberately conservative. It does not claim perfect conversion for every journal or every Word file.

## Download

Prebuilt release packages require no Python installation.

| Platform | Package | Includes |
|---|---|---|
| Windows x64 | [Download latest ZIP](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/releases/latest/download/Word-Journal-Manuscript-Converter-Windows-x64.zip) | Desktop GUI + CLI |
| macOS | [Download latest ZIP](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/releases/latest/download/Word-Journal-Manuscript-Converter-macOS.zip) | Desktop GUI + CLI |
| Linux x64 | [Download latest tar.gz](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/releases/latest/download/Word-Journal-Manuscript-Converter-Linux-x64.tar.gz) | CLI |

Each release asset has a matching `.sha256` checksum. Windows and macOS binaries are currently unsigned/not notarized, so the operating system may show a security prompt.

**Website:** https://mbilal-ou.github.io/Word-Journal-Manuscript-Converter/

## Three workflows

### 1. Journal Conversion

Choose one of the source-dated built-in journal profiles, use a custom JSON profile, or supply a journal Word template.

The v0.5 catalog contains **30 journal-specific profiles plus the generic review profile**, covering PLOS, Nature Portfolio, Oxford University Press, Microbiology Society, ASM, and multiple Frontiers journals. Profiles encode only rules the current engine can actually evaluate.

Supported readiness checks include abstract limits, keyword limits, required sections, citation/reference integrity, tracked changes, and selected review-format targets. The transformed DOCX is kept only if the preservation gate passes.

### Template Mode

If the target journal supplies a `.docx` or `.dotx` Word template, select it directly in the desktop app.

Template Mode transfers a safe formatting subset to a **new copy** of the manuscript:

- page size and orientation
- page margins
- column layout
- line numbering
- selected standard Word styles such as Normal, Title, Heading 1-4, Caption, Quote, and Bibliography when present in both files

Template Mode does **not** copy the template's body text, instructional placeholders, headers, footers, macros, figures, citations, or relationships. Existing manuscript text and scientific content remain the source of truth.

This gives any journal with a Word template a safe adaptation path even when it does not yet have a built-in profile.

### 2. Citation Navigator

Use Citation Navigator when journal conversion is not needed and the goal is citation/reference traceability.

- Detects EndNote, Zotero, CSL, and Mendeley field signatures.
- Maps numbered citations to bibliography entries.
- Generates a local clickable HTML citation map.
- Creates internally linked DOCX copies for safely linkable plain numbered citations.
- Can create an explicitly requested **static linked review copy** from a live citation-manager manuscript.
- Keeps the original EndNote/Zotero/Mendeley manuscript unchanged.
- The Word add-in provides live-safe click-to-jump navigation without rewriting citation-manager payloads.

A linked review copy is appropriate for reading, review, proofreading, sharing, or other static use. Keep the original live manuscript if citations still need to be added, removed, refreshed, or restyled.

### 3. Manuscript Audit

Run structure, citation, figure, table, equation, field, comment, tracked-change, and other preservation-sensitive checks without converting the document and without choosing a journal.

## Journal coverage

Run:

```bash
word-journal-converter profiles
```

The catalog currently includes journal-specific profiles for:

- PLOS ONE, PLOS Biology, PLOS Genetics, PLOS Computational Biology, PLOS Pathogens, PLOS Neglected Tropical Diseases, PLOS Medicine
- Scientific Reports, Nature Communications, Nature Microbiology
- Nucleic Acids Research, Bioinformatics
- Microbial Genomics, Microbiology, Journal of General Virology, Journal of Medical Microbiology, International Journal of Systematic and Evolutionary Microbiology
- mBio, Applied and Environmental Microbiology, mSphere
- Frontiers in Microbiology, Genetics, Bioinformatics, Cellular and Infection Microbiology, Ecology and Evolution, Plant Science, Molecular Biosciences, Immunology, Medicine, and Veterinary Science

Every journal-specific profile records an official source URL and a checked date. Journal instructions can change, so the app warns when a profile becomes old and the official author instructions remain authoritative.

## Citation safety model

A research `.docx` can contain live citation-manager fields, bookmarks, hyperlinks, equations, figures, comments, tracked changes, footnotes, endnotes, custom XML, and internal package relationships.

The converter follows four rules:

1. **Do not rewrite scientific content during formatting operations.**
2. **Do not silently flatten live citation-manager fields.**
3. **When static flattening is requested, do it only in a separate review copy.**
4. **Do not keep a transformed file if the applicable preservation audit fails.**

## CLI quick start

List journal profiles:

```bash
word-journal-converter profiles
```

Run journal-specific analysis:

```bash
word-journal-converter analyze manuscript.docx --profile nature-communications-article --html-out manuscript_report.html
```

Apply supported profile formatting:

```bash
word-journal-converter retarget manuscript.docx --profile frontiers-microbiology-original-research --output manuscript_retargeted.docx
```

Inspect and apply a journal template:

```bash
word-journal-converter template-inspect journal_template.dotx
word-journal-converter template-retarget manuscript.docx --template journal_template.dotx --output manuscript_template_retargeted.docx --report template_report.json
```

Create a static linked review copy from a live citation-manager manuscript:

```bash
word-journal-converter make-navigable manuscript.docx --output manuscript_linked_review_copy.docx --static-review-copy
```

Verify any before/after pair:

```bash
word-journal-converter verify manuscript.docx manuscript_retargeted.docx
```

## Preservation gate

The standard verifier checks visible text, scientific numeric tokens, Word fields, live citation counts, embedded-media hashes, custom XML, package relationships, content types, tracked changes, comments, footnotes/endnotes, equations, tables, bookmarks, and hyperlinks.

Profile-based and template-based formatting outputs are removed automatically if the preservation gate fails.

## Custom journal profiles

Use either a bundled key or your own JSON file.

```bash
word-journal-converter validate-profile scientific-reports-article
word-journal-converter validate-profile path/to/custom-profile.json
```

Requirements that the engine cannot evaluate should not be encoded as if they were automatically checked.

## Privacy

The core does not require network access for manuscript inspection, citation navigation, report generation, template adaptation, retargeting, or preservation verification.

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

## Non-goals

Word Journal Manuscript Converter does not guarantee journal acceptance, supersede current publisher instructions, fabricate references, silently flatten live citations, copy instructional template text into manuscripts, promise pixel-perfect reproduction of every custom template, or submit manuscripts to publisher portals.

## Documentation

- [User guide](docs/USER_GUIDE.md)
- [Journal profiles](docs/JOURNAL_PROFILES.md)
- [Template Mode](docs/TEMPLATE_MODE.md)
- [Citation Navigator](docs/CITATION_NAVIGATOR.md)
- [Architecture](docs/ARCHITECTURE.md)
- [Preservation contract](docs/PRESERVATION_CONTRACT.md)
