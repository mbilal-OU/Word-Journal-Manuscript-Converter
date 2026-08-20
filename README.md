# Word Journal Manuscript Converter

**Prepare the manuscript. Trace every citation. Preserve the science.**

[![CI](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/actions/workflows/ci.yml/badge.svg)](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/actions/workflows/ci.yml)
[![Release builds](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/actions/workflows/release.yml/badge.svg)](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/actions/workflows/release.yml)
[![Pages](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/actions/workflows/pages.yml/badge.svg)](https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/actions/workflows/pages.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

Word Journal Manuscript Converter is a local-first toolkit for existing Microsoft Word research manuscripts. It combines journal preparation, citation/reference navigation, manuscript auditing, and safe adaptation from publisher Word templates.

**Public status: Early Access.**  
Internal package version: `0.5.0`. The first stable public launch is reserved for `1.0.0`. The current engineering release tag remains `v0.5.0-beta.1` for traceability.

Developed by **Muhammad Bilal**.

## Download

For Windows, the installer is the recommended package. It adds Start menu integration and supports in-app update downloads.

| Platform | Package | Use |
|---|---|---|
| Windows | `Word-Journal-Manuscript-Converter-Setup.exe` | Recommended installer |
| Windows x64 | `Word-Journal-Manuscript-Converter-Windows-x64.zip` | Portable GUI + CLI |
| macOS | `Word-Journal-Manuscript-Converter-macOS.zip` | GUI + CLI |
| Linux x64 | `Word-Journal-Manuscript-Converter-Linux-x64.tar.gz` | CLI |

Every release asset has a matching SHA-256 checksum.

Website: https://mbilal-ou.github.io/Word-Journal-Manuscript-Converter/

Early Access Windows and macOS builds are not yet code-signed/notarized, so operating-system security prompts may appear. Code signing is a launch-readiness item.

## Product workflows

### Journal Conversion

Choose a source-dated built-in profile, use a custom JSON profile, or supply a journal Word template.

The catalog currently contains **30 journal-specific profiles plus a generic review profile** across PLOS, Nature Portfolio, Oxford University Press, Microbiology Society, ASM, and multiple Frontiers journals.

Supported profile-based operations are intentionally limited to requirements the engine can evaluate or safely apply.

### Journal Template Mode

If a publisher supplies a `.docx` or `.dotx` template, Template Mode transfers a conservative formatting subset into a **new copy** of the manuscript:

- page size and orientation
- page margins
- column layout
- line numbering
- selected standard Word styles such as Normal, Title, Heading 1-4, Caption, Quote, and Bibliography when present in both files

Template Mode does **not** copy template body text, placeholders, headers, footers, macros, figures, citations, or unrelated relationships.

### Citation Navigator

Use Citation Navigator without selecting a journal.

- detects EndNote, Zotero, CSL, and Mendeley field signatures
- maps numbered citations to bibliography entries
- generates a clickable local HTML citation map
- creates internally linked DOCX copies for safely linkable plain numbered citations
- can create an explicitly requested static linked review copy from a live citation-manager manuscript
- keeps the original EndNote/Zotero/Mendeley manuscript unchanged
- provides a Word add-in for live-safe citation/reference jumps

A linked review copy is suitable for reading, review, proofreading, sharing, and other static use. Keep the original live manuscript whenever citations still need to be added, removed, refreshed, or restyled.

### Manuscript Audit

Run citation, structure, figure, table, equation, field, comment, tracked-change, bookmark, hyperlink, and other preservation-sensitive checks without converting anything.

## Preservation model

The converter follows four rules:

1. Do not rewrite scientific content during formatting operations.
2. Do not silently flatten live citation-manager fields.
3. When static flattening is explicitly requested, do it only in a separate review copy.
4. Do not keep a transformed file if the applicable preservation audit fails.

The preservation gate checks visible text, scientific numeric tokens, Word fields, citation-field counts, embedded media, custom XML, relationships, content types, tracked changes, comments, notes, equations, tables, bookmarks, and hyperlinks.

## Word add-in

The companion Word add-in provides live-safe citation/reference navigation directly inside Microsoft Word.

Installation and testing instructions are available on the hosted [Word Citation Navigator page](https://mbilal-ou.github.io/Word-Journal-Manuscript-Converter/word-addin/) and in [docs/WORD_ADDIN.md](docs/WORD_ADDIN.md).

The add-in web application is hosted over HTTPS on GitHub Pages. For the stable public launch, the intended distribution path is Microsoft Marketplace rather than development sideloading.

## Updates

The desktop app can check GitHub Releases for newer builds.

On Windows, when a newer release contains the setup executable, the app can download and launch that installer. This means an installed user does not need to manually revisit GitHub for every update.

Automatic update checks can be disabled under **Privacy & analytics**.

See [docs/UPDATES.md](docs/UPDATES.md).

## Optional product analytics

Anonymous analytics are **opt-in** in the desktop app and Word add-in. The website also asks before recording anonymous product events.

When enabled, analytics can record:

- product feature used
- app version
- operating-system family/version
- anonymous installation/session identifiers
- download button clicks
- page views
- approximate session duration

Analytics do **not** include:

- manuscript text
- filenames or paths
- citation or reference text
- figures or tables
- document hashes
- journal manuscript metadata

Feedback is submitted only when the user presses **Submit**. Contact email is optional.

See [PRIVACY.md](PRIVACY.md) and [docs/ANALYTICS.md](docs/ANALYTICS.md).

## CLI examples

```bash
word-journal-converter --version
word-journal-converter profiles
word-journal-converter analyze manuscript.docx --profile nature-communications-article
word-journal-converter navigate manuscript.docx --html-out citation_navigation.html
word-journal-converter template-inspect journal_template.dotx
word-journal-converter template-retarget manuscript.docx --template journal_template.dotx --output retargeted.docx
word-journal-converter check-update
```

## Install from source

```bash
git clone https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter.git
cd Word-Journal-Manuscript-Converter
python -m venv .venv
pip install -e .
word-journal-converter --help
```

The core package uses the Python standard library only.

## Tests

```bash
pytest
```

Regression tests use synthetic DOCX packages so unpublished research is never needed in the repository test corpus.

## Launch-readiness status

Early Access is intended for structured real-world testing before `1.0.0`.

Remaining launch gates include:

- broad regression testing across real EndNote, Zotero, and Mendeley documents
- real publisher Word-template testing
- Windows code signing
- macOS notarization
- Microsoft Marketplace submission for the Word add-in
- user feedback and telemetry review
- final privacy, packaging, and update-flow verification

## License

MIT. See [LICENSE](LICENSE).

## Citation

See [CITATION.cff](CITATION.cff).
