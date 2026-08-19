# Word Journal Manuscript Converter

**Retarget the format. Preserve the science.**

Word Journal Manuscript Converter is an open-source, local-first toolkit for existing Microsoft Word research manuscripts. It inspects preservation-sensitive DOCX structures, maps in-text citations to bibliography entries, checks explicit journal requirements, applies a limited set of safe formatting transformations, and verifies that protected scientific content survived unchanged.

> **Status: v0.2.1 alpha.** Word Journal Manuscript Converter is useful now for inspection, citation/reference auditing, controlled formatting, and preservation verification. It does not claim perfect journal conversion.

## Why Word Journal Manuscript Converter is different

A research manuscript is not just text plus font choices. A `.docx` can contain EndNote or Zotero fields, Word fields, bookmarks, hyperlinks, equations, figures, comments, tracked changes, footnotes, endnotes, custom XML, and internal relationships.

Word Journal Manuscript Converter follows three rules:

1. **Do not rewrite scientific content during formatting operations.**
2. **Do not flatten live citation-manager fields.**
3. **Do not keep a transformed file if the preservation audit fails.**

## What works now

### Inspect a manuscript

```bash
word-journal-converter inspect manuscript.docx
```

Inventories:

- paragraphs, tables, sections, and embedded media
- Office Math equations
- comments, footnotes, and endnotes
- tracked insertions and deletions
- Word fields and citation-manager field signatures
- bookmarks and hyperlinks
- custom XML parts

### Map in-text citations to references

```bash
word-journal-converter citations manuscript.docx
```

Current citation graph support:

- plain numbered citations such as `[4]`, `[2, 5]`, and `[7-10]`
- conservative author-year matching
- inventory of live EndNote, Zotero, CSL, and Mendeley signatures
- unresolved-citation and uncited-reference reporting

Live citation-manager payloads are treated as protected content.

### Extract manuscript structure

```bash
word-journal-converter structure manuscript.docx
```

Reports headings, abstract text and word count, keywords, reference section, and manuscript word count.

### Check a journal profile

```bash
word-journal-converter readiness manuscript.docx --profile journal-profile.json
```

Profiles can describe rules such as:

- abstract requirement and word limit
- keyword count
- required sections
- tracked-change policy
- citation/reference resolution
- review-copy formatting targets

Every journal profile should carry an official source URL and the date it was checked.

### Safely retarget a copy

```bash
word-journal-converter retarget manuscript.docx \
  --profile journal-profile.json \
  --output manuscript_retargeted.docx \
  --report retarget-report.json
```

v0.2 can safely target:

- page margins
- line numbering
- Normal-style body font
- body font size
- line spacing

The transformed file is removed automatically if the preservation gate fails.

### Add clickable plain numbered citations

```bash
word-journal-converter link-citations manuscript.docx \
  --output manuscript_linked.docx \
  --report link-report.json
```

For simple plain-text citations like `[12]`, Word Journal Manuscript Converter can add an internal Word link to the corresponding reference and bookmark the reference entry. The operation refuses to run when live citation-manager fields are detected.

### Verify before and after

```bash
word-journal-converter verify manuscript_before.docx manuscript_after.docx
```

The preservation gate currently checks:

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

## Desktop app

Launch the local GUI:

```bash
word-journal-converter gui
```

or:

```bash
word-journal-converter-gui
```

The GUI provides:

- Inspect
- Citation map
- Readiness
- Safe retarget
- Link `[N]` citations

Windows and macOS app binaries are planned through the release workflow in `.github/workflows/release.yml`.

## Microsoft Word add-in

A task-pane starter lives in [`integrations/word-addin/`](integrations/word-addin/).

It performs a lightweight quick check inside the current Word document and reports citation signatures, bookmarks, hyperlinks, equations, numeric tokens, and plain numbered citations. The starter does not send manuscript text to a Word Journal Manuscript Converter service.

Full package-level transformation stays in the desktop/core engine because that is where Word Journal Manuscript Converter can inspect the complete DOCX and enforce the preservation gate.

## Privacy

The current core does not require network access for manuscript inspection, citation mapping, retargeting, or preservation verification.

The intended default is:

```text
local manuscript
    -> local analysis
    -> transformed copy
    -> local preservation audit
    -> user review
```

See [`PRIVACY.md`](PRIVACY.md) and [`docs/RESEARCH_SAFETY.md`](docs/RESEARCH_SAFETY.md).

## Install from source

```bash
git clone https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter.git
cd Word-Journal-Manuscript-Converter
python -m venv .venv
```

Activate the environment and install:

```bash
pip install -e .
```

Then:

```bash
word-journal-converter --help
```

The core package intentionally uses the Python standard library only.

## Tests

```bash
pytest
```

The regression suite includes synthetic Word packages for citation mapping, link creation, field preservation, numeric-change detection, and safe retargeting.

## Journal profiles

Use [`journal-profiles/profile-template.json`](journal-profiles/profile-template.json) as the starting point. Do not treat demonstration profiles as official journal instructions.

See [`docs/JOURNAL_PROFILES.md`](docs/JOURNAL_PROFILES.md).

## Project structure

```text
src/word_journal_manuscript_converter/              core engine, CLI, GUI
journal-profiles/           profile templates and safe demos
profiles/                   schema and early profile examples
integrations/word-addin/    Word task-pane starter
packaging/                  executable build entry points
site/                       GitHub Pages landing page
scripts/                    local launch helpers
tests/                      OOXML regression tests
docs/                       architecture, safety, user guide
.github/workflows/          CI, release builds, Pages deployment
```

## Access strategy

**Primary:** signed desktop application for Windows and macOS.

**Companion:** Microsoft Word task pane for quick in-document checks and later controlled handoff to the desktop core.

**Technical:** CLI and Python API for reproducible research workflows.

**Later:** browser/PWA only where manuscript privacy can remain explicit and safe.

## Non-goals

Word Journal Manuscript Converter does not:

- guarantee journal acceptance
- fabricate references
- rewrite scientific claims during formatting
- silently flatten EndNote or Zotero fields
- submit manuscripts to publisher portals
- treat third-party journal instructions as authoritative when an official source exists

## Documentation

- [User guide](docs/USER_GUIDE.md)
- [Architecture](docs/ARCHITECTURE.md)
- [Preservation contract](docs/PRESERVATION_CONTRACT.md)
- [Research safety](docs/RESEARCH_SAFETY.md)
- [Journal profiles](docs/JOURNAL_PROFILES.md)
- [Word add-in strategy](docs/WORD_ADDIN.md)
- [Access strategy](docs/ACCESS.md)
- [Roadmap](ROADMAP.md)
- [FAQ](docs/FAQ.md)

## License

MIT. See [`LICENSE`](LICENSE).
