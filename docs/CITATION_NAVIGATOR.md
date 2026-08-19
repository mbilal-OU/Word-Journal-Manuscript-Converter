# Citation Navigator

Citation Navigator is the journal-independent workflow in Word Journal Manuscript Converter. It is intended for students, researchers, thesis writers, reviewers, and collaborators who want citations and bibliography entries to be easier to trace without changing the manuscript's scientific content.

## Two navigation modes

### Live-safe Word navigation

Use this for documents containing live EndNote, Zotero, or CSL/Mendeley citation-manager fields.

The desktop/core application detects the live fields and deliberately refuses to rewrite or wrap them to force hyperlinks. The Word add-in instead scans the visible citation numbers and bibliography, then moves the Word selection to the requested citation or reference paragraph.

This changes navigation state only. It does not rewrite citation-manager field instructions.

### Clickable DOCX export

Use this for plain numbered citations such as `[4]`, `[2, 5]`, or `[7-10]` when no live citation-manager fields are present.

The converter can create a separate DOCX copy, add bookmarks to bibliography entries, and add internal hyperlinks to simple `[N]` citation tokens. The output is kept only if the preservation audit passes.

## Desktop workflow

1. Open **Citation Navigator**.
2. Choose a `.docx` manuscript.
3. Click **Analyze navigation**.
4. Review the detected citation manager, citation mode, reference count, matched links, and unresolved keys.
5. Save the local citation-navigation HTML report if desired.
6. If the document is plain numbered text, use **Create clickable copy**.
7. If live citation-manager fields are detected, use the Word add-in for non-mutating in-document navigation.

No journal profile is required.

## CLI

Analyze citation navigation:

```bash
word-journal-converter navigate manuscript.docx --html-out citation_navigation.html
```

Create a clickable copy when safe:

```bash
word-journal-converter make-navigable manuscript.docx --output manuscript_navigable.docx
```

A live-field document returns a safe refusal instead of modifying the DOCX.

## Word add-in

The task pane scans the visible Word paragraphs, identifies a `References` or `Bibliography` heading, maps numbered citations in the manuscript body, and provides two controls for each detected key:

- **Citation**: jump to the first paragraph containing that citation key.
- **Reference**: jump to the matching bibliography paragraph.

The add-in uses Word's selection/navigation APIs and does not alter citation-manager payloads.

## Current limits

- Live-safe Word navigation currently focuses on numbered citation styles.
- Plain clickable DOCX export is intentionally conservative and does not force links into live fields.
- Author-year citation analysis is available in the core, but the first Word add-in navigation interface is optimized for numbered references.
- A malformed or unusually structured bibliography may require manual review.

## Safety rule

Citation traceability must never be gained by sacrificing citation-manager integrity. When the converter cannot prove that a DOCX mutation is safe, it leaves the manuscript unchanged and provides an analysis or navigation-only path instead.
