# Citation Navigator

Citation Navigator is the journal-independent workflow in Word Journal Manuscript Converter. It is intended for students, researchers, thesis writers, reviewers, and collaborators who want citations and bibliography entries to be easier to trace without changing the scientific content of the master manuscript.

## Three navigation paths

### Live-safe Word navigation

Use this for documents containing live EndNote, Zotero, or CSL/Mendeley citation-manager fields when you still need to edit, refresh, or reformat those citations.

The Word add-in scans the visible citation numbers and bibliography and moves the Word selection to the requested citation or reference location. It does not rewrite citation-manager field instructions.

### Linked review copy

Use this when a live EndNote/Zotero/Mendeley manuscript needs a separate, static, easy-to-review copy with internal citation/reference navigation.

The desktop app detects live fields **before** asking for a save location and gives you an explicit choice:

- **Use Live Navigation** keeps the master manuscript live and opens the Word add-in guide.
- **Create Linked Review Copy** creates a separate static `.docx`.
- **Cancel** changes nothing.

For the linked review copy, the converter flattens supported citation-manager fields in the copy only, preserves their visible text, then adds bibliography bookmarks and internal citation links where supported. The original manuscript is never modified.

The linked review copy is intentionally static. Do not use it as the master document for refreshing or reformatting EndNote, Zotero, or Mendeley citations.

### Plain clickable DOCX export

For documents that already contain plain numbered citations and no live citation-manager fields, Citation Navigator can create a clickable DOCX directly.

## Static review-copy preservation gate

A linked review copy is kept only if the dedicated static-copy audit confirms all expected protected content remains intact. The audit checks:

- visible text
- scientific numeric tokens
- non-citation Word fields
- embedded media
- custom XML
- package relationships
- content types
- tracked changes
- comments
- footnotes and endnotes
- equations
- tables
- bookmark retention
- hyperlink retention

The only expected structural difference is that supported citation-manager field instructions are removed from the **review copy**.

## Desktop workflow

1. Open **Citation Navigator**.
2. Choose a `.docx` manuscript.
3. Click **Analyze navigation**.
4. Review the detected citation manager, live-field count, reference count, matched links, and unresolved keys.
5. Save the citation-navigation HTML report if desired.
6. Click **Create navigable copy**.
7. If live fields are detected, choose either **Use Live Navigation** or **Create Linked Review Copy**.
8. If creating a review copy, choose a new output filename and keep the original manuscript as your editable citation-manager master.

No journal profile is required.

## CLI

Analyze citation navigation:

```bash
word-journal-converter navigate manuscript.docx --html-out citation_navigation.html
```

Create a clickable copy for a plain numbered document:

```bash
word-journal-converter make-navigable manuscript.docx --output manuscript_navigable.docx
```

Explicitly create a static linked review copy from a live citation-manager manuscript:

```bash
word-journal-converter make-navigable manuscript.docx \
  --output manuscript_linked_review_copy.docx \
  --static-review-copy
```

Without `--static-review-copy`, a live-field document returns a safe refusal instead of flattening citation fields.

## Word add-in

The task pane scans the open Word document, identifies a `References` or `Bibliography` heading, maps numbered citations, and provides controls to jump to the first citation occurrence or the matching bibliography entry.

This is the preferred path when the manuscript must remain fully live in EndNote, Zotero, or Mendeley.

## Current limits

- Live-safe Word navigation currently focuses on numbered citation styles.
- Static linked review copies preserve visible field results but are not citation-manager-editable by design.
- Internal DOCX linking currently targets simple visible numbered citation tokens conservatively; complex range/group presentation may remain partially unlinked.
- Author-year citation analysis is available in the core, but the first navigation interface is optimized for numbered references.
- Unusual or malformed Word field structures may be refused rather than flattened.

## Safety rule

Citation traceability must never be gained by silently sacrificing the editable master manuscript. Live-field flattening requires explicit user choice and occurs only in a separate review copy.
