# Preservation contract

Word Journal Manuscript Converter's primary safety invariant is simple:

> A journal-retargeting operation must not silently change scientific content or destroy document semantics that the researcher may need later.

## Protected by default

- manuscript prose
- numeric values and units
- live EndNote, Zotero, Mendeley/CSL, and Word citation fields when detectable
- bibliography field data
- bookmarks and internal hyperlinks
- REF/PAGEREF fields
- equations
- figures and embedded media
- table content
- footnotes and endnotes
- comments
- tracked changes
- custom XML
- package relationships required by protected objects
- package content-type declarations

## Citation rule

An in-text citation and its reference entry are treated as one linked system. A transformation must never renumber, rewrite, or flatten one side independently.

For live citation-manager fields, automatic transformations preserve the field instructions. For plain-text citations, Word Journal Manuscript Converter may create Word-native internal links only when it can map the citation to a reference with a supported deterministic rule.

## Explicit exceptions

A protected feature may be changed only when:

1. the user explicitly requests the change;
2. the operation records the change in its transformation manifest;
3. the relevant preservation rule is adjusted intentionally; and
4. a recoverable source copy remains available.

## Current preservation gate

The v0.2 gate compares:

- visible text
- numeric-token sequence
- complete Word field instructions
- live citation-field counts
- embedded-media hashes
- custom XML hashes
- relationship-part hashes
- `[Content_Types].xml`
- tracked insertion/deletion counts
- comment counts
- footnote/endnote counts
- equation counts
- table counts
- bookmark loss
- hyperlink loss

Adding bookmarks or hyperlinks is allowed for the explicit citation-link operation, but losing them is not.

## Future verification levels

- package validity
- semantic inventory preservation
- stronger unit/value fingerprints
- citation-manager fixture suites
- document-property and numbering audits
- before/after visual regression in desktop releases
- optional Word-open validation on supported platforms
