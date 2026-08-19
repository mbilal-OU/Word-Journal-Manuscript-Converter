# Word Journal Manuscript Converter v0.4.0

v0.4.0 turns the project into a dual-purpose research tool.

## Three workflows

### Journal Conversion
Use a verified journal profile to audit explicit requirements and apply only supported formatting changes to a new DOCX copy.

### Citation Navigator
Use the converter without selecting any journal. The navigator detects the citation-manager environment, maps visible citations to bibliography entries, and provides a local clickable HTML citation map. Plain numbered citations can be exported as a new internally linked DOCX. Documents with live EndNote, Zotero, CSL, or Mendeley fields remain protected and are routed to non-mutating Word navigation.

### Manuscript Audit
Inspect manuscript structure and preservation-sensitive DOCX content without conversion.

## Safety model

Live citation-manager payloads are not rewritten to force hyperlinks. The navigation strategy is selected conservatively:

- live citation-manager fields: non-mutating Word add-in navigation
- plain numbered citations: optional clickable DOCX export
- unsupported or ambiguous citation patterns: analysis only

All core analysis remains local-first.
