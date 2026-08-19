# v0.4.0 validation checklist

This release introduces the dual-purpose workflow: Journal Conversion, Citation Navigator, and Manuscript Audit.

Validation targets:

- `pytest`
- CLI version command
- `navigate` command on a synthetic numbered-citation DOCX
- `make-navigable` creates a new clickable copy for plain numbered citations
- live EndNote/Zotero/Mendeley fields trigger safe non-mutating mode
- desktop GUI starts and exposes the three workflows
- frozen release builds retain bundled journal profiles

The Citation Navigator never rewrites live citation-manager field payloads in automatic mode.
