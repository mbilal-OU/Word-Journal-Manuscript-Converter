# Word add-in strategy

The Word task pane is a companion interface, not a replacement for the package-level core.

## Current starter

`integrations/word-addin/` contains a task pane that reads the current Word body OOXML and reports quick counts for citation signatures, bookmarks, hyperlinks, equations, numeric tokens, and plain numbered citations.

The quick check does not send manuscript text to a Word Journal Manuscript Converter service.

## Why the full transformation remains in the desktop core

The desktop engine can inspect the complete DOCX package, including package parts outside the body OOXML, and can run a strict before/after preservation audit. The add-in cannot provide the same package-level guarantee by itself.

## Long-term model

Word task pane -> signed local Word Journal Manuscript Converter desktop bridge -> transformed copy -> preservation audit -> user review.
