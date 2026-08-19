# Word Journal Manuscript Converter Word add-in

The Word task pane now provides a **live-safe Citation Navigator** for open manuscripts. It can scan visible numbered citations and bibliography entries, detect EndNote, Zotero, and CSL/Mendeley field signatures, and jump the Word selection between a citation and its matching reference without rewriting citation-manager field payloads.

## Why the add-in matters

Live EndNote, Zotero, and Mendeley citations are structured Word fields. Wrapping or rewriting those fields just to force hyperlinks can make later citation refreshes unreliable. Citation Navigator therefore keeps those fields intact and uses Word's navigation APIs to select existing paragraphs/ranges instead.

For plain numbered citations with no live citation-manager fields, the desktop/CLI application can also create a separate clickable DOCX copy with internal bookmarks and hyperlinks.

## Current task-pane workflows

- **Citation Navigator**: scan numbered citations, map bibliography entries, show unresolved keys, jump to the first citation occurrence, and jump to the matching reference.
- **Quick integrity check**: count scientific numeric tokens, citation-manager signatures, bookmarks, hyperlinks, and equation nodes.

Both operations run inside Word through Office.js. They do not transmit manuscript text to a Word Journal Manuscript Converter service.

## Development

Office add-ins are web applications. For development, serve this directory over trusted HTTPS at `https://localhost:3000` and sideload `manifest.xml` into Word.

The current add-in is intentionally non-destructive in behavior. Full DOCX-package transformations remain in the local Python engine, where the complete package can be preservation-audited.

## Planned bridge

A future signed desktop integration can let the task pane request controlled transformations from the local core. The desktop engine will always operate on a copy, run the preservation gate, and return only a verified output.
