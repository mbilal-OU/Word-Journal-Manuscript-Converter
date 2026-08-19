# Word Journal Manuscript Converter Word add-in starter

This task-pane companion performs a lightweight, local inspection of the currently open Word document. It uses Word's OOXML access to count preservation-sensitive structures and citation-manager signatures without sending manuscript text to a Word Journal Manuscript Converter service.

## Development

Office add-ins are web applications. For development, serve this directory over trusted HTTPS at `https://localhost:3000` and sideload `manifest.xml` into Word.

The current add-in is intentionally **read-only in behavior**, even though the manifest requests `ReadWriteDocument` for the later controlled-fix workflow. Full DOCX-package transformations remain in the local Python engine, where the complete package can be preservation-audited.

## Planned bridge

A later release can connect the task pane to a signed local Word Journal Manuscript Converter desktop service. The add-in will request a transformation, the desktop core will operate on a copy of the `.docx`, run the preservation gate, and only then return the new document to the user.
