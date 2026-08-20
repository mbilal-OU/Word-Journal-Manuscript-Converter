# Word add-in

This directory contains the Early Access Word Citation Navigator add-in.

Files:

- `manifest.xml` - Office add-in-only manifest
- `taskpane.html` - task pane UI
- `taskpane.js` - citation/reference navigation logic
- `assets/` - add-in icons

The hosted installation page is:

`https://mbilal-ou.github.io/Word-Journal-Manuscript-Converter/word-addin/`

The hosted task pane is deployed through GitHub Pages at:

`https://mbilal-ou.github.io/Word-Journal-Manuscript-Converter/addin/taskpane.html`

For installation and deployment guidance, see:

- [`docs/WORD_ADDIN.md`](../../docs/WORD_ADDIN.md)
- https://mbilal-ou.github.io/Word-Journal-Manuscript-Converter/word-addin/

## Safety

Citation jumps change only the Word selection. They do not rewrite EndNote, Zotero, or Mendeley field payloads.

Optional add-in analytics are off until the user enables them. Analytics never include document text, filenames, citation text, reference text, or document metadata.

Word Journal Manuscript Converter is independently developed and is not affiliated with or endorsed by Microsoft, Clarivate/EndNote, Zotero, Mendeley, or any journal publisher.

Developed by Muhammad Bilal.
