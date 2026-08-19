# Template Mode

Template Mode adapts an existing manuscript to formatting information from a journal-supplied Microsoft Word template while keeping the manuscript itself as the source of scientific content.

## Accepted template files

Template Mode accepts:

- `.docx`
- `.dotx`

The source manuscript must remain a `.docx`.

## What is transferred

When present in the template, the current engine can transfer this conservative formatting subset:

- page size and orientation
- page margins
- column layout
- Word line numbering
- selected standard styles that exist in both files:
  - Normal
  - Title
  - Subtitle
  - Heading 1 through Heading 4
  - Caption
  - Quote
  - Intense Quote
  - List Paragraph
  - Bibliography

For styles, only paragraph and run formatting properties are transferred. The manuscript text is not rebuilt from template content.

## What is never copied from the template

Template Mode deliberately does not copy:

- template body text
- instructional placeholder paragraphs
- sample author names or affiliations
- headers or footers
- macros
- citations or bibliography fields
- figures or media
- comments or tracked changes
- template package relationships
- custom XML from the template

This prevents journal instructions or example content from being inserted into a research manuscript accidentally.

## Preservation gate

Template Mode writes a new `.docx` and then runs the standard fail-closed preservation verifier against the original manuscript.

The output is retained only when protected manuscript content passes verification, including visible text, scientific numeric tokens, Word fields, citation-manager fields, media hashes, custom XML, relationships, content types, tracked changes, comments, notes, equations, tables, bookmarks, and hyperlinks.

If verification fails, the output is deleted.

## Desktop workflow

1. Select the manuscript.
2. Open **Journal Conversion**.
3. Under **Journal Word template**, choose the journal `.docx` or `.dotx`.
4. Click **Inspect template** to see the transferable formatting detected.
5. Click **Apply template safely...**.
6. Save to a new filename.
7. Review the resulting document in Microsoft Word before submission.

The original manuscript is never overwritten.

## CLI

Inspect a template:

```bash
word-journal-converter template-inspect journal_template.dotx
```

Apply the safe template subset:

```bash
word-journal-converter template-retarget manuscript.docx \
  --template journal_template.dotx \
  --output manuscript_template_retargeted.docx \
  --report template_report.json
```

## Important limitation

A Word template can contain complex direct formatting, custom styles, section-specific layout, text boxes, content controls, macros, and publisher-specific automation. v0.5.0 does not claim pixel-perfect cloning of those features.

Direct formatting already applied to individual manuscript paragraphs or runs can override transferred styles. Always inspect the new copy in Word and compare it with the journal's current instructions before submission.
