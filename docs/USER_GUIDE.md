# User guide

## 1. Start with a copy

Word Journal Manuscript Converter never intends to overwrite the source manuscript. Keep the original `.docx` as the source of truth.

## 2. Inspect

```bash
word-journal-converter inspect manuscript.docx
```

This inventories sensitive Word structures before any transformation.

## 3. Check citation/reference integrity

```bash
word-journal-converter citations manuscript.docx
```

For plain numbered references, Word Journal Manuscript Converter maps citation numbers to reference entries. For live EndNote, Zotero, CSL, or Mendeley signatures, it reports the fields and leaves their payloads untouched.

## 4. Check a journal profile

```bash
word-journal-converter readiness manuscript.docx --profile journal-profiles/profile-template.json
```

Profiles must be populated from current official journal instructions. Word Journal Manuscript Converter does not treat an old profile as timeless truth.

## 5. Retarget a copy

```bash
word-journal-converter retarget manuscript.docx \
  --profile journal-profile.json \
  --output manuscript_retargeted.docx \
  --report retarget-report.json
```

The output is removed automatically if the preservation audit fails.

## 6. Add internal links for plain numbered citations

```bash
word-journal-converter link-citations manuscript.docx \
  --output manuscript_linked.docx \
  --report citation-link-report.json
```

This release supports simple plain-text `[N]` citations. It refuses the operation when live citation-manager fields are detected.

## 7. Desktop GUI

```bash
word-journal-converter gui
```

or

```bash
word-journal-converter-gui
```

The GUI exposes Inspect, Citation map, Readiness, Safe retarget, and Link `[N]` citations.
