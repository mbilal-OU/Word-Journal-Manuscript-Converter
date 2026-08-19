# User guide

## 1. Download or install

For most users, download the Windows or macOS package from the latest GitHub Release. Linux users can use the packaged CLI. Source installation remains available for reproducible workflows.

The current desktop binaries are unsigned/not notarized, so Windows SmartScreen or macOS Gatekeeper may show a warning.

## 2. Start with the original manuscript

Word Journal Manuscript Converter writes transformed results to a new file. Keep the original `.docx` as the source of truth.

## 3. Run a full local analysis

Desktop: choose a manuscript, select a bundled journal profile, then click **Full analysis**.

CLI:

```bash
word-journal-converter analyze manuscript.docx \
  --profile scientific-reports-article \
  --json-out manuscript_report.json \
  --html-out manuscript_report.html
```

The HTML report contains manuscript structure, a DOCX integrity inventory, the citation graph, journal-readiness checks, profile provenance, and the machine-readable audit record.

## 4. Inspect citation/reference integrity

```bash
word-journal-converter citations manuscript.docx
```

For plain numbered references, the converter maps citation numbers to reference entries. For live EndNote, Zotero, CSL, or Mendeley signatures, it reports the fields and leaves their payloads untouched.

## 5. Choose a journal profile

```bash
word-journal-converter profiles
word-journal-converter validate-profile plos-one-research-article
```

Bundled profiles include official source URLs and checked dates. Always open the current official instructions before actual submission.

## 6. Check readiness

```bash
word-journal-converter readiness manuscript.docx \
  --profile frontiers-microbiology-original-research
```

## 7. Retarget a copy

```bash
word-journal-converter retarget manuscript.docx \
  --profile frontiers-microbiology-original-research \
  --output manuscript_retargeted.docx \
  --report retarget-report.json
```

The output is removed automatically if the preservation audit fails.

## 8. Add internal links for plain numbered citations

```bash
word-journal-converter link-citations manuscript.docx \
  --output manuscript_linked.docx \
  --report citation-link-report.json
```

This operation supports simple plain-text `[N]` citations and refuses to run when live citation-manager fields are detected.

## 9. Verify an arbitrary before/after pair

```bash
word-journal-converter verify manuscript_before.docx manuscript_after.docx
```

A non-zero exit status indicates a preservation failure.
