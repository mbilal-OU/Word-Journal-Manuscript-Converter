# Architecture

## Product boundary

Word Journal Manuscript Converter is an integrity-preserving retargeting system for existing Word research manuscripts.

The current architecture separates six concerns:

1. **Package inspection**: direct read-only access to DOCX/OOXML parts.
2. **Manuscript structure**: headings, abstract, keywords, and references without flattening the DOCX.
3. **Citation graph**: links in-text citations to bibliography entries while treating live citation-manager fields as protected content.
4. **Journal profiles**: explicit, sourced, article-type-specific requirements.
5. **Safe transformation**: narrow OOXML/style operations with defined preconditions.
6. **Verification**: a fail-closed preservation gate after mutation.

## Current package layout

```text
src/word_journal_manuscript_converter/
  docx_package.py   package-safe OOXML reads and hashes
  audit.py          inventory and before/after preservation verification
  structure.py      manuscript structure extraction
  citations.py      citation/reference graph
  journal.py        journal-profile readiness checks
  retarget.py       safe profile-driven formatting transforms
  linking.py        internal links for simple plain numbered citations
  gui.py            local desktop GUI
  cli.py            command-line interface
```

## Why direct OOXML inspection

High-level DOCX libraries are useful for normal authoring tasks, but they do not guarantee round-trip preservation of every Word feature. Word Journal Manuscript Converter therefore starts from the ZIP package and XML parts, and changes only explicit package scopes that are covered by regression tests.

## Current safe transformation scope

`retarget.py` currently changes only:

- section page margins in `word/document.xml`
- line-numbering properties in `word/document.xml`
- Normal-style font, size, and line spacing in `word/styles.xml`

All other package parts are copied through unchanged.

`linking.py` changes only `word/document.xml`, adding internal bookmarks and internal hyperlinks for simple plain-text `[N]` citations. It refuses to run if live citation-manager fields are detected.

## Fail-closed mutation rule

A transformation is eligible for automatic execution only when:

- its preconditions can be detected;
- the changed OOXML scope is explicit;
- protected content can be fingerprinted before the change;
- the output remains a valid DOCX package;
- the post-transform preservation gate passes.

If the preservation gate fails, the transformed output is deleted.

## Interface separation

The core is interface-independent.

```text
CLI ---------\
Desktop GUI --- > Word Journal Manuscript Converter core -> DOCX copy -> preservation gate
Word add-in --/            |
                            +-> JSON audit/readiness reports
```

The Word add-in is intentionally a companion interface. Full package-level operations remain in the local desktop/core layer because it can inspect the entire DOCX package.
