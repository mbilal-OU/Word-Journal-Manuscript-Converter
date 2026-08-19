# Research safety model

Word Journal Manuscript Converter is designed for unpublished scientific manuscripts.

Default rules:

1. Local processing for manuscript inspection and transformation.
2. No manuscript-body upload required for core use.
3. Original file is not overwritten by automatic transformations.
4. Live citation-manager fields are preserved by default.
5. Scientific text and numeric tokens are immutable during formatting-only operations.
6. Post-transform verification is mandatory.
7. Journal lookups, when added, should send only public journal identifiers or metadata needed for the lookup.
8. Debug logs must not contain full manuscript text by default.
9. Public bug reports should use synthetic or redacted documents.
