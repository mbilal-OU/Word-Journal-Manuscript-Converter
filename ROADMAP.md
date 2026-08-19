# Word Journal Manuscript Converter roadmap

## v0.3: public beta

Completed in the current repository:

- DOCX package inspection and preservation audit
- citation/reference graph for numbered and conservative author-year citations
- live EndNote, Zotero, CSL, and Mendeley signature inventory
- bundled, source-dated journal profiles
- profile freshness warnings
- combined full-analysis workflow
- local HTML report export
- safe formatting transforms for margins, line numbering, Normal style font, size, and line spacing
- fail-closed post-transform preservation gate
- internal linking for simple plain numbered citations
- desktop GUI with bundled profile selector
- Word task-pane starter
- direct release packaging with checksums
- GitHub Pages landing page
- synthetic regression tests

## v0.4: citation-manager fidelity

- deeper Zotero field parsing without payload mutation
- deeper EndNote field parsing without payload mutation
- dedicated Mendeley/CSL regression fixtures
- Word native bibliography and CITATION field support
- stronger author-year disambiguation
- grouped/ranged plain-numbered citation navigation
- explicit opt-in DOI/PMID validation

## v0.5: manuscript structure adaptation

- title-page model
- authors and affiliations
- corresponding-author metadata
- structured abstracts
- declaration-section modeling
- anonymous-review copy generation
- caption and table-placement checks
- page-number checks and controlled page-number insertion
- transformation preview/manifest

## v0.6: desktop hardening

- code-signed Windows builds
- notarized macOS builds
- drag-and-drop manuscript workflow
- larger human-readable audit dashboard
- local profile manager with stale-profile warnings
- settings with no manuscript-content telemetry

## v0.7: Word integration

- production task pane
- navigate to detected manuscript problems
- controlled local bridge to the desktop engine
- explicit confirmation before document-changing operations
- add-in deployment documentation

## v1.0: stable release

A v1.0 label will be used only after the project has a broad synthetic/consented DOCX regression corpus, reproducible signed installers, documented preservation guarantees, stable APIs, and citation-manager coverage sufficient to justify a stable compatibility claim.
