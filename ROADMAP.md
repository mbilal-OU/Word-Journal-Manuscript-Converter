# Word Journal Manuscript Converter roadmap

## v0.2: usable integrity-first alpha

Completed in the current repository:

- DOCX package inspection
- live citation-manager signature inventory
- numbered and conservative author-year citation/reference graph
- manuscript structure extraction
- journal profile readiness checks
- safe formatting transforms for margins, line numbering, Normal style font, font size, and line spacing
- fail-closed post-transform preservation gate
- internal linking for simple plain numbered `[N]` citations
- local Tk desktop GUI
- Word task-pane starter
- landing page and release build workflows
- synthetic regression tests

## v0.3: citation and reference fidelity

- deeper Zotero field parsing without payload mutation
- deeper EndNote field parsing without payload mutation
- Mendeley/CSL adapter tests
- Word native bibliography and CITATION field support
- grouped/ranged plain numbered citation linking
- stronger author-year disambiguation for same-author/same-year references
- DOI and PMID metadata validation as an explicit opt-in network feature

## v0.4: manuscript structure and journal adaptation

- title-page model
- authors and affiliations
- corresponding-author metadata
- structured abstracts
- declaration sections
- anonymous-review copy generation
- caption and table-placement checks
- review-copy page numbering
- transformation preview and manifest

## v0.5: desktop release

- signed Windows application
- signed macOS application
- drag-and-drop workflow
- human-readable audit dashboard
- side-by-side transformation manifest
- profile manager with source/date warnings
- local settings with no manuscript-content telemetry

## v0.6: Word integration

- production task pane
- current-document readiness checks
- navigate to detected problems
- controlled local bridge to the desktop engine
- user confirmation before document-changing operations
- add-in distribution documentation

## v0.7: journal profile registry

- versioned profiles
- article-type-specific rules
- official source URLs
- checked dates
- staleness warnings
- review process for community-submitted profiles

## v1.0: stable release

- documented preservation guarantees
- broad citation-manager regression corpus
- broad DOCX feature corpus
- reproducible installers
- semantic versioning and migration policy
- stable Python API
- public issue templates that discourage uploading unpublished research
