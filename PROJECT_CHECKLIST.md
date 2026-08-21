# Product Development Checklist

Last updated: 2026-08-20

This checklist is the short, persistent record of what is done, what is being validated, and what remains. `PROJECT_STATUS.md` contains the more detailed engineering evidence and readiness estimates.

Status meanings:

- [x] Implemented and validated to the level stated
- [ ] Not complete yet
- **VALIDATING** means implemented but still needs broader CI, Microsoft Word, or real-document evidence
- **PARTIAL** means useful support exists but important cases remain

## Current product priorities

1. Make Manuscript Clickable / Citation Navigator
2. Preflight Conversion Eligibility
3. Cross-discipline journal and article-type coverage
4. Word Template Coverage Expansion
5. Broader citation-manager and real-document validation
6. Signing, notarization, Word add-in distribution, and stable 1.0

## Release and document-safety foundation

- [x] Original manuscript is never overwritten by conversion/navigation workflows.
- [x] Visible text preservation gate.
- [x] Scientific numeric-token preservation gate.
- [x] Word field-instruction preservation checks.
- [x] EndNote/Zotero/CSL candidate-field inventory.
- [x] Embedded media preservation checks.
- [x] Custom XML, relationships, and content-types preservation checks.
- [x] Comments, footnotes/endnotes, tracked changes, tables, and equation inventory.
- [x] Defensive OOXML structural validation.
- [x] Markup-compatibility namespace preservation after XML rewrites.
- [x] Fail-closed removal of output when a safety gate fails.
- [x] Tested real EndNote manuscript opens without Word repair after journal conversion.
- [x] Tested real EndNote manuscript opens without Word repair after template adaptation.
- [x] Tested real EndNote static linked review copy opens without Word repair.
- [x] Official Early Access release workflow passed on Windows, macOS, and Linux (run 32437167856).
- [ ] Broaden Microsoft Word openability testing across 30-50 diverse real/sanitized manuscripts.

## Make Manuscript Clickable / Citation Navigator

### User experience

- [x] No journal is required to use citation navigation.
- [x] Original manuscript remains unchanged; navigation is created in a separate copy.
- [x] Citation analysis can run without modifying the document.
- [x] Citation-manager master file is protected when live fields are present.
- [x] Separate static linked review copy is available by explicit user choice for live citation-manager documents.
- [ ] Rename/reframe the desktop primary workflow more prominently as **Make Manuscript Clickable** throughout the interface. **PARTIAL**
- [ ] Add a simple citation-style override only when automatic detection confidence is low.

### Citation-style detection

- [x] Numbered square-bracket detection such as `[1]`.
- [x] Numbered parenthetical detection such as `(1)` implemented. **VALIDATING**
- [x] Superscript-numbered detection implemented. **VALIDATING**
- [x] Author-year detection implemented for common parenthetical and narrative forms. **VALIDATING**
- [x] Detection confidence score added.
- [x] Ambiguous author-year matches are left unresolved rather than guessed.
- [x] Full Python 3.10-3.13 CI is green for the current navigation branch (run 32439674266).
- [ ] Footnote/endnote citation-system navigation. **PLANNED**
- [ ] Legal and other specialized citation systems. **PLANNED**

### Forward navigation: in-text citation -> reference

- [x] Simple numbered citations link to bibliography entries.
- [x] Explicit numbers inside grouped citations can be linked individually, e.g. `[1, 2]`. **VALIDATING**
- [x] Parenthetical numbered citations can be linked. **VALIDATING**
- [x] Superscript numbered citations can be linked. **VALIDATING**
- [x] Unambiguous author-year citations can be linked to bibliography entries. **VALIDATING**
- [ ] Ranges such as `[2-5]`: visible endpoints can be linked without changing manuscript text, but implied interior references cannot each have a separate static hyperlink target. **PARTIAL**
- [ ] Cross-run citations split by unusual Word formatting need broader support. **PARTIAL**

### Reverse navigation: reference -> in-text citation

- [x] Citation occurrence bookmarks implemented. **VALIDATING**
- [x] Matched references link back to the first linked in-text occurrence. **VALIDATING**
- [x] Reverse linking does not add visible text to the bibliography.
- [ ] Add-in navigation to previous/next/all citation occurrences for references cited multiple times. **PLANNED**
- [ ] Optional static-document all-occurrence backlink UI only if it can be done without cluttering manuscript content. **PLANNED**

### Citation integrity/reporting

- [x] Citation graph reports matched and unmatched references.
- [x] Citation graph now records citation style and detection confidence. **VALIDATING**
- [x] Ambiguous author-year citations are reported separately. **VALIDATING**
- [x] Navigation HTML report remains available.
- [x] New linking code uses namespace-preserving XML serialization.
- [ ] Real Microsoft Word test of the new bidirectional numbered copy.
- [ ] Real Microsoft Word test of author-year navigation.
- [ ] Real Microsoft Word test of grouped/ranged numeric citations.
- [ ] Real Zotero document navigation test.
- [ ] Real Mendeley document navigation test.

## Preflight Conversion Eligibility

- [ ] Add a pre-conversion eligibility engine before journal/template conversion.
- [ ] Technical safety gate: valid DOCX, safe OOXML, relationships, namespaces.
- [ ] Scientific-content preservation gate: fields, equations, tables, figures, comments, notes, tracked changes.
- [ ] Semantic mappability score for title/authors/affiliations/abstract/headings/references/etc.
- [ ] Journal/article-type fit score.
- [ ] Template-compatibility score for Template Mode.
- [ ] Four clear outcomes: AUTO-CONVERSION ELIGIBLE / ELIGIBLE WITH REVIEW / FORMAT-ONLY ELIGIBLE / BLOCKED.
- [ ] Missing journal-required content must not be invented automatically.
- [ ] Missing declarations may permit formatting but must prevent a false submission-ready claim.
- [ ] Generate a mandatory author review checklist after conversion.

## Cross-discipline scholarly support

- [x] Product direction is discipline-agnostic, not life-science-only.
- [ ] Replace life-science assumptions in semantic mapping with article-type/profile-driven rules.
- [ ] Life sciences and medicine coverage.
- [ ] Chemistry coverage, including schemes/compound-heavy manuscripts.
- [ ] Physics coverage, including equation-heavy layouts.
- [ ] Engineering coverage.
- [ ] Computer science / ACM / IEEE coverage.
- [ ] Mathematics coverage, including theorem/proof-heavy manuscripts.
- [ ] Social sciences and psychology coverage.
- [ ] Economics/business coverage.
- [ ] Humanities coverage, including footnote/endnote-heavy manuscripts.
- [ ] Legal/specialized citation workflows where safely supportable.
- [ ] Publisher-family abstraction plus journal-specific profiles plus bring-your-own-template mode.

## Journal Profile Conversion

- [x] 30+ bundled profiles available.
- [x] Journal readiness analysis.
- [x] Margins, line numbering, basic body style, and spacing transforms.
- [x] Conversion Assurance and fail-closed output.
- [x] No universal publisher-compliance claim.
- [ ] Expand profiles across disciplines and article types.
- [ ] Add eligibility gate before conversion.
- [ ] Strengthen current-guideline/profile maintenance workflow.

## Word Template Mode

- [x] `.docx` and `.dotx` inspection.
- [x] Page size/orientation, margins, columns, and line-numbering adaptation.
- [x] Matching styles, safe custom styles, document defaults, semantic-role inference, and safe theme handling.
- [x] Supported Fidelity and Template Coverage are reported separately.
- [x] Low coverage is not presented as complete conversion.
- [x] Real tested template: 100% supported fidelity, 33% coverage, Word openability PASS.
- [ ] Raise real-template coverage materially above 33% without weakening document safety. **PR #8 workstream**
- [ ] Style dependency import.
- [ ] Relationship-safe numbering transfer.
- [ ] Relationship-safe headers/footers/page numbering.
- [ ] Table/caption styles.
- [ ] More complete section-layout support.
- [ ] Safe content-control/placeholder analysis.
- [ ] At least 10 distinct real templates before Release Candidate.
- [ ] At least 20 distinct real templates before stable 1.0.

## Manuscript Audit

- [x] Structure, paragraphs, sections, references, tables, media, equations, comments, notes, tracked changes, fields, bookmarks, and hyperlinks inventory.
- [x] Citation/reference graph.
- [ ] Integrate citation-style confidence and navigation eligibility into the main audit summary.
- [ ] Add clearer discipline/article-type semantic structure reporting.

## Word add-in

- [x] Office.js task pane and manifest exist.
- [x] Early Access sideloading guidance exists.
- [x] Non-mutating navigation foundation exists.
- [ ] Show all occurrences of a selected reference with Previous / Next / occurrence list.
- [ ] Make live citation navigation the preferred path when citation-manager fields must remain editable.
- [ ] Broader real Word testing.
- [ ] Microsoft Marketplace submission.

## Packaging and stable-release gates

- [x] Windows portable build.
- [x] Windows installer.
- [x] macOS package build.
- [x] Linux package build.
- [x] Cross-platform GitHub Release attachment workflow fixed and validated.
- [ ] Windows code signing.
- [ ] macOS notarization.
- [ ] 30-50 diverse real/sanitized manuscripts.
- [ ] Real EndNote, Zotero, and Mendeley validation.
- [ ] Cross-discipline journal/template corpus.
- [ ] No known critical Word integrity blocker.
- [ ] Release Candidate.
- [ ] Stable 1.0.
