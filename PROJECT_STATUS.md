# Project Status and Launch Checklist

Last reviewed: 2026-08-20

This document is the working project checklist for Word Journal Manuscript Converter. It is intentionally conservative. A feature is not marked complete just because CI passes. Real Microsoft Word testing, real citation-manager documents, and real publisher templates are required for launch readiness.

## Status legend

- **Working**: implemented and has meaningful automated or real-world evidence.
- **Partial**: implemented, but important cases or validation are still missing.
- **Blocked**: a known issue prevents release-quality use.
- **Planned**: not yet implemented.

The percentage is a practical readiness estimate, not a mathematical guarantee.

## Current critical blockers

| Area | Status | Readiness | Current issue | Exit condition |
|---|---|---:|---|---|
| Word-openable converted DOCX | Blocked | 55% | A real retargeted manuscript triggered Word's unreadable-content repair dialog | Same real manuscript opens in desktop Word without repair after conversion |
| Arbitrary journal template adaptation | Partial | 55% | Safe Template Mode does not yet reproduce every custom style, header/footer, numbering system, placeholder, or direct-formatting pattern | Real publisher templates show verified high supported fidelity and clearly reported remaining coverage gaps |
| Static linked review copy | Partial | 70% | Bookmark-order fix exists, but real complex EndNote document needs another Word-openability retest | Real EndNote review copy opens without repair and visible text remains identical |
| Word add-in installation | Partial | 55% | Early Access uses Microsoft sideloading, which is too technical for normal users | Guided Early Access setup plus Microsoft Marketplace distribution for stable release |
| Zotero/Mendeley validation | Partial | 35% | Synthetic regression coverage is stronger than real-user validation | Real documents from external testers pass navigation and integrity tests |
| Code signing and notarization | Planned | 10% | Windows/macOS packages are not yet signed/notarized | Signed Windows release and notarized macOS release |

## Core document safety

| Feature | Status | Readiness | Evidence / next check |
|---|---|---:|---|
| Original manuscript never overwritten | Working | 95% | Output-path checks and copy-based workflows |
| DOCX ZIP/package validation | Working | 90% | Required package-part checks |
| Visible text preservation | Working | 90% | Before/after fingerprint gate |
| Scientific numeric-token preservation | Working | 90% | Before/after numeric token comparison |
| Word field-instruction preservation | Working | 90% | Field-instruction multiset comparison |
| Citation-field count preservation | Working | 90% | EndNote/Zotero/CSL candidate field counts |
| Embedded media preservation | Working | 90% | SHA-256 media comparison |
| Custom XML preservation | Working | 85% | Hash comparison |
| Relationship preservation | Working | 85% | Relationship hash comparison |
| Tracked changes preservation | Working | 85% | Insert/delete counts |
| Comments and notes preservation | Working | 85% | Comment, footnote, endnote counts |
| Table preservation | Working | 90% | Table counts |
| Equation preservation | Partial | 80% | Native OMML and known embedded equation objects are counted; equations inserted only as ordinary pictures cannot yet be distinguished from figures |
| Defensive OOXML structure validation | Partial | 80% | PR #7 adds known child-order validation for edited containers; real Word retest still required |
| Microsoft Word openability gate | Partial | 60% | Defensive XML validation exists; real Word application-level validation remains a required manual release gate |

## Journal profile conversion

| Feature | Status | Readiness | Evidence / next check |
|---|---|---:|---|
| 30+ bundled journal profiles | Working | 85% | Bundled profile regression tests |
| Source-dated profile metadata | Working | 80% | Official source URLs and checked dates are stored where available |
| Page margins | Working | 85% | Transformation plus independent post-save verification |
| Line numbering | Working | 85% | Transformation plus independent verification |
| Body font and size | Partial | 75% | Normal style can be retargeted; direct formatting can override it |
| Line spacing | Partial | 75% | Normal style can be retargeted; direct paragraph formatting can override it |
| Required-section checks | Working | 80% | Readiness engine |
| Abstract length checks | Working | 85% | Readiness engine |
| Keyword checks | Working | 85% | Readiness engine |
| Citation/reference resolution checks | Partial | 75% | Strongest for numbered citations; complex and author-year cases need broader coverage |
| Conversion Assurance engine | Partial | 80% | PR #7 independently reopens and verifies output; real Word test is the current gate |
| Fail-closed conversion output | Working | 85% | Unsafe or structurally suspect output is removed rather than presented as successful |
| No misleading universal-compliance claim | Working | 95% | Product wording requires final author review and current journal guidance check |

## Journal Template Mode

| Feature | Status | Readiness | Evidence / next check |
|---|---|---:|---|
| `.docx` and `.dotx` template inspection | Working | 85% | Template inventory and validation |
| Page size and orientation transfer | Working | 80% | Machine-verifiable fidelity check |
| Margin transfer | Working | 80% | Machine-verifiable fidelity check |
| Column transfer | Working | 80% | Machine-verifiable fidelity check |
| Line-numbering transfer | Working | 80% | Machine-verifiable fidelity check |
| Matching paragraph-style transfer | Partial | 75% | PR #7 expands beyond a small fixed style list and matches by ID/name |
| Document-default transfer | Partial | 70% | Added in PR #7; needs real template testing |
| Compatible theme/font-part transfer | Partial | 65% | Added when both packages contain compatible target parts |
| Semantic paragraph-style bridge | Partial | 60% | Added for common manuscript roles; custom publisher semantics still vary |
| Template supported-fidelity score | Working | 80% | Independently compares transferred formatting |
| Template coverage score | Working | 80% | Explicitly counts unsupported template surfaces |
| Headers and footers | Planned | 25% | Relationship-safe import design required |
| Numbering definitions | Planned | 25% | Must avoid corrupting existing lists/citation numbering |
| Custom placeholder/content-control mapping | Planned | 20% | Requires safe semantic mapping |
| Exact arbitrary-template reproduction | Partial | 35% | Not claimed. Some publisher templates will still require manual review or a future Word-assisted mode |

## Citation Navigator

| Feature | Status | Readiness | Evidence / next check |
|---|---|---:|---|
| EndNote field detection | Working | 90% | Real manuscript detected 113 live EndNote citation fields |
| Zotero field detection | Partial | 65% | Synthetic coverage; real beta documents needed |
| Mendeley/CSL detection | Partial | 65% | Synthetic coverage; real beta documents needed |
| Plain numbered citation graph | Working | 80% | Regression tests and clickable map |
| Simple `[N]` internal links | Working | 80% | Separate navigable copy |
| Grouped/range citation linking | Partial | 40% | Complex groups/ranges are intentionally left unchanged in current release |
| Static linked review copy | Partial | 70% | Real EndNote test previously created 41 links and 88 reference bookmarks; Word-openability retest required after structural fixes |
| Live citation fields protected | Working | 90% | Desktop app refuses to overlay plain links on live citation-manager fields |
| Word add-in live navigation | Partial | 70% | Hosted task pane and manifest exist; Early Access sideloading remains inconvenient |
| Named citation-manager choice dialog | In progress | 75% | PR #7 UI work replaces ambiguous Yes/No wording with explicit actions |

## Manuscript Audit

| Feature | Status | Readiness | Evidence / next check |
|---|---|---:|---|
| Paragraph count | Working | 95% | Structural inventory |
| Section count | Working | 90% | Structural inventory |
| Table count | Working | 95% | Structural inventory |
| Embedded media count | Working | 90% | DOCX media inventory |
| Equation count | Partial | 80% | OMML and known equation-editor objects supported |
| Comments | Working | 90% | comments.xml inventory |
| Footnotes/endnotes | Working | 90% | Word note-part inventory |
| Tracked insertions/deletions | Working | 90% | Review markup inventory |
| Word fields | Working | 90% | Field-instruction inventory |
| Hyperlinks/bookmarks | Working | 85% | Structural inventory |

## Desktop user interface

| Feature | Status | Readiness | Evidence / next check |
|---|---|---:|---|
| Three-workflow navigation | Working | 85% | Journal Conversion, Citation Navigator, Manuscript Audit |
| Active-operation highlighting | In progress | 75% | PR #7 UI work adds explicit active-task state and selected action styling |
| Clear conversion vs template mode separation | In progress | 75% | PR #7 UI refinement |
| Conversion assurance summary | In progress | 75% | Scores/verdict made visible in Summary view |
| Template fidelity summary | In progress | 75% | Fidelity and coverage surfaced in Summary view |
| Named live-citation actions | In progress | 75% | Replaces ambiguous Yes/No dialog |
| Guided Word add-in setup | In progress | 60% | Early Access helper explains easiest route and local manifest location |
| Background/non-blocking heavy operations | Planned | 45% | Some network tasks are threaded; document transformations can still block the Tk event loop briefly |
| Accessibility/keyboard review | Planned | 30% | Dedicated pass required |

## Word add-in

| Feature | Status | Readiness | Evidence / next check |
|---|---|---:|---|
| Office.js task pane | Working | 80% | Hosted HTTPS task pane |
| Manifest | Working | 80% | Included in release package and installer |
| Citation/reference scan | Partial | 75% | Needs broader real-document validation |
| Non-mutating navigation | Partial | 75% | Intended to change Word selection rather than citation payloads |
| Integrity check | Partial | 70% | Lightweight in-Word check |
| Early Access installation | Partial | 55% | Microsoft sideloading is intentionally a testing path, not final UX |
| Guided setup from desktop app | In progress | 60% | PR #7 UX work |
| Microsoft Marketplace publication | Planned | 10% | Requires Partner Center submission and Microsoft review |
| Normal one-click user installation | Planned | 10% | Target stable UX through Marketplace |

## Packaging, updates, privacy, and release engineering

| Feature | Status | Readiness | Evidence / next check |
|---|---|---:|---|
| Windows portable build | Working | 90% | GitHub Actions artifact |
| Windows installer | Working | 85% | Inno Setup build succeeds |
| macOS package | Working | 70% | Build works; notarization remains |
| Linux package | Working | 80% | Build artifact works through CI smoke path |
| App icon | Working | 90% | Validated Windows/macOS release icons |
| In-app update check | Working | 80% | GitHub Releases based |
| Release asset automation | Partial | 80% | Build pipeline works; release upload path has required troubleshooting history and should remain regression-tested |
| Optional privacy-minimized analytics | Working | 80% | Opt-in, excludes manuscript content |
| Feedback submission | Working | 75% | Optional user feedback path |
| Windows code signing | Planned | 10% | Launch gate |
| macOS notarization | Planned | 10% | Launch gate |

## Testing matrix

| Test area | Status | Readiness | Next target |
|---|---|---:|---|
| Python 3.10-3.13 CI | Working | 95% | Keep green on every PR |
| Synthetic DOCX regression suite | Working | 85% | Add more malformed/complex OOXML fixtures |
| Real EndNote manuscripts | Partial | 70% | Retest linked review copy and journal conversion after PR #7 structural patch |
| Real Zotero manuscripts | Planned | 25% | Recruit external testers |
| Real Mendeley manuscripts | Planned | 25% | Recruit external testers |
| Real publisher templates | Partial | 35% | Build a template corpus across publishers and article types |
| Complex equations | Partial | 50% | Test OMML, MathType, Equation Editor, equation-as-image cases |
| Tracked changes/comments/notes | Partial | 60% | Expand real-world fixtures |
| Large review/thesis documents | Partial | 40% | Stress test long manuscripts and many citations |
| 30-50 diverse real/sanitized manuscripts | Planned | 15% | Required before 1.0 |

## Release milestones

| Milestone | Status | Readiness | Definition |
|---|---|---:|---|
| Early Access | Working | 80% | Usable for structured testing, with limitations clearly disclosed |
| Release Candidate | Partial | 55% | No known document-corruption blockers; broad real-document validation; installer/update path stable |
| 1.0 stable | Partial | 45% | Real-world validation complete, signing/notarization addressed, add-in distribution simplified, no known critical integrity bugs |

## PR #7 gate

Current PR: `feat/conversion-assurance`

Initial CI passed on Python 3.10, 3.11, 3.12, and 3.13. A subsequent real-world test exposed a Microsoft Word unreadable-content warning in a retargeted manuscript and insufficient visual adaptation for the supplied template. Therefore PR #7 remains a draft and is not merge-ready.

PR #7 may move toward merge only after all of the following are true:

- [ ] Full CI matrix passes after the structural-order patch.
- [ ] Same real EndNote manuscript opens after journal conversion with no Word repair dialog.
- [ ] Same real manuscript keeps all 113 live citation fields in the live master workflow.
- [ ] Equations are detected and preserved according to the expanded inventory.
- [ ] Tables and embedded media remain unchanged.
- [ ] Linked review copy opens without Word repair.
- [ ] Template Mode reports Supported Fidelity and Template Coverage clearly.
- [ ] The supplied real template shows a meaningful visual adaptation where supported.
- [ ] Unsupported template features are explicitly listed instead of being silently ignored.
- [ ] GUI clearly shows which workflow/action is selected and running.
- [ ] Live-citation dialog uses named actions instead of Yes/No semantics.
- [ ] Word add-in Early Access setup is explained as sideloading, with an easier guided path.

## Next engineering priorities

1. Eliminate Word repair dialogs in every generated DOCX path.
2. Retest the user's real EndNote manuscript and real template.
3. Expand template semantic adaptation without weakening preservation guarantees.
4. Simplify Word add-in setup for Early Access testers.
5. Build a real template/manuscript validation corpus.
6. Recruit Zotero and Mendeley testers.
7. Prepare code signing, notarization, and Microsoft Marketplace submission for stable launch.
