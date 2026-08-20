# Project Status and Launch Checklist

Last reviewed: 2026-08-20

This is the working launch checklist for Word Journal Manuscript Converter. Status is intentionally conservative. Automated CI is necessary, but release-quality claims also require representative Microsoft Word, citation-manager, and real-template testing.

## Status legend

- **Working**: implemented with meaningful automated or real-world evidence.
- **Partial**: implemented, but important cases or validation are still missing.
- **Blocked**: a known issue prevents release-quality use.
- **Planned**: not yet implemented.

Readiness percentages are engineering estimates. They are not publisher-compliance guarantees.

## Current critical items

| Area | Status | Readiness | Current evidence | Next exit condition |
|---|---|---:|---|---|
| Microsoft Word openability after XML rewrite | Working | 85% | Same real EndNote manuscript now opens without repair after journal conversion, template adaptation, and static linked review-copy creation | Repeat across a broader real-document corpus |
| Markup-compatibility namespace preservation | Working | 90% | Namespace-preservation regression passes and the real repair failure no longer reproduces on the tested manuscript | Keep regression green across future XML transforms |
| Arbitrary Word template adaptation | Partial | 45% | Real template opened safely but produced 100% supported fidelity with only 33% coverage and a LIMITED verdict | Increase safe coverage across real publisher templates without weakening Word openability |
| Template success messaging | Working | 95% | Low coverage is now reported as partial/limited rather than a complete conversion | Keep coverage and fidelity visibly separate |
| Static linked review copy | Working | 80% | Fresh real EndNote review copy opens in Word without repair | Broaden grouped/range and multi-manager testing |
| Word add-in installation | Partial | 60% | Guided Early Access setup exists, but Microsoft sideloading is still technical | Microsoft Marketplace distribution or equivalent supported deployment |
| Zotero and Mendeley real-world validation | Partial | 35% | Detection and synthetic regression exist, but real-document validation is limited | External tester documents pass navigation and integrity testing |
| Code signing and notarization | Planned | 10% | Windows and macOS packages are not yet signed/notarized | Signed Windows build and notarized macOS build |

## Core document safety

| Feature | Status | Readiness | Evidence or next check |
|---|---|---:|---|
| Original manuscript never overwritten | Working | 95% | New-copy workflows and same-path rejection |
| DOCX ZIP/package validation | Working | 90% | Required package-part checks |
| Visible text preservation | Working | 90% | Before/after fingerprint gate |
| Scientific numeric-token preservation | Working | 90% | Before/after token comparison |
| Word field-instruction preservation | Working | 90% | Field-instruction multiset comparison |
| Citation-field count preservation | Working | 90% | EndNote, Zotero, and CSL candidate-field counts |
| Embedded media preservation | Working | 90% | SHA-256 media comparison |
| Custom XML preservation | Working | 85% | Hash comparison |
| Relationship preservation | Working | 85% | Relationship hash comparison |
| Content types preservation | Working | 90% | Package hash comparison |
| Tracked changes preservation | Working | 85% | Insert/delete counts |
| Comments preservation | Working | 85% | comments.xml inventory |
| Footnote and endnote preservation | Working | 85% | Note-part inventory |
| Table preservation | Working | 90% | Table counts |
| Equation preservation | Partial | 80% | Native OMML and known embedded equation objects are counted; equation-as-image cannot yet be distinguished from an ordinary image |
| Known OOXML child-order validation | Working | 90% | Paragraph, run, style, table, and section ordering checks plus real Word retest |
| Markup-compatibility prefix validation | Working | 90% | Added after real Word repair failures and validated against the tested manuscript |
| Microsoft Word application-level openability | Working | 85% | Three generated-output paths now open cleanly for the tested real manuscript |

## Journal profile conversion

| Feature | Status | Readiness | Evidence or next check |
|---|---|---:|---|
| 30+ bundled journal profiles | Working | 85% | Bundled profile regression tests |
| Source-dated profile metadata | Working | 80% | Official source URLs and checked dates stored where available |
| Page margins | Working | 85% | Transform plus post-save verification |
| Line numbering | Working | 85% | Transform plus post-save verification |
| Body font and size | Partial | 75% | Normal style can be retargeted; direct formatting can override it |
| Line spacing | Partial | 75% | Normal style can be retargeted; direct paragraph formatting can override it |
| Required-section checks | Working | 80% | Readiness engine |
| Abstract length checks | Working | 85% | Readiness engine |
| Keyword checks | Working | 85% | Readiness engine |
| Citation/reference resolution checks | Partial | 75% | Strongest for numbered citations; complex and author-year cases need broader coverage |
| Conversion Assurance engine | Working | 85% | Reopens and verifies saved output; tested real conversion reported 100% formatting compliance and 83/100 manuscript readiness |
| Fail-closed conversion output | Working | 90% | Unsafe or structurally suspect output is removed |
| No universal publisher-compliance claim | Working | 95% | Final author review and current journal guidance remain required |
| UI text avoids long dash separators | Working | 95% | Journal labels and workflow text use plain separators |

## Word Template Mode

| Feature | Status | Readiness | Evidence or next check |
|---|---|---:|---|
| `.docx` and `.dotx` template inspection | Working | 85% | Template package inventory |
| Page size and orientation | Working | 80% | Machine-verifiable comparison |
| Margins | Working | 80% | Machine-verifiable comparison |
| Columns | Working | 80% | Machine-verifiable comparison |
| Line numbering | Working | 80% | Machine-verifiable comparison |
| Matching paragraph styles | Partial | 80% | Match by ID/name |
| Safe custom paragraph styles | Partial | 70% | Custom styles without unsafe dependencies can be imported |
| Document defaults | Partial | 75% | Imported and verified; more real templates needed |
| Template semantic role inference | Partial | 65% | Uses template body and style names to infer title, headings, abstract, captions, keywords, and references |
| Theme transfer | Partial | 70% | Only self-contained theme parts are copied automatically |
| Embedded font-table transfer | Blocked | 20% | fontTable can depend on relationships and embedded fonts; unsafe standalone copy is disabled |
| Template supported-fidelity score | Working | 90% | Tested template reported 100% fidelity for the subset actually transferred |
| Template coverage score | Working | 90% | Tested template reported 33% coverage, correctly exposing the remaining gap |
| High/partial/limited verdict | Working | 95% | Low coverage is not presented as a full conversion |
| Headers and footers | Planned | 25% | Relationship-safe merge required |
| Numbering definitions | Planned | 25% | Must avoid corrupting lists and citation numbering |
| Content-control placeholder mapping | Planned | 20% | Requires semantic and relationship-safe handling |
| Table-style transfer | Planned | 25% | Needs safe style and relationship handling |
| Page-number/header conventions | Planned | 20% | Depends on relationship-safe header/footer support |
| Macros | Out of safe scope | 10% | Macro import needs a separate signed macro-aware path |
| Exact arbitrary-template reproduction | Partial | 35% | Not claimed. Complex publisher templates may require a future Word-assisted mode |

## Citation Navigator

| Feature | Status | Readiness | Evidence or next check |
|---|---|---:|---|
| EndNote field detection | Working | 90% | Real manuscript detected 113 live EndNote citation fields |
| Zotero field detection | Partial | 65% | Synthetic coverage; real tester documents needed |
| Mendeley/CSL detection | Partial | 65% | Synthetic coverage; real tester documents needed |
| Plain numbered citation graph | Working | 80% | Regression tests and local citation map |
| Simple `[N]` internal links | Working | 80% | Separate navigable copy |
| Grouped/range citation linking | Partial | 40% | Complex groups and ranges remain intentionally conservative |
| Static linked review copy | Working | 80% | Fresh real EndNote review copy opens without Word repair |
| Live citation fields protected | Working | 90% | App refuses to overlay plain links on live manager fields |
| Explicit citation-manager action dialog | Working | 90% | Named buttons replace ambiguous Yes/No workflow |
| Word add-in live navigation | Partial | 70% | Hosted task pane and manifest exist; broader real-document validation needed |

## Manuscript Audit

| Feature | Status | Readiness | Evidence or next check |
|---|---|---:|---|
| Paragraph count | Working | 95% | Structural inventory |
| Section count | Working | 90% | Structural inventory |
| Table count | Working | 95% | Structural inventory |
| Embedded media count | Working | 90% | DOCX media inventory |
| Native Word equation count | Partial | 85% | OMML detection |
| Embedded equation object count | Partial | 75% | Known Equation Editor/MathType-style objects |
| Comments | Working | 90% | comments.xml inventory |
| Footnotes/endnotes | Working | 90% | Note-part inventory |
| Tracked insertions/deletions | Working | 90% | Review markup inventory |
| Word fields | Working | 90% | Field-instruction inventory |
| Hyperlinks/bookmarks | Working | 85% | Structural inventory |

## Desktop interface

| Feature | Status | Readiness | Evidence or next check |
|---|---|---:|---|
| Three primary workflows | Working | 90% | Journal Conversion, Citation Navigator, Manuscript Audit |
| Strong selected main-tab color | Working | 95% | Selected workflow uses bright blue with white text |
| Separate Journal Profile and Word Template sub-tabs | Working | 95% | Conversion paths are visually separated |
| Running action indicator | Working | 90% | RUNNING/DONE/STOPPED state strip and progress indicator |
| Conversion assurance summary | Working | 85% | Compliance, integrity, blockers, verdict |
| Template fidelity and coverage summary | Working | 90% | Fidelity and coverage shown together |
| Low-coverage warning | Working | 95% | 33% coverage produced LIMITED TEMPLATE ADAPTATION and manual-formatting warning |
| Named live-citation choices | Working | 90% | Open Word add-in / Create static review copy / Cancel |
| Guided Word add-in setup | Partial | 65% | Local manifest finder and short Early Access instructions |
| Background document transformation | Planned | 45% | Heavy local transforms can still block Tk briefly |
| Accessibility and keyboard pass | Planned | 30% | Dedicated review required |

## Word add-in

| Feature | Status | Readiness | Evidence or next check |
|---|---|---:|---|
| Office.js task pane | Working | 80% | Hosted HTTPS task pane |
| Manifest | Working | 80% | Included in portable package and installer |
| Citation/reference scan | Partial | 75% | Needs broader real-document validation |
| Non-mutating navigation | Partial | 75% | Intended to move Word selection, not rewrite manager payloads |
| Integrity check | Partial | 70% | Lightweight in-Word check |
| Early Access sideloading | Partial | 60% | Microsoft testing route is inherently more technical than store installation |
| Desktop guided setup | Partial | 65% | Explains fastest route and locates manifest |
| Microsoft Marketplace submission | Planned | 10% | Requires Partner Center submission and Microsoft review |
| Normal one-click Word installation | Planned | 10% | Stable target through Marketplace |

## Packaging, privacy, and release engineering

| Feature | Status | Readiness | Evidence or next check |
|---|---|---:|---|
| Windows portable build | Working | 90% | GitHub Actions artifact |
| Windows installer | Working | 85% | Inno Setup build |
| macOS package | Working | 70% | Build works; notarization remains |
| Linux package | Working | 80% | CI smoke path |
| App icon | Working | 90% | Validated Windows/macOS icons |
| In-app update check | Working | 80% | GitHub Releases based |
| Release asset upload | In progress | 85% | File-only upload logic replaces the directory-glob failure; official post-merge manual release run must validate it |
| GitHub Actions Node runtime migration | Working | 95% | CI passed using checkout@v7 and setup-python@v7; release workflow uses Node 24-capable action majors and awaits one post-merge release run |
| Optional privacy-minimized analytics | Working | 80% | Opt-in and excludes manuscript content |
| Feedback submission | Working | 75% | Optional feedback path |
| Windows code signing | Planned | 10% | Launch gate |
| macOS notarization | Planned | 10% | Launch gate |

## Testing matrix

| Test area | Status | Readiness | Next target |
|---|---|---:|---|
| Python 3.10-3.13 CI | Working | 95% | Keep green on every PR |
| Synthetic DOCX regression suite | Working | 85% | Continue adding complex OOXML fixtures |
| `mc:Ignorable` namespace regression | Working | 90% | Keep green across all future XML rewrites |
| Real EndNote manuscripts | Partial | 75% | One real manuscript now passes all three output-openability paths; broaden corpus |
| Real Zotero manuscripts | Planned | 25% | Recruit testers |
| Real Mendeley manuscripts | Planned | 25% | Recruit testers |
| Real publisher templates | Partial | 35% | Tested template is safe but only 33% covered; build a cross-publisher corpus |
| Complex equations | Partial | 50% | OMML, MathType, Equation Editor, equation-as-image |
| Tracked changes/comments/notes | Partial | 60% | Expand real-world fixtures |
| Large review/thesis documents | Partial | 40% | Stress-test long manuscripts and many citations |
| 30-50 diverse real/sanitized manuscripts | Planned | 15% | Required before 1.0 |

## Release milestones

| Milestone | Status | Readiness | Definition |
|---|---|---:|---|
| Early Access | Working | 85% | Core safety and real EndNote openability test pass; limitations are clearly disclosed |
| Release Candidate | Partial | 55% | Needs broader document/template validation and stable signed packaging |
| 1.0 stable | Partial | 45% | Real-world validation complete, signing/notarization addressed, add-in distribution simplified, no known critical integrity bugs |

## PR #7 merge gate

Current PR: `feat/conversion-assurance`

The original real-world Word repair blocker has been resolved for the tested manuscript. The same fresh build now opens all three generated-document paths without a repair dialog:

- journal profile conversion: PASS
- Word template adaptation: PASS for Word openability
- static linked review copy: PASS

The tested journal conversion reported 100% formatting compliance and 83/100 manuscript readiness. The tested Word template reported 100% supported fidelity and 33% coverage and was correctly classified as LIMITED TEMPLATE ADAPTATION with manual formatting required.

Merge gate for the safety/assurance foundation:

- [x] Python 3.10, 3.11, 3.12, and 3.13 CI passed after namespace-preservation changes.
- [x] `mc:Ignorable` namespace regression passed.
- [x] Same real manuscript opens after journal conversion with no Word repair dialog.
- [x] Same real manuscript opens after Word template adaptation with no Word repair dialog.
- [x] Static linked review copy opens with no Word repair dialog.
- [x] Low template coverage is no longer presented as a complete conversion.
- [x] Selected primary workflow tab is visually unmistakable on Windows.
- [x] Journal Profile and Word Template are clearly separated in the interface.
- [x] Live-citation dialog uses named actions rather than Yes/No semantics.
- [x] Word add-in Early Access setup is explained as sideloading with a guided path.
- [x] Fresh CI passed after final release-workflow/runtime cleanup.
- [ ] Official Early Access release build from merged `main` succeeds and attaches release files correctly.

Template coverage remains a separate product-quality workstream and does not weaken the completed Word-openability fix.

## Next focused engineering phase: Template Coverage Expansion

Priorities:

1. Increase safe custom-style import and semantic role mapping.
2. Add relationship-safe header and footer transfer.
3. Add numbering-definition transfer without altering citation numbering.
4. Add table-style and caption-style transfer.
5. Improve title, author, affiliation, abstract, keyword, heading, and reference mapping.
6. Add page-number and section-break conventions where safely supported.
7. Build a real publisher-template regression corpus.
8. Keep Word openability, namespace preservation, field preservation, and fail-closed output as non-negotiable gates.

## Stable 1.0 launch gate

- [ ] No known Microsoft Word repair/corruption blocker across a broad test corpus.
- [ ] 30-50 diverse real or sanitized manuscripts tested.
- [ ] Real EndNote, Zotero, and Mendeley coverage.
- [ ] Multiple real publisher templates tested.
- [ ] Complex equations, tables, figures, tracked changes, comments, footnotes, and endnotes covered.
- [ ] Windows installer signed.
- [ ] macOS package notarized.
- [ ] Privacy and analytics final review complete.
- [ ] Update path verified from an installed prior build.
- [ ] Word add-in Marketplace submission completed or a clearly supported distribution alternative approved.
- [ ] Documentation matches actual supported behavior with no universal compliance or exact-template guarantee.
