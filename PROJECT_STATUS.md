# Project Status and Launch Checklist

Last reviewed: 2026-08-20

This is the working launch checklist for Word Journal Manuscript Converter. Status is intentionally conservative. Automated CI is necessary, but a feature is not considered release-ready until it also survives representative Microsoft Word and citation-manager testing.

## Status legend

- **Working**: implemented with meaningful automated or real-world evidence.
- **Partial**: implemented, but important cases or validation are still missing.
- **Blocked**: a known issue prevents release-quality use.
- **Planned**: not yet implemented.

Readiness percentages are engineering estimates. They are not compliance guarantees.

## Current critical blockers

| Area | Status | Readiness | Current evidence | Exit condition |
|---|---|---:|---|---|
| Microsoft Word openability after XML rewrite | Blocked | 45% | Two real manuscript tests produced Word's unreadable-content repair dialog even though generic XML and preservation checks passed | The same real EndNote manuscript opens without repair after journal conversion, template adaptation, and static review-copy creation |
| Markup-compatibility namespace preservation | In progress | 70% | Real-world failure is consistent with `mc:Ignorable` prefixes being lost by ElementTree serialization | Regression tests prove compatibility-only namespace declarations survive rewriting and generated DOCX opens cleanly in Word |
| Arbitrary Word template adaptation | Partial | 45% | A real template reported 100% supported fidelity but only 50% coverage and did not visually reproduce the supplied format | Real templates reach high verified coverage, or remaining unsupported surfaces are clearly reported as partial rather than full conversion |
| Template success messaging | In progress | 80% | 100% supported fidelity was easy to misread as full conversion when coverage was 50% | UI explicitly classifies high, partial, limited, or blocked adaptation and never calls low-coverage output a full conversion |
| Static linked review copy | Blocked | 65% | Real EndNote review-copy workflow previously triggered Word repair | Same real EndNote review copy opens with no repair and visible text remains identical |
| Word add-in installation | Partial | 60% | Early Access still uses Microsoft sideloading | Guided test setup plus Microsoft Marketplace distribution for stable release |
| Zotero and Mendeley real-world validation | Partial | 35% | Detection and synthetic regression exist, but real documents are limited | External tester documents pass navigation and integrity testing |
| Code signing and notarization | Planned | 10% | Windows and macOS packages are not yet signed/notarized | Signed Windows build and notarized macOS build |

## Core document safety

| Feature | Status | Readiness | Evidence or next check |
|---|---|---:|---|
| Original manuscript never overwritten | Working | 95% | New-copy workflows and same-path rejection |
| DOCX ZIP/package validation | Working | 90% | Required package-part checks |
| Visible text preservation | Working | 90% | Before/after fingerprint gate |
| Scientific numeric-token preservation | Working | 90% | Before/after token comparison |
| Word field-instruction preservation | Working | 90% | Field-instruction multiset comparison |
| Citation-field count preservation | Working | 90% | EndNote, Zotero, CSL candidate field counts |
| Embedded media preservation | Working | 90% | SHA-256 media comparison |
| Custom XML preservation | Working | 85% | Hash comparison |
| Relationship preservation | Working | 85% | Relationship hash comparison |
| Content types preservation | Working | 90% | Package hash comparison |
| Tracked changes preservation | Working | 85% | Insert/delete counts |
| Comments preservation | Working | 85% | comments.xml inventory |
| Footnote and endnote preservation | Working | 85% | Note-part inventory |
| Table preservation | Working | 90% | Table counts |
| Equation preservation | Partial | 80% | Native OMML and known embedded equation objects counted; equation-as-image remains indistinguishable from ordinary images |
| Known OOXML child-order validation | Partial | 85% | Paragraph, run, style, table, and section ordering checks exist |
| Markup-compatibility prefix validation | In progress | 70% | Added after repeated real Word repair failures exposed a gap not caught by generic XML parsing |
| Microsoft Word application-level openability | Blocked | 45% | Must be demonstrated in real desktop Word before merge |

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
| Conversion Assurance engine | Partial | 80% | Reopens and verifies saved output; Word application-level validation remains the current gate |
| Fail-closed conversion output | Working | 85% | Unsafe or suspect output is removed |
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
| Safe custom paragraph styles | In progress | 70% | Custom styles without numbering dependencies can be imported |
| Document defaults | Partial | 70% | Imported and verified; real-template testing still required |
| Template semantic role inference | In progress | 65% | Uses supplied template body and style names to infer title, headings, abstract, captions, keywords, and references |
| Theme transfer | Partial | 70% | Only self-contained theme part is copied automatically |
| Embedded font-table transfer | Blocked | 20% | fontTable can depend on relationships and embedded fonts; unsafe standalone copy is disabled |
| Template supported-fidelity score | Working | 85% | Measures only transferred and machine-verified formatting |
| Template coverage score | Working | 85% | Measures unsupported discovered template surfaces |
| High/partial/limited verdict | In progress | 85% | Low coverage is no longer presented as full conversion |
| Headers and footers | Planned | 25% | Relationship-safe merge required |
| Numbering definitions | Planned | 25% | Must avoid corrupting lists and citation numbering |
| Content-control placeholder mapping | Planned | 20% | Requires semantic and relationship-safe handling |
| Macros | Out of safe scope | 10% | Macro preservation/import needs a separate signed macro-aware path |
| Exact arbitrary-template reproduction | Partial | 35% | Not claimed. A future Word-assisted mode may be needed for complex publisher templates |

## Citation Navigator

| Feature | Status | Readiness | Evidence or next check |
|---|---|---:|---|
| EndNote field detection | Working | 90% | Real manuscript detected 113 live EndNote citation fields |
| Zotero field detection | Partial | 65% | Synthetic coverage; real tester documents needed |
| Mendeley/CSL detection | Partial | 65% | Synthetic coverage; real tester documents needed |
| Plain numbered citation graph | Working | 80% | Regression tests and local citation map |
| Simple `[N]` internal links | Working | 80% | Separate navigable copy |
| Grouped/range citation linking | Partial | 40% | Complex groups and ranges remain intentionally conservative |
| Static linked review copy | Blocked | 65% | Real Word openability still failing |
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
| Three primary workflows | Working | 85% | Journal Conversion, Citation Navigator, Manuscript Audit |
| Strong selected main-tab color | In progress | 90% | Selected workflow uses bright blue with white text |
| Separate journal-profile and Word-template sub-tabs | In progress | 90% | Only the selected conversion path is shown |
| Running action indicator | In progress | 85% | RUNNING/DONE/STOPPED state strip and progress indicator |
| Conversion assurance summary | Working | 80% | Compliance, integrity, blockers, verdict |
| Template fidelity and coverage summary | Working | 85% | Fidelity and coverage shown together |
| Low-coverage warning | In progress | 90% | Partial/limited template results use warning dialog rather than success dialog |
| Named live-citation choices | Working | 90% | Open Word add-in / Create static review copy / Cancel |
| Guided Word add-in setup | Partial | 65% | Local manifest finder and short Early Access instructions |
| Background document transformation | Planned | 45% | Heavy local transformations can still block Tk briefly |
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
| Release asset automation | Partial | 80% | Prior directory-glob upload failure fixed separately; keep regression-tested |
| Optional privacy-minimized analytics | Working | 80% | Opt-in and excludes manuscript content |
| Feedback submission | Working | 75% | Optional feedback path |
| Windows code signing | Planned | 10% | Launch gate |
| macOS notarization | Planned | 10% | Launch gate |

## Testing matrix

| Test area | Status | Readiness | Next target |
|---|---|---:|---|
| Python 3.10-3.13 CI | Working | 95% | Keep green on every PR |
| Synthetic DOCX regression suite | Working | 85% | Add namespace and complex OOXML fixtures |
| `mc:Ignorable` namespace regression | In progress | 80% | Must fail on dangling prefix and pass after safe rewrite |
| Real EndNote manuscripts | Partial | 65% | Repeat same manuscript after namespace-preservation patch |
| Real Zotero manuscripts | Planned | 25% | Recruit testers |
| Real Mendeley manuscripts | Planned | 25% | Recruit testers |
| Real publisher templates | Partial | 35% | Build a corpus across publishers/article types |
| Complex equations | Partial | 50% | OMML, MathType, Equation Editor, equation-as-image |
| Tracked changes/comments/notes | Partial | 60% | Expand real-world fixtures |
| Large review/thesis documents | Partial | 40% | Stress-test long manuscripts and many citations |
| 30-50 diverse real/sanitized manuscripts | Planned | 15% | Required before 1.0 |

## Release milestones

| Milestone | Status | Readiness | Definition |
|---|---|---:|---|
| Early Access | Working | 75% | Usable for structured testing with limitations clearly disclosed |
| Release Candidate | Blocked | 45% | Requires no known Word-repair/document-corruption blocker and broader real-document validation |
| 1.0 stable | Partial | 40% | Real-world validation complete, signing/notarization addressed, add-in distribution simplified, no known critical integrity bugs |

## PR #7 merge gate

Current PR: `feat/conversion-assurance`

PR #7 remains a draft. It is not merge-ready because repeated real Microsoft Word tests still produce an unreadable-content repair dialog and the supplied Word template produced only 50% coverage despite 100% supported fidelity.

PR #7 may move toward merge only after all of the following are true:

- [ ] Python 3.10, 3.11, 3.12, and 3.13 CI passes after the namespace-preservation patch.
- [ ] `mc:Ignorable` namespace regression tests pass.
- [ ] Same real EndNote manuscript opens after journal conversion with no Word repair dialog.
- [ ] Same real manuscript keeps all 113 live citation fields in the live master workflow.
- [ ] Static linked review copy opens with no Word repair dialog.
- [ ] Equation count is preserved before/after conversion.
- [ ] Tables, figures, comments, notes, tracked changes, fields, relationships, and custom XML remain preserved.
- [ ] Same real template is re-tested and the UI reports high, partial, limited, or blocked adaptation truthfully.
- [ ] A low-coverage template result is never presented as a complete conversion.
- [ ] Selected primary workflow tab is visually unmistakable on Windows.
- [ ] Journal Profile and Word Template are clearly separated in the interface.
- [ ] Word add-in setup path is understandable to a non-developer tester.

## Stable 1.0 launch gate

- [ ] No known Microsoft Word repair/corruption blocker.
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
