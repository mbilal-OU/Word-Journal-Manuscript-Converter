# Template Coverage Expansion Roadmap

This document defines the next focused engineering phase for Word Journal Manuscript Converter after the Conversion Assurance and Word-openability foundation was validated in PR #7.

## Baseline

The tested real Word template opened safely after adaptation and reported:

- Supported fidelity: 100%
- Template coverage: 33%
- Verdict: LIMITED TEMPLATE ADAPTATION - MANUAL FORMATTING REQUIRED

That result is considered safe and truthful, but the coverage is too low for mature Template Mode behavior.

## Non-negotiable safety gates

Every template-coverage change must preserve all existing safety requirements:

1. Original manuscript is never overwritten.
2. Visible text and scientific numeric tokens remain unchanged unless a workflow explicitly permits otherwise.
3. Live citation-manager fields remain protected.
4. Tables, figures, equations, comments, footnotes, endnotes, tracked changes, relationships, custom XML, and content types remain protected by preservation checks.
5. Markup-compatibility namespaces and `mc:Ignorable` prefixes remain valid after every XML rewrite.
6. Generated DOCX output must pass defensive OOXML checks.
7. Any failed preservation or structural check must withhold the output.
8. Low template coverage must never be presented as a complete conversion.

## Phase 1: semantic style coverage

Target common manuscript roles and map them to template styles using both style definitions and template-body examples.

- article title
- subtitle
- author names
- affiliations
- corresponding-author block
- abstract heading and abstract body
- keyword heading and keyword body
- Heading 1 to Heading 4
- normal body text
- figure captions
- table captions
- references heading
- reference entries
- acknowledgments
- funding
- data availability
- author contributions
- competing interests

Required tests:

- role inference from style ID
- role inference from style name
- role inference from actual template paragraphs
- preservation of existing citation fields and visible text
- no Word repair after semantic remapping

## Phase 2: style dependency import

Import a custom paragraph style only when its dependencies can be resolved safely.

Dependencies to model:

- `basedOn`
- `next`
- linked character style
- paragraph properties
- run properties
- numbering references
- theme-font references

A style with unresolved or unsafe dependencies must remain unsupported and be reported explicitly.

## Phase 3: numbering definitions

Add relationship-safe transfer of numbering definitions for manuscript lists and headings.

Guardrails:

- never rewrite citation-manager numbering
- never treat bibliography numbering as a general list without verification
- copy only numbering definitions referenced by imported template styles
- remap `numId` and `abstractNumId` values when collisions occur
- verify list numbering after save

## Phase 4: headers, footers, and page numbering

Add safe section-level import for publisher header/footer conventions.

Requirements:

- copy only relationships required by the imported header/footer
- remap relationship IDs safely
- preserve existing manuscript relationships
- support page-number fields without flattening them
- allow user-facing coverage reporting when a header/footer is intentionally skipped

Do not import template placeholder text as manuscript content unless a future explicit placeholder-mapping workflow is used.

## Phase 5: tables and captions

Expand support for:

- table styles
- table alignment
- cell margins
- header-row formatting
- caption styles
- spacing before and after tables/captions

All table counts and visible cell text must remain preserved.

## Phase 6: section layout

Expand section adaptation for:

- first-page differences
- section breaks
- column transitions
- page numbering restart behavior
- vertical alignment
- text direction where relevant

Changes must remain compatible with existing equations, figures, fields, and references.

## Phase 7: content controls and placeholders

Inspect publisher templates for structured document tags and placeholder patterns.

Initial behavior:

- detect and report placeholders
- classify likely semantic roles
- do not move manuscript content automatically unless the mapping is high confidence and separately validated

Any automatic placeholder filling should be introduced only after a dedicated test corpus exists.

## Template coverage score

Coverage must be based on discovered template features, not only features the engine already knows how to transfer.

Recommended categories:

- page/section layout
- paragraph styles
- character styles
- numbering
- theme/fonts
- headers/footers
- tables
- captions
- content controls/placeholders
- page numbering

The report should continue to expose both:

- **Supported Fidelity**: how accurately transferred features match the template
- **Template Coverage**: how much of the relevant template formatting surface was actually handled

## Verdict thresholds

These remain descriptive, not publisher-compliance guarantees.

- 90-100% coverage with no blocking failures: HIGH TEMPLATE ADAPTATION
- 70-89%: PARTIAL TEMPLATE ADAPTATION
- 40-69%: LIMITED TEMPLATE ADAPTATION
- below 40%: LIMITED TEMPLATE ADAPTATION - MANUAL FORMATTING REQUIRED
- any blocking preservation/structural failure: OUTPUT WITHHELD

## Real-template regression corpus

Before Template Mode is considered release-candidate quality, build a sanitized corpus containing templates from multiple publishers and document types.

Minimum target:

- at least 10 distinct real publisher or conference templates before Release Candidate
- at least 20 before 1.0
- include single-column, two-column, author-year, numbered-reference, review, research-article, and conference layouts

For every template, record:

- template features detected
- supported fidelity
- coverage
- unsupported features
- Word-openability result
- preservation result
- visual-review result

## Merge criteria for this workstream

A Template Coverage Expansion PR should not be considered complete merely because synthetic CI passes.

It should demonstrate:

- full Python CI matrix green
- no namespace or OOXML regressions
- at least one real complex template with materially higher coverage than the 33% baseline
- no Word repair dialog
- no lost citation fields, figures, tables, equations, comments, notes, or tracked changes
- coverage and verdict remain honest when unsupported features remain
