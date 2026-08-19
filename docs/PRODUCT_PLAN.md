# Word Journal Manuscript Converter Product Plan

## 1. Product definition

Word Journal Manuscript Converter is a **Word-native, integrity-preserving manuscript retargeting system**. Its job is to help a researcher move an existing `.docx` manuscript toward a new journal workflow while proving what changed and what did not.

The primary differentiator is not automatic formatting. It is the combination of:

- safe handling of real Word document structures;
- citation/reference integrity;
- journal-readiness checking;
- conservative, auditable transformations;
- local-first privacy for unpublished research.

## 2. Non-negotiable guarantees

### Research privacy

Normal inspection, auditing, and local transformations must not require uploading the manuscript. Public journal metadata may be retrieved separately without sending the manuscript body.

### Scientific content preservation

Scientific prose, values, equations, figures, and tables are immutable unless a user explicitly requests a content edit outside the retargeting workflow.

### Citation/reference integrity

- preserve live EndNote/Zotero/Mendeley/Word citation fields when present;
- do not flatten citation-manager fields by default;
- do not renumber one side of a citation/reference relationship independently;
- audit orphan citations and uncited references;
- support clickable in-text citation -> reference navigation where the source or target workflow allows it;
- preserve existing bookmarks and internal hyperlinks.

### Auditability

Every mutation must produce a machine-readable transformation manifest and a human-readable report.

## 3. Competitive position

Word Journal Manuscript Converter should not compete as another generic "format my paper" service. Existing products already automate journal templates, export styles, and submission checks.

The defensible niche is:

**existing Word manuscript + Word-native feature preservation + safe retargeting + citation integrity + proof of preservation**.

## 4. Technical architecture

```text
DOCX package
   |
   v
Package validator / inventory
   |
   v
Protected feature map
   |-- citation-manager fields
   |-- hyperlinks/bookmarks
   |-- equations
   |-- figures/media
   |-- tables
   |-- comments/tracked changes
   |-- footnotes/endnotes
   |-- custom XML
   |
   v
Manuscript structural map
   |
   +------ target journal profile
   |          |
   |          v
   |     requirements engine
   |          |
   +----------+
   |
   v
Transformation planner
   |
   +-- automatic safe changes
   +-- manual actions for unsafe/ambiguous cases
   |
   v
Surgical OOXML/style mutation
   |
   v
Preservation verifier
   |
   v
Retargeted DOCX + compliance report + transformation manifest
```

## 5. Journal intelligence model

A journal rule should never be an untraceable hard-coded statement. Profiles should contain:

- journal name;
- publisher;
- article type;
- source URL;
- date checked;
- requirement category;
- hard vs soft rule;
- whether the rule is format-free at initial submission;
- whether it is auto-fixable;
- validation logic;
- confidence/notes when interpretation is ambiguous.

A journal may have different requirements for original research, reviews, brief reports, methods papers, etc., so profiles must be article-type-specific.

## 6. Citation system plan

### Phase A: identify

Detect:
- EndNote Cite While You Write fields;
- Zotero fields;
- Mendeley/CSL fields;
- native Word citation/bibliography fields;
- plain-text author-year citations;
- plain-text numbered citations;
- bookmarks, REF/PAGEREF, and internal hyperlinks.

### Phase B: map

Build a citation graph:

```text
in-text occurrence -> citation object -> bibliography entry
```

Each edge receives a source and confidence level.

### Phase C: preserve

If the document uses a live citation manager, preserve that manager's field representation and let the manager remain authoritative for style regeneration whenever possible.

### Phase D: navigation

For plain-text manuscripts, optionally add Word bookmarks and internal hyperlinks so in-text citations can jump to reference entries. This is a separate, explicit transformation and must not alter citation text.

## 7. MVP transformation matrix

The initial safe transformation set should be narrow:

- margins and page size;
- line numbering;
- page numbering;
- paragraph spacing;
- heading style mapping;
- font/style normalization where it does not touch field internals;
- title-page separation;
- declaration section ordering;
- anonymized review copy generation;
- required-section placeholders;
- table/figure caption style mapping;
- internal citation/reference linking for plain-text references only after validated matching.

Do not initially automate:

- citation-manager conversion between EndNote/Zotero/Mendeley;
- destructive bibliography rewriting;
- complex multi-column reconstruction;
- equation conversion;
- tracked-change acceptance;
- reference metadata correction without explicit user approval.

## 8. User experience

### Desktop workflow

1. Drop a DOCX.
2. Word Journal Manuscript Converter shows a privacy indicator: `Local processing - manuscript not uploaded`.
3. Integrity scan summarizes citation manager, fields, figures, tables, comments, equations, and review markup.
4. User chooses a journal or imports a profile/template.
5. Readiness screen separates:
   - already compliant;
   - safe automatic fixes;
   - manual decisions;
   - format-free items that do not need changing.
6. User previews proposed changes.
7. Word Journal Manuscript Converter writes a **new copy**, never overwriting the source by default.
8. Preservation audit runs automatically.
9. User gets the new DOCX and a concise audit report.

## 9. Accessibility/distribution

### CLI

Best for reproducibility and early validation.

### Desktop app - recommended primary product

Best balance of usability, privacy, and access to full DOCX files. A Tauri shell with a local core service is the preferred polished architecture; PySide6 is a faster prototype option.

### Word add-in

Best as an in-Word task pane for readiness checks, jump-to-problem navigation, and controlled fixes. Package-level transformations should still be delegated to the local Word Journal Manuscript Converter integrity engine when Office.js does not expose enough fidelity.

### Web/PWA

Useful later for reach. Prefer client-side processing or a local companion. If cloud processing is added, it must be an explicit mode with clear retention and data-handling disclosures.

### Python API

Supports institutional workflows, reproducible pipelines, and future publisher integrations.

## 10. Release sequence

### v0.1 - audit foundation
Current repository state.

### v0.2 - citation graph + manuscript structure
Build authoritative detection/mapping and plain-text reference matching.

### v0.3 - safe transformations
Implement only operations covered by preservation tests.

### v0.4 - desktop alpha
Researcher-friendly drag-and-drop local application.

### v0.5 - Word add-in beta
Readiness panel and controlled fixes inside Word.

### v0.6 - journal profile registry
Community-maintained, source-linked, dated profiles with automated staleness checks.

### v1.0 - validated public release
Signed desktop apps, transformation compatibility matrix, extensive synthetic/de-identified regression corpus, and stable extension APIs.

## 11. Success criteria

Word Journal Manuscript Converter should be judged by preservation and trust, not only by how much formatting it automates.

Key metrics:

- protected-field preservation rate;
- citation/reference integrity rate;
- numeric-content drift rate (target: zero for formatting transformations);
- media hash preservation rate;
- percentage of proposed changes classified as safe/manual with correct rationale;
- false-positive/false-negative rate of journal compliance checks;
- percentage of operations that can be reproduced from the transformation manifest.
