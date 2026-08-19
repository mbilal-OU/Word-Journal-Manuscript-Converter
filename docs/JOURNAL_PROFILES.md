# Journal profiles

A journal profile is machine-readable, source-dated evidence for manuscript rules that Word Journal Manuscript Converter can evaluate safely. It is not a permanent replacement for a journal's official author instructions.

## v0.5.0 catalog

The v0.5.0 catalog contains 30 journal-specific profiles plus the generic review-copy demonstration profile.

Coverage includes:

- PLOS ONE, PLOS Biology, PLOS Genetics, PLOS Computational Biology, PLOS Pathogens, PLOS Neglected Tropical Diseases, PLOS Medicine
- Scientific Reports, Nature Communications, Nature Microbiology
- Nucleic Acids Research, Bioinformatics
- Microbial Genomics, Microbiology, Journal of General Virology, Journal of Medical Microbiology, International Journal of Systematic and Evolutionary Microbiology
- mBio, Applied and Environmental Microbiology, mSphere
- Frontiers in Microbiology, Genetics, Bioinformatics, Cellular and Infection Microbiology, Ecology and Evolution, Plant Science, Molecular Biosciences, Immunology, Medicine, and Veterinary Science

List the installed catalog with:

```bash
word-journal-converter profiles
```

Every journal-specific profile records an official source URL and a `checked_on` date. The app reports profile age and warns when a profile becomes stale.

## What a profile does

A profile can encode only requirements supported by the current engine, including:

- abstract presence
- hard abstract maximum word count
- recommended abstract word-count target
- keyword minimum/maximum
- required manuscript sections
- citation/reference resolution
- tracked-change policy
- comment limits
- figure presence
- live citation-field presence

The safe profile retargeter can currently apply:

- page margins
- line numbering
- body font
- body font size
- line spacing

Requirements that the engine cannot evaluate should not be represented as automatically checked.

## Profiles are not citation styles

A journal profile describes manuscript requirements. It does not replace EndNote, Zotero, Mendeley, or a CSL style for citation formatting.

Citation-manager fields remain protected unless the user explicitly creates a separate static linked review copy.

## Format-free journals

Some publishers allow format-free or flexible initial submission. Those profiles are intentionally conservative. The app checks supported content requirements and may apply review-friendly formatting where the official guidance explicitly supports it, but it does not pretend to reproduce production typesetting.

## Template Mode

If the journal supplies an official Word `.docx` or `.dotx` template, Template Mode is the preferred complement to a journal profile.

Template Mode transfers a safe formatting subset from the supplied Word file to a new copy of the manuscript. It does not copy template body text or instructional placeholders. See [Template Mode](TEMPLATE_MODE.md).

This also provides a practical path for journals that are not yet in the built-in catalog.

## Validate a profile

```bash
word-journal-converter validate-profile scientific-reports-article
word-journal-converter validate-profile path/to/custom-profile.json
```

The validator checks required profile structure, URL syntax, and ISO `checked_on` dates.

## Freshness

Readiness reports calculate the age of a source-dated profile. Profiles older than 120 days generate a freshness warning. Journal instructions change, so the journal's current official guidance remains authoritative.

## Custom profiles

Use `journal-profiles/profile-template.json` as the starting point. Populate only rules supported by current official instructions and by the engine. Record ambiguity in `notes` rather than guessing.
