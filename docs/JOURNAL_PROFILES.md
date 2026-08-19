# Journal profiles

A journal profile is machine-readable, source-dated evidence for rules that Word Journal Manuscript Converter knows how to evaluate. It is not a permanent replacement for the journal's official author instructions.

## Bundled profiles

```bash
word-journal-converter profiles
```

The v0.3 beta includes:

- `plos-one-research-article`
- `scientific-reports-article`
- `frontiers-microbiology-original-research`
- `generic-review-copy`

The three journal-specific profiles were checked against official publisher guidance on 2026-08-19. Every profile records its official source URL and checked date.

## Validate a profile

```bash
word-journal-converter validate-profile scientific-reports-article
word-journal-converter validate-profile path/to/custom-profile.json
```

The validator checks required profile structure, URL syntax, and ISO `checked_on` dates.

## Freshness

Readiness reports calculate the age of a source-dated profile. Profiles older than 120 days generate a freshness warning. This is deliberately conservative because journal instructions change.

## Supported profile rules

The current engine can evaluate:

- abstract presence
- abstract maximum word count
- keyword minimum/maximum
- required manuscript sections
- citation/reference resolution
- tracked-change policy
- comment limits
- figure presence
- live citation-field presence

The current safe retargeting engine can apply:

- margins
- line numbering
- body font
- body font size
- line spacing

Requirements that the engine cannot evaluate should not be represented as if they were automatically checked.

## Custom profiles

Use `journal-profiles/profile-template.json` as the starting point. Populate rules only from current official instructions. Record ambiguity in `notes` rather than guessing.
