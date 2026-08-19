# Access strategy

Word Journal Manuscript Converter should be easy for a researcher who never uses a terminal, while preserving a reproducible technical core.

## 1. Desktop application: primary interface

This is the preferred mainstream product.

Workflow:

1. Choose or drag in a `.docx` manuscript.
2. Inspect document integrity.
3. Review citation/reference status.
4. Choose or load a journal profile.
5. Preview auto-fixable formatting rules.
6. Save to a new output file.
7. Run the preservation gate automatically.
8. Show a human-readable audit summary.

Why desktop first:

- unpublished manuscripts stay local
- the engine can inspect the complete DOCX package
- no browser upload is required
- the same core works without an internet connection

The repository includes a Tk GUI now and release-build automation for later standalone binaries.

## 2. Microsoft Word add-in: companion interface

The add-in should make integrity checks visible where researchers already edit manuscripts.

The current starter performs a local quick check inside Word. It should remain conservative. Package-level transformations belong in the desktop core until the same integrity guarantee can be maintained.

Long-term flow:

```text
Word task pane
    -> user requests safe operation
    -> signed local Word Journal Manuscript Converter bridge
    -> transformed DOCX copy
    -> preservation audit
    -> user reviews result
```

## 3. CLI and Python API: reproducible interface

Use cases:

- research groups
- publishing offices
- batch checking
- CI regression tests
- institutional workflows

## 4. GitHub Pages: discovery and documentation

The repository contains a static landing page under `site/`. Once GitHub Pages is enabled, the Pages workflow can publish it automatically.

## 5. Browser/PWA: later

A browser version is possible, but it should not begin as a traditional upload-to-server service for unpublished manuscripts.

Preferred future approaches:

- client-side processing where technically possible
- explicit local desktop handoff
- opt-in server features only for public journal metadata

## Release path

1. Source release on GitHub.
2. Windows and macOS standalone binaries.
3. Microsoft Word add-in beta.
4. Signed desktop app distribution.
5. Optional installer/package-manager channels.
6. Word add-in deployment or marketplace route after privacy and integration testing.
