# Contributing

Contributions are welcome, especially around DOCX preservation, citation-manager fixtures, journal-profile validation, and user accessibility.

## Research confidentiality

Do not commit, attach, or paste unpublished manuscripts. Tests should use synthetic documents or fully de-identified material.

## Development setup

```bash
python -m venv .venv
pip install -e . --no-build-isolation
pip install pytest
pytest
```

## Rules for DOCX mutations

A new automatic transformation should include:

1. explicit preconditions;
2. a narrowly defined OOXML scope;
3. a before/after preservation test;
4. fail-closed behavior when the preservation contract is not met;
5. documentation of what can and cannot change.

## Journal profiles

Use current official journal or publisher instructions. Record the source URL and checked date. Do not submit an unverifiable third-party style summary as an authoritative profile.

## Pull requests

Keep changes focused and include regression tests for behavior that touches manuscript contents or OOXML structure.
