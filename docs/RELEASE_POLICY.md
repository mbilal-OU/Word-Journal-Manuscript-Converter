# Release policy

The project separates user-facing release names from engineering version identifiers.

## Pre-launch

Users should see simple names:

- Pre-Launch Beta 1
- Pre-Launch Beta 2
- Release Candidate 1

Engineering systems still require ordered versions. The current pre-launch build uses:

- package version `0.5.0`
- release tag `v0.5.0-beta.1`

## Stable launch

The first stable public product release is reserved for `1.0.0`.

The stable label should not be used until the launch-readiness gates in the README are complete.

## Historical v0.x releases

Earlier `v0.2.x` through `v0.5.0` releases are engineering history. They should be treated as pre-release builds rather than as separate public product generations.
