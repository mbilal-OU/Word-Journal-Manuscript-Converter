# Updates

## Windows

The recommended Windows package is the installer:

`Word-Journal-Manuscript-Converter-Setup.exe`

After the app is installed once, it can check GitHub Releases for newer builds.

When a newer Windows release includes the setup executable, the app can download and launch the new installer. This removes the need to manually visit GitHub and download a ZIP for every update.

The user remains in control:

- automatic update checks can be disabled
- downloads occur only after confirmation
- the installer is launched visibly
- manuscript files are not involved in update checks

## Portable Windows ZIP

Portable ZIP users can also use the update checker, but installing the new setup package is recommended if they want the smoother update path.

## macOS and Linux

The app can detect newer releases. The current pre-launch update flow opens or downloads the platform package rather than silently replacing the application.

## Release channels

Public-facing pre-launch labels are intentionally simple:

- Pre-Launch Beta 1
- Pre-Launch Beta 2
- Release Candidate 1
- 1.0.0 stable

Internal package versions and Git tags remain machine-readable for build and update logic.

Current technical version:

- Python package: `0.5.0`
- Git tag: `v0.5.0-beta.1`
- Display label: `Pre-Launch Beta 1`
