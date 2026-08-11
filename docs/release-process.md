# Release Process

Prototype releases are built entirely by the Windows workflow from a version
tag. Release credentials are not available to pull-request build steps.

## Prerequisites

- The version tag must use `vMAJOR.MINOR.PATCH` and match `VersionPrefix` in
  `Directory.Build.props`.
- The tagged commit must pass the solution tests, public-safety scan,
  dependency vulnerability scan, COM contract check, task-pane contract check,
  installer smoke test, and uninstall cleanup test.
- The installer requires Windows 10 or newer and .NET Framework 4.8 or newer.
  The local agent worker is self-contained and does not require a separate
  .NET 8 runtime.

## Build and release controls

- GitHub Actions and security actions are pinned to immutable commit hashes.
- Inno Setup is installed at an exact package version in CI.
- CI publishes separate x86 and x64 workers. Setup selects the worker matching
  the Windows operating-system architecture. The out-of-process worker works
  with either 32-bit or 64-bit Excel.
- The installer registers the add-in per user in both Office registry views on
  64-bit Windows. No administrator permission is requested.
- A smoke test activates both COM classes through their registered ProgIDs in
  32-bit and 64-bit Windows PowerShell, completes the current-user named-pipe
  handshake with the installed worker, uninstalls the product, and verifies
  file and registry cleanup.
- The tag-only release job receives `contents`, `id-token`, and `attestations`
  write permissions. Build and pull-request jobs have read-only repository
  access.

## Unsigned prototype warning

Until an Authenticode certificate is configured, every setup filename and
GitHub prerelease title contains `unsigned`. Windows will show an
unknown-publisher warning. Each prerelease includes a SHA-256 checksum, SPDX
SBOM, and GitHub build-provenance attestation.

Do not rename an unsigned setup file in a way that removes the warning label.
