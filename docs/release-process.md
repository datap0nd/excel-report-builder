# Release Process

Prototype releases are built entirely by the Windows workflow from a version
tag. Release credentials are not available to pull-request build steps.

## Prerequisites

- The version tag must use `vMAJOR.MINOR.PATCH` and match `VersionPrefix` in
  `Directory.Build.props`.
- The tagged commit must already be contained in `origin/main`. CI fetches the
  complete tag and branch history and rejects a tag from any other commit.
- The tagged commit must pass the solution tests, public-safety scan,
  dependency vulnerability scan, COM contract check, task-pane contract check,
  installer smoke test, and uninstall cleanup test.
- The installer requires Windows 10 or newer and .NET Framework 4.8 or newer.
  The local agent worker is self-contained and does not require a separate
  .NET 8 runtime.

## Build and release controls

- GitHub Actions and security actions are pinned to immutable commit hashes.
- Inno Setup is installed at an exact package version in CI.
- CI publishes separate x86 and x64 workers and executes the authenticated
  handshake and unauthenticated fail-closed smoke against both binaries. Setup
  selects the worker matching the Windows operating-system architecture. The
  out-of-process worker works with either 32-bit or 64-bit Excel.
- The installer registers the add-in per user in both Office registry views on
  64-bit Windows. No administrator permission is requested.
- A smoke test activates both COM classes through their registered ProgIDs in
  32-bit and 64-bit Windows PowerShell, completes and independently verifies the
  one-use authenticated named-pipe handshake with the installed worker, verifies
  that a worker without a launch secret fails closed, uninstalls the product,
  and verifies file and registry cleanup.
- The tag-only release job receives `contents`, `id-token`, and `attestations`
  write permissions. Build and pull-request jobs have read-only repository
  access.
- The SPDX SBOM is generated from a staged copy of the exact add-in DLLs and
  both published worker payloads before installer packaging. CI verifies every
  staged file and SHA-256 entry, plus package evidence for the add-in, Core,
  Excel, Agent, and both worker components, before the SBOM can become a release
  asset.

## Unsigned prototype warning

Until an Authenticode certificate is configured, every setup filename and
GitHub prerelease title contains `unsigned`. Windows will show an
unknown-publisher warning. Each prerelease includes a SHA-256 checksum, SPDX
SBOM, and GitHub build-provenance attestation.

Do not rename an unsigned setup file in a way that removes the warning label.
