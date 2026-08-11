# Agent Handoff

## Current Objective

The public `v0.1.0` unsigned prototype is released and independently verified.
The next activity is a separately authorized field test in managed Excel,
starting with generated synthetic data and managed draft sheets. No managed
workstation or remote-desktop session was used for this release.

## Repo State

- Branch: `main`
- Public repo: `datap0nd/excel-report-builder`
- Release source commit: `6e512ea3174eab5bb2d9b471ca4ee24aa14af3e5`
- Tag: `v0.1.0`
- Prerelease: <https://github.com/datap0nd/excel-report-builder/releases/tag/v0.1.0>
- Tasks 1-6 in `plan.md` are complete. Field validation remains separate.
- Main-branch protection is enabled after the final evidence commit passes its
  required Windows and security checks.

## Delivered Product Boundary

- The product is a clean public implementation using generated synthetic
  fixtures and generic language only.
- The native .NET Framework 4.8 COM add-in provides Data, Build, Chat, and Checks
  surfaces in a WPF task pane with continuous typed progress and a 15-second
  no-silence heartbeat.
- Manual configuration and the guarded model worker produce the same versioned
  report specification. The worker has no arbitrary code, formula, COM,
  filesystem, save, publish, delete, email, or unrestricted network tool.
- Long and wide period layouts, deterministic transformations, worksheet and
  Data Model routing, native PivotTables, hidden-pivot dense blocks, typed
  measures, validation, repair, publish approval, and rollback are implemented.
- The source sheet remains unchanged. Autonomous work is limited to owned
  managed drafts, and the add-in never saves the workbook automatically.
- Endpoint credentials are protected for the current Windows user. Every worker
  launch uses a random current-user pipe and one-time HMAC proof before any
  credential, prompt, or workbook sample is sent.

## Verification Evidence

- Local Release build: zero warnings and zero errors.
- Full synthetic suite: 484 passed, 0 failed, 0 skipped.
- Final pre-tag Windows run: <https://github.com/datap0nd/excel-report-builder/actions/runs/31515081616>
- Final security run: <https://github.com/datap0nd/excel-report-builder/actions/runs/31515081587>
- Tagged build and release run: <https://github.com/datap0nd/excel-report-builder/actions/runs/31515470291>
- Windows verified public safety, build, all tests, vulnerable dependencies,
  complete-payload SBOM, COM contracts, WPF rendering, installer construction,
  per-user x86/x64 registry values and value kinds, repair, real COM activation,
  authenticated and fail-closed worker launches, uninstall, and cleanup.
- The final 420 by 900 task-pane render was inspected. Source field names are
  readable, and the operation identity uses constrained trimming at the minimum
  pane width.
- The public release manifest independently matched the downloaded installer
  and SPDX SBOM.
- Installer SHA-256:
  `eee4c80f3b5acf03e1ad61c8bfffd0cc83685ce4aa47079526e13a317afd545b`
- SBOM SHA-256:
  `aba6a35a66147143883bbb0c647cd785582add7fddc3c8d69ed2ad774a09082a`
- The installer and SBOM each passed SLSA provenance verification against the
  tagged source commit and `.github/workflows/windows-build.yml`.

## Remaining Boundary

- The installer is intentionally unsigned, so Windows displays an
  unknown-publisher warning.
- A real Excel LTSC 2021 workbook session has not been exercised. That test must
  begin with a small generated synthetic workbook. A real workbook requires a
  separate explicit authorization and must remain confined to managed drafts.
- No code or release blocker is known.

## Next Step

Wait for explicit authorization to run the managed Excel field test. Do not use
a remote-desktop session or inspect a real workbook merely because the public
release is complete.
