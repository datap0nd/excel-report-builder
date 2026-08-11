# Agent Handoff

## Current Objective

The public `v0.1.0` unsigned prototype is released and independently verified.
An authorized synthetic-data field test on Microsoft 365 desktop Excel exposed
and repaired invalid hand-written Office COM callback metadata. The repaired
add-in now starts and connects without crashing. The next activity is to
exercise the visible task pane through a complete synthetic managed-draft flow.

## Repo State

- Branch: `main`
- Public repo: `datap0nd/excel-report-builder`
- Release source commit: `6e512ea3174eab5bb2d9b471ca4ee24aa14af3e5`
- Tag: `v0.1.0`
- Prerelease: <https://github.com/datap0nd/excel-report-builder/releases/tag/v0.1.0>
- Tasks 1-7 in `plan.md` are complete. Interactive managed-draft validation
  remains separate.
- Main-branch protection is enabled after the final evidence commit passes its
  required Windows and security checks.
- The published `v0.1.0` installer predates the COM ABI repair. A patch release
  has not yet been cut.

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
- The repaired Release build completes with zero warnings and zero errors; all
  484 synthetic tests, the COM ABI contract, task-pane contract, and public
  safety checks pass locally.
- A minimal A/B COM probe reproduced the CLR access violation with the old
  `IDTExtensibility2` declaration and completed `OnConnection`,
  `OnAddInsUpdate`, and `OnStartupComplete` after matching Office's
  `SAFEARRAY(VARIANT)` contract.
- Microsoft 365 desktop Excel `16.0.20228.20158` x64 opened the generated
  `sales_long.csv` source with `ExcelReportBuilder.AddIn` connected,
  `LoadBehavior=3`, and no new Excel or .NET Runtime crash event.

## Remaining Boundary

- The installer is intentionally unsigned, so Windows displays an
  unknown-publisher warning.
- A real Excel LTSC 2021 workbook session has not been exercised.
- Microsoft 365 startup and connection are field-validated with generated
  synthetic data; the full visible task-pane build, check, and publish flow is
  not yet recorded as field evidence.
- The running Excel session has the ABI-repaired field binary loaded. Additional
  RCW lifecycle hardening in the current source is built and tested, but the
  locked installed DLL cannot be replaced until Excel closes normally. No
  Windows restart or Office reinstall is required.
- A real workbook requires separate explicit authorization and must remain
  confined to managed drafts.
- Do not reinstall the original `v0.1.0` asset over the repaired local field
  installation.

## Next Step

Ask the user to verify **Data > Report Builder** with a generated source, build
a managed draft, run checks, and publish only after explicit review. Do not
control Excel's UI or inspect a real workbook without separate authorization.
After that evidence is recorded, publish a patch release containing the
repaired COM contracts.
