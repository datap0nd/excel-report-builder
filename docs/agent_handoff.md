# Agent Handoff

## Current Objective

Verify the Windows build and per-user installer from `main`, inspect the
Windows-rendered task pane, publish the `v0.1.0` unsigned prerelease, and record
the result. Managed-workstation Excel validation remains separately authorized.

## Confirmed Decisions

- This repository is a clean implementation with no imported private source or
  history.
- Repository and product name: Excel Report Builder.
- Windows Excel LTSC 2021 is the primary runtime.
- One raw extract produces one output; no joins or SQL.
- One-row month and metric-month normalization is the first functional gate.
- Large normalized results route to Excel's Data Model without truncation.
- The agent works only in managed drafts and continuously reports progress.
- Managed-workstation field testing is out of scope until separately
  authorized.

## Implemented

- `ReportSpecV1` JSON Schema, strict JSON shape validation, typed transforms,
  measures, period slices, report blocks, presentation, checks, and ownership.
- Long and wide period profiling, explicit missing-year rejection,
  metric-month unpivoting, worksheet versus Data Model routing, and independent
  canonical-data auditing.
- Manual preparation, field placement, subtotal/order/format controls,
  calculated metrics, stable multi-block layout, managed extents, and
  report-specific checks.
- Native PivotTables, hidden-pivot dense reports, controlled generated
  formulas, rebuild idempotency, draft-only mutation, publish approval, and one
  managed rollback copy.
- Current-user local worker, restricted named pipe, protected endpoint
  credentials, guarded tool calls, checkpoints, cancellation, finite repairs,
  exact specification handoff, and continuous progress with 15-second
  heartbeats.
- Pinned Windows build, COM contract, x86/x64 worker packaging, installer smoke
  test, dependency review, secret scanning, CodeQL, SBOM, checksum, and
  provenance workflows.

## Verification

- Local Release build: clean, zero warnings.
- Synthetic tests: 295 passing.
- Dependency vulnerability scan: no vulnerable NuGet packages reported.
- Impeccable task-pane detector: no findings.
- Windows workflow, rendered preview, installer execution, and tagged release:
  pending the first `main` push.

## Public Safety

Use only generated synthetic fixtures and generic names. Never commit real
workbooks, screenshots, data, paths, credentials, endpoints, transcripts, or
private-repository content.

## Next Step

Push `main`, resolve any Windows-only failures, inspect the preview artifact,
tag `v0.1.0`, verify all release assets, then update this handoff with the
release evidence.
