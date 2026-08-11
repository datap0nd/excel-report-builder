# Agent Handoff

## Current Objective

Publish the hardened `v0.1.0` unsigned prototype from `main`, verify the Windows
installer and rendered task pane, then record the release evidence. Real Excel
field validation remains separately authorized.

## Repo State

- Path: repository root
- Branch: `main`
- Latest published commit before this release batch: `cde179b`
- Public repo: `datap0nd/excel-report-builder`
- Push status: the release-hardening batch is verified locally and not yet
  pushed

## Decisions Made

- The product remains a clean public implementation with generated synthetic
  fixtures only.
- The add-in mutates exact managed workbook objects only and never saves a
  workbook.
- Full-source transformations are independently evaluated for exact row and
  additive-total reconciliation. The reference evaluator and generated M share
  exact-case fields, bounded en-US conversions, null filter behavior, error
  recovery, numeric period handling, and finite decimal arithmetic.
- Wide normalization expands one explicit record per mapped cell, including
  null cells, so projected and actual row counts cannot diverge.
- Worksheet and Data Model PivotCaches use bounded backend-specific ownership
  slots. Inactive managed backend objects can remain when Excel-owned caches
  depend on them.
- Registry-only ownership is not enough to mutate a live query, workbook name,
  connection, or PivotTable. Exact content or source contracts must also match.
- Publishing creates static values and formats. Final and rollback worksheets
  contain neither formulas nor live PivotTables.
- Every worker launch uses a random current-user pipe and one-time HMAC proof
  before endpoint credentials, prompts, or workbook samples are sent.
- Managed-workstation testing is out of scope until explicitly authorized.

## Files Changed

- `src/ExcelReportBuilder.Excel`: exact source reconciliation, backend routing,
  PivotCache ownership, output auditing, and transactional static publishing.
- `src/ExcelReportBuilder.AddIn`: exact saved-setup rebuilding, endpoint-scoped
  credential state, multi-block manual projection, and task-pane binding fixes.
- `src/ExcelReportBuilder.Agent` and `src/ExcelReportBuilder.Worker`: authenticated
  one-use worker transport and endpoint credential scoping.
- `.github/workflows/windows-build.yml`, `scripts`, and release docs: x86/x64
  worker smoke tests, complete-payload SBOM validation, release ancestry, and
  public-safety gates.
- `tests`: adversarial coverage for every hardened contract.

## Commands And Checks

- Release solution build: passed with zero warnings.
- Full synthetic test suite: 484 passed, 0 failed, 0 skipped.
- NuGet transitive vulnerability scan: no vulnerable packages reported.
- JSON Schema and XAML parsing: passed.
- Public identifier, credential, private-artifact, em-dash, and diff checks:
  passed locally; the complete PowerShell gate will run in Windows CI.
- The pinned SBOM generator was checked locally against the Release add-in
  payload and recognized every first-party assembly and required runtime
  package under the names enforced by CI.
- Windows build, COM activation, installer, task-pane preview, SBOM, checksum,
  provenance, and tagged release: not yet run for this unpushed batch.

## Open Questions

- No code blocker is known. Windows CI remains the authoritative check for COM,
  PowerShell, WPF rendering, Inno Setup, and installed-worker behavior.

## Next Step

Commit and push the scoped release-hardening batch to `origin/main`, then follow
the Windows and security workflows through completion before creating the
`v0.1.0` tag.
