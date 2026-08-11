# Excel Report Builder Plan

## Numbered Tasks

1. [x] Bootstrap the public repository, product and security documentation,
   native add-in shell, worker protocol, continuous activity feed, Windows CI,
   and per-user installer.
2. [x] Implement source profiling, period detection, one-row metric-month
   normalization, bounded transformation specifications, worksheet/Data Model
   routing, and total-preservation tests.
3. [x] Implement the versioned report specification, manual Rows, Columns,
   Values, Filters and formatting builder, native PivotTable execution, and
   workbook persistence.
4. [x] Implement dense output blocks, typed measures, managed ownership,
   independent validation, publish approval, and rollback.
5. [x] Implement OpenAI-compatible model settings, the guarded local worker,
   schema-aware chat, validation-driven repair, checkpoints, cancellation, and
   final change reports.
6. [x] Complete public release verification. User-authorized Excel field
   validation remains a separate next activity and was not part of this
   release.
7. [x] Repair the Office COM callback ABI, add marshaling regression checks,
   and validate add-in startup against generated synthetic data in Microsoft
   365 desktop Excel.
8. [x] Replace the released dense-report product contract with PivotTable+:
   preserve the hardened COM/installer/worker infrastructure, define the
   supported real-PivotTable feature matrix, and mark the old managed-output
   workflow for retirement.
9. [x] Implement active PivotTable/source discovery, ordinary-to-Data-Model
   enablement, native field placement, and workbook-owned PivotTable+ metadata.
10. [ ] Implement workbook DAX measure authoring and ordered MDX named sets for
    ratios, comparisons, period selections, asymmetric branches, and custom
    row/column order inside a real Excel PivotTable.
11. [ ] Replace the Data/Build/Chat/Checks workbench with a compact,
    keyboard-accessible PivotTable+ task pane and contextual Ribbon actions
    that operate on the selected native PivotTable.
12. [ ] Reframe the guarded local-model worker around typed PivotTable+ actions,
    preview, deterministic validation, and explicit apply/undo boundaries.
13. [ ] Retire dense formula reports, publishing, and obsolete managed-output
    UI; update branding, documentation, installer metadata, synthetic samples,
    and migration notes.
14. [ ] Complete the PivotTable+ release: full automated suite, COM contract
    checks, installer verification, generated synthetic workbook evidence, and
    a committed and published GitHub release candidate.

## Completion Contract

The PivotTable+ rebuild is complete only when Tasks 8-14 are on `main`, the
Windows build and installer checks pass, and the generated synthetic scenario
proves that standard fields, generated measures, and asymmetric named sets all
remain part of one refreshable native Excel PivotTable. The default workflow
must not create a formula-backed companion report. Model requests remain typed
and bounded; they cannot execute formulas, VBA, arbitrary COM, shell, file, or
save operations. User-operated visual validation is separate from automated
and non-UI repository verification unless UI control is explicitly authorized.
