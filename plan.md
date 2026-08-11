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

## Completion Contract

The first public prototype is complete only when Tasks 1-5 are on `main`, the
Windows build and installer checks pass, a tagged release is available, and the
synthetic end-to-end scenario produces a checked managed draft through both the
manual builder and chatbot specification path. Field validation on a managed
work PC remains separately authorized and is not implied by repository work.
Post-release compatibility repairs must preserve the same safety boundary and
pass both synthetic verification and an explicitly authorized Excel field test.
