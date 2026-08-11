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
6. [ ] Complete public release verification and later user-authorized Excel
   field validation. Packaging, accessibility contracts, diagnostics, and
   security workflows are implemented; field validation remains separate.

## Completion Contract

The first public prototype is complete only when Tasks 1-5 are on `main`, the
Windows build and installer checks pass, a tagged release is available, and the
synthetic end-to-end scenario produces a checked managed draft through both the
manual builder and chatbot specification path. Field validation on a managed
work PC remains separately authorized and is not implied by repository work.
