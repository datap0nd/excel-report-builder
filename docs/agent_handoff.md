# Agent Handoff

## Current Objective

Tasks 1-9 are complete on `codex/pivottable-plus`. Task 9 establishes the safe,
native PivotTable+ foundation: read-only discovery, validated layout mutation,
explicit classic-to-Data-Model conversion, durable recovery, and exact ownership
of generated workbook artifacts.

The next numbered task is Task 10: implement typed, deterministic DAX measures
and ordered MDX named sets for ratios, comparisons, period selections,
asymmetric branches, and custom row or column order inside one real PivotTable.

## Repository State

- Branch: `codex/pivottable-plus`
- Tasks 1-9 in `plan.md` are complete; Tasks 10-14 remain.
- Task 9 was implemented and verified without controlling the Excel UI or using
  computer automation. All repository evidence uses generated synthetic data
  and generic names.
- The PivotTable+ rebuild remains intentionally incomplete until Tasks 10-14
  are implemented and verified.

## Task 9 Product Contract

### Core contracts and validation

- Immutable PivotTable+ contracts describe a path-free workbook, worksheet,
  PivotTable, source, field catalog, placements, filters, layout, and format.
- Layout mutation is explicit: a definition must contain placements unless
  `ClearAll` is set, and `ClearAll` cannot be combined with placements.
- Rows, Columns, Filters, and Values have deterministic positions. The special
  Excel Values axis is typed as Automatic, Rows, or Columns with a one-based
  position; two or more Values require an explicit axis.
- Validation derives required capabilities from the requested operation and
  enforces the classic worksheet, Data Model, and external OLAP source truth
  table. Capability flags describe Excel and source potential, while each
  executor still fails closed for operations it does not implement.
- A regular source field cannot occupy more than one non-Values area. Classic
  Values may repeat a source field when captions remain unique. Data Model
  implicit Values support Sum, Count, Average, Minimum, and Maximum with unique
  field-and-function pairs. External OLAP Values must reference existing
  measures.
- Workbook-local identifiers and field bindings reject path-like values,
  unknown fields, invalid enum states, ambiguous captions, and source or target
  contradictions before native mutation.

### Read-only discovery and workbook identity

- Discovery inspects the active, selected native PivotTable and identifies a
  classic worksheet source, the workbook Data Model, or external OLAP without
  changing the workbook.
- It captures fields, placements, layout, formatting, source capabilities, and
  the Excel Values pseudo-field while excluding that pseudo-field from the
  regular field catalog through COM identity comparison.
- Workbook identity resolution is read-only. It returns an existing stored ID
  or a session-stable `workbook_<guid>` token, so opening the pane or selecting
  a PivotTable does not add Custom XML or dirty the workbook.
- Apply persists that exact token once, after live target and source preflight
  plus rollback-state capture succeed and immediately before the first native
  mutation. Exact-token persistence is idempotent and rejects collisions.

### Native PivotTable mutation

- The native service binds the supplied COM object to the requested workbook,
  worksheet, PivotTable, cache, and source before it captures or changes state.
- Classic and OLAP/Data Model fields can be placed in Rows, Columns, Filters,
  and Values with supported aggregation, position, row-axis layout, totals,
  subtotals, repeated labels, style, and number-format metadata.
- The Values pseudo-field axis and position are captured, applied, restored,
  and verified explicitly instead of being treated as an ordinary source
  field.
- Preflight fails closed for unreadable mutation-relevant COM state, active
  native filters that cannot be preserved, unsupported Show Values As state,
  calculated fields or items, mixed per-row repeated labels, source drift, and
  unsupported source or aggregation combinations.
- Apply uses bounded native batching, refresh, exact postcondition checks, and
  rollback. Rollback restores layout, formatting, filters and captions that are
  represented by the contract, including pre-existing implicit measure
  captions. Newly created implicit measures are tracked for cleanup on failure.
- Task 9 never substitutes a formula-backed companion report and never claims
  ownership of the user's PivotTable or source data.

### Explicit classic-to-Data-Model conversion

- Conversion is a separate, explicit operation for an ordinary classic
  PivotTable. It can bind a worksheet table or a workbook-scoped source name
  for a raw range without converting or reformatting the user's source.
- Generated source name, query, connection, and temporary PivotTable artifacts
  are planned and fingerprinted before creation. The original classic
  PivotTable remains available until the replacement is independently bound,
  refreshed, formatted, and verified.
- Ownership schema 1.3 provides write-ahead Pending state, exact temporary
  worksheet and PivotTable receipts, recovery checkpoints, and an Active
  promotion boundary. `RecoverPending` durably completes or safely converges an
  interrupted conversion without growing artifacts or deleting unowned state.
- Retry and cleanup require exact IDs, fingerprints, source lineage, and target
  receipts. Ambiguous, changed, or contaminated workbook state fails closed.

### Workbook-owned metadata

- Versioned, deterministic Custom XML stores only path-free identifiers,
  fingerprints, the target worksheet and PivotTable, bounded undo or recovery
  metadata, and lifecycle state.
- Ownership covers only generated measures, named sets, queries, connections,
  workbook-scoped source names, and temporary conversion artifacts. Exact
  collision guards prevent one setup from claiming another setup's artifacts.
- Workbook paths, source data, prompts, credentials, endpoint details, and
  measure formulas are not stored in PivotTable+ ownership metadata.

## Verification Evidence

- Release builds complete with zero warnings and zero errors for both Excel
  target frameworks: .NET Framework 4.8 and .NET Standard 2.0.
- The final automated suite has 786 passing tests: 170 Core, 98 Agent, and 518
  Excel tests.
- Coverage includes validation, read-only discovery, session identity,
  classic and OLAP layout planning, fail-closed COM capture, exact verification
  and rollback, schema 1.3 ownership, raw-range binding, interrupted conversion
  recovery, collision handling, and idempotent retry.
- The repository public-safety check passes against tracked and nonignored
  files. Task 9 fixtures and documentation contain only generated synthetic or
  generic content.

## Remaining Boundary and Next Step

- Automated tests use synthetic late-bound hosts and do not replace live Excel
  evidence. No Excel UI or computer control was used for Task 9.
- Live smoke coverage for Excel LTSC 2021 and Microsoft 365 remains part of
  Task 14. The main host risks are COM/RCW identity behavior, Values
  pseudo-field placement, PivotCache replacement, provider and Data Model
  differences, refresh timing, and rollback after real-host COM failures.
- Task 10 is next. It must expose only typed, validated measure and named-set
  operations; it must not expose arbitrary DAX, MDX, formulas, COM, filesystem,
  save, publish, or deletion capabilities to the model.
- Task 11 will add the PivotTable+ UI after the native and calculation layers
  exist. Live-host and installer release evidence remains deferred to Task 14.
