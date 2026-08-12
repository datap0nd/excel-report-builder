# Agent Handoff

## Current Objective

Tasks 1-11 are complete on `codex/pivottable-plus`. Task 11 replaces the old
workbench at the COM composition boundary with the compact, keyboard-accessible
PivotTable+ pane and Ribbon actions. A local unsigned 0.2.0.1 candidate is
installed and registered on the development PC.

The next numbered task is Task 12: reframe the optional local-model workflow as
a typed PivotTable+ proposal and preview flow. It must not apply without an
explicit user action.

## Repository State

- Branch: `codex/pivottable-plus`
- Tasks 1-11 in `plan.md` are complete; Tasks 12-14 remain.
- Task 9 was implemented and verified without controlling the Excel UI or using
  computer automation. All repository evidence uses generated synthetic data
  and generic names.
- The PivotTable+ rebuild remains intentionally incomplete until Tasks 12-14
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
- The final automated suite has 1,025 passing tests: 244 Core, 98 Agent, and
  683 Excel tests.
- Coverage includes validation, read-only discovery, session identity,
  classic and OLAP layout planning, fail-closed COM capture, exact verification
  and rollback, schema 1.3 ownership, raw-range binding, interrupted conversion
  recovery, collision handling, and idempotent retry.
- The repository public-safety check passes against tracked and nonignored
  files. Task 9 fixtures and documentation contain only generated synthetic or
  generic content.

## Task 10 Semantic PivotTable Contract

### Typed calculations and periods

- Closed, immutable calculation contracts compile only trusted schema field
  IDs into deterministic DAX. Supported operations include aggregate,
  weighted aggregate, difference, safe ratio, parent and filtered-total share,
  growth, achievement, variance, variance percent, and percentage-point delta.
- Period and scenario slices explicitly replace or intersect current axis
  context. Coverage validation blocks impossible output such as manufacturing
  quarters from yearly-only source data.
- Dependencies form a bounded acyclic graph. Measures carry deterministic
  definition and formula fingerprints, typed formatting, stable generated host
  names, and exact creation and display order. No public raw-DAX node exists.
- Exact typed equality uses strict set membership semantics, preventing DAX
  blank coercion from treating blank as zero, false, or empty text.

### Asymmetric named sets

- A bounded schema catalog binds opaque hierarchy, level, and member IDs to
  provider unique names. The model cannot supply or execute arbitrary MDX.
- Closed named-set expressions support exact ordered tuples, hierarchy default
  members, scoped/asymmetric branches, and typed Top N against an exact owned
  DAX measure. They compile deterministically to set MDX and remain source- and
  measure-fingerprint bound.
- Excel mutation uses the supported CalculatedMembers.Add(xlCalculatedSet),
  CubeFields.AddSet, MakeConnection, IsValid, and CubeField placement path.
  Partial host commits, orphaned calculated-member/cube-field pairs, source
  drift, sibling use, dynamic references, and inventory drift fail closed or
  reconcile to an exact state.
- A labeled Others member is deliberately unsupported: a named set can select
  existing tuples but cannot create the complement as a new dimension member.

### Combined one-PivotTable transaction

- `PivotSemanticMutationService` composes measure upserts, named-set upserts,
  the complete Rows/Columns/Values layout, named-set deletes, and measure
  deletes under one write-ahead journal and exactly one selected-PivotTable
  refresh.
- The semantic layout preserves Filters exactly, handles the Excel Values
  pseudo-field explicitly, interleaves generated and existing Values, and
  never creates a helper worksheet or formula-backed companion table.
- Verification rescans measures, calculated members, named sets, sibling
  PivotTables, dependencies, filter state, and final layout before one combined
  ownership commit. Failure rolls back in reverse dependency order and proves
  the prior state before clearing the journal.
- Same-session retries converge an ambiguous post-commit layout without a
  duplicate host artifact. One-level Undo keeps formulas and prior native
  definitions only in memory, journals the inverse transition, refreshes once,
  and restores the exact prior layout and ownership. Post-restart semantic Undo
  is intentionally unavailable because formulas are never persisted in
  workbook metadata.
- Ownership schema 1.4 stores only bounded IDs, operation kinds, target
  references, and hashes for pending Measure and NamedSet transitions. Tests
  assert that DAX and MDX never enter Custom XML.

## Task 11 Pane and Local Installation

- The COM task-pane host now composes `PivotPlusView` and
  `ExcelPivotPlusHostService`; the old report workbench is no longer reachable
  from the Ribbon or task-pane bootstrapper.
- The pane reads only the PivotTable under the active cell, searches its field
  catalog, shows familiar Filters/Columns/Rows/Values areas, and holds edits as
  a preview until the user chooses Apply.
- Excel's built-in PivotTable Fields pane remains the primary drag-and-drop
  workflow and is available from both the PivotTable+ pane and Ribbon.
- An explicit Enable PivotTable+ action invokes the verified Task 9
  classic-to-Data-Model conversion. The first visible extra is a typed parent
  Portion measure with session Undo; it never accepts raw DAX or creates a
  helper table.
- The rendered 420-by-900 synthetic pane was inspected. The task-pane contract,
  COM contract, public-safety check, dual-target AddIn build, and the full 1,025
  test suite pass.
- Inno Setup 6.7.1 built unsigned candidate `0.2.0.1`. Installation completed
  without a Windows restart. Installed AddIn/Core/Excel/Agent DLL hashes match
  the Release payload, both registry views have `FriendlyName=PivotTable+` and
  `LoadBehavior=3`, and both 32-bit and 64-bit PowerShell successfully activate
  the add-in and ActiveX task-pane COM classes.
- No Excel UI was controlled during verification. The user must perform the
  final visible open-pane and PivotTable interaction check.

## Remaining Boundary and Next Step

- Automated tests use synthetic late-bound hosts and do not replace live Excel
  evidence. No Excel UI or computer control was used for Task 9.
- Live smoke coverage for Excel LTSC 2021 and Microsoft 365 remains part of
  Task 14. The main host risks are COM/RCW identity behavior, Values
  pseudo-field placement, PivotCache replacement, provider and Data Model
  differences, refresh timing, and rollback after real-host COM failures.
- Task 12 will map local-model requests to the same typed calculations, named
  sets, and layout plan. It must not expose arbitrary DAX, MDX, formulas, COM,
  filesystem, save, publish, or deletion capabilities to the model.
- Live-host and installer release evidence remains deferred to Task 14. Task 10
  used no Excel UI or computer control; all host evidence is generated,
  late-bound, synthetic, and path-free.
