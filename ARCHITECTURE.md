# Architecture

## System boundary

PivotTable+ is an authoring layer over a real Excel PivotTable. It does not own
the aggregation grid and does not render a formula-backed companion report.

```text
Built-in Excel Field List              PivotTable+ task pane / Ribbon
            |                                      |
            +------------------+-------------------+
                               v
                    typed PivotPlusSpecV1 plan
                               |
             +-----------------+------------------+
             |                 |                  |
             v                 v                  v
      standard Pivot       DAX compiler       MDX set compiler
       object model             |                  |
             |                  v                  v
             |          ModelMeasures.Add   CalculatedMembers.Add
             |                                   + CubeFields.AddSet
             +-----------------+------------------+
                               v
                    one native Excel PivotTable
```

## Compatibility chassis

The rebuild keeps the already verified .NET Framework 4.8 COM bootstrap,
Office callback ABI, WinForms/ElementHost task-pane control, per-user x86/x64
registration, installer, and authenticated out-of-process worker. Assembly
names, ProgIDs, CLSIDs, and the installer application ID remain stable during
the product migration.

The former dense-report specification, renderer, publishing transaction, and
four-surface workbench are not part of the new runtime path.

## Projects

- `ExcelReportBuilder.Core`: temporary compatibility project name containing
  PivotTable+ specifications, typed measures, period semantics, compilers, and
  validation. It contains no COM objects.
- `ExcelReportBuilder.Excel`: active PivotTable discovery, capability
  inspection, native mutations, Data Model enablement, ownership metadata,
  refresh verification, and rollback.
- `ExcelReportBuilder.Agent`: endpoint policy, typed PivotTable+ tool schemas,
  prompt construction, response validation, and orchestration.
- `ExcelReportBuilder.Worker`: authenticated one-use out-of-process model host.
- `ExcelReportBuilder.AddIn`: COM entry point, Ribbon, compact task pane,
  preview/apply/undo commands, progress, and cancellation.

Project and assembly renaming is deliberately separated from behavior changes
so installation compatibility is not put at risk.

## Core specification

`PivotPlusSpecV1` identifies a target PivotTable without a workbook path and
contains only typed operations:

- ordinary field placements and native layout settings;
- measure definitions expressed as validated expression nodes;
- ordered period/scenario outputs;
- asymmetric axis branches and exact member paths;
- filters, Top N scoring context, and member order;
- labels, number formats, subtotal and grand-total behavior.

Raw DAX or MDX is not part of the public specification and cannot be supplied
by a model. Deterministic compilers are the sole source of executable formulas.

## Pivot capabilities

The context resolver classifies the selected PivotTable as:

| Source | Standard edits | DAX measures | Named sets |
| --- | --- | --- | --- |
| Classic worksheet/range | Yes | After explicit Data Model enablement | After explicit Data Model enablement |
| Excel Data Model | Yes | Yes | Yes |
| External OLAP | Yes, provider permitting | No workbook DAX; provider/private MDX rules apply | Yes, provider permitting |

Capability checks occur before preview. A plan that exceeds the current source
is rejected or offers an explicit conversion; it never changes backend
silently.

## Native mutation transaction

Every Apply follows one coordinator:

1. Resolve and revalidate the active PivotTable identity and capabilities.
2. Load the last saved PivotTable+ metadata and verify artifact ownership.
3. Snapshot affected field orientations, filters, measures, named sets, style,
   and totals.
4. Set a depth-based event/reentrancy guard and preserve `ManualUpdate`.
5. Upsert owned DAX measures using native model format objects.
6. Upsert owned MDX named sets, expose them with `CubeFields.AddSet`, and place
   them on the requested axis.
7. Apply ordinary field, filter, order, total, and formatting changes.
8. Restore update state, refresh once, reacquire COM objects, and verify native
   fields, values, ordering, and metadata.
9. Commit the ownership record and bounded undo snapshot.
10. On any failure, restore the snapshot and remove only newly created owned
    artifacts.

The user's PivotTable is a mutation target, never an owned object. A measure or
set with the requested name but a different ownership fingerprint is a hard
collision.

## Data Model enablement

Advanced actions on a classic PivotTable require an explicit enablement
transaction. The source is inspected and, if necessary, represented as a
managed query/model connection. PivotTable layout, filters, formatting, and
location are snapshotted before a Data Model PivotTable is created or rebound.
The original remains recoverable until refresh and semantic verification pass.

No claim is made that Excel can toggle an ordinary PivotCache into a Data Model
cache in place.

## DAX measures

Typed expression nodes support aggregate, filtered aggregate, weighted result,
reference, constant, arithmetic, safe divide, ratio, difference, share,
variance, growth, achievement, and percentage-point delta. A dependency graph
is validated for missing references and cycles before compilation.

The Excel adapter uses `Workbook.Model.ModelMeasures.Add` for new measures and
updates owned measures in place when safe. `FormatInformation` is a native
`ModelFormat*` object, not a format string. Associated table, display name,
description, and format are part of the ownership fingerprint.

## Asymmetric named sets

An asymmetric axis is represented as an ordered list of validated member
paths/tuples plus display and hierarchy options. The compiler escapes names and
members and emits MDX only after every referenced hierarchy and member is
resolved against the current cube schema.

The Excel adapter uses `CalculatedMembers.Add` with `xlCalculatedSet`, then
`CubeFields.AddSet` before orientation and position are assigned. Excel does
not apply ordinary filters directly to named sets, so an advanced branch edit
regenerates its owned set. Ordinary report filters and slicers remain native
when their semantics are compatible.

## Agent trust boundary

```text
bounded Pivot snapshot + user request
                |
                v
 authenticated one-use pipe -> local model worker
                |                     |
                |                     v
                |          typed action proposal only
                +---------------------+
                v
 deterministic validation and preview
                |
          explicit user Apply
                v
 native mutation transaction
```

The worker never receives an Excel application object. It cannot execute DAX,
MDX, worksheet formulas, VBA, COM, shell, filesystem, save, publish, delete,
email, or arbitrary network operations. Credentials are protected for the
current Windows user and transferred only after one-use HMAC authentication.

## Persistence

Workbook Custom XML stores path-free PivotTable+ metadata:

- workbook identity and target sheet/PivotTable names;
- source and cube-schema fingerprints;
- owned measure/set/query/connection identifiers and definition fingerprints;
- the typed setup and last bounded undo snapshot;
- format version and migration state.

The workbook remains refreshable and readable without the add-in. Native model
measures and sets remain workbook objects; only editing conveniences disappear.

## Verification gates

- Unit tests for specifications, dependency graphs, DAX/MDX escaping and golden
  compilation, period grain, scoped branches, and plan validation.
- Fake-COM tests for exact call order, capability checks, idempotency,
  reentrancy, ownership collisions, rollback, and refresh/reacquire behavior.
- Static COM ABI, Ribbon, task-pane, installer, SBOM, and public-safety checks.
- User-operated smoke matrix on Excel LTSC 2021 and Microsoft 365 using only
  generated synthetic data: native Field List visibility, values/order,
  refresh, save/reopen, add-in-disabled behavior, and owned-only cleanup.
