# Product

## Name

PivotTable+

The existing assembly names, COM ProgIDs, CLSIDs, and installer application ID
remain stable during the rebuild so an installed development copy can be
upgraded without reinstalling Office or Windows. They are compatibility
identifiers, not the product name.

## Purpose

PivotTable+ extends a real Microsoft Excel PivotTable with advanced layout,
calculation, period, and automation controls. Analysts keep Excel's familiar
PivotTable, PivotTable Analyze and Design tabs, and native Field List. The
add-in supplies a compact companion pane for operations that Excel exposes
poorly or cannot express through the ordinary field-list workflow.

The default result is always one refreshable native PivotTable. PivotTable+
does not generate a formula-backed report beside a PivotTable for a supported
feature.

## Platform

- Windows 10 or newer.
- Excel LTSC 2021 is the compatibility floor.
- Compatible Microsoft 365 desktop Excel is supported.
- 32-bit and 64-bit Office.
- .NET Framework 4.8 in-process COM add-in with a docked WPF task pane.
- Optional out-of-process local or explicitly configured OpenAI-compatible
  model endpoint.

Excel for the web and macOS are outside the first package because the required
Data Model, DAX, MDX named-set, and COM APIs are desktop capabilities.

## Users

Excel-fluent analysts who already understand Rows, Columns, Values, Filters,
subtotals, and PivotTable refresh. They need precise recurring reports without
maintaining a second grid of brittle formulas.

## Native-first contract

1. Standard placement, filtering, sorting, formatting, expand/collapse, and
   totals use Excel's PivotTable object model.
2. Advanced calculations are typed operations compiled to workbook DAX model
   measures. The measures appear in Excel's native Field List.
3. Asymmetric row and column layouts are typed operations compiled to ordered
   MDX named sets on a Data Model or OLAP PivotTable.
4. A classic range PivotTable remains usable for compatible actions. Features
   that require the Data Model show an explicit enablement preview and preserve
   or restore the original PivotTable if conversion fails.
5. The selected PivotTable remains user-owned. PivotTable+ records ownership
   only for artifacts it creates and refuses name collisions with user-owned
   measures or sets.
6. Unsupported constructs fail with a specific explanation. There is no silent
   fallback to a formula report.

## Capabilities

### Familiar PivotTable editing

- Identify the PivotTable under the active cell and its classic, Data Model, or
  external OLAP source type.
- Search and place fields in Rows, Columns, Values, and Filters.
- Reorder fields, choose aggregations, apply native number formats, choose
  subtotals and grand totals, sort, filter, and refresh.
- Keep the built-in Field List usable alongside the PivotTable+ pane.

### Advanced rows and ordering

- Add detail only beneath an exact parent path.
- Select a subset of members for that scoped detail.
- Continue the remaining hierarchy beneath the inserted detail.
- Reorder a scoped block among normal children and order its children
  independently.
- Emit an ordered Value, percentage of immediate parent, and percentage of
  filtered total stack.
- Apply Top N within each parent, optional Others, and an independent scoring
  context while preserving all displayed periods.

### Advanced columns and periods

- Compose exact ordered period and scenario outputs rather than a symmetric
  cross-product.
- Support explicit Year, Half, Quarter, Month, and full-date source grain.
- Select all or latest periods and order child periods with requested rollups.
- Create layouts such as Actual Jan, Actual Feb, Actual Mar, Q1 Plan, and Q1
  Variance without producing unwanted monthly Plan columns.
- Reject an impossible rollup, such as manufacturing quarters from yearly-only
  source data.

### Measures and comparisons

- Sum, Count, Distinct Count, Average, Minimum, Maximum, and weighted average.
- Filtered aggregate, ratio, share, difference, percent change, growth rate,
  achievement rate, variance, variance percentage, and percentage-point delta.
- Measures may depend on other PivotTable+ measures through a validated acyclic
  dependency graph.
- Labels, number formats, descriptions, and ordering persist in the workbook
  and native Field List.
- Zero-denominator behavior is explicit and defaults to blank.

### Filters and reusable setups

- One typed filter contract for report, row, scoped-detail, column, period, and
  value context.
- Exact searchable members with include, exclude, contains, comparison, and
  list operations where the source supports them.
- Saved PivotTable+ setup metadata reconnects to the same PivotTable after
  reopen and refresh without storing workbook paths or source values.
- Built-in presets include Portion, Growth/Rate, Actual/Rate, Variance, and an
  Actual-versus-Plan period pack.

### Agent-compatible authoring

- A local model may describe a proposed typed action plan; it never calculates
  workbook values or emits executable DAX, MDX, formulas, VBA, or COM calls.
- The deterministic host validates field names, members, grain, dependencies,
  capabilities, ownership, and estimated change scope.
- The user sees a complete preview before Apply and receives a bounded Undo for
  the last applied change.
- A selected blank formatted range or example layout may supply deterministic
  geometry and label hints. An image is supplementary context, never the sole
  source of truth.

## Trust and performance

- DAX evaluation and PivotTable rendering remain in Excel's Data Model engine.
- Multiple edits are batched under `ManualUpdate` and refreshed once.
- Added measures and named sets are benchmarked against equivalent manually
  authored native definitions; no claim of identical speed is made before the
  representative scenario passes.
- Apply is transactional: snapshot, validate, mutate, refresh, verify, and
  restore on failure.
- Refresh, reopen, and operation with the add-in disabled are release gates.
- No source edits, automatic workbook saves, email, scheduling, or arbitrary
  external-file automation occur in the first PivotTable+ package.

## Design commitments

- The Microsoft PivotTable and Field List remain visually dominant.
- The pane is compact, flat, keyboard accessible, and uses Segoe UI and native
  Windows control behavior.
- Default language is workbook-oriented: PivotTable, Rows, Columns, Values,
  Filters, Portion, Variance, Periods, Preview, Apply, and Undo.
- DAX, MDX, cube, tuple, and model terminology appears only in diagnostics or
  advanced help.
- No chat-first dashboard, decorative cards, or duplicate field-list clone.

## Deferred extras

Word exhibits, commentary generation, scheduling, delivery, change alerts,
approvals, and script export are valuable later modules. They do not block the
native PivotTable+ package and must not distort its first-run experience.

## Evidence policy

The public repository contains generated synthetic fixtures only. Real
workbooks, screenshots, prompts, endpoints, credentials, organization names,
and copied private formulas are prohibited.
