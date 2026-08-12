# PivotTable+ Feature and Acceptance Contract

This document is the implementation boundary for the PivotTable+ rebuild. It
summarizes generic capabilities evidenced in the public Composer project and
maps them to one real Excel PivotTable. It does not copy Composer source or
private report content.

## Required package

### P0: native editing core

1. Operate on the real PivotTable selected in Excel. Do not create a copy or
   formula-backed companion report for a supported feature.
2. Preserve familiar Rows, Columns, Values, Filters, native Field List, native
   formatting, refresh, and PivotTable Analyze/Design behavior.
3. Provide scoped detail beneath an exact parent path, member subset selection,
   continued hierarchy, and independent block/child ordering.
4. Provide ordered Value, percentage of parent, and percentage of filtered
   total row outputs, per-level subtotals, and report totals.
5. Provide Top N per parent with optional Others and an independent scoring
   context that does not remove displayed periods.
6. Provide an asymmetric column composer for exact scenario/period/measure
   tuples.
7. Provide a calculation dependency graph for aggregate, filtered aggregate,
   weighted result, difference, ratio, share, growth, achievement, variance,
   variance percentage, and percentage-point delta.
8. Validate source period grain and member coverage before creating Year,
   Half, Quarter, Month, or full-date outputs.
9. Provide one typed filter contract across report, row, scoped detail, column,
   period, value, comparison, and Top N scoring context.
10. Preview every mutation, apply transactionally, refresh once, verify the
    result, and provide bounded Undo.

### P1: reusable and agent-compatible

11. Save a path-free setup in workbook metadata and reconnect after reopen and
    source refresh.
12. Accept a natural-language request through a local or explicitly configured
    OpenAI-compatible endpoint and return only typed PivotTable+ actions.
13. Use selected blank-range structure and formatting as optional deterministic
    layout hints; flag ambiguity instead of inventing structure.
14. Check missing fields/members, changed grain, zero denominators, stale owned
    artifacts, unsupported source capabilities, and expected change scope.
15. Include Portion, Growth/Rate, Actual/Rate, Variance, and
    Actual-versus-Plan period presets.

### P2: separate extras

Word exhibits, narrative commentary, saved-layout libraries, script handoff,
scheduling, delivery, change alerts, and approvals are separate modules. They
must not add clutter or a second report engine to the core PivotTable+ pane.

## Native implementation map

| User capability | Native implementation |
| --- | --- |
| Standard fields, totals, sorting, filters, style | PivotTable/PivotField/CubeField object model |
| Portion and standard show-values calculations | Native calculation where semantics match; otherwise owned DAX measure |
| Reusable ratios, comparisons, period/scenario values | Owned DAX model measures |
| Scoped parent-specific detail and per-parent order | Ordered MDX named set |
| Exact asymmetric period/scenario columns | Ordered measures plus MDX named set when tuple control is required |
| Classic PivotTable requiring an advanced feature | Explicit, reversible Data Model enablement |
| Local-model request | Typed proposal, deterministic compiler, preview, explicit Apply |

## Acceptance matrix

1. Selecting an existing PivotTable identifies its name, source type, and
   supported capabilities without creating another table.
2. Adding or reordering a normal field changes the native PivotTable and
   remains functional after refresh and reopen.
3. Adding SKU detail only beneath Family A leaves SKU absent beneath Family B
   and leaves Family B's total unchanged.
4. A scoped detail block can be placed between two normal children and keeps
   that position after refresh.
5. Value, percentage of parent, and percentage of total use the correct
   denominator in every displayed period; valid sibling portions total 100%.
6. Per-level subtotal and grand-total toggles affect only the requested native
   totals.
7. Top 2 scored on Current Year with Others may still display Prior Year, Plan,
   and Current Year; Others reconciles to the full total.
8. Actual Jan, Actual Feb, Actual Mar, Q1 Plan, and Q1 Variance appear in that
   exact order in one real PivotTable without monthly Plan columns.
9. Four named ratio measures remain present and visible after the pane closes,
   refreshes, and reopens.
10. A percentage-point delta between two calculated ratios equals Current
    Ratio minus Plan Ratio and has explicit blank-on-zero behavior.
11. Monthly source produces Q1 from Jan-Mar; yearly-only source requesting Q1
    is blocked with a grain error.
12. Report, row, scoped-detail, period, value, comparison, and scoring filters
    share their declared filter context.
13. Member search returns exact available members and does not report an empty
    list merely because the source is a Data Model.
14. Custom labels appear in the native Field List/PivotTable and persist.
15. A model-generated preview lists every field, measure, set, filter, and
    ordering change; the workbook is untouched before Apply.
16. A selected formatted blank range produces a geometry/label proposal or a
    visible ambiguity result.
17. Undo removes only artifacts from the last Apply and restores the prior
    PivotTable state.
18. Save, close, reopen, and refresh preserve the setup, owned measures, named
    sets, ordering, and native PivotTable behavior.
19. An unsupported construct explains the limitation and does not generate a
    formula table.
20. A representative large Data Model refresh is benchmarked against the same
    manually authored DAX/MDX definitions and stays within the release budget.

## Known native constraints

- MDX named sets require a Data Model or compatible OLAP PivotTable.
- Excel does not apply ordinary filtering directly to a named set; PivotTable+
  must regenerate an owned set when an advanced branch rule changes.
- External OLAP providers may restrict private calculations or sets.
- Ordinary PivotTables cannot be converted by toggling a PivotCache flag; Data
  Model enablement is an explicit migration with rollback.
- Visual behavior and performance on LTSC 2021 and Microsoft 365 require the
  generated synthetic smoke matrix before release.

## Non-goals

- Cloning or injecting controls into Microsoft's built-in Field List.
- A custom pivot aggregation engine.
- A formula-backed dense report as an automatic fallback.
- Arbitrary DAX, MDX, formulas, VBA, COM, shell, filesystem, workbook-save, or
  delivery capabilities exposed to a model.
