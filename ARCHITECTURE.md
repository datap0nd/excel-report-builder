# Architecture

## Trust boundary

Excel Report Builder separates model inference from Excel mutation.

```text
Excel ribbon and task pane
        |
        v
Validated ReportSpecV1
        |
        +----> deterministic transform and report compiler
        |                |
        |                v
        |        owned Excel objects only
        |
        +----> current-user named pipe ----> local agent worker
                                              |
                                              v
                                  OpenAI-compatible endpoint
```

The worker never receives an Excel application object and cannot execute
formulas, VBA, COM calls, shell commands, workbook saves, publishing, deletion,
or arbitrary filesystem operations. It proposes only typed report-specification
changes and allowlisted job operations. Its only local write is bounded,
non-sensitive checkpoint metadata under the current user's application-data
directory. The add-in validates every request and owns the complete Excel
execution path.

## Projects

- `ExcelReportBuilder.Core`: versioned specifications, source profiling,
  transformation plans, measure expressions, compilation, and validation.
- `ExcelReportBuilder.Excel`: late-bound Excel execution, object ownership,
  PivotTables, managed drafts, persistence, publishing, and rollback.
- `ExcelReportBuilder.Agent`: endpoint policy, tool protocol, job state, model
  client, and agent orchestration shared by the add-in and worker.
- `ExcelReportBuilder.Worker`: out-of-process model and planning host.
- `ExcelReportBuilder.AddIn`: COM entry point, ribbon, task pane, manual builder,
  Chat, Checks, progress, and cancellation.

## Data flow

1. The user selects a table or rectangular source range.
2. The source profiler inspects headers and bounded samples without modifying
   values.
3. Period detection proposes a long-date or wide metric-month mapping.
4. A typed transformation plan compiles to a restricted workbook query that
   can reference only the selected source.
5. Normalized data loads to a managed worksheet table when it fits, otherwise
   to Excel's Data Model.
6. The report compiler creates a backend-neutral pivot and layout plan.
7. The Excel executor builds hidden or visible native PivotTables and renders
   only managed draft sheets.
8. An independent validation plan reconciles row counts, totals, periods,
   subtotals, ratios, formulas, and refresh state.
9. Publishing remains a user action and never saves the workbook.

## Large-data behavior

Projected normalized row count determines the backend before a refresh begins.
The worksheet backend is used only when the complete result fits. Larger
results use the Data Model. A successful job always records source,
normalization, pivot, and finished-output row counts. Truncation is never an
accepted fallback.

## Continuous feedback

Every operation emits typed progress events with stage, message, elapsed time,
optional counts, object identity, and completed checks. A heartbeat fills quiet
periods while the host remains responsive. Progress is part of the protocol and
test contract, not task-pane decoration.
