# Known Limitations

This is an unsigned PivotTable+ prototype. Use a generated or disposable copy
of an important workbook until the add-in has been validated in your own Excel
environment.

- Windows and classic desktop Excel are required. Excel for the web and macOS
  do not expose the required Data Model, DAX, MDX, and COM APIs.
- Excel LTSC 2021 is the compatibility floor. Microsoft 365 desktop Excel is
  also supported, but both hosts remain part of the live release matrix.
- Standard Rows, Columns, Values, Filters, formatting, sorting, and
  expand/collapse remain Excel features. PivotTable+ does not inject controls
  into Microsoft's built-in Fields pane; it supplies a separate native task
  pane and a button that toggles the built-in pane.
- Advanced measures and asymmetric named sets require a Data Model PivotTable.
  A supported worksheet-backed PivotTable may be explicitly enabled. The
  operation can increase workbook size because the source becomes part of the
  workbook Data Model.
- Classic-to-Data-Model enablement deliberately refuses PivotTables whose
  filters, custom cell metadata, cache policy, calculated fields/items,
  grouped fields, slicers, charts, or formatting cannot be preserved exactly.
- Typed named sets support exact ordered tuples, scoped asymmetric branches,
  hierarchy defaults, and Top N. A labeled Others member is not supported;
  representing the complement as an existing All member would be incorrect.
- Named sets cannot receive Excel's normal native filters. PivotTable+ must
  regenerate an owned set when its typed member selection changes.
- Parent Portion currently requires the user to choose the numeric source
  column and the detail row field. It creates a workbook DAX measure and places
  it inside the selected native PivotTable.
- Undo for semantic measure/set changes is bounded to the current add-in
  session. Prior formulas are intentionally never stored in Custom XML.
- Unsupported or ambiguous workbook state fails closed. PivotTable+ never
  silently substitutes a formula table.
- The setup package is not Authenticode-signed. Verify its SHA-256 checksum
  before running it.
- Automated tests use generated, late-bound Excel-shaped hosts. They do not
  replace live save/reopen/refresh testing on both LTSC 2021 and Microsoft 365.
