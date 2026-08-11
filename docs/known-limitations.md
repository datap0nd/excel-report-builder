# Known Limitations

This is an unsigned `v0.1.0` prototype. Use a copy of an important workbook
until the add-in has been validated in your own Excel environment.

The published `v0.1.0` installer predates the repaired Office COM callback
contracts in the current source. Use a later patch release when one is
available.

- Windows and classic desktop Excel are required. The add-in does not run in
  Excel for the web or on macOS.
- The supported source is one selected in-workbook table or rectangular range
  with one header row. Joins, appends, external files, databases, and SQL are
  intentionally excluded.
- Text period normalization accepts supported explicit formats in years 1900
  through 2099. Month headers without a year require user input.
- Manual preparation steps run before automatic wide-to-long normalization.
  Filter the generated Period or Metric through report Filters. Post-unpivot
  transformation editing is not exposed in the first manual UI.
- Calculated metrics require dense report blocks. Standard PivotTable and
  metric-stack blocks support direct aggregate Values.
- Report blocks reserve a user-visible managed rectangle. A build fails before
  writing outside that rectangle or when rectangles overlap.
- Saved setups that use advanced agent-only features remain exact and can be
  rebuilt or changed through Chat, but their bounded manual projection is
  visibly read-only. Representable manual setups with up to eight blocks reopen
  editable when every block shares the same Rows, Columns, Values, Filters, and
  layout controls.
- Data Model refresh behavior depends on the installed Excel build. Large
  sources are never truncated, but no universal completion time is promised.
- Excel exposes PivotCaches by workbook index but does not expose a cache name
  or delete operation. The add-in therefore records an exact managed cache
  identity, index, and source contract in workbook metadata. Normal rebuilds
  reuse that exclusive cache. Each block has one bounded cache slot for the
  worksheet route and one for the Data Model route, so switching routes and
  switching back reuses the prior compatible cache. A changed source contract,
  a shared cache, or a missing cache retires only the add-in's exact slot and
  creates a replacement; unmanaged caches are never modified. Excel may retain
  an inaccessible orphan cache after an external deletion or source replacement.
  Because deleting a managed PivotTable does not delete its cache, v0.1 retains
  these two exact cache-slot records for each historical managed block. Re-adding
  that block can reuse the cache instead of creating an untraceable duplicate.
- Managed Power Query and Data Model objects keep stable names and refresh in
  place on same-route rebuilds. When the route changes, the inactive managed
  backend can remain in the workbook because an Excel-owned cache may still
  depend on it. At most one exact owned canonical object is retained per backend.
- Setup is not Authenticode-signed. Verify the published SHA-256 checksum before
  running it.
- CI verifies builds, contracts, COM registration, worker startup, installation,
  and uninstall cleanup on Windows. Microsoft 365 desktop Excel startup has
  been field-tested with generated synthetic data; LTSC 2021 and a complete
  interactive managed-draft flow still require explicitly authorized field
  validation.
