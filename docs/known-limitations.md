# Known Limitations

This is an unsigned `v0.1.0` prototype. Use a copy of an important workbook
until the add-in has been validated in your own Excel environment.

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
- Data Model refresh behavior depends on the installed Excel build. Large
  sources are never truncated, but no universal completion time is promised.
- Setup is not Authenticode-signed. Verify the published SHA-256 checksum before
  running it.
- CI verifies builds, contracts, COM registration, worker startup, installation,
  and uninstall cleanup on Windows. A real Excel field test is intentionally
  deferred until a user explicitly authorizes a managed workstation session.
