# PivotTable+

## Download

**[Download PivotTable+ for Windows (.exe)](https://github.com/datap0nd/excel-report-builder/releases/download/v0.2.0/ExcelReportBuilderSetup-0.2.0-unsigned.exe)**

This is an unsigned prototype for Excel LTSC 2021 and compatible Microsoft 365
desktop Excel. Close Excel before running the installer; no PC restart is
required.

PivotTable+ is a Windows Excel add-in that enhances a real native PivotTable.
Excel's PivotTable, Analyze/Design tabs, refresh behavior, and built-in Fields
pane remain usable. The add-in supplies a compact companion pane for previewed
field placement and advanced features such as parent portions, typed Data Model
measures, and asymmetric named sets.

The default result is always one refreshable PivotTable. Supported features do
not create a formula-backed companion report.

## Current status

The development branch includes:

- selected-PivotTable discovery for worksheet, Data Model, and external OLAP
  sources;
- a familiar Rows, Columns, Values, and Filters pane with explicit preview and
  Apply;
- a direct button for Excel's built-in PivotTable Fields pane;
- verified classic-to-Data-Model enablement with durable recovery;
- typed DAX measures for ratios, portions, variance, comparisons, and period
  slices;
- typed MDX named sets for exact ordering and asymmetric branches;
- one combined transactional refresh, verification, rollback, and session Undo;
- per-workbook ownership metadata containing hashes and identifiers, never DAX,
  MDX, workbook paths, prompts, or source values.

The package is an unsigned prototype. Use generated or disposable workbooks
until the live-host release matrix is complete.

## Install

Download the setup executable above and matching SHA-256 file from
[GitHub Releases](https://github.com/datap0nd/excel-report-builder/releases),
verify the checksum, close Excel, and run setup. Installation is per-user and
does not require administrator rights. Windows may show an unknown-publisher
warning because the prototype is not Authenticode-signed.

Open Excel and choose **PivotTable+** on the Data tab. Select a cell inside an
existing PivotTable and choose Refresh in the pane. Standard edits may be made
with Excel's native Field List or previewed in PivotTable+. Advanced extras
require a Data Model PivotTable; the pane offers an explicit enablement action
for supported classic worksheet-backed PivotTables.

## Platform

- Windows 10 or newer
- Microsoft Excel LTSC 2021 or compatible Microsoft 365 desktop Excel
- 32-bit or 64-bit Office
- .NET Framework 4.8 for the in-process add-in

Excel for the web and macOS are outside the first package because the required
Data Model, DAX, MDX, and COM APIs are desktop capabilities.

## Development

```powershell
dotnet restore ExcelReportBuilder.sln
dotnet build ExcelReportBuilder.sln -c Release --no-restore
dotnet test ExcelReportBuilder.sln -c Release --no-build
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/Test-TaskPaneContract.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/Test-PublicSafety.ps1
```

See [PRODUCT.md](PRODUCT.md) for the native-first product contract and
[docs/known-limitations.md](docs/known-limitations.md) before using the unsigned
prototype on important workbooks.

## Privacy and safety

This public repository uses generated synthetic fixtures only. Do not commit
real workbooks, exports, screenshots, prompts, transcripts, endpoints,
credentials, or company-specific material. See [SECURITY.md](SECURITY.md) and
[docs/public-safety.md](docs/public-safety.md).

## License

MIT
