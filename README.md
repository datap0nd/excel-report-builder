# Excel Report Builder

Build validated dense management reports inside Microsoft Excel with manual
controls and an optional local AI agent.

Excel Report Builder is a Windows add-in for analysts who receive a raw extract
and need a repeatable report with period columns, nested rows, subtotals,
ratios, comparisons, and explicit checks. The source remains unchanged. Native
PivotTables perform base aggregation, while a deterministic layout layer builds
the finished management report in a managed draft.

## Project status

The `v0.1.0` prototype implements:

- long-date and wide Jan-Dec source detection;
- one-row metric-month normalization;
- worksheet and Excel Data Model paths without silent truncation;
- manual Data, Rows, Columns, Values, Filters, formatting, and checks;
- dense report blocks backed by native PivotTables;
- an OpenAI-compatible local agent that edits the same report specification;
- continuous, specific build feedback and cancellation;
- managed drafts that require a user action before publishing.

It does not use SQL, join unrelated sources, save workbooks automatically, or
give the model arbitrary Excel, formula, VBA, shell, or filesystem access.

## Install

Download the unsigned setup executable and matching SHA-256 file from
[GitHub Releases](https://github.com/datap0nd/excel-report-builder/releases).
Verify the checksum, close Excel, and run setup. Installation is per user and
does not request administrator rights. The prototype is unsigned, so Windows
will show an unknown-publisher warning.

> **Release note:** the published `v0.1.0` installer predates the Office COM
> callback ABI repair in the current source. Use a later patch release when
> available; do not reinstall the original asset over a repaired development
> installation.

Open Excel and choose **Excel Report Builder** on the ribbon. The default path
is manual and does not need a model. Chat requires an OpenAI-compatible
endpoint and an editable model ID. No model is bundled.

## Workflow

1. Select one table or rectangular range with one header row.
2. Confirm the detected long or wide period layout. Month-only headers require
   an explicit reporting year.
3. Add bounded preparation steps, then place fields in Rows, Columns, Values,
   and Filters.
4. Add calculated metrics, report blocks, layout controls, and report-specific
   checks, or ask Chat to propose the same typed report setup.
5. Build a managed draft, review the activity timeline and checks, then publish
   with an explicit click. The add-in never saves the workbook automatically.

## Platform

- Windows 10 or newer
- Microsoft Excel LTSC 2021 or compatible classic desktop Excel
- 32-bit or 64-bit Office
- .NET Framework 4.8 for the in-process add-in
- Optional OpenAI-compatible model endpoint for Chat

## Development

Windows development prerequisites:

- Visual Studio 2022 Build Tools with the .NET desktop workload
- .NET 8 SDK
- .NET Framework 4.8 targeting pack
- Inno Setup 6 for installer builds

Build and test:

```powershell
dotnet restore ExcelReportBuilder.sln
dotnet build ExcelReportBuilder.sln -c Release --no-restore
dotnet test ExcelReportBuilder.sln -c Release --no-build
powershell -NoProfile -File scripts/Test-PublicSafety.ps1
```

The current synthetic suite covers source profiling, period normalization,
Power Query generation, pivot planning, dense rendering, reconciliation,
ownership, publishing, agent permissions, endpoint policy, progress ordering,
checkpointing, cancellation, and resume behavior.

See [known limitations](docs/known-limitations.md) before using the unsigned
prototype on important workbooks.

## Privacy and safety

The repository contains generated synthetic fixtures only. Do not attach or
commit real workbooks, exports, screenshots, prompts, transcripts, endpoints,
credentials, or company-specific report material. See [SECURITY.md](SECURITY.md)
and [docs/public-safety.md](docs/public-safety.md).

## License

MIT
