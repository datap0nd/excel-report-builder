# Product

<!-- impeccable:product-schema 1 -->

## Platform

windows

## Stack

Delegated: a Windows-native C# COM add-in targeting .NET Framework 4.8, with a
WPF interface hosted in a COM-compatible task pane, a self-contained local
worker, per-user installation, and Excel LTSC 2021 as the primary runtime.

## Users

Excel-fluent analysts who receive recurring raw extracts and must turn them
into dense, formatted management reports with nested rows, period columns,
subtotals, ratios, comparisons, and checks. They work inside managed Windows
environments and expect Excel-native behavior rather than database or
developer terminology.

## Product Purpose

Excel Report Builder converts one raw extract in the active workbook into one
repeatable, validated report setup. Users can configure it manually or ask a
local AI agent to build the same configuration. Success means the source stays
unchanged, wide periods normalize without truncation, generated reports remain
traceable to native Excel pivots, and every completed result passes explicit
checks.

## Positioning

The manual builder and the AI agent edit one bounded, versioned report
specification. The model never writes arbitrary formulas or receives general
Excel automation. A deterministic host owns transformation, PivotTable
construction, report rendering, validation, and managed-draft publishing.

## Operating Context

- One raw extract is already present as a table or rectangular range in Excel.
- Sources may contain a date column or one-row wide period headers such as
  month names and metric-month combinations.
- Finished workbooks use dense row and column hierarchies rather than dashboard
  cards.
- Long-running builds may continue beyond thirty minutes and must provide
  continuous, specific progress, cancellation, checkpoints, and a final audit.
- The product operates only on managed draft objects until the user publishes.

## Capabilities and Constraints

- No SQL, database sources, joins, email, scheduling, external-file automation,
  or automatic workbook saves in version 1.
- Supports manual Data, Rows, Columns, Values, Filters, transformations,
  subtotals, ordering, metrics, formatting, Preview, and Checks.
- Supports local or explicitly configured OpenAI-compatible AI endpoints with
  model discovery, capability checks, protected credentials, and explicit
  consent before a remote endpoint receives bounded workbook samples.
- Uses native PivotTables for aggregation and controlled report formulas for
  dense layouts.
- Uses Excel's built-in data preparation and Data Model internally when a
  normalized result exceeds worksheet capacity.
- Never silently truncates source data or invents a missing year.
- Windows and Excel LTSC 2021 are the primary production environment.

## Brand Commitments

- Product name: Excel Report Builder.
- Public vocabulary: dense management report, Data, Rows, Columns, Values,
  Filters, Preview, Checks, and Saved report setup.
- The interface is precise, operational, and spreadsheet-native.
- No emoji and no decorative marketing language.

## Evidence on Hand

The repository contains no real workbook, customer data, screenshot, or
organization-specific report. All demonstrations and tests use generated
synthetic data.

## Product Principles

1. One specification powers both manual configuration and AI assistance.
2. Pivots aggregate; deterministic layout code presents; independent checks
   decide whether a result is trustworthy.
3. Source data is preserved and every generated Excel object is explicitly
   owned.
4. Ambiguity is shown and resolved, never hidden behind a confident guess.
5. Long work continuously reports what it is doing and remains cancellable.

## Accessibility & Inclusion

All task-pane actions must be keyboard accessible, expose useful accessible
names, preserve visible focus, avoid color-only status, and remain usable at
the narrow width of a docked Excel task pane.
