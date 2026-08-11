using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ExcelReportBuilder.Core.Measures;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Core.Validation;

namespace ExcelReportBuilder.AddIn.Host
{
    /// <summary>
    /// Applies the bounded manual layout controls to an already configured
    /// report block. The configured block remains untouched and is explicitly
    /// cloned for every independently managed output block.
    /// </summary>
    internal sealed class ManualLayoutTranslator
    {
        private const int MaximumManualBlocks = 8;
        private const int ExcelMaximumRows = 1048576;
        private const int ExcelMaximumColumns = 16384;

        private static readonly Regex CellPattern = new Regex(
            @"^\$?(?<column>[A-Za-z]{1,3})\$?(?<row>[1-9][0-9]{0,6})$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        /// <summary>
        /// Replaces the report blocks and configured checks with validated
        /// manual selections. The mutation is rolled back if the resulting
        /// ReportSpecV1 fails structural validation.
        /// </summary>
        public ReportSpecV1 Apply(
            ReportBlockSpec configuredTemplate,
            IReadOnlyList<ManualReportBlockSnapshot> blocks,
            ManualLayoutSnapshot layout,
            IReadOnlyList<ManualCheckSnapshot> checks,
            ReportSpecV1 specification)
        {
            if (configuredTemplate == null)
            {
                throw new ArgumentNullException(nameof(configuredTemplate));
            }

            if (blocks == null)
            {
                throw new ArgumentNullException(nameof(blocks));
            }

            if (layout == null)
            {
                throw new ArgumentNullException(nameof(layout));
            }

            if (checks == null)
            {
                throw new ArgumentNullException(nameof(checks));
            }

            if (specification == null)
            {
                throw new ArgumentNullException(nameof(specification));
            }

            if (!string.Equals(
                    specification.SchemaVersion,
                    ReportSpecV1.CurrentSchemaVersion,
                    StringComparison.Ordinal))
            {
                throw new NotSupportedException(
                    "Manual layout cannot be applied to an unsupported report specification version.");
            }

            ValidateTemplate(configuredTemplate);
            ManualLayoutValues layoutValues = ValidateLayout(layout);
            List<ReportBlockSpec> translatedBlocks = TranslateBlocks(
                configuredTemplate,
                blocks,
                layoutValues,
                specification.OwnershipId);
            List<ReportCheckSpec> translatedChecks = TranslateChecks(checks, specification.Measures);

            List<ReportBlockSpec> previousBlocks = specification.Blocks;
            List<ReportCheckSpec> previousChecks = specification.Checks;
            specification.Blocks = translatedBlocks;
            specification.Checks = translatedChecks;
            try
            {
                ValidationResult validation = ReportSpecValidator.Validate(specification);
                if (!validation.IsValid)
                {
                    throw new InvalidOperationException(FormatValidationFailure(validation));
                }
            }
            catch
            {
                specification.Blocks = previousBlocks;
                specification.Checks = previousChecks;
                throw;
            }

            return specification;
        }

        private static List<ReportBlockSpec> TranslateBlocks(
            ReportBlockSpec template,
            IReadOnlyList<ManualReportBlockSnapshot> snapshots,
            ManualLayoutValues layout,
            string rootOwnershipId)
        {
            if (snapshots.Count < 1 || snapshots.Count > MaximumManualBlocks)
            {
                throw new InvalidOperationException(
                    "Manual layout requires between 1 and 8 report blocks.");
            }

            var result = new List<ReportBlockSpec>(snapshots.Count);
            var occupied = new List<OwnedRectangle>();
            var blockIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var ownershipIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (!string.IsNullOrWhiteSpace(rootOwnershipId))
            {
                ownershipIds.Add(rootOwnershipId);
            }

            for (var index = 0; index < snapshots.Count; index++)
            {
                ManualReportBlockSnapshot snapshot = snapshots[index]
                    ?? throw new InvalidOperationException(
                        "Report block " + (index + 1).ToString(CultureInfo.InvariantCulture) + " is missing.");
                string path = "Report block " + (index + 1).ToString(CultureInfo.InvariantCulture);
                string worksheetName = ValidateWorksheetName(snapshot.WorksheetName, path);
                ParsedCell anchor = ParseCell(snapshot.AnchorCell, path);
                string title = ValidateOptionalTitle(snapshot.Title, path);
                ReportOutputMode outputMode = ParseOutputMode(snapshot.OutputStyle, path);
                int ownedRows = ValidateOwnedExtent(
                    snapshot.OwnedRows,
                    ExcelMaximumRows,
                    path + " managed rows");
                int ownedColumns = ValidateOwnedExtent(
                    snapshot.OwnedColumns,
                    ExcelMaximumColumns,
                    path + " managed columns");

                var block = CloneBlock(template);
                if (string.IsNullOrWhiteSpace(snapshot.StableId))
                {
                    block.Id = CreateLegacyBlockId(
                        template.Id,
                        "report_block",
                        index,
                        blockIds,
                        path + " identifier");
                    block.OwnershipId = CreateLegacyBlockId(
                        template.OwnershipId,
                        "managed_block_owner",
                        index,
                        ownershipIds,
                        path + " ownership identifier");
                }
                else
                {
                    block.Id = CreateStableManagedId(
                        snapshot.StableId!,
                        "report_block",
                        blockIds,
                        path + " stable identifier");
                    block.OwnershipId = CreateStableManagedId(
                        "owned_" + snapshot.StableId,
                        "managed_block_owner",
                        ownershipIds,
                        path + " ownership identifier");
                }

                block.Title = string.IsNullOrEmpty(title) ? null : title;
                block.WorksheetName = worksheetName;
                block.AnchorCell = anchor.NormalizedAddress;
                block.OutputMode = outputMode;
                block.OwnedExtent = new OwnedRangeExtentSpec
                {
                    RowCount = ownedRows,
                    ColumnCount = ownedColumns
                };
                block.Layout.DenseLayout = new DenseLayoutOptions
                {
                    RepeatRowLabels = layout.RepeatRowLabels,
                    InsertBlankRows = layout.InsertBlankRows,
                    FreezeHeaders = layout.FreezeHeaders,
                    ShowRowGrandTotals = layout.ShowRowGrandTotals,
                    ShowColumnGrandTotals = layout.ShowColumnGrandTotals,
                    RowIndent = layout.RowIndent
                };
                GrandTotalsSpec templateTotals = template.Layout.GrandTotals;
                block.Layout.GrandTotals = new GrandTotalsSpec
                {
                    ShowRows = layout.ShowRowGrandTotals,
                    ShowColumns = layout.ShowColumnGrandTotals,
                    RowPlacement = templateTotals.RowPlacement,
                    ColumnPlacement = templateTotals.ColumnPlacement,
                    RowLabel = layout.RowGrandTotalLabel,
                    ColumnLabel = layout.ColumnGrandTotalLabel,
                    StyleId = templateTotals.StyleId
                };

                EnsureExtentWithinWorksheet(block, anchor, path);
                var rectangle = new OwnedRectangle(
                    worksheetName,
                    anchor.Row,
                    checked(anchor.Row + block.OwnedExtent.RowCount - 1),
                    anchor.Column,
                    checked(anchor.Column + block.OwnedExtent.ColumnCount - 1),
                    path);
                OwnedRectangle? overlap = occupied.FirstOrDefault(item => item.Overlaps(rectangle));
                if (overlap != null)
                {
                    throw new InvalidOperationException(
                        path + " owns " + rectangle.DisplayAddress + ", which overlaps " +
                        overlap.Label + " at " + overlap.DisplayAddress + " on worksheet '" +
                        worksheetName + "'. Move one anchor so the managed ranges do not overlap.");
                }

                occupied.Add(rectangle);
                result.Add(block);
            }

            return result;
        }

        private static List<ReportCheckSpec> TranslateChecks(
            IReadOnlyList<ManualCheckSnapshot> snapshots,
            List<MeasureDefinition> measures)
        {
            if (measures == null)
            {
                throw new InvalidOperationException(
                    "The configured report has no measure collection for manual checks.");
            }

            MeasureLookup lookup = BuildMeasureLookup(measures);
            var result = new List<ReportCheckSpec>(snapshots.Count);
            var semanticChecks = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var checkIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (var index = 0; index < snapshots.Count; index++)
            {
                ManualCheckSnapshot snapshot = snapshots[index]
                    ?? throw new InvalidOperationException(
                        "Check " + (index + 1).ToString(CultureInfo.InvariantCulture) + " is missing.");
                string path = "Check " + (index + 1).ToString(CultureInfo.InvariantCulture);
                if (snapshot.Tolerance < 0m)
                {
                    throw new InvalidOperationException(path + " tolerance cannot be negative.");
                }

                ParsedCheckKind parsedKind = ParseCheckKind(snapshot.Kind, path);
                string? measureId;
                string? comparedMeasureId;
                if (parsedKind.Kind == ReportCheckKind.NoTruncation)
                {
                    RejectMetric(snapshot.Metric, path, "Metric");
                    RejectMetric(snapshot.ComparedMetric, path, "Compared metric");
                    measureId = null;
                    comparedMeasureId = null;
                }
                else if (parsedKind.Kind == ReportCheckKind.Balance)
                {
                    measureId = ResolveMetric(snapshot.Metric, path, "Metric", lookup);
                    comparedMeasureId = ResolveMetric(
                        snapshot.ComparedMetric,
                        path,
                        "Compared metric",
                        lookup);
                    if (string.Equals(measureId, comparedMeasureId, StringComparison.OrdinalIgnoreCase))
                    {
                        throw new InvalidOperationException(
                            path + " must compare two different metrics.");
                    }
                }
                else
                {
                    measureId = ResolveMetric(snapshot.Metric, path, "Metric", lookup);
                    RejectMetric(snapshot.ComparedMetric, path, "Compared metric");
                    comparedMeasureId = null;
                    if (parsedKind.Kind == ReportCheckKind.TotalPreservation)
                    {
                        MeasureDefinition measure = lookup.ById[measureId];
                        if (!(measure.Expression is AggregateMeasureExpression aggregate) ||
                            aggregate.Function != AggregateFunction.Sum ||
                            !string.IsNullOrWhiteSpace(aggregate.PeriodSliceId))
                        {
                            throw new InvalidOperationException(
                                path + " total preservation requires a direct, unsliced Sum metric.");
                        }
                    }
                }

                string semanticKey = parsedKind.SemanticName + "|" +
                    (measureId ?? string.Empty) + "|" + (comparedMeasureId ?? string.Empty);
                if (!semanticChecks.Add(semanticKey))
                {
                    throw new InvalidOperationException(
                        path + " duplicates an earlier " + parsedKind.DisplayName + " check.");
                }

                result.Add(new ReportCheckSpec
                {
                    Id = CreateUniqueIndexedId(
                        "manual_" + parsedKind.SemanticName,
                        "manual_check",
                        index + 1,
                        checkIds),
                    Kind = parsedKind.Kind,
                    MeasureId = measureId,
                    ComparedMeasureId = comparedMeasureId,
                    Tolerance = snapshot.Tolerance
                });
            }

            return result;
        }

        private static ParsedCheckKind ParseCheckKind(string value, string path)
        {
            string normalized = NormalizeOption(value);
            switch (normalized)
            {
                case "total preservation":
                    return new ParsedCheckKind(
                        ReportCheckKind.TotalPreservation,
                        "total_preservation",
                        "total preservation");
                case "no truncation":
                    return new ParsedCheckKind(
                        ReportCheckKind.NoTruncation,
                        "no_truncation",
                        "no truncation");
                case "required values":
                    return new ParsedCheckKind(
                        ReportCheckKind.RequiredValues,
                        "required_values",
                        "required values");
                case "non negative":
                    return new ParsedCheckKind(
                        ReportCheckKind.NonNegative,
                        "non_negative",
                        "non-negative");
                case "balance":
                    return new ParsedCheckKind(
                        ReportCheckKind.Balance,
                        "balance",
                        "balance");
                case "rendered output":
                    // ReportSpecV1 currently represents a metric-level rendered
                    // reconciliation with the typed TotalPreservation kind. The
                    // planner also adds its mandatory block-level rendered-output
                    // reconciliation independently.
                    return new ParsedCheckKind(
                        ResolveRenderedOutputKind(),
                        "rendered_output",
                        "rendered output");
                default:
                    throw new InvalidOperationException(
                        path + " has unsupported kind '" + (value ?? string.Empty) +
                        "'. Choose Total preservation, No truncation, Required values, " +
                        "Non-negative, Balance, or Rendered output.");
            }
        }

        private static ReportCheckKind ResolveRenderedOutputKind()
        {
            ReportCheckKind parsed;
            return Enum.TryParse("RenderedOutput", ignoreCase: false, out parsed)
                && Enum.IsDefined(typeof(ReportCheckKind), parsed)
                    ? parsed
                    : ReportCheckKind.TotalPreservation;
        }

        private static MeasureLookup BuildMeasureLookup(List<MeasureDefinition> measures)
        {
            var byId = new Dictionary<string, MeasureDefinition>(StringComparer.OrdinalIgnoreCase);
            var byLabel = new Dictionary<string, List<MeasureDefinition>>(StringComparer.OrdinalIgnoreCase);
            for (var index = 0; index < measures.Count; index++)
            {
                MeasureDefinition measure = measures[index]
                    ?? throw new InvalidOperationException(
                        "Configured metric " + (index + 1).ToString(CultureInfo.InvariantCulture) + " is missing.");
                if (string.IsNullOrWhiteSpace(measure.Id))
                {
                    throw new InvalidOperationException(
                        "Every configured metric requires a stable identifier before checks can be added.");
                }

                if (byId.ContainsKey(measure.Id))
                {
                    throw new InvalidOperationException(
                        "Configured metric identifier '" + measure.Id + "' is duplicated.");
                }

                byId.Add(measure.Id, measure);

                string label = (measure.Label ?? string.Empty).Trim();
                if (label.Length == 0)
                {
                    continue;
                }

                List<MeasureDefinition> matches;
                if (!byLabel.TryGetValue(label, out matches!))
                {
                    matches = new List<MeasureDefinition>();
                    byLabel.Add(label, matches);
                }

                matches.Add(measure);
            }

            return new MeasureLookup(byId, byLabel);
        }

        private static string ResolveMetric(
            string value,
            string path,
            string fieldLabel,
            MeasureLookup lookup)
        {
            string candidate = (value ?? string.Empty).Trim();
            if (candidate.Length == 0)
            {
                throw new InvalidOperationException(
                    path + " requires a " + fieldLabel.ToLowerInvariant() + ".");
            }

            MeasureDefinition direct;
            if (lookup.ById.TryGetValue(candidate, out direct!))
            {
                return direct.Id;
            }

            List<MeasureDefinition> matches;
            if (!lookup.ByLabel.TryGetValue(candidate, out matches!) || matches.Count == 0)
            {
                throw new InvalidOperationException(
                    path + " references unknown " + fieldLabel.ToLowerInvariant() + " '" + candidate +
                    "'. Choose a configured metric label or identifier.");
            }

            if (matches.Count > 1)
            {
                throw new InvalidOperationException(
                    path + " uses ambiguous metric label '" + candidate +
                    "'. Use the metric identifier instead.");
            }

            return matches[0].Id;
        }

        private static void RejectMetric(string value, string path, string fieldLabel)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                throw new InvalidOperationException(
                    path + " does not accept a " + fieldLabel.ToLowerInvariant() + ".");
            }
        }

        private static void ValidateTemplate(ReportBlockSpec template)
        {
            if (template.Layout == null)
            {
                throw new InvalidOperationException(
                    "The configured report block has no layout to copy.");
            }

            if (template.Layout.Rows == null || template.Layout.Columns == null ||
                template.Layout.Values == null || template.Layout.Filters == null)
            {
                throw new InvalidOperationException(
                    "The configured report block has an incomplete Rows, Columns, Values, or Filters collection.");
            }

            if (template.Layout.GrandTotals == null)
            {
                throw new InvalidOperationException(
                    "The configured report block has no grand-total settings to copy.");
            }

            if (template.OwnedExtent == null ||
                template.OwnedExtent.RowCount < 1 || template.OwnedExtent.ColumnCount < 1)
            {
                throw new InvalidOperationException(
                    "The configured report block requires a positive managed row and column extent.");
            }

            if (template.PeriodSlices == null || template.Headers == null || template.Spacers == null)
            {
                throw new InvalidOperationException(
                    "The configured report block has an incomplete period, header, or spacer collection.");
            }
        }

        private static ManualLayoutValues ValidateLayout(ManualLayoutSnapshot layout)
        {
            if (layout.RowIndent < 0 || layout.RowIndent > 15)
            {
                throw new InvalidOperationException("Row indent must be between 0 and 15.");
            }

            string rowLabel = ValidateRequiredLabel(
                layout.RowGrandTotalLabel,
                "Row grand-total label");
            string columnLabel = ValidateRequiredLabel(
                layout.ColumnGrandTotalLabel,
                "Column grand-total label");
            return new ManualLayoutValues(
                layout.RepeatRowLabels,
                layout.InsertBlankRows,
                layout.FreezeHeaders,
                layout.ShowRowGrandTotals,
                layout.ShowColumnGrandTotals,
                layout.RowIndent,
                rowLabel,
                columnLabel);
        }

        private static string ValidateRequiredLabel(string value, string label)
        {
            string normalized = (value ?? string.Empty).Trim();
            if (normalized.Length == 0)
            {
                throw new InvalidOperationException(label + " cannot be blank.");
            }

            if (normalized.Length > 120 || normalized.Any(char.IsControl))
            {
                throw new InvalidOperationException(
                    label + " must contain at most 120 visible characters.");
            }

            return normalized;
        }

        private static string ValidateOptionalTitle(string value, string path)
        {
            string normalized = (value ?? string.Empty).Trim();
            if (normalized.Length > 120 || normalized.Any(char.IsControl))
            {
                throw new InvalidOperationException(
                    path + " title must contain at most 120 visible characters.");
            }

            return normalized;
        }

        private static string ValidateWorksheetName(string value, string path)
        {
            string normalized = (value ?? string.Empty).Trim();
            if (normalized.Length == 0)
            {
                throw new InvalidOperationException(path + " requires a worksheet name.");
            }

            if (normalized.Length > 31)
            {
                throw new InvalidOperationException(
                    path + " worksheet name must contain at most 31 characters.");
            }

            if (normalized.Any(char.IsControl) ||
                normalized.IndexOfAny(new[] { '[', ']', ':', '*', '?', '/', '\\' }) >= 0 ||
                normalized[0] == '\'' || normalized[normalized.Length - 1] == '\'')
            {
                throw new InvalidOperationException(
                    path + " worksheet name contains a character Excel does not allow.");
            }

            if (string.Equals(normalized, "History", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    path + " cannot use Excel's reserved worksheet name 'History'.");
            }

            return normalized;
        }

        private static ParsedCell ParseCell(string value, string path)
        {
            string normalized = (value ?? string.Empty).Trim();
            Match match = CellPattern.Match(normalized);
            int row;
            if (!match.Success ||
                !int.TryParse(
                    match.Groups["row"].Value,
                    NumberStyles.None,
                    CultureInfo.InvariantCulture,
                    out row))
            {
                throw new InvalidOperationException(
                    path + " anchor must be an A1-style cell address within the worksheet.");
            }

            int column = ColumnNumber(match.Groups["column"].Value);
            if (row < 1 || row > ExcelMaximumRows || column < 1 || column > ExcelMaximumColumns)
            {
                throw new InvalidOperationException(
                    path + " anchor must be within A1:XFD1048576.");
            }

            return new ParsedCell(
                row,
                column,
                ColumnLetters(column) + row.ToString(CultureInfo.InvariantCulture));
        }

        private static void EnsureExtentWithinWorksheet(
            ReportBlockSpec block,
            ParsedCell anchor,
            string path)
        {
            long finalRow = (long)anchor.Row + block.OwnedExtent.RowCount - 1L;
            long finalColumn = (long)anchor.Column + block.OwnedExtent.ColumnCount - 1L;
            if (finalRow > ExcelMaximumRows || finalColumn > ExcelMaximumColumns)
            {
                throw new InvalidOperationException(
                    path + " managed extent runs past the worksheet boundary. " +
                    "Choose an earlier anchor or reduce the configured managed extent.");
            }
        }

        private static int ValidateOwnedExtent(int value, int maximum, string label)
        {
            if (value < 1 || value > maximum)
            {
                throw new InvalidOperationException(
                    label + " must be between 1 and " +
                    maximum.ToString("N0", CultureInfo.InvariantCulture) + ".");
            }

            return value;
        }

        private static ReportOutputMode ParseOutputMode(string value, string path)
        {
            switch (NormalizeOption(value))
            {
                case "standard matrix": return ReportOutputMode.StandardMatrix;
                case "metric stack": return ReportOutputMode.MetricStack;
                case "dense management block": return ReportOutputMode.DenseGrid;
                default:
                    throw new InvalidOperationException(
                        path + " has unsupported output style '" + (value ?? string.Empty) +
                        "'. Choose Standard matrix, Metric stack, or Dense management block.");
            }
        }

        private static string NormalizeOption(string value)
        {
            string source = (value ?? string.Empty).Trim().Replace('-', ' ');
            var result = new StringBuilder(source.Length);
            var previousWhitespace = false;
            foreach (char character in source)
            {
                if (char.IsWhiteSpace(character))
                {
                    if (!previousWhitespace)
                    {
                        result.Append(' ');
                    }

                    previousWhitespace = true;
                }
                else
                {
                    result.Append(char.ToLowerInvariant(character));
                    previousWhitespace = false;
                }
            }

            return result.ToString();
        }

        private static string CreateUniqueIndexedId(
            string candidate,
            string fallback,
            int index,
            ISet<string> existing)
        {
            string baseValue = ToIdentifier(candidate, fallback);
            var attempt = index;
            while (true)
            {
                string suffix = "_" + attempt.ToString(CultureInfo.InvariantCulture);
                int maximumBaseLength = 64 - suffix.Length;
                string boundedBase = baseValue.Length <= maximumBaseLength
                    ? baseValue
                    : baseValue.Substring(0, maximumBaseLength).TrimEnd('_', '-');
                if (boundedBase.Length == 0 || !IsAsciiLetter(boundedBase[0]))
                {
                    boundedBase = "b";
                }

                string value = boundedBase + suffix;
                if (existing.Add(value))
                {
                    return value;
                }

                attempt++;
            }
        }

        private static string CreateLegacyBlockId(
            string candidate,
            string fallback,
            int zeroBasedIndex,
            ISet<string> existing,
            string label)
        {
            if (zeroBasedIndex != 0)
            {
                return CreateUniqueIndexedId(
                    candidate,
                    fallback,
                    zeroBasedIndex + 1,
                    existing);
            }

            string value = BoundIdentifier(ToIdentifier(candidate, fallback));
            if (!existing.Add(value))
            {
                throw new InvalidOperationException(
                    label + " conflicts with another managed identifier.");
            }

            return value;
        }

        private static string CreateStableManagedId(
            string candidate,
            string fallback,
            ISet<string> existing,
            string label)
        {
            string value = BoundIdentifier(ToIdentifier(candidate, fallback));
            if (!existing.Add(value))
            {
                throw new InvalidOperationException(
                    label + " is duplicated. Remove and add the duplicate block again.");
            }

            return value;
        }

        private static string BoundIdentifier(string value)
        {
            string bounded = value.Length <= 64
                ? value
                : value.Substring(0, 64).TrimEnd('_', '-');
            return bounded.Length == 0 || !IsAsciiLetter(bounded[0])
                ? "b_" + bounded
                : bounded;
        }

        private static string ToIdentifier(string value, string fallback)
        {
            var result = new StringBuilder();
            foreach (char character in value ?? string.Empty)
            {
                if (IsAsciiLetter(character) || (character >= '0' && character <= '9') ||
                    character == '_' || character == '-')
                {
                    result.Append(character);
                }
                else
                {
                    result.Append('_');
                }
            }

            string normalized = result.ToString().Trim('_', '-');
            if (normalized.Length == 0)
            {
                normalized = fallback;
            }

            if (!IsAsciiLetter(normalized[0]))
            {
                normalized = "b_" + normalized;
            }

            return normalized;
        }

        private static bool IsAsciiLetter(char value)
        {
            return value >= 'A' && value <= 'Z' || value >= 'a' && value <= 'z';
        }

        private static int ColumnNumber(string letters)
        {
            var result = 0;
            foreach (char character in letters.ToUpperInvariant())
            {
                result = checked(result * 26 + character - 'A' + 1);
            }

            return result;
        }

        private static string ColumnLetters(int column)
        {
            var value = column;
            var result = string.Empty;
            while (value > 0)
            {
                value--;
                result = (char)('A' + value % 26) + result;
                value /= 26;
            }

            return result;
        }

        private static string FormatValidationFailure(ValidationResult validation)
        {
            string[] errors = validation.Issues
                .Where(issue => issue.Severity == ValidationSeverity.Error)
                .Take(5)
                .Select(issue => issue.Path + ": " + issue.Message)
                .ToArray();
            return "Manual layout produced an invalid report setup. " + string.Join(" ", errors);
        }

        private static ReportBlockSpec CloneBlock(ReportBlockSpec source)
        {
            return new ReportBlockSpec
            {
                Id = source.Id,
                OwnershipId = source.OwnershipId,
                Title = source.Title,
                WorksheetName = source.WorksheetName,
                AnchorCell = source.AnchorCell,
                OutputMode = source.OutputMode,
                OwnedExtent = new OwnedRangeExtentSpec
                {
                    RowCount = source.OwnedExtent.RowCount,
                    ColumnCount = source.OwnedExtent.ColumnCount
                },
                Layout = CloneLayout(source.Layout),
                PeriodSlices = source.PeriodSlices.Select(ClonePeriodSlice).ToList(),
                Headers = source.Headers.Select(CloneHeader).ToList(),
                Spacers = source.Spacers.Select(CloneSpacer).ToList(),
                HeaderStyleId = source.HeaderStyleId,
                BodyStyleId = source.BodyStyleId,
                SubtotalStyleId = source.SubtotalStyleId,
                GrandTotalStyleId = source.GrandTotalStyleId
            };
        }

        private static ReportLayoutSpec CloneLayout(ReportLayoutSpec source)
        {
            if (source.Rows == null || source.Columns == null ||
                source.Values == null || source.Filters == null)
            {
                throw new InvalidOperationException(
                    "The configured report block has an incomplete layout collection.");
            }

            return new ReportLayoutSpec
            {
                Rows = source.Rows.Select(CloneFieldPlacement).ToList(),
                Columns = source.Columns.Select(CloneFieldPlacement).ToList(),
                Values = source.Values.Select(CloneValuePlacement).ToList(),
                Filters = source.Filters.Select(CloneFilterPlacement).ToList(),
                DenseLayout = CloneDenseLayout(source.DenseLayout),
                GrandTotals = CloneGrandTotals(source.GrandTotals)
            };
        }

        private static FieldPlacementSpec CloneFieldPlacement(FieldPlacementSpec source)
        {
            if (source == null)
            {
                throw new InvalidOperationException(
                    "The configured layout contains a missing row or column placement.");
            }

            if (source.MemberOrder == null || source.GroupBuckets == null)
            {
                throw new InvalidOperationException(
                    "The configured layout contains an incomplete member order or grouping collection.");
            }

            return new FieldPlacementSpec
            {
                Field = source.Field,
                Caption = source.Caption,
                Subtotals = source.Subtotals == null
                    ? throw new InvalidOperationException(
                        "The configured layout contains a placement without subtotal settings.")
                    : new SubtotalSpec
                    {
                        Mode = source.Subtotals.Mode,
                        Placement = source.Subtotals.Placement,
                        Label = source.Subtotals.Label,
                        StyleId = source.Subtotals.StyleId
                    },
                Sort = source.Sort,
                MemberOrder = source.MemberOrder.Select(CloneScalar).ToList(),
                GroupBuckets = source.GroupBuckets.Select(CloneGroupBucket).ToList(),
                TopN = source.TopN == null
                    ? null
                    : new TopNSpec
                    {
                        Count = source.TopN.Count,
                        MeasureId = source.TopN.MeasureId,
                        Direction = source.TopN.Direction,
                        IncludeOthers = source.TopN.IncludeOthers,
                        OthersLabel = source.TopN.OthersLabel
                    }
            };
        }

        private static MemberGroupBucketSpec CloneGroupBucket(MemberGroupBucketSpec source)
        {
            if (source == null)
            {
                throw new InvalidOperationException(
                    "The configured layout contains a missing member group.");
            }

            if (source.Members == null)
            {
                throw new InvalidOperationException(
                    "The configured layout contains a member group without a member collection.");
            }

            return new MemberGroupBucketSpec
            {
                Id = source.Id,
                Label = source.Label,
                Members = source.Members.Select(CloneScalar).ToList(),
                IncludeUnmatched = source.IncludeUnmatched
            };
        }

        private static ValuePlacementSpec CloneValuePlacement(ValuePlacementSpec source)
        {
            if (source == null)
            {
                throw new InvalidOperationException(
                    "The configured layout contains a missing Value placement.");
            }

            if (source.PeriodSliceIds == null)
            {
                throw new InvalidOperationException(
                    "The configured layout contains a Value without a period-slice collection.");
            }

            return new ValuePlacementSpec
            {
                MeasureId = source.MeasureId,
                Caption = source.Caption,
                NumberFormat = source.NumberFormat,
                PeriodSliceIds = source.PeriodSliceIds.ToList(),
                StyleId = source.StyleId
            };
        }

        private static FilterPlacementSpec CloneFilterPlacement(FilterPlacementSpec source)
        {
            if (source == null)
            {
                throw new InvalidOperationException(
                    "The configured layout contains a missing Filter placement.");
            }

            if (source.SelectedValues == null)
            {
                throw new InvalidOperationException(
                    "The configured layout contains a Filter without a selected-values collection.");
            }

            return new FilterPlacementSpec
            {
                Field = source.Field,
                SelectedValues = source.SelectedValues.Select(CloneScalar).ToList(),
                IncludeBlank = source.IncludeBlank
            };
        }

        private static DenseLayoutOptions CloneDenseLayout(DenseLayoutOptions source)
        {
            if (source == null)
            {
                throw new InvalidOperationException(
                    "The configured report block has no dense-layout settings.");
            }

            return new DenseLayoutOptions
            {
                RepeatRowLabels = source.RepeatRowLabels,
                ShowRowGrandTotals = source.ShowRowGrandTotals,
                ShowColumnGrandTotals = source.ShowColumnGrandTotals,
                InsertBlankRows = source.InsertBlankRows,
                RowIndent = source.RowIndent,
                FreezeHeaders = source.FreezeHeaders
            };
        }

        private static GrandTotalsSpec CloneGrandTotals(GrandTotalsSpec source)
        {
            if (source == null)
            {
                throw new InvalidOperationException(
                    "The configured report block has no grand-total settings.");
            }

            return new GrandTotalsSpec
            {
                ShowRows = source.ShowRows,
                ShowColumns = source.ShowColumns,
                RowPlacement = source.RowPlacement,
                ColumnPlacement = source.ColumnPlacement,
                RowLabel = source.RowLabel,
                ColumnLabel = source.ColumnLabel,
                StyleId = source.StyleId
            };
        }

        private static ScalarValue CloneScalar(ScalarValue source)
        {
            if (source == null)
            {
                throw new InvalidOperationException(
                    "The configured layout contains a missing literal value.");
            }

            return new ScalarValue
            {
                Kind = source.Kind,
                Text = source.Text,
                Number = source.Number,
                Boolean = source.Boolean,
                Temporal = source.Temporal
            };
        }

        private static PeriodSliceSpec ClonePeriodSlice(PeriodSliceSpec source)
        {
            if (source == null)
            {
                throw new InvalidOperationException(
                    "The configured report block contains a missing period slice.");
            }

            return new PeriodSliceSpec
            {
                Id = source.Id,
                Label = source.Label,
                Kind = source.Kind,
                SelectedStart = source.SelectedStart,
                SelectedEnd = source.SelectedEnd,
                BasedOnSliceId = source.BasedOnSliceId
            };
        }

        private static ReportHeaderSpec CloneHeader(ReportHeaderSpec source)
        {
            if (source == null)
            {
                throw new InvalidOperationException(
                    "The configured report block contains a missing header.");
            }

            return new ReportHeaderSpec
            {
                Text = source.Text,
                RelativeRow = source.RelativeRow,
                RelativeColumn = source.RelativeColumn,
                ColumnSpan = source.ColumnSpan,
                StyleId = source.StyleId
            };
        }

        private static SpacerSpec CloneSpacer(SpacerSpec source)
        {
            if (source == null)
            {
                throw new InvalidOperationException(
                    "The configured report block contains a missing spacer.");
            }

            return new SpacerSpec
            {
                Axis = source.Axis,
                BeforeLevel = source.BeforeLevel,
                Count = source.Count,
                Size = source.Size
            };
        }

        private sealed class ManualLayoutValues
        {
            public ManualLayoutValues(
                bool repeatRowLabels,
                bool insertBlankRows,
                bool freezeHeaders,
                bool showRowGrandTotals,
                bool showColumnGrandTotals,
                int rowIndent,
                string rowGrandTotalLabel,
                string columnGrandTotalLabel)
            {
                RepeatRowLabels = repeatRowLabels;
                InsertBlankRows = insertBlankRows;
                FreezeHeaders = freezeHeaders;
                ShowRowGrandTotals = showRowGrandTotals;
                ShowColumnGrandTotals = showColumnGrandTotals;
                RowIndent = rowIndent;
                RowGrandTotalLabel = rowGrandTotalLabel;
                ColumnGrandTotalLabel = columnGrandTotalLabel;
            }

            public bool RepeatRowLabels { get; }

            public bool InsertBlankRows { get; }

            public bool FreezeHeaders { get; }

            public bool ShowRowGrandTotals { get; }

            public bool ShowColumnGrandTotals { get; }

            public int RowIndent { get; }

            public string RowGrandTotalLabel { get; }

            public string ColumnGrandTotalLabel { get; }
        }

        private sealed class MeasureLookup
        {
            public MeasureLookup(
                Dictionary<string, MeasureDefinition> byId,
                Dictionary<string, List<MeasureDefinition>> byLabel)
            {
                ById = byId;
                ByLabel = byLabel;
            }

            public Dictionary<string, MeasureDefinition> ById { get; }

            public Dictionary<string, List<MeasureDefinition>> ByLabel { get; }
        }

        private sealed class ParsedCheckKind
        {
            public ParsedCheckKind(
                ReportCheckKind kind,
                string semanticName,
                string displayName)
            {
                Kind = kind;
                SemanticName = semanticName;
                DisplayName = displayName;
            }

            public ReportCheckKind Kind { get; }

            public string SemanticName { get; }

            public string DisplayName { get; }
        }

        private sealed class ParsedCell
        {
            public ParsedCell(int row, int column, string normalizedAddress)
            {
                Row = row;
                Column = column;
                NormalizedAddress = normalizedAddress;
            }

            public int Row { get; }

            public int Column { get; }

            public string NormalizedAddress { get; }
        }

        private sealed class OwnedRectangle
        {
            public OwnedRectangle(
                string worksheetName,
                int startRow,
                int endRow,
                int startColumn,
                int endColumn,
                string label)
            {
                WorksheetName = worksheetName;
                StartRow = startRow;
                EndRow = endRow;
                StartColumn = startColumn;
                EndColumn = endColumn;
                Label = label;
            }

            public string WorksheetName { get; }

            public int StartRow { get; }

            public int EndRow { get; }

            public int StartColumn { get; }

            public int EndColumn { get; }

            public string Label { get; }

            public string DisplayAddress => ColumnLetters(StartColumn) +
                StartRow.ToString(CultureInfo.InvariantCulture) + ":" +
                ColumnLetters(EndColumn) + EndRow.ToString(CultureInfo.InvariantCulture);

            public bool Overlaps(OwnedRectangle other)
            {
                return string.Equals(
                        WorksheetName,
                        other.WorksheetName,
                        StringComparison.OrdinalIgnoreCase) &&
                    StartRow <= other.EndRow && EndRow >= other.StartRow &&
                    StartColumn <= other.EndColumn && EndColumn >= other.StartColumn;
            }
        }
    }
}
