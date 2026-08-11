using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelReportBuilder.Core.Measures;
using ExcelReportBuilder.Core.Periods;
using ExcelReportBuilder.Core.Planning;
using ExcelReportBuilder.Core.Profiling;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Core.Transforms;

namespace ExcelReportBuilder.Core.Validation
{
    public static class ReportSpecValidator
    {
        private const int MaximumTransforms = 100;
        private const int MaximumMeasures = 128;
        private const int MaximumBlocks = 64;
        private const int MaximumStyles = 128;
        private const int MaximumChecks = 128;
        private const int MaximumPeriodSlices = 64;
        private const int MaximumAxisFields = 32;
        private const int MaximumValues = 128;
        private const int MaximumFilters = 32;
        private const int MaximumGroupBuckets = 256;
        private const int MaximumMembers = 1000;
        private const int MaximumHeaders = 128;
        private const int MaximumSpacers = 64;
        private const int MaximumPeriodColumns = 16384;
        private const int MaximumFilterLiterals = 256;
        private const int MaximumExpressionDepth = 32;
        private const int MaximumExpressionNodes = 256;

        private static readonly Regex IdPattern = new Regex(
            @"^[A-Za-z][A-Za-z0-9_-]{0,63}$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly Regex WorkbookObjectPattern = new Regex(
            @"^[A-Za-z_\\][A-Za-z0-9_.\\]*$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly Regex CellPattern = new Regex(
            @"^\$?(?<column>[A-Za-z]{1,3})\$?(?<row>[1-9][0-9]{0,6})$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly Regex ColorPattern = new Regex(
            @"^#[0-9A-Fa-f]{6}$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly Regex Sha256Pattern = new Regex(
            @"^[0-9a-f]{64}$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        public static ValidationResult Validate(ReportSpecV1 specification, SourceProfile? sourceProfile = null)
        {
            var result = new ValidationResult();
            if (specification == null)
            {
                result.AddError("SPEC_REQUIRED", "$", "A report specification is required.");
                return result;
            }

            ValidateIdentity(specification, result);
            ValidateSource(specification.Source, sourceProfile, result);
            ValidatePeriodMapping(specification.PeriodMapping, sourceProfile, result, "$.periodMapping");
            var outputFields = ValidateTransforms(specification, sourceProfile, result);
            ValidateFinalPeriodField(specification.PeriodMapping, outputFields, result);
            var measures = ValidateMeasures(specification.Measures, outputFields, result);
            ValidatePeriodMappingUsage(specification, measures, result);
            ValidateWeightedConstructions(specification.Measures, specification.Transforms, result);
            var styles = ValidateStyles(specification.Styles, result);
            ValidateBlocks(specification.Blocks, specification.OwnershipId, measures, styles, outputFields, result);
            ValidateChecks(specification.Checks, measures, result);

            if (sourceProfile != null)
            {
                var projection = RowProjectionCalculator.Project(sourceProfile.RowCount, specification.PeriodMapping);
                if (projection.Route == SourceLoadRoute.DataModel)
                {
                    result.AddWarning(
                        "DATA_MODEL_REQUIRED",
                        "$.periodMapping",
                        "The normalized result is projected to contain "
                            + projection.ProjectedRows.ToString(CultureInfo.InvariantCulture)
                            + " rows and must be loaded to the Data Model without truncation.");
                }
            }

            if (specification.Measures != null
                && specification.Measures.Any(measure => measure != null && ContainsDistinctCount(measure.Expression)))
            {
                result.AddWarning(
                    "DATA_MODEL_REQUIRED",
                    "$.measures",
                    "Distinct Count requires a Data Model-backed aggregate even when the prepared rows fit on a worksheet.");
            }

            return result;
        }

        private static bool ContainsDistinctCount(MeasureExpression expression)
        {
            switch (expression)
            {
                case AggregateMeasureExpression aggregate:
                    return aggregate.Function == AggregateFunction.DistinctCount;
                case FilteredAggregateMeasureExpression filtered:
                    return filtered.Function == AggregateFunction.DistinctCount;
                case WeightedAggregateMeasureExpression weighted:
                    return ContainsDistinctCount(weighted.Numerator) || ContainsDistinctCount(weighted.Denominator);
                case BinaryMeasureExpression binary:
                    return ContainsDistinctCount(binary.Left) || ContainsDistinctCount(binary.Right);
                case SafeDivideMeasureExpression divide:
                    return ContainsDistinctCount(divide.Numerator) || ContainsDistinctCount(divide.Denominator);
                case RatioMeasureExpression ratio:
                    return ContainsDistinctCount(ratio.Numerator) || ContainsDistinctCount(ratio.Denominator);
                case DifferenceMeasureExpression difference:
                    return ContainsDistinctCount(difference.Current) || ContainsDistinctCount(difference.Baseline);
                case ShareMeasureExpression share:
                    return ContainsDistinctCount(share.Part) || ContainsDistinctCount(share.Whole);
                default:
                    return false;
            }
        }

        private static void ValidateFinalPeriodField(
            PeriodMappingSpec? mapping,
            HashSet<string>? outputFields,
            ValidationResult result)
        {
            if (mapping == null || outputFields == null)
            {
                return;
            }

            var periodField = mapping.Kind == PeriodMappingKind.LongDateColumn
                ? mapping.DateColumn
                : mapping.PeriodColumnName;
            if (!string.IsNullOrWhiteSpace(periodField) && !outputFields.Contains(periodField!))
            {
                result.AddError(
                    "PERIOD_FIELD_REMOVED_BY_TRANSFORM",
                    "$.transforms",
                    "The final prepared source must retain the mapped period field '" + periodField + "'.");
            }
        }

        private static void ValidatePeriodMappingUsage(
            ReportSpecV1 specification,
            Dictionary<string, MeasureDefinition> measures,
            ValidationResult result)
        {
            if (specification.PeriodMapping != null)
            {
                return;
            }

            var usesSlices = (specification.Blocks ?? new List<ReportBlockSpec>())
                .Where(block => block != null)
                .Any(block => (block.PeriodSlices != null && block.PeriodSlices.Count != 0)
                    || block.Layout != null
                        && block.Layout.Values != null
                        && block.Layout.Values.Any(value => value != null
                            && value.PeriodSliceIds != null
                            && value.PeriodSliceIds.Count != 0));
            if (!usesSlices)
            {
                foreach (var measure in measures.Values)
                {
                    var expressionSlices = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    CollectExpressionSliceIds(
                        measure.Expression,
                        measures,
                        new HashSet<string>(StringComparer.OrdinalIgnoreCase),
                        expressionSlices);
                    if (expressionSlices.Count != 0)
                    {
                        usesSlices = true;
                        break;
                    }
                }
            }

            if (usesSlices)
            {
                result.AddError(
                    "PERIOD_MAPPING_REQUIRED",
                    "$.periodMapping",
                    "Period slices require an explicit source period mapping.");
            }
        }

        private static void ValidateIdentity(ReportSpecV1 specification, ValidationResult result)
        {
            if (!string.Equals(
                specification.SchemaVersion,
                ReportSpecV1.CurrentSchemaVersion,
                StringComparison.Ordinal))
            {
                result.AddError(
                    "SCHEMA_VERSION_UNSUPPORTED",
                    "$.schemaVersion",
                    "schemaVersion must be '" + ReportSpecV1.CurrentSchemaVersion + "'.");
            }

            ValidateId(specification.Id, "$.id", "SPEC_ID_INVALID", result);
            ValidateId(specification.OwnershipId, "$.ownershipId", "OWNERSHIP_ID_INVALID", result);
            ValidateRequiredText(specification.Name, "$.name", 120, result);
        }

        private static void ValidateSource(
            WorkbookSourceSpec source,
            SourceProfile? sourceProfile,
            ValidationResult result)
        {
            if (source == null)
            {
                result.AddError("SOURCE_REQUIRED", "$.source", "A workbook source is required.");
                return;
            }

            ValidateEnum(source.Kind, "$.source.kind", "SOURCE_KIND_INVALID", result);

            if (string.IsNullOrWhiteSpace(source.WorkbookObjectName)
                || source.WorkbookObjectName.Length > 255
                || !WorkbookObjectPattern.IsMatch(source.WorkbookObjectName))
            {
                result.AddError(
                    "SOURCE_NAME_INVALID",
                    "$.source.workbookObjectName",
                    "The source must be a valid Excel table or named-range identifier exposed by Excel.CurrentWorkbook.");
            }

            if (source.HeaderRowCount != 1)
            {
                result.AddError(
                    "ONE_HEADER_ROW_REQUIRED",
                    "$.source.headerRowCount",
                    "Version 1 supports exactly one header row.");
            }

            ValidateSourceFingerprint(source.Fingerprint, sourceProfile, result);

            if (sourceProfile == null)
            {
                return;
            }

            foreach (var issue in sourceProfile.Issues)
            {
                result.AddError(
                    "SOURCE_" + issue.Code.ToString().ToUpperInvariant(),
                    issue.ColumnIndex.HasValue
                        ? "$.source.columns[" + issue.ColumnIndex.Value.ToString(CultureInfo.InvariantCulture) + "]"
                        : "$.source",
                    issue.Message);
            }
        }

        private static void ValidateSourceFingerprint(
            SourceFingerprintSpec fingerprint,
            SourceProfile? sourceProfile,
            ValidationResult result)
        {
            const string path = "$.source.fingerprint";
            if (fingerprint == null)
            {
                result.AddError("SOURCE_FINGERPRINT_REQUIRED", path, "A source fingerprint is required for saved report compatibility.");
                return;
            }

            if (!string.Equals(
                fingerprint.Algorithm,
                SourceFingerprintSpec.CurrentAlgorithm,
                StringComparison.Ordinal))
            {
                result.AddError(
                    "SOURCE_FINGERPRINT_ALGORITHM_UNSUPPORTED",
                    path + ".algorithm",
                    "The source fingerprint algorithm must be '" + SourceFingerprintSpec.CurrentAlgorithm + "'.");
            }

            if (string.IsNullOrWhiteSpace(fingerprint.HeaderHash)
                || !Sha256Pattern.IsMatch(fingerprint.HeaderHash))
            {
                result.AddError(
                    "SOURCE_HEADER_HASH_INVALID",
                    path + ".headerHash",
                    "The source header hash must be a lowercase SHA-256 value.");
            }

            if (fingerprint.ColumnCount < 1 || fingerprint.ColumnCount > 16384)
            {
                result.AddError(
                    "SOURCE_FINGERPRINT_COLUMN_COUNT_INVALID",
                    path + ".columnCount",
                    "The fingerprint column count must be within worksheet bounds.");
            }

            if (fingerprint.SampleRowCount.HasValue != (fingerprint.SampleHash != null))
            {
                result.AddError(
                    "SOURCE_SAMPLE_FINGERPRINT_INCOMPLETE",
                    path,
                    "Sample hash and sample row count must either both be present or both be absent.");
            }

            if (fingerprint.SampleRowCount.HasValue
                && (fingerprint.SampleRowCount.Value < 1 || fingerprint.SampleRowCount.Value > 64))
            {
                result.AddError(
                    "SOURCE_SAMPLE_ROW_COUNT_INVALID",
                    path + ".sampleRowCount",
                    "A source fingerprint may summarize 1-64 sampled rows.");
            }

            if (fingerprint.SampleHash != null
                && !Sha256Pattern.IsMatch(fingerprint.SampleHash))
            {
                result.AddError(
                    "SOURCE_SAMPLE_HASH_INVALID",
                    path + ".sampleHash",
                    "The source sample hash must be a lowercase SHA-256 value.");
            }

            if (sourceProfile == null)
            {
                return;
            }

            if (fingerprint.ColumnCount != sourceProfile.ColumnCount)
            {
                result.AddError(
                    "SOURCE_FINGERPRINT_SHAPE_MISMATCH",
                    path + ".columnCount",
                    "The selected source column count no longer matches the saved report setup.");
            }

            var actualHeaderHash = SourceFingerprint.ComputeHeaderHash(
                sourceProfile.Columns.OrderBy(column => column.Index).Select(column => column.Name));
            if (!string.Equals(fingerprint.HeaderHash, actualHeaderHash, StringComparison.Ordinal))
            {
                result.AddError(
                    "SOURCE_FINGERPRINT_HEADER_MISMATCH",
                    path + ".headerHash",
                    "The selected source headers no longer match the saved report setup.");
            }
        }

        private static void ValidatePeriodMapping(
            PeriodMappingSpec? mapping,
            SourceProfile? sourceProfile,
            ValidationResult result,
            string path)
        {
            if (mapping == null)
            {
                return;
            }

            ValidateId(mapping.Id, path + ".id", "PERIOD_MAPPING_ID_INVALID", result);
            ValidateEnum(mapping.Kind, path + ".kind", "PERIOD_MAPPING_KIND_INVALID", result);
            if (mapping.Grain.HasValue)
            {
                ValidateEnum(mapping.Grain.Value, path + ".grain", "PERIOD_GRAIN_INVALID", result);
            }

            ValidateColumnName(mapping.PeriodColumnName, path + ".periodColumnName", result);
            ValidateColumnName(mapping.ValueColumnName, path + ".valueColumnName", result);
            ValidateColumnName(mapping.MetricColumnName, path + ".metricColumnName", result);
            ValidateDistinct(
                new[] { mapping.PeriodColumnName, mapping.ValueColumnName, mapping.MetricColumnName },
                path,
                "PERIOD_OUTPUT_NAMES_DUPLICATE",
                result);

            if (mapping.ReportingYear.HasValue
                && (mapping.ReportingYear.Value < 1900 || mapping.ReportingYear.Value > 9999))
            {
                result.AddError(
                    "REPORTING_YEAR_INVALID",
                    path + ".reportingYear",
                    "The reporting year must be between 1900 and 9999.");
            }

            if (mapping.Kind == PeriodMappingKind.LongDateColumn)
            {
                ValidateColumnName(mapping.DateColumn, path + ".dateColumn", result);
                if (mapping.Columns == null)
                {
                    result.AddError("PERIOD_COLUMNS_REQUIRED", path + ".columns", "The period-column collection is required.");
                }
                else if (mapping.Columns.Count != 0)
                {
                    result.AddError(
                        "LONG_PERIOD_COLUMNS_NOT_ALLOWED",
                        path + ".columns",
                        "A long date mapping cannot also contain wide period columns.");
                }

                ValidateProfileColumn(mapping.DateColumn, sourceProfile, path + ".dateColumn", result);
                SourceColumnProfile? periodColumn = sourceProfile == null ||
                    string.IsNullOrWhiteSpace(mapping.DateColumn)
                    ? null
                    : sourceProfile.FindColumn(mapping.DateColumn!);
                if (periodColumn != null)
                {
                    if (periodColumn.PeriodLikeWithoutYearCount > 0 && !mapping.ReportingYear.HasValue)
                    {
                        result.AddError(
                            "REPORTING_YEAR_REQUIRED",
                            path + ".reportingYear",
                            "The selected period column contains month or quarter values without a year. Choose a reporting year before building.");
                    }

                    int observedGrainCount = (periodColumn.DayGrainCount > 0 ? 1 : 0) +
                        (periodColumn.MonthGrainCount > 0 ? 1 : 0) +
                        (periodColumn.QuarterGrainCount > 0 ? 1 : 0);
                    if (observedGrainCount > 1)
                    {
                        result.AddError(
                            "MIXED_PERIOD_GRAINS",
                            path + ".grain",
                            "The selected period column mixes day, month, or quarter values.");
                    }
                }

                return;
            }

            if (mapping.DateColumn != null)
            {
                result.AddError(
                    "WIDE_DATE_COLUMN_NOT_ALLOWED",
                    path + ".dateColumn",
                    "A wide period mapping must use explicit header mappings, not a date column.");
            }

            var wideGrain = mapping.Grain ?? PeriodGrain.Month;
            if (wideGrain == PeriodGrain.Day)
            {
                result.AddError(
                    "WIDE_PERIOD_GRAIN_INVALID",
                    path + ".grain",
                    "Wide period headers must use month or quarter grain.");
            }

            ValidateColumnList(mapping.KeyColumns, path + ".keyColumns", false, result);
            var normalizationOutputs = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                mapping.PeriodColumnName,
                mapping.ValueColumnName
            };
            if (mapping.Kind == PeriodMappingKind.MetricMonthHeaders)
            {
                normalizationOutputs.Add(mapping.MetricColumnName);
            }

            foreach (var keyColumn in mapping.KeyColumns ?? new List<string>())
            {
                if (normalizationOutputs.Contains(keyColumn))
                {
                    result.AddError(
                        "PERIOD_OUTPUT_KEY_COLLISION",
                        path + ".keyColumns",
                        "A period-normalization output name cannot also be a key column.");
                }
            }

            if (mapping.Columns == null || mapping.Columns.Count == 0)
            {
                result.AddError("PERIOD_COLUMNS_REQUIRED", path + ".columns", "At least one period column is required.");
                return;
            }

            if (mapping.Columns.Count > MaximumPeriodColumns)
            {
                result.AddError(
                    "TOO_MANY_PERIOD_COLUMNS",
                    path + ".columns",
                    "A period mapping may contain at most 16,384 source columns.");
            }

            var sourceColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var periodKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var periodsByMetric = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
            for (var index = 0; index < mapping.Columns.Count; index++)
            {
                var column = mapping.Columns[index];
                var columnPath = path + ".columns[" + index.ToString(CultureInfo.InvariantCulture) + "]";
                if (column == null)
                {
                    result.AddError("PERIOD_COLUMN_REQUIRED", columnPath, "A period-column mapping cannot be null.");
                    continue;
                }

                ValidateColumnName(column.SourceColumn, columnPath + ".sourceColumn", result);
                ValidateProfileColumn(column.SourceColumn, sourceProfile, columnPath + ".sourceColumn", result);
                if (!sourceColumns.Add(column.SourceColumn))
                {
                    result.AddError(
                        "PERIOD_SOURCE_COLUMN_DUPLICATE",
                        columnPath + ".sourceColumn",
                        "A source column can appear only once in the period mapping.");
                }

                if (column.Month < 1 || column.Month > 12)
                {
                    result.AddError("PERIOD_MONTH_INVALID", columnPath + ".month", "Month must be between 1 and 12.");
                }
                else if (wideGrain == PeriodGrain.Quarter
                    && column.Month != 1 && column.Month != 4
                    && column.Month != 7 && column.Month != 10)
                {
                    result.AddError(
                        "QUARTER_START_MONTH_INVALID",
                        columnPath + ".month",
                        "Quarter mappings must use January, April, July, or October as the canonical start month.");
                }

                var effectiveYear = column.Year ?? mapping.ReportingYear;
                if (!effectiveYear.HasValue)
                {
                    result.AddError(
                        "REPORTING_YEAR_REQUIRED",
                        columnPath + ".year",
                        "A month header without a year requires an explicit reportingYear.");
                }
                else if (effectiveYear.Value < 1900 || effectiveYear.Value > 9999)
                {
                    result.AddError("PERIOD_YEAR_INVALID", columnPath + ".year", "Year must be between 1900 and 9999.");
                }

                var periodKey = (effectiveYear.HasValue
                        ? effectiveYear.Value.ToString("0000", CultureInfo.InvariantCulture)
                        : "????")
                    + "-"
                    + column.Month.ToString("00", CultureInfo.InvariantCulture);
                if (mapping.Kind == PeriodMappingKind.MonthHeaders)
                {
                    if (column.Metric != null)
                    {
                        result.AddError(
                            "METRIC_NOT_ALLOWED",
                            columnPath + ".metric",
                            "Month-only mappings cannot specify a metric.");
                    }

                    if (!periodKeys.Add(periodKey))
                    {
                        result.AddError(
                            "PERIOD_DUPLICATE",
                            columnPath,
                            "Only one source column may map to each period.");
                    }
                }
                else
                {
                    ValidateRequiredText(column.Metric, columnPath + ".metric", 120, result);
                    var metric = column.Metric ?? string.Empty;
                    HashSet<string> metricPeriods;
                    if (!periodsByMetric.TryGetValue(metric, out metricPeriods))
                    {
                        metricPeriods = new HashSet<string>(StringComparer.Ordinal);
                        periodsByMetric.Add(metric, metricPeriods);
                    }

                    if (!metricPeriods.Add(periodKey))
                    {
                        result.AddError(
                            "METRIC_PERIOD_DUPLICATE",
                            columnPath,
                            "Only one source column may map to each metric and period.");
                    }

                    periodKeys.Add(periodKey);
                }
            }

            foreach (var keyColumn in mapping.KeyColumns ?? new List<string>())
            {
                ValidateProfileColumn(keyColumn, sourceProfile, path + ".keyColumns", result);
                if (sourceColumns.Contains(keyColumn))
                {
                    result.AddError(
                        "KEY_PERIOD_COLUMN_OVERLAP",
                        path + ".keyColumns",
                        "A key column cannot also be a period column.");
                }
            }

            if (mapping.Kind == PeriodMappingKind.MetricMonthHeaders)
            {
                foreach (var pair in periodsByMetric)
                {
                    if (!pair.Value.SetEquals(periodKeys))
                    {
                        result.AddError(
                            "METRIC_PERIOD_MATRIX_INCOMPLETE",
                            path + ".columns",
                            "Metric '" + pair.Key + "' does not contain every mapped period.");
                    }
                }
            }
        }

        private static HashSet<string>? ValidateTransforms(
            ReportSpecV1 specification,
            SourceProfile? sourceProfile,
            ValidationResult result)
        {
            if (specification.Transforms == null)
            {
                result.AddError("TRANSFORMS_REQUIRED", "$.transforms", "The transforms collection is required.");
                return null;
            }

            if (specification.Transforms.Count > MaximumTransforms)
            {
                result.AddError(
                    "TOO_MANY_TRANSFORMS",
                    "$.transforms",
                    "A report may contain at most " + MaximumTransforms.ToString(CultureInfo.InvariantCulture) + " transforms.");
            }

            HashSet<string>? fields = sourceProfile == null
                ? null
                : new HashSet<string>(sourceProfile.Columns.Select(column => column.Name), StringComparer.OrdinalIgnoreCase);
            var transformIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var normalizeCount = 0;
            for (var index = 0; index < specification.Transforms.Count; index++)
            {
                var transform = specification.Transforms[index];
                var path = "$.transforms[" + index.ToString(CultureInfo.InvariantCulture) + "]";
                if (transform == null)
                {
                    result.AddError("TRANSFORM_REQUIRED", path, "A transform cannot be null.");
                    continue;
                }

                ValidateId(transform.Id, path + ".id", "TRANSFORM_ID_INVALID", result);
                if (!transformIds.Add(transform.Id ?? string.Empty))
                {
                    result.AddError("TRANSFORM_ID_DUPLICATE", path + ".id", "Transform IDs must be unique.");
                }

                switch (transform)
                {
                    case SelectColumnsTransform select:
                        ValidateColumnList(select.Columns, path + ".columns", true, result);
                        ValidateExistingColumns(select.Columns, fields, path + ".columns", result);
                        fields = fields == null || select.Columns == null
                            ? null
                            : new HashSet<string>(select.Columns, StringComparer.OrdinalIgnoreCase);
                        break;
                    case KeepColumnsTransform keep:
                        ValidateColumnList(keep.Columns, path + ".columns", true, result);
                        ValidateExistingColumns(keep.Columns, fields, path + ".columns", result);
                        fields = fields == null || keep.Columns == null
                            ? null
                            : new HashSet<string>(keep.Columns, StringComparer.OrdinalIgnoreCase);
                        break;
                    case RemoveColumnsTransform remove:
                        ValidateColumnList(remove.Columns, path + ".columns", true, result);
                        ValidateExistingColumns(remove.Columns, fields, path + ".columns", result);
                        if (fields != null && remove.Columns != null)
                        {
                            foreach (var column in remove.Columns)
                            {
                                fields.Remove(column);
                            }
                        }

                        break;
                    case ReorderColumnsTransform reorder:
                        ValidateColumnList(reorder.Columns, path + ".columns", true, result);
                        ValidateExistingColumns(reorder.Columns, fields, path + ".columns", result);
                        if (fields != null && reorder.Columns != null && !fields.SetEquals(reorder.Columns))
                        {
                            result.AddError(
                                "REORDER_MUST_LIST_ALL_COLUMNS",
                                path + ".columns",
                                "Reordering must list every current column exactly once.");
                        }

                        break;
                    case RenameColumnTransform rename:
                        ValidateColumnName(rename.From, path + ".from", result);
                        ValidateColumnName(rename.To, path + ".to", result);
                        ValidateExistingColumn(rename.From, fields, path + ".from", result);
                        if (fields != null && fields.Contains(rename.To)
                            && !string.Equals(rename.From, rename.To, StringComparison.OrdinalIgnoreCase))
                        {
                            result.AddError("RENAME_TARGET_EXISTS", path + ".to", "The renamed column already exists.");
                        }

                        if (fields != null)
                        {
                            fields.Remove(rename.From);
                            fields.Add(rename.To);
                        }

                        break;
                    case ChangeColumnTypeTransform changeType:
                        ValidateColumnOperation(changeType.Column, fields, path + ".column", result);
                        ValidateEnum(changeType.DataType, path + ".dataType", "COLUMN_DATA_TYPE_INVALID", result);
                        break;
                    case TrimTextTransform trim:
                        ValidateColumnOperationList(trim.Columns, fields, path + ".columns", result);
                        break;
                    case ReplaceValueTransform replace:
                        ValidateColumnOperation(replace.Column, fields, path + ".column", result);
                        ValidateScalar(replace.Find, path + ".find", result);
                        ValidateScalar(replace.ReplaceWith, path + ".replaceWith", result);
                        break;
                    case NormalizeBlanksTransform blanks:
                        ValidateColumnOperationList(blanks.Columns, fields, path + ".columns", result);
                        ValidateScalar(blanks.Replacement, path + ".replacement", result);
                        break;
                    case NormalizeErrorsTransform errors:
                        ValidateColumnOperationList(errors.Columns, fields, path + ".columns", result);
                        ValidateScalar(errors.Replacement, path + ".replacement", result);
                        break;
                    case FillDownTransform fillDown:
                        ValidateColumnOperationList(fillDown.Columns, fields, path + ".columns", result);
                        break;
                    case MapValuesTransform map:
                        ValidateMap(map, fields, path, result);
                        break;
                    case FilterRowsTransform filter:
                        ValidateFilter(filter, fields, path, result);
                        break;
                    case ExcludeTotalRowsTransform exclude:
                        ValidateTotalRowExclusion(exclude, fields, path, result);
                        break;
                    case DerivePeriodPartsTransform derive:
                        ValidateDerivePeriodParts(derive, fields, path, result);
                        break;
                    case AddArithmeticColumnTransform arithmetic:
                        ValidateArithmetic(arithmetic, fields, path, result);
                        break;
                    case NormalizePeriodsTransform normalize:
                        normalizeCount++;
                        fields = ValidateNormalizeTransform(normalize, specification.PeriodMapping, fields, path, result);
                        break;
                    default:
                        result.AddError("TRANSFORM_KIND_UNSUPPORTED", path + ".kind", "The transform kind is not supported.");
                        break;
                }
            }

            if (normalizeCount > 1)
            {
                result.AddError(
                    "MULTIPLE_PERIOD_NORMALIZATIONS",
                    "$.transforms",
                    "Only one period-normalization transform is allowed.");
            }

            if (specification.PeriodMapping != null
                && specification.PeriodMapping.Kind != PeriodMappingKind.LongDateColumn
                && normalizeCount == 0)
            {
                result.AddError(
                    "PERIOD_NORMALIZATION_REQUIRED",
                    "$.transforms",
                    "A wide period mapping requires a normalizePeriods transform.");
            }

            return fields;
        }

        private static HashSet<string>? ValidateNormalizeTransform(
            NormalizePeriodsTransform transform,
            PeriodMappingSpec? rootMapping,
            HashSet<string>? fields,
            string path,
            ValidationResult result)
        {
            if (string.IsNullOrWhiteSpace(transform.PeriodMappingId))
            {
                result.AddError(
                    "PERIOD_MAPPING_REFERENCE_REQUIRED",
                    path + ".periodMappingId",
                    "The normalize transform must reference the report's explicit periodMapping.");
                return fields;
            }

            if (rootMapping == null
                || !string.Equals(rootMapping.Id, transform.PeriodMappingId, StringComparison.OrdinalIgnoreCase))
            {
                result.AddError(
                    "PERIOD_MAPPING_REFERENCE_UNKNOWN",
                    path + ".periodMappingId",
                    "The normalize transform must reference the report's explicit periodMapping.");
                return fields;
            }

            var mapping = rootMapping;

            if (mapping.Kind == PeriodMappingKind.LongDateColumn)
            {
                if (!mapping.Grain.HasValue)
                {
                    result.AddError(
                        "LONG_PERIOD_GRAIN_REQUIRED",
                        path,
                        "A normalized long period column requires an explicit day, month, or quarter grain.");
                }

                ValidateExistingColumn(mapping.DateColumn, fields, path + ".dateColumn", result);
                return fields;
            }

            ValidateExistingColumns(mapping.KeyColumns, fields, path + ".keyColumns", result);
            ValidateExistingColumns(mapping.Columns.Select(column => column.SourceColumn), fields, path + ".periodColumns", result);
            if (fields == null)
            {
                return null;
            }

            var output = new HashSet<string>(mapping.KeyColumns, StringComparer.OrdinalIgnoreCase)
            {
                mapping.PeriodColumnName
            };
            if (mapping.Kind == PeriodMappingKind.MonthHeaders)
            {
                output.Add(mapping.ValueColumnName);
            }
            else
            {
                output.Add(mapping.MetricColumnName);
                output.Add(mapping.ValueColumnName);
            }

            return output;
        }

        private static void ValidateMap(
            MapValuesTransform map,
            HashSet<string>? fields,
            string path,
            ValidationResult result)
        {
            ValidateColumnOperation(map.Column, fields, path + ".column", result);
            if (map.Entries == null || map.Entries.Count == 0)
            {
                result.AddError("MAP_ENTRIES_REQUIRED", path + ".entries", "At least one value mapping is required.");
                return;
            }

            if (map.Entries.Count > 256)
            {
                result.AddError("MAP_TOO_LARGE", path + ".entries", "A value map may contain at most 256 entries.");
            }

            var mappedInputs = new HashSet<string>(StringComparer.Ordinal);
            for (var index = 0; index < map.Entries.Count; index++)
            {
                var entryPath = path + ".entries[" + index.ToString(CultureInfo.InvariantCulture) + "]";
                var entry = map.Entries[index];
                if (entry == null)
                {
                    result.AddError("MAP_ENTRY_REQUIRED", entryPath, "A value-map entry cannot be null.");
                    continue;
                }

                ValidateScalar(entry.From, entryPath + ".from", result);
                ValidateScalar(entry.To, entryPath + ".to", result);
                if (!mappedInputs.Add(ScalarKey(entry.From)))
                {
                    result.AddError(
                        "MAP_INPUT_DUPLICATE",
                        entryPath + ".from",
                        "Each input literal can appear only once in a value map.");
                }
            }
        }

        private static void ValidateFilter(
            FilterRowsTransform filter,
            HashSet<string>? fields,
            string path,
            ValidationResult result)
        {
            ValidateColumnOperation(filter.Column, fields, path + ".column", result);
            ValidateEnum(filter.Operator, path + ".operator", "FILTER_OPERATOR_INVALID", result);
            var requiresValue = filter.Operator != RowFilterOperator.IsBlank
                && filter.Operator != RowFilterOperator.IsNotBlank;
            if (requiresValue && filter.Value == null)
            {
                result.AddError("FILTER_VALUE_REQUIRED", path + ".value", "This filter operator requires a literal value.");
            }
            else if (!requiresValue && filter.Value != null)
            {
                result.AddError("FILTER_VALUE_NOT_ALLOWED", path + ".value", "Blank filters do not accept a value.");
            }
            else if (filter.Value != null)
            {
                ValidateScalar(filter.Value, path + ".value", result);
                if ((filter.Operator == RowFilterOperator.Contains
                        || filter.Operator == RowFilterOperator.StartsWith
                        || filter.Operator == RowFilterOperator.EndsWith)
                    && filter.Value.Kind != ScalarValueKind.Text)
                {
                    result.AddError("TEXT_FILTER_VALUE_REQUIRED", path + ".value", "Text matching requires a text literal.");
                }
            }
        }

        private static void ValidateTotalRowExclusion(
            ExcludeTotalRowsTransform exclude,
            HashSet<string>? fields,
            string path,
            ValidationResult result)
        {
            if (exclude.Evidence == null || exclude.Evidence.Count == 0)
            {
                result.AddError(
                    "TOTAL_ROW_EVIDENCE_REQUIRED",
                    path + ".evidence",
                    "Total rows may be excluded only with explicit observed evidence.");
                return;
            }

            if (exclude.Evidence.Count > 32)
            {
                result.AddError("TOO_MUCH_TOTAL_ROW_EVIDENCE", path + ".evidence", "At most 32 evidence conditions are allowed.");
            }

            for (var index = 0; index < exclude.Evidence.Count; index++)
            {
                var evidence = exclude.Evidence[index];
                var evidencePath = path + ".evidence[" + index.ToString(CultureInfo.InvariantCulture) + "]";
                if (evidence == null)
                {
                    result.AddError("TOTAL_ROW_EVIDENCE_REQUIRED", evidencePath, "A total-row evidence item cannot be null.");
                    continue;
                }

                ValidateColumnOperation(evidence.Column, fields, evidencePath + ".column", result);
                ValidateEnum(evidence.MatchKind, evidencePath + ".matchKind", "TOTAL_ROW_MATCH_KIND_INVALID", result);
                ValidateEnum(evidence.Source, evidencePath + ".source", "EVIDENCE_SOURCE_INVALID", result);
                if (evidence.ObservedMatchCount <= 0)
                {
                    result.AddError(
                        "TOTAL_ROW_MATCH_COUNT_REQUIRED",
                        evidencePath + ".observedMatchCount",
                        "Evidence must state how many matching rows were observed.");
                }

                if (evidence.MatchKind != TotalRowMatchKind.IsBlank
                    && (evidence.Values == null || evidence.Values.Count == 0))
                {
                    result.AddError(
                        "TOTAL_ROW_EVIDENCE_VALUES_REQUIRED",
                        evidencePath + ".values",
                        "This evidence match requires at least one literal value.");
                }

                if (evidence.Values != null && evidence.Values.Count > MaximumFilterLiterals)
                {
                    result.AddError(
                        "TOO_MANY_TOTAL_ROW_EVIDENCE_VALUES",
                        evidencePath + ".values",
                        "A total-row evidence condition may contain at most 256 literal values.");
                }

                if (evidence.MatchKind == TotalRowMatchKind.IsBlank
                    && evidence.Values != null
                    && evidence.Values.Count != 0)
                {
                    result.AddError(
                        "TOTAL_ROW_EVIDENCE_VALUES_NOT_ALLOWED",
                        evidencePath + ".values",
                        "Blank total-row evidence cannot contain literal values.");
                }

                foreach (var value in evidence.Values ?? new List<ScalarValue>())
                {
                    ValidateScalar(value, evidencePath + ".values", result);
                    if ((evidence.MatchKind == TotalRowMatchKind.StartsWith
                            || evidence.MatchKind == TotalRowMatchKind.Contains)
                        && value.Kind != ScalarValueKind.Text)
                    {
                        result.AddError(
                            "TOTAL_ROW_TEXT_EVIDENCE_REQUIRED",
                            evidencePath + ".values",
                            "Text evidence matching requires text literals.");
                    }
                }
            }
        }

        private static void ValidateDerivePeriodParts(
            DerivePeriodPartsTransform derive,
            HashSet<string>? fields,
            string path,
            ValidationResult result)
        {
            ValidateColumnOperation(derive.DateColumn, fields, path + ".dateColumn", result);
            if (derive.Columns == null || derive.Columns.Count == 0)
            {
                result.AddError("DERIVED_PERIOD_COLUMNS_REQUIRED", path + ".columns", "At least one period part is required.");
                return;
            }

            if (derive.Columns.Count > 6)
            {
                result.AddError("TOO_MANY_DERIVED_PERIOD_COLUMNS", path + ".columns", "At most six distinct period parts can be derived.");
            }

            var parts = new HashSet<DerivedPeriodPart>();
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var index = 0; index < derive.Columns.Count; index++)
            {
                var column = derive.Columns[index];
                var columnPath = path + ".columns[" + index.ToString(CultureInfo.InvariantCulture) + "]";
                if (column == null)
                {
                    result.AddError("DERIVED_PERIOD_COLUMN_REQUIRED", columnPath, "A derived period column cannot be null.");
                    continue;
                }

                ValidateEnum(column.Part, columnPath + ".part", "DERIVED_PERIOD_PART_INVALID", result);
                ValidateColumnName(column.OutputColumn, columnPath + ".outputColumn", result);
                if (!parts.Add(column.Part))
                {
                    result.AddError("DERIVED_PERIOD_PART_DUPLICATE", columnPath + ".part", "Each period part may be derived once.");
                }

                if (!names.Add(column.OutputColumn) || (fields != null && fields.Contains(column.OutputColumn)))
                {
                    result.AddError("DERIVED_COLUMN_EXISTS", columnPath + ".outputColumn", "The derived output column already exists.");
                }
            }

            if (fields != null)
            {
                foreach (var name in names)
                {
                    fields.Add(name);
                }
            }
        }

        private static void ValidateArithmetic(
            AddArithmeticColumnTransform arithmetic,
            HashSet<string>? fields,
            string path,
            ValidationResult result)
        {
            ValidateColumnName(arithmetic.OutputColumn, path + ".outputColumn", result);
            ValidateEnum(arithmetic.Operator, path + ".operator", "ARITHMETIC_OPERATOR_INVALID", result);
            if (fields != null && fields.Contains(arithmetic.OutputColumn))
            {
                result.AddError("ARITHMETIC_COLUMN_EXISTS", path + ".outputColumn", "The arithmetic output column already exists.");
            }

            if (arithmetic.ResultType != ColumnDataType.WholeNumber
                && arithmetic.ResultType != ColumnDataType.DecimalNumber)
            {
                result.AddError(
                    "ARITHMETIC_RESULT_TYPE_INVALID",
                    path + ".resultType",
                    "Typed arithmetic can produce only whole-number or decimal-number columns.");
            }

            if (arithmetic.Operator == ArithmeticOperator.Divide
                && arithmetic.ResultType != ColumnDataType.DecimalNumber)
            {
                result.AddError(
                    "ARITHMETIC_DIVIDE_RESULT_TYPE_INVALID",
                    path + ".resultType",
                    "Division must produce a decimal-number column.");
            }

            if (arithmetic.Operator == ArithmeticOperator.Divide
                && !arithmetic.ReturnNullOnZeroDenominator)
            {
                result.AddError(
                    "ARITHMETIC_DIVIDE_NULL_ON_ZERO_REQUIRED",
                    path + ".returnNullOnZeroDenominator",
                    "Division must return blank when the denominator is zero.");
            }

            ValidateArithmeticOperand(arithmetic.Left, fields, path + ".left", result);
            ValidateArithmeticOperand(arithmetic.Right, fields, path + ".right", result);
            if (arithmetic.Operator == ArithmeticOperator.Divide
                && arithmetic.Right != null
                && arithmetic.Right.Kind == ArithmeticOperandKind.Number
                && arithmetic.Right.Number == 0m)
            {
                result.AddError("ARITHMETIC_LITERAL_DIVIDE_BY_ZERO", path + ".right", "A literal denominator cannot be zero.");
            }

            fields?.Add(arithmetic.OutputColumn);
        }

        private static void ValidateArithmeticOperand(
            ArithmeticOperand operand,
            HashSet<string>? fields,
            string path,
            ValidationResult result)
        {
            if (operand == null)
            {
                result.AddError("ARITHMETIC_OPERAND_REQUIRED", path, "An arithmetic operand is required.");
                return;
            }

            ValidateEnum(operand.Kind, path + ".kind", "ARITHMETIC_OPERAND_KIND_INVALID", result);

            if (operand.Kind == ArithmeticOperandKind.Column)
            {
                ValidateColumnOperation(operand.Column, fields, path + ".column", result);
                if (operand.Number.HasValue)
                {
                    result.AddError("ARITHMETIC_NUMBER_NOT_ALLOWED", path + ".number", "A column operand cannot contain a number.");
                }
            }
            else
            {
                if (!operand.Number.HasValue)
                {
                    result.AddError("ARITHMETIC_NUMBER_REQUIRED", path + ".number", "A number operand requires a literal number.");
                }

                if (!string.IsNullOrWhiteSpace(operand.Column))
                {
                    result.AddError("ARITHMETIC_COLUMN_NOT_ALLOWED", path + ".column", "A number operand cannot reference a column.");
                }
            }
        }

        private static Dictionary<string, MeasureDefinition> ValidateMeasures(
            List<MeasureDefinition> definitions,
            HashSet<string>? fields,
            ValidationResult result)
        {
            var measures = new Dictionary<string, MeasureDefinition>(StringComparer.OrdinalIgnoreCase);
            if (definitions == null)
            {
                result.AddError("MEASURES_REQUIRED", "$.measures", "The measures collection is required.");
                return measures;
            }

            if (definitions.Count > MaximumMeasures)
            {
                result.AddError("TOO_MANY_MEASURES", "$.measures", "A report may contain at most 128 measures.");
            }

            for (var index = 0; index < definitions.Count; index++)
            {
                var definition = definitions[index];
                var path = "$.measures[" + index.ToString(CultureInfo.InvariantCulture) + "]";
                if (definition == null)
                {
                    result.AddError("MEASURE_REQUIRED", path, "A measure cannot be null.");
                    continue;
                }

                ValidateId(definition.Id, path + ".id", "MEASURE_ID_INVALID", result);
                ValidateRequiredText(definition.Label, path + ".label", 120, result);
                ValidateEnum(definition.ValueType, path + ".valueType", "MEASURE_VALUE_TYPE_INVALID", result);
                ValidateNumberFormat(definition.NumberFormat, path + ".numberFormat", result);
                var measureKey = definition.Id ?? string.Empty;
                if (measures.ContainsKey(measureKey))
                {
                    result.AddError("MEASURE_ID_DUPLICATE", path + ".id", "Measure IDs must be unique.");
                }
                else
                {
                    measures.Add(measureKey, definition);
                }
            }

            var state = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var definition in measures.Values)
            {
                ValidateMeasureDefinition(definition, measures, fields, state, result);
            }

            return measures;
        }

        private static MeasureValueType? ValidateMeasureDefinition(
            MeasureDefinition definition,
            Dictionary<string, MeasureDefinition> measures,
            HashSet<string>? fields,
            Dictionary<string, int> state,
            ValidationResult result)
        {
            int currentState;
            state.TryGetValue(definition.Id, out currentState);
            if (currentState == 2)
            {
                return definition.ValueType;
            }

            if (currentState == 1)
            {
                result.AddError(
                    "MEASURE_REFERENCE_CYCLE",
                    "$.measures",
                    "The measure graph contains a reference cycle involving '" + definition.Id + "'.");
                return null;
            }

            state[definition.Id] = 1;
            var nodeCount = 0;
            var expressionStack = new HashSet<MeasureExpression>();
            var inferred = ValidateExpression(
                definition.Expression,
                definition.Id,
                measures,
                fields,
                state,
                expressionStack,
                0,
                ref nodeCount,
                result);
            if (inferred.HasValue && inferred.Value != definition.ValueType)
            {
                result.AddError(
                    "MEASURE_TYPE_MISMATCH",
                    "$.measures[" + definition.Id + "].valueType",
                    "The declared measure type does not match its expression type.");
            }

            state[definition.Id] = 2;
            return inferred;
        }

        private static MeasureValueType? ValidateExpression(
            MeasureExpression expression,
            string ownerId,
            Dictionary<string, MeasureDefinition> measures,
            HashSet<string>? fields,
            Dictionary<string, int> state,
            HashSet<MeasureExpression> expressionStack,
            int depth,
            ref int nodeCount,
            ValidationResult result)
        {
            var path = "$.measures[" + ownerId + "].expression";
            if (expression == null)
            {
                result.AddError("MEASURE_EXPRESSION_REQUIRED", path, "A measure expression is required.");
                return null;
            }

            nodeCount++;
            if (depth > MaximumExpressionDepth || nodeCount > MaximumExpressionNodes)
            {
                result.AddError("MEASURE_EXPRESSION_TOO_COMPLEX", path, "The measure expression exceeds the bounded complexity limit.");
                return null;
            }

            if (!expressionStack.Add(expression))
            {
                result.AddError("MEASURE_EXPRESSION_OBJECT_CYCLE", path, "The expression object graph contains a cycle.");
                return null;
            }

            ValidateEnum(expression.ResultType, path + ".resultType", "MEASURE_RESULT_TYPE_INVALID", result);

            MeasureValueType? inferred;
            switch (expression)
            {
                case AggregateMeasureExpression aggregate:
                    ValidateFieldReference(aggregate.Field, fields, path + ".field", result);
                    ValidateOptionalId(aggregate.PeriodSliceId, path + ".periodSliceId", "PERIOD_SLICE_ID_INVALID", result);
                    inferred = ValidateAggregateType(aggregate.Function, aggregate.ResultType, path, result);
                    break;
                case FilteredAggregateMeasureExpression filtered:
                    ValidateFieldReference(filtered.Field, fields, path + ".field", result);
                    ValidateMeasureFilters(filtered.Filters, fields, path + ".filters", result);
                    ValidateOptionalId(filtered.PeriodSliceId, path + ".periodSliceId", "PERIOD_SLICE_ID_INVALID", result);
                    inferred = ValidateAggregateType(filtered.Function, filtered.ResultType, path, result);
                    break;
                case WeightedAggregateMeasureExpression weighted:
                    ValidateEnum(weighted.OnZero, path + ".onZero", "ZERO_DENOMINATOR_BEHAVIOR_INVALID", result);
                    var weightedNumerator = ValidateExpression(weighted.Numerator, ownerId, measures, fields, state, expressionStack, depth + 1, ref nodeCount, result);
                    var weightedDenominator = ValidateExpression(weighted.Denominator, ownerId, measures, fields, state, expressionStack, depth + 1, ref nodeCount, result);
                    inferred = InferDivision(weightedNumerator, weightedDenominator, path, result);
                    ValidateNodeResultType(weighted, inferred, path, result);
                    break;
                case ReferenceMeasureExpression reference:
                    MeasureDefinition target;
                    if (!measures.TryGetValue(reference.MeasureId ?? string.Empty, out target))
                    {
                        result.AddError("MEASURE_REFERENCE_UNKNOWN", path + ".measureId", "The referenced measure does not exist.");
                        inferred = null;
                    }
                    else
                    {
                        inferred = ValidateMeasureDefinition(target, measures, fields, state, result);
                        if (inferred.HasValue && reference.ResultType != inferred.Value)
                        {
                            result.AddError(
                                "REFERENCE_TYPE_MISMATCH",
                                path + ".resultType",
                                "A reference node must use the referenced measure's type.");
                        }
                    }

                    break;
                case ConstantMeasureExpression constant:
                    inferred = constant.ResultType;
                    break;
                case BinaryMeasureExpression binary:
                    ValidateEnum(binary.Operator, path + ".operator", "BINARY_MEASURE_OPERATOR_INVALID", result);
                    var left = ValidateExpression(binary.Left, ownerId, measures, fields, state, expressionStack, depth + 1, ref nodeCount, result);
                    var right = ValidateExpression(binary.Right, ownerId, measures, fields, state, expressionStack, depth + 1, ref nodeCount, result);
                    inferred = InferBinary(binary.Operator, left, right, path, result);
                    ValidateNodeResultType(binary, inferred, path, result);
                    break;
                case SafeDivideMeasureExpression divide:
                    ValidateEnum(divide.OnZero, path + ".onZero", "ZERO_DENOMINATOR_BEHAVIOR_INVALID", result);
                    var numerator = ValidateExpression(divide.Numerator, ownerId, measures, fields, state, expressionStack, depth + 1, ref nodeCount, result);
                    var denominator = ValidateExpression(divide.Denominator, ownerId, measures, fields, state, expressionStack, depth + 1, ref nodeCount, result);
                    inferred = divide.AsPercentage
                        ? ValidateComparableRatio(numerator, denominator, path, result)
                        : InferDivision(numerator, denominator, path, result);
                    if (divide.AsPercentage && inferred.HasValue)
                    {
                        inferred = MeasureValueType.Percentage;
                    }

                    ValidateNodeResultType(divide, inferred, path, result);
                    break;
                case RatioMeasureExpression ratio:
                    ValidateEnum(ratio.OnZero, path + ".onZero", "ZERO_DENOMINATOR_BEHAVIOR_INVALID", result);
                    var ratioNumerator = ValidateExpression(ratio.Numerator, ownerId, measures, fields, state, expressionStack, depth + 1, ref nodeCount, result);
                    var ratioDenominator = ValidateExpression(ratio.Denominator, ownerId, measures, fields, state, expressionStack, depth + 1, ref nodeCount, result);
                    inferred = InferDivision(ratioNumerator, ratioDenominator, path, result);
                    ValidateNodeResultType(ratio, inferred, path, result);
                    break;
                case DifferenceMeasureExpression difference:
                    ValidateEnum(difference.DifferenceKind, path + ".differenceKind", "DIFFERENCE_KIND_INVALID", result);
                    ValidateEnum(difference.OnZero, path + ".onZero", "ZERO_DENOMINATOR_BEHAVIOR_INVALID", result);
                    var current = ValidateExpression(difference.Current, ownerId, measures, fields, state, expressionStack, depth + 1, ref nodeCount, result);
                    var baseline = ValidateExpression(difference.Baseline, ownerId, measures, fields, state, expressionStack, depth + 1, ref nodeCount, result);
                    inferred = InferDifference(difference.DifferenceKind, current, baseline, path, result);
                    ValidateNodeResultType(difference, inferred, path, result);
                    break;
                case ShareMeasureExpression share:
                    ValidateEnum(share.OnZero, path + ".onZero", "ZERO_DENOMINATOR_BEHAVIOR_INVALID", result);
                    ValidateEnum(share.Scope, path + ".scope", "SHARE_SCOPE_INVALID", result);
                    var part = ValidateExpression(share.Part, ownerId, measures, fields, state, expressionStack, depth + 1, ref nodeCount, result);
                    var whole = ValidateExpression(share.Whole, ownerId, measures, fields, state, expressionStack, depth + 1, ref nodeCount, result);
                    inferred = ValidateComparableRatio(part, whole, path, result);
                    if (inferred.HasValue)
                    {
                        inferred = MeasureValueType.Percentage;
                    }

                    ValidateNodeResultType(share, inferred, path, result);
                    break;
                default:
                    result.AddError("MEASURE_EXPRESSION_KIND_UNSUPPORTED", path, "The expression kind is not supported.");
                    inferred = null;
                    break;
            }

            expressionStack.Remove(expression);
            return inferred;
        }

        private static MeasureValueType? ValidateAggregateType(
            AggregateFunction function,
            MeasureValueType declared,
            string path,
            ValidationResult result)
        {
            ValidateEnum(function, path + ".function", "AGGREGATE_FUNCTION_INVALID", result);
            if ((function == AggregateFunction.Count || function == AggregateFunction.DistinctCount)
                && declared != MeasureValueType.WholeNumber)
            {
                result.AddError(
                    "COUNT_TYPE_INVALID",
                    path + ".resultType",
                    "Count and distinct-count expressions must produce a whole number.");
                return MeasureValueType.WholeNumber;
            }

            if (function == AggregateFunction.Average && declared == MeasureValueType.WholeNumber)
            {
                result.AddError(
                    "AVERAGE_TYPE_INVALID",
                    path + ".resultType",
                    "An average of whole-number values must produce a Number.");
                return MeasureValueType.Number;
            }

            return declared;
        }

        private static MeasureValueType? InferBinary(
            BinaryMeasureOperator operation,
            MeasureValueType? left,
            MeasureValueType? right,
            string path,
            ValidationResult result)
        {
            if (!left.HasValue || !right.HasValue)
            {
                return null;
            }

            if (operation == BinaryMeasureOperator.Divide)
            {
                return InferDivision(left, right, path, result);
            }

            if (operation == BinaryMeasureOperator.Add || operation == BinaryMeasureOperator.Subtract)
            {
                if (left == right)
                {
                    return left;
                }

                if (IsPlainNumber(left.Value) && IsPlainNumber(right.Value))
                {
                    return MeasureValueType.Number;
                }

                result.AddError("MEASURE_OPERAND_TYPE_MISMATCH", path, "Addition and subtraction require compatible operand types.");
                return null;
            }

            if (left == MeasureValueType.Percentage)
            {
                return right == MeasureValueType.WholeNumber ? MeasureValueType.Number : right;
            }

            if (right == MeasureValueType.Percentage)
            {
                return left == MeasureValueType.WholeNumber ? MeasureValueType.Number : left;
            }

            if (left == MeasureValueType.Currency && IsPlainNumber(right.Value)
                || right == MeasureValueType.Currency && IsPlainNumber(left.Value))
            {
                return MeasureValueType.Currency;
            }

            if (IsPlainNumber(left.Value) && IsPlainNumber(right.Value))
            {
                return MeasureValueType.Number;
            }

            result.AddError("MEASURE_OPERAND_TYPE_MISMATCH", path, "Multiplication requires compatible numeric operand types.");
            return null;
        }

        private static MeasureValueType? InferDivision(
            MeasureValueType? numerator,
            MeasureValueType? denominator,
            string path,
            ValidationResult result)
        {
            if (!numerator.HasValue || !denominator.HasValue)
            {
                return null;
            }

            if (numerator == MeasureValueType.Currency && IsPlainNumber(denominator.Value))
            {
                return MeasureValueType.Currency;
            }

            if (numerator == denominator
                || IsPlainNumber(numerator.Value) && IsPlainNumber(denominator.Value))
            {
                return MeasureValueType.Number;
            }

            if (numerator == MeasureValueType.Percentage && IsPlainNumber(denominator.Value))
            {
                return MeasureValueType.Percentage;
            }

            result.AddError("MEASURE_OPERAND_TYPE_MISMATCH", path, "Division requires compatible numeric operand types.");
            return null;
        }

        private static MeasureValueType? ValidateComparableRatio(
            MeasureValueType? numerator,
            MeasureValueType? denominator,
            string path,
            ValidationResult result)
        {
            if (!numerator.HasValue || !denominator.HasValue)
            {
                return null;
            }

            var comparable = numerator == denominator
                || IsPlainNumber(numerator.Value) && IsPlainNumber(denominator.Value);
            if (!comparable)
            {
                result.AddError("RATIO_OPERAND_TYPE_MISMATCH", path, "A ratio requires comparable numerator and denominator types.");
                return null;
            }

            return MeasureValueType.Percentage;
        }

        private static MeasureValueType? InferDifference(
            DifferenceKind kind,
            MeasureValueType? current,
            MeasureValueType? baseline,
            string path,
            ValidationResult result)
        {
            if (!current.HasValue || !baseline.HasValue)
            {
                return null;
            }

            if (kind == DifferenceKind.PercentagePoints)
            {
                if (current != MeasureValueType.Percentage || baseline != MeasureValueType.Percentage)
                {
                    result.AddError(
                        "PERCENTAGE_POINT_TYPE_MISMATCH",
                        path,
                        "A percentage-point difference requires two percentage expressions.");
                    return null;
                }

                return MeasureValueType.Percentage;
            }

            if (kind == DifferenceKind.Percentage)
            {
                return ValidateComparableRatio(current, baseline, path, result);
            }

            if (current == baseline)
            {
                return current;
            }

            if (IsPlainNumber(current.Value) && IsPlainNumber(baseline.Value))
            {
                return MeasureValueType.Number;
            }

            result.AddError("DIFFERENCE_TYPE_MISMATCH", path, "An absolute difference requires compatible types.");
            return null;
        }

        private static void ValidateNodeResultType(
            MeasureExpression expression,
            MeasureValueType? inferred,
            string path,
            ValidationResult result)
        {
            if (inferred.HasValue && expression.ResultType != inferred.Value)
            {
                result.AddError(
                    "EXPRESSION_RESULT_TYPE_MISMATCH",
                    path + ".resultType",
                    "The node's declared resultType does not match the operation's result type.");
            }
        }

        private static bool IsPlainNumber(MeasureValueType value)
        {
            return value == MeasureValueType.WholeNumber || value == MeasureValueType.Number;
        }

        private static void ValidateMeasureFilters(
            List<MeasureFilterSpec> filters,
            HashSet<string>? fields,
            string path,
            ValidationResult result,
            bool requireOne = true)
        {
            if (filters == null)
            {
                result.AddError("MEASURE_FILTERS_REQUIRED", path, "The measure filter collection is required.");
                return;
            }

            if (requireOne && filters.Count == 0)
            {
                result.AddError("MEASURE_FILTER_REQUIRED", path, "A filtered aggregate requires at least one filter.");
                return;
            }

            if (filters.Count > 32)
            {
                result.AddError("TOO_MANY_MEASURE_FILTERS", path, "An aggregate may contain at most 32 filters.");
            }


            var literalCount = filters.Where(filter => filter != null && filter.Values != null)
                .Sum(filter => (long)filter.Values.Count);
            if (literalCount > MaximumFilterLiterals)
            {
                result.AddError(
                    "TOO_MANY_MEASURE_FILTER_VALUES",
                    path,
                    "A filtered aggregate may contain at most 256 literal values across all filters.");
            }

            long filterCombinations = 1L;
            foreach (var setFilter in filters.Where(filter => filter != null
                && filter.Values != null
                && (filter.Operator == MeasureFilterOperator.In || filter.Operator == MeasureFilterOperator.NotIn)))
            {
                var setSize = Math.Max(1, setFilter.Values.Count);
                if (filterCombinations > MaximumFilterLiterals / setSize)
                {
                    filterCombinations = MaximumFilterLiterals + 1L;
                    break;
                }

                filterCombinations *= setSize;
            }

            if (filterCombinations > MaximumFilterLiterals)
            {
                result.AddError(
                    "MEASURE_FILTER_COMBINATIONS_TOO_LARGE",
                    path,
                    "Set filters may expand to at most 256 deterministic filter combinations.");
            }

            var filterKeys = new HashSet<string>(StringComparer.Ordinal);
            var filterFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (var index = 0; index < filters.Count; index++)
            {
                var filter = filters[index];
                var filterPath = path + "[" + index.ToString(CultureInfo.InvariantCulture) + "]";
                if (filter == null)
                {
                    result.AddError("MEASURE_FILTER_REQUIRED", filterPath, "A measure filter cannot be null.");
                    continue;
                }

                ValidateFieldReference(filter.Field, fields, filterPath + ".field", result);
                if (!string.IsNullOrWhiteSpace(filter.Field) && !filterFields.Add(filter.Field))
                {
                    result.AddError(
                        "MEASURE_FILTER_FIELD_DUPLICATE",
                        filterPath + ".field",
                        "A filtered aggregate may define at most one condition per field.");
                }
                if (!Enum.IsDefined(typeof(MeasureFilterOperator), filter.Operator))
                {
                    result.AddError("MEASURE_FILTER_OPERATOR_INVALID", filterPath + ".operator", "The measure filter operator is not supported.");
                }

                var blankOperator = filter.Operator == MeasureFilterOperator.IsBlank
                    || filter.Operator == MeasureFilterOperator.IsNotBlank;
                var setOperator = filter.Operator == MeasureFilterOperator.In
                    || filter.Operator == MeasureFilterOperator.NotIn;
                if (filter.Values == null)
                {
                    result.AddError("MEASURE_FILTER_VALUES_REQUIRED", filterPath + ".values", "The filter value collection is required.");
                    continue;
                }

                if (blankOperator && filter.Values.Count != 0)
                {
                    result.AddError("MEASURE_FILTER_VALUES_NOT_ALLOWED", filterPath + ".values", "Blank filters cannot contain values.");
                }
                else if (setOperator && (filter.Values.Count == 0 || filter.Values.Count > 256))
                {
                    result.AddError("MEASURE_FILTER_SET_SIZE_INVALID", filterPath + ".values", "In and NotIn filters require 1-256 values.");
                }
                else if (!blankOperator && !setOperator && filter.Values.Count != 1)
                {
                    result.AddError("MEASURE_FILTER_SINGLE_VALUE_REQUIRED", filterPath + ".values", "This filter requires exactly one value.");
                }

                var valueKeys = new HashSet<string>(StringComparer.Ordinal);
                foreach (var value in filter.Values)
                {
                    ValidateScalar(value, filterPath + ".values", result);
                    if (value != null && value.Kind == ScalarValueKind.Null)
                    {
                        result.AddError("MEASURE_FILTER_NULL_LITERAL_NOT_ALLOWED", filterPath + ".values", "Use IsBlank or IsNotBlank instead of a null literal.");
                    }

                    if (value != null && !valueKeys.Add(ScalarKey(value)))
                    {
                        result.AddError("MEASURE_FILTER_VALUE_DUPLICATE", filterPath + ".values", "Filter values must be unique.");
                    }
                }

                var filterKey = MeasureFilterKey(filter);
                if (!filterKeys.Add(filterKey))
                {
                    result.AddError("MEASURE_FILTER_DUPLICATE", filterPath, "Duplicate aggregate filters are not allowed.");
                }
            }
        }

        private static HashSet<string> ValidateStyles(List<PresentationStyleSpec> styles, ValidationResult result)
        {
            var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (styles == null)
            {
                result.AddError("STYLES_REQUIRED", "$.styles", "The styles collection is required.");
                return ids;
            }

            if (styles.Count > MaximumStyles)
            {
                result.AddError("TOO_MANY_STYLES", "$.styles", "A report may define at most 128 presentation styles.");
            }

            for (var index = 0; index < styles.Count; index++)
            {
                var style = styles[index];
                var path = "$.styles[" + index.ToString(CultureInfo.InvariantCulture) + "]";
                if (style == null)
                {
                    result.AddError("STYLE_REQUIRED", path, "A presentation style cannot be null.");
                    continue;
                }

                ValidateId(style.Id, path + ".id", "STYLE_ID_INVALID", result);
                if (!ids.Add(style.Id ?? string.Empty))
                {
                    result.AddError("STYLE_ID_DUPLICATE", path + ".id", "Style IDs must be unique.");
                }

                ValidateColor(style.FontColor, path + ".fontColor", result);
                ValidateColor(style.FillColor, path + ".fillColor", result);
                ValidateEnum(style.HorizontalAlignment, path + ".horizontalAlignment", "HORIZONTAL_ALIGNMENT_INVALID", result);
                ValidateNumberFormat(style.NumberFormat, path + ".numberFormat", result);
                if (style.DecimalPlaces.HasValue && (style.DecimalPlaces.Value < 0 || style.DecimalPlaces.Value > 15))
                {
                    result.AddError("DECIMAL_PLACES_INVALID", path + ".decimalPlaces", "Decimal places must be between 0 and 15.");
                }
            }

            return ids;
        }

        private static void ValidateWeightedConstructions(
            IEnumerable<MeasureDefinition> measures,
            IEnumerable<TransformStep> transforms,
            ValidationResult result)
        {
            if (measures == null || transforms == null)
            {
                return;
            }

            var arithmetic = new Dictionary<string, AddArithmeticColumnTransform>(StringComparer.OrdinalIgnoreCase);
            foreach (var transform in transforms.OfType<AddArithmeticColumnTransform>())
            {
                if (!string.IsNullOrWhiteSpace(transform.OutputColumn)
                    && !arithmetic.ContainsKey(transform.OutputColumn))
                {
                    arithmetic.Add(transform.OutputColumn, transform);
                }
            }
            foreach (var measure in measures)
            {
                if (measure == null)
                {
                    continue;
                }

                ValidateWeightedExpression(measure.Expression, measure.Id, arithmetic, result);
            }
        }

        private static void ValidateWeightedExpression(
            MeasureExpression expression,
            string measureId,
            Dictionary<string, AddArithmeticColumnTransform> arithmetic,
            ValidationResult result)
        {
            switch (expression)
            {
                case WeightedAggregateMeasureExpression weighted:
                    AggregateDetails numerator;
                    AggregateDetails denominator;
                    var path = "$.measures[" + measureId + "].expression";
                    if (!TryGetAggregateDetails(weighted.Numerator, out numerator)
                        || !TryGetAggregateDetails(weighted.Denominator, out denominator))
                    {
                        result.AddError(
                            "WEIGHTED_AGGREGATES_REQUIRED",
                            path,
                            "A weighted aggregate requires direct typed numerator and denominator aggregates.");
                        break;
                    }

                    if (numerator.Function != AggregateFunction.Sum
                        || denominator.Function != AggregateFunction.Sum)
                    {
                        result.AddError(
                            "WEIGHTED_SUM_REQUIRED",
                            path,
                            "Weighted numerator and denominator components must both use Sum.");
                    }

                    AddArithmeticColumnTransform product;
                    if (!arithmetic.TryGetValue(numerator.Field, out product)
                        || product.Operator != ArithmeticOperator.Multiply
                        || !IsColumnOperand(product.Left)
                        || !IsColumnOperand(product.Right)
                        || !string.Equals(product.Left.Column, denominator.Field, StringComparison.OrdinalIgnoreCase)
                            && !string.Equals(product.Right.Column, denominator.Field, StringComparison.OrdinalIgnoreCase))
                    {
                        result.AddError(
                            "WEIGHTED_PRODUCT_TRANSFORM_REQUIRED",
                            path + ".numerator",
                            "The weighted numerator must sum a typed multiplication column whose operands include the denominator weight field.");
                    }

                    if (!string.Equals(numerator.PeriodSliceId, denominator.PeriodSliceId, StringComparison.OrdinalIgnoreCase)
                        || !MeasureFiltersEquivalent(numerator.Filters, denominator.Filters))
                    {
                        result.AddError(
                            "WEIGHTED_SCOPE_MISMATCH",
                            path,
                            "Weighted numerator and denominator aggregates must use identical filters and period slices.");
                    }

                    ValidateWeightedExpression(weighted.Numerator, measureId, arithmetic, result);
                    ValidateWeightedExpression(weighted.Denominator, measureId, arithmetic, result);
                    break;
                case BinaryMeasureExpression binary:
                    ValidateWeightedExpression(binary.Left, measureId, arithmetic, result);
                    ValidateWeightedExpression(binary.Right, measureId, arithmetic, result);
                    break;
                case SafeDivideMeasureExpression divide:
                    ValidateWeightedExpression(divide.Numerator, measureId, arithmetic, result);
                    ValidateWeightedExpression(divide.Denominator, measureId, arithmetic, result);
                    break;
                case RatioMeasureExpression ratio:
                    ValidateWeightedExpression(ratio.Numerator, measureId, arithmetic, result);
                    ValidateWeightedExpression(ratio.Denominator, measureId, arithmetic, result);
                    break;
                case DifferenceMeasureExpression difference:
                    ValidateWeightedExpression(difference.Current, measureId, arithmetic, result);
                    ValidateWeightedExpression(difference.Baseline, measureId, arithmetic, result);
                    break;
                case ShareMeasureExpression share:
                    ValidateWeightedExpression(share.Part, measureId, arithmetic, result);
                    ValidateWeightedExpression(share.Whole, measureId, arithmetic, result);
                    break;
            }
        }

        private static bool IsColumnOperand(ArithmeticOperand operand)
        {
            return operand != null
                && operand.Kind == ArithmeticOperandKind.Column
                && !string.IsNullOrWhiteSpace(operand.Column);
        }

        private static bool TryGetAggregateDetails(MeasureExpression expression, out AggregateDetails details)
        {
            if (expression is AggregateMeasureExpression aggregate)
            {
                details = new AggregateDetails(
                    aggregate.Field,
                    aggregate.Function,
                    aggregate.PeriodSliceId,
                    new List<MeasureFilterSpec>());
                return true;
            }

            if (expression is FilteredAggregateMeasureExpression filtered)
            {
                details = new AggregateDetails(
                    filtered.Field,
                    filtered.Function,
                    filtered.PeriodSliceId,
                    filtered.Filters);
                return true;
            }

            details = new AggregateDetails(string.Empty, AggregateFunction.Sum, null, new List<MeasureFilterSpec>());
            return false;
        }

        private static bool MeasureFiltersEquivalent(
            IReadOnlyList<MeasureFilterSpec> left,
            IReadOnlyList<MeasureFilterSpec> right)
        {
            if (left.Count != right.Count)
            {
                return false;
            }

            var leftKeys = left.Select(MeasureFilterKey).OrderBy(value => value, StringComparer.Ordinal);
            var rightKeys = right.Select(MeasureFilterKey).OrderBy(value => value, StringComparer.Ordinal);
            return leftKeys.SequenceEqual(rightKeys, StringComparer.Ordinal);
        }

        private static string MeasureFilterKey(MeasureFilterSpec filter)
        {
            if (filter == null)
            {
                return "<null>";
            }

            var values = filter.Values == null
                ? Enumerable.Empty<string>()
                : filter.Values.Select(ScalarKey).OrderBy(value => value, StringComparer.Ordinal);
            return (filter.Field ?? string.Empty).ToUpperInvariant()
                + "|" + filter.Operator
                + "|" + string.Join(",", values);
        }

        private sealed class AggregateDetails
        {
            public AggregateDetails(
                string field,
                AggregateFunction function,
                string? periodSliceId,
                IReadOnlyList<MeasureFilterSpec> filters)
            {
                Field = field;
                Function = function;
                PeriodSliceId = periodSliceId;
                Filters = filters;
            }

            public string Field { get; }

            public AggregateFunction Function { get; }

            public string? PeriodSliceId { get; }

            public IReadOnlyList<MeasureFilterSpec> Filters { get; }
        }

        private static void ValidateBlocks(
            List<ReportBlockSpec> blocks,
            string rootOwnershipId,
            Dictionary<string, MeasureDefinition> measures,
            HashSet<string> styles,
            HashSet<string>? fields,
            ValidationResult result)
        {
            if (blocks == null || blocks.Count == 0)
            {
                result.AddError("REPORT_BLOCK_REQUIRED", "$.blocks", "At least one independently anchored report block is required.");
                return;
            }

            if (blocks.Count > MaximumBlocks)
            {
                result.AddError("TOO_MANY_REPORT_BLOCKS", "$.blocks", "A report may contain at most 64 blocks.");
            }

            var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var ownershipIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (!string.IsNullOrWhiteSpace(rootOwnershipId))
            {
                ownershipIds.Add(rootOwnershipId);
            }

            var ownedRanges = new List<OwnedBlockRange>();
            for (var index = 0; index < blocks.Count; index++)
            {
                var block = blocks[index];
                var path = "$.blocks[" + index.ToString(CultureInfo.InvariantCulture) + "]";
                if (block == null)
                {
                    result.AddError("REPORT_BLOCK_REQUIRED", path, "A report block cannot be null.");
                    continue;
                }

                ValidateId(block.Id, path + ".id", "BLOCK_ID_INVALID", result);
                ValidateId(block.OwnershipId, path + ".ownershipId", "BLOCK_OWNERSHIP_ID_INVALID", result);
                ValidateOptionalText(block.Title, path + ".title", 120, result);
                if (!ids.Add(block.Id ?? string.Empty))
                {
                    result.AddError("BLOCK_ID_DUPLICATE", path + ".id", "Block IDs must be unique.");
                }

                if (!ownershipIds.Add(block.OwnershipId ?? string.Empty))
                {
                    result.AddError(
                        "BLOCK_OWNERSHIP_ID_DUPLICATE",
                        path + ".ownershipId",
                        "Block ownership IDs must be unique and cannot reuse the report ownership ID.");
                }

                ValidateWorksheetName(block.WorksheetName, path + ".worksheetName", result);
                int anchorRow;
                int anchorColumn;
                if (!TryParseCell(block.AnchorCell, out anchorRow, out anchorColumn))
                {
                    result.AddError(
                        "ANCHOR_CELL_INVALID",
                        path + ".anchorCell",
                        "The block anchor must be an A1-style cell address within worksheet bounds.");
                }
                else if (block.OwnedExtent == null)
                {
                    result.AddError("OWNED_EXTENT_REQUIRED", path + ".ownedExtent", "Every block requires an owned write extent.");
                }
                else if (block.OwnedExtent.RowCount < 1 || block.OwnedExtent.ColumnCount < 1)
                {
                    result.AddError(
                        "OWNED_EXTENT_INVALID",
                        path + ".ownedExtent",
                        "Owned row and column counts must both be positive.");
                }
                else
                {
                    var endRow = (long)anchorRow + block.OwnedExtent.RowCount - 1L;
                    var endColumn = (long)anchorColumn + block.OwnedExtent.ColumnCount - 1L;
                    if (endRow > RowProjection.ExcelWorksheetRowLimit || endColumn > 16384L)
                    {
                        result.AddError(
                            "OWNED_EXTENT_OUT_OF_BOUNDS",
                            path + ".ownedExtent",
                            "The owned block extent runs past worksheet bounds.");
                    }
                    else
                    {
                        var ownedRange = new OwnedBlockRange(
                            block.Id ?? string.Empty,
                            block.WorksheetName ?? string.Empty,
                            anchorRow,
                            (int)endRow,
                            anchorColumn,
                            (int)endColumn);
                        var overlap = ownedRanges.FirstOrDefault(existing => existing.Overlaps(ownedRange));
                        if (overlap != null)
                        {
                            result.AddError(
                                "BLOCK_OWNED_RANGE_OVERLAP",
                                path + ".ownedExtent",
                                "Owned block ranges for '" + overlap.BlockId + "' and '" + block.Id + "' overlap.");
                        }

                        ownedRanges.Add(ownedRange);
                    }
                }

                if (!Enum.IsDefined(typeof(ReportOutputMode), block.OutputMode))
                {
                    result.AddError("OUTPUT_MODE_INVALID", path + ".outputMode", "The report output mode is not supported.");
                }

                ValidateStyleReference(block.HeaderStyleId, styles, path + ".headerStyleId", result);
                ValidateStyleReference(block.BodyStyleId, styles, path + ".bodyStyleId", result);
                ValidateStyleReference(block.SubtotalStyleId, styles, path + ".subtotalStyleId", result);
                ValidateStyleReference(block.GrandTotalStyleId, styles, path + ".grandTotalStyleId", result);
                ValidateSlices(block.PeriodSlices, path + ".periodSlices", result);
                ValidateLayout(block.Layout, measures, styles, fields, block.PeriodSlices, path + ".layout", result);
                ValidateHeaders(block.Headers, styles, block.OwnedExtent, path + ".headers", result);
                ValidateSpacers(block.Spacers, path + ".spacers", result);
            }
        }

        private static bool TryParseCell(string? address, out int row, out int column)
        {
            row = 0;
            column = 0;
            if (string.IsNullOrWhiteSpace(address))
            {
                return false;
            }

            var match = CellPattern.Match(address!);
            if (!match.Success
                || !int.TryParse(match.Groups["row"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out row))
            {
                return false;
            }

            foreach (var character in match.Groups["column"].Value.ToUpperInvariant())
            {
                column = checked(column * 26 + character - 'A' + 1);
            }

            return row >= 1
                && row <= RowProjection.ExcelWorksheetRowLimit
                && column >= 1
                && column <= 16384;
        }

        private sealed class OwnedBlockRange
        {
            public OwnedBlockRange(
                string blockId,
                string worksheetName,
                int startRow,
                int endRow,
                int startColumn,
                int endColumn)
            {
                BlockId = blockId;
                WorksheetName = worksheetName;
                StartRow = startRow;
                EndRow = endRow;
                StartColumn = startColumn;
                EndColumn = endColumn;
            }

            public string BlockId { get; }

            public string WorksheetName { get; }

            public int StartRow { get; }

            public int EndRow { get; }

            public int StartColumn { get; }

            public int EndColumn { get; }

            public bool Overlaps(OwnedBlockRange other)
            {
                return string.Equals(WorksheetName, other.WorksheetName, StringComparison.OrdinalIgnoreCase)
                    && StartRow <= other.EndRow
                    && EndRow >= other.StartRow
                    && StartColumn <= other.EndColumn
                    && EndColumn >= other.StartColumn;
            }
        }

        private static void ValidateLayout(
            ReportLayoutSpec layout,
            Dictionary<string, MeasureDefinition> measures,
            HashSet<string> styles,
            HashSet<string>? fields,
            List<PeriodSliceSpec> slices,
            string path,
            ValidationResult result)
        {
            if (layout == null)
            {
                result.AddError("BLOCK_LAYOUT_REQUIRED", path, "A block layout is required.");
                return;
            }

            if (layout.Rows != null && layout.Rows.Count > MaximumAxisFields)
            {
                result.AddError("TOO_MANY_ROW_FIELDS", path + ".rows", "A block may contain at most 32 row fields.");
            }

            if (layout.Columns != null && layout.Columns.Count > MaximumAxisFields)
            {
                result.AddError("TOO_MANY_COLUMN_FIELDS", path + ".columns", "A block may contain at most 32 column fields.");
            }

            if (layout.Values != null && layout.Values.Count > MaximumValues)
            {
                result.AddError("TOO_MANY_VALUES", path + ".values", "A block may contain at most 128 Values.");
            }

            if (layout.Filters != null && layout.Filters.Count > MaximumFilters)
            {
                result.AddError("TOO_MANY_FILTERS", path + ".filters", "A block may contain at most 32 Filters.");
            }

            ValidateFieldPlacements(layout.Rows, measures, styles, fields, path + ".rows", result);
            ValidateFieldPlacements(layout.Columns, measures, styles, fields, path + ".columns", result);
            ValidateFieldPlacements(layout.Filters, fields, path + ".filters", result);
            ValidatePlacementAxisUniqueness(layout, path, result);
            var sliceIds = new HashSet<string>(
                (slices ?? new List<PeriodSliceSpec>())
                    .Where(slice => slice != null)
                    .Select(slice => slice.Id),
                StringComparer.OrdinalIgnoreCase);
            ValidateTopNMeasureSlices(layout.Rows, measures, sliceIds, path + ".rows", result);
            ValidateTopNMeasureSlices(layout.Columns, measures, sliceIds, path + ".columns", result);
            if (layout.Values == null || layout.Values.Count == 0)
            {
                result.AddError("BLOCK_VALUE_REQUIRED", path + ".values", "A report block requires at least one Value.");
            }
            else
            {
                var valueKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                for (var index = 0; index < layout.Values.Count; index++)
                {
                    var value = layout.Values[index];
                    var valuePath = path + ".values[" + index.ToString(CultureInfo.InvariantCulture) + "]";
                    if (value == null)
                    {
                        result.AddError("VALUE_REQUIRED", valuePath, "A Value placement cannot be null.");
                        continue;
                    }

                    var placementSliceIds = value.PeriodSliceIds ?? new List<string>();
                    if (value.PeriodSliceIds == null)
                    {
                        result.AddError("VALUE_SLICES_REQUIRED", valuePath + ".periodSliceIds", "The Value period-slice collection is required.");
                    }

                    if (placementSliceIds.Count > MaximumPeriodSlices)
                    {
                        result.AddError(
                            "TOO_MANY_VALUE_SLICES",
                            valuePath + ".periodSliceIds",
                            "A Value may reference at most 64 period slices.");
                    }

                    var valueKey = value.MeasureId + "|" + string.Join(",", placementSliceIds.OrderBy(id => id, StringComparer.OrdinalIgnoreCase));
                    if (!valueKeys.Add(valueKey))
                    {
                        result.AddError("VALUE_PLACEMENT_DUPLICATE", valuePath, "The same measure and period-slice placement cannot be repeated.");
                    }

                    MeasureDefinition placedMeasure;
                    if (!measures.TryGetValue(value.MeasureId ?? string.Empty, out placedMeasure))
                    {
                        result.AddError("VALUE_MEASURE_UNKNOWN", valuePath + ".measureId", "The Value references an unknown measure.");
                    }
                    else
                    {
                        var expressionSlices = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        CollectExpressionSliceIds(
                            placedMeasure.Expression,
                            measures,
                            new HashSet<string>(StringComparer.OrdinalIgnoreCase),
                            expressionSlices);
                        foreach (var expressionSlice in expressionSlices)
                        {
                            if (!sliceIds.Contains(expressionSlice))
                            {
                                result.AddError(
                                    "MEASURE_SLICE_UNKNOWN_IN_BLOCK",
                                    valuePath + ".measureId",
                                    "The measure requires period slice '" + expressionSlice + "', which is not defined in this block.");
                            }
                        }

                        if (expressionSlices.Count != 0 && placementSliceIds.Count != 0)
                        {
                            result.AddError(
                                "VALUE_SLICE_CONTEXT_CONFLICT",
                                valuePath + ".periodSliceIds",
                                "A slice-bound measure cannot also be expanded across placement period slices.");
                        }

                        var expressionFilterFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        CollectExpressionFilterFields(
                            placedMeasure.Expression,
                            measures,
                            new HashSet<string>(StringComparer.OrdinalIgnoreCase),
                            expressionFilterFields);
                        var layoutFields = new HashSet<string>(
                            (layout.Rows ?? new List<FieldPlacementSpec>())
                                .Where(field => field != null)
                                .Select(field => field.Field)
                                .Concat((layout.Columns ?? new List<FieldPlacementSpec>())
                                    .Where(field => field != null)
                                    .Select(field => field.Field))
                                .Concat((layout.Filters ?? new List<FilterPlacementSpec>())
                                    .Where(filter => filter != null)
                                    .Select(filter => filter.Field)),
                            StringComparer.OrdinalIgnoreCase);
                        if (expressionFilterFields.Overlaps(layoutFields))
                        {
                            result.AddError(
                                "MEASURE_FILTER_LAYOUT_FIELD_CONFLICT",
                                valuePath + ".measureId",
                                "A measure filter cannot reuse a row, column, or report Filter field in the same block.");
                        }
                    }

                    ValidateNumberFormat(value.NumberFormat, valuePath + ".numberFormat", result);
                    ValidateOptionalText(value.Caption, valuePath + ".caption", 120, result);
                    ValidateStyleReference(value.StyleId, styles, valuePath + ".styleId", result);
                    ValidateDistinct(placementSliceIds, valuePath + ".periodSliceIds", "VALUE_SLICE_DUPLICATE", result);
                    foreach (var sliceId in placementSliceIds)
                    {
                        if (!sliceIds.Contains(sliceId))
                        {
                            result.AddError("VALUE_SLICE_UNKNOWN", valuePath + ".periodSliceIds", "The Value references an unknown period slice.");
                        }
                    }
                }
            }

            if (layout.DenseLayout == null)
            {
                result.AddError("DENSE_LAYOUT_REQUIRED", path + ".denseLayout", "Dense layout options are required.");
            }
            else if (layout.DenseLayout.RowIndent < 0 || layout.DenseLayout.RowIndent > 15)
            {
                result.AddError("ROW_INDENT_INVALID", path + ".denseLayout.rowIndent", "Row indent must be between 0 and 15.");
            }

            if (layout.GrandTotals == null)
            {
                result.AddError("GRAND_TOTALS_REQUIRED", path + ".grandTotals", "Grand-total options are required.");
            }
            else
            {
                ValidateStyleReference(layout.GrandTotals.StyleId, styles, path + ".grandTotals.styleId", result);
                ValidateEnum(layout.GrandTotals.RowPlacement, path + ".grandTotals.rowPlacement", "GRAND_TOTAL_PLACEMENT_INVALID", result);
                ValidateEnum(layout.GrandTotals.ColumnPlacement, path + ".grandTotals.columnPlacement", "GRAND_TOTAL_PLACEMENT_INVALID", result);
                ValidateRequiredText(layout.GrandTotals.RowLabel, path + ".grandTotals.rowLabel", 120, result);
                ValidateRequiredText(layout.GrandTotals.ColumnLabel, path + ".grandTotals.columnLabel", 120, result);
                if (layout.DenseLayout != null
                    && (layout.DenseLayout.ShowRowGrandTotals != layout.GrandTotals.ShowRows
                        || layout.DenseLayout.ShowColumnGrandTotals != layout.GrandTotals.ShowColumns))
                {
                    result.AddError(
                        "GRAND_TOTAL_OPTIONS_CONFLICT",
                        path,
                        "Dense-layout and grand-total visibility options must agree.");
                }
            }
        }

        private static void ValidatePlacementAxisUniqueness(
            ReportLayoutSpec layout,
            string path,
            ValidationResult result)
        {
            var axes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            AddAxisFields(layout.Rows, "rows", axes, path, result);
            AddAxisFields(layout.Columns, "columns", axes, path, result);
            if (layout.Filters != null)
            {
                foreach (var filter in layout.Filters.Where(filter => filter != null))
                {
                    if (string.IsNullOrWhiteSpace(filter.Field))
                    {
                        continue;
                    }

                    string priorAxis;
                    if (axes.TryGetValue(filter.Field, out priorAxis))
                    {
                        result.AddError(
                            "FIELD_USED_ON_MULTIPLE_AXES",
                            path + ".filters",
                            "Field '" + filter.Field + "' is already placed on " + priorAxis + ".");
                    }
                    else
                    {
                        axes.Add(filter.Field, "filters");
                    }
                }
            }
        }

        private static void AddAxisFields(
            IEnumerable<FieldPlacementSpec>? placements,
            string axis,
            Dictionary<string, string> axes,
            string path,
            ValidationResult result)
        {
            if (placements == null)
            {
                return;
            }

            foreach (var placement in placements.Where(placement => placement != null))
            {
                if (string.IsNullOrWhiteSpace(placement.Field))
                {
                    continue;
                }

                string priorAxis;
                if (axes.TryGetValue(placement.Field, out priorAxis))
                {
                    result.AddError(
                        "FIELD_USED_ON_MULTIPLE_AXES",
                        path + "." + axis,
                        "Field '" + placement.Field + "' is already placed on " + priorAxis + ".");
                }
                else
                {
                    axes.Add(placement.Field, axis);
                }
            }
        }

        private static void ValidateTopNMeasureSlices(
            List<FieldPlacementSpec>? placements,
            Dictionary<string, MeasureDefinition> measures,
            HashSet<string> sliceIds,
            string path,
            ValidationResult result)
        {
            if (placements == null)
            {
                return;
            }

            for (var index = 0; index < placements.Count; index++)
            {
                var placement = placements[index];
                if (placement == null || placement.TopN == null)
                {
                    continue;
                }

                MeasureDefinition measure;
                if (!measures.TryGetValue(placement.TopN.MeasureId ?? string.Empty, out measure))
                {
                    continue;
                }

                var requiredSlices = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                CollectExpressionSliceIds(
                    measure.Expression,
                    measures,
                    new HashSet<string>(StringComparer.OrdinalIgnoreCase),
                    requiredSlices);
                foreach (var requiredSlice in requiredSlices)
                {
                    if (!sliceIds.Contains(requiredSlice))
                    {
                        result.AddError(
                            "TOP_N_MEASURE_SLICE_UNKNOWN_IN_BLOCK",
                            path + "[" + index.ToString(CultureInfo.InvariantCulture) + "].topN.measureId",
                            "The Top N measure requires period slice '" + requiredSlice + "', which is not defined in this block.");
                    }
                }
            }
        }

        private static void ValidateFieldPlacements(
            List<FieldPlacementSpec>? placements,
            Dictionary<string, MeasureDefinition> measures,
            HashSet<string> styles,
            HashSet<string>? fields,
            string path,
            ValidationResult result)
        {
            if (placements == null)
            {
                result.AddError("FIELD_PLACEMENTS_REQUIRED", path, "The field-placement collection is required.");
                return;
            }

            var placed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var index = 0; index < placements.Count; index++)
            {
                var placement = placements[index];
                var itemPath = path + "[" + index.ToString(CultureInfo.InvariantCulture) + "]";
                if (placement == null)
                {
                    result.AddError("FIELD_PLACEMENT_REQUIRED", itemPath, "A field placement cannot be null.");
                    continue;
                }

                ValidateFieldReference(placement.Field, fields, itemPath + ".field", result);
                ValidateOptionalText(placement.Caption, itemPath + ".caption", 120, result);
                ValidateEnum(placement.Sort, itemPath + ".sort", "SORT_DIRECTION_INVALID", result);
                if (!placed.Add(placement.Field ?? string.Empty))
                {
                    result.AddError("FIELD_PLACEMENT_DUPLICATE", itemPath + ".field", "A field can appear only once in this placement axis.");
                }

                if (placement.Subtotals == null)
                {
                    result.AddError("SUBTOTAL_SPEC_REQUIRED", itemPath + ".subtotals", "Subtotal options are required for each row level.");
                }
                else
                {
                    ValidateStyleReference(placement.Subtotals.StyleId, styles, itemPath + ".subtotals.styleId", result);
                    if (!Enum.IsDefined(typeof(SubtotalMode), placement.Subtotals.Mode)
                        || !Enum.IsDefined(typeof(TotalPlacement), placement.Subtotals.Placement))
                    {
                        result.AddError("SUBTOTAL_OPTION_INVALID", itemPath + ".subtotals", "The subtotal mode or placement is not supported.");
                    }

                    if (placement.Subtotals.Mode == SubtotalMode.None
                        && (placement.Subtotals.Label != null
                            || placement.Subtotals.StyleId != null))
                    {
                        result.AddError(
                            "DISABLED_SUBTOTAL_PRESENTATION_NOT_ALLOWED",
                            itemPath + ".subtotals",
                            "A disabled subtotal cannot define a label or style.");
                    }
                    else if (placement.Subtotals.Label != null)
                    {
                        ValidateRequiredText(placement.Subtotals.Label, itemPath + ".subtotals.label", 120, result);
                    }
                }

                ValidateScalarListUnique(placement.MemberOrder, itemPath + ".memberOrder", result);
                if (placement.MemberOrder != null && placement.MemberOrder.Count > MaximumMembers)
                {
                    result.AddError("TOO_MANY_ORDERED_MEMBERS", itemPath + ".memberOrder", "Member order may contain at most 1,000 values.");
                }
                ValidateGroupBuckets(placement.GroupBuckets, itemPath + ".groupBuckets", result);
                if (placement.MemberOrder != null
                    && placement.MemberOrder.Count != 0
                    && placement.Sort != SortDirection.SourceOrder)
                {
                    result.AddError(
                        "MEMBER_ORDER_SORT_CONFLICT",
                        itemPath,
                        "Manual member order and ascending or descending sort cannot be combined.");
                }

                if (placement.TopN != null)
                {
                    if (!Enum.IsDefined(typeof(TopNDirection), placement.TopN.Direction))
                    {
                        result.AddError("TOP_N_DIRECTION_INVALID", itemPath + ".topN.direction", "The Top N direction is not supported.");
                    }

                    if (placement.TopN.Count < 1 || placement.TopN.Count > 1000)
                    {
                        result.AddError("TOP_N_COUNT_INVALID", itemPath + ".topN.count", "Top N count must be between 1 and 1000.");
                    }

                    MeasureDefinition topNMeasure;
                    if (!measures.TryGetValue(placement.TopN.MeasureId ?? string.Empty, out topNMeasure))
                    {
                        result.AddError("TOP_N_MEASURE_UNKNOWN", itemPath + ".topN.measureId", "Top N requires a known measure.");
                    }
                    else
                    {
                        var rankAggregate = topNMeasure.Expression as AggregateMeasureExpression;
                        if (rankAggregate == null
                            || rankAggregate.Function == AggregateFunction.DistinctCount
                            || !string.IsNullOrWhiteSpace(rankAggregate.PeriodSliceId))
                        {
                            result.AddError(
                                "TOP_N_MEASURE_NOT_RANKABLE",
                                itemPath + ".topN.measureId",
                                "Top N ranking requires an unfiltered, slice-independent aggregate measure.");
                        }
                    }

                    ValidateRequiredText(placement.TopN.OthersLabel, itemPath + ".topN.othersLabel", 120, result);
                }
            }
        }

        private static void CollectExpressionSliceIds(
            MeasureExpression expression,
            Dictionary<string, MeasureDefinition> measures,
            HashSet<string> visitedMeasures,
            HashSet<string> target)
        {
            switch (expression)
            {
                case AggregateMeasureExpression aggregate:
                    if (!string.IsNullOrWhiteSpace(aggregate.PeriodSliceId))
                    {
                        target.Add(aggregate.PeriodSliceId!);
                    }

                    break;
                case FilteredAggregateMeasureExpression filtered:
                    if (!string.IsNullOrWhiteSpace(filtered.PeriodSliceId))
                    {
                        target.Add(filtered.PeriodSliceId!);
                    }

                    break;
                case WeightedAggregateMeasureExpression weighted:
                    CollectExpressionSliceIds(weighted.Numerator, measures, visitedMeasures, target);
                    CollectExpressionSliceIds(weighted.Denominator, measures, visitedMeasures, target);
                    break;
                case ReferenceMeasureExpression reference:
                    MeasureDefinition referenced;
                    if (measures.TryGetValue(reference.MeasureId, out referenced)
                        && visitedMeasures.Add(reference.MeasureId))
                    {
                        CollectExpressionSliceIds(referenced.Expression, measures, visitedMeasures, target);
                        visitedMeasures.Remove(reference.MeasureId);
                    }

                    break;
                case BinaryMeasureExpression binary:
                    CollectExpressionSliceIds(binary.Left, measures, visitedMeasures, target);
                    CollectExpressionSliceIds(binary.Right, measures, visitedMeasures, target);
                    break;
                case SafeDivideMeasureExpression divide:
                    CollectExpressionSliceIds(divide.Numerator, measures, visitedMeasures, target);
                    CollectExpressionSliceIds(divide.Denominator, measures, visitedMeasures, target);
                    break;
                case RatioMeasureExpression ratio:
                    CollectExpressionSliceIds(ratio.Numerator, measures, visitedMeasures, target);
                    CollectExpressionSliceIds(ratio.Denominator, measures, visitedMeasures, target);
                    break;
                case DifferenceMeasureExpression difference:
                    CollectExpressionSliceIds(difference.Current, measures, visitedMeasures, target);
                    CollectExpressionSliceIds(difference.Baseline, measures, visitedMeasures, target);
                    break;
                case ShareMeasureExpression share:
                    CollectExpressionSliceIds(share.Part, measures, visitedMeasures, target);
                    CollectExpressionSliceIds(share.Whole, measures, visitedMeasures, target);
                    break;
            }
        }

        private static void CollectExpressionFilterFields(
            MeasureExpression expression,
            Dictionary<string, MeasureDefinition> measures,
            HashSet<string> visitedMeasures,
            HashSet<string> target)
        {
            switch (expression)
            {
                case FilteredAggregateMeasureExpression filtered:
                    foreach (var filter in filtered.Filters ?? new List<MeasureFilterSpec>())
                    {
                        if (filter != null && !string.IsNullOrWhiteSpace(filter.Field))
                        {
                            target.Add(filter.Field);
                        }
                    }

                    break;
                case WeightedAggregateMeasureExpression weighted:
                    CollectExpressionFilterFields(weighted.Numerator, measures, visitedMeasures, target);
                    CollectExpressionFilterFields(weighted.Denominator, measures, visitedMeasures, target);
                    break;
                case ReferenceMeasureExpression reference:
                    MeasureDefinition referenced;
                    if (measures.TryGetValue(reference.MeasureId, out referenced)
                        && visitedMeasures.Add(reference.MeasureId))
                    {
                        CollectExpressionFilterFields(referenced.Expression, measures, visitedMeasures, target);
                        visitedMeasures.Remove(reference.MeasureId);
                    }

                    break;
                case BinaryMeasureExpression binary:
                    CollectExpressionFilterFields(binary.Left, measures, visitedMeasures, target);
                    CollectExpressionFilterFields(binary.Right, measures, visitedMeasures, target);
                    break;
                case SafeDivideMeasureExpression divide:
                    CollectExpressionFilterFields(divide.Numerator, measures, visitedMeasures, target);
                    CollectExpressionFilterFields(divide.Denominator, measures, visitedMeasures, target);
                    break;
                case RatioMeasureExpression ratio:
                    CollectExpressionFilterFields(ratio.Numerator, measures, visitedMeasures, target);
                    CollectExpressionFilterFields(ratio.Denominator, measures, visitedMeasures, target);
                    break;
                case DifferenceMeasureExpression difference:
                    CollectExpressionFilterFields(difference.Current, measures, visitedMeasures, target);
                    CollectExpressionFilterFields(difference.Baseline, measures, visitedMeasures, target);
                    break;
                case ShareMeasureExpression share:
                    CollectExpressionFilterFields(share.Part, measures, visitedMeasures, target);
                    CollectExpressionFilterFields(share.Whole, measures, visitedMeasures, target);
                    break;
            }
        }

        private static void ValidateFieldPlacements(
            List<FilterPlacementSpec>? placements,
            HashSet<string>? fields,
            string path,
            ValidationResult result)
        {
            if (placements == null)
            {
                result.AddError("FILTER_PLACEMENTS_REQUIRED", path, "The Filters collection is required.");
                return;
            }

            var placed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var index = 0; index < placements.Count; index++)
            {
                var placement = placements[index];
                var itemPath = path + "[" + index.ToString(CultureInfo.InvariantCulture) + "]";
                if (placement == null)
                {
                    result.AddError("FILTER_PLACEMENT_REQUIRED", itemPath, "A Filter placement cannot be null.");
                    continue;
                }

                ValidateFieldReference(placement.Field, fields, itemPath + ".field", result);
                if (!placed.Add(placement.Field ?? string.Empty))
                {
                    result.AddError("FILTER_PLACEMENT_DUPLICATE", itemPath + ".field", "A field can appear only once in Filters.");
                }

                ValidateScalarListUnique(placement.SelectedValues, itemPath + ".selectedValues", result);
                if (placement.SelectedValues != null && placement.SelectedValues.Count > MaximumMembers)
                {
                    result.AddError("TOO_MANY_FILTER_VALUES", itemPath + ".selectedValues", "A report Filter may select at most 1,000 values.");
                }
                // An empty selection intentionally means "All". The field still
                // appears in the native PivotTable Filter area without limiting
                // the report. IncludeBlank is used only when a selection exists.
            }
        }

        private static void ValidateGroupBuckets(
            List<MemberGroupBucketSpec> buckets,
            string path,
            ValidationResult result)
        {
            if (buckets == null)
            {
                result.AddError("GROUP_BUCKETS_REQUIRED", path, "The group-bucket collection is required.");
                return;
            }

            if (buckets.Count > MaximumGroupBuckets)
            {
                result.AddError("TOO_MANY_GROUP_BUCKETS", path, "A field may contain at most 256 group buckets.");
            }

            var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var labels = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var members = new HashSet<string>(StringComparer.Ordinal);
            for (var index = 0; index < buckets.Count; index++)
            {
                var bucket = buckets[index];
                var itemPath = path + "[" + index.ToString(CultureInfo.InvariantCulture) + "]";
                if (bucket == null)
                {
                    result.AddError("GROUP_BUCKET_REQUIRED", itemPath, "A group bucket cannot be null.");
                    continue;
                }

                ValidateId(bucket.Id, itemPath + ".id", "GROUP_BUCKET_ID_INVALID", result);
                ValidateRequiredText(bucket.Label, itemPath + ".label", 120, result);
                if (!ids.Add(bucket.Id ?? string.Empty))
                {
                    result.AddError("GROUP_BUCKET_ID_DUPLICATE", itemPath + ".id", "Group-bucket IDs must be unique within a field.");
                }

                if (!labels.Add(bucket.Label ?? string.Empty))
                {
                    result.AddError("GROUP_BUCKET_LABEL_DUPLICATE", itemPath + ".label", "Group-bucket labels must be unique within a field.");
                }

                if (bucket.Members == null)
                {
                    result.AddError("GROUP_BUCKET_MEMBERS_REQUIRED", itemPath + ".members", "The group member collection is required.");
                    continue;
                }

                if (bucket.Members.Count == 0 && !bucket.IncludeUnmatched)
                {
                    result.AddError("GROUP_BUCKET_EMPTY", itemPath + ".members", "A group bucket requires members or IncludeUnmatched.");
                }

                if (bucket.Members.Count > MaximumMembers)
                {
                    result.AddError("TOO_MANY_GROUP_MEMBERS", itemPath + ".members", "A group bucket may contain at most 1,000 members.");
                }

                foreach (var member in bucket.Members)
                {
                    ValidateScalar(member, itemPath + ".members", result);
                    var key = ScalarKey(member);
                    if (!members.Add(key))
                    {
                        result.AddError("GROUP_MEMBER_DUPLICATE", itemPath + ".members", "A member cannot belong to more than one group bucket.");
                    }
                }
            }

            if (buckets.Count(bucket => bucket != null && bucket.IncludeUnmatched) > 1)
            {
                result.AddError("MULTIPLE_UNMATCHED_BUCKETS", path, "Only one group bucket may include unmatched members.");
            }
        }

        private static void ValidateSlices(List<PeriodSliceSpec> slices, string path, ValidationResult result)
        {
            if (slices == null)
            {
                result.AddError("PERIOD_SLICES_REQUIRED", path, "The period-slice collection is required.");
                return;
            }

            if (slices.Count > MaximumPeriodSlices)
            {
                result.AddError("TOO_MANY_PERIOD_SLICES", path, "A block may contain at most 64 period slices.");
            }

            var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var byId = new Dictionary<string, PeriodSliceSpec>(StringComparer.OrdinalIgnoreCase);
            for (var index = 0; index < slices.Count; index++)
            {
                var slice = slices[index];
                var slicePath = path + "[" + index.ToString(CultureInfo.InvariantCulture) + "]";
                if (slice == null)
                {
                    result.AddError("PERIOD_SLICE_REQUIRED", slicePath, "A period slice cannot be null.");
                    continue;
                }

                ValidateId(slice.Id, slicePath + ".id", "PERIOD_SLICE_ID_INVALID", result);
                ValidateOptionalId(slice.BasedOnSliceId, slicePath + ".basedOnSliceId", "PERIOD_SLICE_BASE_ID_INVALID", result);
                ValidateRequiredText(slice.Label, slicePath + ".label", 120, result);
                if (!ids.Add(slice.Id ?? string.Empty))
                {
                    result.AddError("PERIOD_SLICE_ID_DUPLICATE", slicePath + ".id", "Period-slice IDs must be unique within a block.");
                }
                else
                {
                    byId.Add(slice.Id ?? string.Empty, slice);
                }

                if (!Enum.IsDefined(typeof(PeriodSliceKind), slice.Kind))
                {
                    result.AddError("PERIOD_SLICE_KIND_INVALID", slicePath + ".kind", "The period-slice kind is not supported.");
                }
            }

            for (var index = 0; index < slices.Count; index++)
            {
                var slice = slices[index];
                if (slice == null)
                {
                    continue;
                }

                var slicePath = path + "[" + index.ToString(CultureInfo.InvariantCulture) + "]";
                var absolute = slice.Kind == PeriodSliceKind.Current || slice.Kind == PeriodSliceKind.Selected;
                if (absolute)
                {
                    if (!slice.SelectedStart.HasValue || !slice.SelectedEnd.HasValue)
                    {
                        result.AddError(
                            "ABSOLUTE_SLICE_DATES_REQUIRED",
                            slicePath,
                            "Current and selected period slices require explicit start and end dates.");
                    }
                    else if (slice.SelectedStart.Value.Date > slice.SelectedEnd.Value.Date)
                    {
                        result.AddError("PERIOD_SLICE_RANGE_INVALID", slicePath, "The period start must not be after its end.");
                    }
                    else if (slice.SelectedStart.Value.TimeOfDay != TimeSpan.Zero
                        || slice.SelectedEnd.Value.TimeOfDay != TimeSpan.Zero)
                    {
                        result.AddError("PERIOD_SLICE_TIME_NOT_ALLOWED", slicePath, "Period slice boundaries must be dates without time components.");
                    }
                }
                else if (slice.SelectedStart.HasValue || slice.SelectedEnd.HasValue)
                {
                    result.AddError("SLICE_DATES_NOT_ALLOWED", slicePath, "Relative period slices cannot contain explicit dates.");
                }

                if (slice.Kind == PeriodSliceKind.Prior || slice.Kind == PeriodSliceKind.SamePeriodPriorYear)
                {
                    if (string.IsNullOrWhiteSpace(slice.BasedOnSliceId) || !ids.Contains(slice.BasedOnSliceId!))
                    {
                        result.AddError("SLICE_BASE_UNKNOWN", slicePath + ".basedOnSliceId", "A relative period slice must reference another slice in the block.");
                    }
                    else if (string.Equals(slice.Id, slice.BasedOnSliceId, StringComparison.OrdinalIgnoreCase))
                    {
                        result.AddError("SLICE_SELF_REFERENCE", slicePath + ".basedOnSliceId", "A period slice cannot reference itself.");
                    }
                }
                else if (!string.IsNullOrWhiteSpace(slice.BasedOnSliceId))
                {
                    result.AddError("SLICE_BASE_NOT_ALLOWED", slicePath + ".basedOnSliceId", "Absolute period slices cannot reference another slice.");
                }
            }

            if (slices.Count(slice => slice != null && slice.Kind == PeriodSliceKind.Current) > 1)
            {
                result.AddError("MULTIPLE_CURRENT_SLICES", path, "A block can define at most one current period slice.");
            }

            foreach (var sliceId in byId.Keys)
            {
                DetectSliceCycle(sliceId, byId, new HashSet<string>(StringComparer.OrdinalIgnoreCase), result, path);
            }

            try
            {
                PeriodSliceResolver.Resolve(slices.Where(slice => slice != null));
            }
            catch (PeriodSliceResolutionException exception)
            {
                result.AddError(exception.Code, path, exception.Message);
            }
        }

        private static void DetectSliceCycle(
            string sliceId,
            Dictionary<string, PeriodSliceSpec> slices,
            HashSet<string> pathIds,
            ValidationResult result,
            string path)
        {
            if (!pathIds.Add(sliceId))
            {
                result.AddError("PERIOD_SLICE_REFERENCE_CYCLE", path, "The period-slice reference graph contains a cycle.");
                return;
            }

            PeriodSliceSpec slice;
            if (slices.TryGetValue(sliceId, out slice)
                && !string.IsNullOrWhiteSpace(slice.BasedOnSliceId)
                && slices.ContainsKey(slice.BasedOnSliceId!))
            {
                DetectSliceCycle(slice.BasedOnSliceId!, slices, pathIds, result, path);
            }

            pathIds.Remove(sliceId);
        }

        private static void ValidateHeaders(
            List<ReportHeaderSpec> headers,
            HashSet<string> styles,
            OwnedRangeExtentSpec? ownedExtent,
            string path,
            ValidationResult result)
        {
            if (headers == null)
            {
                result.AddError("HEADERS_REQUIRED", path, "The header collection is required.");
                return;
            }

            if (headers.Count > MaximumHeaders)
            {
                result.AddError("TOO_MANY_HEADERS", path, "A block may contain at most 128 headers.");
            }

            for (var index = 0; index < headers.Count; index++)
            {
                var header = headers[index];
                var itemPath = path + "[" + index.ToString(CultureInfo.InvariantCulture) + "]";
                if (header == null)
                {
                    result.AddError("HEADER_REQUIRED", itemPath, "A report header cannot be null.");
                    continue;
                }

                ValidateRequiredText(header.Text, itemPath + ".text", 240, result);
                if (header.RelativeRow < 0 || header.RelativeColumn < 0)
                {
                    result.AddError("HEADER_POSITION_INVALID", itemPath, "Header offsets cannot be negative.");
                }

                if (header.ColumnSpan < 1 || header.ColumnSpan > 16384)
                {
                    result.AddError("HEADER_SPAN_INVALID", itemPath + ".columnSpan", "Header column span is outside Excel limits.");
                }

                if (ownedExtent != null
                    && header.RelativeRow >= 0
                    && header.RelativeColumn >= 0
                    && header.ColumnSpan >= 1
                    && ((long)header.RelativeRow + 1L > ownedExtent.RowCount
                        || (long)header.RelativeColumn + header.ColumnSpan > ownedExtent.ColumnCount))
                {
                    result.AddError(
                        "HEADER_OUTSIDE_OWNED_EXTENT",
                        itemPath,
                        "The header must fit inside the block's owned write extent.");
                }

                ValidateStyleReference(header.StyleId, styles, itemPath + ".styleId", result);
            }


            for (var leftIndex = 0; leftIndex < headers.Count; leftIndex++)
            {
                var left = headers[leftIndex];
                if (left == null || left.RelativeRow < 0 || left.RelativeColumn < 0 || left.ColumnSpan < 1)
                {
                    continue;
                }

                for (var rightIndex = leftIndex + 1; rightIndex < headers.Count; rightIndex++)
                {
                    var right = headers[rightIndex];
                    if (right == null || right.RelativeRow != left.RelativeRow
                        || right.RelativeColumn < 0 || right.ColumnSpan < 1)
                    {
                        continue;
                    }

                    var leftEnd = (long)left.RelativeColumn + left.ColumnSpan - 1L;
                    var rightEnd = (long)right.RelativeColumn + right.ColumnSpan - 1L;
                    if (left.RelativeColumn <= rightEnd && right.RelativeColumn <= leftEnd)
                    {
                        result.AddError(
                            "HEADER_OVERLAP",
                            path + "[" + rightIndex.ToString(CultureInfo.InvariantCulture) + "]",
                            "Report headers cannot occupy overlapping cells.");
                    }
                }
            }
        }

        private static void ValidateSpacers(List<SpacerSpec> spacers, string path, ValidationResult result)
        {
            if (spacers == null)
            {
                result.AddError("SPACERS_REQUIRED", path, "The spacer collection is required.");
                return;
            }

            if (spacers.Count > MaximumSpacers)
            {
                result.AddError("TOO_MANY_SPACERS", path, "A block may contain at most 64 spacers.");
            }

            for (var index = 0; index < spacers.Count; index++)
            {
                var spacer = spacers[index];
                var itemPath = path + "[" + index.ToString(CultureInfo.InvariantCulture) + "]";
                if (spacer == null)
                {
                    result.AddError("SPACER_REQUIRED", itemPath, "A spacer cannot be null.");
                    continue;
                }

                ValidateEnum(spacer.Axis, itemPath + ".axis", "SPACER_AXIS_INVALID", result);
                if (spacer.BeforeLevel < 0 || spacer.Count < 1 || spacer.Count > 100)
                {
                    result.AddError("SPACER_INVALID", itemPath, "Spacer level and count must be within their bounded ranges.");
                }

                if (spacer.Size.HasValue && (spacer.Size.Value <= 0d || spacer.Size.Value > 409d))
                {
                    result.AddError("SPACER_SIZE_INVALID", itemPath + ".size", "Spacer size must be greater than zero and within Excel limits.");
                }
            }
        }

        private static void ValidateChecks(
            List<ReportCheckSpec> checks,
            Dictionary<string, MeasureDefinition> measures,
            ValidationResult result)
        {
            if (checks == null)
            {
                result.AddError("CHECKS_REQUIRED", "$.checks", "The Checks collection is required.");
                return;
            }

            if (checks.Count > MaximumChecks)
            {
                result.AddError("TOO_MANY_CHECKS", "$.checks", "A report may contain at most 128 Checks.");
            }

            var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var index = 0; index < checks.Count; index++)
            {
                var check = checks[index];
                var path = "$.checks[" + index.ToString(CultureInfo.InvariantCulture) + "]";
                if (check == null)
                {
                    result.AddError("CHECK_REQUIRED", path, "A report check cannot be null.");
                    continue;
                }

                ValidateId(check.Id, path + ".id", "CHECK_ID_INVALID", result);
                ValidateEnum(check.Kind, path + ".kind", "CHECK_KIND_INVALID", result);
                if (!ids.Add(check.Id ?? string.Empty))
                {
                    result.AddError("CHECK_ID_DUPLICATE", path + ".id", "Check IDs must be unique.");
                }

                if (check.Tolerance < 0m)
                {
                    result.AddError("CHECK_TOLERANCE_INVALID", path + ".tolerance", "Check tolerance cannot be negative.");
                }

                ValidateOptionalId(check.MeasureId, path + ".measureId", "CHECK_MEASURE_ID_INVALID", result);
                if (!string.IsNullOrWhiteSpace(check.MeasureId) && !measures.ContainsKey(check.MeasureId!))
                {
                    result.AddError("CHECK_MEASURE_UNKNOWN", path + ".measureId", "The check references an unknown measure.");
                }
                else
                {
                    ValidateCheckMeasureScope(check.MeasureId, measures, path + ".measureId", result);
                }

                ValidateOptionalId(check.ComparedMeasureId, path + ".comparedMeasureId", "CHECK_COMPARISON_ID_INVALID", result);
                if (!string.IsNullOrWhiteSpace(check.ComparedMeasureId) && !measures.ContainsKey(check.ComparedMeasureId!))
                {
                    result.AddError("CHECK_COMPARISON_UNKNOWN", path + ".comparedMeasureId", "The check references an unknown comparison measure.");
                }
                else
                {
                    ValidateCheckMeasureScope(check.ComparedMeasureId, measures, path + ".comparedMeasureId", result);
                }

                if (check.Kind == ReportCheckKind.Balance
                    && (string.IsNullOrWhiteSpace(check.MeasureId) || string.IsNullOrWhiteSpace(check.ComparedMeasureId)))
                {
                    result.AddError("BALANCE_MEASURES_REQUIRED", path, "A balance check requires two measures.");
                }

                if ((check.Kind == ReportCheckKind.TotalPreservation
                        || check.Kind == ReportCheckKind.RequiredValues
                        || check.Kind == ReportCheckKind.NonNegative)
                    && string.IsNullOrWhiteSpace(check.MeasureId))
                {
                    result.AddError("CHECK_MEASURE_REQUIRED", path + ".measureId", "This check requires a measure.");
                }

                if (check.Kind == ReportCheckKind.NoTruncation
                    && (!string.IsNullOrWhiteSpace(check.MeasureId)
                        || !string.IsNullOrWhiteSpace(check.ComparedMeasureId)))
                {
                    result.AddError("CHECK_MEASURE_NOT_ALLOWED", path, "A no-truncation check does not accept measure references.");
                }
            }
        }

        private static void ValidateCheckMeasureScope(
            string? measureId,
            Dictionary<string, MeasureDefinition> measures,
            string path,
            ValidationResult result)
        {
            MeasureDefinition measure;
            if (string.IsNullOrWhiteSpace(measureId) || !measures.TryGetValue(measureId!, out measure))
            {
                return;
            }

            var slices = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            CollectExpressionSliceIds(
                measure.Expression,
                measures,
                new HashSet<string>(StringComparer.OrdinalIgnoreCase),
                slices);
            if (slices.Count != 0)
            {
                result.AddError(
                    "CHECK_PERIOD_SCOPE_AMBIGUOUS",
                    path,
                    "A global Check cannot reference a slice-bound measure. Use an unbound measure or a block-scoped output check.");
            }
        }

        private static void ValidateScalar(ScalarValue value, string path, ValidationResult result)
        {
            if (value == null)
            {
                result.AddError("SCALAR_REQUIRED", path, "A literal value is required.");
                return;
            }

            ValidateEnum(value.Kind, path + ".kind", "SCALAR_KIND_INVALID", result);

            var populated = (value.Text != null ? 1 : 0)
                + (value.Number.HasValue ? 1 : 0)
                + (value.Boolean.HasValue ? 1 : 0)
                + (value.Temporal.HasValue ? 1 : 0);
            var valid = value.Kind == ScalarValueKind.Null && populated == 0
                || value.Kind == ScalarValueKind.Text && value.Text != null && populated == 1
                || value.Kind == ScalarValueKind.Number && value.Number.HasValue && populated == 1
                || value.Kind == ScalarValueKind.Boolean && value.Boolean.HasValue && populated == 1
                || (value.Kind == ScalarValueKind.Date || value.Kind == ScalarValueKind.DateTime)
                    && value.Temporal.HasValue && populated == 1;
            if (!valid)
            {
                result.AddError("SCALAR_SHAPE_INVALID", path, "The literal payload does not match its declared kind.");
            }

            if (value.Kind == ScalarValueKind.Date && value.Temporal.HasValue
                && value.Temporal.Value.TimeOfDay != TimeSpan.Zero)
            {
                result.AddError("DATE_LITERAL_HAS_TIME", path + ".temporal", "A date literal cannot contain a time component.");
            }

            if (value.Kind == ScalarValueKind.Text
                && value.Text != null
                && (value.Text.Length > 1024 || value.Text.Any(char.IsControl)))
            {
                result.AddError(
                    "SCALAR_TEXT_INVALID",
                    path + ".text",
                    "Text literals may contain at most 1,024 characters and cannot contain control characters.");
            }
        }

        private static void ValidateScalarListUnique(List<ScalarValue> values, string path, ValidationResult result)
        {
            if (values == null)
            {
                result.AddError("SCALAR_LIST_REQUIRED", path, "The value list is required.");
                return;
            }

            var keys = new HashSet<string>(StringComparer.Ordinal);
            foreach (var value in values)
            {
                ValidateScalar(value, path, result);
                if (!keys.Add(ScalarKey(value)))
                {
                    result.AddError("MEMBER_ORDER_DUPLICATE", path, "Member order cannot contain duplicate values.");
                }
            }
        }

        private static string ScalarKey(ScalarValue value)
        {
            if (value == null)
            {
                return "<null>";
            }

            return value.Kind + "|" + (value.Text
                ?? (value.Number.HasValue ? value.Number.Value.ToString(CultureInfo.InvariantCulture) : null)
                ?? (value.Boolean.HasValue ? value.Boolean.Value.ToString() : null)
                ?? (value.Temporal.HasValue ? value.Temporal.Value.ToString("O", CultureInfo.InvariantCulture) : string.Empty));
        }

        private static void ValidateColumnOperation(
            string? column,
            HashSet<string>? fields,
            string path,
            ValidationResult result)
        {
            ValidateColumnName(column, path, result);
            ValidateExistingColumn(column, fields, path, result);
        }

        private static void ValidateColumnOperationList(
            List<string> columns,
            HashSet<string>? fields,
            string path,
            ValidationResult result)
        {
            ValidateColumnList(columns, path, true, result);
            ValidateExistingColumns(columns, fields, path, result);
        }

        private static void ValidateColumnList(
            IEnumerable<string> columns,
            string path,
            bool requireOne,
            ValidationResult result)
        {
            if (columns == null)
            {
                result.AddError("COLUMN_LIST_REQUIRED", path, "The column list is required.");
                return;
            }

            var materialized = columns.ToList();
            if (materialized.Count > MaximumPeriodColumns)
            {
                result.AddError("TOO_MANY_COLUMNS", path, "A column list may contain at most 16,384 entries.");
            }
            if (requireOne && materialized.Count == 0)
            {
                result.AddError("COLUMN_LIST_EMPTY", path, "At least one column is required.");
            }

            for (var index = 0; index < materialized.Count; index++)
            {
                ValidateColumnName(materialized[index], path + "[" + index.ToString(CultureInfo.InvariantCulture) + "]", result);
            }

            ValidateDistinct(materialized, path, "COLUMN_LIST_DUPLICATE", result);
        }

        private static void ValidateColumnName(string? value, string path, ValidationResult result)
        {
            ValidateRequiredText(value, path, 255, result);
            if (!string.IsNullOrEmpty(value) && value!.StartsWith("__erb_", StringComparison.OrdinalIgnoreCase))
            {
                result.AddError("RESERVED_COLUMN_NAME", path, "Column names beginning with __erb_ are reserved for deterministic build steps.");
            }
        }

        private static void ValidateFieldReference(
            string? field,
            HashSet<string>? fields,
            string path,
            ValidationResult result)
        {
            ValidateColumnName(field, path, result);
            ValidateExistingColumn(field, fields, path, result);
        }

        private static void ValidateExistingColumns(
            IEnumerable<string> columns,
            HashSet<string>? fields,
            string path,
            ValidationResult result)
        {
            if (columns == null || fields == null)
            {
                return;
            }

            foreach (var column in columns)
            {
                ValidateExistingColumn(column, fields, path, result);
            }
        }

        private static void ValidateExistingColumn(
            string? column,
            HashSet<string>? fields,
            string path,
            ValidationResult result)
        {
            if (fields == null || string.IsNullOrWhiteSpace(column))
            {
                return;
            }

            if (fields.Any(field => string.Equals(field, column, StringComparison.Ordinal)))
            {
                return;
            }

            if (fields.Contains(column!))
            {
                result.AddError(
                    "SOURCE_FIELD_CASE_MISMATCH",
                    path,
                    "The column reference '" + column + "' does not match the source column's letter casing.");
            }
            else
            {
                result.AddError("SOURCE_FIELD_UNKNOWN", path, "The column '" + column + "' is not available at this step.");
            }
        }

        private static void ValidateProfileColumn(
            string? column,
            SourceProfile? sourceProfile,
            string path,
            ValidationResult result)
        {
            if (sourceProfile == null || string.IsNullOrWhiteSpace(column))
            {
                return;
            }

            if (sourceProfile.Columns.Any(profileColumn =>
                    string.Equals(profileColumn.Name, column, StringComparison.Ordinal)))
            {
                return;
            }

            if (sourceProfile.FindColumn(column!) != null)
            {
                result.AddError(
                    "SOURCE_FIELD_CASE_MISMATCH",
                    path,
                    "The column reference '" + column + "' does not match the source column's letter casing.");
            }
            else
            {
                result.AddError("SOURCE_FIELD_UNKNOWN", path, "The source does not contain column '" + column + "'.");
            }
        }

        private static void ValidateRequiredText(
            string? value,
            string path,
            int maximumLength,
            ValidationResult result)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                result.AddError("TEXT_REQUIRED", path, "A non-blank value is required.");
            }
            else if (value!.Length > maximumLength)
            {
                result.AddError("TEXT_TOO_LONG", path, "The value exceeds the bounded length.");
            }
            else if (value.Any(char.IsControl))
            {
                result.AddError("CONTROL_CHARACTER_NOT_ALLOWED", path, "Control characters are not allowed.");
            }
        }

        private static void ValidateId(string? value, string path, string code, ValidationResult result)
        {
            if (string.IsNullOrWhiteSpace(value) || !IdPattern.IsMatch(value))
            {
                result.AddError(code, path, "Use 1-64 letters, numbers, underscores, or hyphens, beginning with a letter.");
            }
        }

        private static void ValidateEnum<TEnum>(
            TEnum value,
            string path,
            string code,
            ValidationResult result)
            where TEnum : struct
        {
            if (!Enum.IsDefined(typeof(TEnum), value))
            {
                result.AddError(code, path, "The value is not supported by this report specification version.");
            }
        }

        private static void ValidateOptionalId(string? value, string path, string code, ValidationResult result)
        {
            if (value != null && !IdPattern.IsMatch(value))
            {
                result.AddError(code, path, "Use 1-64 letters, numbers, underscores, or hyphens, beginning with a letter.");
            }
        }

        private static void ValidateOptionalText(
            string? value,
            string path,
            int maximumLength,
            ValidationResult result)
        {
            if (value != null)
            {
                ValidateRequiredText(value, path, maximumLength, result);
            }
        }

        private static void ValidateDistinct(
            IEnumerable<string> values,
            string path,
            string code,
            ValidationResult result)
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var value in values)
            {
                if (!seen.Add(value ?? string.Empty))
                {
                    result.AddError(code, path, "Values in this collection must be unique.");
                    return;
                }
            }
        }

        private static void ValidateStyleReference(
            string? styleId,
            HashSet<string> styles,
            string path,
            ValidationResult result)
        {
            if (styleId == null)
            {
                return;
            }

            if (!IdPattern.IsMatch(styleId))
            {
                result.AddError("STYLE_REFERENCE_INVALID", path, "A present style reference must be a valid nonblank identifier.");
            }
            else if (!styles.Contains(styleId))
            {
                result.AddError("STYLE_REFERENCE_UNKNOWN", path, "The referenced presentation style does not exist.");
            }
        }

        private static void ValidateColor(string? color, string path, ValidationResult result)
        {
            if (color != null && !ColorPattern.IsMatch(color))
            {
                result.AddError("COLOR_INVALID", path, "Colors must use #RRGGBB notation.");
            }
        }

        private static void ValidateNumberFormat(string? format, string path, ValidationResult result)
        {
            if (format == null)
            {
                return;
            }

            if (format.Length > 128 || format.Any(char.IsControl) || format.StartsWith("=", StringComparison.Ordinal))
            {
                result.AddError("NUMBER_FORMAT_INVALID", path, "The number format is not a bounded Excel number-format string.");
            }
        }

        private static void ValidateWorksheetName(string? name, string path, ValidationResult result)
        {
            ValidateRequiredText(name, path, 31, result);
            if (!string.IsNullOrEmpty(name) && name!.IndexOfAny(new[] { '[', ']', ':', '*', '?', '/', '\\' }) >= 0)
            {
                result.AddError("WORKSHEET_NAME_INVALID", path, "The worksheet name contains a character Excel does not allow.");
            }
        }
    }
}
