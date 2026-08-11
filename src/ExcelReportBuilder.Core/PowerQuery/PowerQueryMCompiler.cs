using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Core.Transforms;

namespace ExcelReportBuilder.Core.PowerQuery
{
    public sealed class MCompilationResult
    {
        public string Query { get; set; } = string.Empty;

        public string FinalStepName { get; set; } = string.Empty;

        public string SourceConnector { get; set; } = "Excel.CurrentWorkbook";

        public List<string> ReferencedWorkbookObjects { get; set; } = new List<string>();

        public List<string> AppliedTransformIds { get; set; } = new List<string>();
    }

    public sealed class MCompilationException : Exception
    {
        public MCompilationException(string code, string message, string? transformId = null)
            : base(message)
        {
            Code = code;
            TransformId = transformId;
        }

        public string Code { get; }

        public string? TransformId { get; }
    }

    /// <summary>
    /// Compiles the closed transform union to Power Query M. The source grammar
    /// has no connector choice: every query begins at Excel.CurrentWorkbook.
    /// </summary>
    public static class PowerQueryMCompiler
    {
        private static readonly Regex WorkbookObjectPattern = new Regex(
            @"^[A-Za-z_\\][A-Za-z0-9_.\\]*$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly string[] BoundedDateFormats =
        {
            "yyyy-MM-dd",
            "yyyy/MM/dd",
            "M/d/yyyy",
            "d-MMM-yyyy",
            "MMM d yyyy",
            "MMMM d yyyy"
        };

        private static readonly string[] BoundedDateTimeFormats =
        {
            "yyyy-MM-dd HH:mm:ss",
            "yyyy-MM-ddTHH:mm:ss",
            "yyyy-MM-ddTHH:mm:ss.FFFFFFF",
            "M/d/yyyy h:mm:ss tt"
        };

        public static MCompilationResult Compile(ReportSpecV1 specification)
        {
            if (specification == null)
            {
                throw new ArgumentNullException(nameof(specification));
            }

            return Compile(specification.Source, specification.Transforms, specification.PeriodMapping);
        }

        public static MCompilationResult Compile(
            WorkbookSourceSpec source,
            IEnumerable<TransformStep> transforms,
            PeriodMappingSpec? periodMapping = null)
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            if (transforms == null)
            {
                throw new ArgumentNullException(nameof(transforms));
            }

            if (string.IsNullOrWhiteSpace(source.WorkbookObjectName)
                || !WorkbookObjectPattern.IsMatch(source.WorkbookObjectName))
            {
                throw new MCompilationException(
                    "SOURCE_NAME_INVALID",
                    "Only a valid Excel table or named-range identifier can be read through Excel.CurrentWorkbook.");
            }

            if (source.HeaderRowCount != 1)
            {
                throw new MCompilationException("ONE_HEADER_ROW_REQUIRED", "Power Query compilation requires exactly one header row.");
            }

            var materialized = transforms.ToList();
            if (materialized.Count > 100)
            {
                throw new MCompilationException("TOO_MANY_TRANSFORMS", "At most 100 bounded transforms can be compiled.");
            }

            var workbookSource = "Excel.CurrentWorkbook(){[Name=" +
                MString(source.WorkbookObjectName) + "]}[Content]";
            var lines = new List<string>();
            switch (source.Kind)
            {
                case WorkbookSourceKind.Table:
                    lines.Add("    Source = " + workbookSource);
                    break;
                case WorkbookSourceKind.NamedRange:
                    lines.Add("    RawSource = " + workbookSource);
                    lines.Add(
                        "    Source = Table.PromoteHeaders(RawSource, " +
                        "[PromoteAllScalars = true, Culture = \"en-US\"])");
                    break;
                default:
                    throw new MCompilationException(
                        "SOURCE_KIND_INVALID",
                        "Only Excel tables and managed named ranges can be compiled.");
            }
            var appliedIds = new List<string>();
            var current = "Source";
            for (var index = 0; index < materialized.Count; index++)
            {
                var transform = materialized[index];
                if (transform == null)
                {
                    throw new MCompilationException("TRANSFORM_REQUIRED", "A transform cannot be null.");
                }

                var stepName = "Step " + (index + 1).ToString("00", CultureInfo.InvariantCulture)
                    + " " + transform.Kind;
                var expression = CompileTransform(current, transform, periodMapping);
                lines.Add("    " + MIdentifier(stepName) + " = " + expression);
                current = MIdentifier(stepName);
                appliedIds.Add(transform.Id);
            }

            var query = new StringBuilder();
            query.AppendLine("let");
            for (var index = 0; index < lines.Count; index++)
            {
                query.Append(lines[index]);
                query.AppendLine(index == lines.Count - 1 ? string.Empty : ",");
            }

            query.Append("in");
            query.AppendLine();
            query.Append("    ");
            query.Append(current);

            return new MCompilationResult
            {
                Query = query.ToString(),
                FinalStepName = current,
                SourceConnector = "Excel.CurrentWorkbook",
                ReferencedWorkbookObjects = new List<string> { source.WorkbookObjectName },
                AppliedTransformIds = appliedIds
            };
        }

        private static string CompileTransform(
            string previous,
            TransformStep transform,
            PeriodMappingSpec? rootMapping)
        {
            switch (transform)
            {
                case SelectColumnsTransform select:
                    return "Table.SelectColumns(" + previous + ", " + MStringList(RequireColumns(select.Columns, transform)) + ", MissingField.Error)";
                case KeepColumnsTransform keep:
                    return "Table.SelectColumns(" + previous + ", " + MStringList(RequireColumns(keep.Columns, transform)) + ", MissingField.Error)";
                case RemoveColumnsTransform remove:
                    return "Table.RemoveColumns(" + previous + ", " + MStringList(RequireColumns(remove.Columns, transform)) + ", MissingField.Error)";
                case ReorderColumnsTransform reorder:
                    return "Table.ReorderColumns(" + previous + ", " + MStringList(RequireColumns(reorder.Columns, transform)) + ", MissingField.Error)";
                case RenameColumnTransform rename:
                    return "Table.RenameColumns(" + previous + ", {{"
                        + MString(RequireColumn(rename.From, transform)) + ", "
                        + MString(RequireColumn(rename.To, transform)) + "}}, MissingField.Error)";
                case ChangeColumnTypeTransform changeType:
                    return CompileChangeColumnType(previous, changeType);
                case TrimTextTransform trim:
                    return CompileTrim(previous, trim);
                case ReplaceValueTransform replace:
                    return "Table.ReplaceValue(" + previous + ", " + MLiteral(replace.Find, transform)
                        + ", " + MLiteral(replace.ReplaceWith, transform)
                        + ", Replacer.ReplaceValue, {" + MString(RequireColumn(replace.Column, transform)) + "})";
                case NormalizeBlanksTransform blanks:
                    return CompileNormalizeBlanks(previous, blanks);
                case NormalizeErrorsTransform errors:
                    return CompileNormalizeErrors(previous, errors);
                case FillDownTransform fillDown:
                    return "Table.FillDown(" + previous + ", " + MStringList(RequireColumns(fillDown.Columns, transform)) + ")";
                case MapValuesTransform map:
                    return CompileMap(previous, map);
                case FilterRowsTransform filter:
                    return "Table.SelectRows(" + previous + ", each " + CompileFilter(filter) + ")";
                case ExcludeTotalRowsTransform exclude:
                    return CompileTotalRowExclusion(previous, exclude);
                case DerivePeriodPartsTransform derive:
                    return CompilePeriodParts(previous, derive);
                case AddArithmeticColumnTransform arithmetic:
                    return CompileArithmetic(previous, arithmetic);
                case NormalizePeriodsTransform normalize:
                    return CompilePeriodNormalization(previous, normalize, rootMapping);
                default:
                    throw new MCompilationException(
                        "TRANSFORM_KIND_UNSUPPORTED",
                        "The transform kind cannot be compiled.",
                        transform.Id);
            }
        }

        private static string CompileTrim(string previous, TrimTextTransform transform)
        {
            var columns = RequireColumns(transform.Columns, transform);
            var operations = columns.Select(column => "{" + MString(column)
                + ", each if _ is null then null else Text.Trim(Text.From(_, \"en-US\")), type nullable text}");
            return "Table.TransformColumns(" + previous + ", {" + string.Join(", ", operations) + "})";
        }

        private static string CompileChangeColumnType(
            string previous,
            ChangeColumnTypeTransform transform)
        {
            var column = RequireColumn(transform.Column, transform);
            string conversion;
            string resultType;
            switch (transform.DataType)
            {
                case ColumnDataType.Text:
                    conversion = "if _ is null then null"
                        + " else if _ is text or _ is logical or _ is number or _ is date or _ is datetime"
                        + " then Text.From(_, \"en-US\")"
                        + " else error Error.Record(\"Invalid text value\", \"Text conversion accepts text, logical, finite or non-finite numbers, dates, date-times, or blank.\", null)";
                    resultType = "type nullable text";
                    break;
                case ColumnDataType.WholeNumber:
                    conversion = "let converted = " + CompileBoundedDecimalValue("_", "Type conversion")
                        + " in if converted is null then null else Int64.From(converted, \"en-US\")";
                    resultType = "Int64.Type";
                    break;
                case ColumnDataType.DecimalNumber:
                    conversion = CompileBoundedDecimalValue("_", "Type conversion");
                    resultType = "type nullable number";
                    break;
                case ColumnDataType.Boolean:
                    conversion = "if _ is null then null else if _ is logical then _"
                        + " else if _ is number then " + CompileBoundedDecimalValue("_", "Logical conversion") + " <> 0"
                        + " else if _ is text then let text = Text.Lower(Text.Trim(_)) in"
                        + " if text = \"true\" then true else if text = \"false\" then false"
                        + " else error Error.Record(\"Invalid logical value\", \"Logical text must be true or false.\", null)"
                        + " else error Error.Record(\"Invalid logical value\", \"Logical values must be logical, numeric, true, false, or blank.\", null)";
                    resultType = "type nullable logical";
                    break;
                case ColumnDataType.Date:
                    conversion = CompileBoundedTemporalValue("_", dateTime: false);
                    resultType = "type nullable date";
                    break;
                case ColumnDataType.DateTime:
                    conversion = CompileBoundedTemporalValue("_", dateTime: true);
                    resultType = "type nullable datetime";
                    break;
                default:
                    throw new MCompilationException(
                        "COLUMN_DATA_TYPE_INVALID",
                        "The column data type cannot be compiled.",
                        transform.Id);
            }

            return "Table.TransformColumns(" + previous + ", {{" + MString(column)
                + ", each " + conversion + ", " + resultType
                + "}}, null, MissingField.Error)";
        }

        private static string CompileNormalizeBlanks(string previous, NormalizeBlanksTransform transform)
        {
            var replacement = MLiteral(transform.Replacement, transform);
            var predicate = transform.TreatWhitespaceAsBlank
                ? "Value.Is(_, type text) and Text.Trim(Text.From(_, \"en-US\")) = \"\""
                : "Value.Is(_, type text) and Text.From(_, \"en-US\") = \"\"";
            var operations = RequireColumns(transform.Columns, transform)
                .Select(column => "{" + MString(column) + ", each if _ is null then " + replacement
                    + " else if " + predicate + " then " + replacement + " else _, type any}");
            return "Table.TransformColumns(" + previous + ", {" + string.Join(", ", operations) + "})";
        }

        private static string CompileNormalizeErrors(string previous, NormalizeErrorsTransform transform)
        {
            var replacement = MLiteral(transform.Replacement, transform);
            var operations = RequireColumns(transform.Columns, transform)
                .Select(column => "{" + MString(column) + ", " + replacement + "}");
            return "Table.ReplaceErrorValues(" + previous + ", {" + string.Join(", ", operations) + "})";
        }

        private static string CompileMap(string previous, MapValuesTransform transform)
        {
            var column = RequireColumn(transform.Column, transform);
            if (transform.Entries == null || transform.Entries.Count == 0 || transform.Entries.Count > 256)
            {
                throw new MCompilationException("MAP_ENTRIES_INVALID", "A value map requires 1-256 entries.", transform.Id);
            }

            var inputs = new HashSet<string>(StringComparer.Ordinal);
            var clauses = new List<string>();
            foreach (var entry in transform.Entries)
            {
                if (entry == null)
                {
                    throw new MCompilationException("MAP_ENTRY_REQUIRED", "A value-map entry cannot be null.", transform.Id);
                }

                var input = MLiteral(entry.From, transform);
                if (!inputs.Add(input))
                {
                    throw new MCompilationException("MAP_INPUT_DUPLICATE", "Each input literal can appear only once in a value map.", transform.Id);
                }

                clauses.Add("if Value.Equals(_, " + input + ") then " + MLiteral(entry.To, transform));
            }

            return "Table.TransformColumns(" + previous + ", {{" + MString(column)
                + ", each " + string.Join(" else ", clauses) + " else _, type any}})";
        }

        private static string CompileFilter(FilterRowsTransform transform)
        {
            var field = MField(RequireColumn(transform.Column, transform));
            switch (transform.Operator)
            {
                case RowFilterOperator.IsBlank:
                    RequireNoFilterValue(transform);
                    return field + " is null or (Value.Is(" + field + ", type text) and Text.Trim(Text.From(" + field + ", \"en-US\")) = \"\")";
                case RowFilterOperator.IsNotBlank:
                    RequireNoFilterValue(transform);
                    return "not (" + field + " is null or (Value.Is(" + field + ", type text) and Text.Trim(Text.From(" + field + ", \"en-US\")) = \"\"))";
            }

            if (transform.Value == null)
            {
                throw new MCompilationException("FILTER_VALUE_REQUIRED", "The filter operator requires a literal value.", transform.Id);
            }

            var literal = MLiteral(transform.Value, transform);
            switch (transform.Operator)
            {
                case RowFilterOperator.Equal:
                    return field + " = " + literal;
                case RowFilterOperator.NotEqual:
                    return field + " <> " + literal;
                case RowFilterOperator.GreaterThan:
                    return field + " <> null and " + literal + " <> null and " + field + " > " + literal;
                case RowFilterOperator.GreaterThanOrEqual:
                    return field + " <> null and " + literal + " <> null and " + field + " >= " + literal;
                case RowFilterOperator.LessThan:
                    return field + " <> null and " + literal + " <> null and " + field + " < " + literal;
                case RowFilterOperator.LessThanOrEqual:
                    return field + " <> null and " + literal + " <> null and " + field + " <= " + literal;
                case RowFilterOperator.Contains:
                    RequireTextLiteral(transform.Value, transform);
                    return field + " <> null and Text.Contains(Text.From(" + field + ", \"en-US\"), " + literal + ", Comparer.Ordinal)";
                case RowFilterOperator.StartsWith:
                    RequireTextLiteral(transform.Value, transform);
                    return field + " <> null and Text.StartsWith(Text.From(" + field + ", \"en-US\"), " + literal + ", Comparer.Ordinal)";
                case RowFilterOperator.EndsWith:
                    RequireTextLiteral(transform.Value, transform);
                    return field + " <> null and Text.EndsWith(Text.From(" + field + ", \"en-US\"), " + literal + ", Comparer.Ordinal)";
                default:
                    throw new MCompilationException("FILTER_OPERATOR_UNSUPPORTED", "The filter operator cannot be compiled.", transform.Id);
            }
        }

        private static string CompileTotalRowExclusion(string previous, ExcludeTotalRowsTransform transform)
        {
            if (transform.Evidence == null || transform.Evidence.Count == 0)
            {
                throw new MCompilationException(
                    "TOTAL_ROW_EVIDENCE_REQUIRED",
                    "Total rows cannot be excluded without explicit observed evidence.",
                    transform.Id);
            }

            var conditions = new List<string>();
            foreach (var evidence in transform.Evidence)
            {
                if (evidence.ObservedMatchCount <= 0)
                {
                    throw new MCompilationException(
                        "TOTAL_ROW_MATCH_COUNT_REQUIRED",
                        "Total-row evidence must include a positive observed match count.",
                        transform.Id);
                }

                var field = MField(RequireColumn(evidence.Column, transform));
                if (evidence.MatchKind == TotalRowMatchKind.IsBlank)
                {
                    conditions.Add("(" + field + " is null or (Value.Is(" + field
                        + ", type text) and Text.Trim(Text.From(" + field + ", \"en-US\")) = \"\"))");
                    continue;
                }

                if (evidence.Values == null || evidence.Values.Count == 0)
                {
                    throw new MCompilationException(
                        "TOTAL_ROW_EVIDENCE_VALUES_REQUIRED",
                        "The total-row evidence condition requires values.",
                        transform.Id);
                }

                var valueConditions = new List<string>();
                foreach (var value in evidence.Values)
                {
                    var literal = MLiteral(value, transform);
                    switch (evidence.MatchKind)
                    {
                        case TotalRowMatchKind.EqualsAny:
                            valueConditions.Add(field + " = " + literal);
                            break;
                        case TotalRowMatchKind.StartsWith:
                            RequireTextLiteral(value, transform);
                            valueConditions.Add(field + " <> null and Text.StartsWith(Text.From(" + field + ", \"en-US\"), "
                                + literal + ", Comparer.Ordinal)");
                            break;
                        case TotalRowMatchKind.Contains:
                            RequireTextLiteral(value, transform);
                            valueConditions.Add(field + " <> null and Text.Contains(Text.From(" + field + ", \"en-US\"), "
                                + literal + ", Comparer.Ordinal)");
                            break;
                        default:
                            throw new MCompilationException(
                                "TOTAL_ROW_MATCH_UNSUPPORTED",
                                "The total-row evidence match cannot be compiled.",
                                transform.Id);
                    }
                }

                conditions.Add("(" + string.Join(" or ", valueConditions) + ")");
            }

            var join = transform.RequireAllEvidence ? " and " : " or ";
            return "Table.SelectRows(" + previous + ", each not (" + string.Join(join, conditions) + "))";
        }

        private static string CompilePeriodParts(string previous, DerivePeriodPartsTransform transform)
        {
            var dateField = MField(RequireColumn(transform.DateColumn, transform));
            if (transform.Columns == null || transform.Columns.Count == 0)
            {
                throw new MCompilationException("DERIVED_PERIOD_COLUMNS_REQUIRED", "At least one period part is required.", transform.Id);
            }

            var expression = previous;
            var outputs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var column in transform.Columns)
            {
                var output = RequireColumn(column.OutputColumn, transform);
                if (!outputs.Add(output))
                {
                    throw new MCompilationException("DERIVED_PERIOD_COLUMN_DUPLICATE", "Derived output columns must be unique.", transform.Id);
                }

                string valueExpression;
                string valueType;
                switch (column.Part)
                {
                    case DerivedPeriodPart.Year:
                        valueExpression = "Date.Year(Date.From(" + dateField + ", \"en-US\"))";
                        valueType = "Int64.Type";
                        break;
                    case DerivedPeriodPart.Half:
                        valueExpression = "if Date.Month(Date.From(" + dateField + ", \"en-US\")) <= 6 then \"H1\" else \"H2\"";
                        valueType = "type text";
                        break;
                    case DerivedPeriodPart.Quarter:
                        valueExpression = "\"Q\" & Number.ToText(Date.QuarterOfYear(Date.From(" + dateField + ", \"en-US\")), \"0\", \"en-US\")";
                        valueType = "type text";
                        break;
                    case DerivedPeriodPart.MonthNumber:
                        valueExpression = "Date.Month(Date.From(" + dateField + ", \"en-US\"))";
                        valueType = "Int64.Type";
                        break;
                    case DerivedPeriodPart.MonthName:
                        valueExpression = "Date.MonthName(Date.From(" + dateField + ", \"en-US\"), \"en-US\")";
                        valueType = "type text";
                        break;
                    case DerivedPeriodPart.YearMonth:
                        valueExpression = "Date.ToText(Date.StartOfMonth(Date.From(" + dateField + ", \"en-US\")), \"yyyy-MM\", \"en-US\")";
                        valueType = "type text";
                        break;
                    default:
                        throw new MCompilationException("PERIOD_PART_UNSUPPORTED", "The derived period part cannot be compiled.", transform.Id);
                }

                expression = "Table.AddColumn(" + expression + ", " + MString(output) + ", each if "
                    + dateField + " is null then null else " + valueExpression + ", " + valueType + ")";
            }

            return expression;
        }

        private static string CompileArithmetic(string previous, AddArithmeticColumnTransform transform)
        {
            var output = RequireColumn(transform.OutputColumn, transform);
            var left = CompileArithmeticOperand(transform.Left, transform);
            var right = CompileArithmeticOperand(transform.Right, transform);
            string operation;
            switch (transform.Operator)
            {
                case ArithmeticOperator.Add:
                    operation = "Value.Add(left, right, Precision.Decimal)";
                    break;
                case ArithmeticOperator.Subtract:
                    operation = "Value.Subtract(left, right, Precision.Decimal)";
                    break;
                case ArithmeticOperator.Multiply:
                    operation = "Value.Multiply(left, right, Precision.Decimal)";
                    break;
                case ArithmeticOperator.Divide:
                    if (transform.Right.Kind == ArithmeticOperandKind.Number && transform.Right.Number == 0m)
                    {
                        throw new MCompilationException(
                            "ARITHMETIC_LITERAL_DIVIDE_BY_ZERO",
                            "A literal denominator cannot be zero.",
                            transform.Id);
                    }

                    if (!transform.ReturnNullOnZeroDenominator)
                    {
                        throw new MCompilationException(
                            "ARITHMETIC_DIVIDE_NULL_ON_ZERO_REQUIRED",
                            "Division must return blank when the denominator is zero.",
                            transform.Id);
                    }

                    operation = "if right is null or right = 0 then null else Value.Divide(left, right, Precision.Decimal)";
                    break;
                default:
                    throw new MCompilationException("ARITHMETIC_OPERATOR_UNSUPPORTED", "The arithmetic operation cannot be compiled.", transform.Id);
            }

            var resultType = transform.ResultType == ColumnDataType.WholeNumber ? "Int64.Type" : "type number";
            if (transform.ResultType != ColumnDataType.WholeNumber && transform.ResultType != ColumnDataType.DecimalNumber)
            {
                throw new MCompilationException(
                    "ARITHMETIC_RESULT_TYPE_INVALID",
                    "Arithmetic output must be whole-number or decimal-number typed.",
                    transform.Id);
            }

            var result = transform.ResultType == ColumnDataType.WholeNumber
                ? "if calculated is null then null else Int64.From(calculated, \"en-US\")"
                : "calculated";
            return "Table.AddColumn(" + previous + ", " + MString(output)
                + ", each let left = " + left + ", right = " + right
                + ", calculated = if left is null or right is null then null else " + operation
                + " in " + result + ", " + resultType + ")";
        }

        private static string CompileArithmeticOperand(ArithmeticOperand operand, TransformStep transform)
        {
            if (operand == null)
            {
                throw new MCompilationException("ARITHMETIC_OPERAND_REQUIRED", "An arithmetic operand is required.", transform.Id);
            }

            if (operand.Kind == ArithmeticOperandKind.Column)
            {
                var field = MField(RequireColumn(operand.Column, transform));
                return CompileBoundedDecimalValue(field, "Arithmetic operand");
            }

            if (!operand.Number.HasValue)
            {
                throw new MCompilationException("ARITHMETIC_NUMBER_REQUIRED", "A numeric operand requires a value.", transform.Id);
            }

            return "Decimal.From(" + MString(operand.Number.Value.ToString(CultureInfo.InvariantCulture))
                + ", \"en-US\")";
        }

        private static string CompileBoundedDecimalValue(string rawExpression, string errorReason)
        {
            return "(let raw = " + rawExpression
                + ", text = if raw is text then Text.Trim(raw) else null"
                + ", validText = text <> null and text <> \"\" and Text.Select(text, {\"0\"..\"9\", \"+\", \"-\", \".\", \",\", \"e\", \"E\"}) = text"
                + ", converted = if raw is null then [HasError = false, Value = null]"
                + " else if raw is number then try Decimal.From(raw, \"en-US\")"
                + " else if validText then try Decimal.From(text, \"en-US\")"
                + " else [HasError = true]"
                + " in if converted[HasError] then error Error.Record(" + MString("Invalid " + errorReason.ToLowerInvariant())
                + ", " + MString(errorReason + "s must be finite decimal numbers, numeric text, or blank.")
                + ", null) else converted[Value])";
        }

        private static string CompileBoundedTemporalValue(string rawExpression, bool dateTime)
        {
            string function = dateTime ? "DateTime" : "Date";
            IEnumerable<string> formats = dateTime
                ? BoundedDateTimeFormats.Concat(BoundedDateFormats)
                : BoundedDateFormats;
            string attempts = string.Join(", ", formats.Select(format =>
                "try " + function + ".FromText(text, [Format = " + MString(format)
                + ", Culture = \"en-US\"])"));
            string directKinds = dateTime
                ? "raw is datetime or raw is date or raw is number"
                : "raw is date or raw is datetime or raw is datetimezone or raw is number";
            string description = dateTime
                ? "Date-time values must be dates, date-times, Excel date serials, supported en-US date-time text, or blank."
                : "Date values must be dates, date-times, Excel date serials, supported en-US date text, or blank.";
            return "(let raw = " + rawExpression
                + ", text = if raw is text then Text.Trim(raw) else null"
                + ", attempts = if text = null or text = \"\" then {} else {" + attempts + "}"
                + ", successful = List.First(List.Select(attempts, each not _[HasError]), null)"
                + ", parsedText = if successful = null then null else successful[Value]"
                + " in if raw is null then null else if " + directKinds
                + " then " + function + ".From(raw, \"en-US\")"
                + " else if parsedText <> null then parsedText"
                + " else error Error.Record(" + MString("Invalid " + function.ToLowerInvariant() + " value")
                + ", " + MString(description) + ", null))";
        }

        private static string CompilePeriodNormalization(
            string previous,
            NormalizePeriodsTransform transform,
            PeriodMappingSpec? rootMapping)
        {
            if (string.IsNullOrWhiteSpace(transform.PeriodMappingId))
            {
                throw new MCompilationException(
                    "PERIOD_MAPPING_REFERENCE_REQUIRED",
                    "The normalize transform must reference an explicit period mapping.",
                    transform.Id);
            }

            if (rootMapping == null
                || !string.Equals(rootMapping.Id, transform.PeriodMappingId, StringComparison.OrdinalIgnoreCase))
            {
                throw new MCompilationException(
                    "PERIOD_MAPPING_REFERENCE_UNKNOWN",
                    "The normalize transform references an unknown explicit period mapping.",
                    transform.Id);
            }

            var mapping = rootMapping;

            if (mapping.Kind == PeriodMappingKind.LongDateColumn)
            {
                return CompileLongPeriodNormalization(previous, mapping, transform);
            }

            return CompileWidePeriodNormalization(previous, mapping, transform);
        }

        private static string CompileLongPeriodNormalization(
            string previous,
            PeriodMappingSpec mapping,
            TransformStep transform)
        {
            var column = RequireColumn(mapping.DateColumn, transform);
            if (!mapping.Grain.HasValue)
            {
                throw new MCompilationException(
                    "LONG_PERIOD_GRAIN_REQUIRED",
                    "A normalized long period column requires an explicit day, month, or quarter grain.",
                    transform.Id);
            }

            var grain = mapping.Grain.Value;
            var reportingYear = mapping.ReportingYear.HasValue
                ? mapping.ReportingYear.Value.ToString(CultureInfo.InvariantCulture)
                : "null";
            var grainName = grain.ToString().ToLowerInvariant();
            var canonicalizer = grain == PeriodGrain.Day
                ? "Date.From(checked)"
                : grain == PeriodGrain.Quarter
                    ? "Date.StartOfQuarter(Date.From(checked))"
                    : "Date.StartOfMonth(Date.From(checked))";

            // This function deliberately mirrors PeriodValueNormalizer's bounded,
            // culture-independent vocabulary. It never falls back to the current
            // year or the current Windows locale.
            var declarations = new List<string>
            {
                "reportingYear = " + reportingYear,
                "expectedGrain = " + MString(grainName),
                "isDigits = (token as nullable text) as logical => token <> null and token <> \"\" and Text.Select(token, {\"0\"..\"9\"}) = token",
                "fourDigitYear = (token as nullable text) as nullable number => let parsed = if token <> null and Text.Length(token) = 4 and isDigits(token) and (Text.StartsWith(token, \"19\") or Text.StartsWith(token, \"20\")) then Number.FromText(token, \"en-US\") else null in parsed",
                "tokenYear = (token as nullable text) as nullable number => let four = fourDigitYear(token), two = if token <> null and Text.Length(token) = 2 and isDigits(token) then Number.FromText(token, \"en-US\") else null in if four <> null then four else if two = null then null else if two <= 29 then 2000 + two else 1900 + two",
                "monthNumber = (token as nullable text) as nullable number => let cleaned = if token = null then null else Text.Upper(Text.Trim(token)), names = {\"JAN\",\"JANUARY\",\"FEB\",\"FEBRUARY\",\"MAR\",\"MARCH\",\"APR\",\"APRIL\",\"MAY\",\"JUN\",\"JUNE\",\"JUL\",\"JULY\",\"AUG\",\"AUGUST\",\"SEP\",\"SEPT\",\"SEPTEMBER\",\"OCT\",\"OCTOBER\",\"NOV\",\"NOVEMBER\",\"DEC\",\"DECEMBER\"}, months = {1,1,2,2,3,3,4,4,5,6,6,7,7,8,8,9,9,9,10,10,11,11,12,12}, position = if cleaned = null then -1 else List.PositionOf(names, cleaned) in if position < 0 then null else months{position}",
                "numericMonth = (token as nullable text) as nullable number => let parsed = if token <> null and (Text.Length(token) = 1 or Text.Length(token) = 2) and isDigits(token) then Number.FromText(token, \"en-US\") else null in if parsed <> null and parsed >= 1 and parsed <= 12 then parsed else null",
                "quarterNumber = (token as nullable text) as nullable number => let parsed = if token <> null and Text.Length(token) = 2 and Text.StartsWith(token, \"Q\") and isDigits(Text.End(token, 1)) then Number.FromText(Text.End(token, 1), \"en-US\") else null in if parsed <> null and parsed >= 1 and parsed <= 4 then parsed else null",
                "dayNumber = (token as nullable text) as nullable number => let parsed = if token <> null and (Text.Length(token) = 1 or Text.Length(token) = 2) and isDigits(token) then Number.FromText(token, \"en-US\") else null in if parsed <> null and parsed >= 1 and parsed <= 31 then parsed else null",
                "parseIsoDate = (token as nullable text) as nullable date => let validShape = token <> null and Text.Length(token) = 10 and (Text.Range(token, 4, 1) = \"-\" or Text.Range(token, 4, 1) = \"/\") and Text.Range(token, 7, 1) = Text.Range(token, 4, 1) and isDigits(Text.Remove(token, {\"-\",\"/\"})), y = if validShape then fourDigitYear(Text.Start(token, 4)) else null, m = if validShape then numericMonth(Text.Range(token, 5, 2)) else null, d = if validShape then dayNumber(Text.End(token, 2)) else null in if y = null or m = null or d = null then null else try #date(y, m, d) otherwise null",
                "validZone = (zone as text) as logical => zone = \"\" or zone = \"Z\" or (Text.Length(zone) = 6 and (Text.StartsWith(zone, \"+\") or Text.StartsWith(zone, \"-\")) and Text.Range(zone, 3, 1) = \":\" and isDigits(Text.Range(zone, 1, 2) & Text.End(zone, 2)) and Number.FromText(Text.Range(zone, 1, 2), \"en-US\") <= 14 and Number.FromText(Text.End(zone, 2), \"en-US\") <= 59)",
                "validTimeTail = (tail as text) as logical => let clock = if Text.Length(tail) >= 8 then Text.Start(tail, 8) else \"\", clockValid = Text.Length(clock) = 8 and Text.Range(clock, 2, 1) = \":\" and Text.Range(clock, 5, 1) = \":\" and isDigits(Text.Remove(clock, {\":\"})) and Number.FromText(Text.Start(clock, 2), \"en-US\") <= 23 and Number.FromText(Text.Range(clock, 3, 2), \"en-US\") <= 59 and Number.FromText(Text.End(clock, 2), \"en-US\") <= 59, suffix = if Text.Length(tail) > 8 then Text.Range(tail, 8) else \"\", zonePosition = if Text.StartsWith(suffix, \".\") then Text.PositionOfAny(suffix, {\"Z\",\"+\",\"-\"}) else -1, fraction = if not Text.StartsWith(suffix, \".\") then null else if zonePosition < 0 then Text.Range(suffix, 1) else Text.Range(suffix, 1, zonePosition - 1), zone = if Text.StartsWith(suffix, \".\") then (if zonePosition < 0 then \"\" else Text.Range(suffix, zonePosition)) else suffix, fractionValid = fraction <> null and Text.Length(fraction) >= 1 and Text.Length(fraction) <= 7 and isDigits(fraction), suffixValid = if suffix = \"\" then true else if Text.StartsWith(suffix, \".\") then fractionValid and validZone(zone) else validZone(suffix) in clockValid and suffixValid",
                "normalizePeriod = (raw as any) as nullable date => let rawText = if raw = null then null else Text.Trim(Text.From(raw, \"en-US\")), upperText = if rawText = null then null else Text.Upper(rawText), compactSpaces = if upperText = null then null else Text.Combine(List.Select(Text.SplitAny(upperText, \" #(tab)#(cr)#(lf)\"), each _ <> \"\"), \" \"), text = if compactSpaces = null then null else Text.Replace(compactSpaces, \"Q \", \"Q\"), digitsOnly = if text = null then null else Text.Select(text, {\"0\"..\"9\"}), compact = text <> null and Text.Length(text) = 6 and digitsOnly = text and fourDigitYear(Text.Start(text, 4)) <> null and numericMonth(Text.End(text, 2)) <> null, nativeDate = if raw is date or raw is datetime or raw is datetimezone then Date.From(raw) else if raw is number and not compact then try Date.From(raw) otherwise null else null, isoDate = parseIsoDate(text), isoPrefix = if text <> null and Text.Length(text) > 10 then parseIsoDate(Text.Start(text, 10)) else null, isoDelimiter = if text <> null and Text.Length(text) > 10 then Text.Range(text, 10, 1) else \"\", isoTimeValid = isoPrefix <> null and Text.Length(text) >= 19 and ((isoDelimiter = \"T\" and validTimeTail(Text.Range(text, 11))) or (isoDelimiter = \" \" and Text.Length(text) = 19 and validTimeTail(Text.Range(text, 11)))), zonedAttempt = if isoTimeValid then try DateTimeZone.FromText(text, [Culture=\"en-US\"]) else null, localAttempt = if isoTimeValid then try DateTime.FromText(text, [Culture=\"en-US\"]) else null, isoDateTime = if zonedAttempt <> null and not zonedAttempt[HasError] then Date.From(zonedAttempt[Value]) else if localAttempt <> null and not localAttempt[HasError] then Date.From(localAttempt[Value]) else null, parts = if text = null then {} else List.Select(Text.SplitAny(text, \" -_./\"), each _ <> \"\"), first = if List.Count(parts) > 0 then parts{0} else null, second = if List.Count(parts) > 1 then parts{1} else null, third = if List.Count(parts) > 2 then parts{2} else null, firstMonthName = monthNumber(first), secondMonthName = monthNumber(second), firstNumericMonth = numericMonth(first), secondNumericMonth = numericMonth(second), firstTokenYear = tokenYear(first), secondTokenYear = tokenYear(second), thirdFourYear = fourDigitYear(third), firstQuarter = quarterNumber(first), secondQuarter = quarterNumber(second), namedDay = if List.Count(parts) = 3 and Text.Combine(parts, \"-\") = text and secondMonthName <> null and thirdFourYear <> null and dayNumber(first) <> null then try #date(thirdFourYear, secondMonthName, dayNumber(first)) otherwise null else if List.Count(parts) = 3 and Text.Combine(parts, \" \") = text and firstMonthName <> null and thirdFourYear <> null and dayNumber(second) <> null then try #date(thirdFourYear, firstMonthName, dayNumber(second)) otherwise null else null, missingYear = List.Count(parts) = 1 and reportingYear = null and (firstQuarter <> null or firstMonthName <> null), parsed = if nativeDate <> null then [Value=nativeDate, Grain=null] else if isoDate <> null then [Value=isoDate, Grain=null] else if isoDateTime <> null then [Value=isoDateTime, Grain=null] else if namedDay <> null then [Value=namedDay, Grain=null] else if compact then [Value=#date(fourDigitYear(Text.Start(text, 4)), numericMonth(Text.End(text, 2)), 1), Grain=\"month\"] else if List.Count(parts) = 1 and firstQuarter <> null and reportingYear <> null then [Value=#date(reportingYear, ((firstQuarter - 1) * 3) + 1, 1), Grain=\"quarter\"] else if List.Count(parts) = 1 and firstMonthName <> null and reportingYear <> null then [Value=#date(reportingYear, firstMonthName, 1), Grain=\"month\"] else if List.Count(parts) = 2 and firstQuarter <> null and secondTokenYear <> null then [Value=#date(secondTokenYear, ((firstQuarter - 1) * 3) + 1, 1), Grain=\"quarter\"] else if List.Count(parts) = 2 and secondQuarter <> null and firstTokenYear <> null then [Value=#date(firstTokenYear, ((secondQuarter - 1) * 3) + 1, 1), Grain=\"quarter\"] else if List.Count(parts) = 2 and firstMonthName <> null and secondTokenYear <> null then [Value=#date(secondTokenYear, firstMonthName, 1), Grain=\"month\"] else if List.Count(parts) = 2 and firstTokenYear <> null and secondMonthName <> null then [Value=#date(firstTokenYear, secondMonthName, 1), Grain=\"month\"] else if List.Count(parts) = 2 and fourDigitYear(first) <> null and secondNumericMonth <> null then [Value=#date(fourDigitYear(first), secondNumericMonth, 1), Grain=\"month\"] else if List.Count(parts) = 2 and firstNumericMonth <> null and fourDigitYear(second) <> null then [Value=#date(fourDigitYear(second), firstNumericMonth, 1), Grain=\"month\"] else null, checked = if raw = null or rawText = \"\" then error Error.Record(\"Blank period\", \"A period value cannot be blank.\", null) else if missingYear then error Error.Record(\"Reporting year required\", \"A month or quarter without a year requires an explicit reporting year.\", rawText) else if parsed = null then error Error.Record(\"Unsupported period\", \"The period value is not one of the supported unambiguous formats.\", rawText) else if parsed[Grain] <> null and parsed[Grain] <> expectedGrain then error Error.Record(\"Mixed period grain\", \"The period value does not match the configured grain.\", rawText) else parsed[Value], result = " + canonicalizer + " in result",
                "converted = Table.TransformColumns(" + previous + ", {{" + MString(column) + ", each normalizePeriod(_), type date}}, null, MissingField.Error)"
            };
            int normalizationDeclarationIndex = declarations.Count - 2;
            declarations[normalizationDeclarationIndex] = declarations[normalizationDeclarationIndex]
                .Replace(
                    "nativeDate = if raw is date or raw is datetime or raw is datetimezone then Date.From(raw) else if raw is number and not compact then try Date.From(raw) otherwise null else null",
                    "numericCompactCandidate = raw is number and raw >= 100000 and raw <= 999999 and Number.Mod(raw, 1) = 0, nativeDate = if raw is date or raw is datetime or raw is datetimezone then Date.From(raw) else if raw is number and not numericCompactCandidate then try Date.From(raw) otherwise null else null")
                .Replace(
                    "parts = if text = null then {} else List.Select(Text.SplitAny(text, \" -_./\"), each _ <> \"\")",
                    "parts = if text = null then {} else Text.SplitAny(text, \" -_./\")");
            return "(let " + string.Join(", ", declarations) + " in converted)";
        }

        private static string CompileWidePeriodNormalization(
            string previous,
            PeriodMappingSpec mapping,
            TransformStep transform)
        {
            var keys = RequireColumns(mapping.KeyColumns, transform, allowEmpty: true);
            var outputNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                RequireColumn(mapping.PeriodColumnName, transform),
                RequireColumn(mapping.ValueColumnName, transform)
            };
            if (mapping.Kind == PeriodMappingKind.MetricMonthHeaders)
            {
                outputNames.Add(RequireColumn(mapping.MetricColumnName, transform));
            }

            if (keys.Any(outputNames.Contains))
            {
                throw new MCompilationException(
                    "PERIOD_OUTPUT_KEY_COLLISION",
                    "A period-normalization output name cannot also be a key column.",
                    transform.Id);
            }

            var periodColumns = mapping.Columns ?? new List<PeriodColumnMapping>();
            if (periodColumns.Count == 0)
            {
                throw new MCompilationException("PERIOD_COLUMNS_REQUIRED", "Period normalization requires mapped columns.", transform.Id);
            }

            var grain = mapping.Grain ?? PeriodGrain.Month;
            if (grain == PeriodGrain.Day)
            {
                throw new MCompilationException(
                    "WIDE_PERIOD_GRAIN_INVALID",
                    "Wide period normalization requires month or quarter grain.",
                    transform.Id);
            }

            var sourceColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var column in periodColumns)
            {
                RequireColumn(column.SourceColumn, transform);
                if (column.Month < 1 || column.Month > 12 || !sourceColumns.Add(column.SourceColumn))
                {
                    throw new MCompilationException("PERIOD_MAPPING_INVALID", "The period mapping contains an invalid or repeated source column.", transform.Id);
                }

                if (grain == PeriodGrain.Quarter
                    && column.Month != 1 && column.Month != 4
                    && column.Month != 7 && column.Month != 10)
                {
                    throw new MCompilationException(
                        "QUARTER_START_MONTH_INVALID",
                        "Quarter mappings must use the first month of each quarter.",
                        transform.Id);
                }

                var year = column.Year ?? mapping.ReportingYear;
                if (!year.HasValue)
                {
                    throw new MCompilationException(
                        "REPORTING_YEAR_REQUIRED",
                        "A month header without a year requires an explicit reporting year.",
                        transform.Id);
                }
            }

            var knownNames = new HashSet<string>(keys, StringComparer.OrdinalIgnoreCase);
            knownNames.UnionWith(periodColumns.Select(column => column.SourceColumn));
            knownNames.Add(mapping.PeriodColumnName);
            knownNames.Add(mapping.ValueColumnName);
            knownNames.Add(mapping.MetricColumnName);
            var headerColumn = ChooseInternalName("__erb_period_header", knownNames);
            knownNames.Add(headerColumn);
            var cellsColumn = ChooseInternalName("__erb_period_cells", knownNames);
            knownNames.Add(cellsColumn);
            var cellHeaderField = ChooseInternalName("__erb_cell_header", knownNames);
            knownNames.Add(cellHeaderField);
            var cellValueField = ChooseInternalName("__erb_cell_value", knownNames);
            var selectedColumns = keys.Concat(periodColumns.Select(column => column.SourceColumn)).ToList();
            var periodCase = CompileMappingCase(headerColumn, periodColumns, mapping.ReportingYear, false, transform);
            // Table.Unpivot omits null cells. Expand an explicit bounded record for
            // every mapped source column so row identity never depends on its value.
            var cellRecords = periodColumns.Select(column =>
                "Record.FromList({" + MString(column.SourceColumn) + ", "
                + MField(column.SourceColumn) + "}, {" + MString(cellHeaderField)
                + ", " + MString(cellValueField) + "})");
            var rowPreservingExpansion = "selected = Table.SelectColumns(" + previous + ", "
                + MStringList(selectedColumns) + ", MissingField.Error), withCells = Table.AddColumn(selected, "
                + MString(cellsColumn) + ", each {" + string.Join(", ", cellRecords)
                + "}, type list), withoutMappedColumns = Table.RemoveColumns(withCells, "
                + MStringList(periodColumns.Select(column => column.SourceColumn))
                + ", MissingField.Error), expandedCells = Table.ExpandListColumn(withoutMappedColumns, "
                + MString(cellsColumn) + "), expanded = Table.ExpandRecordColumn(expandedCells, "
                + MString(cellsColumn) + ", {" + MString(cellHeaderField) + ", "
                + MString(cellValueField) + "}, {" + MString(headerColumn) + ", "
                + MString(mapping.ValueColumnName) + "})";

            if (mapping.Kind == PeriodMappingKind.MonthHeaders)
            {
                if (periodColumns.Any(column => !string.IsNullOrWhiteSpace(column.Metric)))
                {
                    throw new MCompilationException("METRIC_NOT_ALLOWED", "Month-only normalization cannot contain metric names.", transform.Id);
                }

                return "(let " + rowPreservingExpansion
                    + ", withPeriod = Table.AddColumn(expanded, " + MString(mapping.PeriodColumnName)
                    + ", each " + periodCase + ", type date), cleaned = Table.RemoveColumns(withPeriod, {"
                    + MString(headerColumn) + "}, MissingField.Error) in cleaned)";
            }

            if (periodColumns.Any(column => string.IsNullOrWhiteSpace(column.Metric)))
            {
                throw new MCompilationException("METRIC_REQUIRED", "Every metric-month mapping requires a metric name.", transform.Id);
            }

            ValidateCompleteMetricMatrix(periodColumns, mapping.ReportingYear, transform);
            var metricCase = CompileMappingCase(headerColumn, periodColumns, mapping.ReportingYear, true, transform);
            return "(let " + rowPreservingExpansion
                + ", withPeriod = Table.AddColumn(expanded, " + MString(mapping.PeriodColumnName)
                + ", each " + periodCase + ", type date), withMetric = Table.AddColumn(withPeriod, "
                + MString(mapping.MetricColumnName) + ", each " + metricCase
                + ", type text), cleaned = Table.RemoveColumns(withMetric, {" + MString(headerColumn)
                + "}, MissingField.Error) in cleaned)";
        }

        private static string CompileMappingCase(
            string headerColumn,
            IReadOnlyList<PeriodColumnMapping> mappings,
            int? reportingYear,
            bool metric,
            TransformStep transform)
        {
            var field = MField(headerColumn);
            var builder = new StringBuilder();
            foreach (var mapping in mappings)
            {
                var year = mapping.Year ?? reportingYear;
                if (!year.HasValue)
                {
                    throw new MCompilationException("REPORTING_YEAR_REQUIRED", "A reporting year is required.", transform.Id);
                }

                builder.Append("if ");
                builder.Append(field);
                builder.Append(" = ");
                builder.Append(MString(mapping.SourceColumn));
                builder.Append(" then ");
                if (metric)
                {
                    builder.Append(MString(mapping.Metric ?? string.Empty));
                }
                else
                {
                    builder.Append("#date(");
                    builder.Append(year.Value.ToString(CultureInfo.InvariantCulture));
                    builder.Append(", ");
                    builder.Append(mapping.Month.ToString(CultureInfo.InvariantCulture));
                    builder.Append(", 1)");
                }

                builder.Append(" else ");
            }

            builder.Append("error \"Unmapped period header\"");
            return builder.ToString();
        }

        private static void ValidateCompleteMetricMatrix(
            IReadOnlyCollection<PeriodColumnMapping> mappings,
            int? reportingYear,
            TransformStep transform)
        {
            var allPeriods = new HashSet<string>(
                mappings.Select(mapping => PeriodKey(mapping, reportingYear)),
                StringComparer.Ordinal);
            foreach (var group in mappings.GroupBy(mapping => mapping.Metric ?? string.Empty, StringComparer.OrdinalIgnoreCase))
            {
                var periods = new HashSet<string>(group.Select(mapping => PeriodKey(mapping, reportingYear)), StringComparer.Ordinal);
                if (periods.Count != group.Count() || !periods.SetEquals(allPeriods))
                {
                    throw new MCompilationException(
                        "METRIC_PERIOD_MATRIX_INVALID",
                        "Every metric must map exactly once to every period.",
                        transform.Id);
                }
            }
        }

        private static string PeriodKey(PeriodColumnMapping mapping, int? reportingYear)
        {
            var year = mapping.Year ?? reportingYear;
            return (year.HasValue ? year.Value.ToString("0000", CultureInfo.InvariantCulture) : "????")
                + "-" + mapping.Month.ToString("00", CultureInfo.InvariantCulture);
        }

        private static string MLiteral(ScalarValue value, TransformStep transform)
        {
            if (value == null)
            {
                throw new MCompilationException("SCALAR_REQUIRED", "A literal value is required.", transform.Id);
            }

            switch (value.Kind)
            {
                case ScalarValueKind.Null:
                    if (HasScalarPayload(value))
                    {
                        break;
                    }

                    return "null";
                case ScalarValueKind.Text:
                    if (value.Text != null && ScalarPayloadCount(value) == 1)
                    {
                        return MString(value.Text);
                    }

                    break;
                case ScalarValueKind.Number:
                    if (value.Number.HasValue && ScalarPayloadCount(value) == 1)
                    {
                        return value.Number.Value.ToString(CultureInfo.InvariantCulture);
                    }

                    break;
                case ScalarValueKind.Boolean:
                    if (value.Boolean.HasValue && ScalarPayloadCount(value) == 1)
                    {
                        return value.Boolean.Value ? "true" : "false";
                    }

                    break;
                case ScalarValueKind.Date:
                    if (value.Temporal.HasValue && value.Temporal.Value.TimeOfDay == TimeSpan.Zero
                        && ScalarPayloadCount(value) == 1)
                    {
                        var date = value.Temporal.Value;
                        return "#date(" + date.Year.ToString(CultureInfo.InvariantCulture) + ", "
                            + date.Month.ToString(CultureInfo.InvariantCulture) + ", "
                            + date.Day.ToString(CultureInfo.InvariantCulture) + ")";
                    }

                    break;
                case ScalarValueKind.DateTime:
                    if (value.Temporal.HasValue && ScalarPayloadCount(value) == 1)
                    {
                        var dateTime = value.Temporal.Value;
                        var seconds = dateTime.Second + dateTime.Millisecond / 1000m;
                        return "#datetime(" + dateTime.Year.ToString(CultureInfo.InvariantCulture) + ", "
                            + dateTime.Month.ToString(CultureInfo.InvariantCulture) + ", "
                            + dateTime.Day.ToString(CultureInfo.InvariantCulture) + ", "
                            + dateTime.Hour.ToString(CultureInfo.InvariantCulture) + ", "
                            + dateTime.Minute.ToString(CultureInfo.InvariantCulture) + ", "
                            + seconds.ToString(CultureInfo.InvariantCulture) + ")";
                    }

                    break;
            }

            throw new MCompilationException("SCALAR_SHAPE_INVALID", "The literal payload does not match its kind.", transform.Id);
        }

        private static string MType(ColumnDataType type)
        {
            switch (type)
            {
                case ColumnDataType.Text:
                    return "type text";
                case ColumnDataType.WholeNumber:
                    return "Int64.Type";
                case ColumnDataType.DecimalNumber:
                    return "type number";
                case ColumnDataType.Boolean:
                    return "type logical";
                case ColumnDataType.Date:
                    return "type date";
                case ColumnDataType.DateTime:
                    return "type datetime";
                default:
                    throw new MCompilationException("COLUMN_TYPE_UNSUPPORTED", "The column type cannot be compiled.");
            }
        }

        private static string MString(string value)
        {
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            if (value.Any(char.IsControl))
            {
                throw new MCompilationException("CONTROL_CHARACTER_NOT_ALLOWED", "Power Query identifiers and literals cannot contain control characters.");
            }

            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }

        private static string MStringList(IEnumerable<string> values)
        {
            return "{" + string.Join(", ", values.Select(MString)) + "}";
        }

        private static string MField(string column)
        {
            return "[#\"" + column.Replace("\"", "\"\"") + "\"]";
        }

        private static string MIdentifier(string identifier)
        {
            return "#\"" + identifier.Replace("\"", "\"\"") + "\"";
        }

        private static string RequireColumn(string? column, TransformStep transform)
        {
            if (string.IsNullOrWhiteSpace(column) || column!.Length > 255 || column.Any(char.IsControl))
            {
                throw new MCompilationException("COLUMN_NAME_INVALID", "A bounded non-blank column name is required.", transform.Id);
            }

            return column;
        }

        private static List<string> RequireColumns(
            IEnumerable<string> columns,
            TransformStep transform,
            bool allowEmpty = false)
        {
            if (columns == null)
            {
                throw new MCompilationException("COLUMN_LIST_REQUIRED", "A column list is required.", transform.Id);
            }

            var materialized = columns.Select(column => RequireColumn(column, transform)).ToList();
            if (!allowEmpty && materialized.Count == 0)
            {
                throw new MCompilationException("COLUMN_LIST_EMPTY", "At least one column is required.", transform.Id);
            }

            if (materialized.Count != materialized.Distinct(StringComparer.OrdinalIgnoreCase).Count())
            {
                throw new MCompilationException("COLUMN_LIST_DUPLICATE", "Column lists cannot contain duplicates.", transform.Id);
            }

            return materialized;
        }

        private static void RequireNoFilterValue(FilterRowsTransform transform)
        {
            if (transform.Value != null)
            {
                throw new MCompilationException("FILTER_VALUE_NOT_ALLOWED", "Blank filters do not accept a literal value.", transform.Id);
            }
        }

        private static void RequireTextLiteral(ScalarValue value, TransformStep transform)
        {
            if (value == null || value.Kind != ScalarValueKind.Text || value.Text == null)
            {
                throw new MCompilationException("TEXT_LITERAL_REQUIRED", "Text matching requires a text literal.", transform.Id);
            }
        }

        private static int ScalarPayloadCount(ScalarValue value)
        {
            return (value.Text != null ? 1 : 0)
                + (value.Number.HasValue ? 1 : 0)
                + (value.Boolean.HasValue ? 1 : 0)
                + (value.Temporal.HasValue ? 1 : 0);
        }

        private static bool HasScalarPayload(ScalarValue value)
        {
            return ScalarPayloadCount(value) != 0;
        }

        private static string ChooseInternalName(string preferred, HashSet<string> knownNames)
        {
            var candidate = preferred;
            var suffix = 1;
            while (knownNames.Contains(candidate))
            {
                candidate = preferred + "_" + suffix.ToString(CultureInfo.InvariantCulture);
                suffix++;
            }

            return candidate;
        }
    }
}
