using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Core.Transforms;

namespace ExcelReportBuilder.AddIn.Host
{
    /// <summary>
    /// Translates the bounded manual transformation editor into the closed
    /// transformation grammar. The translator never accepts formulas, code, or
    /// expressions. Every result is a typed data operation or literal value.
    /// </summary>
    internal sealed class ManualTransformTranslator
    {
        private const int MaximumTransforms = 100;
        private const int MaximumListItems = 256;
        private const int MaximumColumnNameLength = 255;
        private const int MaximumLiteralLength = 1024;
        private const int MaximumDetailsLength = 262144;
        private static readonly string[] DateTimeLiteralFormats =
        {
            "yyyy-MM-dd'T'HH:mm:ss",
            "yyyy-MM-dd'T'HH:mm:ss.F",
            "yyyy-MM-dd'T'HH:mm:ss.FF",
            "yyyy-MM-dd'T'HH:mm:ss.FFF"
        };

        public List<TransformStep> Translate(IReadOnlyList<ManualTransformSnapshot> snapshots)
        {
            if (snapshots == null)
            {
                throw new ArgumentNullException(nameof(snapshots));
            }

            if (snapshots.Count > MaximumTransforms)
            {
                throw new InvalidOperationException(
                    "A report setup can contain at most 100 manual transformations.");
            }

            var result = new List<TransformStep>(snapshots.Count);
            for (var index = 0; index < snapshots.Count; index++)
            {
                ManualTransformSnapshot snapshot = snapshots[index];
                if (snapshot == null)
                {
                    throw Error(index, "Choose a transformation operation or remove the empty row.");
                }

                ValidateInputLengths(snapshot, index);
                if (index != 0 && string.Equals(
                        NormalizeToken(snapshot.Operation),
                        "excludetotalrows",
                        StringComparison.Ordinal))
                {
                    throw Error(
                        index,
                        "Exclude total rows must be the first preparation step so its evidence can be checked against the unchanged source.");
                }
                result.Add(TranslateOne(snapshot, index));
            }

            return result;
        }

        private static TransformStep TranslateOne(ManualTransformSnapshot snapshot, int index)
        {
            string operation = NormalizeToken(snapshot.Operation);
            switch (operation)
            {
                case "keepcolumns":
                    EnsureBlank(snapshot.OutputColumn, index, "Output column", snapshot.Operation);
                    return WithId(
                        new KeepColumnsTransform
                        {
                            Columns = ParseColumns(
                                ReadColumnListInput(snapshot, index),
                                index,
                                "Keep columns")
                        },
                        index,
                        "keep-columns");

                case "removecolumns":
                    EnsureBlank(snapshot.OutputColumn, index, "Output column", snapshot.Operation);
                    return WithId(
                        new RemoveColumnsTransform
                        {
                            Columns = ParseColumns(
                                ReadColumnListInput(snapshot, index),
                                index,
                                "Remove columns")
                        },
                        index,
                        "remove-columns");

                case "reordercolumns":
                    EnsureBlank(snapshot.OutputColumn, index, "Output column", snapshot.Operation);
                    return WithId(
                        new ReorderColumnsTransform
                        {
                            Columns = ParseColumns(
                                ReadColumnListInput(snapshot, index),
                                index,
                                "Reorder columns")
                        },
                        index,
                        "reorder-columns");

                case "renamecolumn":
                    EnsureBlank(snapshot.Details, index, "Details", snapshot.Operation);
                    return WithId(
                        new RenameColumnTransform
                        {
                            From = RequireColumn(snapshot.Column, index, "Column"),
                            To = RequireColumn(snapshot.OutputColumn, index, "Output column")
                        },
                        index,
                        "rename-column");

                case "converttype":
                case "changetype":
                    EnsureBlank(snapshot.OutputColumn, index, "Output column", snapshot.Operation);
                    return WithId(
                        new ChangeColumnTypeTransform
                        {
                            Column = RequireColumn(snapshot.Column, index, "Column"),
                            DataType = ParseColumnDataType(snapshot.Details, index)
                        },
                        index,
                        "convert-type");

                case "trimtext":
                    EnsureBlank(snapshot.OutputColumn, index, "Output column", snapshot.Operation);
                    return WithId(
                        new TrimTextTransform
                        {
                            Columns = ParseColumns(
                                ReadColumnListInput(snapshot, index),
                                index,
                                "Trim text")
                        },
                        index,
                        "trim-text");

                case "replacevalue":
                    EnsureBlank(snapshot.OutputColumn, index, "Output column", snapshot.Operation);
                    return TranslateReplacement(snapshot, index);

                case "normalizeblanks":
                    EnsureBlank(snapshot.OutputColumn, index, "Output column", snapshot.Operation);
                    return WithId(
                        new NormalizeBlanksTransform
                        {
                            Columns = ParseColumns(snapshot.Column, index, "Normalize blanks"),
                            Replacement = ParseOptionalTypedLiteralOrNull(
                                snapshot.Details,
                                index,
                                "blank replacement"),
                            TreatWhitespaceAsBlank = true
                        },
                        index,
                        "normalize-blanks");

                case "normalizeerrors":
                    EnsureBlank(snapshot.OutputColumn, index, "Output column", snapshot.Operation);
                    return WithId(
                        new NormalizeErrorsTransform
                        {
                            Columns = ParseColumns(snapshot.Column, index, "Normalize errors"),
                            Replacement = ParseOptionalTypedLiteralOrNull(
                                snapshot.Details,
                                index,
                                "error replacement")
                        },
                        index,
                        "normalize-errors");

                case "filldown":
                    EnsureBlank(snapshot.OutputColumn, index, "Output column", snapshot.Operation);
                    return WithId(
                        new FillDownTransform
                        {
                            Columns = ParseColumns(
                                ReadColumnListInput(snapshot, index),
                                index,
                                "Fill down")
                        },
                        index,
                        "fill-down");

                case "mapvalues":
                    EnsureBlank(snapshot.OutputColumn, index, "Output column", snapshot.Operation);
                    return TranslateValueMap(snapshot, index);

                case "filterrows":
                    EnsureBlank(snapshot.OutputColumn, index, "Output column", snapshot.Operation);
                    return TranslateFilter(snapshot, index);

                case "excludetotalrows":
                    EnsureBlank(snapshot.OutputColumn, index, "Output column", snapshot.Operation);
                    return TranslateTotalRowExclusion(snapshot, index);

                case "deriveperiodparts":
                    EnsureBlank(snapshot.OutputColumn, index, "Output column", snapshot.Operation);
                    return TranslatePeriodParts(snapshot, index);

                case "arithmetic":
                case "addarithmeticcolumn":
                    return TranslateArithmetic(snapshot, index);

                default:
                    throw Error(
                        index,
                        "The operation '" + Display(snapshot.Operation) + "' is not supported. "
                        + "Choose one of the operations shown in the transformation list.");
            }
        }

        private static TransformStep TranslateReplacement(ManualTransformSnapshot snapshot, int index)
        {
            string[] parts = SplitSingle(snapshot.Details, "=>", index, "Replace value uses old => new.");
            return WithId(
                new ReplaceValueTransform
                {
                    Column = RequireColumn(snapshot.Column, index, "Column"),
                    Find = ParseTextOrNull(parts[0], index, "value to replace"),
                    ReplaceWith = ParseTextOrNull(parts[1], index, "replacement value")
                },
                index,
                "replace-value");
        }

        private static TransformStep TranslateValueMap(ManualTransformSnapshot snapshot, int index)
        {
            string column = RequireColumn(snapshot.Column, index, "Column");
            IReadOnlyList<string> mappings = SplitList(snapshot.Details, index, "Map values");
            var entries = new List<ValueMapEntry>(mappings.Count);
            var inputs = new HashSet<string>(StringComparer.Ordinal);
            for (var itemIndex = 0; itemIndex < mappings.Count; itemIndex++)
            {
                string[] pair = SplitSingle(
                    mappings[itemIndex],
                    "=>",
                    index,
                    "Each value mapping must use old => new.");
                ScalarValue from = ParseTypedLiteral(pair[0], index, "mapped source value");
                ScalarValue to = ParseTypedLiteral(pair[1], index, "mapped output value");
                if (!inputs.Add(ScalarKey(from)))
                {
                    throw Error(index, "Each source value can appear only once in a value map.");
                }

                entries.Add(new ValueMapEntry { From = from, To = to });
            }

            return WithId(
                new MapValuesTransform
                {
                    Column = column,
                    Entries = entries
                },
                index,
                "map-values");
        }

        private static TransformStep TranslateFilter(ManualTransformSnapshot snapshot, int index)
        {
            string column = RequireColumn(snapshot.Column, index, "Column");
            string details = RequireText(snapshot.Details, index, "Filter details", MaximumDetailsLength);
            int delimiter = details.IndexOf(':');
            string operatorText = delimiter < 0 ? details : details.Substring(0, delimiter);
            string valueText = delimiter < 0 ? string.Empty : details.Substring(delimiter + 1);
            RowFilterOperator filterOperator = ParseFilterOperator(operatorText, index);
            bool requiresValue = filterOperator != RowFilterOperator.IsBlank
                && filterOperator != RowFilterOperator.IsNotBlank;

            if (!requiresValue && delimiter >= 0 && valueText.Trim().Length != 0)
            {
                throw Error(index, "Is blank and Is not blank do not accept a filter value.");
            }

            if (requiresValue && delimiter < 0)
            {
                throw Error(
                    index,
                    "Filter rows uses operator:value, for example equal:text:Open or greaterThan:number:100.");
            }

            ScalarValue? value = requiresValue
                ? ParseTypedLiteral(valueText, index, "filter value")
                : null;
            bool isTextMatch = filterOperator == RowFilterOperator.Contains
                || filterOperator == RowFilterOperator.StartsWith
                || filterOperator == RowFilterOperator.EndsWith;
            if (isTextMatch && (value == null
                    || value.Kind != ScalarValueKind.Text
                    || value.Text == null))
            {
                throw Error(index, "Contains, Starts with, and Ends with require a text filter value.");
            }

            if (isTextMatch && value!.Text!.Length == 0)
            {
                throw Error(index, "Contains, Starts with, and Ends with require non-empty text.");
            }

            return WithId(
                new FilterRowsTransform
                {
                    Column = column,
                    Operator = filterOperator,
                    Value = value
                },
                index,
                "filter-rows");
        }

        private static TransformStep TranslateTotalRowExclusion(
            ManualTransformSnapshot snapshot,
            int index)
        {
            string column = RequireColumn(snapshot.Column, index, "Column");
            IReadOnlyList<string> rawValues = SplitList(
                snapshot.Details,
                index,
                "Exclude total rows");
            var values = new List<ScalarValue>(rawValues.Count);
            var unique = new HashSet<string>(StringComparer.Ordinal);
            foreach (string rawValue in rawValues)
            {
                ScalarValue value = ParseTypedLiteral(rawValue, index, "confirmed total-row value");
                if (!unique.Add(ScalarKey(value)))
                {
                    throw Error(index, "Confirmed total-row values cannot be repeated.");
                }

                values.Add(value);
            }

            return WithId(
                new ExcludeTotalRowsTransform
                {
                    RequireAllEvidence = false,
                    Evidence = new List<TotalRowEvidenceSpec>
                    {
                        new TotalRowEvidenceSpec
                        {
                            Column = column,
                            MatchKind = TotalRowMatchKind.EqualsAny,
                            Values = values,
                            Source = EvidenceSource.UserConfirmation,
                            ObservedMatchCount = 0L
                        }
                    }
                },
                index,
                "exclude-total-rows");
        }

        private static TransformStep TranslatePeriodParts(ManualTransformSnapshot snapshot, int index)
        {
            string dateColumn = RequireColumn(snapshot.Column, index, "Date column");
            IReadOnlyList<string> definitions = SplitList(
                snapshot.Details,
                index,
                "Derive period parts");
            if (definitions.Count > 6)
            {
                throw Error(index, "Derive period parts can create at most six output columns.");
            }

            var columns = new List<DerivedPeriodColumnSpec>(definitions.Count);
            var parts = new HashSet<DerivedPeriodPart>();
            var outputs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string definition in definitions)
            {
                string[] pair = SplitSingle(
                    definition,
                    "=",
                    index,
                    "Period parts use Part=Output column, separated with semicolons.");
                DerivedPeriodPart part = ParseDerivedPeriodPart(pair[0], index);
                string output = RequireColumn(pair[1], index, "Derived output column");
                if (!parts.Add(part))
                {
                    throw Error(index, "Each period part can be derived only once.");
                }

                if (!outputs.Add(output))
                {
                    throw Error(index, "Each derived period output must have a different column name.");
                }

                columns.Add(new DerivedPeriodColumnSpec
                {
                    Part = part,
                    OutputColumn = output
                });
            }

            return WithId(
                new DerivePeriodPartsTransform
                {
                    DateColumn = dateColumn,
                    Columns = columns
                },
                index,
                "derive-period-parts");
        }

        private static TransformStep TranslateArithmetic(ManualTransformSnapshot snapshot, int index)
        {
            string leftColumn = RequireColumn(snapshot.Column, index, "Left column");
            string outputColumn = RequireColumn(snapshot.OutputColumn, index, "Output column");
            string details = RequireText(snapshot.Details, index, "Arithmetic details", MaximumDetailsLength);
            int delimiter = details.IndexOf(':');
            if (delimiter <= 0 || delimiter == details.Length - 1)
            {
                throw Error(
                    index,
                    "Arithmetic uses operation:right operand, for example divide:Units or multiply:1.25.");
            }

            if (details.IndexOf(':', delimiter + 1) >= 0)
            {
                throw Error(index, "Arithmetic details can contain only one colon.");
            }

            ArithmeticOperator arithmeticOperator = ParseArithmeticOperator(
                details.Substring(0, delimiter),
                index);
            string rightText = details.Substring(delimiter + 1).Trim();
            ArithmeticOperand right;
            if (decimal.TryParse(rightText, NumberStyles.Float, CultureInfo.InvariantCulture, out decimal number))
            {
                if (arithmeticOperator == ArithmeticOperator.Divide && number == 0m)
                {
                    throw Error(index, "The arithmetic denominator cannot be the number zero.");
                }

                right = new ArithmeticOperand
                {
                    Kind = ArithmeticOperandKind.Number,
                    Number = number
                };
            }
            else
            {
                right = new ArithmeticOperand
                {
                    Kind = ArithmeticOperandKind.Column,
                    Column = RequireColumn(rightText, index, "Right column")
                };
            }

            return WithId(
                new AddArithmeticColumnTransform
                {
                    OutputColumn = outputColumn,
                    Operator = arithmeticOperator,
                    Left = new ArithmeticOperand
                    {
                        Kind = ArithmeticOperandKind.Column,
                        Column = leftColumn
                    },
                    Right = right,
                    ResultType = ColumnDataType.DecimalNumber,
                    ReturnNullOnZeroDenominator = true
                },
                index,
                "arithmetic");
        }

        private static string ReadColumnListInput(ManualTransformSnapshot snapshot, int index)
        {
            bool hasColumn = !string.IsNullOrWhiteSpace(snapshot.Column);
            bool hasDetails = !string.IsNullOrWhiteSpace(snapshot.Details);
            if (hasColumn == hasDetails)
            {
                throw Error(
                    index,
                    "Enter the semicolon-separated column list in either Column or Details, but not both.");
            }

            return hasColumn ? snapshot.Column : snapshot.Details;
        }

        private static List<string> ParseColumns(string raw, int index, string operation)
        {
            IReadOnlyList<string> items = SplitList(raw, index, operation);
            var columns = new List<string>(items.Count);
            var unique = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string item in items)
            {
                string column = RequireColumn(item, index, operation + " column");
                if (!unique.Add(column))
                {
                    throw Error(index, operation + " cannot list the same column more than once.");
                }

                columns.Add(column);
            }

            return columns;
        }

        private static IReadOnlyList<string> SplitList(string raw, int index, string label)
        {
            string text = RequireText(raw, index, label + " details", MaximumDetailsLength);
            string[] items = text.Split(new[] { ';' }, StringSplitOptions.None);
            if (items.Length > MaximumListItems)
            {
                throw Error(index, label + " can contain at most 256 entries.");
            }

            var result = new List<string>(items.Length);
            for (var itemIndex = 0; itemIndex < items.Length; itemIndex++)
            {
                string item = items[itemIndex].Trim();
                if (item.Length == 0)
                {
                    throw Error(index, label + " contains an empty entry. Remove extra semicolons.");
                }

                result.Add(item);
            }

            return result;
        }

        private static string[] SplitSingle(string raw, string separator, int index, string guidance)
        {
            string text = RequireText(raw, index, "Details", MaximumDetailsLength);
            int delimiter = text.IndexOf(separator, StringComparison.Ordinal);
            if (delimiter < 0
                || text.IndexOf(separator, delimiter + separator.Length, StringComparison.Ordinal) >= 0)
            {
                throw Error(index, guidance);
            }

            string left = text.Substring(0, delimiter).Trim();
            string right = text.Substring(delimiter + separator.Length).Trim();
            if (left.Length == 0 || right.Length == 0)
            {
                throw Error(index, guidance);
            }

            return new[] { left, right };
        }

        private static ColumnDataType ParseColumnDataType(string raw, int index)
        {
            switch (NormalizeToken(raw))
            {
                case "text": return ColumnDataType.Text;
                case "wholenumber": return ColumnDataType.WholeNumber;
                case "decimalnumber": return ColumnDataType.DecimalNumber;
                case "boolean": return ColumnDataType.Boolean;
                case "date": return ColumnDataType.Date;
                case "datetime": return ColumnDataType.DateTime;
                default:
                    throw Error(
                        index,
                        "Convert type must be Text, Whole number, Decimal number, Boolean, Date, or Date time.");
            }
        }

        private static RowFilterOperator ParseFilterOperator(string raw, int index)
        {
            switch (NormalizeToken(raw))
            {
                case "equal":
                case "equals":
                    return RowFilterOperator.Equal;
                case "notequal":
                case "notequals":
                    return RowFilterOperator.NotEqual;
                case "greaterthan": return RowFilterOperator.GreaterThan;
                case "greaterthanorequal":
                case "greaterthanorequals":
                    return RowFilterOperator.GreaterThanOrEqual;
                case "lessthan": return RowFilterOperator.LessThan;
                case "lessthanorequal":
                case "lessthanorequals":
                    return RowFilterOperator.LessThanOrEqual;
                case "contains": return RowFilterOperator.Contains;
                case "startswith": return RowFilterOperator.StartsWith;
                case "endswith": return RowFilterOperator.EndsWith;
                case "isblank": return RowFilterOperator.IsBlank;
                case "isnotblank": return RowFilterOperator.IsNotBlank;
                default:
                    throw Error(
                        index,
                        "The filter operator is not supported. Use Equal, Not equal, Greater than, "
                        + "Greater than or equal, Less than, Less than or equal, Contains, Starts with, "
                        + "Ends with, Is blank, or Is not blank.");
            }
        }

        private static DerivedPeriodPart ParseDerivedPeriodPart(string raw, int index)
        {
            switch (NormalizeToken(raw))
            {
                case "year": return DerivedPeriodPart.Year;
                case "half": return DerivedPeriodPart.Half;
                case "quarter": return DerivedPeriodPart.Quarter;
                case "monthnumber": return DerivedPeriodPart.MonthNumber;
                case "monthname": return DerivedPeriodPart.MonthName;
                case "yearmonth": return DerivedPeriodPart.YearMonth;
                default:
                    throw Error(
                        index,
                        "A period part must be Year, Half, Quarter, Month number, Month name, or Year month.");
            }
        }

        private static ArithmeticOperator ParseArithmeticOperator(string raw, int index)
        {
            switch (NormalizeToken(raw))
            {
                case "add": return ArithmeticOperator.Add;
                case "subtract": return ArithmeticOperator.Subtract;
                case "multiply": return ArithmeticOperator.Multiply;
                case "divide": return ArithmeticOperator.Divide;
                default:
                    throw Error(index, "Arithmetic operation must be Add, Subtract, Multiply, or Divide.");
            }
        }

        private static ScalarValue ParseOptionalTypedLiteralOrNull(string raw, int index, string label)
        {
            return string.IsNullOrWhiteSpace(raw)
                ? ScalarValue.Null()
                : ParseTypedLiteral(raw, index, label);
        }

        private static ScalarValue ParseTextOrNull(string raw, int index, string label)
        {
            string value = RequireText(raw, index, label, MaximumLiteralLength);
            if (string.Equals(value, "null", StringComparison.OrdinalIgnoreCase))
            {
                return ScalarValue.Null();
            }

            if (value.StartsWith("text:", StringComparison.OrdinalIgnoreCase))
            {
                value = value.Substring(5);
            }

            ValidateLiteral(value, index, label, allowEmpty: true);
            return ScalarValue.FromText(value);
        }

        private static ScalarValue ParseTypedLiteral(string raw, int index, string label)
        {
            string value = RequireText(raw, index, label, MaximumLiteralLength);
            if (string.Equals(value, "null", StringComparison.OrdinalIgnoreCase))
            {
                return ScalarValue.Null();
            }

            int typeDelimiter = value.IndexOf(':');
            if (typeDelimiter >= 0)
            {
                if (typeDelimiter == 0)
                {
                    throw Error(
                        index,
                        "The " + label + " has a blank type prefix. Use text:, number:, boolean:, date:, or datetime:.");
                }

                string type = NormalizeToken(value.Substring(0, typeDelimiter));
                string payload = value.Substring(typeDelimiter + 1);
                switch (type)
                {
                    case "text":
                        ValidateLiteral(payload, index, label, allowEmpty: true);
                        return ScalarValue.FromText(payload);
                    case "number":
                        if (!decimal.TryParse(payload, NumberStyles.Float, CultureInfo.InvariantCulture, out decimal number))
                        {
                            throw Error(index, "The " + label + " must use an invariant number such as number:1250.50.");
                        }

                        return ScalarValue.FromNumber(number);
                    case "boolean":
                    case "bool":
                        if (!bool.TryParse(payload, out bool boolean))
                        {
                            throw Error(index, "The " + label + " must be boolean:true or boolean:false.");
                        }

                        return ScalarValue.FromBoolean(boolean);
                    case "date":
                        if (!DateTime.TryParseExact(
                                payload,
                                "yyyy-MM-dd",
                                CultureInfo.InvariantCulture,
                                DateTimeStyles.None,
                                out DateTime date))
                        {
                            throw Error(index, "The " + label + " must use date:yyyy-MM-dd.");
                        }

                        return ScalarValue.FromDate(date);
                    case "datetime":
                        if (!DateTime.TryParseExact(
                                payload,
                                DateTimeLiteralFormats,
                                CultureInfo.InvariantCulture,
                                DateTimeStyles.None,
                                out DateTime dateTime))
                        {
                            throw Error(
                                index,
                                "The " + label
                                + " must use datetime:yyyy-MM-ddTHH:mm:ss with no more than three fractional digits.");
                        }

                        return ScalarValue.FromDateTime(dateTime);
                    default:
                        throw Error(
                            index,
                            "The type prefix '" + value.Substring(0, typeDelimiter)
                            + "' is not supported. Use text:, number:, boolean:, date:, or datetime:. "
                            + "Prefix text containing a colon with text:.");
                }
            }

            if (bool.TryParse(value, out bool inferredBoolean))
            {
                return ScalarValue.FromBoolean(inferredBoolean);
            }

            if (decimal.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out decimal inferredNumber))
            {
                return ScalarValue.FromNumber(inferredNumber);
            }

            if (DateTime.TryParseExact(
                    value,
                    "yyyy-MM-dd",
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out DateTime inferredDate))
            {
                return ScalarValue.FromDate(inferredDate);
            }

            if (DateTime.TryParseExact(
                    value,
                    DateTimeLiteralFormats,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out DateTime inferredDateTime))
            {
                return ScalarValue.FromDateTime(inferredDateTime);
            }

            ValidateLiteral(value, index, label, allowEmpty: false);
            return ScalarValue.FromText(value);
        }

        private static string ScalarKey(ScalarValue value)
        {
            switch (value.Kind)
            {
                case ScalarValueKind.Null:
                    return "null";
                case ScalarValueKind.Text:
                    return "text|" + value.Text;
                case ScalarValueKind.Number:
                    return "number|" + value.Number.GetValueOrDefault().ToString(CultureInfo.InvariantCulture);
                case ScalarValueKind.Boolean:
                    return "boolean|" + value.Boolean.GetValueOrDefault().ToString(CultureInfo.InvariantCulture);
                case ScalarValueKind.Date:
                    return "date|" + value.Temporal.GetValueOrDefault().ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                case ScalarValueKind.DateTime:
                    return "datetime|" + value.Temporal.GetValueOrDefault().ToString("O", CultureInfo.InvariantCulture);
                default:
                    throw new InvalidOperationException("The literal type is not supported.");
            }
        }

        private static T WithId<T>(T transform, int index, string operation) where T : TransformStep
        {
            transform.Id = "manual-"
                + (index + 1).ToString("000", CultureInfo.InvariantCulture)
                + "-"
                + operation;
            return transform;
        }

        private static string RequireColumn(string raw, int index, string label)
        {
            string column = RequireText(raw, index, label, MaximumColumnNameLength);
            if (column.StartsWith("__erb_", StringComparison.OrdinalIgnoreCase))
            {
                throw Error(index, "Column names beginning with __erb_ are reserved by the report builder.");
            }

            return column;
        }

        private static string RequireText(string raw, int index, string label, int maximumLength)
        {
            string value = (raw ?? string.Empty).Trim();
            if (value.Length == 0)
            {
                throw Error(index, label + " is required.");
            }

            if (value.Length > maximumLength)
            {
                throw Error(
                    index,
                    label + " is too long. The maximum is "
                    + maximumLength.ToString(CultureInfo.InvariantCulture)
                    + " characters.");
            }

            if (value.Any(char.IsControl))
            {
                throw Error(index, label + " cannot contain control characters.");
            }

            return value;
        }

        private static void ValidateLiteral(
            string value,
            int index,
            string label,
            bool allowEmpty)
        {
            if (!allowEmpty && value.Length == 0)
            {
                throw Error(index, label + " is required.");
            }

            if (value.Length > MaximumLiteralLength)
            {
                throw Error(index, "The " + label + " can contain at most 1,024 characters.");
            }

            if (value.Any(char.IsControl))
            {
                throw Error(index, "The " + label + " cannot contain control characters.");
            }
        }

        private static void ValidateInputLengths(ManualTransformSnapshot snapshot, int index)
        {
            ValidateBounded(snapshot.Operation, index, "Operation", 80);
            ValidateBounded(snapshot.Column, index, "Column", MaximumDetailsLength);
            ValidateBounded(snapshot.OutputColumn, index, "Output column", MaximumColumnNameLength);
            ValidateBounded(snapshot.Details, index, "Details", MaximumDetailsLength);
        }

        private static void ValidateBounded(string raw, int index, string label, int maximumLength)
        {
            string value = raw ?? string.Empty;
            if (value.Length > maximumLength)
            {
                throw Error(
                    index,
                    label + " is too long. The maximum is "
                    + maximumLength.ToString(CultureInfo.InvariantCulture)
                    + " characters.");
            }

            if (value.Any(char.IsControl))
            {
                throw Error(index, label + " cannot contain control characters.");
            }
        }

        private static void EnsureBlank(
            string raw,
            int index,
            string label,
            string operation)
        {
            if (!string.IsNullOrWhiteSpace(raw))
            {
                throw Error(
                    index,
                    label + " is not used by " + Display(operation)
                    + ". Clear it so no instruction is silently ignored.");
            }
        }

        private static string NormalizeToken(string raw)
        {
            if (raw == null)
            {
                return string.Empty;
            }

            return new string(raw
                .Where(character => !char.IsWhiteSpace(character) && character != '-' && character != '_')
                .Select(char.ToLowerInvariant)
                .ToArray());
        }

        private static string Display(string value)
        {
            string result = (value ?? string.Empty).Trim();
            return result.Length == 0 ? "blank" : result;
        }

        private static InvalidOperationException Error(int index, string message)
        {
            return new InvalidOperationException(
                "Transformation "
                + (index + 1).ToString(CultureInfo.InvariantCulture)
                + ": "
                + message);
        }
    }
}
