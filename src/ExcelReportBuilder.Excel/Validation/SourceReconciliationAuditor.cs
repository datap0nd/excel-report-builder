using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using ExcelReportBuilder.Core.Measures;
using ExcelReportBuilder.Core.Periods;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Core.Transforms;

namespace ExcelReportBuilder.Excel.Validation
{
    /// <summary>
    /// Independent full-source evidence used by the mandatory no-truncation
    /// and source-total checks whenever deterministic transforms are applied.
    /// </summary>
    public sealed class SourceReconciliationAudit
    {
        public long SourceRows { get; set; }

        public long ExpectedNormalizedRows { get; set; }

        public IReadOnlyDictionary<string, long> RemovedRowsByTransform { get; set; } =
            new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);

        public IReadOnlyDictionary<string, decimal> ExpectedTotals { get; set; } =
            new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Applies the closed typed transform grammar as a small independent
    /// reference evaluator. It never executes M, formulas, code, or workbook
    /// mutations. Production calls read only the source columns needed by the
    /// transforms and auditable additive Values.
    /// </summary>
    public sealed class SourceReconciliationAuditor
    {
        private const int SourceReadBatchRows = 16384;
        private static readonly CultureInfo EnglishUnitedStates =
            CultureInfo.GetCultureInfo("en-US");
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

        public SourceReconciliationAudit AuditRange(
            object sourceRange,
            ReportSpecV1 specification,
            long expectedSourceRows,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (sourceRange == null) throw new ArgumentNullException(nameof(sourceRange));
            if (specification == null) throw new ArgumentNullException(nameof(specification));

            dynamic range = sourceRange;
            dynamic? listObject = TryGetContainingListObject(range);
            dynamic headerRange;
            dynamic dataRange;
            long actualSourceRows;
            int firstDataRow;
            int columnCount;
            if (listObject != null)
            {
                headerRange = listObject.HeaderRowRange;
                actualSourceRows = Convert.ToInt64(
                    listObject.ListRows.Count,
                    CultureInfo.InvariantCulture);
                dataRange = actualSourceRows == 0L ? range : listObject.DataBodyRange;
                firstDataRow = 1;
                columnCount = Convert.ToInt32(
                    listObject.ListColumns.Count,
                    CultureInfo.InvariantCulture);
            }
            else
            {
                var totalRows = Convert.ToInt64(range.Rows.Count, CultureInfo.InvariantCulture);
                if (totalRows < 1)
                {
                    throw new InvalidOperationException("The selected source no longer contains its header row.");
                }

                headerRange = range;
                dataRange = range;
                actualSourceRows = totalRows - 1L;
                firstDataRow = 2;
                columnCount = Convert.ToInt32(range.Columns.Count, CultureInfo.InvariantCulture);
            }

            if (actualSourceRows != expectedSourceRows)
            {
                throw new InvalidOperationException(
                    "The selected source row count changed after planning. Reinspect the Data before building.");
            }

            if (actualSourceRows > int.MaxValue)
            {
                throw new InvalidOperationException(
                    "The selected Excel range contains more rows than the bounded host auditor can index.");
            }

            var headerIndexes = new Dictionary<string, int>(StringComparer.Ordinal);
            for (var columnIndex = 1; columnIndex <= columnCount; columnIndex++)
            {
                string? header = Convert.ToString(
                    headerRange.Cells[1, columnIndex].Value2,
                    CultureInfo.InvariantCulture);
                if (!string.IsNullOrWhiteSpace(header))
                {
                    headerIndexes.Add(header!, columnIndex);
                }
            }

            var requiredColumnIndexes = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (string header in RequiredSourceColumns(specification))
            {
                if (headerIndexes.TryGetValue(header, out int columnIndex))
                {
                    requiredColumnIndexes.Add(header, columnIndex);
                }
            }

            var context = new AuditContext(specification, cancellationToken);
            for (long batchOffset = 0L; batchOffset < actualSourceRows; batchOffset += SourceReadBatchRows)
            {
                cancellationToken.ThrowIfCancellationRequested();
                int batchRows = checked((int)Math.Min(
                    SourceReadBatchRows,
                    actualSourceRows - batchOffset));
                var valuesByColumn = new Dictionary<string, object?[]>(
                    StringComparer.Ordinal);
                foreach (KeyValuePair<string, int> column in requiredColumnIndexes)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    valuesByColumn[column.Key] = ReadDataColumn(
                        dataRange,
                        checked((int)batchOffset + firstDataRow),
                        column.Value,
                        batchRows);
                }

                for (var rowIndex = 0; rowIndex < batchRows; rowIndex++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var row = new Dictionary<string, object?>(StringComparer.Ordinal);
                    foreach (var column in valuesByColumn)
                    {
                        row[column.Key] = column.Value[rowIndex];
                    }

                    context.SourceRows++;
                    Process(row, 0, context);
                }
            }

            context.ValidateTotalRowEvidence();
            return context.ToResult();
        }

        public SourceReconciliationAudit AuditRows(
            IEnumerable<IReadOnlyDictionary<string, object?>> sourceRows,
            ReportSpecV1 specification,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (sourceRows == null) throw new ArgumentNullException(nameof(sourceRows));
            if (specification == null) throw new ArgumentNullException(nameof(specification));

            var context = new AuditContext(specification, cancellationToken);
            foreach (IReadOnlyDictionary<string, object?> sourceRow in sourceRows)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (sourceRow == null)
                {
                    throw new InvalidOperationException("The source audit encountered a missing row.");
                }

                context.SourceRows++;
                Process(CloneRow(sourceRow), 0, context);
            }

            context.ValidateTotalRowEvidence();
            return context.ToResult();
        }

        private static void Process(
            Dictionary<string, object?> row,
            int transformIndex,
            AuditContext context)
        {
            context.CancellationToken.ThrowIfCancellationRequested();
            if (transformIndex >= context.Specification.Transforms.Count)
            {
                context.Accept(row);
                return;
            }

            TransformStep transform = context.Specification.Transforms[transformIndex];
            switch (transform)
            {
                case SelectColumnsTransform select:
                    KeepOnly(row, select.Columns);
                    break;
                case KeepColumnsTransform keep:
                    KeepOnly(row, keep.Columns);
                    break;
                case RemoveColumnsTransform remove:
                    foreach (string column in remove.Columns)
                    {
                        row.Remove(column);
                    }
                    break;
                case ReorderColumnsTransform _:
                    break;
                case RenameColumnTransform rename:
                    Rename(row, rename.From, rename.To);
                    break;
                case ChangeColumnTypeTransform change:
                    row[change.Column] = ConvertColumnValue(
                        RequiredValue(row, change.Column),
                        change.DataType,
                        change.Id);
                    break;
                case TrimTextTransform trim:
                    foreach (string column in trim.Columns)
                    {
                        object? value = RequiredValue(row, column);
                        row[column] = IsNull(value) || IsError(value)
                            ? value
                            : TextFrom(value!).Trim();
                    }
                    break;
                case ReplaceValueTransform replace:
                    object? replaceValue = RequiredValue(row, replace.Column);
                    if (!IsError(replaceValue)
                        && ValuesEqual(replaceValue, ScalarObject(replace.Find)))
                    {
                        row[replace.Column] = ScalarObject(replace.ReplaceWith);
                    }
                    break;
                case NormalizeBlanksTransform blanks:
                    foreach (string column in blanks.Columns)
                    {
                        object? value = RequiredValue(row, column);
                        if (IsNull(value) || value is string text &&
                            (blanks.TreatWhitespaceAsBlank
                                ? string.IsNullOrWhiteSpace(text)
                                : text.Length == 0))
                        {
                            row[column] = ScalarObject(blanks.Replacement);
                        }
                    }
                    break;
                case NormalizeErrorsTransform errors:
                    foreach (string column in errors.Columns)
                    {
                        if (IsError(RequiredValue(row, column)))
                        {
                            row[column] = ScalarObject(errors.Replacement);
                        }
                    }
                    break;
                case FillDownTransform fill:
                    ApplyFillDown(row, transformIndex, fill, context);
                    break;
                case MapValuesTransform map:
                    ApplyMap(row, map);
                    break;
                case FilterRowsTransform filter:
                    if (!MatchesFilter(RequiredValue(row, filter.Column), filter))
                    {
                        context.RecordRemoval(filter.Id);
                        return;
                    }
                    break;
                case ExcludeTotalRowsTransform exclusion:
                    if (context.MatchesAndRecordsEvidence(row, exclusion))
                    {
                        context.RecordRemoval(exclusion.Id);
                        return;
                    }
                    break;
                case DerivePeriodPartsTransform derive:
                    ApplyPeriodParts(row, derive);
                    break;
                case AddArithmeticColumnTransform arithmetic:
                    row[arithmetic.OutputColumn] = ApplyArithmetic(row, arithmetic);
                    break;
                case NormalizePeriodsTransform normalize:
                    foreach (Dictionary<string, object?> normalized in
                             NormalizePeriods(row, normalize, context.Specification.PeriodMapping))
                    {
                        Process(normalized, transformIndex + 1, context);
                    }
                    return;
                default:
                    throw new InvalidOperationException(
                        "The source audit does not support transform '" + transform.Kind + "'.");
            }

            Process(row, transformIndex + 1, context);
        }

        private static void KeepOnly(
            Dictionary<string, object?> row,
            IEnumerable<string> columns)
        {
            var keep = new HashSet<string>(columns, StringComparer.Ordinal);
            foreach (string column in row.Keys.Where(column => !keep.Contains(column)).ToList())
            {
                row.Remove(column);
            }
        }

        private static void Rename(
            Dictionary<string, object?> row,
            string from,
            string to)
        {
            object? value = RequiredValue(row, from);
            row.Remove(from);
            row.Add(to, value);
        }

        private static void ApplyFillDown(
            Dictionary<string, object?> row,
            int transformIndex,
            FillDownTransform transform,
            AuditContext context)
        {
            Dictionary<string, object?> state = context.FillDownState(transformIndex);
            foreach (string column in transform.Columns)
            {
                object? value = RequiredValue(row, column);
                if (IsNull(value))
                {
                    if (state.TryGetValue(column, out object? previous))
                    {
                        row[column] = previous;
                    }
                }
                else
                {
                    state[column] = value;
                }
            }
        }

        private static void ApplyMap(
            Dictionary<string, object?> row,
            MapValuesTransform transform)
        {
            object? value = RequiredValue(row, transform.Column);
            if (IsError(value))
            {
                return;
            }

            foreach (ValueMapEntry entry in transform.Entries)
            {
                if (ValuesEqual(value, ScalarObject(entry.From)))
                {
                    row[transform.Column] = ScalarObject(entry.To);
                    return;
                }
            }
        }

        private static bool MatchesFilter(object? value, FilterRowsTransform transform)
        {
            ThrowIfError(value, transform.Id);
            switch (transform.Operator)
            {
                case RowFilterOperator.IsBlank:
                    return IsBlank(value);
                case RowFilterOperator.IsNotBlank:
                    return !IsBlank(value);
            }

            if (transform.Value == null)
            {
                throw new InvalidOperationException(
                    "Filter transform '" + transform.Id + "' has no comparison value.");
            }

            object? expected = ScalarObject(transform.Value);
            switch (transform.Operator)
            {
                case RowFilterOperator.Equal:
                    return ValuesEqual(value, expected);
                case RowFilterOperator.NotEqual:
                    return !ValuesEqual(value, expected);
                case RowFilterOperator.GreaterThan:
                    return !IsNull(value) && !IsNull(expected)
                        && Compare(value, expected, transform.Id) > 0;
                case RowFilterOperator.GreaterThanOrEqual:
                    return !IsNull(value) && !IsNull(expected)
                        && Compare(value, expected, transform.Id) >= 0;
                case RowFilterOperator.LessThan:
                    return !IsNull(value) && !IsNull(expected)
                        && Compare(value, expected, transform.Id) < 0;
                case RowFilterOperator.LessThanOrEqual:
                    return !IsNull(value) && !IsNull(expected)
                        && Compare(value, expected, transform.Id) <= 0;
                case RowFilterOperator.Contains:
                    return !IsNull(value) && TextFrom(value!).IndexOf(
                        TextFrom(expected!),
                        StringComparison.Ordinal) >= 0;
                case RowFilterOperator.StartsWith:
                    return !IsNull(value) && TextFrom(value!).StartsWith(
                        TextFrom(expected!),
                        StringComparison.Ordinal);
                case RowFilterOperator.EndsWith:
                    return !IsNull(value) && TextFrom(value!).EndsWith(
                        TextFrom(expected!),
                        StringComparison.Ordinal);
                default:
                    throw new InvalidOperationException(
                        "The source audit does not support filter operator '" + transform.Operator + "'.");
            }
        }

        private static void ApplyPeriodParts(
            Dictionary<string, object?> row,
            DerivePeriodPartsTransform transform)
        {
            object? raw = RequiredValue(row, transform.DateColumn);
            DateTime? date = null;
            AuditError? conversionError = null;
            if (IsError(raw))
            {
                conversionError = ToAuditError(raw!, transform.Id);
            }
            else if (!IsNull(raw))
            {
                try
                {
                    date = DateFrom(raw!, transform.Id);
                }
                catch (Exception exception) when (
                    exception is InvalidOperationException ||
                    exception is ArgumentException ||
                    exception is OverflowException)
                {
                    conversionError = new AuditError(transform.Id, exception.Message);
                }
            }

            foreach (DerivedPeriodColumnSpec column in transform.Columns)
            {
                object? value;
                if (conversionError != null)
                {
                    value = conversionError;
                }
                else if (!date.HasValue)
                {
                    value = null;
                }
                else
                {
                    switch (column.Part)
                    {
                        case DerivedPeriodPart.Year:
                            value = (long)date.Value.Year;
                            break;
                        case DerivedPeriodPart.Half:
                            value = date.Value.Month <= 6 ? "H1" : "H2";
                            break;
                        case DerivedPeriodPart.Quarter:
                            value = "Q" + (((date.Value.Month - 1) / 3) + 1)
                                .ToString(CultureInfo.InvariantCulture);
                            break;
                        case DerivedPeriodPart.MonthNumber:
                            value = (long)date.Value.Month;
                            break;
                        case DerivedPeriodPart.MonthName:
                            value = date.Value.ToString("MMMM", EnglishUnitedStates);
                            break;
                        case DerivedPeriodPart.YearMonth:
                            value = date.Value.ToString("yyyy-MM", CultureInfo.InvariantCulture);
                            break;
                        default:
                            throw new InvalidOperationException(
                                "The source audit does not support period part '" + column.Part + "'.");
                    }
                }

                row.Add(column.OutputColumn, value);
            }
        }

        private static object? ApplyArithmetic(
            IReadOnlyDictionary<string, object?> row,
            AddArithmeticColumnTransform transform)
        {
            try
            {
                decimal? left = ArithmeticValue(row, transform.Left, transform.Id);
                decimal? right = ArithmeticValue(row, transform.Right, transform.Id);
                if (!left.HasValue || !right.HasValue)
                {
                    return null;
                }

                decimal calculated;
                switch (transform.Operator)
                {
                    case ArithmeticOperator.Add:
                        calculated = checked(left.Value + right.Value);
                        break;
                    case ArithmeticOperator.Subtract:
                        calculated = checked(left.Value - right.Value);
                        break;
                    case ArithmeticOperator.Multiply:
                        calculated = checked(left.Value * right.Value);
                        break;
                    case ArithmeticOperator.Divide:
                        if (right.Value == 0m)
                        {
                            if (transform.ReturnNullOnZeroDenominator)
                            {
                                return null;
                            }

                            throw new InvalidOperationException(
                                "Arithmetic transform '" + transform.Id + "' divides by zero.");
                        }

                        calculated = left.Value / right.Value;
                        break;
                    default:
                        throw new InvalidOperationException(
                            "The source audit does not support arithmetic operator '" + transform.Operator + "'.");
                }

                if (transform.ResultType == ColumnDataType.WholeNumber)
                {
                    decimal rounded = decimal.Round(calculated, 0, MidpointRounding.ToEven);
                    return checked((long)rounded);
                }

                return calculated;
            }
            catch (Exception exception) when (
                exception is InvalidOperationException ||
                exception is OverflowException ||
                exception is DivideByZeroException)
            {
                return new AuditError(transform.Id, exception.Message);
            }
        }

        private static decimal? ArithmeticValue(
            IReadOnlyDictionary<string, object?> row,
            ArithmeticOperand operand,
            string transformId)
        {
            if (operand.Kind == ArithmeticOperandKind.Number)
            {
                return operand.Number;
            }

            object? value = RequiredValue(row, operand.Column ?? string.Empty);
            if (IsNull(value))
            {
                return null;
            }

            return NumberFrom(value!, transformId);
        }

        private static IEnumerable<Dictionary<string, object?>> NormalizePeriods(
            IReadOnlyDictionary<string, object?> row,
            NormalizePeriodsTransform transform,
            PeriodMappingSpec? mapping)
        {
            if (mapping == null || !string.Equals(
                    mapping.Id,
                    transform.PeriodMappingId,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "Period normalization references an unavailable mapping during source audit.");
            }

            if (mapping.Kind == PeriodMappingKind.LongDateColumn)
            {
                var normalized = CloneRow(row);
                string dateColumn = mapping.DateColumn ?? string.Empty;
                object? raw = RequiredValue(normalized, dateColumn);
                if (IsError(raw))
                {
                    normalized[dateColumn] = ToAuditError(raw!, transform.Id);
                    yield return normalized;
                    yield break;
                }

                try
                {
                    if (IsNull(raw))
                    {
                        throw new InvalidOperationException(
                            "A blank period cannot be normalized during source audit.");
                    }

                    DateTime period = NormalizeLongPeriodValue(raw!, mapping);
                    normalized[dateColumn] = new AuditTemporal(period, AuditTemporalKind.Date);
                }
                catch (Exception exception) when (
                    exception is ArgumentException ||
                    exception is InvalidOperationException ||
                    exception is ArgumentOutOfRangeException ||
                    exception is OverflowException)
                {
                    normalized[dateColumn] = new AuditError(transform.Id, exception.Message);
                }

                yield return normalized;
                yield break;
            }

            foreach (PeriodColumnMapping periodColumn in mapping.Columns)
            {
                var normalized = new Dictionary<string, object?>(StringComparer.Ordinal);
                foreach (string key in mapping.KeyColumns)
                {
                    normalized.Add(key, RequiredValue(row, key));
                }

                int? year = periodColumn.Year ?? mapping.ReportingYear;
                if (!year.HasValue)
                {
                    throw new InvalidOperationException(
                        "A reporting year is required during source row-count audit.");
                }

                normalized.Add(
                    mapping.PeriodColumnName,
                    new AuditTemporal(
                        new DateTime(year.Value, periodColumn.Month, 1),
                        AuditTemporalKind.Date));
                if (mapping.Kind == PeriodMappingKind.MetricMonthHeaders)
                {
                    normalized.Add(mapping.MetricColumnName, periodColumn.Metric);
                }

                normalized.Add(
                    mapping.ValueColumnName,
                    RequiredValue(row, periodColumn.SourceColumn));
                yield return normalized;
            }
        }

        private static DateTime NormalizeLongPeriodValue(
            object raw,
            PeriodMappingSpec mapping)
        {
            try
            {
                return PeriodValueNormalizer.Normalize(
                    raw,
                    mapping.ReportingYear,
                    mapping.Grain);
            }
            catch (ArgumentException)
            {
                if (!TryNumber(raw, out decimal serial))
                {
                    throw;
                }

                if (serial >= 100000m && serial <= 999999m
                    && decimal.Truncate(serial) == serial)
                {
                    throw new ArgumentException(
                        "A six-digit numeric period must be a valid YYYYMM value.",
                        nameof(raw));
                }

                return PeriodValueNormalizer.Normalize(
                    DateTime.FromOADate(Convert.ToDouble(serial, CultureInfo.InvariantCulture)),
                    mapping.ReportingYear,
                    mapping.Grain);
            }
        }

        private static object? ConvertColumnValue(
            object? value,
            ColumnDataType dataType,
            string transformId)
        {
            if (IsNull(value))
            {
                return null;
            }

            if (IsError(value))
            {
                return ToAuditError(value!, transformId);
            }

            try
            {
                switch (dataType)
                {
                    case ColumnDataType.Text:
                        if (!IsSupportedTextValue(value!))
                        {
                            throw new InvalidOperationException(
                                "Text conversion received an unsupported full-source value type.");
                        }

                        return TextFrom(value!);
                    case ColumnDataType.WholeNumber:
                        decimal wholeSource = DecimalFrom(value!, transformId);
                        decimal rounded = decimal.Round(wholeSource, 0, MidpointRounding.ToEven);
                        return checked((long)rounded);
                    case ColumnDataType.DecimalNumber:
                        return DecimalFrom(value!, transformId);
                    case ColumnDataType.Boolean:
                        return BooleanFrom(value!, transformId);
                    case ColumnDataType.Date:
                        return new AuditTemporal(
                            TypedTemporalFrom(value!, dateTime: false, transformId: transformId).Date,
                            AuditTemporalKind.Date);
                    case ColumnDataType.DateTime:
                        return new AuditTemporal(
                            TypedTemporalFrom(value!, dateTime: true, transformId: transformId),
                            AuditTemporalKind.DateTime);
                    default:
                        throw new InvalidOperationException(
                            "The source audit does not support column type '" + dataType + "'.");
                }
            }
            catch (Exception exception) when (
                exception is FormatException ||
                exception is InvalidCastException ||
                exception is OverflowException ||
                exception is ArgumentException ||
                exception is InvalidOperationException)
            {
                return new AuditError(
                    transformId,
                    "The full-source value cannot be converted to " + dataType + ". " +
                    exception.Message);
            }
        }

        private static int Compare(object? left, object? right, string transformId)
        {
            if (IsNull(left) || IsNull(right))
            {
                throw new InvalidOperationException(
                    "Filter transform '" + transformId + "' cannot order a blank value.");
            }

            ThrowIfError(left, transformId);
            ThrowIfError(right, transformId);
            if (TryNumber(left, out decimal leftNumber) &&
                TryNumber(right, out decimal rightNumber))
            {
                return leftNumber.CompareTo(rightNumber);
            }

            if (TryTemporal(left, out AuditTemporal leftDate)
                && TryTemporal(right, out AuditTemporal rightDate)
                && leftDate.Kind == rightDate.Kind)
            {
                return leftDate.Value.CompareTo(rightDate.Value);
            }

            if (left is string leftText && right is string rightText)
            {
                return string.Compare(leftText, rightText, StringComparison.Ordinal);
            }

            throw new InvalidOperationException(
                "Filter transform '" + transformId +
                "' compares incompatible full-source value types.");
        }

        private static bool ValuesEqual(object? left, object? right)
        {
            if (IsNull(left) || IsNull(right))
            {
                return IsNull(left) && IsNull(right);
            }

            if (TryNumber(left, out decimal leftNumber) &&
                TryNumber(right, out decimal rightNumber))
            {
                return leftNumber == rightNumber;
            }

            if (TryTemporal(left, out AuditTemporal leftDate)
                && TryTemporal(right, out AuditTemporal rightDate))
            {
                return leftDate.Kind == rightDate.Kind
                    && leftDate.Value == rightDate.Value;
            }

            return left!.GetType() == right!.GetType() && left.Equals(right);
        }

        private static bool IsBlank(object? value)
        {
            return IsNull(value) || value is string text && string.IsNullOrWhiteSpace(text);
        }

        private static bool IsNull(object? value)
        {
            return value == null || value == DBNull.Value;
        }

        private static void ThrowIfError(object? value, string transformId)
        {
            if (value is AuditError auditError)
            {
                throw new InvalidOperationException(
                    "Transform '" + auditError.TransformId + "' produced an error before it was normalized. " +
                    auditError.Message);
            }

            if (value is ErrorWrapper)
            {
                throw new InvalidOperationException(
                    "Transform '" + transformId +
                    "' encountered an Excel error before it was normalized.");
            }
        }

        private static object? RequiredValue(
            IReadOnlyDictionary<string, object?> row,
            string column)
        {
            if (row.TryGetValue(column, out object? value))
            {
                return value;
            }

            throw new InvalidOperationException(
                "The full-source audit cannot resolve column '" + column + "'.");
        }

        private static object? ScalarObject(ScalarValue value)
        {
            switch (value.Kind)
            {
                case ScalarValueKind.Null:
                    return null;
                case ScalarValueKind.Text:
                    return value.Text;
                case ScalarValueKind.Number:
                    return value.Number;
                case ScalarValueKind.Boolean:
                    return value.Boolean;
                case ScalarValueKind.Date:
                    return value.Temporal.HasValue
                        ? new AuditTemporal(value.Temporal.Value.Date, AuditTemporalKind.Date)
                        : null;
                case ScalarValueKind.DateTime:
                    return value.Temporal.HasValue
                        ? new AuditTemporal(value.Temporal.Value, AuditTemporalKind.DateTime)
                        : null;
                default:
                    throw new InvalidOperationException(
                        "The source audit does not support scalar kind '" + value.Kind + "'.");
            }
        }

        private static decimal NumberFrom(object value, string transformId)
        {
            return DecimalFrom(value, transformId);
        }

        private static decimal DecimalFrom(object value, string transformId)
        {
            ThrowIfError(value, transformId);
            if (TryNumber(value, out decimal number))
            {
                return number;
            }

            if (value is string text)
            {
                string trimmed = text.Trim();
                if (trimmed.Length > 0
                    && trimmed.All(IsArithmeticTextCharacter)
                    && decimal.TryParse(
                        trimmed,
                        NumberStyles.Float | NumberStyles.AllowThousands,
                        EnglishUnitedStates,
                        out number))
                {
                    return number;
                }
            }

            throw new InvalidOperationException(
                "Transform '" + transformId + "' requires a numeric full-source value.");
        }

        private static bool IsArithmeticTextCharacter(char value)
        {
            return value >= '0' && value <= '9'
                || value == '+'
                || value == '-'
                || value == '.'
                || value == ','
                || value == 'e'
                || value == 'E';
        }

        private static bool BooleanFrom(object value, string transformId)
        {
            ThrowIfError(value, transformId);
            if (value is bool boolean)
            {
                return boolean;
            }

            if (TryNumber(value, out decimal number))
            {
                return number != 0m;
            }

            if (value is string text)
            {
                if (string.Equals(text.Trim(), "true", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                if (string.Equals(text.Trim(), "false", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            throw new InvalidOperationException(
                "Transform '" + transformId + "' requires a logical full-source value.");
        }

        private static bool TryNumber(object? value, out decimal number)
        {
            number = 0m;
            if (IsNull(value) || value is bool || value is string || value is DateTime ||
                value is DateTimeOffset || value is AuditTemporal || IsError(value))
            {
                return false;
            }

            try
            {
                number = Convert.ToDecimal(value, CultureInfo.InvariantCulture);
                return true;
            }
            catch (Exception exception) when (
                exception is FormatException ||
                exception is InvalidCastException ||
                exception is OverflowException)
            {
                return false;
            }
        }

        private static bool IsSupportedTextValue(object value)
        {
            return value is string
                || value is bool
                || value is DateTime
                || value is AuditTemporal
                || value is byte
                || value is sbyte
                || value is short
                || value is ushort
                || value is int
                || value is uint
                || value is long
                || value is ulong
                || value is float
                || value is double
                || value is decimal;
        }

        private static DateTime TypedTemporalFrom(
            object value,
            bool dateTime,
            string transformId)
        {
            ThrowIfError(value, transformId);
            if (value is AuditTemporal temporal)
            {
                return temporal.Value;
            }

            if (value is DateTime date)
            {
                return date;
            }

            if (!dateTime && value is DateTimeOffset offset)
            {
                return offset.DateTime;
            }

            if (TryNumber(value, out decimal serial))
            {
                return DateTime.FromOADate(Convert.ToDouble(serial, CultureInfo.InvariantCulture));
            }

            if (value is string text)
            {
                string[] formats = dateTime
                    ? BoundedDateTimeFormats.Concat(BoundedDateFormats).ToArray()
                    : BoundedDateFormats;
                if (DateTime.TryParseExact(
                        text.Trim(),
                        formats,
                        EnglishUnitedStates,
                        DateTimeStyles.None,
                        out date))
                {
                    return date;
                }
            }

            throw new InvalidOperationException(
                "Transform '" + transformId + "' requires a supported " +
                (dateTime ? "date-time" : "date") + " full-source value.");
        }

        private static DateTime DateFrom(object value, string transformId)
        {
            ThrowIfError(value, transformId);
            if (value is AuditTemporal temporal)
            {
                return temporal.Value;
            }

            if (value is DateTime date)
            {
                return date;
            }

            if (value is DateTimeOffset offset)
            {
                return offset.DateTime;
            }

            if (TryNumber(value, out decimal serial))
            {
                return DateTime.FromOADate(Convert.ToDouble(serial, CultureInfo.InvariantCulture));
            }

            if (value is string text && DateTime.TryParse(
                    text,
                    EnglishUnitedStates,
                    DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.RoundtripKind,
                    out date))
            {
                return date;
            }

            throw new InvalidOperationException(
                "Transform '" + transformId + "' requires a date full-source value.");
        }

        private static string TextFrom(object value)
        {
            if (value is AuditTemporal temporal)
            {
                return temporal.Kind == AuditTemporalKind.Date
                    ? temporal.Value.ToString("M/d/yyyy", EnglishUnitedStates)
                    : temporal.Value.ToString("M/d/yyyy h:mm:ss tt", EnglishUnitedStates);
            }

            if (value is DateTime date)
            {
                return date.ToString("M/d/yyyy h:mm:ss tt", EnglishUnitedStates);
            }

            if (value is bool boolean)
            {
                return boolean ? "true" : "false";
            }

            return Convert.ToString(value, EnglishUnitedStates) ?? string.Empty;
        }

        private static bool TryTemporal(object? value, out AuditTemporal temporal)
        {
            if (value is AuditTemporal typed)
            {
                temporal = typed;
                return true;
            }

            if (value is DateTime dateTime)
            {
                temporal = new AuditTemporal(dateTime, AuditTemporalKind.DateTime);
                return true;
            }

            if (value is DateTimeOffset offset)
            {
                temporal = new AuditTemporal(offset.DateTime, AuditTemporalKind.DateTime);
                return true;
            }

            temporal = null!;
            return false;
        }

        private static bool IsError(object? value)
        {
            return value is ErrorWrapper || value is AuditError;
        }

        private static AuditError ToAuditError(object value, string transformId)
        {
            return value as AuditError
                ?? new AuditError(transformId, "The source value is an Excel error.");
        }

        private static object?[] ReadDataColumn(
            dynamic sourceRange,
            int firstSourceRow,
            int columnIndex,
            int rowCount)
        {
            var result = new object?[rowCount];
            if (rowCount == 0)
            {
                return result;
            }

            dynamic first = sourceRange.Cells[firstSourceRow, columnIndex];
            dynamic last = sourceRange.Cells[firstSourceRow + rowCount - 1, columnIndex];
            dynamic dataRange = sourceRange.Worksheet.Range[first, last];
            object? values = dataRange.Value2;
            if (values is Array array && array.Rank == 2)
            {
                int lowerRow = array.GetLowerBound(0);
                int lowerColumn = array.GetLowerBound(1);
                for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
                {
                    result[rowIndex] = array.GetValue(lowerRow + rowIndex, lowerColumn);
                }
            }
            else if (rowCount == 1)
            {
                result[0] = values;
            }
            else
            {
                throw new InvalidOperationException(
                    "Excel did not return the complete source column during independent audit.");
            }

            return result;
        }

        private static dynamic? TryGetContainingListObject(dynamic range)
        {
            try
            {
                return range.ListObject;
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    "Excel could not determine whether the audited source is a table or named range.",
                    exception);
            }
        }

        private static Dictionary<string, object?> CloneRow(
            IReadOnlyDictionary<string, object?> source)
        {
            var result = new Dictionary<string, object?>(StringComparer.Ordinal);
            foreach (KeyValuePair<string, object?> value in source)
            {
                result.Add(value.Key, value.Value);
            }

            return result;
        }

        private static HashSet<string> RequiredSourceColumns(ReportSpecV1 specification)
        {
            var result = new HashSet<string>(StringComparer.Ordinal);
            if (specification.PeriodMapping != null)
            {
                Add(result, specification.PeriodMapping.DateColumn);
                foreach (string key in specification.PeriodMapping.KeyColumns)
                {
                    Add(result, key);
                }

                foreach (PeriodColumnMapping column in specification.PeriodMapping.Columns)
                {
                    Add(result, column.SourceColumn);
                }
            }

            foreach (MeasureDefinition measure in specification.Measures)
            {
                if (measure.Expression is AggregateMeasureExpression aggregate &&
                    aggregate.Function == AggregateFunction.Sum &&
                    string.IsNullOrWhiteSpace(aggregate.PeriodSliceId))
                {
                    Add(result, aggregate.Field);
                }
            }

            foreach (TransformStep transform in specification.Transforms)
            {
                switch (transform)
                {
                    case RenameColumnTransform rename:
                        Add(result, rename.From);
                        Add(result, rename.To);
                        break;
                    case ChangeColumnTypeTransform change:
                        Add(result, change.Column);
                        break;
                    case TrimTextTransform trim:
                        Add(result, trim.Columns);
                        break;
                    case ReplaceValueTransform replace:
                        Add(result, replace.Column);
                        break;
                    case NormalizeBlanksTransform blanks:
                        Add(result, blanks.Columns);
                        break;
                    case NormalizeErrorsTransform errors:
                        Add(result, errors.Columns);
                        break;
                    case FillDownTransform fill:
                        Add(result, fill.Columns);
                        break;
                    case MapValuesTransform map:
                        Add(result, map.Column);
                        break;
                    case FilterRowsTransform filter:
                        Add(result, filter.Column);
                        break;
                    case ExcludeTotalRowsTransform exclusion:
                        foreach (TotalRowEvidenceSpec evidence in exclusion.Evidence)
                        {
                            Add(result, evidence.Column);
                        }
                        break;
                    case DerivePeriodPartsTransform derive:
                        Add(result, derive.DateColumn);
                        foreach (DerivedPeriodColumnSpec column in derive.Columns)
                        {
                            Add(result, column.OutputColumn);
                        }
                        break;
                    case AddArithmeticColumnTransform arithmetic:
                        Add(result, arithmetic.OutputColumn);
                        Add(result, arithmetic.Left.Column);
                        Add(result, arithmetic.Right.Column);
                        break;
                }
            }

            return result;
        }

        private static void Add(ISet<string> target, IEnumerable<string> values)
        {
            foreach (string value in values)
            {
                Add(target, value);
            }
        }

        private static void Add(ISet<string> target, string? value)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                target.Add(value!);
            }
        }

        private sealed class AuditContext
        {
            private readonly Dictionary<int, Dictionary<string, object?>> fillDown =
                new Dictionary<int, Dictionary<string, object?>>();
            private readonly Dictionary<TotalRowEvidenceSpec, long> evidenceMatches =
                new Dictionary<TotalRowEvidenceSpec, long>();
            private readonly List<AuditableTotal> auditableTotals = new List<AuditableTotal>();

            public AuditContext(ReportSpecV1 specification, CancellationToken cancellationToken)
            {
                Specification = specification;
                CancellationToken = cancellationToken;
                foreach (MeasureDefinition measure in specification.Measures)
                {
                    if (measure.Expression is AggregateMeasureExpression aggregate &&
                        aggregate.Function == AggregateFunction.Sum &&
                        string.IsNullOrWhiteSpace(aggregate.PeriodSliceId))
                    {
                        auditableTotals.Add(new AuditableTotal(measure.Id, aggregate.Field));
                        ExpectedTotals[measure.Id] = 0m;
                    }
                }
            }

            public ReportSpecV1 Specification { get; }

            public CancellationToken CancellationToken { get; }

            public long SourceRows { get; set; }

            public long ExpectedNormalizedRows { get; private set; }

            public Dictionary<string, long> RemovedRowsByTransform { get; } =
                new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);

            public Dictionary<string, decimal> ExpectedTotals { get; } =
                new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);

            public Dictionary<string, object?> FillDownState(int transformIndex)
            {
                if (!fillDown.TryGetValue(transformIndex, out Dictionary<string, object?>? state))
                {
                    state = new Dictionary<string, object?>(StringComparer.Ordinal);
                    fillDown.Add(transformIndex, state);
                }

                return state;
            }

            public void RecordRemoval(string transformId)
            {
                RemovedRowsByTransform[transformId] =
                    RemovedRowsByTransform.TryGetValue(transformId, out long count)
                        ? checked(count + 1L)
                        : 1L;
            }

            public bool MatchesAndRecordsEvidence(
                IReadOnlyDictionary<string, object?> row,
                ExcludeTotalRowsTransform transform)
            {
                foreach (TotalRowEvidenceSpec evidence in transform.Evidence)
                {
                    object? value = RequiredValue(row, evidence.Column);
                    if (!IsError(value) && MatchesEvidence(value, evidence))
                    {
                        evidenceMatches[evidence] = evidenceMatches.TryGetValue(evidence, out long count)
                            ? checked(count + 1L)
                            : 1L;
                    }
                }

                foreach (TotalRowEvidenceSpec evidence in transform.Evidence)
                {
                    bool matched = MatchesEvidence(RequiredValue(row, evidence.Column), evidence);
                    if (transform.RequireAllEvidence && !matched)
                    {
                        return false;
                    }

                    if (!transform.RequireAllEvidence && matched)
                    {
                        return true;
                    }
                }

                return transform.RequireAllEvidence;
            }

            public void Accept(IReadOnlyDictionary<string, object?> row)
            {
                ExpectedNormalizedRows = checked(ExpectedNormalizedRows + 1L);
                foreach (AuditableTotal total in auditableTotals)
                {
                    object? value = RequiredValue(row, total.Field);
                    ThrowIfError(value, "source-total-" + total.MeasureId);
                    if (IsNull(value))
                    {
                        continue;
                    }

                    if (!TryNumber(value, out decimal number))
                    {
                        throw new InvalidOperationException(
                            "Additive source total '" + total.MeasureId +
                            "' encountered a nonnumeric value in column '" + total.Field + "'.");
                    }

                    ExpectedTotals[total.MeasureId] = checked(
                        ExpectedTotals[total.MeasureId] + number);
                }
            }

            public void ValidateTotalRowEvidence()
            {
                foreach (ExcludeTotalRowsTransform transform in
                         Specification.Transforms.OfType<ExcludeTotalRowsTransform>())
                {
                    foreach (TotalRowEvidenceSpec evidence in transform.Evidence)
                    {
                        long actual = evidenceMatches.TryGetValue(evidence, out long count) ? count : 0L;
                        if (actual != evidence.ObservedMatchCount)
                        {
                            throw new InvalidOperationException(
                                "Total-row evidence for transform '" + transform.Id +
                                "' expected " + evidence.ObservedMatchCount.ToString(CultureInfo.InvariantCulture) +
                                " matching rows but the current full source contains " +
                                actual.ToString(CultureInfo.InvariantCulture) + ". Reconfirm the source evidence.");
                        }
                    }

                    if (!RemovedRowsByTransform.TryGetValue(transform.Id, out long removed) || removed == 0L)
                    {
                        throw new InvalidOperationException(
                            "Total-row evidence for transform '" + transform.Id +
                            "' did not exclude any current source row.");
                    }
                }
            }

            public SourceReconciliationAudit ToResult()
            {
                return new SourceReconciliationAudit
                {
                    SourceRows = SourceRows,
                    ExpectedNormalizedRows = ExpectedNormalizedRows,
                    RemovedRowsByTransform = new Dictionary<string, long>(
                        RemovedRowsByTransform,
                        StringComparer.OrdinalIgnoreCase),
                    ExpectedTotals = new Dictionary<string, decimal>(
                        ExpectedTotals,
                        StringComparer.OrdinalIgnoreCase)
                };
            }

            private static bool MatchesEvidence(object? value, TotalRowEvidenceSpec evidence)
            {
                ThrowIfError(value, "total-row-evidence");
                if (evidence.MatchKind == TotalRowMatchKind.IsBlank)
                {
                    return IsBlank(value);
                }

                foreach (ScalarValue expectedValue in evidence.Values)
                {
                    object? expected = ScalarObject(expectedValue);
                    switch (evidence.MatchKind)
                    {
                        case TotalRowMatchKind.EqualsAny when ValuesEqual(value, expected):
                            return true;
                        case TotalRowMatchKind.StartsWith when !IsNull(value) &&
                            TextFrom(value!).StartsWith(
                                TextFrom(expected!),
                                StringComparison.Ordinal):
                            return true;
                        case TotalRowMatchKind.Contains when !IsNull(value) &&
                            TextFrom(value!).IndexOf(
                                TextFrom(expected!),
                                StringComparison.Ordinal) >= 0:
                            return true;
                    }
                }

                return false;
            }
        }

        private sealed class AuditableTotal
        {
            public AuditableTotal(string measureId, string field)
            {
                MeasureId = measureId;
                Field = field;
            }

            public string MeasureId { get; }

            public string Field { get; }
        }

        private enum AuditTemporalKind
        {
            Date,
            DateTime
        }

        private sealed class AuditTemporal
        {
            public AuditTemporal(DateTime value, AuditTemporalKind kind)
            {
                Value = value;
                Kind = kind;
            }

            public DateTime Value { get; }

            public AuditTemporalKind Kind { get; }
        }

        private sealed class AuditError
        {
            public AuditError(string transformId, string message)
            {
                TransformId = transformId;
                Message = message;
            }

            public string TransformId { get; }

            public string Message { get; }
        }
    }
}
