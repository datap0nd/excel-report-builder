using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelReportBuilder.Core.Periods;
using ExcelReportBuilder.Core.Specifications;

namespace ExcelReportBuilder.Core.Profiling
{
    public static class SourceProfiler
    {
        private static readonly string[] DateFormats =
        {
            "yyyy-MM-dd",
            "yyyy/MM/dd",
            "d-MMM-yyyy",
            "dd-MMM-yyyy",
            "MMM d yyyy",
            "MMMM d yyyy"
        };

        private static readonly string[] DateTimeFormats =
        {
            "yyyy-MM-dd HH:mm:ss",
            "yyyy-MM-ddTHH:mm:ss",
            "yyyy-MM-ddTHH:mm:ss.FFFFFFF",
            "yyyy-MM-ddTHH:mm:ssK",
            "yyyy-MM-ddTHH:mm:ss.FFFFFFFK"
        };

        public static SourceProfile Profile(
            IReadOnlyList<string> headers,
            IReadOnlyList<object?[]> rows)
        {
            if (headers == null)
            {
                throw new ArgumentNullException(nameof(headers));
            }

            if (rows == null)
            {
                throw new ArgumentNullException(nameof(rows));
            }

            var profile = new SourceProfile
            {
                RowCount = rows.Count,
                ColumnCount = headers.Count
            };

            var accumulators = new ColumnAccumulator[headers.Count];
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var columnIndex = 0; columnIndex < headers.Count; columnIndex++)
            {
                var header = headers[columnIndex] ?? string.Empty;
                accumulators[columnIndex] = new ColumnAccumulator(columnIndex, header);

                if (string.IsNullOrWhiteSpace(header))
                {
                    profile.Issues.Add(new SourceProfileIssue
                    {
                        Code = SourceProfileIssueCode.BlankHeader,
                        ColumnIndex = columnIndex,
                        Message = "Column " + (columnIndex + 1).ToString(CultureInfo.InvariantCulture) + " has no header."
                    });
                }
                else if (!names.Add(header))
                {
                    profile.Issues.Add(new SourceProfileIssue
                    {
                        Code = SourceProfileIssueCode.DuplicateHeader,
                        ColumnIndex = columnIndex,
                        Message = "The header '" + header + "' appears more than once."
                    });
                }
            }

            for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                var row = rows[rowIndex];
                if (row == null || row.Length != headers.Count)
                {
                    profile.Issues.Add(new SourceProfileIssue
                    {
                        Code = SourceProfileIssueCode.RaggedRow,
                        RowIndex = rowIndex,
                        Message = "Row " + (rowIndex + 1).ToString(CultureInfo.InvariantCulture)
                            + " has a different number of cells than the header row."
                    });
                }

                for (var columnIndex = 0; columnIndex < headers.Count; columnIndex++)
                {
                    var value = row != null && columnIndex < row.Length ? row[columnIndex] : null;
                    accumulators[columnIndex].Observe(value);
                }
            }

            for (var columnIndex = 0; columnIndex < accumulators.Length; columnIndex++)
            {
                profile.Columns.Add(accumulators[columnIndex].Build());
            }

            return profile;
        }

        private enum ObservedType
        {
            Text,
            WholeNumber,
            DecimalNumber,
            Boolean,
            Date,
            DateTime
        }

        private sealed class ColumnAccumulator
        {
            private readonly int _index;
            private readonly string _name;
            private readonly Dictionary<ObservedType, long> _typeCounts = new Dictionary<ObservedType, long>();
            private readonly HashSet<string> _distinct = new HashSet<string>(StringComparer.Ordinal);
            private long _blankCount;
            private long _nonBlankCount;
            private long _dateLikeCount;
            private long _periodLikeWithoutYearCount;
            private long _dayGrainCount;
            private long _monthGrainCount;
            private long _quarterGrainCount;
            private long _numericCount;
            private DateTime? _minimumDate;
            private DateTime? _maximumDate;
            private decimal? _minimumNumber;
            private decimal? _maximumNumber;

            public ColumnAccumulator(int index, string name)
            {
                _index = index;
                _name = name;
            }

            public void Observe(object? value)
            {
                if (value == null || value == DBNull.Value || (value is string text && string.IsNullOrWhiteSpace(text)))
                {
                    _blankCount++;
                    return;
                }

                _nonBlankCount++;
                DateTime temporal;
                decimal number;
                ObservedType observedType;

                if (TryGetDate(value, out temporal, out observedType))
                {
                    _dateLikeCount++;
                    ObserveGrain(InferDateGrain(temporal));
                    ObserveDate(temporal);
                    AddType(observedType);
                    _distinct.Add("d:" + temporal.ToString("O", CultureInfo.InvariantCulture));
                    return;
                }

                ParsedPeriodToken parsedPeriod;
                if (PeriodTextParser.TryParseWholeValue(value, out parsedPeriod))
                {
                    ObserveGrain(parsedPeriod.Grain);
                    if (parsedPeriod.RequiresReportingYear)
                    {
                        _periodLikeWithoutYearCount++;
                        AddType(ObservedType.Text);
                        _distinct.Add(
                            "p?:" + Convert.ToString(value, CultureInfo.InvariantCulture));
                    }
                    else
                    {
                        temporal = parsedPeriod.CanonicalPeriod!.Value;
                        _dateLikeCount++;
                        ObserveDate(temporal);
                        AddType(ObservedType.Date);
                        _distinct.Add(
                            "p:" + parsedPeriod.Grain.ToString() + ":" +
                            temporal.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
                    }

                    return;
                }

                if (TryGetNumber(value, out number, out observedType))
                {
                    _numericCount++;
                    ObserveNumber(number);
                    AddType(observedType);
                    _distinct.Add("n:" + number.ToString(CultureInfo.InvariantCulture));
                    return;
                }

                if (value is bool boolean)
                {
                    AddType(ObservedType.Boolean);
                    _distinct.Add("b:" + (boolean ? "1" : "0"));
                    return;
                }

                var normalized = Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
                bool parsedBoolean;
                if (bool.TryParse(normalized, out parsedBoolean))
                {
                    AddType(ObservedType.Boolean);
                    _distinct.Add("b:" + (parsedBoolean ? "1" : "0"));
                    return;
                }

                AddType(ObservedType.Text);
                _distinct.Add("t:" + normalized);
            }

            public SourceColumnProfile Build()
            {
                return new SourceColumnProfile
                {
                    Index = _index,
                    Name = _name,
                    InferredType = InferType(),
                    BlankCount = _blankCount,
                    NonBlankCount = _nonBlankCount,
                    DistinctCount = _distinct.Count,
                    DateLikeCount = _dateLikeCount,
                    PeriodLikeWithoutYearCount = _periodLikeWithoutYearCount,
                    DayGrainCount = _dayGrainCount,
                    MonthGrainCount = _monthGrainCount,
                    QuarterGrainCount = _quarterGrainCount,
                    NumericCount = _numericCount,
                    MinimumDate = _minimumDate,
                    MaximumDate = _maximumDate,
                    MinimumNumber = _minimumNumber,
                    MaximumNumber = _maximumNumber
                };
            }

            private void AddType(ObservedType observedType)
            {
                long count;
                _typeCounts.TryGetValue(observedType, out count);
                _typeCounts[observedType] = count + 1;
            }

            private SourceValueType InferType()
            {
                if (_nonBlankCount == 0)
                {
                    return SourceValueType.Empty;
                }

                var hasWhole = _typeCounts.ContainsKey(ObservedType.WholeNumber);
                var hasDecimal = _typeCounts.ContainsKey(ObservedType.DecimalNumber);
                if (_typeCounts.Count == 1)
                {
                    foreach (var pair in _typeCounts)
                    {
                        return Map(pair.Key);
                    }
                }

                if (_typeCounts.Count == 2 && hasWhole && hasDecimal)
                {
                    return SourceValueType.DecimalNumber;
                }

                var hasDate = _typeCounts.ContainsKey(ObservedType.Date);
                var hasDateTime = _typeCounts.ContainsKey(ObservedType.DateTime);
                if (_typeCounts.Count == 2 && hasDate && hasDateTime)
                {
                    return SourceValueType.DateTime;
                }

                return SourceValueType.Mixed;
            }

            private static SourceValueType Map(ObservedType type)
            {
                switch (type)
                {
                    case ObservedType.Text:
                        return SourceValueType.Text;
                    case ObservedType.WholeNumber:
                        return SourceValueType.WholeNumber;
                    case ObservedType.DecimalNumber:
                        return SourceValueType.DecimalNumber;
                    case ObservedType.Boolean:
                        return SourceValueType.Boolean;
                    case ObservedType.Date:
                        return SourceValueType.Date;
                    case ObservedType.DateTime:
                        return SourceValueType.DateTime;
                    default:
                        return SourceValueType.Mixed;
                }
            }

            private void ObserveDate(DateTime value)
            {
                if (!_minimumDate.HasValue || value < _minimumDate.Value)
                {
                    _minimumDate = value;
                }

                if (!_maximumDate.HasValue || value > _maximumDate.Value)
                {
                    _maximumDate = value;
                }
            }

            private void ObserveGrain(PeriodGrain grain)
            {
                switch (grain)
                {
                    case PeriodGrain.Day:
                        _dayGrainCount++;
                        break;
                    case PeriodGrain.Month:
                        _monthGrainCount++;
                        break;
                    case PeriodGrain.Quarter:
                        _quarterGrainCount++;
                        break;
                    default:
                        throw new InvalidOperationException("Unsupported period grain.");
                }
            }

            private static PeriodGrain InferDateGrain(DateTime value)
            {
                return value.Day == 1 && value.TimeOfDay == TimeSpan.Zero
                    ? PeriodGrain.Month
                    : PeriodGrain.Day;
            }

            private void ObserveNumber(decimal value)
            {
                if (!_minimumNumber.HasValue || value < _minimumNumber.Value)
                {
                    _minimumNumber = value;
                }

                if (!_maximumNumber.HasValue || value > _maximumNumber.Value)
                {
                    _maximumNumber = value;
                }
            }
        }

        private static bool TryGetDate(object value, out DateTime result, out ObservedType observedType)
        {
            if (value is DateTime dateTime)
            {
                result = dateTime;
                observedType = dateTime.TimeOfDay == TimeSpan.Zero ? ObservedType.Date : ObservedType.DateTime;
                return true;
            }

            if (value is DateTimeOffset dateTimeOffset)
            {
                result = dateTimeOffset.DateTime;
                observedType = result.TimeOfDay == TimeSpan.Zero ? ObservedType.Date : ObservedType.DateTime;
                return true;
            }

            var text = value as string;
            if (text != null)
            {
                if (DateTime.TryParseExact(
                    text.Trim(),
                    DateFormats,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AllowWhiteSpaces,
                    out result))
                {
                    observedType = ObservedType.Date;
                    return true;
                }

                if (DateTime.TryParseExact(
                    text.Trim(),
                    DateTimeFormats,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.RoundtripKind,
                    out result))
                {
                    observedType = result.TimeOfDay == TimeSpan.Zero ? ObservedType.Date : ObservedType.DateTime;
                    return true;
                }
            }

            result = default(DateTime);
            observedType = default(ObservedType);
            return false;
        }

        private static bool TryGetNumber(object value, out decimal result, out ObservedType observedType)
        {
            if (value is byte || value is sbyte || value is short || value is ushort
                || value is int || value is uint || value is long || value is ulong)
            {
                try
                {
                    result = Convert.ToDecimal(value, CultureInfo.InvariantCulture);
                    observedType = ObservedType.WholeNumber;
                    return true;
                }
                catch (OverflowException)
                {
                    result = default(decimal);
                    observedType = default(ObservedType);
                    return false;
                }
            }

            if (value is decimal || value is double || value is float)
            {
                try
                {
                    result = Convert.ToDecimal(value, CultureInfo.InvariantCulture);
                    observedType = decimal.Truncate(result) == result
                        ? ObservedType.WholeNumber
                        : ObservedType.DecimalNumber;
                    return true;
                }
                catch (Exception exception) when (exception is OverflowException || exception is FormatException)
                {
                    result = default(decimal);
                    observedType = default(ObservedType);
                    return false;
                }
            }

            var text = value as string;
            if (text != null && decimal.TryParse(
                text.Trim(),
                NumberStyles.Number | NumberStyles.AllowExponent,
                CultureInfo.InvariantCulture,
                out result))
            {
                observedType = decimal.Truncate(result) == result
                    ? ObservedType.WholeNumber
                    : ObservedType.DecimalNumber;
                return true;
            }

            result = default(decimal);
            observedType = default(ObservedType);
            return false;
        }
    }
}
