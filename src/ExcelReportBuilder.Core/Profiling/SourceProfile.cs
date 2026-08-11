using System;
using System.Collections.Generic;

namespace ExcelReportBuilder.Core.Profiling
{
    public enum SourceValueType
    {
        Empty,
        Text,
        WholeNumber,
        DecimalNumber,
        Boolean,
        Date,
        DateTime,
        Mixed
    }

    public enum SourceProfileIssueCode
    {
        BlankHeader,
        DuplicateHeader,
        RaggedRow
    }

    public sealed class SourceProfileIssue
    {
        public SourceProfileIssueCode Code { get; set; }

        public string Message { get; set; } = string.Empty;

        public int? RowIndex { get; set; }

        public int? ColumnIndex { get; set; }
    }

    public sealed class SourceProfile
    {
        public long RowCount { get; set; }

        public int ColumnCount { get; set; }

        public List<SourceColumnProfile> Columns { get; set; } = new List<SourceColumnProfile>();

        public List<SourceProfileIssue> Issues { get; set; } = new List<SourceProfileIssue>();

        public SourceColumnProfile? FindColumn(string name)
        {
            if (name == null)
            {
                throw new ArgumentNullException(nameof(name));
            }

            for (var index = 0; index < Columns.Count; index++)
            {
                if (string.Equals(Columns[index].Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return Columns[index];
                }
            }

            return null;
        }
    }

    public sealed class SourceColumnProfile
    {
        public int Index { get; set; }

        public string Name { get; set; } = string.Empty;

        public SourceValueType InferredType { get; set; }

        public long BlankCount { get; set; }

        public long NonBlankCount { get; set; }

        public long DistinctCount { get; set; }

        public long DateLikeCount { get; set; }

        /// <summary>
        /// Period tokens such as Jan or Q1 that are structurally valid but
        /// cannot become dates until a reporting year is explicitly supplied.
        /// </summary>
        public long PeriodLikeWithoutYearCount { get; set; }

        public long DayGrainCount { get; set; }

        public long MonthGrainCount { get; set; }

        public long QuarterGrainCount { get; set; }

        public long NumericCount { get; set; }

        public DateTime? MinimumDate { get; set; }

        public DateTime? MaximumDate { get; set; }

        public decimal? MinimumNumber { get; set; }

        public decimal? MaximumNumber { get; set; }

        public double DateLikeRatio => NonBlankCount == 0 ? 0d : (double)DateLikeCount / NonBlankCount;

        public double PeriodLikeRatio => NonBlankCount == 0
            ? 0d
            : (double)(DateLikeCount + PeriodLikeWithoutYearCount) / NonBlankCount;

        public double NumericRatio => NonBlankCount == 0 ? 0d : (double)NumericCount / NonBlankCount;
    }
}
