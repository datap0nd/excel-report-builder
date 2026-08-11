using System;

namespace ExcelReportBuilder.Core.Specifications
{
    public enum ScalarValueKind
    {
        Null,
        Text,
        Number,
        Boolean,
        Date,
        DateTime
    }

    /// <summary>
    /// A bounded literal value used by filters and replacements. It deliberately
    /// cannot carry an Excel formula or Power Query expression.
    /// </summary>
    public sealed class ScalarValue
    {
        public ScalarValueKind Kind { get; set; }

        public string? Text { get; set; }

        public decimal? Number { get; set; }

        public bool? Boolean { get; set; }

        public DateTime? Temporal { get; set; }

        public static ScalarValue Null()
        {
            return new ScalarValue { Kind = ScalarValueKind.Null };
        }

        public static ScalarValue FromText(string value)
        {
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            return new ScalarValue { Kind = ScalarValueKind.Text, Text = value };
        }

        public static ScalarValue FromNumber(decimal value)
        {
            return new ScalarValue { Kind = ScalarValueKind.Number, Number = value };
        }

        public static ScalarValue FromBoolean(bool value)
        {
            return new ScalarValue { Kind = ScalarValueKind.Boolean, Boolean = value };
        }

        public static ScalarValue FromDate(DateTime value)
        {
            return new ScalarValue { Kind = ScalarValueKind.Date, Temporal = value.Date };
        }

        public static ScalarValue FromDateTime(DateTime value)
        {
            return new ScalarValue { Kind = ScalarValueKind.DateTime, Temporal = value };
        }
    }
}
