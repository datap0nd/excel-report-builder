using System;
using System.Globalization;
using ExcelReportBuilder.AddIn.Host;

namespace ExcelReportBuilder.AddIn.Presentation
{
    public sealed class ManualTransformRule : ObservableObject
    {
        private string _operation = "Trim text";
        private string _column = string.Empty;
        private string _outputColumn = string.Empty;
        private string _details = string.Empty;

        public string Operation
        {
            get => _operation;
            set
            {
                if (SetProperty(ref _operation, value ?? string.Empty)) RaisePropertyChanged(nameof(Summary));
            }
        }

        public string Column
        {
            get => _column;
            set
            {
                if (SetProperty(ref _column, value ?? string.Empty)) RaisePropertyChanged(nameof(Summary));
            }
        }

        public string OutputColumn
        {
            get => _outputColumn;
            set => SetProperty(ref _outputColumn, value ?? string.Empty);
        }

        public string Details
        {
            get => _details;
            set => SetProperty(ref _details, value ?? string.Empty);
        }

        public string Summary => string.IsNullOrWhiteSpace(Column)
            ? Operation
            : Operation + " · " + Column;

        public ManualTransformSnapshot ToSnapshot()
        {
            return new ManualTransformSnapshot(Operation, Column, OutputColumn, Details);
        }
    }

    public sealed class ManualCalculatedMetricRule : ObservableObject
    {
        private string _label = "Calculated metric";
        private string _kind = "Ratio";
        private string _primary = string.Empty;
        private string _secondary = string.Empty;
        private string _details = string.Empty;
        private string _numberFormat = "0.0%";

        public string Label
        {
            get => _label;
            set
            {
                if (SetProperty(ref _label, value ?? string.Empty)) RaisePropertyChanged(nameof(Summary));
            }
        }

        public string Kind
        {
            get => _kind;
            set
            {
                if (SetProperty(ref _kind, value ?? string.Empty)) RaisePropertyChanged(nameof(Summary));
            }
        }

        public string Primary
        {
            get => _primary;
            set => SetProperty(ref _primary, value ?? string.Empty);
        }

        public string Secondary
        {
            get => _secondary;
            set => SetProperty(ref _secondary, value ?? string.Empty);
        }

        public string Details
        {
            get => _details;
            set => SetProperty(ref _details, value ?? string.Empty);
        }

        public string NumberFormat
        {
            get => _numberFormat;
            set => SetProperty(ref _numberFormat, value ?? "General");
        }

        public string Summary => Label + " · " + Kind;

        public ManualCalculatedMetricSnapshot ToSnapshot()
        {
            return new ManualCalculatedMetricSnapshot(
                Label,
                Kind,
                Primary,
                Secondary,
                Details,
                NumberFormat);
        }
    }

    public sealed class ManualReportBlockRule : ObservableObject
    {
        private const int MaximumOwnedRows = 1048576;
        private const int MaximumOwnedColumns = 16384;

        private string _title = "Management report";
        private string _worksheetName = "Report";
        private string _anchorCell = "A1";
        private string _outputStyle = "Dense management block";
        private int _ownedRows = 500;
        private int _ownedColumns = 64;

        public ManualReportBlockRule(string? stableId = null)
        {
            StableId = string.IsNullOrWhiteSpace(stableId)
                ? "report_block_" + Guid.NewGuid().ToString("N")
                : stableId!.Trim();
        }

        public string StableId { get; }

        public string Title
        {
            get => _title;
            set
            {
                if (SetProperty(ref _title, value ?? string.Empty)) RaisePropertyChanged(nameof(Summary));
            }
        }

        public string WorksheetName
        {
            get => _worksheetName;
            set
            {
                if (SetProperty(ref _worksheetName, value ?? string.Empty)) RaisePropertyChanged(nameof(Summary));
            }
        }

        public string AnchorCell
        {
            get => _anchorCell;
            set
            {
                if (SetProperty(ref _anchorCell, value ?? string.Empty)) RaisePropertyChanged(nameof(Summary));
            }
        }

        public string OutputStyle
        {
            get => _outputStyle;
            set => SetProperty(ref _outputStyle, value ?? "Dense management block");
        }

        public int OwnedRows
        {
            get => _ownedRows;
            set
            {
                int bounded = Math.Max(1, Math.Min(MaximumOwnedRows, value));
                if (SetProperty(ref _ownedRows, bounded)) RaisePropertyChanged(nameof(Summary));
            }
        }

        public int OwnedColumns
        {
            get => _ownedColumns;
            set
            {
                int bounded = Math.Max(1, Math.Min(MaximumOwnedColumns, value));
                if (SetProperty(ref _ownedColumns, bounded)) RaisePropertyChanged(nameof(Summary));
            }
        }

        public string Summary => WorksheetName + "!" + AnchorCell + " · " + Title +
            " · " + OwnedRows + " rows × " + OwnedColumns + " columns";

        public ManualReportBlockSnapshot ToSnapshot()
        {
            return new ManualReportBlockSnapshot(
                Title,
                WorksheetName,
                AnchorCell,
                OutputStyle,
                StableId,
                OwnedRows,
                OwnedColumns);
        }
    }

    public sealed class ManualCheckRule : ObservableObject
    {
        private string _kind = "Total preservation";
        private string _metric = string.Empty;
        private string _comparedMetric = string.Empty;
        private string _toleranceText = "0";

        public string Kind
        {
            get => _kind;
            set
            {
                if (SetProperty(ref _kind, value ?? string.Empty)) RaisePropertyChanged(nameof(Summary));
            }
        }

        public string Metric
        {
            get => _metric;
            set
            {
                if (SetProperty(ref _metric, value ?? string.Empty)) RaisePropertyChanged(nameof(Summary));
            }
        }

        public string ComparedMetric
        {
            get => _comparedMetric;
            set => SetProperty(ref _comparedMetric, value ?? string.Empty);
        }

        public string ToleranceText
        {
            get => _toleranceText;
            set => SetProperty(ref _toleranceText, value ?? string.Empty);
        }

        public string Summary => Kind + (string.IsNullOrWhiteSpace(Metric) ? string.Empty : " · " + Metric);

        public ManualCheckSnapshot ToSnapshot()
        {
            if (!decimal.TryParse(
                    ToleranceText,
                    NumberStyles.Number,
                    CultureInfo.InvariantCulture,
                    out decimal tolerance) || tolerance < 0m)
            {
                throw new InvalidOperationException(
                    "Check tolerance must be a non-negative number using a period as the decimal separator.");
            }

            return new ManualCheckSnapshot(Kind, Metric, ComparedMetric, tolerance);
        }
    }
}
