using System;
using System.Collections.Generic;
using System.Linq;
using ExcelReportBuilder.AddIn.Host;

namespace ExcelReportBuilder.AddIn.Presentation
{
    public enum ShellSurface
    {
        Data,
        Build,
        Chat,
        Checks
    }

    public sealed class FieldPlacement : ManualEditableObject
    {
        private string _setting;
        private bool _showSubtotals;
        private string _subtotalPlacement;
        private string _memberOrderText;
        private string _numberFormat;

        public FieldPlacement(
            PlacementBucket bucket,
            string columnName,
            string setting,
            bool showSubtotals = true,
            string subtotalPlacement = "After members",
            string memberOrderText = "",
            string numberFormat = "#,##0.00")
        {
            Id = Guid.NewGuid();
            Bucket = bucket;
            ColumnName = columnName ?? throw new ArgumentNullException(nameof(columnName));
            _setting = setting ?? string.Empty;
            _showSubtotals = showSubtotals;
            _subtotalPlacement = subtotalPlacement ?? "After members";
            _memberOrderText = memberOrderText ?? string.Empty;
            _numberFormat = numberFormat ?? "#,##0.00";
        }

        public Guid Id { get; }

        public PlacementBucket Bucket { get; }

        public string BucketLabel => Bucket.ToString();

        public string ColumnName { get; }

        public string Setting
        {
            get => _setting;
            set
            {
                if (CanEdit) SetProperty(ref _setting, value ?? string.Empty);
            }
        }

        public bool ShowSubtotals
        {
            get => _showSubtotals;
            set
            {
                if (CanEdit) SetProperty(ref _showSubtotals, value);
            }
        }

        public string SubtotalPlacement
        {
            get => _subtotalPlacement;
            set
            {
                if (CanEdit) SetProperty(ref _subtotalPlacement, value ?? "After members");
            }
        }

        public string MemberOrderText
        {
            get => _memberOrderText;
            set
            {
                if (CanEdit) SetProperty(ref _memberOrderText, value ?? string.Empty);
            }
        }

        public string NumberFormat
        {
            get => _numberFormat;
            set
            {
                if (CanEdit) SetProperty(ref _numberFormat, value ?? "General");
            }
        }

        public FieldPlacementSnapshot ToSnapshot()
        {
            IReadOnlyList<string> selectedValues = Bucket == PlacementBucket.Filters
                ? ParseSelectedValues(Setting)
                : Array.Empty<string>();
            return new FieldPlacementSnapshot(
                Bucket,
                ColumnName,
                Setting,
                ShowSubtotals,
                selectedValues,
                SubtotalPlacement,
                ParseSelectedValues(MemberOrderText),
                NumberFormat);
        }

        private static IReadOnlyList<string> ParseSelectedValues(string setting)
        {
            if (string.IsNullOrWhiteSpace(setting) ||
                string.Equals(setting.Trim(), "All", StringComparison.OrdinalIgnoreCase))
            {
                return Array.Empty<string>();
            }

            return setting.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(value => value.Trim())
                .Where(value => value.Length > 0)
                .Distinct(StringComparer.Ordinal)
                .ToArray();
        }
    }

    public sealed class ChatLine
    {
        public ChatLine(string speaker, string message)
        {
            Speaker = speaker ?? throw new ArgumentNullException(nameof(speaker));
            Message = message ?? throw new ArgumentNullException(nameof(message));
        }

        public string Speaker { get; }

        public string Message { get; }

        public string AccessibleSummary => $"{Speaker}: {Message}";
    }

    public sealed class CheckLine
    {
        public CheckLine(string name, string status, string detail, bool passed)
        {
            Name = name ?? throw new ArgumentNullException(nameof(name));
            Status = status ?? throw new ArgumentNullException(nameof(status));
            Detail = detail ?? throw new ArgumentNullException(nameof(detail));
            Passed = passed;
        }

        public string Name { get; }

        public string Status { get; }

        public string Detail { get; }

        public bool Passed { get; }

        public string AccessibleSummary => $"{Name}, {Status}. {Detail}";
    }
}
