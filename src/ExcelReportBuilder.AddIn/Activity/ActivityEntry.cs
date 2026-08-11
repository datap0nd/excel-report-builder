using System;

namespace ExcelReportBuilder.AddIn.Activity
{
    public sealed class ActivityEntry
    {
        public ActivityEntry(
            DateTimeOffset timestamp,
            ActivityStage stage,
            ActivityKind kind,
            string message,
            string detail)
        {
            Timestamp = timestamp;
            Stage = stage;
            Kind = kind;
            Message = message ?? throw new ArgumentNullException(nameof(message));
            Detail = detail ?? string.Empty;
        }

        public DateTimeOffset Timestamp { get; }

        public ActivityStage Stage { get; }

        public ActivityKind Kind { get; }

        public string Message { get; }

        public string Detail { get; }

        public string DisplayTime => Timestamp.ToLocalTime().ToString("HH:mm:ss");

        public string StageLabel => ActivityLabels.Stage(Stage);

        public string KindLabel => ActivityLabels.Kind(Kind);

        public string AccessibleSummary => string.IsNullOrWhiteSpace(Detail)
            ? $"{DisplayTime}, {KindLabel}, {StageLabel}, {Message}"
            : $"{DisplayTime}, {KindLabel}, {StageLabel}, {Message}. {Detail}";
    }
}
