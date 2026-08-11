using System;

namespace ExcelReportBuilder.Excel.Execution
{
    public enum ExcelBuildStage
    {
        Inspecting,
        Normalizing,
        Planning,
        BuildingPivots,
        Rendering,
        Calculating,
        Checking,
        Repairing,
        Complete
    }

    public sealed class ExcelProgress
    {
        public ExcelBuildStage Stage { get; set; }

        public string Operation { get; set; } = string.Empty;

        public string? ManagedObject { get; set; }

        public long? SourceRows { get; set; }

        public long? ProjectedRows { get; set; }

        public int CompletedChecks { get; set; }

        public TimeSpan Elapsed { get; set; }

        public bool IsHeartbeat { get; set; }
    }

    public interface IExcelProgressSink
    {
        void Report(ExcelProgress progress);
    }

    public sealed class NullExcelProgressSink : IExcelProgressSink
    {
        public static readonly NullExcelProgressSink Instance = new NullExcelProgressSink();

        private NullExcelProgressSink()
        {
        }

        public void Report(ExcelProgress progress)
        {
        }
    }
}
