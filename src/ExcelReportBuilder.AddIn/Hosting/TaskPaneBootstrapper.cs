using System;
using ExcelReportBuilder.AddIn.Host;

namespace ExcelReportBuilder.AddIn.Hosting
{
    /// <summary>
    /// Allows the Excel host composition root to replace the synthetic shell service
    /// before Office asks COM to instantiate the task-pane control.
    /// </summary>
    public static class TaskPaneBootstrapper
    {
        public static Func<IPivotPlusHostService>? HostServiceFactory { get; set; }

        internal static IPivotPlusHostService CreateHostService()
        {
            return HostServiceFactory?.Invoke() ?? new SyntheticPivotPlusHostService();
        }
    }
}
