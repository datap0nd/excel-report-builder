using System;
using System.Collections.Generic;
using System.Linq;
using ExcelReportBuilder.Core.Planning;
using ExcelReportBuilder.Excel.Ownership;

namespace ExcelReportBuilder.Excel.Execution
{
    public sealed class ManagedOutputWorksheetPlan
    {
        public string LogicalWorksheetName { get; set; } = string.Empty;

        public ManagedObjectIdentity DraftIdentity { get; set; } = null!;

        public IReadOnlyList<DenseReportBlockPlan> Blocks { get; set; } =
            Array.Empty<DenseReportBlockPlan>();
    }

    public static class ManagedOutputLayoutPlanner
    {
        public static IReadOnlyList<ManagedOutputWorksheetPlan> Group(
            string reportId,
            IReadOnlyList<DenseReportBlockPlan> blocks)
        {
            if (string.IsNullOrWhiteSpace(reportId))
            {
                throw new ArgumentException("A report identifier is required.", nameof(reportId));
            }

            if (blocks == null)
            {
                throw new ArgumentNullException(nameof(blocks));
            }

            return blocks
                .GroupBy(
                    block => ManagedOutputIdentity.LogicalKey(block.WorksheetName),
                    StringComparer.Ordinal)
                .Select(group => new ManagedOutputWorksheetPlan
                {
                    LogicalWorksheetName = group.First().WorksheetName,
                    DraftIdentity = ManagedOutputIdentity.Draft(
                        reportId,
                        group.First().WorksheetName),
                    Blocks = group.ToList()
                })
                .ToList();
        }
    }
}
