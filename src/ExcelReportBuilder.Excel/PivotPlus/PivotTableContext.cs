using System;
using ExcelReportBuilder.Core.PivotPlus;

namespace ExcelReportBuilder.Excel.PivotPlus
{
    /// <summary>
    /// Excel-host state that is intentionally not part of the portable PivotTable+
    /// layout contract. All identity, source, field, capability, and placement data
    /// lives in <see cref="PivotLayoutDefinition"/>.
    /// </summary>
    public sealed class PivotTableContext
    {
        public PivotTableContext(
            PivotLayoutDefinition definition,
            bool isConnected,
            bool sourceFieldsComplete)
        {
            Definition = definition ?? throw new ArgumentNullException(nameof(definition));
            IsConnected = isConnected;
            SourceFieldsComplete = sourceFieldsComplete;
        }

        public PivotLayoutDefinition Definition { get; }

        /// <summary>
        /// False when Excel still exposes a cached OLAP layout but its workbook
        /// connection can no longer be read.
        /// </summary>
        public bool IsConnected { get; }

        /// <summary>
        /// False when discovery had to reconstruct the field inventory from the
        /// visible axes because Excel could not expose CubeFields.
        /// </summary>
        public bool SourceFieldsComplete { get; }
    }
}
