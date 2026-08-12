using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ExcelReportBuilder.Core.PivotPlus;

namespace ExcelReportBuilder.AddIn.Host
{
    public sealed class PivotPlusFieldSnapshot
    {
        public PivotPlusFieldSnapshot(
            string name,
            string caption,
            bool isMeasure,
            bool isCalculated,
            PivotFieldAreaSupport supportedAreas)
        {
            Name = name ?? string.Empty;
            Caption = caption ?? string.Empty;
            IsMeasure = isMeasure;
            IsCalculated = isCalculated;
            SupportedAreas = supportedAreas;
        }

        public string Name { get; }
        public string Caption { get; }
        public bool IsMeasure { get; }
        public bool IsCalculated { get; }
        public PivotFieldAreaSupport SupportedAreas { get; }
    }

    public sealed class PivotPlusPlacementSnapshot
    {
        public PivotPlusPlacementSnapshot(
            string fieldName,
            string caption,
            PivotFieldArea area,
            int position,
            PivotAggregationFunction? aggregation,
            string numberFormatCode)
        {
            FieldName = fieldName ?? string.Empty;
            Caption = caption ?? string.Empty;
            Area = area;
            Position = position;
            Aggregation = aggregation;
            NumberFormatCode = numberFormatCode ?? string.Empty;
        }

        public string FieldName { get; }
        public string Caption { get; }
        public PivotFieldArea Area { get; }
        public int Position { get; }
        public PivotAggregationFunction? Aggregation { get; }
        public string NumberFormatCode { get; }
    }

    public sealed class PivotPlusPaneSnapshot
    {
        public PivotPlusPaneSnapshot(
            string worksheetName,
            string pivotTableName,
            PivotSourceKind sourceKind,
            IEnumerable<PivotPlusFieldSnapshot> fields,
            IEnumerable<PivotPlusPlacementSnapshot> placements)
        {
            WorksheetName = worksheetName ?? string.Empty;
            PivotTableName = pivotTableName ?? string.Empty;
            SourceKind = sourceKind;
            Fields = new List<PivotPlusFieldSnapshot>(fields ??
                Array.Empty<PivotPlusFieldSnapshot>()).AsReadOnly();
            Placements = new List<PivotPlusPlacementSnapshot>(placements ??
                Array.Empty<PivotPlusPlacementSnapshot>()).AsReadOnly();
        }

        public string WorksheetName { get; }
        public string PivotTableName { get; }
        public PivotSourceKind SourceKind { get; }
        public IReadOnlyList<PivotPlusFieldSnapshot> Fields { get; }
        public IReadOnlyList<PivotPlusPlacementSnapshot> Placements { get; }
        public bool SupportsExtras => SourceKind == PivotSourceKind.DataModel;
        public bool CanEnableDataModel => SourceKind == PivotSourceKind.WorksheetTable ||
                                          SourceKind == PivotSourceKind.WorksheetRange;
    }

    public sealed class PivotPlusPlacementRequest
    {
        public PivotPlusPlacementRequest(
            string fieldName,
            PivotFieldArea area,
            int position,
            string caption,
            PivotAggregationFunction? aggregation,
            string numberFormatCode)
        {
            FieldName = fieldName ?? string.Empty;
            Area = area;
            Position = position;
            Caption = caption ?? string.Empty;
            Aggregation = aggregation;
            NumberFormatCode = numberFormatCode ?? string.Empty;
        }

        public string FieldName { get; }
        public PivotFieldArea Area { get; }
        public int Position { get; }
        public string Caption { get; }
        public PivotAggregationFunction? Aggregation { get; }
        public string NumberFormatCode { get; }
    }

    public interface IPivotPlusHostService
    {
        Task<PivotPlusPaneSnapshot> InspectAsync(CancellationToken cancellationToken);

        Task<PivotPlusPaneSnapshot> ApplyLayoutAsync(
            IReadOnlyList<PivotPlusPlacementRequest> placements,
            CancellationToken cancellationToken);

        Task<PivotPlusPaneSnapshot> EnableDataModelAsync(
            CancellationToken cancellationToken);

        Task<PivotPlusPaneSnapshot> AddParentPortionAsync(
            string valueFieldName,
            string detailFieldName,
            string measureCaption,
            CancellationToken cancellationToken);

        Task<PivotPlusPaneSnapshot> UndoLastExtraAsync(
            CancellationToken cancellationToken);

        void OpenExcelFieldList();
    }
}
