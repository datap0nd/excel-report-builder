using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExcelReportBuilder.Core.PivotPlus;

namespace ExcelReportBuilder.AddIn.Host
{
    internal sealed class SyntheticPivotPlusHostService : IPivotPlusHostService
    {
        private PivotPlusPaneSnapshot snapshot = new PivotPlusPaneSnapshot(
            "Sheet1",
            "PivotTable1",
            PivotSourceKind.DataModel,
            new[]
            {
                new PivotPlusFieldSnapshot("Region", "Region", false, false, PivotFieldAreaSupport.All),
                new PivotPlusFieldSnapshot("Department", "Department", false, false, PivotFieldAreaSupport.All),
                new PivotPlusFieldSnapshot("Date", "Date", false, false, PivotFieldAreaSupport.All),
                new PivotPlusFieldSnapshot("Cost", "Cost", false, false, PivotFieldAreaSupport.All)
            },
            new[]
            {
                new PivotPlusPlacementSnapshot("Region", "Region", PivotFieldArea.Row, 1, null, string.Empty),
                new PivotPlusPlacementSnapshot("Department", "Department", PivotFieldArea.Row, 2, null, string.Empty),
                new PivotPlusPlacementSnapshot("Cost", "Sum of Cost", PivotFieldArea.Values, 1, PivotAggregationFunction.Sum, "#,##0")
            });

        public Task<PivotPlusPaneSnapshot> InspectAsync(CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult(snapshot);
        }

        public Task<PivotPlusPaneSnapshot> ApplyLayoutAsync(
            IReadOnlyList<PivotPlusPlacementRequest> placements,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            snapshot = new PivotPlusPaneSnapshot(
                snapshot.WorksheetName,
                snapshot.PivotTableName,
                snapshot.SourceKind,
                snapshot.Fields,
                placements.Select(item => new PivotPlusPlacementSnapshot(
                    item.FieldName,
                    string.IsNullOrWhiteSpace(item.Caption) ? item.FieldName : item.Caption,
                    item.Area,
                    item.Position,
                    item.Aggregation,
                    item.NumberFormatCode)));
            return Task.FromResult(snapshot);
        }

        public Task<PivotPlusPaneSnapshot> AddParentPortionAsync(
            string valueFieldName,
            string detailFieldName,
            string measureCaption,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var placements = snapshot.Placements.ToList();
            placements.Add(new PivotPlusPlacementSnapshot(
                "[Measures].[" + (string.IsNullOrWhiteSpace(measureCaption) ? "Portion %" : measureCaption) + "]",
                string.IsNullOrWhiteSpace(measureCaption) ? "Portion %" : measureCaption,
                PivotFieldArea.Values,
                placements.Count(item => item.Area == PivotFieldArea.Values) + 1,
                null,
                "0.0%"));
            snapshot = new PivotPlusPaneSnapshot(
                snapshot.WorksheetName,
                snapshot.PivotTableName,
                snapshot.SourceKind,
                snapshot.Fields,
                placements);
            return Task.FromResult(snapshot);
        }

        public Task<PivotPlusPaneSnapshot> EnableDataModelAsync(CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            snapshot = new PivotPlusPaneSnapshot(
                snapshot.WorksheetName,
                snapshot.PivotTableName,
                PivotSourceKind.DataModel,
                snapshot.Fields,
                snapshot.Placements);
            return Task.FromResult(snapshot);
        }

        public Task<PivotPlusPaneSnapshot> GroupDateAsync(
            string fieldName,
            PivotDateGrouping grouping,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult(snapshot);
        }

        public Task<PivotPlusPaneSnapshot> UndoLastExtraAsync(CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            IReadOnlyList<PivotPlusPlacementSnapshot> placements = snapshot.Placements
                .Where(item => !item.Caption.EndsWith("%", StringComparison.Ordinal))
                .ToList();
            snapshot = new PivotPlusPaneSnapshot(
                snapshot.WorksheetName,
                snapshot.PivotTableName,
                snapshot.SourceKind,
                snapshot.Fields,
                placements);
            return Task.FromResult(snapshot);
        }

        public void OpenExcelFieldList()
        {
        }
    }
}
