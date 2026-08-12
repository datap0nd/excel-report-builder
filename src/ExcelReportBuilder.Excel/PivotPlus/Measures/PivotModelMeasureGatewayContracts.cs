using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using ExcelReportBuilder.Core.PivotPlus;
using ExcelReportBuilder.Core.PivotPlus.Calculations;

namespace ExcelReportBuilder.Excel.PivotPlus.Measures
{
    internal enum ExcelModelMeasureFormatKind
    {
        General,
        Boolean,
        WholeNumber,
        DecimalNumber,
        PercentageNumber,
        ScientificNumber,
        Currency,
        Date
    }

    internal sealed class ModelMeasureFormatSnapshot
    {
        public ModelMeasureFormatSnapshot(
            ExcelModelMeasureFormatKind kind,
            int? decimalPlaces = null,
            bool? useThousandsSeparator = null,
            string? currencySymbol = null,
            string? dateFormatString = null)
        {
            Kind = kind;
            DecimalPlaces = decimalPlaces;
            UseThousandsSeparator = useThousandsSeparator;
            CurrencySymbol = currencySymbol;
            DateFormatString = dateFormatString;
        }

        public ExcelModelMeasureFormatKind Kind { get; }

        public int? DecimalPlaces { get; }

        public bool? UseThousandsSeparator { get; }

        public string? CurrencySymbol { get; }

        public string? DateFormatString { get; }
    }

    internal sealed class DesiredModelMeasure
    {
        public DesiredModelMeasure(
            string definitionId,
            int displayOrder,
            int creationOrder,
            string homeTableName,
            string name,
            string formula,
            PivotMeasureFormat format,
            IEnumerable<string>? directDependencyDefinitionIds,
            string definitionFingerprint,
            string descriptionMarker)
        {
            DefinitionId = definitionId;
            DisplayOrder = displayOrder;
            CreationOrder = creationOrder;
            HomeTableName = homeTableName;
            Name = name;
            Formula = formula;
            Format = format;
            DirectDependencyDefinitionIds = new ReadOnlyCollection<string>(
                (directDependencyDefinitionIds ?? Enumerable.Empty<string>()).ToList());
            DefinitionFingerprint = definitionFingerprint;
            DescriptionMarker = descriptionMarker;
        }

        public string DefinitionId { get; }

        public int DisplayOrder { get; }

        public int CreationOrder { get; }

        public string HomeTableName { get; }

        public string Name { get; }

        public string Formula { get; }

        public PivotMeasureFormat Format { get; }

        public IReadOnlyList<string> DirectDependencyDefinitionIds { get; }

        public string DefinitionFingerprint { get; }

        public string DescriptionMarker { get; }
    }

    internal sealed class LiveModelMeasureSnapshot
    {
        public LiveModelMeasureSnapshot(
            string name,
            string associatedTableName,
            string associatedTableLineageFingerprint,
            string formula,
            string description,
            ModelMeasureFormatSnapshot format,
            string liveFingerprint)
        {
            Name = name;
            AssociatedTableName = associatedTableName;
            AssociatedTableLineageFingerprint = associatedTableLineageFingerprint;
            Formula = formula;
            Description = description;
            Format = format;
            LiveFingerprint = liveFingerprint;
        }

        public string Name { get; }

        public string AssociatedTableName { get; }

        public string AssociatedTableLineageFingerprint { get; }

        public string Formula { get; }

        public string Description { get; }

        public ModelMeasureFormatSnapshot Format { get; }

        public string LiveFingerprint { get; }
    }

    internal sealed class ModelDataFieldSnapshot
    {
        public ModelDataFieldSnapshot(
            string uniqueName,
            string caption,
            string captionFingerprint,
            string numberFormat,
            int position,
            bool isModelMeasure,
            string? modelMeasureName = null)
        {
            UniqueName = uniqueName;
            Caption = caption;
            CaptionFingerprint = captionFingerprint;
            NumberFormat = numberFormat;
            Position = position;
            IsModelMeasure = isModelMeasure;
            ModelMeasureName = modelMeasureName;
        }

        public string UniqueName { get; }

        public string Caption { get; }

        public string CaptionFingerprint { get; }

        public string NumberFormat { get; }

        public int Position { get; }

        public bool IsModelMeasure { get; }

        public string? ModelMeasureName { get; }
    }

    internal sealed class ModelPivotUsageSnapshot
    {
        public ModelPivotUsageSnapshot(
            string worksheetName,
            string pivotTableName,
            bool isSelectedTarget,
            IEnumerable<ModelDataFieldSnapshot>? dataFields,
            PivotValuesAxis valuesAxis,
            int valuesPosition)
        {
            WorksheetName = worksheetName;
            PivotTableName = pivotTableName;
            IsSelectedTarget = isSelectedTarget;
            DataFields = new ReadOnlyCollection<ModelDataFieldSnapshot>(
                (dataFields ?? Enumerable.Empty<ModelDataFieldSnapshot>())
                    .OrderBy(field => field.Position)
                    .ToList());
            ValuesAxis = valuesAxis;
            ValuesPosition = valuesPosition;
        }

        public string WorksheetName { get; }

        public string PivotTableName { get; }

        public bool IsSelectedTarget { get; }

        public IReadOnlyList<ModelDataFieldSnapshot> DataFields { get; }

        public PivotValuesAxis ValuesAxis { get; }

        public int ValuesPosition { get; }
    }

    internal sealed class ModelMeasureWorkbookSnapshot
    {
        public ModelMeasureWorkbookSnapshot(
            IEnumerable<LiveModelMeasureSnapshot>? measures,
            IEnumerable<ModelPivotUsageSnapshot>? pivotUsages,
            string selectedPivotFingerprint)
        {
            Measures = new ReadOnlyCollection<LiveModelMeasureSnapshot>(
                (measures ?? Enumerable.Empty<LiveModelMeasureSnapshot>()).ToList());
            PivotUsages = new ReadOnlyCollection<ModelPivotUsageSnapshot>(
                (pivotUsages ?? Enumerable.Empty<ModelPivotUsageSnapshot>()).ToList());
            SelectedPivotFingerprint = selectedPivotFingerprint;
        }

        public IReadOnlyList<LiveModelMeasureSnapshot> Measures { get; }

        public IReadOnlyList<ModelPivotUsageSnapshot> PivotUsages { get; }

        public string SelectedPivotFingerprint { get; }

        public ModelPivotUsageSnapshot SelectedPivot => PivotUsages.Single(usage =>
            usage.IsSelectedTarget);
    }

    internal sealed class BoundModelMeasureTarget
    {
        public BoundModelMeasureTarget(
            object workbook,
            object pivotTable,
            object model,
            object dataModelConnection,
            PivotTargetIdentity identity)
        {
            Workbook = workbook;
            PivotTable = pivotTable;
            Model = model;
            DataModelConnection = dataModelConnection;
            Identity = identity;
        }

        public object Workbook { get; }

        public object PivotTable { get; }

        public object Model { get; }

        public object DataModelConnection { get; }

        public PivotTargetIdentity Identity { get; }
    }

    internal interface IPivotModelMeasureGateway
    {
        BoundModelMeasureTarget Bind(
            object workbook,
            object pivotTable,
            PivotTableContext context);

        ModelMeasureWorkbookSnapshot Capture(BoundModelMeasureTarget target);

        LiveModelMeasureSnapshot CreateMeasure(
            BoundModelMeasureTarget target,
            DesiredModelMeasure definition);

        LiveModelMeasureSnapshot UpdateMeasure(
            BoundModelMeasureTarget target,
            LiveModelMeasureSnapshot before,
            DesiredModelMeasure definition);

        LiveModelMeasureSnapshot RestoreMeasure(
            BoundModelMeasureTarget target,
            LiveModelMeasureSnapshot before);

        void DeleteMeasure(
            BoundModelMeasureTarget target,
            LiveModelMeasureSnapshot expected);

        void ApplyPlacement(
            BoundModelMeasureTarget target,
            PivotMeasurePlacementPlan placement,
            IReadOnlyDictionary<string, DesiredModelMeasure> definitionsById,
            ModelMeasureWorkbookSnapshot before);

        void RestorePlacement(
            BoundModelMeasureTarget target,
            ModelPivotUsageSnapshot before);

        void Refresh(BoundModelMeasureTarget target);
    }
}
