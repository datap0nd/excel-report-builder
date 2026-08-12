using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using ExcelReportBuilder.Core.PivotPlus;
using ExcelReportBuilder.Excel.PivotPlus.Persistence;

namespace ExcelReportBuilder.Excel.PivotPlus.Semantics
{
    /// <summary>
    /// Identifies an existing row or column field exactly as it appeared at
    /// preview. The provider unique name is the host identity; caption and
    /// position fingerprints prevent a later, same-named field from being
    /// accepted accidentally.
    /// </summary>
    public sealed class PivotExistingAxisFieldIdentity
    {
        public PivotExistingAxisFieldIdentity(
            string uniqueName,
            string currentCaptionFingerprint,
            PivotFieldArea currentArea,
            int currentPosition)
        {
            UniqueName = uniqueName ?? string.Empty;
            CurrentCaptionFingerprint = currentCaptionFingerprint ?? string.Empty;
            CurrentArea = currentArea;
            CurrentPosition = currentPosition;
        }

        public string UniqueName { get; }

        public string CurrentCaptionFingerprint { get; }

        public PivotFieldArea CurrentArea { get; }

        /// <summary>
        /// One-based regular-field position with Excel's Values pseudo-field
        /// removed from the sequence.
        /// </summary>
        public int CurrentPosition { get; }
    }

    /// <summary>
    /// Identifies one exact existing Values occurrence. Position is part of
    /// the identity so repeated instances never bind by caption alone.
    /// </summary>
    public sealed class PivotExistingSemanticValueIdentity
    {
        public PivotExistingSemanticValueIdentity(
            string uniqueName,
            string currentCaptionFingerprint,
            string currentNumberFormatFingerprint,
            int currentPosition)
        {
            UniqueName = uniqueName ?? string.Empty;
            CurrentCaptionFingerprint = currentCaptionFingerprint ?? string.Empty;
            CurrentNumberFormatFingerprint = currentNumberFormatFingerprint ?? string.Empty;
            CurrentPosition = currentPosition;
        }

        public string UniqueName { get; }

        public string CurrentCaptionFingerprint { get; }

        public string CurrentNumberFormatFingerprint { get; }

        public int CurrentPosition { get; }
    }

    /// <summary>
    /// One row/column entry. Exactly one of DefinitionId and ExistingField is
    /// populated. DefinitionId is resolved only through a trusted named-set
    /// compilation map supplied by the semantic coordinator.
    /// </summary>
    public sealed class PivotSemanticAxisPlacement
    {
        public PivotSemanticAxisPlacement(int position, string definitionId)
        {
            Position = position;
            DefinitionId = definitionId ?? string.Empty;
        }

        public PivotSemanticAxisPlacement(
            int position,
            PivotExistingAxisFieldIdentity existingField)
        {
            Position = position;
            ExistingField = existingField ??
                throw new ArgumentNullException(nameof(existingField));
        }

        public int Position { get; }

        public string? DefinitionId { get; }

        public PivotExistingAxisFieldIdentity? ExistingField { get; }

        public bool IsGeneratedNamedSet => ExistingField == null;
    }

    /// <summary>
    /// One Values entry. Exactly one of DefinitionId and ExistingDataField is
    /// populated. DefinitionId is resolved only through a trusted compiled
    /// model-measure map.
    /// </summary>
    public sealed class PivotSemanticValuePlacement
    {
        public PivotSemanticValuePlacement(int position, string definitionId)
        {
            Position = position;
            DefinitionId = definitionId ?? string.Empty;
        }

        public PivotSemanticValuePlacement(
            int position,
            PivotExistingSemanticValueIdentity existingDataField)
        {
            Position = position;
            ExistingDataField = existingDataField ??
                throw new ArgumentNullException(nameof(existingDataField));
        }

        public int Position { get; }

        public string? DefinitionId { get; }

        public PivotExistingSemanticValueIdentity? ExistingDataField { get; }

        public bool IsGeneratedMeasure => ExistingDataField == null;
    }

    /// <summary>
    /// Complete final semantic layout for Rows, Columns, and Values. These are
    /// replacement sequences: leaving an existing placement out explicitly
    /// removes it. Filters are deliberately absent because this layer captures
    /// and preserves the exact existing filter layout and filter state.
    /// </summary>
    public sealed class PivotSemanticLayoutPlan
    {
        public PivotSemanticLayoutPlan(
            IEnumerable<PivotSemanticAxisPlacement>? rows,
            IEnumerable<PivotSemanticAxisPlacement>? columns,
            IEnumerable<PivotSemanticValuePlacement>? values,
            PivotValuesAxis valuesAxis,
            int valuesPosition)
        {
            Rows = Copy(rows);
            Columns = Copy(columns);
            Values = Copy(values);
            ValuesAxis = valuesAxis;
            ValuesPosition = valuesPosition;
        }

        public IReadOnlyList<PivotSemanticAxisPlacement> Rows { get; }

        public IReadOnlyList<PivotSemanticAxisPlacement> Columns { get; }

        public IReadOnlyList<PivotSemanticValuePlacement> Values { get; }

        public PivotValuesAxis ValuesAxis { get; }

        /// <summary>
        /// One-based insertion position among the regular fields of the chosen
        /// axis. For zero or one Values field the required sentinel is one and
        /// ValuesAxis must be Automatic.
        /// </summary>
        public int ValuesPosition { get; }

        private static IReadOnlyList<T> Copy<T>(IEnumerable<T>? values)
        {
            return new ReadOnlyCollection<T>(
                (values ?? Enumerable.Empty<T>()).ToList());
        }
    }

    public static class PivotSemanticLayoutFingerprint
    {
        public static string CreateCaptionFingerprint(string caption)
        {
            if (caption == null) throw new ArgumentNullException(nameof(caption));
            return PivotPlusFingerprint.Create("semantic.caption.v1", caption);
        }

        public static string CreateNumberFormatFingerprint(string numberFormat)
        {
            if (numberFormat == null)
            {
                throw new ArgumentNullException(nameof(numberFormat));
            }

            return PivotPlusFingerprint.Create(
                "semantic.number-format.v1",
                numberFormat);
        }
    }

    internal sealed class PivotSemanticAxisFieldSnapshot
    {
        public PivotSemanticAxisFieldSnapshot(
            string uniqueName,
            string caption,
            string captionFingerprint,
            PivotFieldArea area,
            int position,
            int cubeFieldType)
        {
            UniqueName = uniqueName;
            Caption = caption;
            CaptionFingerprint = captionFingerprint;
            Area = area;
            Position = position;
            CubeFieldType = cubeFieldType;
        }

        public string UniqueName { get; }
        public string Caption { get; }
        public string CaptionFingerprint { get; }
        public PivotFieldArea Area { get; }
        public int Position { get; }
        public int CubeFieldType { get; }
    }

    internal sealed class PivotSemanticValueFieldSnapshot
    {
        public PivotSemanticValueFieldSnapshot(
            string uniqueName,
            string caption,
            string captionFingerprint,
            string numberFormat,
            string numberFormatFingerprint,
            int position,
            int cubeFieldType)
        {
            UniqueName = uniqueName;
            Caption = caption;
            CaptionFingerprint = captionFingerprint;
            NumberFormat = numberFormat;
            NumberFormatFingerprint = numberFormatFingerprint;
            Position = position;
            CubeFieldType = cubeFieldType;
        }

        public string UniqueName { get; }
        public string Caption { get; }
        public string CaptionFingerprint { get; }
        public string NumberFormat { get; }
        public string NumberFormatFingerprint { get; }
        public int Position { get; }
        public int CubeFieldType { get; }
    }

    internal sealed class PivotSemanticFilterFieldSnapshot
    {
        public PivotSemanticFilterFieldSnapshot(
            string uniqueName,
            string caption,
            int position,
            string stateFingerprint)
        {
            UniqueName = uniqueName;
            Caption = caption;
            Position = position;
            StateFingerprint = stateFingerprint;
        }

        public string UniqueName { get; }
        public string Caption { get; }
        public int Position { get; }
        public string StateFingerprint { get; }
    }

    internal sealed class PivotSemanticLayoutSnapshot
    {
        public PivotSemanticLayoutSnapshot(
            PivotTargetIdentity identity,
            IEnumerable<PivotSemanticAxisFieldSnapshot>? rows,
            IEnumerable<PivotSemanticAxisFieldSnapshot>? columns,
            IEnumerable<PivotSemanticValueFieldSnapshot>? values,
            IEnumerable<PivotSemanticFilterFieldSnapshot>? filters,
            PivotValuesAxis valuesAxis,
            int valuesPosition,
            string filterFingerprint,
            string layoutFingerprint)
        {
            Identity = identity ?? throw new ArgumentNullException(nameof(identity));
            Rows = Copy(rows);
            Columns = Copy(columns);
            Values = Copy(values);
            Filters = Copy(filters);
            ValuesAxis = valuesAxis;
            ValuesPosition = valuesPosition;
            FilterFingerprint = filterFingerprint;
            LayoutFingerprint = layoutFingerprint;
        }

        public PivotTargetIdentity Identity { get; }
        public IReadOnlyList<PivotSemanticAxisFieldSnapshot> Rows { get; }
        public IReadOnlyList<PivotSemanticAxisFieldSnapshot> Columns { get; }
        public IReadOnlyList<PivotSemanticValueFieldSnapshot> Values { get; }
        public IReadOnlyList<PivotSemanticFilterFieldSnapshot> Filters { get; }
        public PivotValuesAxis ValuesAxis { get; }
        public int ValuesPosition { get; }
        public string FilterFingerprint { get; }
        public string LayoutFingerprint { get; }

        private static IReadOnlyList<T> Copy<T>(IEnumerable<T>? values)
        {
            return new ReadOnlyCollection<T>(
                (values ?? Enumerable.Empty<T>()).ToList());
        }
    }

    internal sealed class BoundPivotSemanticLayoutTarget
    {
        public BoundPivotSemanticLayoutTarget(
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

    internal interface IPivotSemanticLayoutGateway
    {
        BoundPivotSemanticLayoutTarget Bind(
            object workbook,
            object pivotTable,
            PivotTableContext context);

        PivotSemanticLayoutSnapshot Capture(
            BoundPivotSemanticLayoutTarget target);

        PivotSemanticPreparedPlacement Prepare(
            BoundPivotSemanticLayoutTarget target,
            PivotSemanticLayoutPlan plan,
            IReadOnlyDictionary<string, string> namedSetUniqueNamesByDefinitionId,
            IReadOnlyDictionary<string, string> measureUniqueNamesByDefinitionId,
            PivotSemanticLayoutSnapshot before);
    }

    /// <summary>
    /// Placement participant for the future composite semantic transaction.
    /// It performs no refresh and owns no workbook artifact. Apply is locally
    /// rollback-safe; Restore remains callable if a later transaction phase
    /// fails.
    /// </summary>
    internal sealed class PivotSemanticPreparedPlacement
    {
        private readonly LateBoundPivotSemanticLayoutGateway gateway;
        private readonly BoundPivotSemanticLayoutTarget target;
        private readonly PivotSemanticLayoutPlan plan;
        private readonly IReadOnlyDictionary<string, string> namedSets;
        private readonly IReadOnlyDictionary<string, string> measures;
        private readonly PivotSemanticLayoutSnapshot before;
        private readonly bool alreadyApplied;

        internal PivotSemanticPreparedPlacement(
            LateBoundPivotSemanticLayoutGateway gateway,
            BoundPivotSemanticLayoutTarget target,
            PivotSemanticLayoutPlan plan,
            IReadOnlyDictionary<string, string> namedSets,
            IReadOnlyDictionary<string, string> measures,
            PivotSemanticLayoutSnapshot before,
            bool alreadyApplied = false)
        {
            this.gateway = gateway;
            this.target = target;
            this.plan = plan;
            this.namedSets = namedSets;
            this.measures = measures;
            this.before = before;
            this.alreadyApplied = alreadyApplied;
        }

        public PivotSemanticLayoutSnapshot Before => before;

        public void Apply()
        {
            if (alreadyApplied)
            {
                gateway.VerifyDesired(target, plan, namedSets, measures, before);
                return;
            }

            try
            {
                gateway.ApplyExact(target, plan, namedSets, measures, before);
                gateway.VerifyDesired(target, plan, namedSets, measures, before);
            }
            catch (Exception applyFailure)
            {
                try
                {
                    gateway.RestoreExact(target, before);
                    gateway.VerifySnapshot(target, before);
                }
                catch (Exception restoreFailure)
                {
                    throw new PivotSemanticPlacementException(
                        "Excel failed to apply the semantic layout and could not restore the exact prior placement.",
                        rollbackCompleted: false,
                        new AggregateException(applyFailure, restoreFailure));
                }

                throw new PivotSemanticPlacementException(
                    "Excel failed to apply the semantic layout; the exact prior placement was restored.",
                    rollbackCompleted: true,
                    applyFailure);
            }
        }

        public void Restore()
        {
            gateway.RestoreExact(target, before);
            gateway.VerifySnapshot(target, before);
        }

        public void VerifyDesired()
        {
            gateway.VerifyDesired(target, plan, namedSets, measures, before);
        }
    }

    internal sealed class PivotSemanticPlacementException : InvalidOperationException
    {
        public PivotSemanticPlacementException(
            string message,
            bool rollbackCompleted,
            Exception innerException)
            : base(message, innerException)
        {
            RollbackCompleted = rollbackCompleted;
        }

        public bool RollbackCompleted { get; }
    }
}
