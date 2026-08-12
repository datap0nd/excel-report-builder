using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using ExcelReportBuilder.Core.PivotPlus;
using ExcelReportBuilder.Excel.PivotPlus.Persistence;

namespace ExcelReportBuilder.Excel.PivotPlus.Measures
{
    /// <summary>
    /// Identifies an existing Values field without trusting its display caption
    /// as a stable identity. The caption itself remains in Excel; only its hash
    /// is carried through a placement request.
    /// </summary>
    public sealed class PivotExistingDataFieldIdentity
    {
        public PivotExistingDataFieldIdentity(
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

        /// <summary>
        /// Hash of the exact native NumberFormat observed during preview. It
        /// binds retries without persisting the user-facing format string.
        /// </summary>
        public string CurrentNumberFormatFingerprint { get; }

        /// <summary>
        /// One-based position observed during preview. This disambiguates
        /// repeated instances of the same source field with the same caption.
        /// </summary>
        public int CurrentPosition { get; }
    }

    /// <summary>
    /// One entry in the complete final Values sequence. Exactly one of
    /// DefinitionId and ExistingDataField is populated.
    /// </summary>
    public sealed class PivotMeasureValuePlacement
    {
        public PivotMeasureValuePlacement(int position, string definitionId)
        {
            Position = position;
            DefinitionId = definitionId ?? string.Empty;
        }

        public PivotMeasureValuePlacement(
            int position,
            PivotExistingDataFieldIdentity existingDataField)
        {
            Position = position;
            ExistingDataField = existingDataField ??
                throw new ArgumentNullException(nameof(existingDataField));
        }

        public int Position { get; }

        public string? DefinitionId { get; }

        public PivotExistingDataFieldIdentity? ExistingDataField { get; }

        public bool IsGeneratedMeasure => ExistingDataField == null;
    }

    /// <summary>
    /// A complete, one-based final Values layout. Existing unowned values must
    /// appear exactly once, which prevents measure authoring from silently
    /// removing a user's native PivotTable field.
    /// </summary>
    public sealed class PivotMeasurePlacementPlan
    {
        public PivotMeasurePlacementPlan(
            IEnumerable<PivotMeasureValuePlacement>? values,
            PivotValuesAxis valuesAxis,
            int valuesPosition)
        {
            Values = new ReadOnlyCollection<PivotMeasureValuePlacement>(
                (values ?? Enumerable.Empty<PivotMeasureValuePlacement>()).ToList());
            ValuesAxis = valuesAxis;
            ValuesPosition = valuesPosition;
        }

        public IReadOnlyList<PivotMeasureValuePlacement> Values { get; }

        public PivotValuesAxis ValuesAxis { get; }

        public int ValuesPosition { get; }
    }

    /// <summary>
    /// Produces the hash used by a placement plan to bind an existing value to
    /// the caption observed during preview. The caption text is never persisted
    /// by PivotTable+ ownership metadata.
    /// </summary>
    public static class PivotMeasurePlacementFingerprint
    {
        public static string CreateCaptionFingerprint(string caption)
        {
            if (caption == null) throw new ArgumentNullException(nameof(caption));
            return PivotPlusFingerprint.Create("pivot.caption.v1", caption);
        }

        public static string CreateNumberFormatFingerprint(string numberFormat)
        {
            if (numberFormat == null) throw new ArgumentNullException(nameof(numberFormat));
            return PivotPlusFingerprint.Create("pivot.number-format.v1", numberFormat);
        }
    }

    public enum PivotModelMeasureApplyStatus
    {
        Applied,
        NoChange,
        RecoveryRequired
    }

    public sealed class PivotModelMeasureApplyResult
    {
        internal PivotModelMeasureApplyResult(
            string applyId,
            PivotModelMeasureApplyStatus status,
            int created,
            int updated,
            int deleted,
            bool undoAvailable)
        {
            ApplyId = applyId;
            Status = status;
            Created = created;
            Updated = updated;
            Deleted = deleted;
            UndoAvailable = undoAvailable;
        }

        public string ApplyId { get; }

        public PivotModelMeasureApplyStatus Status { get; }

        public int Created { get; }

        public int Updated { get; }

        public int Deleted { get; }

        public bool UndoAvailable { get; }
    }

    public sealed class PivotModelMeasureMutationException : InvalidOperationException
    {
        internal PivotModelMeasureMutationException(
            string message,
            bool rollbackCompleted,
            bool recoveryRequired,
            Exception innerException)
            : base(message, innerException)
        {
            RollbackCompleted = rollbackCompleted;
            RecoveryRequired = recoveryRequired;
        }

        public bool RollbackCompleted { get; }

        public bool RecoveryRequired { get; }
    }

    public sealed class PivotModelMeasureUndoUnavailableException : InvalidOperationException
    {
        internal PivotModelMeasureUndoUnavailableException(string message)
            : base(message)
        {
        }
    }
}
