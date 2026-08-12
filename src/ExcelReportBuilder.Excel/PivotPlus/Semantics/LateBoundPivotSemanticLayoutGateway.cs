using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;
using ExcelReportBuilder.Core.PivotPlus;

namespace ExcelReportBuilder.Excel.PivotPlus.Semantics
{
    /// <summary>
    /// Strict late-bound boundary for complete Data Model PivotTable placement.
    /// It changes only Orientation, Position, Caption, and NumberFormat. It
    /// never refreshes, deletes, or creates a CubeField, ModelMeasure, or
    /// CalculatedMember.
    /// </summary>
    internal sealed class LateBoundPivotSemanticLayoutGateway : IPivotSemanticLayoutGateway
    {
        private const int OrientationHidden = 0;
        private const int OrientationRow = 1;
        private const int OrientationColumn = 2;
        private const int OrientationPage = 3;
        private const int OrientationData = 4;
        private const int DataModelConnectionType = 7;
        private const int CubeFieldTypeMeasure = 2;
        private const int CubeFieldTypeSet = 3;

        private const int MaximumAxisFields = 256;
        private const int MaximumDataFields = 512;
        private const int MaximumFilterFields = 256;
        private const int MaximumFilterItems = 4096;
        private const int MaximumCubeFields = 4096;
        private const int MaximumNameCharacters = 1024;
        private const int MaximumUniqueNameCharacters = 2048;
        private const int MaximumFormatCharacters = 255;
        private const int MaximumFilterTokenCharacters = 4096;

        public BoundPivotSemanticLayoutTarget Bind(
            object workbook,
            object pivotTable,
            PivotTableContext context)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            if (context == null) throw new ArgumentNullException(nameof(context));

            PivotLayoutDefinition definition = context.Definition;
            const PivotCapability required = PivotCapability.DataModel |
                                             PivotCapability.NativeFieldPlacement;
            if (!context.IsConnected ||
                !context.SourceFieldsComplete ||
                definition.Source.Kind != PivotSourceKind.DataModel ||
                (definition.Source.Capabilities & required) != required)
            {
                throw new NotSupportedException(
                    "Semantic placement requires the selected native workbook Data Model PivotTable.");
            }

            dynamic pivot = pivotTable;
            string pivotName = ReadBoundedRequiredString(
                () => (object?)pivot.Name,
                MaximumNameCharacters,
                "selected PivotTable name");
            object worksheet = ReadRequired(
                () => (object?)pivot.Parent,
                "Excel did not expose the selected PivotTable worksheet.");
            dynamic nativeWorksheet = worksheet;
            string worksheetName = ReadBoundedRequiredString(
                () => (object?)nativeWorksheet.Name,
                MaximumNameCharacters,
                "selected PivotTable worksheet name");
            object liveWorkbook = ReadRequired(
                () => (object?)nativeWorksheet.Parent,
                "Excel did not expose the selected PivotTable workbook.");
            if (!ComObjectIdentity.AreSame(workbook, liveWorkbook))
            {
                throw new InvalidOperationException(
                    "The supplied workbook is not the selected PivotTable's workbook.");
            }

            PivotTargetIdentity expected = definition.Target;
            if (!string.Equals(
                    pivotName,
                    expected.PivotTableName,
                    StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(
                    worksheetName,
                    expected.WorksheetName,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "The selected PivotTable no longer matches the discovered target.");
            }

            string workbookId = new StoredWorkbookIdentityResolver().Resolve(workbook);
            if (!string.Equals(workbookId, expected.WorkbookId, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The selected workbook no longer matches the discovered path-free identity.");
            }

            object cache = ReadPivotCache(pivotTable);
            dynamic nativeCache = cache;
            if (!ReadRequiredBoolean(
                    () => (object?)nativeCache.OLAP,
                    "selected PivotCache.OLAP"))
            {
                throw new NotSupportedException(
                    "The selected PivotTable is not a Data Model PivotTable.");
            }

            dynamic nativeWorkbook = workbook;
            object model = ReadRequired(
                () => (object?)nativeWorkbook.Model,
                "Excel did not expose the workbook Data Model.");
            dynamic nativeModel = model;
            object modelConnection = ReadRequired(
                () => (object?)nativeModel.DataModelConnection,
                "Excel did not expose Workbook.Model.DataModelConnection.");
            dynamic nativeModelConnection = modelConnection;
            if (ReadRequiredInt(
                    () => (object?)nativeModelConnection.Type,
                    "Data Model connection type") != DataModelConnectionType)
            {
                throw new InvalidOperationException(
                    "Workbook.Model.DataModelConnection is not the special Data Model connection.");
            }

            object cacheConnection = ReadRequired(
                () => (object?)nativeCache.WorkbookConnection,
                "Excel did not expose the selected PivotCache workbook connection.");
            if (!ComObjectIdentity.AreSame(cacheConnection, modelConnection))
            {
                throw new NotSupportedException(
                    "The selected PivotTable does not use this workbook's exact Data Model connection.");
            }

            DemandCollectionCapability(pivotTable, "CubeFields", MaximumCubeFields);
            DemandCollectionCapability(pivotTable, "RowFields", MaximumAxisFields + 1);
            DemandCollectionCapability(pivotTable, "ColumnFields", MaximumAxisFields + 1);
            DemandCollectionCapability(pivotTable, "DataFields", MaximumDataFields);
            DemandCollectionCapability(pivotTable, "PageFields", MaximumFilterFields);

            return new BoundPivotSemanticLayoutTarget(
                workbook,
                pivotTable,
                model,
                modelConnection,
                expected);
        }

        public PivotSemanticLayoutSnapshot Capture(
            BoundPivotSemanticLayoutTarget target)
        {
            if (target == null) throw new ArgumentNullException(nameof(target));
            return ReadHostState(target, requireContiguousPositions: true).Snapshot;
        }

        public PivotSemanticPreparedPlacement Prepare(
            BoundPivotSemanticLayoutTarget target,
            PivotSemanticLayoutPlan plan,
            IReadOnlyDictionary<string, string> namedSetUniqueNamesByDefinitionId,
            IReadOnlyDictionary<string, string> measureUniqueNamesByDefinitionId,
            PivotSemanticLayoutSnapshot before)
        {
            if (target == null) throw new ArgumentNullException(nameof(target));
            if (plan == null) throw new ArgumentNullException(nameof(plan));
            if (namedSetUniqueNamesByDefinitionId == null)
            {
                throw new ArgumentNullException(nameof(namedSetUniqueNamesByDefinitionId));
            }

            if (measureUniqueNamesByDefinitionId == null)
            {
                throw new ArgumentNullException(nameof(measureUniqueNamesByDefinitionId));
            }

            if (before == null) throw new ArgumentNullException(nameof(before));
            DemandSnapshotIdentity(target, before);
            DemandSnapshotIntegrity(before);
            PivotSemanticLayoutSnapshot current = Capture(target);
            DemandSameLayout(current, before, "The selected PivotTable layout changed after preview.");

            var namedSets = new ReadOnlyDictionary<string, string>(
                namedSetUniqueNamesByDefinitionId.ToDictionary(
                    item => item.Key,
                    item => item.Value,
                    StringComparer.Ordinal));
            var measures = new ReadOnlyDictionary<string, string>(
                measureUniqueNamesByDefinitionId.ToDictionary(
                    item => item.Key,
                    item => item.Value,
                    StringComparer.Ordinal));
            PivotSemanticLayoutCanonical.ValidatePlanAndMappings(
                plan,
                namedSets,
                measures,
                before);
            return new PivotSemanticPreparedPlacement(
                this,
                target,
                plan,
                namedSets,
                measures,
                before);
        }

        internal PivotSemanticPreparedPlacement PrepareRecoveredFinal(
            BoundPivotSemanticLayoutTarget target,
            PivotSemanticLayoutPlan plan,
            IReadOnlyDictionary<string, string> namedSetUniqueNamesByDefinitionId,
            IReadOnlyDictionary<string, string> measureUniqueNamesByDefinitionId,
            PivotSemanticLayoutSnapshot originalBefore)
        {
            if (target == null) throw new ArgumentNullException(nameof(target));
            if (plan == null) throw new ArgumentNullException(nameof(plan));
            if (namedSetUniqueNamesByDefinitionId == null)
            {
                throw new ArgumentNullException(nameof(namedSetUniqueNamesByDefinitionId));
            }

            if (measureUniqueNamesByDefinitionId == null)
            {
                throw new ArgumentNullException(nameof(measureUniqueNamesByDefinitionId));
            }

            if (originalBefore == null)
            {
                throw new ArgumentNullException(nameof(originalBefore));
            }

            DemandSnapshotIdentity(target, originalBefore);
            DemandSnapshotIntegrity(originalBefore);
            var namedSets = new ReadOnlyDictionary<string, string>(
                namedSetUniqueNamesByDefinitionId.ToDictionary(
                    item => item.Key,
                    item => item.Value,
                    StringComparer.Ordinal));
            var measures = new ReadOnlyDictionary<string, string>(
                measureUniqueNamesByDefinitionId.ToDictionary(
                    item => item.Key,
                    item => item.Value,
                    StringComparer.Ordinal));
            VerifyDesired(target, plan, namedSets, measures, originalBefore);
            return new PivotSemanticPreparedPlacement(
                this,
                target,
                plan,
                namedSets,
                measures,
                originalBefore,
                alreadyApplied: true);
        }

        internal void ApplyExact(
            BoundPivotSemanticLayoutTarget target,
            PivotSemanticLayoutPlan plan,
            IReadOnlyDictionary<string, string> namedSets,
            IReadOnlyDictionary<string, string> measures,
            PivotSemanticLayoutSnapshot before)
        {
            DemandSnapshotIdentity(target, before);
            PivotSemanticLayoutCanonical.ValidatePlanAndMappings(
                plan,
                namedSets,
                measures,
                before);
            HostState current = ReadHostState(target, requireContiguousPositions: true);
            DemandSameLayout(
                current.Snapshot,
                before,
                "The selected PivotTable layout changed before semantic placement.");

            var retainedAxis = new HashSet<object>(ReferenceEqualityComparer.Instance);
            foreach (PivotSemanticAxisPlacement placement in plan.Rows.Concat(plan.Columns))
            {
                if (!placement.IsGeneratedNamedSet)
                {
                    retainedAxis.Add(ResolveExistingAxis(current, placement.ExistingField!).CubeField);
                }
            }

            foreach (AxisFieldHandle field in current.Rows.Concat(current.Columns))
            {
                if (!retainedAxis.Contains(field.CubeField))
                {
                    SetOrientation(field.CubeField, OrientationHidden, "axis field removal");
                }
            }

            foreach (PivotSemanticAxisPlacement placement in plan.Rows)
            {
                OrientAxisTarget(
                    target,
                    current,
                    placement,
                    namedSets,
                    OrientationRow);
            }

            foreach (PivotSemanticAxisPlacement placement in plan.Columns)
            {
                OrientAxisTarget(
                    target,
                    current,
                    placement,
                    namedSets,
                    OrientationColumn);
            }

            ApplyValues(target, current, plan, measures);
            ApplyValuesAxis(target.PivotTable, plan);
            ApplyAxisPositions(target, plan, namedSets);
            RestoreExistingPresentation(target, plan, before);
            DemandFiltersUnchanged(target, before);
        }

        internal void RestoreExact(
            BoundPivotSemanticLayoutTarget target,
            PivotSemanticLayoutSnapshot before)
        {
            if (target == null) throw new ArgumentNullException(nameof(target));
            if (before == null) throw new ArgumentNullException(nameof(before));
            DemandSnapshotIdentity(target, before);
            DemandSnapshotIntegrity(before);
            DemandFiltersUnchanged(target, before);

            HostState current = ReadHostState(target, requireContiguousPositions: false);
            var expectedAxisNames = new HashSet<string>(
                before.Rows.Concat(before.Columns).Select(item => item.UniqueName),
                StringComparer.OrdinalIgnoreCase);
            foreach (AxisFieldHandle field in current.Rows.Concat(current.Columns))
            {
                if (!expectedAxisNames.Contains(field.Snapshot.UniqueName))
                {
                    SetOrientation(field.CubeField, OrientationHidden, "axis rollback removal");
                }
            }

            foreach (PivotSemanticAxisFieldSnapshot field in before.Rows)
            {
                CubeFieldHandle cube = ResolveCubeField(
                    target.PivotTable,
                    field.UniqueName,
                    field.CubeFieldType);
                SetOrientation(cube.Native, OrientationRow, "row-axis rollback");
            }

            foreach (PivotSemanticAxisFieldSnapshot field in before.Columns)
            {
                CubeFieldHandle cube = ResolveCubeField(
                    target.PivotTable,
                    field.UniqueName,
                    field.CubeFieldType);
                SetOrientation(cube.Native, OrientationColumn, "column-axis rollback");
            }

            RestoreValues(target, before);
            ApplyValuesAxis(
                target.PivotTable,
                before.Values.Count,
                before.ValuesAxis,
                before.ValuesPosition,
                before.Rows.Count,
                before.Columns.Count);
            ApplySnapshotAxisPositions(target, before);
            RestoreSnapshotPresentation(target, before);
            DemandFiltersUnchanged(target, before);
        }

        internal void VerifySnapshot(
            BoundPivotSemanticLayoutTarget target,
            PivotSemanticLayoutSnapshot expected)
        {
            if (target == null) throw new ArgumentNullException(nameof(target));
            if (expected == null) throw new ArgumentNullException(nameof(expected));
            DemandSnapshotIdentity(target, expected);
            DemandSnapshotIntegrity(expected);
            DemandSameLayout(
                Capture(target),
                expected,
                "Excel did not restore the exact semantic layout.");
        }

        internal void VerifyDesired(
            BoundPivotSemanticLayoutTarget target,
            PivotSemanticLayoutPlan plan,
            IReadOnlyDictionary<string, string> namedSets,
            IReadOnlyDictionary<string, string> measures,
            PivotSemanticLayoutSnapshot before)
        {
            if (target == null) throw new ArgumentNullException(nameof(target));
            PivotSemanticLayoutCanonical.ValidatePlanAndMappings(
                plan,
                namedSets,
                measures,
                before);
            PivotSemanticLayoutSnapshot live = Capture(target);
            if (!string.Equals(
                    live.FilterFingerprint,
                    before.FilterFingerprint,
                    StringComparison.Ordinal) ||
                live.Rows.Count != plan.Rows.Count ||
                live.Columns.Count != plan.Columns.Count ||
                live.Values.Count != plan.Values.Count ||
                live.ValuesAxis != plan.ValuesAxis ||
                live.ValuesPosition != plan.ValuesPosition ||
                !MatchesDesiredAxis(live.Rows, plan.Rows, namedSets) ||
                !MatchesDesiredAxis(live.Columns, plan.Columns, namedSets) ||
                !MatchesDesiredValues(live.Values, plan.Values, measures))
            {
                throw new InvalidOperationException(
                    "Excel did not expose the exact requested semantic layout with unchanged Filters.");
            }
        }

        private static bool MatchesDesiredAxis(
            IReadOnlyList<PivotSemanticAxisFieldSnapshot> live,
            IReadOnlyList<PivotSemanticAxisPlacement> desired,
            IReadOnlyDictionary<string, string> namedSets)
        {
            for (var index = 0; index < desired.Count; index++)
            {
                PivotSemanticAxisPlacement placement = desired[index];
                PivotSemanticAxisFieldSnapshot field = live[index];
                string uniqueName = placement.IsGeneratedNamedSet
                    ? namedSets[placement.DefinitionId!]
                    : placement.ExistingField!.UniqueName;
                if (field.Position != index + 1 ||
                    !string.Equals(field.UniqueName, uniqueName, StringComparison.Ordinal) ||
                    (!placement.IsGeneratedNamedSet &&
                     !string.Equals(
                         field.CaptionFingerprint,
                         placement.ExistingField!.CurrentCaptionFingerprint,
                         StringComparison.Ordinal)))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool MatchesDesiredValues(
            IReadOnlyList<PivotSemanticValueFieldSnapshot> live,
            IReadOnlyList<PivotSemanticValuePlacement> desired,
            IReadOnlyDictionary<string, string> measures)
        {
            for (var index = 0; index < desired.Count; index++)
            {
                PivotSemanticValuePlacement placement = desired[index];
                PivotSemanticValueFieldSnapshot field = live[index];
                string uniqueName = placement.IsGeneratedMeasure
                    ? measures[placement.DefinitionId!]
                    : placement.ExistingDataField!.UniqueName;
                if (field.Position != index + 1 ||
                    !string.Equals(field.UniqueName, uniqueName, StringComparison.Ordinal))
                {
                    return false;
                }

                if (!placement.IsGeneratedMeasure &&
                    (!string.Equals(
                         field.CaptionFingerprint,
                         placement.ExistingDataField!.CurrentCaptionFingerprint,
                         StringComparison.Ordinal) ||
                     !string.Equals(
                         field.NumberFormatFingerprint,
                         placement.ExistingDataField.CurrentNumberFormatFingerprint,
                         StringComparison.Ordinal)))
                {
                    return false;
                }
            }

            return true;
        }

        private static void OrientAxisTarget(
            BoundPivotSemanticLayoutTarget target,
            HostState before,
            PivotSemanticAxisPlacement placement,
            IReadOnlyDictionary<string, string> namedSets,
            int orientation)
        {
            if (!placement.IsGeneratedNamedSet)
            {
                AxisFieldHandle existing = ResolveExistingAxis(
                    before,
                    placement.ExistingField!);
                SetOrientation(existing.CubeField, orientation, "existing axis placement");
                return;
            }

            CubeFieldHandle generated = ResolveCubeField(
                target.PivotTable,
                namedSets[placement.DefinitionId!],
                CubeFieldTypeSet);
            SetOrientation(generated.Native, orientation, "generated named-set placement");
        }

        private static void ApplyValues(
            BoundPivotSemanticLayoutTarget target,
            HostState before,
            PivotSemanticLayoutPlan plan,
            IReadOnlyDictionary<string, string> measures)
        {
            var retained = new HashSet<object>(ReferenceEqualityComparer.Instance);
            foreach (PivotSemanticValuePlacement placement in plan.Values)
            {
                if (!placement.IsGeneratedMeasure)
                {
                    retained.Add(ResolveExistingValue(before, placement.ExistingDataField!).NativeField);
                }
            }

            foreach (ValueFieldHandle field in before.Values)
            {
                if (!retained.Contains(field.NativeField))
                {
                    SetOrientation(field.NativeField, OrientationHidden, "Values field removal");
                }
            }

            // Retained existing DataField occurrences stay visible. Reorient
            // only generated measure CubeFields so repeated existing Values
            // occurrences are neither manufactured nor collapsed.
            foreach (PivotSemanticValuePlacement placement in plan.Values
                         .Where(item => item.IsGeneratedMeasure))
            {
                CubeFieldHandle generated = ResolveCubeField(
                    target.PivotTable,
                    measures[placement.DefinitionId!],
                    CubeFieldTypeMeasure);
                SetOrientation(generated.Native, OrientationData, "generated measure placement");
            }

            IReadOnlyList<ValueFieldHandle> live = ReadValueFields(
                target.PivotTable,
                requireContiguousPositions: false);
            var assigned = new HashSet<object>(ReferenceEqualityComparer.Instance);
            foreach (PivotSemanticValuePlacement placement in plan.Values
                         .OrderBy(item => item.Position))
            {
                ValueFieldHandle field;
                if (placement.IsGeneratedMeasure)
                {
                    string name = measures[placement.DefinitionId!];
                    field = live.SingleOrDefault(item =>
                        !assigned.Contains(item.NativeField) &&
                        string.Equals(
                            item.Snapshot.UniqueName,
                            name,
                            StringComparison.Ordinal)) ??
                        throw new InvalidOperationException(
                            "Excel did not materialize the generated measure DataField exactly once.");
                }
                else
                {
                    ValueFieldHandle prior = ResolveExistingValue(
                        before,
                        placement.ExistingDataField!);
                    field = live.SingleOrDefault(item =>
                        !assigned.Contains(item.NativeField) &&
                        ComObjectIdentity.AreSame(item.NativeField, prior.NativeField)) ??
                        live.SingleOrDefault(item =>
                            !assigned.Contains(item.NativeField) &&
                            PivotSemanticLayoutCanonical.Matches(
                                item.Snapshot,
                                placement.ExistingDataField!)) ??
                        throw new InvalidOperationException(
                            "Excel did not preserve an existing Values occurrence.");
                }

                assigned.Add(field.NativeField);
                SetPosition(field.NativeField, placement.Position, "Values field position");
            }

            if (live.Count != plan.Values.Count || assigned.Count != live.Count)
            {
                throw new InvalidOperationException(
                    "Excel exposed an unexpected Values field after semantic placement.");
            }
        }

        private static void RestoreValues(
            BoundPivotSemanticLayoutTarget target,
            PivotSemanticLayoutSnapshot before)
        {
            IReadOnlyList<ValueFieldHandle> current = ReadValueFields(
                target.PivotTable,
                requireContiguousPositions: false);
            var expectedNames = new HashSet<string>(
                before.Values.Select(item => item.UniqueName),
                StringComparer.OrdinalIgnoreCase);
            foreach (ValueFieldHandle field in current)
            {
                if (!expectedNames.Contains(field.Snapshot.UniqueName))
                {
                    SetOrientation(field.NativeField, OrientationHidden, "Values rollback removal");
                }
            }

            current = ReadValueFields(
                target.PivotTable,
                requireContiguousPositions: false);
            foreach (PivotSemanticValueFieldSnapshot expected in before.Values)
            {
                if (!current.Any(item =>
                        string.Equals(
                            item.Snapshot.UniqueName,
                            expected.UniqueName,
                            StringComparison.OrdinalIgnoreCase)))
                {
                    CubeFieldHandle cube = ResolveCubeField(
                        target.PivotTable,
                        expected.UniqueName,
                        CubeFieldTypeMeasure);
                    SetOrientation(cube.Native, OrientationData, "Values rollback placement");
                    current = ReadValueFields(
                        target.PivotTable,
                        requireContiguousPositions: false);
                }
            }

            var available = current.ToList();
            foreach (PivotSemanticValueFieldSnapshot expected in before.Values
                         .OrderBy(item => item.Position))
            {
                ValueFieldHandle field = available
                    .Where(item => string.Equals(
                        item.Snapshot.UniqueName,
                        expected.UniqueName,
                        StringComparison.OrdinalIgnoreCase))
                    .OrderByDescending(item =>
                        string.Equals(
                            item.Snapshot.CaptionFingerprint,
                            expected.CaptionFingerprint,
                            StringComparison.Ordinal) &&
                        string.Equals(
                            item.Snapshot.NumberFormatFingerprint,
                            expected.NumberFormatFingerprint,
                            StringComparison.Ordinal))
                    .ThenBy(item => Math.Abs(item.Snapshot.Position - expected.Position))
                    .FirstOrDefault() ?? throw new InvalidOperationException(
                        "Excel did not restore an exact prior Values occurrence.");
                available.Remove(field);
                SetPosition(field.NativeField, expected.Position, "Values rollback position");
                WritePresentation(
                    field.NativeField,
                    field.Snapshot.Caption,
                    expected.Caption,
                    field.Snapshot.NumberFormat,
                    expected.NumberFormat);
            }

            if (available.Count != 0)
            {
                throw new InvalidOperationException(
                    "Excel retained an unexpected Values occurrence during rollback.");
            }
        }

        private static void RestoreExistingPresentation(
            BoundPivotSemanticLayoutTarget target,
            PivotSemanticLayoutPlan plan,
            PivotSemanticLayoutSnapshot before)
        {
            HostState current = ReadHostState(target, requireContiguousPositions: true);
            foreach (PivotSemanticAxisPlacement placement in plan.Rows.Concat(plan.Columns))
            {
                if (placement.IsGeneratedNamedSet) continue;
                PivotSemanticAxisFieldSnapshot expected = before.Rows.Concat(before.Columns)
                    .Single(item => PivotSemanticLayoutCanonical.Matches(
                        item,
                        placement.ExistingField!));
                AxisFieldHandle live = current.Rows.Concat(current.Columns).Single(item =>
                    string.Equals(
                        item.Snapshot.UniqueName,
                        expected.UniqueName,
                        StringComparison.OrdinalIgnoreCase));
                WriteCaption(live.NativeField, live.Snapshot.Caption, expected.Caption);
            }

            foreach (PivotSemanticValuePlacement placement in plan.Values)
            {
                if (placement.IsGeneratedMeasure) continue;
                PivotSemanticValueFieldSnapshot expected = before.Values.Single(item =>
                    PivotSemanticLayoutCanonical.Matches(
                        item,
                        placement.ExistingDataField!));
                ValueFieldHandle live = current.Values.Single(item =>
                    item.Snapshot.Position == placement.Position &&
                    string.Equals(
                        item.Snapshot.UniqueName,
                        expected.UniqueName,
                        StringComparison.OrdinalIgnoreCase));
                WritePresentation(
                    live.NativeField,
                    live.Snapshot.Caption,
                    expected.Caption,
                    live.Snapshot.NumberFormat,
                    expected.NumberFormat);
            }
        }

        private static void RestoreSnapshotPresentation(
            BoundPivotSemanticLayoutTarget target,
            PivotSemanticLayoutSnapshot before)
        {
            HostState current = ReadHostState(target, requireContiguousPositions: true);
            foreach (PivotSemanticAxisFieldSnapshot expected in before.Rows.Concat(before.Columns))
            {
                AxisFieldHandle live = current.Rows.Concat(current.Columns).Single(item =>
                    item.Snapshot.Area == expected.Area &&
                    item.Snapshot.Position == expected.Position &&
                    string.Equals(
                        item.Snapshot.UniqueName,
                        expected.UniqueName,
                        StringComparison.OrdinalIgnoreCase));
                WriteCaption(live.NativeField, live.Snapshot.Caption, expected.Caption);
            }

            foreach (PivotSemanticValueFieldSnapshot expected in before.Values)
            {
                ValueFieldHandle live = current.Values.Single(item =>
                    item.Snapshot.Position == expected.Position &&
                    string.Equals(
                        item.Snapshot.UniqueName,
                        expected.UniqueName,
                        StringComparison.OrdinalIgnoreCase));
                WritePresentation(
                    live.NativeField,
                    live.Snapshot.Caption,
                    expected.Caption,
                    live.Snapshot.NumberFormat,
                    expected.NumberFormat);
            }
        }

        private static void ApplyValuesAxis(
            object pivotTable,
            PivotSemanticLayoutPlan plan)
        {
            ApplyValuesAxis(
                pivotTable,
                plan.Values.Count,
                plan.ValuesAxis,
                plan.ValuesPosition,
                plan.Rows.Count,
                plan.Columns.Count);
        }

        private static void ApplyValuesAxis(
            object pivotTable,
            int valueCount,
            PivotValuesAxis axis,
            int position,
            int rowCount,
            int columnCount)
        {
            if (valueCount <= 1)
            {
                if (axis != PivotValuesAxis.Automatic || position != 1)
                {
                    throw new InvalidOperationException(
                        "A zero- or single-value layout must use the automatic Values axis.");
                }

                return;
            }

            int maximum = axis == PivotValuesAxis.Rows
                ? rowCount + 1
                : columnCount + 1;
            if ((axis != PivotValuesAxis.Rows && axis != PivotValuesAxis.Columns) ||
                position <= 0 || position > maximum)
            {
                throw new InvalidOperationException(
                    "The Values pseudo-field has an invalid semantic axis position.");
            }

            dynamic pivot = pivotTable;
            object dataPivotField = ReadRequired(
                () => (object?)pivot.DataPivotField,
                "Excel did not expose the Values pseudo-field for placement.");
            SetOrientation(
                dataPivotField,
                axis == PivotValuesAxis.Rows ? OrientationRow : OrientationColumn,
                "Values pseudo-field orientation");
            SetPosition(dataPivotField, position, "Values pseudo-field position");
        }

        private static void ApplyAxisPositions(
            BoundPivotSemanticLayoutTarget target,
            PivotSemanticLayoutPlan plan,
            IReadOnlyDictionary<string, string> namedSets)
        {
            foreach (PivotSemanticAxisPlacement placement in plan.Rows.OrderBy(item => item.Position))
            {
                string name = placement.IsGeneratedNamedSet
                    ? namedSets[placement.DefinitionId!]
                    : placement.ExistingField!.UniqueName;
                CubeFieldHandle cube = ResolveCubeField(target.PivotTable, name, expectedType: null);
                SetPosition(
                    cube.Native,
                    RawAxisPosition(
                        placement.Position,
                        PivotValuesAxis.Rows,
                        plan.Values.Count,
                        plan.ValuesAxis,
                        plan.ValuesPosition),
                    "row-axis field position");
            }

            foreach (PivotSemanticAxisPlacement placement in plan.Columns.OrderBy(item => item.Position))
            {
                string name = placement.IsGeneratedNamedSet
                    ? namedSets[placement.DefinitionId!]
                    : placement.ExistingField!.UniqueName;
                CubeFieldHandle cube = ResolveCubeField(target.PivotTable, name, expectedType: null);
                SetPosition(
                    cube.Native,
                    RawAxisPosition(
                        placement.Position,
                        PivotValuesAxis.Columns,
                        plan.Values.Count,
                        plan.ValuesAxis,
                        plan.ValuesPosition),
                    "column-axis field position");
            }
        }

        private static void ApplySnapshotAxisPositions(
            BoundPivotSemanticLayoutTarget target,
            PivotSemanticLayoutSnapshot before)
        {
            foreach (PivotSemanticAxisFieldSnapshot field in before.Rows)
            {
                CubeFieldHandle cube = ResolveCubeField(
                    target.PivotTable,
                    field.UniqueName,
                    field.CubeFieldType);
                SetPosition(
                    cube.Native,
                    RawAxisPosition(
                        field.Position,
                        PivotValuesAxis.Rows,
                        before.Values.Count,
                        before.ValuesAxis,
                        before.ValuesPosition),
                    "row-axis rollback position");
            }

            foreach (PivotSemanticAxisFieldSnapshot field in before.Columns)
            {
                CubeFieldHandle cube = ResolveCubeField(
                    target.PivotTable,
                    field.UniqueName,
                    field.CubeFieldType);
                SetPosition(
                    cube.Native,
                    RawAxisPosition(
                        field.Position,
                        PivotValuesAxis.Columns,
                        before.Values.Count,
                        before.ValuesAxis,
                        before.ValuesPosition),
                    "column-axis rollback position");
            }
        }

        private static int RawAxisPosition(
            int normalizedPosition,
            PivotValuesAxis fieldAxis,
            int valueCount,
            PivotValuesAxis valuesAxis,
            int valuesPosition)
        {
            return valueCount > 1 &&
                   fieldAxis == valuesAxis &&
                   normalizedPosition >= valuesPosition
                ? normalizedPosition + 1
                : normalizedPosition;
        }

        private static AxisFieldHandle ResolveExistingAxis(
            HostState state,
            PivotExistingAxisFieldIdentity identity)
        {
            return state.Rows.Concat(state.Columns).SingleOrDefault(item =>
                PivotSemanticLayoutCanonical.Matches(item.Snapshot, identity)) ??
                throw new InvalidOperationException(
                    "An existing axis field no longer matches its preview identity.");
        }

        private static ValueFieldHandle ResolveExistingValue(
            HostState state,
            PivotExistingSemanticValueIdentity identity)
        {
            return state.Values.SingleOrDefault(item =>
                PivotSemanticLayoutCanonical.Matches(item.Snapshot, identity)) ??
                throw new InvalidOperationException(
                    "An existing Values field no longer matches its preview identity.");
        }

        private static HostState ReadHostState(
            BoundPivotSemanticLayoutTarget target,
            bool requireContiguousPositions)
        {
            DemandStillBound(target);
            IReadOnlyList<ValueFieldHandle> values = ReadValueFields(
                target.PivotTable,
                requireContiguousPositions);
            object? dataPivotField = null;
            PivotValuesAxis valuesAxis = PivotValuesAxis.Automatic;
            var valuesPosition = 1;
            if (values.Count > 1)
            {
                dynamic pivot = target.PivotTable;
                dataPivotField = ReadRequired(
                    () => (object?)pivot.DataPivotField,
                    "Excel did not expose the Values pseudo-field for a multi-value PivotTable.");
                int orientation = ReadRequiredInt(
                    () => (object?)((dynamic)dataPivotField).Orientation,
                    "Values pseudo-field orientation");
                valuesAxis = orientation == OrientationRow
                    ? PivotValuesAxis.Rows
                    : orientation == OrientationColumn
                        ? PivotValuesAxis.Columns
                        : throw new NotSupportedException(
                            "The Values pseudo-field is outside Rows and Columns.");
                valuesPosition = ReadRequiredPositiveInt(
                    () => (object?)((dynamic)dataPivotField).Position,
                    "Values pseudo-field position");
            }

            IReadOnlyList<AxisFieldHandle> rows = ReadAxisFields(
                target.PivotTable,
                "RowFields",
                PivotFieldArea.Row,
                OrientationRow,
                dataPivotField,
                requireContiguousPositions,
                out int rowPseudoCount);
            IReadOnlyList<AxisFieldHandle> columns = ReadAxisFields(
                target.PivotTable,
                "ColumnFields",
                PivotFieldArea.Column,
                OrientationColumn,
                dataPivotField,
                requireContiguousPositions,
                out int columnPseudoCount);
            if (values.Count > 1)
            {
                int expectedRows = valuesAxis == PivotValuesAxis.Rows ? 1 : 0;
                int expectedColumns = valuesAxis == PivotValuesAxis.Columns ? 1 : 0;
                if (rowPseudoCount != expectedRows || columnPseudoCount != expectedColumns)
                {
                    throw new NotSupportedException(
                        "Excel did not expose the Values pseudo-field exactly once on its declared axis.");
                }

                int axisCount = valuesAxis == PivotValuesAxis.Rows
                    ? rows.Count
                    : columns.Count;
                if (valuesPosition > axisCount + 1)
                {
                    throw new NotSupportedException(
                        "The Values pseudo-field position is outside its native axis.");
                }
            }

            if (rows.Concat(columns)
                .GroupBy(item => item.Snapshot.UniqueName, StringComparer.OrdinalIgnoreCase)
                .Any(group => group.Count() != 1))
            {
                throw new NotSupportedException(
                    "Excel exposed multiple visible PivotFields for one hierarchy CubeField; semantic placement will not pretend they are independently orientable.");
            }

            IReadOnlyList<PivotSemanticFilterFieldSnapshot> filters = ReadFilters(
                target.PivotTable);
            string filterFingerprint = PivotSemanticLayoutCanonical.CreateFilterFingerprint(
                filters);
            string layoutFingerprint = PivotSemanticLayoutCanonical.CreateLayoutFingerprint(
                rows.Select(item => item.Snapshot),
                columns.Select(item => item.Snapshot),
                values.Select(item => item.Snapshot),
                valuesAxis,
                valuesPosition,
                filterFingerprint);
            var snapshot = new PivotSemanticLayoutSnapshot(
                target.Identity,
                rows.Select(item => item.Snapshot),
                columns.Select(item => item.Snapshot),
                values.Select(item => item.Snapshot),
                filters,
                valuesAxis,
                valuesPosition,
                filterFingerprint,
                layoutFingerprint);
            return new HostState(rows, columns, values, dataPivotField, snapshot);
        }

        private static IReadOnlyList<AxisFieldHandle> ReadAxisFields(
            object pivotTable,
            string collectionName,
            PivotFieldArea area,
            int requiredOrientation,
            object? dataPivotField,
            bool requireContiguousPositions,
            out int pseudoCount)
        {
            dynamic pivot = pivotTable;
            object collection = collectionName == "RowFields"
                ? ReadRequired(
                    () => (object?)pivot.RowFields,
                    "Excel did not expose PivotTable RowFields.")
                : ReadRequired(
                    () => (object?)pivot.ColumnFields,
                    "Excel did not expose PivotTable ColumnFields.");
            IReadOnlyList<object> items = ReadCollection(
                collection,
                MaximumAxisFields + 1,
                "PivotTable " + collectionName);
            var raw = new List<RawAxisField>();
            pseudoCount = 0;
            foreach (object fieldObject in items)
            {
                if (dataPivotField != null &&
                    ComObjectIdentity.AreSame(fieldObject, dataPivotField))
                {
                    pseudoCount++;
                    continue;
                }

                dynamic field = fieldObject;
                if (ReadRequiredInt(
                        () => (object?)field.Orientation,
                        collectionName + " field orientation") != requiredOrientation)
                {
                    throw new NotSupportedException(
                        "Excel exposed a field in the wrong PivotTable axis collection.");
                }

                object cubeObject = ReadRequired(
                    () => (object?)field.CubeField,
                    "Excel did not expose an axis field CubeField.");
                dynamic cube = cubeObject;
                string uniqueName = ReadBoundedRequiredString(
                    () => (object?)cube.Name,
                    MaximumUniqueNameCharacters,
                    "axis CubeField unique name");
                int cubeType = ReadRequiredInt(
                    () => (object?)cube.CubeFieldType,
                    "axis CubeField type");
                if (cubeType == CubeFieldTypeMeasure)
                {
                    throw new NotSupportedException(
                        "Excel exposed a measure as a regular row or column field.");
                }

                if (ReadRequiredInt(
                        () => (object?)cube.Orientation,
                        "axis CubeField orientation") != requiredOrientation)
                {
                    throw new NotSupportedException(
                        "Excel exposed a PivotField whose hierarchy CubeField is on another axis.");
                }

                string caption = ReadBoundedRequiredString(
                    () => (object?)field.Caption,
                    MaximumNameCharacters,
                    "axis field caption");
                int rawPosition = ReadRequiredPositiveInt(
                    () => (object?)field.Position,
                    "axis field position");
                raw.Add(new RawAxisField(
                    fieldObject,
                    cubeObject,
                    uniqueName,
                    caption,
                    rawPosition,
                    cubeType));
            }

            int expectedNativeCount = raw.Count + pseudoCount;
            if (raw.GroupBy(
                    item => item.UniqueName,
                    StringComparer.OrdinalIgnoreCase)
                .Any(group => group.Count() != 1))
            {
                throw new NotSupportedException(
                    "Excel exposed multiple visible PivotFields for one hierarchy CubeField; semantic placement will not pretend they are independently orientable.");
            }

            if (requireContiguousPositions &&
                items.Select(item => ReadRequiredPositiveInt(
                        () => (object?)((dynamic)item).Position,
                        collectionName + " native position"))
                    .OrderBy(position => position)
                    .Where((position, index) => position != index + 1)
                    .Any())
            {
                throw new NotSupportedException(
                    "Excel exposed duplicate or non-contiguous native axis positions.");
            }

            if (expectedNativeCount != items.Count)
            {
                throw new InvalidOperationException(
                    "Excel exposed an inconsistent native axis collection.");
            }

            var result = new List<AxisFieldHandle>(raw.Count);
            int normalized = 0;
            foreach (RawAxisField field in raw.OrderBy(item => item.RawPosition))
            {
                normalized++;
                result.Add(new AxisFieldHandle(
                    field.NativeField,
                    field.CubeField,
                    field.RawPosition,
                    new PivotSemanticAxisFieldSnapshot(
                        field.UniqueName,
                        field.Caption,
                        PivotSemanticLayoutFingerprint.CreateCaptionFingerprint(
                            field.Caption),
                        area,
                        normalized,
                        field.CubeFieldType)));
            }

            return result;
        }

        private static IReadOnlyList<ValueFieldHandle> ReadValueFields(
            object pivotTable,
            bool requireContiguousPositions)
        {
            dynamic pivot = pivotTable;
            object collection = ReadRequired(
                () => (object?)pivot.DataFields,
                "Excel did not expose PivotTable DataFields.");
            IReadOnlyList<object> items = ReadCollection(
                collection,
                MaximumDataFields,
                "PivotTable DataFields");
            var result = new List<ValueFieldHandle>(items.Count);
            foreach (object fieldObject in items)
            {
                dynamic field = fieldObject;
                if (ReadRequiredInt(
                        () => (object?)field.Orientation,
                        "DataField orientation") != OrientationData)
                {
                    throw new NotSupportedException(
                        "Excel exposed a non-data field in PivotTable DataFields.");
                }

                object cubeObject = ReadRequired(
                    () => (object?)field.CubeField,
                    "Excel did not expose a DataField CubeField.");
                dynamic cube = cubeObject;
                string uniqueName = ReadBoundedRequiredString(
                    () => (object?)cube.Name,
                    MaximumUniqueNameCharacters,
                    "DataField CubeField unique name");
                int cubeType = ReadRequiredInt(
                    () => (object?)cube.CubeFieldType,
                    "DataField CubeField type");
                if (cubeType != CubeFieldTypeMeasure)
                {
                    throw new NotSupportedException(
                        "A Values field is not backed by an exact cube measure.");
                }

                if (ReadRequiredInt(
                        () => (object?)cube.Orientation,
                        "DataField CubeField orientation") != OrientationData)
                {
                    throw new NotSupportedException(
                        "A Values field hierarchy is not oriented as data.");
                }

                string caption = ReadBoundedRequiredString(
                    () => (object?)field.Caption,
                    MaximumNameCharacters,
                    "DataField caption");
                string format = ReadBoundedOptionalString(
                    () => (object?)field.NumberFormat,
                    MaximumFormatCharacters,
                    "DataField number format");
                int position = ReadRequiredPositiveInt(
                    () => (object?)field.Position,
                    "DataField position");
                result.Add(new ValueFieldHandle(
                    fieldObject,
                    cubeObject,
                    new PivotSemanticValueFieldSnapshot(
                        uniqueName,
                        caption,
                        PivotSemanticLayoutFingerprint.CreateCaptionFingerprint(caption),
                        format,
                        PivotSemanticLayoutFingerprint.CreateNumberFormatFingerprint(format),
                        position,
                        cubeType)));
            }

            if (requireContiguousPositions &&
                result.Select(item => item.Snapshot.Position)
                    .OrderBy(position => position)
                    .Where((position, index) => position != index + 1)
                    .Any())
            {
                throw new NotSupportedException(
                    "Excel exposed duplicate or non-contiguous Values positions.");
            }

            return result.OrderBy(item => item.Snapshot.Position).ToList();
        }

        private static IReadOnlyList<PivotSemanticFilterFieldSnapshot> ReadFilters(
            object pivotTable)
        {
            dynamic pivot = pivotTable;
            object collection = ReadRequired(
                () => (object?)pivot.PageFields,
                "Excel did not expose PivotTable PageFields.");
            IReadOnlyList<object> fields = ReadCollection(
                collection,
                MaximumFilterFields,
                "PivotTable PageFields");
            var result = new List<PivotSemanticFilterFieldSnapshot>(fields.Count);
            foreach (object fieldObject in fields)
            {
                dynamic field = fieldObject;
                if (ReadRequiredInt(
                        () => (object?)field.Orientation,
                        "PageField orientation") != OrientationPage)
                {
                    throw new NotSupportedException(
                        "Excel exposed a non-filter field in PivotTable PageFields.");
                }

                object cubeObject = ReadRequired(
                    () => (object?)field.CubeField,
                    "Excel did not expose a PageField CubeField.");
                dynamic cube = cubeObject;
                string uniqueName = ReadBoundedRequiredString(
                    () => (object?)cube.Name,
                    MaximumUniqueNameCharacters,
                    "PageField CubeField unique name");
                string caption = ReadBoundedRequiredString(
                    () => (object?)field.Caption,
                    MaximumNameCharacters,
                    "PageField caption");
                if (ReadRequiredInt(
                        () => (object?)cube.Orientation,
                        "PageField CubeField orientation") != OrientationPage)
                {
                    throw new NotSupportedException(
                        "A Filters field hierarchy is not oriented as a page field.");
                }
                int position = ReadRequiredPositiveInt(
                    () => (object?)field.Position,
                    "PageField position");
                result.Add(new PivotSemanticFilterFieldSnapshot(
                    uniqueName,
                    caption,
                    position,
                    ReadFilterStateFingerprint(fieldObject)));
            }

            if (result.Select(item => item.Position)
                    .OrderBy(position => position)
                    .Where((position, index) => position != index + 1)
                    .Any() ||
                result.GroupBy(item => item.UniqueName, StringComparer.OrdinalIgnoreCase)
                    .Any(group => group.Count() != 1))
            {
                throw new NotSupportedException(
                    "Excel exposed an ambiguous Filters layout.");
            }

            return result.OrderBy(item => item.Position).ToList();
        }

        private static string ReadFilterStateFingerprint(object fieldObject)
        {
            dynamic field = fieldObject;
            object cubeObject = ReadRequired(
                () => (object?)field.CubeField,
                "Excel did not expose the PageField CubeField for filter state.");
            dynamic cube = cubeObject;
            bool multiple = ReadRequiredBoolean(
                () => (object?)cube.EnableMultiplePageItems,
                "PageField CubeField EnableMultiplePageItems");
            bool allVisible = ReadRequiredBoolean(
                () => (object?)cube.AllItemsVisible,
                "PageField CubeField AllItemsVisible");
            bool includeNewItems = ReadRequiredBoolean(
                () => (object?)cube.IncludeNewItemsInFilter,
                "PageField CubeField IncludeNewItemsInFilter");
            string currentPageName = multiple
                ? "multiple-items"
                : ReadBoundedOptionalString(
                    () => (object?)field.CurrentPageName,
                    MaximumFilterTokenCharacters,
                    "PageField CurrentPageName");
            string currentPageList = multiple
                ? ReadBoundedVariantToken(
                    () => (object?)field.CurrentPageList,
                    "PageField CurrentPageList")
                : "single-item";
            string visibleItems = ReadBoundedVariantToken(
                () => (object?)field.VisibleItemsList,
                "PageField VisibleItemsList");
            string hiddenItems = ReadBoundedVariantToken(
                () => (object?)field.HiddenItemsList,
                "PageField HiddenItemsList");

            object filtersObject = ReadRequired(
                () => (object?)field.PivotFilters,
                "Excel did not expose PageField PivotFilters.");
            IReadOnlyList<object> filters = ReadCollection(
                filtersObject,
                MaximumFilterFields,
                "PageField PivotFilters");
            if (filters.Count != 0)
            {
                throw new NotSupportedException(
                    "Semantic placement fails closed when a Filters-area field has an advanced PivotFilter.");
            }

            var canonical = new StringBuilder("semantic-filter-state-v1");
            Append(canonical, multiple ? "true" : "false");
            Append(canonical, allVisible ? "true" : "false");
            Append(canonical, includeNewItems ? "true" : "false");
            Append(canonical, currentPageName);
            Append(canonical, currentPageList);
            Append(canonical, visibleItems);
            Append(canonical, hiddenItems);

            return Persistence.PivotPlusFingerprint.Create(
                "semantic.filter-state.v1",
                canonical.ToString());
        }

        private static string ReadBoundedVariantToken(
            Func<object?> reader,
            string label)
        {
            if (!PivotLateBound.TryRead(reader, out object? value))
            {
                throw new InvalidOperationException("Excel did not expose " + label + ".");
            }

            var values = new List<string>();
            if (value == null || ReferenceEquals(value, Type.Missing))
            {
                values.Add("missing");
            }
            else if (value is string text)
            {
                values.Add(DemandFilterToken(text, label));
            }
            else if (value is Array array)
            {
                if (array.Length > MaximumFilterItems)
                {
                    throw new NotSupportedException(
                        "Excel " + label + " exceeds its bounded limit.");
                }

                foreach (object? item in array)
                {
                    if (!(item is string itemText))
                    {
                        throw new NotSupportedException(
                            "Excel exposed a non-text " + label + " entry.");
                    }

                    values.Add(DemandFilterToken(itemText, label));
                }
            }
            else
            {
                throw new NotSupportedException(
                    "Excel exposed an unsupported " + label + " value.");
            }

            if (values.GroupBy(item => item, StringComparer.OrdinalIgnoreCase)
                .Any(group => group.Count() != 1))
            {
                throw new NotSupportedException(
                    "Excel exposed duplicate " + label + " entries.");
            }

            var canonical = new StringBuilder();
            foreach (string item in values.OrderBy(
                         item => item,
                         StringComparer.Ordinal))
            {
                Append(canonical, item);
            }

            return canonical.ToString();
        }

        private static string DemandFilterToken(string value, string label)
        {
            if (value.Length > MaximumFilterTokenCharacters || value.Any(char.IsControl))
            {
                throw new NotSupportedException(
                    "Excel exposed an invalid or unbounded " + label + " entry.");
            }

            return value;
        }

        private static CubeFieldHandle ResolveCubeField(
            object pivotTable,
            string uniqueName,
            int? expectedType)
        {
            dynamic pivot = pivotTable;
            object collection = ReadRequired(
                () => (object?)pivot.CubeFields,
                "Excel did not expose PivotTable CubeFields.");
            var matches = new List<CubeFieldHandle>();
            foreach (object cubeObject in ReadCollection(
                         collection,
                         MaximumCubeFields,
                         "PivotTable CubeFields"))
            {
                dynamic cube = cubeObject;
                string liveName = ReadBoundedRequiredString(
                    () => (object?)cube.Name,
                    MaximumUniqueNameCharacters,
                    "CubeField unique name");
                if (!string.Equals(liveName, uniqueName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                matches.Add(new CubeFieldHandle(
                    cubeObject,
                    liveName,
                    ReadRequiredInt(
                        () => (object?)cube.CubeFieldType,
                        "CubeField type")));
            }

            if (matches.Count != 1 ||
                !string.Equals(matches[0].UniqueName, uniqueName, StringComparison.Ordinal) ||
                (expectedType.HasValue && matches[0].CubeFieldType != expectedType.Value))
            {
                throw new InvalidOperationException(
                    "Excel did not expose exactly one exact typed CubeField for semantic placement.");
            }

            return matches[0];
        }

        private static void DemandFiltersUnchanged(
            BoundPivotSemanticLayoutTarget target,
            PivotSemanticLayoutSnapshot before)
        {
            IReadOnlyList<PivotSemanticFilterFieldSnapshot> filters = ReadFilters(
                target.PivotTable);
            string fingerprint = PivotSemanticLayoutCanonical.CreateFilterFingerprint(filters);
            if (!string.Equals(
                    fingerprint,
                    before.FilterFingerprint,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The selected PivotTable Filters changed during semantic placement.");
            }
        }

        private static void DemandSnapshotIdentity(
            BoundPivotSemanticLayoutTarget target,
            PivotSemanticLayoutSnapshot snapshot)
        {
            if (!string.Equals(
                    snapshot.Identity.WorkbookId,
                    target.Identity.WorkbookId,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    snapshot.Identity.WorksheetName,
                    target.Identity.WorksheetName,
                    StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(
                    snapshot.Identity.PivotTableName,
                    target.Identity.PivotTableName,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "The semantic layout snapshot is not the bound PivotTable.");
            }
        }

        private static void DemandSnapshotIntegrity(PivotSemanticLayoutSnapshot snapshot)
        {
            string filter = PivotSemanticLayoutCanonical.CreateFilterFingerprint(
                snapshot.Filters);
            string layout = PivotSemanticLayoutCanonical.CreateLayoutFingerprint(
                snapshot.Rows,
                snapshot.Columns,
                snapshot.Values,
                snapshot.ValuesAxis,
                snapshot.ValuesPosition,
                filter);
            if (!string.Equals(filter, snapshot.FilterFingerprint, StringComparison.Ordinal) ||
                !string.Equals(layout, snapshot.LayoutFingerprint, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The semantic layout snapshot fingerprint is inconsistent.");
            }
        }

        private static void DemandSameLayout(
            PivotSemanticLayoutSnapshot live,
            PivotSemanticLayoutSnapshot expected,
            string message)
        {
            if (!string.Equals(
                    live.LayoutFingerprint,
                    expected.LayoutFingerprint,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(message);
            }
        }

        private static void DemandStillBound(BoundPivotSemanticLayoutTarget target)
        {
            dynamic pivot = target.PivotTable;
            object worksheet = ReadRequired(
                () => (object?)pivot.Parent,
                "Excel did not expose the selected PivotTable worksheet.");
            dynamic nativeWorksheet = worksheet;
            object workbook = ReadRequired(
                () => (object?)nativeWorksheet.Parent,
                "Excel did not expose the selected PivotTable workbook.");
            if (!ComObjectIdentity.AreSame(workbook, target.Workbook))
            {
                throw new InvalidOperationException(
                    "The selected PivotTable moved to another workbook.");
            }

            string workbookId = new StoredWorkbookIdentityResolver().Resolve(target.Workbook);
            if (!string.Equals(workbookId, target.Identity.WorkbookId, StringComparison.Ordinal))
            {
                throw new InvalidOperationException("The bound workbook identity changed.");
            }

            if (!string.Equals(
                    ReadBoundedRequiredString(
                        () => (object?)pivot.Name,
                        MaximumNameCharacters,
                        "selected PivotTable name"),
                    target.Identity.PivotTableName,
                    StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(
                    ReadBoundedRequiredString(
                        () => (object?)nativeWorksheet.Name,
                        MaximumNameCharacters,
                        "selected PivotTable worksheet name"),
                    target.Identity.WorksheetName,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "The selected PivotTable identity changed.");
            }

            dynamic workbookObject = target.Workbook;
            object model = ReadRequired(
                () => (object?)workbookObject.Model,
                "Excel did not expose the bound workbook Data Model.");
            if (!ComObjectIdentity.AreSame(model, target.Model))
            {
                throw new InvalidOperationException(
                    "The bound workbook Data Model object changed.");
            }

            dynamic modelObject = model;
            object connection = ReadRequired(
                () => (object?)modelObject.DataModelConnection,
                "Excel did not expose the bound Data Model connection.");
            if (!ComObjectIdentity.AreSame(connection, target.DataModelConnection) ||
                ReadRequiredInt(
                    () => (object?)((dynamic)connection).Type,
                    "bound Data Model connection type") != DataModelConnectionType)
            {
                throw new InvalidOperationException(
                    "The workbook Data Model connection changed.");
            }

            object cache = ReadPivotCache(target.PivotTable);
            dynamic nativeCache = cache;
            if (!ReadRequiredBoolean(
                    () => (object?)nativeCache.OLAP,
                    "selected PivotCache.OLAP") ||
                !ComObjectIdentity.AreSame(
                    ReadRequired(
                        () => (object?)nativeCache.WorkbookConnection,
                        "Excel did not expose the selected PivotCache connection."),
                    target.DataModelConnection))
            {
                throw new InvalidOperationException(
                    "The selected PivotTable no longer uses the bound Data Model connection.");
            }
        }

        private static void DemandCollectionCapability(
            object pivotTable,
            string member,
            int maximum)
        {
            dynamic pivot = pivotTable;
            object collection;
            switch (member)
            {
                case "CubeFields":
                    collection = ReadRequired(
                        () => (object?)pivot.CubeFields,
                        "Excel did not expose selected PivotTable CubeFields.");
                    break;
                case "RowFields":
                    collection = ReadRequired(
                        () => (object?)pivot.RowFields,
                        "Excel did not expose selected PivotTable RowFields.");
                    break;
                case "ColumnFields":
                    collection = ReadRequired(
                        () => (object?)pivot.ColumnFields,
                        "Excel did not expose selected PivotTable ColumnFields.");
                    break;
                case "DataFields":
                    collection = ReadRequired(
                        () => (object?)pivot.DataFields,
                        "Excel did not expose selected PivotTable DataFields.");
                    break;
                case "PageFields":
                    collection = ReadRequired(
                        () => (object?)pivot.PageFields,
                        "Excel did not expose selected PivotTable PageFields.");
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(member));
            }

            _ = ReadCollection(collection, maximum, "selected PivotTable " + member);
        }

        private static object ReadPivotCache(object pivotTable)
        {
            dynamic pivot = pivotTable;
            if (PivotLateBound.TryRead(
                    () => (object?)pivot.PivotCache(),
                    out object? methodValue) &&
                methodValue != null)
            {
                return methodValue;
            }

            return ReadRequired(
                () => (object?)pivot.PivotCache,
                "Excel did not expose the selected PivotTable cache.");
        }

        private static IReadOnlyList<object> ReadCollection(
            object collectionObject,
            int maximum,
            string label)
        {
            dynamic collection = collectionObject;
            int count = ReadRequiredInt(
                () => (object?)collection.Count,
                label + " count");
            if (count < 0 || count > maximum)
            {
                throw new NotSupportedException(
                    "The Excel " + label + " collection exceeds its bounded limit.");
            }

            var result = new List<object>(count);
            for (var index = 1; index <= count; index++)
            {
                int captured = index;
                result.Add(ReadRequired(
                    () => (object?)collection.Item(captured),
                    "Excel did not expose item " +
                    captured.ToString(CultureInfo.InvariantCulture) +
                    " in " + label + "."));
            }

            return result;
        }

        private static object ReadRequired(Func<object?> reader, string message)
        {
            if (!PivotLateBound.TryRead(reader, out object? value) || value == null)
            {
                throw new InvalidOperationException(message);
            }

            return value;
        }

        private static string ReadBoundedRequiredString(
            Func<object?> reader,
            int maximum,
            string label)
        {
            string value = ReadBoundedOptionalString(reader, maximum, label);
            if (string.IsNullOrWhiteSpace(value))
            {
                throw new InvalidOperationException(
                    "Excel did not expose a non-empty " + label + ".");
            }

            return value;
        }

        private static string ReadBoundedOptionalString(
            Func<object?> reader,
            int maximum,
            string label)
        {
            if (!PivotLateBound.TryRead(reader, out object? value))
            {
                throw new InvalidOperationException("Excel did not expose " + label + ".");
            }

            string result;
            if (value == null || ReferenceEquals(value, Type.Missing))
            {
                result = string.Empty;
            }
            else if (value is string text)
            {
                result = text;
            }
            else
            {
                throw new NotSupportedException(
                    "Excel exposed a non-text " + label + ".");
            }

            if (result.Length > maximum || result.Any(char.IsControl))
            {
                throw new NotSupportedException(
                    "Excel exposed an invalid or unbounded " + label + ".");
            }

            return result;
        }

        private static int ReadRequiredInt(Func<object?> reader, string label)
        {
            object value = ReadRequired(reader, "Excel did not expose " + label + ".");
            try
            {
                return Convert.ToInt32(value, CultureInfo.InvariantCulture);
            }
            catch (Exception exception) when (
                exception is FormatException ||
                exception is InvalidCastException ||
                exception is OverflowException)
            {
                throw new NotSupportedException(
                    "Excel exposed an invalid " + label + ".",
                    exception);
            }
        }

        private static int ReadRequiredPositiveInt(Func<object?> reader, string label)
        {
            int value = ReadRequiredInt(reader, label);
            if (value <= 0)
            {
                throw new NotSupportedException(
                    "Excel exposed an invalid " + label + ".");
            }

            return value;
        }

        private static bool ReadRequiredBoolean(Func<object?> reader, string label)
        {
            object value = ReadRequired(reader, "Excel did not expose " + label + ".");
            try
            {
                return Convert.ToBoolean(value, CultureInfo.InvariantCulture);
            }
            catch (Exception exception) when (
                exception is FormatException || exception is InvalidCastException)
            {
                throw new NotSupportedException(
                    "Excel exposed an invalid " + label + ".",
                    exception);
            }
        }

        private static void SetOrientation(object target, int value, string label)
        {
            try
            {
                ((dynamic)target).Orientation = value;
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    "Excel failed to apply " + label + ".",
                    exception);
            }
        }

        private static void SetPosition(object target, int value, string label)
        {
            try
            {
                ((dynamic)target).Position = value;
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    "Excel failed to apply " + label + ".",
                    exception);
            }
        }

        private static void WriteCaption(object target, string current, string expected)
        {
            if (string.Equals(current, expected, StringComparison.Ordinal)) return;
            try
            {
                ((dynamic)target).Caption = expected;
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    "Excel failed to restore an exact field caption.",
                    exception);
            }
        }

        private static void WritePresentation(
            object target,
            string currentCaption,
            string expectedCaption,
            string currentFormat,
            string expectedFormat)
        {
            WriteCaption(target, currentCaption, expectedCaption);
            if (string.Equals(currentFormat, expectedFormat, StringComparison.Ordinal)) return;
            try
            {
                ((dynamic)target).NumberFormat = expectedFormat;
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    "Excel failed to restore an exact Values number format.",
                    exception);
            }
        }

        private static void Append(StringBuilder target, string value)
        {
            string actual = value ?? string.Empty;
            target.Append('|')
                .Append(actual.Length.ToString(CultureInfo.InvariantCulture))
                .Append(':')
                .Append(actual);
        }

        private sealed class HostState
        {
            public HostState(
                IReadOnlyList<AxisFieldHandle> rows,
                IReadOnlyList<AxisFieldHandle> columns,
                IReadOnlyList<ValueFieldHandle> values,
                object? dataPivotField,
                PivotSemanticLayoutSnapshot snapshot)
            {
                Rows = rows;
                Columns = columns;
                Values = values;
                DataPivotField = dataPivotField;
                Snapshot = snapshot;
            }

            public IReadOnlyList<AxisFieldHandle> Rows { get; }
            public IReadOnlyList<AxisFieldHandle> Columns { get; }
            public IReadOnlyList<ValueFieldHandle> Values { get; }
            public object? DataPivotField { get; }
            public PivotSemanticLayoutSnapshot Snapshot { get; }
        }

        private sealed class RawAxisField
        {
            public RawAxisField(
                object nativeField,
                object cubeField,
                string uniqueName,
                string caption,
                int rawPosition,
                int cubeFieldType)
            {
                NativeField = nativeField;
                CubeField = cubeField;
                UniqueName = uniqueName;
                Caption = caption;
                RawPosition = rawPosition;
                CubeFieldType = cubeFieldType;
            }

            public object NativeField { get; }
            public object CubeField { get; }
            public string UniqueName { get; }
            public string Caption { get; }
            public int RawPosition { get; }
            public int CubeFieldType { get; }
        }

        private sealed class AxisFieldHandle
        {
            public AxisFieldHandle(
                object nativeField,
                object cubeField,
                int rawPosition,
                PivotSemanticAxisFieldSnapshot snapshot)
            {
                NativeField = nativeField;
                CubeField = cubeField;
                RawPosition = rawPosition;
                Snapshot = snapshot;
            }

            public object NativeField { get; }
            public object CubeField { get; }
            public int RawPosition { get; }
            public PivotSemanticAxisFieldSnapshot Snapshot { get; }
        }

        private sealed class ValueFieldHandle
        {
            public ValueFieldHandle(
                object nativeField,
                object cubeField,
                PivotSemanticValueFieldSnapshot snapshot)
            {
                NativeField = nativeField;
                CubeField = cubeField;
                Snapshot = snapshot;
            }

            public object NativeField { get; }
            public object CubeField { get; }
            public PivotSemanticValueFieldSnapshot Snapshot { get; }
        }

        private sealed class CubeFieldHandle
        {
            public CubeFieldHandle(object native, string uniqueName, int cubeFieldType)
            {
                Native = native;
                UniqueName = uniqueName;
                CubeFieldType = cubeFieldType;
            }

            public object Native { get; }
            public string UniqueName { get; }
            public int CubeFieldType { get; }
        }

        private sealed class ReferenceEqualityComparer : IEqualityComparer<object>
        {
            public static readonly ReferenceEqualityComparer Instance =
                new ReferenceEqualityComparer();

            public new bool Equals(object? x, object? y)
            {
                return ReferenceEquals(x, y);
            }

            public int GetHashCode(object obj)
            {
                return System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
            }
        }
    }
}
