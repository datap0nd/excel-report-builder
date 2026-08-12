using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using ExcelReportBuilder.Core.PivotPlus;
using ExcelReportBuilder.Core.PivotPlus.Calculations;
using ExcelReportBuilder.Excel.PivotPlus.Persistence;

namespace ExcelReportBuilder.Excel.PivotPlus.Measures
{
    /// <summary>
    /// Strict late-bound boundary for workbook Data Model measures. All DAX is
    /// supplied by the validated compiler layer; this class only translates a
    /// compiled operation to the documented Excel object model.
    /// </summary>
    internal sealed class LateBoundPivotModelMeasureGateway : IPivotModelMeasureGateway
    {
        private const int OrientationHidden = 0;
        private const int OrientationRow = 1;
        private const int OrientationColumn = 2;
        private const int OrientationData = 4;
        private const int CubeFieldTypeMeasure = 2;
        private const int CubeFieldSubTypeImplicitMeasure = 11;
        private const int DataModelConnectionType = 7;

        private const int MaximumWorksheets = 1024;
        private const int MaximumPivotTables = 4096;
        private const int MaximumModelTables = 512;
        private const int MaximumModelMeasures = 512;
        private const int MaximumDataFields = 512;
        private const int MaximumCubeFields = 4096;
        private const int MaximumNameCharacters = 255;
        private const int MaximumDescriptionCharacters = 4096;
        private const int MaximumFormulaCharacters = 1024 * 1024;
        private const int MaximumFormatCharacters = 255;
        private const int MaximumFormatDecimalPlaces = 30;

        private readonly Func<object, string> typeNameResolver;

        public LateBoundPivotModelMeasureGateway()
            : this(ResolveRuntimeTypeName)
        {
        }

        internal LateBoundPivotModelMeasureGateway(Func<object, string> typeNameResolver)
        {
            this.typeNameResolver = typeNameResolver ??
                throw new ArgumentNullException(nameof(typeNameResolver));
        }

        public BoundModelMeasureTarget Bind(
            object workbook,
            object pivotTable,
            PivotTableContext context)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            if (context == null) throw new ArgumentNullException(nameof(context));

            PivotLayoutDefinition definition = context.Definition;
            if (!context.IsConnected ||
                !context.SourceFieldsComplete ||
                definition.Source.Kind != PivotSourceKind.DataModel ||
                (definition.Source.Capabilities & PivotCapability.DataModel) == 0 ||
                (definition.Source.Capabilities & PivotCapability.ModelMeasures) == 0)
            {
                throw new NotSupportedException(
                    "Model measures require the selected native workbook Data Model PivotTable.");
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
            object dataModelConnection = ReadRequired(
                () => (object?)nativeModel.DataModelConnection,
                "Excel did not expose Workbook.Model.DataModelConnection.");
            dynamic nativeDataModelConnection = dataModelConnection;
            if (ReadRequiredInt(
                    () => (object?)nativeDataModelConnection.Type,
                    "Data Model connection type") != DataModelConnectionType)
            {
                throw new InvalidOperationException(
                    "Workbook.Model.DataModelConnection is not the special Data Model connection.");
            }

            object cacheConnection = ReadRequired(
                () => (object?)nativeCache.WorkbookConnection,
                "Excel did not expose the selected PivotCache workbook connection.");
            if (!ComObjectIdentity.AreSame(cacheConnection, dataModelConnection))
            {
                throw new NotSupportedException(
                    "The selected PivotTable does not use this workbook's exact Data Model connection.");
            }

            // Read both collections now so a successful bind is also a strict
            // capability check for the APIs used by subsequent mutations.
            ReadCollection(
                ReadRequired(
                    () => (object?)nativeModel.ModelTables,
                    "Excel did not expose Data Model tables."),
                MaximumModelTables,
                "Data Model tables");
            ReadCollection(
                ReadRequired(
                    () => (object?)nativeModel.ModelMeasures,
                    "Excel did not expose Data Model measures."),
                MaximumModelMeasures,
                "Data Model measures");

            return new BoundModelMeasureTarget(
                workbook,
                pivotTable,
                model,
                dataModelConnection,
                expected);
        }

        public ModelMeasureWorkbookSnapshot Capture(BoundModelMeasureTarget target)
        {
            if (target == null) throw new ArgumentNullException(nameof(target));
            DemandStillBound(target);

            IReadOnlyList<ModelTableHandle> tables = ReadModelTables(target);
            IReadOnlyList<MeasureHandle> measures = ReadModelMeasures(target, tables);
            var measureNamesByUniqueName = measures.ToDictionary(
                item => MeasureCubeUniqueName(item.Snapshot.Name),
                item => item.Snapshot.Name,
                StringComparer.OrdinalIgnoreCase);

            dynamic workbook = target.Workbook;
            object worksheets = ReadRequired(
                () => (object?)workbook.Worksheets,
                "Excel did not expose workbook worksheets.");
            var usages = new List<ModelPivotUsageSnapshot>();
            foreach (object worksheetObject in ReadCollection(
                         worksheets,
                         MaximumWorksheets,
                         "workbook worksheets"))
            {
                dynamic worksheet = worksheetObject;
                string worksheetName = ReadBoundedRequiredString(
                    () => (object?)worksheet.Name,
                    MaximumNameCharacters,
                    "worksheet name");
                object pivotTables = ReadRequiredPivotTables(worksheet);
                foreach (object pivotTable in ReadCollection(
                             pivotTables,
                             MaximumPivotTables,
                             "worksheet PivotTables"))
                {
                    if (!IsWorkbookModelPivot(pivotTable, target.DataModelConnection))
                    {
                        continue;
                    }

                    bool isSelected = ComObjectIdentity.AreSame(
                        pivotTable,
                        target.PivotTable);
                    usages.Add(ReadPivotUsage(
                        pivotTable,
                        worksheetName,
                        isSelected,
                        measureNamesByUniqueName));
                }
            }

            if (usages.Count(item => item.IsSelectedTarget) != 1)
            {
                throw new InvalidOperationException(
                    "Excel did not expose the selected PivotTable exactly once in this workbook.");
            }

            ModelPivotUsageSnapshot selected = usages.Single(item => item.IsSelectedTarget);
            DemandUsageIdentity(selected, target.Identity);
            return new ModelMeasureWorkbookSnapshot(
                measures.Select(item => item.Snapshot),
                usages,
                PivotModelMeasureCanonical.CreatePivotFingerprint(selected));
        }

        public LiveModelMeasureSnapshot CreateMeasure(
            BoundModelMeasureTarget target,
            DesiredModelMeasure definition)
        {
            if (target == null) throw new ArgumentNullException(nameof(target));
            if (definition == null) throw new ArgumentNullException(nameof(definition));
            DemandStillBound(target);

            IReadOnlyList<ModelTableHandle> tables = ReadModelTables(target);
            ModelTableHandle table = ResolveTable(
                tables,
                definition.HomeTableName,
                expectedLineageFingerprint: null);
            if (FindMeasure(target, tables, definition.Name) != null)
            {
                throw new InvalidOperationException(
                    "A Data Model measure already uses the requested generated name.");
            }

            ModelMeasureFormatSnapshot desiredFormat = ToFormatSnapshot(definition.Format);
            object formatTemplate = ReadModelFormatTemplate(target.Model, desiredFormat.Kind);
            object? created = null;
            Exception? addFailure = null;
            try
            {
                dynamic model = target.Model;
                object measures = ReadRequired(
                    () => (object?)model.ModelMeasures,
                    "Excel did not expose Data Model measures for creation.");
                dynamic nativeMeasures = measures;
                created = nativeMeasures.Add(
                    definition.Name,
                    table.Native,
                    definition.Formula,
                    formatTemplate,
                    definition.DescriptionMarker);
            }
            catch (Exception exception)
            {
                addFailure = exception;
            }

            MeasureHandle? live = FindMeasure(target, tables, definition.Name);
            if (live == null)
            {
                throw new InvalidOperationException(
                    "Excel did not create the requested Data Model measure.",
                    addFailure);
            }

            if (created != null && !ComObjectIdentity.AreSame(created, live.Native))
            {
                throw new InvalidOperationException(
                    "Excel returned a different measure from the one committed to the model.",
                    addFailure);
            }

            if (!MatchesDesiredCore(live.Snapshot, definition, table, desiredFormat.Kind))
            {
                throw new InvalidOperationException(
                    "Excel may have committed a different measure definition; ownership is ambiguous.",
                    addFailure);
            }

            try
            {
                ConfigureMeasureFormat(
                    live.Native,
                    formatTemplate,
                    desiredFormat);
                MeasureHandle verified = ReadMeasureHandle(live.Native, tables);
                DemandDesired(verified.Snapshot, definition, table, desiredFormat);
                return verified.Snapshot;
            }
            catch (Exception exception)
            {
                MeasureHandle? reconciled = FindMeasure(target, tables, definition.Name);
                if (reconciled != null &&
                    MatchesDesired(reconciled.Snapshot, definition, table, desiredFormat))
                {
                    return reconciled.Snapshot;
                }

                throw new InvalidOperationException(
                    "Excel created the measure but its exact final state is ambiguous.",
                    addFailure == null ? exception : new AggregateException(addFailure, exception));
            }
        }

        public LiveModelMeasureSnapshot UpdateMeasure(
            BoundModelMeasureTarget target,
            LiveModelMeasureSnapshot before,
            DesiredModelMeasure definition)
        {
            if (target == null) throw new ArgumentNullException(nameof(target));
            if (before == null) throw new ArgumentNullException(nameof(before));
            if (definition == null) throw new ArgumentNullException(nameof(definition));
            DemandStillBound(target);
            if (!string.Equals(before.Name, definition.Name, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "Model measures are updated in place and cannot be renamed.");
            }

            IReadOnlyList<ModelTableHandle> tables = ReadModelTables(target);
            MeasureHandle current = FindMeasure(target, tables, before.Name) ??
                throw new InvalidOperationException(
                    "The measure selected for update no longer exists.");
            DemandSameSnapshot(current.Snapshot, before, "measure selected for update");
            ModelTableHandle desiredTable = ResolveTable(
                tables,
                definition.HomeTableName,
                expectedLineageFingerprint: null);
            ModelMeasureFormatSnapshot desiredFormat = ToFormatSnapshot(definition.Format);
            object formatTemplate = ReadModelFormatTemplate(target.Model, desiredFormat.Kind);

            Exception? mutationFailure = null;
            try
            {
                WriteMeasureDefinition(
                    current.Native,
                    desiredTable.Native,
                    definition.Formula,
                    definition.DescriptionMarker,
                    formatTemplate,
                    desiredFormat);
            }
            catch (Exception exception)
            {
                mutationFailure = exception;
            }

            MeasureHandle? live = FindMeasure(target, tables, before.Name);
            if (live != null &&
                MatchesDesired(live.Snapshot, definition, desiredTable, desiredFormat))
            {
                return live.Snapshot;
            }

            if (live != null && MatchesDesiredCore(
                    live.Snapshot,
                    definition,
                    desiredTable,
                    desiredFormat.Kind))
            {
                try
                {
                    ConfigureMeasureFormat(live.Native, formatTemplate, desiredFormat);
                    MeasureHandle configured = ReadMeasureHandle(live.Native, tables);
                    if (MatchesDesired(
                            configured.Snapshot,
                            definition,
                            desiredTable,
                            desiredFormat))
                    {
                        return configured.Snapshot;
                    }
                }
                catch (Exception exception)
                {
                    mutationFailure = mutationFailure == null
                        ? exception
                        : new AggregateException(mutationFailure, exception);
                }
            }

            if (live != null && SameSnapshot(live.Snapshot, before))
            {
                throw new InvalidOperationException(
                    "Excel did not commit the measure update.",
                    mutationFailure);
            }

            Exception cause = mutationFailure ?? new InvalidOperationException(
                "Excel exposed a partial measure update.");
            try
            {
                RestoreExistingMeasure(target, tables, before);
            }
            catch (Exception rollbackFailure)
            {
                throw new InvalidOperationException(
                    "Excel partially updated the measure and exact restoration is ambiguous.",
                    new AggregateException(cause, rollbackFailure));
            }

            throw new InvalidOperationException(
                "Excel did not commit the measure update; the prior exact definition was restored.",
                cause);
        }

        public LiveModelMeasureSnapshot RestoreMeasure(
            BoundModelMeasureTarget target,
            LiveModelMeasureSnapshot before)
        {
            if (target == null) throw new ArgumentNullException(nameof(target));
            if (before == null) throw new ArgumentNullException(nameof(before));
            DemandStillBound(target);
            IReadOnlyList<ModelTableHandle> tables = ReadModelTables(target);
            MeasureHandle? current = FindMeasure(target, tables, before.Name);
            if (current != null)
            {
                return RestoreExistingMeasure(target, tables, before);
            }

            ModelTableHandle table = ResolveTable(
                tables,
                before.AssociatedTableName,
                before.AssociatedTableLineageFingerprint);
            object template = ReadModelFormatTemplate(target.Model, before.Format.Kind);
            object? restored = null;
            Exception? addFailure = null;
            try
            {
                dynamic model = target.Model;
                dynamic measures = ReadRequired(
                    () => (object?)model.ModelMeasures,
                    "Excel did not expose Data Model measures for restoration.");
                restored = measures.Add(
                    before.Name,
                    table.Native,
                    before.Formula,
                    template,
                    before.Description);
            }
            catch (Exception exception)
            {
                addFailure = exception;
            }

            MeasureHandle? live = FindMeasure(target, tables, before.Name);
            if (live == null)
            {
                throw new InvalidOperationException(
                    "Excel did not restore the deleted measure.",
                    addFailure);
            }

            if (restored != null && !ComObjectIdentity.AreSame(restored, live.Native))
            {
                throw new InvalidOperationException(
                    "Excel returned a different restored measure from the one committed to the model.",
                    addFailure);
            }

            if (!MatchesRestoreCore(live.Snapshot, before))
            {
                throw new InvalidOperationException(
                    "Excel may have restored a different measure definition.",
                    addFailure);
            }

            try
            {
                ConfigureMeasureFormat(live.Native, template, before.Format);
                MeasureHandle verified = ReadMeasureHandle(live.Native, tables);
                DemandSameSnapshot(verified.Snapshot, before, "restored measure");
                return verified.Snapshot;
            }
            catch (Exception exception)
            {
                MeasureHandle? reconciled = FindMeasure(target, tables, before.Name);
                if (reconciled != null && SameSnapshot(reconciled.Snapshot, before))
                {
                    return reconciled.Snapshot;
                }

                throw new InvalidOperationException(
                    "Excel restored the measure but its exact state is ambiguous.",
                    addFailure == null ? exception : new AggregateException(addFailure, exception));
            }
        }

        public void DeleteMeasure(
            BoundModelMeasureTarget target,
            LiveModelMeasureSnapshot expected)
        {
            if (target == null) throw new ArgumentNullException(nameof(target));
            if (expected == null) throw new ArgumentNullException(nameof(expected));
            DemandStillBound(target);
            IReadOnlyList<ModelTableHandle> tables = ReadModelTables(target);
            MeasureHandle current = FindMeasure(target, tables, expected.Name) ??
                throw new InvalidOperationException(
                    "The measure selected for deletion no longer exists.");
            DemandSameSnapshot(current.Snapshot, expected, "measure selected for deletion");

            Exception? deleteFailure = null;
            try
            {
                dynamic native = current.Native;
                native.Delete();
            }
            catch (Exception exception)
            {
                deleteFailure = exception;
            }

            MeasureHandle? survivor = FindMeasure(target, tables, expected.Name);
            if (survivor == null)
            {
                return;
            }

            if (SameSnapshot(survivor.Snapshot, expected))
            {
                throw new InvalidOperationException(
                    "Excel did not commit the measure deletion.",
                    deleteFailure);
            }

            throw new InvalidOperationException(
                "Excel exposed a different measure after deletion; ownership is ambiguous.",
                deleteFailure);
        }

        public void ApplyPlacement(
            BoundModelMeasureTarget target,
            PivotMeasurePlacementPlan placement,
            IReadOnlyDictionary<string, DesiredModelMeasure> definitionsById,
            ModelMeasureWorkbookSnapshot before)
        {
            if (target == null) throw new ArgumentNullException(nameof(target));
            if (placement == null) throw new ArgumentNullException(nameof(placement));
            if (definitionsById == null) throw new ArgumentNullException(nameof(definitionsById));
            if (before == null) throw new ArgumentNullException(nameof(before));
            DemandStillBound(target);
            DemandUsageIdentity(before.SelectedPivot, target.Identity);
            ValidatePlacement(placement, definitionsById);

            IReadOnlyList<ModelTableHandle> tables = ReadModelTables(target);
            IReadOnlyList<MeasureHandle> measures = ReadModelMeasures(target, tables);
            var modelNames = measures.ToDictionary(
                item => MeasureCubeUniqueName(item.Snapshot.Name),
                item => item.Snapshot.Name,
                StringComparer.OrdinalIgnoreCase);
            ModelPivotUsageSnapshot current = ReadSelectedUsage(target, modelNames);
            if (!string.Equals(
                    PivotModelMeasureCanonical.CreatePivotFingerprint(current),
                    PivotModelMeasureCanonical.CreatePivotFingerprint(before.SelectedPivot),
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The selected PivotTable Values layout changed after preview.");
            }

            IReadOnlyList<DataFieldHandle> originalFields = ReadDataFieldHandles(
                target.PivotTable,
                modelNames);
            var originalExisting = originalFields.ToDictionary(
                item => ExistingFieldKey(
                    item.Snapshot.UniqueName,
                    item.Snapshot.CaptionFingerprint,
                    PivotMeasurePlacementFingerprint.CreateNumberFormatFingerprint(
                        item.Snapshot.NumberFormat),
                    item.Snapshot.Position),
                item => item,
                StringComparer.OrdinalIgnoreCase);
            var selectedDefinitionIds = new HashSet<string>(
                placement.Values
                    .Where(item => item.IsGeneratedMeasure)
                    .Select(item => item.DefinitionId!),
                StringComparer.Ordinal);

            // Only definitions in this validated request are eligible for
            // hiding. Other data fields are left for the service's owned
            // deletion phase and are never guessed from their captions.
            foreach (DesiredModelMeasure definition in definitionsById.Values)
            {
                CubeFieldHandle? cube = FindCubeField(
                    target.PivotTable,
                    MeasureCubeUniqueName(definition.Name));
                if (cube != null && !selectedDefinitionIds.Contains(definition.DefinitionId))
                {
                    dynamic nativeCube = cube.Native;
                    nativeCube.Orientation = OrientationHidden;
                }
            }

            foreach (PivotMeasureValuePlacement value in placement.Values
                         .Where(item => item.IsGeneratedMeasure)
                         .OrderBy(item => item.Position))
            {
                DesiredModelMeasure definition = definitionsById[value.DefinitionId!];
                MeasureHandle liveMeasure = measures.SingleOrDefault(item =>
                    string.Equals(
                        item.Snapshot.Name,
                        definition.Name,
                        StringComparison.OrdinalIgnoreCase)) ??
                    throw new InvalidOperationException(
                        "The generated model measure is not live for placement.");
                if (!string.Equals(
                        liveMeasure.Snapshot.Name,
                        definition.Name,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "The generated measure name casing changed in Excel.");
                }

                CubeFieldHandle cube = ResolveAuthoredMeasureCubeField(
                    target.PivotTable,
                    definition.Name);
                dynamic nativeCube = cube.Native;
                nativeCube.Orientation = OrientationData;
            }

            IReadOnlyList<DataFieldHandle> liveFields = ReadDataFieldHandles(
                target.PivotTable,
                modelNames);
            var byUniqueName = liveFields
                .GroupBy(item => item.Snapshot.UniqueName, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(
                    group => group.Key,
                    group => group.ToList(),
                    StringComparer.OrdinalIgnoreCase);
            var ordered = new List<DataFieldHandle>(placement.Values.Count);
            foreach (PivotMeasureValuePlacement value in placement.Values.OrderBy(item => item.Position))
            {
                DataFieldHandle field;
                if (value.IsGeneratedMeasure)
                {
                    DesiredModelMeasure definition = definitionsById[value.DefinitionId!];
                    string uniqueName = MeasureCubeUniqueName(definition.Name);
                    field = DemandSingleDataField(byUniqueName, uniqueName);
                }
                else
                {
                    PivotExistingDataFieldIdentity identity = value.ExistingDataField!;
                    string key = ExistingFieldKey(
                        identity.UniqueName,
                        identity.CurrentCaptionFingerprint,
                        identity.CurrentNumberFormatFingerprint,
                        identity.CurrentPosition);
                    if (!originalExisting.TryGetValue(key, out DataFieldHandle? original))
                    {
                        throw new InvalidOperationException(
                            "An existing Values field no longer matches its preview identity.");
                    }

                    field = liveFields.SingleOrDefault(item =>
                        ComObjectIdentity.AreSame(item.Native, original.Native)) ??
                        throw new InvalidOperationException(
                            "Excel did not preserve an existing Values field during placement.");
                }

                ordered.Add(field);
            }

            // The service admits a complete final plan only after proving that
            // every unowned current value is represented. It is therefore safe
            // here to remove omitted authored values before their owned model
            // measures are deleted later in the transaction. We act on exact
            // CubeField objects and never infer ownership from a caption.
            foreach (DataFieldHandle live in liveFields)
            {
                if (!ordered.Any(item => ComObjectIdentity.AreSame(item.Native, live.Native)))
                {
                    dynamic nativeCube = live.CubeField;
                    nativeCube.Orientation = OrientationHidden;
                }
            }

            liveFields = ReadDataFieldHandles(target.PivotTable, modelNames);
            byUniqueName = liveFields
                .GroupBy(item => item.Snapshot.UniqueName, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(
                    group => group.Key,
                    group => group.ToList(),
                    StringComparer.OrdinalIgnoreCase);
            ordered.Clear();
            foreach (PivotMeasureValuePlacement value in placement.Values.OrderBy(item => item.Position))
            {
                if (value.IsGeneratedMeasure)
                {
                    DesiredModelMeasure definition = definitionsById[value.DefinitionId!];
                    ordered.Add(DemandSingleDataField(
                        byUniqueName,
                        MeasureCubeUniqueName(definition.Name)));
                }
                else
                {
                    PivotExistingDataFieldIdentity identity = value.ExistingDataField!;
                    string key = ExistingFieldKey(
                        identity.UniqueName,
                        identity.CurrentCaptionFingerprint,
                        identity.CurrentNumberFormatFingerprint,
                        identity.CurrentPosition);
                    DataFieldHandle original = originalExisting[key];
                    ordered.Add(liveFields.Single(item =>
                        ComObjectIdentity.AreSame(item.Native, original.Native)));
                }
            }

            for (var index = 0; index < ordered.Count; index++)
            {
                dynamic nativeField = ordered[index].Native;
                nativeField.Position = index + 1;
            }

            // Reapply exact user-facing settings for every field that existed
            // at preview, including an already-placed owned model measure.
            foreach (DataFieldHandle original in originalFields)
            {
                DataFieldHandle? live = ReadDataFieldHandles(target.PivotTable, modelNames)
                    .SingleOrDefault(item =>
                        ComObjectIdentity.AreSame(item.Native, original.Native));
                if (live == null) continue;
                dynamic nativeField = live.Native;
                if (!string.Equals(
                        live.Snapshot.Caption,
                        original.Snapshot.Caption,
                        StringComparison.Ordinal))
                {
                    nativeField.Caption = original.Snapshot.Caption;
                }

                if (!string.Equals(
                        live.Snapshot.NumberFormat,
                        original.Snapshot.NumberFormat,
                        StringComparison.Ordinal))
                {
                    nativeField.NumberFormat = original.Snapshot.NumberFormat;
                }
            }

            ApplyValuesAxis(
                target.PivotTable,
                placement.ValuesAxis,
                placement.ValuesPosition,
                placement.Values.Count);
        }

        public void RestorePlacement(
            BoundModelMeasureTarget target,
            ModelPivotUsageSnapshot before)
        {
            if (target == null) throw new ArgumentNullException(nameof(target));
            if (before == null) throw new ArgumentNullException(nameof(before));
            DemandStillBound(target);
            DemandUsageIdentity(before, target.Identity);

            IReadOnlyList<ModelTableHandle> tables = ReadModelTables(target);
            IReadOnlyList<MeasureHandle> measures = ReadModelMeasures(target, tables);
            var modelNames = measures.ToDictionary(
                item => MeasureCubeUniqueName(item.Snapshot.Name),
                item => item.Snapshot.Name,
                StringComparer.OrdinalIgnoreCase);
            var expectedUniqueNames = new HashSet<string>(
                before.DataFields.Select(item => item.UniqueName),
                StringComparer.OrdinalIgnoreCase);

            foreach (DataFieldHandle current in ReadDataFieldHandles(target.PivotTable, modelNames))
            {
                if (!expectedUniqueNames.Contains(current.Snapshot.UniqueName))
                {
                    dynamic nativeCube = current.CubeField;
                    nativeCube.Orientation = OrientationHidden;
                }
            }

            foreach (ModelDataFieldSnapshot field in before.DataFields.OrderBy(item => item.Position))
            {
                CubeFieldHandle cube = ResolveMeasureCubeField(
                    target.PivotTable,
                    field.UniqueName,
                    allowImplicit: true);
                dynamic nativeCube = cube.Native;
                nativeCube.Orientation = OrientationData;
            }

            IReadOnlyList<DataFieldHandle> restored = ReadDataFieldHandles(
                target.PivotTable,
                modelNames);
            var available = restored.ToList();
            foreach (ModelDataFieldSnapshot expected in before.DataFields.OrderBy(item => item.Position))
            {
                List<DataFieldHandle> sameSource = available.Where(item => string.Equals(
                        item.Snapshot.UniqueName,
                        expected.UniqueName,
                        StringComparison.OrdinalIgnoreCase))
                    .ToList();
                if (sameSource.Count == 0)
                {
                    throw new InvalidOperationException(
                        "Excel did not restore an exact Values field occurrence.");
                }

                DataFieldHandle field = sameSource
                    .Where(item =>
                        string.Equals(
                            item.Snapshot.CaptionFingerprint,
                            expected.CaptionFingerprint,
                            StringComparison.Ordinal) &&
                        string.Equals(
                            item.Snapshot.NumberFormat,
                            expected.NumberFormat,
                            StringComparison.Ordinal))
                    .OrderBy(item => Math.Abs(item.Snapshot.Position - expected.Position))
                    .FirstOrDefault() ?? sameSource
                    .OrderBy(item => Math.Abs(item.Snapshot.Position - expected.Position))
                    .First();
                available.Remove(field);
                dynamic nativeField = field.Native;
                nativeField.Position = expected.Position;
                if (!string.Equals(
                        field.Snapshot.Caption,
                        expected.Caption,
                        StringComparison.Ordinal))
                {
                    nativeField.Caption = expected.Caption;
                }

                if (!string.Equals(
                        field.Snapshot.NumberFormat,
                        expected.NumberFormat,
                        StringComparison.Ordinal))
                {
                    nativeField.NumberFormat = expected.NumberFormat;
                }
            }

            ApplyValuesAxis(
                target.PivotTable,
                before.ValuesAxis,
                before.ValuesPosition,
                before.DataFields.Count);

            ModelPivotUsageSnapshot verified = ReadSelectedUsage(target, modelNames);
            if (!string.Equals(
                    PivotModelMeasureCanonical.CreatePivotFingerprint(verified),
                    PivotModelMeasureCanonical.CreatePivotFingerprint(before),
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "Excel did not restore the exact prior Values layout.");
            }
        }

        public void Refresh(BoundModelMeasureTarget target)
        {
            if (target == null) throw new ArgumentNullException(nameof(target));
            DemandStillBound(target);
            dynamic pivot = target.PivotTable;
            object result;
            try
            {
                result = pivot.RefreshTable();
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    "Excel failed to refresh the selected PivotTable.",
                    exception);
            }

            bool refreshed;
            try
            {
                refreshed = Convert.ToBoolean(result, CultureInfo.InvariantCulture);
            }
            catch (Exception exception) when (
                exception is FormatException || exception is InvalidCastException)
            {
                throw new InvalidOperationException(
                    "Excel returned an invalid RefreshTable result.",
                    exception);
            }

            if (!refreshed)
            {
                throw new InvalidOperationException(
                    "Excel reported that the selected PivotTable refresh failed.");
            }
        }

        private LiveModelMeasureSnapshot RestoreExistingMeasure(
            BoundModelMeasureTarget target,
            IReadOnlyList<ModelTableHandle> tables,
            LiveModelMeasureSnapshot before)
        {
            MeasureHandle current = FindMeasure(target, tables, before.Name) ??
                throw new InvalidOperationException(
                    "The measure selected for restoration no longer exists.");
            ModelTableHandle table = ResolveTable(
                tables,
                before.AssociatedTableName,
                before.AssociatedTableLineageFingerprint);
            object template = ReadModelFormatTemplate(target.Model, before.Format.Kind);
            Exception? restoreFailure = null;
            try
            {
                WriteMeasureDefinition(
                    current.Native,
                    table.Native,
                    before.Formula,
                    before.Description,
                    template,
                    before.Format);
            }
            catch (Exception exception)
            {
                restoreFailure = exception;
            }

            MeasureHandle? restored = FindMeasure(target, tables, before.Name);
            if (restored != null && SameSnapshot(restored.Snapshot, before))
            {
                return restored.Snapshot;
            }

            throw new InvalidOperationException(
                "Excel did not restore the exact prior measure definition.",
                restoreFailure);
        }

        private void WriteMeasureDefinition(
            object measureObject,
            object table,
            string formula,
            string description,
            object formatTemplate,
            ModelMeasureFormatSnapshot format)
        {
            dynamic measure = measureObject;
            // Name is deliberately absent: renaming can invalidate dependent
            // DAX and active CubeField identities.
            measure.AssociatedTable = table;
            measure.Formula = formula;
            measure.FormatInformation = formatTemplate;
            measure.Description = description;
            ConfigureMeasureFormat(measureObject, formatTemplate, format);
        }

        private void ConfigureMeasureFormat(
            object measureObject,
            object assignedTemplate,
            ModelMeasureFormatSnapshot desired)
        {
            dynamic measure = measureObject;
            object liveFormat = ReadRequired(
                () => (object?)measure.FormatInformation,
                "Excel did not expose the measure format after assignment.");
            ModelMeasureFormatSnapshot current = ReadFormat(liveFormat);
            if (current.Kind != desired.Kind)
            {
                throw new InvalidOperationException(
                    "Excel assigned a different model-measure format type.");
            }

            if (FormatNeedsPropertyWrites(desired) &&
                ComObjectIdentity.AreSame(liveFormat, assignedTemplate))
            {
                throw new NotSupportedException(
                    "Excel aliased the measure format to the workbook format template; safe formatting is unavailable.");
            }

            dynamic format = liveFormat;
            switch (desired.Kind)
            {
                case ExcelModelMeasureFormatKind.General:
                case ExcelModelMeasureFormatKind.Boolean:
                    break;
                case ExcelModelMeasureFormatKind.WholeNumber:
                    format.UseThousandSeparator = desired.UseThousandsSeparator!.Value;
                    break;
                case ExcelModelMeasureFormatKind.DecimalNumber:
                case ExcelModelMeasureFormatKind.PercentageNumber:
                    format.DecimalPlaces = desired.DecimalPlaces!.Value;
                    format.UseThousandSeparator = desired.UseThousandsSeparator!.Value;
                    break;
                case ExcelModelMeasureFormatKind.ScientificNumber:
                    format.DecimalPlaces = desired.DecimalPlaces!.Value;
                    break;
                case ExcelModelMeasureFormatKind.Currency:
                    format.DecimalPlaces = desired.DecimalPlaces!.Value;
                    format.Symbol = desired.CurrencySymbol!;
                    break;
                case ExcelModelMeasureFormatKind.Date:
                    format.FormatString = desired.DateFormatString!;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(desired));
            }

            ModelMeasureFormatSnapshot verified = ReadFormat(liveFormat);
            if (!SameFormat(verified, desired))
            {
                throw new InvalidOperationException(
                    "Excel did not preserve the exact requested model-measure format.");
            }
        }

        private ModelMeasureFormatSnapshot ReadFormat(object formatObject)
        {
            string rawTypeName;
            try
            {
                rawTypeName = typeNameResolver(formatObject);
            }
            catch (Exception exception)
            {
                throw new NotSupportedException(
                    "Excel did not expose the model-measure format COM type.",
                    exception);
            }

            ExcelModelMeasureFormatKind kind = ParseFormatKind(rawTypeName);
            dynamic format = formatObject;
            switch (kind)
            {
                case ExcelModelMeasureFormatKind.General:
                case ExcelModelMeasureFormatKind.Boolean:
                    return new ModelMeasureFormatSnapshot(kind);
                case ExcelModelMeasureFormatKind.WholeNumber:
                    return new ModelMeasureFormatSnapshot(
                        kind,
                        useThousandsSeparator: ReadRequiredBoolean(
                            () => (object?)format.UseThousandSeparator,
                            "whole-number thousands separator"));
                case ExcelModelMeasureFormatKind.DecimalNumber:
                case ExcelModelMeasureFormatKind.PercentageNumber:
                    return new ModelMeasureFormatSnapshot(
                        kind,
                        decimalPlaces: ReadFormatDecimalPlaces(
                            () => (object?)format.DecimalPlaces,
                            kind + " decimal places"),
                        useThousandsSeparator: ReadRequiredBoolean(
                            () => (object?)format.UseThousandSeparator,
                            kind + " thousands separator"));
                case ExcelModelMeasureFormatKind.ScientificNumber:
                    return new ModelMeasureFormatSnapshot(
                        kind,
                        decimalPlaces: ReadFormatDecimalPlaces(
                            () => (object?)format.DecimalPlaces,
                            "scientific decimal places"));
                case ExcelModelMeasureFormatKind.Currency:
                    return new ModelMeasureFormatSnapshot(
                        kind,
                        decimalPlaces: ReadFormatDecimalPlaces(
                            () => (object?)format.DecimalPlaces,
                            "currency decimal places"),
                        currencySymbol: ReadBoundedOptionalString(
                            () => (object?)format.Symbol,
                            MaximumFormatCharacters,
                            "currency symbol"));
                case ExcelModelMeasureFormatKind.Date:
                    return new ModelMeasureFormatSnapshot(
                        kind,
                        dateFormatString: ReadBoundedOptionalString(
                            () => (object?)format.FormatString,
                            MaximumFormatCharacters,
                            "date format string"));
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        private static ModelMeasureFormatSnapshot ToFormatSnapshot(PivotMeasureFormat format)
        {
            if (format == null) throw new ArgumentNullException(nameof(format));
            switch (format.Kind)
            {
                case PivotMeasureFormatKind.WholeNumber:
                    return new ModelMeasureFormatSnapshot(
                        ExcelModelMeasureFormatKind.WholeNumber,
                        useThousandsSeparator: format.UseThousandsSeparator);
                case PivotMeasureFormatKind.DecimalNumber:
                    return new ModelMeasureFormatSnapshot(
                        ExcelModelMeasureFormatKind.DecimalNumber,
                        format.DecimalPlaces,
                        format.UseThousandsSeparator);
                case PivotMeasureFormatKind.Currency:
                    return new ModelMeasureFormatSnapshot(
                        ExcelModelMeasureFormatKind.Currency,
                        format.DecimalPlaces,
                        currencySymbol: format.CurrencySymbolOrCode ?? string.Empty);
                case PivotMeasureFormatKind.Percentage:
                    return new ModelMeasureFormatSnapshot(
                        ExcelModelMeasureFormatKind.PercentageNumber,
                        format.DecimalPlaces,
                        format.UseThousandsSeparator);
                case PivotMeasureFormatKind.PercentagePoints:
                    // The compiler emits percentage-point results multiplied
                    // by 100, so they are decimal numbers, not percentage cells.
                    return new ModelMeasureFormatSnapshot(
                        ExcelModelMeasureFormatKind.DecimalNumber,
                        format.DecimalPlaces,
                        format.UseThousandsSeparator);
                default:
                    throw new NotSupportedException(
                        "The compiled measure format is not supported by Excel model measures.");
            }
        }

        private static object ReadModelFormatTemplate(
            object modelObject,
            ExcelModelMeasureFormatKind kind)
        {
            dynamic model = modelObject;
            switch (kind)
            {
                case ExcelModelMeasureFormatKind.General:
                    return ReadRequired(
                        () => (object?)model.ModelFormatGeneral,
                        "Excel did not expose ModelFormatGeneral.");
                case ExcelModelMeasureFormatKind.Boolean:
                    return ReadRequired(
                        () => (object?)model.ModelFormatBoolean,
                        "Excel did not expose ModelFormatBoolean.");
                case ExcelModelMeasureFormatKind.WholeNumber:
                    return ReadRequired(
                        () => (object?)model.ModelFormatWholeNumber,
                        "Excel did not expose ModelFormatWholeNumber.");
                case ExcelModelMeasureFormatKind.DecimalNumber:
                    return ReadRequired(
                        () => (object?)model.ModelFormatDecimalNumber,
                        "Excel did not expose ModelFormatDecimalNumber.");
                case ExcelModelMeasureFormatKind.PercentageNumber:
                    return ReadRequired(
                        () => (object?)model.ModelFormatPercentageNumber,
                        "Excel did not expose ModelFormatPercentageNumber.");
                case ExcelModelMeasureFormatKind.ScientificNumber:
                    return ReadRequired(
                        () => (object?)model.ModelFormatScientificNumber,
                        "Excel did not expose ModelFormatScientificNumber.");
                case ExcelModelMeasureFormatKind.Currency:
                    return ReadRequired(
                        () => (object?)model.ModelFormatCurrency,
                        "Excel did not expose ModelFormatCurrency.");
                case ExcelModelMeasureFormatKind.Date:
                    return ReadRequired(
                        () => (object?)model.ModelFormatDate,
                        "Excel did not expose ModelFormatDate.");
                default:
                    throw new ArgumentOutOfRangeException(nameof(kind));
            }
        }

        private IReadOnlyList<ModelTableHandle> ReadModelTables(BoundModelMeasureTarget target)
        {
            dynamic model = target.Model;
            object collection = ReadRequired(
                () => (object?)model.ModelTables,
                "Excel did not expose Data Model tables.");
            var result = new List<ModelTableHandle>();
            foreach (object tableObject in ReadCollection(
                         collection,
                         MaximumModelTables,
                         "Data Model tables"))
            {
                dynamic table = tableObject;
                string name = ReadBoundedRequiredString(
                    () => (object?)table.Name,
                    MaximumNameCharacters,
                    "Data Model table name");
                object sourceConnection = ReadRequired(
                    () => (object?)table.SourceWorkbookConnection,
                    "Excel did not expose the Data Model table source connection.");
                dynamic connection = sourceConnection;
                string connectionName = ReadBoundedRequiredString(
                    () => (object?)connection.Name,
                    MaximumNameCharacters,
                    "Data Model table source connection name");
                int connectionType = ReadRequiredInt(
                    () => (object?)connection.Type,
                    "Data Model table source connection type");
                string lineage = PivotPlusFingerprint.Create(
                    "model.table.lineage.v1",
                    CanonicalToken(name) + CanonicalToken(connectionName) +
                    connectionType.ToString(CultureInfo.InvariantCulture));
                result.Add(new ModelTableHandle(
                    tableObject,
                    name,
                    sourceConnection,
                    lineage));
            }

            DemandUniqueNames(result.Select(item => item.Name), "Data Model tables");
            return result;
        }

        private IReadOnlyList<MeasureHandle> ReadModelMeasures(
            BoundModelMeasureTarget target,
            IReadOnlyList<ModelTableHandle> tables)
        {
            dynamic model = target.Model;
            object collection = ReadRequired(
                () => (object?)model.ModelMeasures,
                "Excel did not expose Data Model measures.");
            var result = ReadCollection(
                    collection,
                    MaximumModelMeasures,
                    "Data Model measures")
                .Select(item => ReadMeasureHandle(item, tables))
                .ToList();
            DemandUniqueNames(
                result.Select(item => item.Snapshot.Name),
                "Data Model measures");
            return result;
        }

        private MeasureHandle ReadMeasureHandle(
            object measureObject,
            IReadOnlyList<ModelTableHandle> tables)
        {
            dynamic measure = measureObject;
            string name = ReadBoundedRequiredString(
                () => (object?)measure.Name,
                MaximumNameCharacters,
                "Data Model measure name");
            object associatedTable = ReadRequired(
                () => (object?)measure.AssociatedTable,
                "Excel did not expose the measure associated table.");
            List<ModelTableHandle> matchingTables = tables
                .Where(table => ComObjectIdentity.AreSame(table.Native, associatedTable))
                .ToList();
            if (matchingTables.Count != 1)
            {
                throw new InvalidOperationException(
                    "The measure is not associated with exactly one live Data Model table.");
            }

            string formula = ReadBoundedRequiredString(
                () => (object?)measure.Formula,
                MaximumFormulaCharacters,
                "Data Model measure formula");
            string description = ReadBoundedOptionalString(
                () => (object?)measure.Description,
                MaximumDescriptionCharacters,
                "Data Model measure description");
            object formatObject = ReadRequired(
                () => (object?)measure.FormatInformation,
                "Excel did not expose the measure format.");
            ModelMeasureFormatSnapshot format = ReadFormat(formatObject);
            ModelTableHandle table = matchingTables[0];
            string fingerprint = PivotModelMeasureCanonical.CreateLiveFingerprint(
                name,
                table.Name,
                table.LineageFingerprint,
                formula,
                description,
                format);
            return new MeasureHandle(
                measureObject,
                new LiveModelMeasureSnapshot(
                    name,
                    table.Name,
                    table.LineageFingerprint,
                    formula,
                    description,
                    format,
                    fingerprint));
        }

        private MeasureHandle? FindMeasure(
            BoundModelMeasureTarget target,
            IReadOnlyList<ModelTableHandle> tables,
            string name)
        {
            List<MeasureHandle> matches = ReadModelMeasures(target, tables)
                .Where(item => string.Equals(
                    item.Snapshot.Name,
                    name,
                    StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (matches.Count > 1)
            {
                throw new InvalidOperationException(
                    "Excel exposed multiple measures with the same name.");
            }

            return matches.Count == 0 ? null : matches[0];
        }

        private static ModelTableHandle ResolveTable(
            IReadOnlyList<ModelTableHandle> tables,
            string name,
            string? expectedLineageFingerprint)
        {
            List<ModelTableHandle> matches = tables
                .Where(item => string.Equals(
                    item.Name,
                    name,
                    StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (matches.Count != 1)
            {
                throw new InvalidOperationException(
                    "Excel did not expose exactly one requested Data Model table.");
            }

            ModelTableHandle table = matches[0];
            if (!string.Equals(table.Name, name, StringComparison.Ordinal) ||
                (expectedLineageFingerprint != null &&
                 !string.Equals(
                     table.LineageFingerprint,
                     expectedLineageFingerprint,
                     StringComparison.Ordinal)))
            {
                throw new InvalidOperationException(
                    "The requested Data Model table lineage changed.");
            }

            return table;
        }

        private static void DemandStillBound(BoundModelMeasureTarget target)
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
            if (!string.Equals(
                    workbookId,
                    target.Identity.WorkbookId,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The bound workbook identity changed.");
            }

            string pivotName = ReadBoundedRequiredString(
                () => (object?)pivot.Name,
                MaximumNameCharacters,
                "selected PivotTable name");
            string worksheetName = ReadBoundedRequiredString(
                () => (object?)nativeWorksheet.Name,
                MaximumNameCharacters,
                "selected PivotTable worksheet name");
            if (!string.Equals(
                    pivotName,
                    target.Identity.PivotTableName,
                    StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(
                    worksheetName,
                    target.Identity.WorksheetName,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "The selected PivotTable identity changed.");
            }

            dynamic nativeWorkbook = target.Workbook;
            object liveModel = ReadRequired(
                () => (object?)nativeWorkbook.Model,
                "Excel did not expose the bound workbook Data Model.");
            if (!ComObjectIdentity.AreSame(liveModel, target.Model))
            {
                throw new InvalidOperationException(
                    "The bound workbook Data Model object changed.");
            }

            dynamic model = target.Model;
            object liveModelConnection = ReadRequired(
                () => (object?)model.DataModelConnection,
                "Excel did not expose the bound Data Model connection.");
            if (!ComObjectIdentity.AreSame(
                    liveModelConnection,
                    target.DataModelConnection))
            {
                throw new InvalidOperationException(
                    "The workbook Data Model connection changed.");
            }

            dynamic nativeModelConnection = liveModelConnection;
            if (ReadRequiredInt(
                    () => (object?)nativeModelConnection.Type,
                    "bound Data Model connection type") != DataModelConnectionType)
            {
                throw new InvalidOperationException(
                    "The bound workbook connection is no longer the Data Model connection.");
            }

            dynamic cache = ReadPivotCache(target.PivotTable);
            if (!ReadRequiredBoolean(
                    () => (object?)cache.OLAP,
                    "selected PivotCache.OLAP"))
            {
                throw new InvalidOperationException(
                    "The selected PivotTable is no longer a Data Model PivotTable.");
            }

            object connection = ReadRequired(
                () => (object?)cache.WorkbookConnection,
                "Excel did not expose the selected PivotCache connection.");
            if (!ComObjectIdentity.AreSame(connection, target.DataModelConnection))
            {
                throw new InvalidOperationException(
                    "The selected PivotTable no longer uses the bound Data Model connection.");
            }
        }

        private static bool IsWorkbookModelPivot(
            object pivotTable,
            object dataModelConnection)
        {
            dynamic cache = ReadPivotCache(pivotTable);
            if (!ReadRequiredBoolean(
                    () => (object?)cache.OLAP,
                    "workbook PivotCache.OLAP"))
            {
                return false;
            }

            object connection = ReadRequired(
                () => (object?)cache.WorkbookConnection,
                "Excel did not expose an OLAP PivotTable workbook connection.");
            return ComObjectIdentity.AreSame(connection, dataModelConnection);
        }

        private static ModelPivotUsageSnapshot ReadPivotUsage(
            object pivotTable,
            string worksheetName,
            bool isSelected,
            IReadOnlyDictionary<string, string> modelMeasureNamesByUniqueName)
        {
            dynamic pivot = pivotTable;
            string pivotName = ReadBoundedRequiredString(
                () => (object?)pivot.Name,
                MaximumNameCharacters,
                "PivotTable name");
            IReadOnlyList<DataFieldHandle> fields = ReadDataFieldHandles(
                pivotTable,
                modelMeasureNamesByUniqueName);
            ReadValuesAxis(
                pivotTable,
                fields.Count,
                out PivotValuesAxis axis,
                out int position);
            return new ModelPivotUsageSnapshot(
                worksheetName,
                pivotName,
                isSelected,
                fields.Select(item => item.Snapshot),
                axis,
                position);
        }

        private static ModelPivotUsageSnapshot ReadSelectedUsage(
            BoundModelMeasureTarget target,
            IReadOnlyDictionary<string, string> modelMeasureNamesByUniqueName)
        {
            return ReadPivotUsage(
                target.PivotTable,
                target.Identity.WorksheetName,
                isSelected: true,
                modelMeasureNamesByUniqueName);
        }

        private static IReadOnlyList<DataFieldHandle> ReadDataFieldHandles(
            object pivotTable,
            IReadOnlyDictionary<string, string> modelMeasureNamesByUniqueName)
        {
            dynamic pivot = pivotTable;
            object collection = ReadRequired(
                () => (object?)pivot.DataFields,
                "Excel did not expose PivotTable DataFields.");
            var result = new List<DataFieldHandle>();
            foreach (object fieldObject in ReadCollection(
                         collection,
                         MaximumDataFields,
                         "PivotTable DataFields"))
            {
                dynamic field = fieldObject;
                object cubeObject = ReadRequired(
                    () => (object?)field.CubeField,
                    "Excel did not expose a DataField CubeField.");
                dynamic cube = cubeObject;
                string uniqueName = ReadBoundedRequiredString(
                    () => (object?)cube.Name,
                    MaximumNameCharacters * 4,
                    "DataField CubeField unique name");
                int cubeType = ReadRequiredInt(
                    () => (object?)cube.CubeFieldType,
                    "DataField CubeField type");
                if (cubeType != CubeFieldTypeMeasure)
                {
                    throw new NotSupportedException(
                        "A Values field is not backed by an exact cube measure.");
                }

                int cubeSubType = ReadRequiredInt(
                    () => (object?)cube.CubeFieldSubType,
                    "DataField CubeField subtype");
                string caption = ReadBoundedRequiredString(
                    () => (object?)field.Caption,
                    MaximumNameCharacters,
                    "DataField caption");
                string numberFormat = ReadBoundedOptionalString(
                    () => (object?)field.NumberFormat,
                    MaximumFormatCharacters,
                    "DataField number format");
                int position = ReadRequiredPositiveInt(
                    () => (object?)field.Position,
                    "DataField position");
                string? modelMeasureName = null;
                bool authored = cubeSubType != CubeFieldSubTypeImplicitMeasure &&
                    modelMeasureNamesByUniqueName.TryGetValue(
                        uniqueName,
                        out modelMeasureName);
                result.Add(new DataFieldHandle(
                    fieldObject,
                    cubeObject,
                    new ModelDataFieldSnapshot(
                        uniqueName,
                        caption,
                        PivotMeasurePlacementFingerprint.CreateCaptionFingerprint(caption),
                        numberFormat,
                        position,
                        authored,
                        authored ? modelMeasureName : null)));
            }

            if (result.Select(item => item.Snapshot.Position).Distinct().Count() != result.Count ||
                result.Any(item => item.Snapshot.Position > result.Count))
            {
                throw new NotSupportedException(
                    "Excel exposed an invalid or duplicate Values field position.");
            }

            return result.OrderBy(item => item.Snapshot.Position).ToList();
        }

        private static void ReadValuesAxis(
            object pivotTable,
            int dataFieldCount,
            out PivotValuesAxis axis,
            out int position)
        {
            if (dataFieldCount <= 1)
            {
                axis = PivotValuesAxis.Automatic;
                // There is no live DataPivotField in this state. The portable
                // layout contract uses one as the bounded sentinel position.
                position = 1;
                return;
            }

            dynamic pivot = pivotTable;
            object dataPivotField = ReadRequired(
                () => (object?)pivot.DataPivotField,
                "Excel did not expose the Values pseudo-axis for a multi-value PivotTable.");
            dynamic nativeField = dataPivotField;
            int orientation = ReadRequiredInt(
                () => (object?)nativeField.Orientation,
                "Values pseudo-axis orientation");
            if (orientation == OrientationRow)
            {
                axis = PivotValuesAxis.Rows;
            }
            else if (orientation == OrientationColumn)
            {
                axis = PivotValuesAxis.Columns;
            }
            else
            {
                throw new NotSupportedException(
                    "A multi-value PivotTable has an invalid Values pseudo-axis orientation.");
            }

            position = ReadRequiredPositiveInt(
                () => (object?)nativeField.Position,
                "Values pseudo-axis position");
        }

        private static void ApplyValuesAxis(
            object pivotTable,
            PivotValuesAxis axis,
            int position,
            int plannedDataFieldCount)
        {
            if (plannedDataFieldCount <= 1)
            {
                if (axis != PivotValuesAxis.Automatic || position != 1)
                {
                    throw new InvalidOperationException(
                        "A single-value layout must use the automatic Values axis.");
                }

                return;
            }

            if ((axis != PivotValuesAxis.Rows && axis != PivotValuesAxis.Columns) ||
                position <= 0)
            {
                throw new InvalidOperationException(
                    "A multi-value layout requires an exact row or column Values pseudo-axis position.");
            }

            dynamic pivot = pivotTable;
            object dataPivotField = ReadRequired(
                () => (object?)pivot.DataPivotField,
                "Excel did not expose the Values pseudo-axis for placement.");
            dynamic nativeField = dataPivotField;
            nativeField.Orientation = axis == PivotValuesAxis.Rows
                ? OrientationRow
                : OrientationColumn;
            nativeField.Position = position;
        }

        private static CubeFieldHandle ResolveAuthoredMeasureCubeField(
            object pivotTable,
            string measureName)
        {
            return ResolveMeasureCubeField(
                pivotTable,
                MeasureCubeUniqueName(measureName),
                allowImplicit: false);
        }

        private static CubeFieldHandle ResolveMeasureCubeField(
            object pivotTable,
            string uniqueName,
            bool allowImplicit)
        {
            CubeFieldHandle field = FindCubeField(pivotTable, uniqueName) ??
                throw new InvalidOperationException(
                    "Excel did not expose the exact authored measure CubeField.");
            if (field.CubeFieldType != CubeFieldTypeMeasure ||
                (!allowImplicit &&
                 field.CubeFieldSubType == CubeFieldSubTypeImplicitMeasure))
            {
                throw new NotSupportedException(
                    "The requested Values field is not an authored model measure.");
            }

            return field;
        }

        private static CubeFieldHandle? FindCubeField(
            object pivotTable,
            string uniqueName)
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
                string name = ReadBoundedRequiredString(
                    () => (object?)cube.Name,
                    MaximumNameCharacters * 4,
                    "CubeField unique name");
                if (!string.Equals(name, uniqueName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                matches.Add(new CubeFieldHandle(
                    cubeObject,
                    name,
                    ReadRequiredInt(
                        () => (object?)cube.CubeFieldType,
                        "CubeField type"),
                    ReadRequiredInt(
                        () => (object?)cube.CubeFieldSubType,
                        "CubeField subtype")));
            }

            if (matches.Count > 1)
            {
                throw new InvalidOperationException(
                    "Excel exposed duplicate CubeFields for the requested measure.");
            }

            if (matches.Count == 1 &&
                !string.Equals(matches[0].UniqueName, uniqueName, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The exact CubeField name casing changed.");
            }

            return matches.Count == 0 ? null : matches[0];
        }

        private static DataFieldHandle DemandSingleDataField(
            IReadOnlyDictionary<string, List<DataFieldHandle>> fields,
            string uniqueName)
        {
            if (!fields.TryGetValue(uniqueName, out List<DataFieldHandle>? matches) ||
                matches.Count != 1)
            {
                throw new InvalidOperationException(
                    "Excel did not expose exactly one Values field for the authored measure.");
            }

            return matches[0];
        }

        private static void ValidatePlacement(
            PivotMeasurePlacementPlan placement,
            IReadOnlyDictionary<string, DesiredModelMeasure> definitionsById)
        {
            int[] positions = placement.Values.Select(item => item.Position).ToArray();
            if (!positions.OrderBy(item => item)
                    .SequenceEqual(Enumerable.Range(1, positions.Length)))
            {
                throw new InvalidOperationException(
                    "The Values placement must be a complete one-based sequence.");
            }

            var generated = new HashSet<string>(StringComparer.Ordinal);
            var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (PivotMeasureValuePlacement value in placement.Values)
            {
                if (value.IsGeneratedMeasure)
                {
                    if (string.IsNullOrWhiteSpace(value.DefinitionId) ||
                        !definitionsById.ContainsKey(value.DefinitionId!) ||
                        !generated.Add(value.DefinitionId!))
                    {
                        throw new InvalidOperationException(
                            "The Values placement contains an unknown or duplicate generated measure.");
                    }
                }
                else
                {
                    PivotExistingDataFieldIdentity identity = value.ExistingDataField!;
                    string key = ExistingFieldKey(
                        identity.UniqueName,
                        identity.CurrentCaptionFingerprint,
                        identity.CurrentNumberFormatFingerprint,
                        identity.CurrentPosition);
                    if (identity.CurrentPosition < 1)
                    {
                        throw new InvalidOperationException(
                            "The Values placement contains an invalid preview position.");
                    }
                    if (!existing.Add(key))
                    {
                        throw new InvalidOperationException(
                            "The Values placement contains a duplicate existing field.");
                    }
                }
            }

            if (placement.Values.Count <= 1)
            {
                if (placement.ValuesAxis != PivotValuesAxis.Automatic ||
                    placement.ValuesPosition != 1)
                {
                    throw new InvalidOperationException(
                        "A single-value layout must use the automatic Values axis.");
                }
            }
            else if ((placement.ValuesAxis != PivotValuesAxis.Rows &&
                      placement.ValuesAxis != PivotValuesAxis.Columns) ||
                     placement.ValuesPosition <= 0)
            {
                throw new InvalidOperationException(
                    "A multi-value layout requires a row or column Values pseudo-axis.");
            }
        }

        private static bool MatchesDesiredCore(
            LiveModelMeasureSnapshot live,
            DesiredModelMeasure desired,
            ModelTableHandle table,
            ExcelModelMeasureFormatKind formatKind)
        {
            return string.Equals(live.Name, desired.Name, StringComparison.Ordinal) &&
                   string.Equals(
                       live.AssociatedTableName,
                       table.Name,
                       StringComparison.Ordinal) &&
                   string.Equals(
                       live.AssociatedTableLineageFingerprint,
                       table.LineageFingerprint,
                       StringComparison.Ordinal) &&
                   string.Equals(
                       live.Description,
                       desired.DescriptionMarker,
                       StringComparison.Ordinal) &&
                   live.Format.Kind == formatKind;
        }

        private static bool MatchesDesired(
            LiveModelMeasureSnapshot live,
            DesiredModelMeasure desired,
            ModelTableHandle table,
            ModelMeasureFormatSnapshot format)
        {
            return MatchesDesiredCore(live, desired, table, format.Kind) &&
                   SameFormat(live.Format, format);
        }

        private static bool MatchesRestoreCore(
            LiveModelMeasureSnapshot live,
            LiveModelMeasureSnapshot before)
        {
            return string.Equals(live.Name, before.Name, StringComparison.Ordinal) &&
                   string.Equals(
                       live.AssociatedTableName,
                       before.AssociatedTableName,
                       StringComparison.Ordinal) &&
                   string.Equals(
                       live.AssociatedTableLineageFingerprint,
                       before.AssociatedTableLineageFingerprint,
                       StringComparison.Ordinal) &&
                   string.Equals(live.Formula, before.Formula, StringComparison.Ordinal) &&
                   string.Equals(live.Description, before.Description, StringComparison.Ordinal) &&
                   live.Format.Kind == before.Format.Kind;
        }

        private static void DemandDesired(
            LiveModelMeasureSnapshot live,
            DesiredModelMeasure desired,
            ModelTableHandle table,
            ModelMeasureFormatSnapshot format)
        {
            if (!MatchesDesired(live, desired, table, format))
            {
                throw new InvalidOperationException(
                    "Excel did not preserve the exact requested measure definition.");
            }
        }

        private static void DemandSameSnapshot(
            LiveModelMeasureSnapshot live,
            LiveModelMeasureSnapshot expected,
            string label)
        {
            if (!SameSnapshot(live, expected))
            {
                throw new InvalidOperationException(
                    "The exact " + label + " changed after preview.");
            }
        }

        private static bool SameSnapshot(
            LiveModelMeasureSnapshot left,
            LiveModelMeasureSnapshot right)
        {
            return string.Equals(
                left.LiveFingerprint,
                right.LiveFingerprint,
                StringComparison.Ordinal) &&
                string.Equals(left.Name, right.Name, StringComparison.Ordinal) &&
                string.Equals(
                    left.AssociatedTableName,
                    right.AssociatedTableName,
                    StringComparison.Ordinal) &&
                string.Equals(
                    left.AssociatedTableLineageFingerprint,
                    right.AssociatedTableLineageFingerprint,
                    StringComparison.Ordinal) &&
                string.Equals(left.Formula, right.Formula, StringComparison.Ordinal) &&
                string.Equals(left.Description, right.Description, StringComparison.Ordinal) &&
                SameFormat(left.Format, right.Format);
        }

        private static bool SameFormat(
            ModelMeasureFormatSnapshot left,
            ModelMeasureFormatSnapshot right)
        {
            return left.Kind == right.Kind &&
                   left.DecimalPlaces == right.DecimalPlaces &&
                   left.UseThousandsSeparator == right.UseThousandsSeparator &&
                   string.Equals(
                       left.CurrencySymbol,
                       right.CurrencySymbol,
                       StringComparison.Ordinal) &&
                   string.Equals(
                       left.DateFormatString,
                       right.DateFormatString,
                       StringComparison.Ordinal);
        }

        private static bool FormatNeedsPropertyWrites(ModelMeasureFormatSnapshot format)
        {
            return format.Kind != ExcelModelMeasureFormatKind.General &&
                   format.Kind != ExcelModelMeasureFormatKind.Boolean;
        }

        private static ExcelModelMeasureFormatKind ParseFormatKind(string rawTypeName)
        {
            string typeName = rawTypeName ?? string.Empty;
            int separator = typeName.LastIndexOf('.');
            if (separator >= 0) typeName = typeName.Substring(separator + 1);
            typeName = typeName.TrimStart('_');
            if (typeName.EndsWith("Class", StringComparison.Ordinal))
            {
                typeName = typeName.Substring(0, typeName.Length - "Class".Length);
            }

            switch (typeName)
            {
                case "ModelFormatGeneral":
                    return ExcelModelMeasureFormatKind.General;
                case "ModelFormatBoolean":
                    return ExcelModelMeasureFormatKind.Boolean;
                case "ModelFormatWholeNumber":
                    return ExcelModelMeasureFormatKind.WholeNumber;
                case "ModelFormatDecimalNumber":
                    return ExcelModelMeasureFormatKind.DecimalNumber;
                case "ModelFormatPercentageNumber":
                    return ExcelModelMeasureFormatKind.PercentageNumber;
                case "ModelFormatScientificNumber":
                    return ExcelModelMeasureFormatKind.ScientificNumber;
                case "ModelFormatCurrency":
                    return ExcelModelMeasureFormatKind.Currency;
                case "ModelFormatDate":
                    return ExcelModelMeasureFormatKind.Date;
                default:
                    throw new NotSupportedException(
                        "Excel exposed an unknown model-measure format COM type.");
            }
        }

        private static string ResolveRuntimeTypeName(object value)
        {
            if (value == null) throw new ArgumentNullException(nameof(value));
            if (!Marshal.IsComObject(value))
            {
                return value.GetType().Name;
            }

            var dispatch = (IDispatch)value;
            uint count;
            int result = dispatch.GetTypeInfoCount(out count);
            Marshal.ThrowExceptionForHR(result);
            if (count != 1)
            {
                throw new NotSupportedException(
                    "The model format COM object did not expose exactly one type library entry.");
            }

            ITypeInfo? typeInfo = null;
            try
            {
                result = dispatch.GetTypeInfo(
                    0,
                    unchecked((uint)CultureInfo.CurrentCulture.LCID),
                    out typeInfo!);
                Marshal.ThrowExceptionForHR(result);
                return Marshal.GetTypeInfoName(typeInfo);
            }
            finally
            {
                if (typeInfo != null && Marshal.IsComObject(typeInfo))
                {
                    Marshal.ReleaseComObject(typeInfo);
                }
            }
        }

        private static object ReadPivotCache(object pivotTable)
        {
            dynamic pivot = pivotTable;
            if (PivotLateBound.TryRead(
                    () => (object?)pivot.PivotCache(),
                    out object? methodValue) && methodValue != null)
            {
                return methodValue;
            }

            return ReadRequired(
                () => (object?)pivot.PivotCache,
                "Excel did not expose the PivotTable cache.");
        }

        private static object ReadRequiredPivotTables(dynamic worksheet)
        {
            if (PivotLateBound.TryRead(
                    () => (object?)worksheet.PivotTables(),
                    out object? methodValue) && methodValue != null)
            {
                return methodValue;
            }

            return ReadRequired(
                () => (object?)worksheet.PivotTables,
                "Excel did not expose worksheet PivotTables.");
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
                object item = ReadRequired(
                    () => (object?)collection.Item(captured),
                    "Excel did not expose item " + captured.ToString(CultureInfo.InvariantCulture) +
                    " in " + label + ".");
                result.Add(item);
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
                throw new InvalidOperationException(
                    "Excel did not expose " + label + ".");
            }

            string result;
            if (value == null)
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
            if (result.Length > maximum || result.Any(character =>
                    char.IsControl(character) &&
                    character != '\r' &&
                    character != '\n' &&
                    character != '\t'))
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

        private static int ReadFormatDecimalPlaces(Func<object?> reader, string label)
        {
            int value = ReadRequiredInt(reader, label);
            if (value < 0 || value > MaximumFormatDecimalPlaces)
            {
                throw new NotSupportedException(
                    "Excel exposed unbounded model-format decimal places.");
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

        private static void DemandUniqueNames(IEnumerable<string> names, string label)
        {
            if (names.GroupBy(item => item, StringComparer.OrdinalIgnoreCase)
                .Any(group => group.Count() != 1))
            {
                throw new InvalidOperationException(
                    "Excel exposed duplicate names in " + label + ".");
            }
        }

        private static void DemandUsageIdentity(
            ModelPivotUsageSnapshot usage,
            PivotTargetIdentity target)
        {
            if (!usage.IsSelectedTarget ||
                !string.Equals(
                    usage.WorksheetName,
                    target.WorksheetName,
                    StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(
                    usage.PivotTableName,
                    target.PivotTableName,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "The selected PivotTable usage snapshot is not the bound target.");
            }
        }

        private static string MeasureCubeUniqueName(string measureName)
        {
            return "[Measures].[" + measureName.Replace("]", "]]" ) + "]";
        }

        private static string ExistingFieldKey(
            string uniqueName,
            string captionFingerprint,
            string numberFormatFingerprint,
            int position)
        {
            return uniqueName + "\u001f" + captionFingerprint + "\u001f" +
                   numberFormatFingerprint + "\u001f" +
                   position.ToString(CultureInfo.InvariantCulture);
        }

        private static string CanonicalToken(string value)
        {
            return value.Length.ToString(CultureInfo.InvariantCulture) + ":" + value + "|";
        }

        private sealed class ModelTableHandle
        {
            public ModelTableHandle(
                object native,
                string name,
                object sourceConnection,
                string lineageFingerprint)
            {
                Native = native;
                Name = name;
                SourceConnection = sourceConnection;
                LineageFingerprint = lineageFingerprint;
            }

            public object Native { get; }
            public string Name { get; }
            public object SourceConnection { get; }
            public string LineageFingerprint { get; }
        }

        private sealed class MeasureHandle
        {
            public MeasureHandle(object native, LiveModelMeasureSnapshot snapshot)
            {
                Native = native;
                Snapshot = snapshot;
            }

            public object Native { get; }
            public LiveModelMeasureSnapshot Snapshot { get; }
        }

        private sealed class DataFieldHandle
        {
            public DataFieldHandle(
                object native,
                object cubeField,
                ModelDataFieldSnapshot snapshot)
            {
                Native = native;
                CubeField = cubeField;
                Snapshot = snapshot;
            }

            public object Native { get; }
            public object CubeField { get; }
            public ModelDataFieldSnapshot Snapshot { get; }
        }

        private sealed class CubeFieldHandle
        {
            public CubeFieldHandle(
                object native,
                string uniqueName,
                int cubeFieldType,
                int cubeFieldSubType)
            {
                Native = native;
                UniqueName = uniqueName;
                CubeFieldType = cubeFieldType;
                CubeFieldSubType = cubeFieldSubType;
            }

            public object Native { get; }
            public string UniqueName { get; }
            public int CubeFieldType { get; }
            public int CubeFieldSubType { get; }
        }

        [ComImport]
        [Guid("00020400-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IDispatch
        {
            [PreserveSig]
            int GetTypeInfoCount(out uint count);

            [PreserveSig]
            int GetTypeInfo(uint index, uint lcid, out ITypeInfo typeInfo);
        }
    }
}
