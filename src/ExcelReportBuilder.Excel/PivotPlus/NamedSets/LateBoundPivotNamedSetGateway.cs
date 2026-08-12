using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using ExcelReportBuilder.Core.PivotPlus;
using ExcelReportBuilder.Core.PivotPlus.Calculations;
using ExcelReportBuilder.Core.PivotPlus.NamedSets;
using ExcelReportBuilder.Excel.PivotPlus.Persistence;

namespace ExcelReportBuilder.Excel.PivotPlus.NamedSets
{
    /// <summary>
    /// Strict late-bound boundary for Data Model named-set discovery, capture,
    /// reconciliation, and host mutation.
    /// </summary>
    internal sealed class LateBoundPivotNamedSetGateway : IPivotNamedSetGateway
    {
        private const int DataModelConnectionType = 7;
        private const int CubeFieldTypeHierarchy = 1;
        private const int CubeFieldTypeSet = 3;
        private const int CalculatedMemberTypeSet = 1;
        private const int MaximumWorksheets = 1024;
        private const int MaximumPivotTables = 4096;
        private const int MaximumModelTables = 512;
        private const int MaximumCubeFields = 4096;
        private const int MaximumHierarchies = 64;
        private const int MaximumLevels = 256;
        private const int MaximumMembers = 4096;
        private const int MaximumCalculatedMembers = 512;
        private const int MaximumModelMeasures = 512;
        private const int MaximumNameCharacters = 255;
        private const int MaximumProviderUniqueNameCharacters = 2048;
        private const int MaximumCaptionCharacters = 1024;
        private const int MaximumFormulaCharacters = 24 * 1024;

        public BoundPivotNamedSetTarget Bind(
            object workbook,
            object pivotTable,
            PivotTableContext context)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            if (context == null) throw new ArgumentNullException(nameof(context));

            PivotLayoutDefinition definition = context.Definition;
            const PivotCapability required = PivotCapability.DataModel |
                                             PivotCapability.CalculatedMembers |
                                             PivotCapability.NamedSets;
            if (!context.IsConnected ||
                !context.SourceFieldsComplete ||
                definition.Source.Kind != PivotSourceKind.DataModel ||
                (definition.Source.Capabilities & required) != required)
            {
                throw new NotSupportedException(
                    "Named sets require the selected native workbook Data Model PivotTable.");
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

            ReadCollection(
                ReadRequired(
                    () => (object?)nativeModel.ModelTables,
                    "Excel did not expose Data Model tables."),
                MaximumModelTables,
                "Data Model tables");
            ReadCollection(
                ReadRequired(
                    () => (object?)pivot.CubeFields,
                    "Excel did not expose selected PivotTable CubeFields."),
                MaximumCubeFields,
                "selected PivotTable CubeFields");
            ReadCollection(
                ReadRequired(
                    () => (object?)pivot.CalculatedMembers,
                    "Excel did not expose selected PivotTable CalculatedMembers."),
                MaximumCalculatedMembers,
                "selected PivotTable CalculatedMembers");

            return new BoundPivotNamedSetTarget(
                workbook,
                pivotTable,
                model,
                dataModelConnection,
                expected);
        }

        public PivotNamedSetSchemaDiscoveryResult DiscoverSchema(
            BoundPivotNamedSetTarget target)
        {
            if (target == null) throw new ArgumentNullException(nameof(target));
            DemandStillBound(target);

            string modelLineage = ReadModelLineageFingerprint(target);
            dynamic pivot = target.PivotTable;
            object cubeFieldsObject = ReadRequired(
                () => (object?)pivot.CubeFields,
                "Excel did not expose selected PivotTable CubeFields.");
            IReadOnlyList<object> cubeFields = ReadCollection(
                cubeFieldsObject,
                MaximumCubeFields,
                "selected PivotTable CubeFields");
            var diagnostics = new List<PivotNamedSetDiscoveryDiagnostic>();
            var hierarchies = new List<PivotNamedSetHierarchySchema>();
            var hierarchyNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var levelNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var memberNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var catalogTokens = new List<string>();
            var levelCount = 0;
            var memberCount = 0;

            foreach (object cubeFieldObject in cubeFields)
            {
                dynamic cubeField = cubeFieldObject;
                if (ReadCubeFieldType(cubeField) != CubeFieldTypeHierarchy)
                {
                    continue;
                }

                if (hierarchies.Count >= MaximumHierarchies)
                {
                    throw new NotSupportedException(
                        "The Data Model hierarchy catalog exceeds its bounded limit.");
                }

                string hierarchyName = ReadProviderUniqueName(
                    cubeField,
                    "Data Model hierarchy unique name");
                if (!hierarchyNames.Add(hierarchyName))
                {
                    throw new InvalidOperationException(
                        "Excel exposed duplicate Data Model hierarchy identities.");
                }

                string hierarchyCaption = ReadBoundedOptionalString(
                    () => (object?)cubeField.Caption,
                    MaximumCaptionCharacters,
                    "Data Model hierarchy caption");
                string hierarchyId = PivotNamedSetCanonical.CreateStableCatalogId(
                    "hierarchy",
                    hierarchyName);
                var levels = new List<PivotNamedSetLevelSchema>();
                if (!TryReadCollectionMember(
                        () => (object?)cubeField.PivotFields,
                        () => (object?)cubeField.PivotFields(),
                        out object? pivotFieldsObject) ||
                    pivotFieldsObject == null)
                {
                    diagnostics.Add(new PivotNamedSetDiscoveryDiagnostic(
                        "PIVOT_SET_DISCOVERY_PIVOTFIELDS_UNAVAILABLE",
                        "hierarchies[" + hierarchyId + "]",
                        "Excel has not materialized PivotFields for this hierarchy."));
                }
                else
                {
                    IReadOnlyList<object> pivotFields = ReadCollection(
                        pivotFieldsObject,
                        MaximumLevels,
                        "existing hierarchy PivotFields");
                    for (var levelIndex = 0; levelIndex < pivotFields.Count; levelIndex++)
                    {
                        levelCount++;
                        if (levelCount > MaximumLevels)
                        {
                            throw new NotSupportedException(
                                "The Data Model level catalog exceeds its bounded limit.");
                        }

                        dynamic pivotField = pivotFields[levelIndex];
                        string levelName = ReadProviderUniqueName(
                            pivotField,
                            "Data Model level unique name");
                        if (!levelNames.Add(levelName))
                        {
                            throw new InvalidOperationException(
                                "Excel exposed duplicate Data Model level identities.");
                        }

                        string levelId = PivotNamedSetCanonical.CreateStableCatalogId(
                            "level",
                            levelName);
                        var members = new List<PivotNamedSetMemberSchema>();
                        bool membersComplete = TryReadCollectionMember(
                            () => (object?)pivotField.PivotItems,
                            () => (object?)pivotField.PivotItems(),
                            out object? pivotItemsObject) &&
                            pivotItemsObject != null;
                        if (!membersComplete)
                        {
                            diagnostics.Add(new PivotNamedSetDiscoveryDiagnostic(
                                "PIVOT_SET_DISCOVERY_PIVOTITEMS_UNAVAILABLE",
                                "levels[" + levelId + "]",
                                "Excel has not materialized PivotItems for this level."));
                        }
                        else
                        {
                            IReadOnlyList<object> pivotItems = ReadCollection(
                                pivotItemsObject!,
                                MaximumMembers,
                                "existing level PivotItems");
                            foreach (object pivotItemObject in pivotItems)
                            {
                                memberCount++;
                                if (memberCount > MaximumMembers)
                                {
                                    throw new NotSupportedException(
                                        "The Data Model member catalog exceeds its bounded limit.");
                                }

                                dynamic pivotItem = pivotItemObject;
                                string memberName = ReadProviderUniqueName(
                                    pivotItem,
                                    "Data Model member unique name");
                                if (!memberNames.Add(memberName))
                                {
                                    throw new InvalidOperationException(
                                        "Excel exposed duplicate Data Model member identities.");
                                }

                                string memberCaption = ReadBoundedOptionalString(
                                    () => (object?)pivotItem.Caption,
                                    MaximumCaptionCharacters,
                                    "Data Model member caption");
                                string memberId = PivotNamedSetCanonical.CreateStableCatalogId(
                                    "member",
                                    memberName);
                                members.Add(new PivotNamedSetMemberSchema(
                                    memberId,
                                    memberName,
                                    NullIfEmpty(memberCaption),
                                    parentMemberId: null,
                                    isAllMember: false));
                                catalogTokens.Add(CatalogToken(
                                    "member",
                                    levelName,
                                    memberName,
                                    memberCaption));
                            }
                        }

                        string levelCaption = ReadBoundedOptionalString(
                            () => (object?)pivotField.Caption,
                            MaximumCaptionCharacters,
                            "Data Model level caption");
                        levels.Add(new PivotNamedSetLevelSchema(
                            levelId,
                            levelName,
                            levelIndex,
                            membersComplete,
                            members));
                        catalogTokens.Add(CatalogToken(
                            "level",
                            hierarchyName,
                            levelName,
                            levelCaption,
                            levelIndex.ToString(CultureInfo.InvariantCulture),
                            membersComplete ? "complete" : "incomplete"));
                    }
                }

                hierarchies.Add(new PivotNamedSetHierarchySchema(
                    hierarchyId,
                    hierarchyName,
                    identityComplete: true,
                    levels,
                    NullIfEmpty(hierarchyCaption)));
                catalogTokens.Add(CatalogToken(
                    "hierarchy",
                    hierarchyName,
                    hierarchyCaption));
            }

            string sourceFingerprint = PivotNamedSetCanonical.CreateSourceFingerprint(
                modelLineage,
                catalogTokens);
            return new PivotNamedSetSchemaDiscoveryResult(
                new PivotNamedSetSchema(
                    sourceFingerprint,
                    PivotNamedSetProviderKind.DataModel,
                    hierarchies),
                diagnostics);
        }

        public PivotNamedSetWorkbookSnapshot Capture(BoundPivotNamedSetTarget target)
        {
            return CaptureCore(target, requireConnectionRefresh: true);
        }

        private PivotNamedSetWorkbookSnapshot CaptureCore(
            BoundPivotNamedSetTarget target,
            bool requireConnectionRefresh)
        {
            if (target == null) throw new ArgumentNullException(nameof(target));
            DemandStillBound(target);
            string sourceFingerprint = DiscoverSchema(target).Schema.SourceFingerprint;
            string modelLineage = ReadModelLineageFingerprint(target);

            dynamic workbook = target.Workbook;
            object worksheetsObject = ReadRequired(
                () => (object?)workbook.Worksheets,
                "Excel did not expose workbook worksheets.");
            var pivots = new List<PivotNamedSetPivotSnapshot>();
            foreach (object worksheetObject in ReadCollection(
                         worksheetsObject,
                         MaximumWorksheets,
                         "workbook worksheets"))
            {
                dynamic worksheet = worksheetObject;
                string worksheetName = ReadBoundedRequiredString(
                    () => (object?)worksheet.Name,
                    MaximumNameCharacters,
                    "worksheet name");
                object pivotTablesObject = ReadRequiredPivotTables(worksheet);
                foreach (object pivotTable in ReadCollection(
                             pivotTablesObject,
                             MaximumPivotTables,
                             "worksheet PivotTables"))
                {
                    if (!IsWorkbookModelPivot(
                            pivotTable,
                            target.DataModelConnection))
                    {
                        continue;
                    }

                    bool isSelected = ComObjectIdentity.AreSame(
                        pivotTable,
                        target.PivotTable);
                    pivots.Add(ReadPivotSnapshot(
                        pivotTable,
                        worksheetName,
                        isSelected,
                        sourceFingerprint,
                        requireConnectionRefresh,
                        modelLineage));
                }
            }

            if (pivots.Count(pivot => pivot.IsSelectedTarget) != 1)
            {
                throw new InvalidOperationException(
                    "Excel did not expose the selected PivotTable exactly once in this workbook.");
            }

            PivotNamedSetPivotSnapshot selected = pivots.Single(
                pivot => pivot.IsSelectedTarget);
            if (!string.Equals(
                    selected.WorksheetName,
                    target.Identity.WorksheetName,
                    StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(
                    selected.PivotTableName,
                    target.Identity.PivotTableName,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "The named-set snapshot is not the bound target.");
            }

            return new PivotNamedSetWorkbookSnapshot(
                pivots,
                sourceFingerprint,
                modelLineage);
        }

        public LivePivotNamedSetSnapshot CreateSet(
            BoundPivotNamedSetTarget target,
            DesiredPivotNamedSet definition)
        {
            if (target == null) throw new ArgumentNullException(nameof(target));
            if (definition == null) throw new ArgumentNullException(nameof(definition));
            DemandStillBound(target);
            DemandDesiredDefinition(definition);
            _ = DemandLiveCompatibility(target, definition);
            PivotNamedSetWorkbookSnapshot before = Capture(target);
            DemandSourceFingerprint(before, definition.SourceFingerprint);
            LivePivotNamedSetSnapshot? existing = FindSelectedArtifact(
                before,
                definition.Name);
            DemandSafeCreateUse(before, definition, existing);
            return CreateDesiredWithReconciliation(
                target,
                definition,
                before,
                existing);
        }

        public LivePivotNamedSetSnapshot ReplaceSet(
            BoundPivotNamedSetTarget target,
            LivePivotNamedSetSnapshot before,
            DesiredPivotNamedSet definition)
        {
            if (target == null) throw new ArgumentNullException(nameof(target));
            if (before == null) throw new ArgumentNullException(nameof(before));
            if (definition == null) throw new ArgumentNullException(nameof(definition));
            DemandStillBound(target);
            DemandDesiredDefinition(definition);
            _ = DemandLiveCompatibility(target, definition);
            if (!string.Equals(before.Name, definition.Name, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "A named-set replacement must retain its generated name.");
            }

            PivotNamedSetWorkbookSnapshot snapshot = Capture(target);
            DemandSourceFingerprint(snapshot, definition.SourceFingerprint);
            DemandSourceFingerprint(snapshot, before.SourceFingerprint);
            LivePivotNamedSetSnapshot current = DemandSelectedArtifact(
                snapshot,
                before.Name);
            DemandSameSnapshot(current, before, "named set");
            DemandSafeDestructiveUse(snapshot, before);
            try
            {
                DeleteSetCore(target, before);
                PivotNamedSetWorkbookSnapshot absent = CaptureStatic(target);
                DemandExactInventoryTransition(
                    snapshot,
                    absent,
                    before.Name,
                    before,
                    expectedAfter: null);
                LivePivotNamedSetSnapshot created = CreateDesiredWithReconciliation(
                    target,
                    definition,
                    absent,
                    baselineArtifact: null);
                PivotNamedSetWorkbookSnapshot completed = CaptureStatic(target);
                LivePivotNamedSetSnapshot desired = DemandSelectedArtifact(
                    completed,
                    definition.Name);
                DemandExactInventoryTransition(
                    snapshot,
                    completed,
                    before.Name,
                    before,
                    desired);
                DemandExactMeasureDependencies(
                    target,
                    definition.DirectMeasureDependencies);
                return created;
            }
            catch (Exception failure)
            {
                PivotNamedSetWorkbookSnapshot observedWorkbook;
                try
                {
                    observedWorkbook = CaptureForRecovery(target);
                }
                catch (Exception captureFailure)
                {
                    throw new PivotNamedSetRecoveryRequiredException(
                        "The failed named-set replacement could not be inventoried.",
                        captureFailure);
                }

                LivePivotNamedSetSnapshot? observed = FindSelectedArtifact(
                    observedWorkbook,
                    definition.Name);
                if (observed != null &&
                    MatchesDesired(observed, definition) &&
                    observedWorkbook.SelectedPivot.ConnectionRefreshed)
                {
                    try
                    {
                        DemandExactInventoryTransition(
                            snapshot,
                            observedWorkbook,
                            before.Name,
                            before,
                            observed);
                        DemandExactMeasureDependencies(
                            target,
                            definition.DirectMeasureDependencies);
                        return observed;
                    }
                    catch (InvalidOperationException)
                    {
                        // The intended replacement may be exact while an unrelated
                        // workbook artifact changed. Roll it back below and require
                        // full original-inventory verification.
                    }
                }

                try
                {
                    if (observed != null &&
                        IsRecognizedDesiredIntermediate(observed, definition))
                    {
                        DeleteRecognizedDesiredArtifact(target, observed, definition);
                    }
                    else if (observed != null &&
                             !SameSnapshot(observed, before) &&
                             !IsRecognizedSnapshotIntermediate(observed, before))
                    {
                        throw new PivotNamedSetRecoveryRequiredException(
                            "The replacement left an unrecognized collision at the generated identity.");
                    }

                    LivePivotNamedSetSnapshot restored = RestoreSnapshotCore(target, before);
                    DemandSameSnapshot(restored, before, "restored named set");
                    PivotNamedSetWorkbookSnapshot restoredWorkbook = CaptureStatic(target);
                    DemandExactInventoryTransition(
                        snapshot,
                        restoredWorkbook,
                        before.Name,
                        before,
                        restored);
                }
                catch (Exception rollbackFailure)
                {
                    throw new PivotNamedSetRecoveryRequiredException(
                        "The named-set replacement left an ambiguous host state and exact rollback failed.",
                        rollbackFailure);
                }

                throw new InvalidOperationException(
                    "The named-set replacement failed and the exact prior definition was restored.",
                    failure);
            }
        }

        public LivePivotNamedSetSnapshot RestoreSet(
            BoundPivotNamedSetTarget target,
            LivePivotNamedSetSnapshot before)
        {
            if (target == null) throw new ArgumentNullException(nameof(target));
            if (before == null) throw new ArgumentNullException(nameof(before));
            DemandStillBound(target);
            DemandRestorableSnapshot(before);
            DemandLiveSourceFingerprint(target, before.SourceFingerprint);
            PivotNamedSetWorkbookSnapshot baseline = Capture(target);
            DemandSourceFingerprint(baseline, before.SourceFingerprint);
            LivePivotNamedSetSnapshot? current = FindSelectedArtifact(
                baseline,
                before.Name);
            if (current != null &&
                !SameSnapshot(current, before) &&
                !IsRecognizedSnapshotIntermediate(current, before))
            {
                throw new PivotNamedSetRecoveryRequiredException(
                    "The current named-set state is not an exact restorable intermediate.");
            }

            LivePivotNamedSetSnapshot restored = RestoreSnapshotCore(target, before);
            PivotNamedSetWorkbookSnapshot completed = CaptureStatic(target);
            DemandExactInventoryTransition(
                baseline,
                completed,
                before.Name,
                current,
                restored);
            return restored;
        }

        public void DeleteSet(
            BoundPivotNamedSetTarget target,
            LivePivotNamedSetSnapshot expected)
        {
            if (target == null) throw new ArgumentNullException(nameof(target));
            if (expected == null) throw new ArgumentNullException(nameof(expected));
            DemandStillBound(target);
            DemandRestorableSnapshot(expected);
            DemandLiveSourceFingerprint(target, expected.SourceFingerprint);
            PivotNamedSetWorkbookSnapshot snapshot = Capture(target);
            DemandSourceFingerprint(snapshot, expected.SourceFingerprint);
            LivePivotNamedSetSnapshot current = DemandSelectedArtifact(
                snapshot,
                expected.Name);
            DemandSameSnapshot(current, expected, "named set");
            DemandSafeDestructiveUse(snapshot, expected);

            Exception? firstFailure = null;
            for (var attempt = 0; attempt < 2; attempt++)
            {
                try
                {
                    DeleteSetCore(target, expected);
                    PivotNamedSetWorkbookSnapshot deleted = CaptureStatic(target);
                    DemandExactInventoryTransition(
                        snapshot,
                        deleted,
                        expected.Name,
                        expected,
                        expectedAfter: null);
                    return;
                }
                catch (Exception failure)
                {
                    firstFailure = firstFailure ?? failure;
                    PivotNamedSetWorkbookSnapshot observedWorkbook;
                    try
                    {
                        observedWorkbook = CaptureForRecovery(target);
                    }
                    catch (Exception captureFailure)
                    {
                        throw new PivotNamedSetRecoveryRequiredException(
                            "The failed named-set deletion could not be inventoried.",
                            captureFailure);
                    }

                    Exception? inventoryFailure = null;
                    try
                    {
                        DemandUnrelatedInventoryUnchanged(
                            snapshot,
                            observedWorkbook,
                            expected.Name);
                    }
                    catch (Exception changed)
                    {
                        inventoryFailure = changed;
                    }

                    LivePivotNamedSetSnapshot? observed = FindSelectedArtifact(
                        observedWorkbook,
                        expected.Name);
                    if (observed == null && inventoryFailure == null)
                    {
                        DemandExactInventoryTransition(
                            snapshot,
                            observedWorkbook,
                            expected.Name,
                            expected,
                            expectedAfter: null);
                        return;
                    }

                    if (observed != null &&
                        SameSnapshot(observed, expected) &&
                        inventoryFailure == null &&
                        attempt == 0)
                    {
                        continue;
                    }

                    if (observed != null &&
                        SameSnapshot(observed, expected) &&
                        inventoryFailure == null)
                    {
                        throw new InvalidOperationException(
                            "Excel rejected the named-set deletion before changing the host.",
                            firstFailure);
                    }

                    if (observed != null &&
                        !IsRecognizedSnapshotIntermediate(observed, expected))
                    {
                        throw new PivotNamedSetRecoveryRequiredException(
                            "The deletion left an unrecognized named-set intermediate.",
                            firstFailure);
                    }

                    try
                    {
                        LivePivotNamedSetSnapshot restored = RestoreSnapshotCore(
                            target,
                            expected);
                        DemandSameSnapshot(restored, expected, "restored named set");
                        PivotNamedSetWorkbookSnapshot restoredWorkbook = CaptureStatic(target);
                        DemandExactInventoryTransition(
                            snapshot,
                            restoredWorkbook,
                            expected.Name,
                            expected,
                            restored);
                    }
                    catch (Exception rollbackFailure)
                    {
                        throw new PivotNamedSetRecoveryRequiredException(
                            "The named-set deletion left an ambiguous host state and exact rollback failed.",
                            inventoryFailure ?? rollbackFailure);
                    }

                    throw new InvalidOperationException(
                        "The named-set deletion failed and the exact prior definition was restored.",
                        firstFailure);
                }
            }
        }

        private static LivePivotNamedSetSnapshot CreateDesiredWithReconciliation(
            BoundPivotNamedSetTarget target,
            DesiredPivotNamedSet definition,
            PivotNamedSetWorkbookSnapshot baseline,
            LivePivotNamedSetSnapshot? baselineArtifact)
        {
            Exception? firstFailure = null;
            for (var attempt = 0; attempt < 2; attempt++)
            {
                try
                {
                    PivotNamedSetWorkbookSnapshot current = attempt == 0
                        ? baseline
                        : CaptureStatic(target);
                    DemandUnrelatedInventoryUnchanged(
                        baseline,
                        current,
                        definition.Name);
                    LivePivotNamedSetSnapshot? artifact = FindSelectedArtifact(
                        current,
                        definition.Name);
                    if (artifact == null)
                    {
                        _ = CreateDesiredCore(target, definition);
                    }
                    else if (MatchesDesired(artifact, definition) &&
                             current.SelectedPivot.ConnectionRefreshed)
                    {
                        DemandExactInventoryTransition(
                            baseline,
                            current,
                            definition.Name,
                            baselineArtifact,
                            artifact);
                        DemandExactMeasureDependencies(
                            target,
                            definition.DirectMeasureDependencies);
                        return artifact;
                    }
                    else if (IsRecognizedDesiredIntermediate(artifact, definition))
                    {
                        ConvergeDesiredIntermediate(target, artifact, definition);
                    }
                    else
                    {
                        throw new PivotNamedSetRecoveryRequiredException(
                            "An unrecognized object occupies the generated named-set identity.");
                    }

                    PivotNamedSetWorkbookSnapshot completed = CaptureStatic(target);
                    DemandUnrelatedInventoryUnchanged(
                        baseline,
                        completed,
                        definition.Name);
                    LivePivotNamedSetSnapshot desired = DemandSelectedArtifact(
                        completed,
                        definition.Name);
                    if (!MatchesDesired(desired, definition))
                    {
                        throw new InvalidOperationException(
                            "Excel did not preserve the exact compiled named-set definition.");
                    }

                    DemandExactInventoryTransition(
                        baseline,
                        completed,
                        definition.Name,
                        baselineArtifact,
                        desired);
                    DemandExactMeasureDependencies(
                        target,
                        definition.DirectMeasureDependencies);
                    return desired;
                }
                catch (Exception failure)
                {
                    firstFailure = firstFailure ?? failure;
                    PivotNamedSetWorkbookSnapshot observed;
                    try
                    {
                        observed = CaptureForRecovery(target);
                    }
                    catch (Exception captureFailure)
                    {
                        throw new PivotNamedSetRecoveryRequiredException(
                            "Excel failed during named-set creation and the resulting host state could not be read.",
                            captureFailure);
                    }

                    Exception? inventoryFailure = null;
                    try
                    {
                        DemandUnrelatedInventoryUnchanged(
                            baseline,
                            observed,
                            definition.Name);
                    }
                    catch (Exception changed)
                    {
                        inventoryFailure = changed;
                    }

                    LivePivotNamedSetSnapshot? observedArtifact = FindSelectedArtifact(
                        observed,
                        definition.Name);
                    if (inventoryFailure == null &&
                        observedArtifact != null &&
                        MatchesDesired(observedArtifact, definition) &&
                        observed.SelectedPivot.ConnectionRefreshed)
                    {
                        DemandExactInventoryTransition(
                            baseline,
                            observed,
                            definition.Name,
                            baselineArtifact,
                            observedArtifact);
                        DemandExactMeasureDependencies(
                            target,
                            definition.DirectMeasureDependencies);
                        return observedArtifact;
                    }

                    if (inventoryFailure == null &&
                        observedArtifact != null &&
                        IsRecognizedDesiredIntermediate(observedArtifact, definition))
                    {
                        try
                        {
                            ConvergeDesiredIntermediate(
                                target,
                                observedArtifact,
                                definition);
                            PivotNamedSetWorkbookSnapshot converged = CaptureStatic(target);
                            DemandUnrelatedInventoryUnchanged(
                                baseline,
                                converged,
                                definition.Name);
                            LivePivotNamedSetSnapshot desired = DemandSelectedArtifact(
                                converged,
                                definition.Name);
                            if (MatchesDesired(desired, definition))
                            {
                                DemandExactInventoryTransition(
                                    baseline,
                                    converged,
                                    definition.Name,
                                    baselineArtifact,
                                    desired);
                                DemandExactMeasureDependencies(
                                    target,
                                    definition.DirectMeasureDependencies);
                                return desired;
                            }
                        }
                        catch (Exception convergenceFailure)
                        {
                            firstFailure = firstFailure ?? convergenceFailure;
                        }
                    }

                    PivotNamedSetWorkbookSnapshot cleanupState;
                    try
                    {
                        cleanupState = CaptureForRecovery(target);
                    }
                    catch (Exception captureFailure)
                    {
                        throw new PivotNamedSetRecoveryRequiredException(
                            "Excel left a named-set state that could not be inspected for rollback.",
                            captureFailure);
                    }

                    LivePivotNamedSetSnapshot? cleanupArtifact = FindSelectedArtifact(
                        cleanupState,
                        definition.Name);
                    if (cleanupArtifact != null)
                    {
                        if (!IsRecognizedDesiredIntermediate(cleanupArtifact, definition))
                        {
                            throw new PivotNamedSetRecoveryRequiredException(
                                "Excel left an unrecognized named-set object after creation.",
                                firstFailure);
                        }

                        try
                        {
                            DeleteRecognizedDesiredArtifact(
                                target,
                                cleanupArtifact,
                                definition);
                        }
                        catch (Exception cleanupFailure)
                        {
                            throw new PivotNamedSetRecoveryRequiredException(
                                "Exact rollback of a partial named-set creation failed.",
                                cleanupFailure);
                        }
                    }

                    PivotNamedSetWorkbookSnapshot cleaned;
                    try
                    {
                        cleaned = CaptureStatic(target);
                        DemandExactInventoryTransition(
                            baseline,
                            cleaned,
                            definition.Name,
                            baselineArtifact,
                            expectedAfter: null);
                    }
                    catch (Exception verificationFailure)
                    {
                        throw new PivotNamedSetRecoveryRequiredException(
                            "Named-set rollback could not restore the exact workbook inventory.",
                            inventoryFailure ?? verificationFailure);
                    }

                    if (attempt == 0) continue;
                    throw new InvalidOperationException(
                        "Excel rejected named-set creation and exact absence was restored.",
                        firstFailure);
                }
            }

            throw new InvalidOperationException("The bounded named-set create retry was lost.");
        }

        private static LivePivotNamedSetSnapshot CreateDesiredCore(
            BoundPivotNamedSetTarget target,
            DesiredPivotNamedSet definition)
        {
            AddHostSet(
                target,
                definition.Name,
                definition.RawMdx,
                definition.Dynamic,
                definition.DisplayFolderMarker,
                definition.HierarchizeDistinct,
                definition.Caption,
                definition.FlattenHierarchies);
            LivePivotNamedSetSnapshot created = DemandSelectedArtifact(
                CaptureStatic(target),
                definition.Name);
            if (!MatchesDesired(created, definition))
            {
                throw new InvalidOperationException(
                    "Excel did not preserve the exact compiled named-set definition.");
            }

            return created;
        }

        private static void AddHostSet(
            BoundPivotNamedSetTarget target,
            string name,
            string rawFormula,
            bool dynamic,
            string displayFolder,
            bool hierarchizeDistinct,
            string caption,
            bool flattenHierarchies)
        {
            dynamic pivot = target.PivotTable;
            object calculatedMembersObject = ReadRequiredCollectionMember(
                () => (object?)pivot.CalculatedMembers,
                () => (object?)pivot.CalculatedMembers(),
                "Excel did not expose PivotTable CalculatedMembers for creation.");
            dynamic calculatedMembers = calculatedMembersObject;
            _ = calculatedMembers.Add(
                name,
                PivotNamedSetFormulaTransport.EncodeForExcel(rawFormula),
                Type.Missing,
                CalculatedMemberTypeSet,
                dynamic,
                displayFolder,
                hierarchizeDistinct);

            object cubeFieldsObject = ReadRequiredCollectionMember(
                () => (object?)pivot.CubeFields,
                () => (object?)pivot.CubeFields(),
                "Excel did not expose PivotTable CubeFields for AddSet.");
            dynamic cubeFields = cubeFieldsObject;
            _ = cubeFields.AddSet(name, caption);

            NativeSetPair pair = FindNativePair(target, name);
            if (pair.CalculatedMember == null || pair.CubeField == null)
            {
                throw new PivotNamedSetRecoveryRequiredException(
                    "Excel did not expose the complete calculated-set/CubeField pair after AddSet.");
            }

            dynamic calculatedMember = pair.CalculatedMember.Native;
            calculatedMember.FlattenHierarchies = flattenHierarchies;
            dynamic cubeField = pair.CubeField.Native;
            cubeField.FlattenHierarchies = flattenHierarchies;
            cubeField.HierarchizeDistinct = hierarchizeDistinct;

            object cacheObject = ReadPivotCache(target.PivotTable);
            dynamic cache = cacheObject;
            cache.MakeConnection();
            if (!ReadRequiredBoolean(
                    () => (object?)calculatedMember.IsValid,
                    "created calculated-set IsValid"))
            {
                throw new InvalidOperationException(
                    "Excel reported the created named set as invalid.");
            }

            string formulaReadback = ReadBoundedRequiredString(
                () => (object?)calculatedMember.Formula,
                MaximumFormulaCharacters,
                "created calculated-set formula");
            PivotNamedSetFormulaTransport.DemandExactReadback(
                formulaReadback,
                rawFormula);
        }

        private static LivePivotNamedSetSnapshot RestoreSnapshotCore(
            BoundPivotNamedSetTarget target,
            LivePivotNamedSetSnapshot before)
        {
            DemandRestorableSnapshot(before);
            PivotNamedSetWorkbookSnapshot workbook = CaptureStatic(target);
            foreach (LivePivotNamedSetSnapshot sibling in workbook.Artifacts.Where(artifact =>
                         string.Equals(
                             artifact.Name,
                             before.Name,
                             StringComparison.OrdinalIgnoreCase) &&
                         !artifact.IsSelectedTarget))
            {
                _ = sibling;
                throw new PivotNamedSetRecoveryRequiredException(
                    "A sibling PivotTable now exposes the named-set identity being restored.");
            }

            LivePivotNamedSetSnapshot? current = workbook.SelectedPivot.Artifacts
                .SingleOrDefault(artifact => string.Equals(
                    artifact.Name,
                    before.Name,
                    StringComparison.OrdinalIgnoreCase));
            if (current != null && SameSnapshot(current, before)) return current;
            if (current == null)
            {
                AddHostSetFromSnapshot(target, before);
            }
            else if (current.PairState == PivotNamedSetPairState.CalculatedMemberOnly &&
                     SameCalculatedSide(current, before))
            {
                AddCubeFieldForExistingCalculatedMember(target, before);
            }
            else if (current.PairState == PivotNamedSetPairState.CubeFieldOnly &&
                     IsRecognizedSnapshotIntermediate(current, before))
            {
                NativeSetPair pair = FindNativePair(target, before.Name);
                if (pair.CubeField == null)
                {
                    throw new PivotNamedSetRecoveryRequiredException(
                        "The orphan CubeField disappeared during rollback.");
                }

                dynamic cubeField = pair.CubeField.Native;
                cubeField.Delete();
                AddHostSetFromSnapshot(target, before);
            }
            else
            {
                throw new PivotNamedSetRecoveryRequiredException(
                    "The current named-set state cannot be reconciled with the exact rollback snapshot.");
            }

            LivePivotNamedSetSnapshot restored = DemandSelectedArtifact(
                CaptureStatic(target),
                before.Name);
            DemandSameSnapshot(restored, before, "restored named set");
            return restored;
        }

        private static void AddHostSetFromSnapshot(
            BoundPivotNamedSetTarget target,
            LivePivotNamedSetSnapshot before)
        {
            AddHostSet(
                target,
                before.Name,
                before.RawFormula,
                before.Dynamic!.Value,
                before.DisplayFolder,
                before.CalculatedMemberHierarchizeDistinct!.Value,
                before.Caption,
                before.CalculatedMemberFlattenHierarchies!.Value);
        }

        private static void AddCubeFieldForExistingCalculatedMember(
            BoundPivotNamedSetTarget target,
            LivePivotNamedSetSnapshot before)
        {
            dynamic pivot = target.PivotTable;
            object cubeFieldsObject = ReadRequiredCollectionMember(
                () => (object?)pivot.CubeFields,
                () => (object?)pivot.CubeFields(),
                "Excel did not expose PivotTable CubeFields for rollback AddSet.");
            dynamic cubeFields = cubeFieldsObject;
            _ = cubeFields.AddSet(before.Name, before.Caption);
            NativeSetPair pair = FindNativePair(target, before.Name);
            if (pair.CalculatedMember == null || pair.CubeField == null)
            {
                throw new PivotNamedSetRecoveryRequiredException(
                    "Excel did not complete the named-set pair during rollback.");
            }

            dynamic calculatedMember = pair.CalculatedMember.Native;
            calculatedMember.FlattenHierarchies =
                before.CalculatedMemberFlattenHierarchies!.Value;
            dynamic cubeField = pair.CubeField.Native;
            cubeField.FlattenHierarchies = before.CubeFieldFlattenHierarchies!.Value;
            cubeField.HierarchizeDistinct = before.CubeFieldHierarchizeDistinct!.Value;
            object cacheObject = ReadPivotCache(target.PivotTable);
            dynamic cache = cacheObject;
            cache.MakeConnection();
            if (!ReadRequiredBoolean(
                    () => (object?)calculatedMember.IsValid,
                    "restored calculated-set IsValid"))
            {
                throw new PivotNamedSetRecoveryRequiredException(
                    "Excel reported the restored named set as invalid.");
            }
        }

        private static void DeleteSetCore(
            BoundPivotNamedSetTarget target,
            LivePivotNamedSetSnapshot expected)
        {
            NativeSetPair pair = FindNativePair(target, expected.Name);
            if (pair.CalculatedMember == null || pair.CubeField == null)
            {
                throw new PivotNamedSetRecoveryRequiredException(
                    "The named-set pair became incomplete before deletion.");
            }

            if (!NativePairMatchesSnapshot(pair, expected))
            {
                throw new PivotNamedSetRecoveryRequiredException(
                    "The named-set pair changed after the exact pre-mutation capture.");
            }

            if (pair.CubeField.Orientation != 0)
            {
                throw new InvalidOperationException(
                    "Named-set deletion requires a hidden CubeField until axis placement is implemented.");
            }

            dynamic cubeField = pair.CubeField.Native;
            cubeField.Delete();
            dynamic calculatedMember = pair.CalculatedMember.Native;
            calculatedMember.Delete();
        }

        private static bool NativePairMatchesSnapshot(
            NativeSetPair pair,
            LivePivotNamedSetSnapshot expected)
        {
            CalculatedMemberHandle calculated = pair.CalculatedMember!;
            CubeSetFieldHandle cube = pair.CubeField!;
            return string.Equals(calculated.Name, expected.Name, StringComparison.Ordinal) &&
                   calculated.Type == expected.CalculatedMemberType &&
                   string.Equals(
                       calculated.RawFormula,
                       expected.RawFormula,
                       StringComparison.Ordinal) &&
                   string.Equals(
                       calculated.DisplayFolder,
                       expected.DisplayFolder,
                       StringComparison.Ordinal) &&
                   calculated.Dynamic == expected.Dynamic &&
                   calculated.FlattenHierarchies ==
                       expected.CalculatedMemberFlattenHierarchies &&
                   calculated.HierarchizeDistinct ==
                       expected.CalculatedMemberHierarchizeDistinct &&
                   calculated.IsValid == expected.IsValid &&
                   string.Equals(cube.SourceName, expected.SourceName, StringComparison.Ordinal) &&
                   string.Equals(cube.Caption, expected.Caption, StringComparison.Ordinal) &&
                   cube.Type == expected.CubeFieldType &&
                   cube.FlattenHierarchies == expected.CubeFieldFlattenHierarchies &&
                   cube.HierarchizeDistinct == expected.CubeFieldHierarchizeDistinct &&
                   cube.ShowInFieldList == expected.ShowInFieldList &&
                   cube.Orientation == expected.Orientation;
        }

        private static NativeSetPair FindNativePair(
            BoundPivotNamedSetTarget target,
            string name,
            bool refreshAndReadValidity = true)
        {
            if (refreshAndReadValidity) DemandMakeConnection(target.PivotTable);
            dynamic pivot = target.PivotTable;
            object calculatedMembersObject = ReadRequiredCollectionMember(
                () => (object?)pivot.CalculatedMembers,
                () => (object?)pivot.CalculatedMembers(),
                "Excel did not expose PivotTable CalculatedMembers.");
            var calculatedMatches = new List<CalculatedMemberHandle>();
            foreach (object item in ReadCollection(
                         calculatedMembersObject,
                         MaximumCalculatedMembers,
                         "PivotTable CalculatedMembers"))
            {
                dynamic calculatedMember = item;
                int type = ReadRequiredInt(
                    () => (object?)calculatedMember.Type,
                    "calculated-member type");
                if (type != CalculatedMemberTypeSet) continue;
                string itemName = ReadBoundedRequiredString(
                    () => (object?)calculatedMember.Name,
                    MaximumNameCharacters,
                    "calculated-set name");
                if (!string.Equals(itemName, name, StringComparison.OrdinalIgnoreCase)) continue;
                string readback = ReadBoundedRequiredString(
                    () => (object?)calculatedMember.Formula,
                    MaximumFormulaCharacters,
                    "calculated-set formula");
                calculatedMatches.Add(new CalculatedMemberHandle(
                    item,
                    itemName,
                    type,
                    PivotNamedSetFormulaTransport.DecodeRequired(readback),
                    ReadBoundedOptionalString(
                        () => (object?)calculatedMember.DisplayFolder,
                        MaximumNameCharacters,
                        "calculated-set DisplayFolder"),
                    ReadRequiredBoolean(
                        () => (object?)calculatedMember.Dynamic,
                        "calculated-set Dynamic"),
                    ReadRequiredBoolean(
                        () => (object?)calculatedMember.FlattenHierarchies,
                        "calculated-set FlattenHierarchies"),
                    ReadRequiredBoolean(
                        () => (object?)calculatedMember.HierarchizeDistinct,
                        "calculated-set HierarchizeDistinct"),
                    refreshAndReadValidity
                        ? ReadOptionalBoolean(
                            () => (object?)calculatedMember.IsValid,
                            "calculated-set IsValid")
                        : null));
            }

            object cubeFieldsObject = ReadRequiredCollectionMember(
                () => (object?)pivot.CubeFields,
                () => (object?)pivot.CubeFields(),
                "Excel did not expose PivotTable CubeFields.");
            var cubeMatches = new List<CubeSetFieldHandle>();
            foreach (object item in ReadCollection(
                         cubeFieldsObject,
                         MaximumCubeFields,
                         "PivotTable CubeFields"))
            {
                dynamic cubeField = item;
                int type = ReadCubeFieldType(cubeField);
                if (type != CubeFieldTypeSet) continue;
                string sourceName = ReadProviderUniqueName(
                    cubeField,
                    "named-set CubeField source name");
                if (!string.Equals(sourceName, name, StringComparison.OrdinalIgnoreCase)) continue;
                cubeMatches.Add(new CubeSetFieldHandle(
                    item,
                    sourceName,
                    ReadBoundedOptionalString(
                        () => (object?)cubeField.Caption,
                        MaximumCaptionCharacters,
                        "named-set CubeField caption"),
                    type,
                    ReadRequiredBoolean(
                        () => (object?)cubeField.FlattenHierarchies,
                        "named-set CubeField FlattenHierarchies"),
                    ReadRequiredBoolean(
                        () => (object?)cubeField.HierarchizeDistinct,
                        "named-set CubeField HierarchizeDistinct"),
                    ReadRequiredBoolean(
                        () => (object?)cubeField.ShowInFieldList,
                        "named-set CubeField ShowInFieldList"),
                    ReadRequiredInt(
                        () => (object?)cubeField.Orientation,
                        "named-set CubeField orientation")));
            }

            if (calculatedMatches.Count > 1 || cubeMatches.Count > 1)
            {
                throw new InvalidOperationException(
                    "Excel exposed duplicate objects for one named-set identity.");
            }

            return new NativeSetPair(
                calculatedMatches.SingleOrDefault(),
                cubeMatches.SingleOrDefault());
        }

        private static LivePivotNamedSetSnapshot DemandSelectedArtifact(
            PivotNamedSetWorkbookSnapshot snapshot,
            string name)
        {
            LivePivotNamedSetSnapshot? artifact = snapshot.SelectedPivot.Artifacts
                .SingleOrDefault(item => string.Equals(
                    item.Name,
                    name,
                    StringComparison.OrdinalIgnoreCase));
            if (artifact == null)
            {
                throw new InvalidOperationException(
                    "Excel did not expose the requested named set in the selected PivotTable.");
            }

            return artifact;
        }

        private static PivotNamedSetWorkbookSnapshot CaptureStatic(
            BoundPivotNamedSetTarget target)
        {
            return new LateBoundPivotNamedSetGateway().Capture(target);
        }

        private static PivotNamedSetWorkbookSnapshot CaptureForRecovery(
            BoundPivotNamedSetTarget target)
        {
            return new LateBoundPivotNamedSetGateway().CaptureCore(
                target,
                requireConnectionRefresh: false);
        }

        private static void DemandSafeCreateUse(
            PivotNamedSetWorkbookSnapshot snapshot,
            DesiredPivotNamedSet desired,
            LivePivotNamedSetSnapshot? existing)
        {
            if (snapshot.Artifacts.Any(artifact =>
                    !artifact.IsSelectedTarget &&
                    string.Equals(
                        artifact.Name,
                        desired.Name,
                        StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidOperationException(
                    "Another Data Model PivotTable exposes the generated named-set identity.");
            }

            if (existing != null && !IsRecognizedDesiredIntermediate(existing, desired))
            {
                throw new InvalidOperationException(
                    "An unowned or stale named set occupies the generated identity.");
            }

            foreach (PivotCalculatedMemberReferenceSnapshot member in snapshot.Pivots
                         .SelectMany(pivot => pivot.CalculatedMembers))
            {
                if (!string.Equals(
                        member.Name,
                        desired.Name,
                        StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                bool exactSelectedSide = existing != null &&
                                         string.Equals(
                                             member.WorksheetName,
                                             existing.WorksheetName,
                                             StringComparison.OrdinalIgnoreCase) &&
                                         string.Equals(
                                             member.PivotTableName,
                                             existing.PivotTableName,
                                             StringComparison.OrdinalIgnoreCase) &&
                                         string.Equals(
                                             member.Name,
                                             existing.Name,
                                             StringComparison.OrdinalIgnoreCase);
                if (!exactSelectedSide)
                {
                    throw new InvalidOperationException(
                        "A calculated member already uses the generated named-set identity.");
                }
            }

            DemandNoFormulaReferences(snapshot, desired.Name, existing);
        }

        private static bool IsRecognizedDesiredIntermediate(
            LivePivotNamedSetSnapshot live,
            DesiredPivotNamedSet desired)
        {
            if (!live.IsSelectedTarget ||
                !string.Equals(live.Name, desired.Name, StringComparison.Ordinal))
            {
                return false;
            }

            bool hasCalculated = live.PairState != PivotNamedSetPairState.CubeFieldOnly;
            bool hasCube = live.PairState != PivotNamedSetPairState.CalculatedMemberOnly;
            bool calculatedMatches = !hasCalculated ||
                                     (live.CalculatedMemberType == CalculatedMemberTypeSet &&
                                      string.Equals(
                                          live.RawFormula,
                                          desired.RawMdx,
                                          StringComparison.Ordinal) &&
                                      string.Equals(
                                          live.FormulaFingerprint,
                                          desired.FormulaFingerprint,
                                          StringComparison.Ordinal) &&
                                      string.Equals(
                                          live.DisplayFolder,
                                          desired.DisplayFolderMarker,
                                          StringComparison.Ordinal) &&
                                      live.Dynamic == desired.Dynamic &&
                                      IsDefaultOrDesired(
                                          live.CalculatedMemberFlattenHierarchies,
                                          desired.FlattenHierarchies) &&
                                      live.CalculatedMemberHierarchizeDistinct ==
                                          desired.HierarchizeDistinct);
            bool cubeMatches = !hasCube ||
                               (live.CubeFieldType == CubeFieldTypeSet &&
                                string.Equals(
                                    live.SourceName,
                                    desired.Name,
                                    StringComparison.Ordinal) &&
                                string.Equals(
                                    live.Caption,
                                    desired.Caption,
                                    StringComparison.Ordinal) &&
                                IsDefaultOrDesired(
                                    live.CubeFieldFlattenHierarchies,
                                    desired.FlattenHierarchies) &&
                                live.CubeFieldHierarchizeDistinct ==
                                    desired.HierarchizeDistinct &&
                                live.ShowInFieldList == true &&
                                live.Orientation == 0);
            return calculatedMatches && cubeMatches;
        }

        private static bool IsDefaultOrDesired(bool? live, bool desired)
        {
            return live.HasValue && (!live.Value || live.Value == desired);
        }

        private static void ConvergeDesiredIntermediate(
            BoundPivotNamedSetTarget target,
            LivePivotNamedSetSnapshot observed,
            DesiredPivotNamedSet desired)
        {
            if (!IsRecognizedDesiredIntermediate(observed, desired))
            {
                throw new PivotNamedSetRecoveryRequiredException(
                    "The partial named-set state is not an exact intended intermediate.");
            }

            if (observed.PairState == PivotNamedSetPairState.CubeFieldOnly)
            {
                throw new InvalidOperationException(
                    "A CubeField-only intermediate must be rolled back before retry.");
            }

            if (observed.PairState == PivotNamedSetPairState.CalculatedMemberOnly)
            {
                dynamic pivot = target.PivotTable;
                object cubeFieldsObject = ReadRequiredCollectionMember(
                    () => (object?)pivot.CubeFields,
                    () => (object?)pivot.CubeFields(),
                    "Excel did not expose PivotTable CubeFields for AddSet reconciliation.");
                dynamic cubeFields = cubeFieldsObject;
                _ = cubeFields.AddSet(desired.Name, desired.Caption);
            }

            NormalizeDesiredPair(target, desired);
        }

        private static void NormalizeDesiredPair(
            BoundPivotNamedSetTarget target,
            DesiredPivotNamedSet desired)
        {
            NativeSetPair pair = FindNativePair(
                target,
                desired.Name,
                refreshAndReadValidity: false);
            if (pair.CalculatedMember == null || pair.CubeField == null ||
                !string.Equals(
                    pair.CalculatedMember.Name,
                    desired.Name,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    pair.CalculatedMember.RawFormula,
                    desired.RawMdx,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    pair.CalculatedMember.DisplayFolder,
                    desired.DisplayFolderMarker,
                    StringComparison.Ordinal) ||
                pair.CalculatedMember.Type != CalculatedMemberTypeSet ||
                pair.CalculatedMember.Dynamic != desired.Dynamic ||
                pair.CalculatedMember.HierarchizeDistinct !=
                    desired.HierarchizeDistinct ||
                !IsDefaultOrDesired(
                    pair.CalculatedMember.FlattenHierarchies,
                    desired.FlattenHierarchies) ||
                !string.Equals(
                    pair.CubeField.SourceName,
                    desired.Name,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    pair.CubeField.Caption,
                    desired.Caption,
                    StringComparison.Ordinal) ||
                pair.CubeField.Type != CubeFieldTypeSet ||
                !IsDefaultOrDesired(
                    pair.CubeField.FlattenHierarchies,
                    desired.FlattenHierarchies) ||
                pair.CubeField.HierarchizeDistinct != desired.HierarchizeDistinct ||
                !pair.CubeField.ShowInFieldList ||
                pair.CubeField.Orientation != 0)
            {
                throw new PivotNamedSetRecoveryRequiredException(
                    "The host pair changed before named-set reconciliation.");
            }

            dynamic calculatedMember = pair.CalculatedMember.Native;
            calculatedMember.FlattenHierarchies = desired.FlattenHierarchies;
            dynamic cubeField = pair.CubeField.Native;
            cubeField.FlattenHierarchies = desired.FlattenHierarchies;
            cubeField.HierarchizeDistinct = desired.HierarchizeDistinct;
            DemandMakeConnection(target.PivotTable);
            if (!ReadRequiredBoolean(
                    () => (object?)calculatedMember.IsValid,
                    "reconciled calculated-set IsValid"))
            {
                throw new InvalidOperationException(
                    "Excel reported the reconciled named set as invalid.");
            }

            string readback = ReadBoundedRequiredString(
                () => (object?)calculatedMember.Formula,
                MaximumFormulaCharacters,
                "reconciled calculated-set formula");
            PivotNamedSetFormulaTransport.DemandExactReadback(
                readback,
                desired.RawMdx);
        }

        private static void DeleteRecognizedDesiredArtifact(
            BoundPivotNamedSetTarget target,
            LivePivotNamedSetSnapshot observed,
            DesiredPivotNamedSet desired)
        {
            if (!IsRecognizedDesiredIntermediate(observed, desired))
            {
                throw new PivotNamedSetRecoveryRequiredException(
                    "Refusing to delete an unrecognized named-set intermediate.");
            }

            NativeSetPair pair = FindNativePair(
                target,
                desired.Name,
                refreshAndReadValidity: false);
            bool expectsCalculated =
                observed.PairState != PivotNamedSetPairState.CubeFieldOnly;
            bool expectsCube =
                observed.PairState != PivotNamedSetPairState.CalculatedMemberOnly;
            if ((pair.CalculatedMember != null) != expectsCalculated ||
                (pair.CubeField != null) != expectsCube ||
                (pair.CalculatedMember != null &&
                 (!string.Equals(
                      pair.CalculatedMember.RawFormula,
                      observed.RawFormula,
                      StringComparison.Ordinal) ||
                  !string.Equals(
                      pair.CalculatedMember.DisplayFolder,
                      observed.DisplayFolder,
                      StringComparison.Ordinal))) ||
                (pair.CubeField != null &&
                 (!string.Equals(
                      pair.CubeField.Caption,
                      observed.Caption,
                      StringComparison.Ordinal) ||
                  pair.CubeField.Orientation != observed.Orientation)))
            {
                throw new PivotNamedSetRecoveryRequiredException(
                    "The intended partial named set changed before exact rollback.");
            }

            if (pair.CubeField != null)
            {
                dynamic cubeField = pair.CubeField.Native;
                cubeField.Delete();
            }

            if (pair.CalculatedMember != null)
            {
                dynamic calculatedMember = pair.CalculatedMember.Native;
                calculatedMember.Delete();
            }
        }

        private static void DemandSafeDestructiveUse(
            PivotNamedSetWorkbookSnapshot snapshot,
            LivePivotNamedSetSnapshot expected)
        {
            DemandRestorableSnapshot(expected);
            if (expected.Orientation != 0)
            {
                throw new InvalidOperationException(
                    "Named-set replacement or deletion requires a hidden set until axis placement is implemented.");
            }

            if (snapshot.Artifacts.Any(artifact =>
                    !artifact.IsSelectedTarget &&
                    string.Equals(
                        artifact.Name,
                        expected.Name,
                        StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidOperationException(
                    "Another Data Model PivotTable exposes this named-set identity.");
            }

            DemandNoFormulaReferences(snapshot, expected.Name, expected);
        }

        private static void DemandNoFormulaReferences(
            PivotNamedSetWorkbookSnapshot snapshot,
            string name,
            LivePivotNamedSetSnapshot? excluded)
        {
            foreach (PivotCalculatedMemberReferenceSnapshot member in snapshot.Pivots
                         .SelectMany(pivot => pivot.CalculatedMembers))
            {
                bool isExcluded = excluded != null &&
                                  string.Equals(
                                      member.WorksheetName,
                                      excluded.WorksheetName,
                                      StringComparison.OrdinalIgnoreCase) &&
                                  string.Equals(
                                      member.PivotTableName,
                                      excluded.PivotTableName,
                                      StringComparison.OrdinalIgnoreCase) &&
                                  string.Equals(
                                      member.Name,
                                      excluded.Name,
                                      StringComparison.OrdinalIgnoreCase);
                if (isExcluded) continue;
                if (!member.FormulaScanComplete)
                {
                    throw new InvalidOperationException(
                        "A live calculated-member formula could not be scanned safely for named-set use.");
                }

                if (MdxNamedSetReferenceScanner.MightReference(member.RawFormula, name))
                {
                    throw new InvalidOperationException(
                        "A live calculated-member formula references the named set.");
                }
            }
        }

        private static void DemandDesiredDefinition(DesiredPivotNamedSet definition)
        {
            if (!IsProviderUniqueName(definition.Name) ||
                string.IsNullOrWhiteSpace(definition.SourceFingerprint) ||
                string.IsNullOrWhiteSpace(definition.CompilationFingerprint) ||
                string.IsNullOrWhiteSpace(definition.Caption) ||
                definition.Caption.Length > MaximumCaptionCharacters ||
                definition.Caption.Any(char.IsControl) ||
                string.IsNullOrWhiteSpace(definition.DisplayFolderMarker) ||
                definition.DisplayFolderMarker.Length > MaximumNameCharacters ||
                definition.DisplayFolderMarker.Any(char.IsControl) ||
                definition.HierarchizeDistinct ||
                !string.Equals(
                    PivotMdxFingerprint.ComputeFormula(definition.RawMdx),
                    definition.FormulaFingerprint,
                    StringComparison.Ordinal))
            {
                throw new ArgumentException(
                    "The trusted compiled named-set host definition is invalid.",
                    nameof(definition));
            }

            PivotPlusMetadataValidator.ValidateFingerprint(
                definition.SourceFingerprint,
                "named-set source fingerprint");
            PivotPlusMetadataValidator.ValidateFingerprint(
                definition.CompilationFingerprint,
                "named-set compilation fingerprint");
            var dependencyIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var dependencyNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (DesiredPivotNamedSetMeasureDependency dependency in
                     definition.DirectMeasureDependencies)
            {
                PivotPlusMetadataValidator.ValidateId(
                    dependency.DefinitionId,
                    "named-set measure dependency identifier");
                PivotPlusMetadataValidator.ValidateArtifactName(
                    dependency.GeneratedMeasureName);
                PivotPlusMetadataValidator.ValidateFingerprint(
                    dependency.MeasureDefinitionFingerprint,
                    "named-set dependency definition fingerprint");
                PivotPlusMetadataValidator.ValidateFingerprint(
                    dependency.MeasureFormulaFingerprint,
                    "named-set dependency formula fingerprint");
                if (!dependencyIds.Add(dependency.DefinitionId) ||
                    !dependencyNames.Add(dependency.GeneratedMeasureName) ||
                    string.IsNullOrWhiteSpace(dependency.ExpectedDescriptionMarker) ||
                    dependency.ExpectedDescriptionMarker.Any(char.IsControl))
                {
                    throw new ArgumentException(
                        "The trusted named-set dependency binding is invalid.",
                        nameof(definition));
                }
            }

            _ = PivotNamedSetFormulaTransport.EncodeForExcel(definition.RawMdx);
        }

        private PivotNamedSetSchemaDiscoveryResult DemandLiveCompatibility(
            BoundPivotNamedSetTarget target,
            DesiredPivotNamedSet definition)
        {
            PivotNamedSetSchemaDiscoveryResult discovery = DiscoverSchema(target);
            if (!string.Equals(
                    discovery.Schema.SourceFingerprint,
                    definition.SourceFingerprint,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The live Data Model schema no longer matches the Core-compiled named set.");
            }

            DemandExactMeasureDependencies(target, definition.DirectMeasureDependencies);
            return discovery;
        }

        private void DemandLiveSourceFingerprint(
            BoundPivotNamedSetTarget target,
            string expectedSourceFingerprint)
        {
            PivotNamedSetSchemaDiscoveryResult discovery = DiscoverSchema(target);
            if (!string.Equals(
                    discovery.Schema.SourceFingerprint,
                    expectedSourceFingerprint,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The live Data Model schema changed after the named-set snapshot.");
            }
        }

        private static void DemandSourceFingerprint(
            PivotNamedSetWorkbookSnapshot snapshot,
            string expectedSourceFingerprint)
        {
            if (!string.Equals(
                    snapshot.SourceFingerprint,
                    expectedSourceFingerprint,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The named-set host snapshot is bound to a different Data Model schema.");
            }
        }

        private static void DemandExactMeasureDependencies(
            BoundPivotNamedSetTarget target,
            IReadOnlyList<DesiredPivotNamedSetMeasureDependency> dependencies)
        {
            if (dependencies.Count == 0) return;
            dynamic model = target.Model;
            object collection = ReadRequired(
                () => (object?)model.ModelMeasures,
                "Excel did not expose Data Model measures required by the named set.");
            var live = new Dictionary<string, LiveMeasureDependency>(
                StringComparer.OrdinalIgnoreCase);
            foreach (object item in ReadCollection(
                         collection,
                         MaximumModelMeasures,
                         "Data Model measures"))
            {
                dynamic measure = item;
                string name = ReadBoundedRequiredString(
                    () => (object?)measure.Name,
                    MaximumNameCharacters,
                    "Data Model measure name");
                if (live.ContainsKey(name))
                {
                    throw new InvalidOperationException(
                        "Excel exposed duplicate Data Model measure names.");
                }

                string formula = ReadBoundedRequiredString(
                    () => (object?)measure.Formula,
                    MaximumFormulaCharacters,
                    "Data Model measure formula");
                string description = ReadBoundedOptionalString(
                    () => (object?)measure.Description,
                    MaximumCaptionCharacters,
                    "Data Model measure description");
                live.Add(name, new LiveMeasureDependency(
                    name,
                    PivotDaxFingerprint.ComputeFormula(formula),
                    description));
            }

            foreach (DesiredPivotNamedSetMeasureDependency dependency in dependencies)
            {
                if (!live.TryGetValue(
                        dependency.GeneratedMeasureName,
                        out LiveMeasureDependency? measure) ||
                    !string.Equals(
                        measure.Name,
                        dependency.GeneratedMeasureName,
                        StringComparison.Ordinal) ||
                    !string.Equals(
                        measure.FormulaFingerprint,
                        dependency.MeasureFormulaFingerprint,
                        StringComparison.Ordinal) ||
                    !string.Equals(
                        measure.Description,
                        dependency.ExpectedDescriptionMarker,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "A Core-bound Data Model measure dependency is missing or stale.");
                }
            }
        }

        private static void DemandRestorableSnapshot(LivePivotNamedSetSnapshot snapshot)
        {
            if (!snapshot.IsSelectedTarget ||
                snapshot.PairState != PivotNamedSetPairState.Complete ||
                !string.Equals(snapshot.Name, snapshot.SourceName, StringComparison.Ordinal) ||
                !IsProviderUniqueName(snapshot.Name) ||
                snapshot.CalculatedMemberType != CalculatedMemberTypeSet ||
                snapshot.CubeFieldType != CubeFieldTypeSet ||
                string.IsNullOrWhiteSpace(snapshot.RawFormula) ||
                !string.Equals(
                    PivotMdxFingerprint.ComputeFormula(snapshot.RawFormula),
                    snapshot.FormulaFingerprint,
                    StringComparison.Ordinal) ||
                string.IsNullOrWhiteSpace(snapshot.Caption) ||
                !IsOwnedDisplayFolderMarker(snapshot.DisplayFolder) ||
                snapshot.Dynamic == null ||
                snapshot.CalculatedMemberFlattenHierarchies == null ||
                snapshot.CubeFieldFlattenHierarchies == null ||
                snapshot.CalculatedMemberFlattenHierarchies !=
                    snapshot.CubeFieldFlattenHierarchies ||
                snapshot.CalculatedMemberHierarchizeDistinct == null ||
                snapshot.CubeFieldHierarchizeDistinct == null ||
                snapshot.CalculatedMemberHierarchizeDistinct !=
                    snapshot.CubeFieldHierarchizeDistinct ||
                snapshot.ShowInFieldList != true ||
                snapshot.Orientation != 0 ||
                snapshot.IsValid != true ||
                string.IsNullOrWhiteSpace(snapshot.SourceFingerprint) ||
                string.IsNullOrWhiteSpace(snapshot.ModelLineageFingerprint))
            {
                throw new InvalidOperationException(
                    "Only an exact hidden valid selected-PivotTable named-set snapshot can be restored.");
            }

            _ = PivotNamedSetFormulaTransport.EncodeForExcel(snapshot.RawFormula);
            PivotPlusMetadataValidator.ValidateFingerprint(
                snapshot.SourceFingerprint,
                "named-set snapshot source fingerprint");
            PivotPlusMetadataValidator.ValidateFingerprint(
                snapshot.ModelLineageFingerprint,
                "named-set snapshot model-lineage fingerprint");
            string canonical = PivotNamedSetCanonical.CreateLiveFingerprint(
                snapshot.SourceFingerprint,
                snapshot.ModelLineageFingerprint,
                snapshot.Name,
                snapshot.PairState,
                snapshot.FormulaFingerprint,
                snapshot.DisplayFolder,
                snapshot.SourceName,
                snapshot.Caption,
                snapshot.CalculatedMemberType,
                snapshot.CubeFieldType,
                snapshot.Dynamic,
                snapshot.CalculatedMemberFlattenHierarchies,
                snapshot.CubeFieldFlattenHierarchies,
                snapshot.CalculatedMemberHierarchizeDistinct,
                snapshot.CubeFieldHierarchizeDistinct,
                snapshot.ShowInFieldList,
                snapshot.Orientation,
                snapshot.IsValid);
            if (!string.Equals(canonical, snapshot.LiveFingerprint, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The named-set snapshot fingerprint does not match its exact pair state.");
            }
        }

        private static bool IsOwnedDisplayFolderMarker(string value)
        {
            const string prefix = "PivotTable+|set|namedset.semantic.v1:sha256:";
            if (value == null ||
                !value.StartsWith(prefix, StringComparison.Ordinal) ||
                value.Length != prefix.Length + 64)
            {
                return false;
            }

            return value.Substring(prefix.Length).All(character =>
                (character >= '0' && character <= '9') ||
                (character >= 'a' && character <= 'f'));
        }

        private static bool MatchesDesired(
            LivePivotNamedSetSnapshot live,
            DesiredPivotNamedSet desired)
        {
            return live.IsSelectedTarget &&
                   live.PairState == PivotNamedSetPairState.Complete &&
                   string.Equals(
                       live.SourceFingerprint,
                       desired.SourceFingerprint,
                       StringComparison.Ordinal) &&
                   string.Equals(live.Name, desired.Name, StringComparison.Ordinal) &&
                   string.Equals(live.SourceName, desired.Name, StringComparison.Ordinal) &&
                   string.Equals(live.Caption, desired.Caption, StringComparison.Ordinal) &&
                   string.Equals(live.RawFormula, desired.RawMdx, StringComparison.Ordinal) &&
                   string.Equals(
                       live.FormulaFingerprint,
                       desired.FormulaFingerprint,
                       StringComparison.Ordinal) &&
                   string.Equals(
                       live.DisplayFolder,
                       desired.DisplayFolderMarker,
                       StringComparison.Ordinal) &&
                   live.CalculatedMemberType == CalculatedMemberTypeSet &&
                   live.CubeFieldType == CubeFieldTypeSet &&
                   live.Dynamic == desired.Dynamic &&
                   live.CalculatedMemberFlattenHierarchies == desired.FlattenHierarchies &&
                   live.CubeFieldFlattenHierarchies == desired.FlattenHierarchies &&
                   live.CalculatedMemberHierarchizeDistinct == desired.HierarchizeDistinct &&
                   live.CubeFieldHierarchizeDistinct == desired.HierarchizeDistinct &&
                   live.ShowInFieldList == true &&
                   live.Orientation == 0 &&
                   live.IsValid == true;
        }

        private static bool SameCalculatedSide(
            LivePivotNamedSetSnapshot live,
            LivePivotNamedSetSnapshot expected)
        {
            return live.PairState == PivotNamedSetPairState.CalculatedMemberOnly &&
                   live.IsSelectedTarget == expected.IsSelectedTarget &&
                   string.Equals(
                       live.SourceFingerprint,
                       expected.SourceFingerprint,
                       StringComparison.Ordinal) &&
                   string.Equals(
                       live.ModelLineageFingerprint,
                       expected.ModelLineageFingerprint,
                       StringComparison.Ordinal) &&
                   string.Equals(live.Name, expected.Name, StringComparison.Ordinal) &&
                   string.Equals(live.RawFormula, expected.RawFormula, StringComparison.Ordinal) &&
                   string.Equals(
                       live.FormulaFingerprint,
                       expected.FormulaFingerprint,
                       StringComparison.Ordinal) &&
                   string.Equals(
                       live.DisplayFolder,
                       expected.DisplayFolder,
                       StringComparison.Ordinal) &&
                   live.CalculatedMemberType == expected.CalculatedMemberType &&
                   live.Dynamic == expected.Dynamic &&
                   live.CalculatedMemberFlattenHierarchies ==
                       expected.CalculatedMemberFlattenHierarchies &&
                    live.CalculatedMemberHierarchizeDistinct ==
                        expected.CalculatedMemberHierarchizeDistinct;
        }

        private static bool IsRecognizedSnapshotIntermediate(
            LivePivotNamedSetSnapshot live,
            LivePivotNamedSetSnapshot expected)
        {
            if (!live.IsSelectedTarget ||
                !string.Equals(
                    live.SourceFingerprint,
                    expected.SourceFingerprint,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    live.ModelLineageFingerprint,
                    expected.ModelLineageFingerprint,
                    StringComparison.Ordinal) ||
                !string.Equals(live.Name, expected.Name, StringComparison.Ordinal))
            {
                return false;
            }

            bool hasCalculated = live.PairState != PivotNamedSetPairState.CubeFieldOnly;
            bool hasCube = live.PairState != PivotNamedSetPairState.CalculatedMemberOnly;
            bool calculatedMatches = !hasCalculated ||
                                     (live.CalculatedMemberType ==
                                          expected.CalculatedMemberType &&
                                      string.Equals(
                                          live.RawFormula,
                                          expected.RawFormula,
                                          StringComparison.Ordinal) &&
                                      string.Equals(
                                          live.FormulaFingerprint,
                                          expected.FormulaFingerprint,
                                          StringComparison.Ordinal) &&
                                      string.Equals(
                                          live.DisplayFolder,
                                          expected.DisplayFolder,
                                          StringComparison.Ordinal) &&
                                      live.Dynamic == expected.Dynamic &&
                                      live.CalculatedMemberFlattenHierarchies ==
                                          expected.CalculatedMemberFlattenHierarchies &&
                                      live.CalculatedMemberHierarchizeDistinct ==
                                          expected.CalculatedMemberHierarchizeDistinct);
            bool cubeMatches = !hasCube ||
                               (live.CubeFieldType == expected.CubeFieldType &&
                                string.Equals(
                                    live.SourceName,
                                    expected.SourceName,
                                    StringComparison.Ordinal) &&
                                string.Equals(
                                    live.Caption,
                                    expected.Caption,
                                    StringComparison.Ordinal) &&
                                live.CubeFieldFlattenHierarchies ==
                                    expected.CubeFieldFlattenHierarchies &&
                                live.CubeFieldHierarchizeDistinct ==
                                    expected.CubeFieldHierarchizeDistinct &&
                                live.ShowInFieldList == expected.ShowInFieldList &&
                                live.Orientation == expected.Orientation);
            return calculatedMatches && cubeMatches;
        }

        private static void DemandSameSnapshot(
            LivePivotNamedSetSnapshot live,
            LivePivotNamedSetSnapshot expected,
            string label)
        {
            if (!SameSnapshot(live, expected))
            {
                throw new InvalidOperationException(
                    "The exact " + label + " changed after capture.");
            }
        }

        private static bool SameSnapshot(
            LivePivotNamedSetSnapshot left,
            LivePivotNamedSetSnapshot right)
        {
            return string.Equals(
                       left.LiveFingerprint,
                       right.LiveFingerprint,
                       StringComparison.Ordinal) &&
                    string.Equals(
                        left.WorksheetName,
                        right.WorksheetName,
                        StringComparison.Ordinal) &&
                    string.Equals(
                        left.PivotTableName,
                        right.PivotTableName,
                        StringComparison.Ordinal) &&
                    left.IsSelectedTarget == right.IsSelectedTarget &&
                    string.Equals(left.Name, right.Name, StringComparison.Ordinal) &&
                    string.Equals(left.RawFormula, right.RawFormula, StringComparison.Ordinal) &&
                    string.Equals(
                        left.FormulaFingerprint,
                        right.FormulaFingerprint,
                        StringComparison.Ordinal) &&
                   string.Equals(left.SourceName, right.SourceName, StringComparison.Ordinal) &&
                   string.Equals(left.Caption, right.Caption, StringComparison.Ordinal) &&
                   string.Equals(
                       left.DisplayFolder,
                       right.DisplayFolder,
                       StringComparison.Ordinal) &&
                   left.PairState == right.PairState &&
                   left.CalculatedMemberType == right.CalculatedMemberType &&
                   left.CubeFieldType == right.CubeFieldType &&
                   left.Dynamic == right.Dynamic &&
                   left.CalculatedMemberFlattenHierarchies ==
                       right.CalculatedMemberFlattenHierarchies &&
                   left.CubeFieldFlattenHierarchies == right.CubeFieldFlattenHierarchies &&
                   left.CalculatedMemberHierarchizeDistinct ==
                       right.CalculatedMemberHierarchizeDistinct &&
                   left.CubeFieldHierarchizeDistinct ==
                       right.CubeFieldHierarchizeDistinct &&
                   left.ShowInFieldList == right.ShowInFieldList &&
                   left.Orientation == right.Orientation &&
                    left.IsValid == right.IsValid &&
                    string.Equals(
                        left.SourceFingerprint,
                        right.SourceFingerprint,
                        StringComparison.Ordinal) &&
                    string.Equals(
                        left.ModelLineageFingerprint,
                        right.ModelLineageFingerprint,
                        StringComparison.Ordinal);
        }

        private static void DemandUnrelatedInventoryUnchanged(
            PivotNamedSetWorkbookSnapshot before,
            PivotNamedSetWorkbookSnapshot after,
            string? allowedSelectedArtifactName)
        {
            if (!string.Equals(
                    before.SourceFingerprint,
                    after.SourceFingerprint,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    before.ModelLineageFingerprint,
                    after.ModelLineageFingerprint,
                    StringComparison.Ordinal) ||
                before.Pivots.Count != after.Pivots.Count)
            {
                throw new InvalidOperationException(
                    "The Data Model source or PivotTable inventory changed during named-set mutation.");
            }

            var afterPivots = new Dictionary<string, PivotNamedSetPivotSnapshot>(
                StringComparer.OrdinalIgnoreCase);
            foreach (PivotNamedSetPivotSnapshot pivot in after.Pivots)
            {
                string key = PivotInventoryKey(pivot.WorksheetName, pivot.PivotTableName);
                if (afterPivots.ContainsKey(key))
                {
                    throw new InvalidOperationException(
                        "Excel exposed an ambiguous PivotTable inventory after named-set mutation.");
                }

                afterPivots.Add(key, pivot);
            }

            var beforeKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (PivotNamedSetPivotSnapshot prior in before.Pivots)
            {
                string key = PivotInventoryKey(prior.WorksheetName, prior.PivotTableName);
                if (!beforeKeys.Add(key) ||
                    !afterPivots.TryGetValue(key, out PivotNamedSetPivotSnapshot? current) ||
                    !string.Equals(
                        prior.WorksheetName,
                        current.WorksheetName,
                        StringComparison.Ordinal) ||
                    !string.Equals(
                        prior.PivotTableName,
                        current.PivotTableName,
                        StringComparison.Ordinal) ||
                    prior.IsSelectedTarget != current.IsSelectedTarget)
                {
                    throw new InvalidOperationException(
                        "The exact PivotTable roster changed during named-set mutation.");
                }

                string? excluded = prior.IsSelectedTarget
                    ? allowedSelectedArtifactName
                    : null;
                DemandSameArtifactInventory(prior.Artifacts, current.Artifacts, excluded);
                DemandSameCalculatedMemberInventory(
                    prior.CalculatedMembers,
                    current.CalculatedMembers,
                    excluded);
            }
        }

        private static void DemandExactInventoryTransition(
            PivotNamedSetWorkbookSnapshot before,
            PivotNamedSetWorkbookSnapshot after,
            string name,
            LivePivotNamedSetSnapshot? expectedBefore,
            LivePivotNamedSetSnapshot? expectedAfter)
        {
            DemandUnrelatedInventoryUnchanged(before, after, name);
            LivePivotNamedSetSnapshot? actualBefore = FindSelectedArtifact(before, name);
            LivePivotNamedSetSnapshot? actualAfter = FindSelectedArtifact(after, name);
            if ((expectedBefore == null) != (actualBefore == null) ||
                (expectedAfter == null) != (actualAfter == null) ||
                (expectedBefore != null && !SameSnapshot(actualBefore!, expectedBefore)) ||
                (expectedAfter != null && !SameSnapshot(actualAfter!, expectedAfter)))
            {
                throw new InvalidOperationException(
                    "The selected named-set state transition was not exact.");
            }

            if (expectedAfter != null && !after.SelectedPivot.ConnectionRefreshed)
            {
                throw new InvalidOperationException(
                    "The final named-set state was not validated after PivotCache.MakeConnection.");
            }
        }

        private static LivePivotNamedSetSnapshot? FindSelectedArtifact(
            PivotNamedSetWorkbookSnapshot snapshot,
            string name)
        {
            return snapshot.SelectedPivot.Artifacts.SingleOrDefault(artifact => string.Equals(
                artifact.Name,
                name,
                StringComparison.OrdinalIgnoreCase));
        }

        private static void DemandSameArtifactInventory(
            IReadOnlyList<LivePivotNamedSetSnapshot> before,
            IReadOnlyList<LivePivotNamedSetSnapshot> after,
            string? excludedName)
        {
            List<LivePivotNamedSetSnapshot> prior = before
                .Where(value => excludedName == null || !string.Equals(
                    value.Name,
                    excludedName,
                    StringComparison.OrdinalIgnoreCase))
                .OrderBy(value => value.Name, StringComparer.Ordinal)
                .ThenBy(value => value.PairState)
                .ToList();
            List<LivePivotNamedSetSnapshot> current = after
                .Where(value => excludedName == null || !string.Equals(
                    value.Name,
                    excludedName,
                    StringComparison.OrdinalIgnoreCase))
                .OrderBy(value => value.Name, StringComparer.Ordinal)
                .ThenBy(value => value.PairState)
                .ToList();
            if (prior.Count != current.Count ||
                prior.Where((value, index) => !SameSnapshot(value, current[index])).Any())
            {
                throw new InvalidOperationException(
                    "An unrelated named-set artifact changed during mutation.");
            }
        }

        private static void DemandSameCalculatedMemberInventory(
            IReadOnlyList<PivotCalculatedMemberReferenceSnapshot> before,
            IReadOnlyList<PivotCalculatedMemberReferenceSnapshot> after,
            string? excludedName)
        {
            List<PivotCalculatedMemberReferenceSnapshot> prior = before
                .Where(value => excludedName == null || !string.Equals(
                    value.Name,
                    excludedName,
                    StringComparison.OrdinalIgnoreCase))
                .OrderBy(value => value.Name, StringComparer.Ordinal)
                .ThenBy(value => value.Type)
                .ToList();
            List<PivotCalculatedMemberReferenceSnapshot> current = after
                .Where(value => excludedName == null || !string.Equals(
                    value.Name,
                    excludedName,
                    StringComparison.OrdinalIgnoreCase))
                .OrderBy(value => value.Name, StringComparer.Ordinal)
                .ThenBy(value => value.Type)
                .ToList();
            if (prior.Count != current.Count)
            {
                throw new InvalidOperationException(
                    "The calculated-member inventory changed during named-set mutation.");
            }

            for (var index = 0; index < prior.Count; index++)
            {
                PivotCalculatedMemberReferenceSnapshot left = prior[index];
                PivotCalculatedMemberReferenceSnapshot right = current[index];
                if (!string.Equals(
                        left.WorksheetName,
                        right.WorksheetName,
                        StringComparison.Ordinal) ||
                    !string.Equals(
                        left.PivotTableName,
                        right.PivotTableName,
                        StringComparison.Ordinal) ||
                    !string.Equals(left.Name, right.Name, StringComparison.Ordinal) ||
                    left.Type != right.Type ||
                    !string.Equals(left.RawFormula, right.RawFormula, StringComparison.Ordinal) ||
                    left.FormulaScanComplete != right.FormulaScanComplete)
                {
                    throw new InvalidOperationException(
                        "An unrelated calculated-member changed during named-set mutation.");
                }
            }
        }

        private static string PivotInventoryKey(string worksheetName, string pivotTableName)
        {
            return worksheetName.Length.ToString(CultureInfo.InvariantCulture) +
                   ":" + worksheetName + "|" + pivotTableName;
        }

        private static void DemandStillBound(BoundPivotNamedSetTarget target)
        {
            dynamic pivot = target.PivotTable;
            object worksheet = ReadRequired(
                () => (object?)pivot.Parent,
                "Excel no longer exposes the selected PivotTable worksheet.");
            dynamic nativeWorksheet = worksheet;
            object workbook = ReadRequired(
                () => (object?)nativeWorksheet.Parent,
                "Excel no longer exposes the selected PivotTable workbook.");
            if (!ComObjectIdentity.AreSame(workbook, target.Workbook))
            {
                throw new InvalidOperationException(
                    "The bound PivotTable moved to a different workbook.");
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
                    StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(
                    new StoredWorkbookIdentityResolver().Resolve(target.Workbook),
                    target.Identity.WorkbookId,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The bound PivotTable identity changed after discovery.");
            }

            object cache = ReadPivotCache(target.PivotTable);
            dynamic nativeCache = cache;
            if (!ReadRequiredBoolean(
                    () => (object?)nativeCache.OLAP,
                    "selected PivotCache.OLAP"))
            {
                throw new InvalidOperationException(
                    "The bound PivotTable is no longer OLAP-backed.");
            }

            object connection = ReadRequired(
                () => (object?)nativeCache.WorkbookConnection,
                "Excel no longer exposes the selected PivotCache connection.");
            if (!ComObjectIdentity.AreSame(connection, target.DataModelConnection))
            {
                throw new InvalidOperationException(
                    "The bound PivotTable no longer uses the exact Data Model connection.");
            }

            dynamic nativeWorkbook = target.Workbook;
            object model = ReadRequired(
                () => (object?)nativeWorkbook.Model,
                "Excel no longer exposes the workbook Data Model.");
            if (!ComObjectIdentity.AreSame(model, target.Model))
            {
                throw new InvalidOperationException(
                    "The workbook Data Model identity changed after discovery.");
            }

            dynamic nativeModel = model;
            object dataModelConnection = ReadRequired(
                () => (object?)nativeModel.DataModelConnection,
                "Excel no longer exposes Workbook.Model.DataModelConnection.");
            if (!ComObjectIdentity.AreSame(
                    dataModelConnection,
                    target.DataModelConnection))
            {
                throw new InvalidOperationException(
                    "The workbook Data Model connection identity changed after discovery.");
            }
        }

        private static PivotNamedSetPivotSnapshot ReadPivotSnapshot(
            object pivotTable,
            string worksheetName,
            bool isSelected,
            string sourceFingerprint,
            bool requireConnectionRefresh,
            string modelLineageFingerprint)
        {
            dynamic pivot = pivotTable;
            var connectionRefreshed = true;
            var connectionAttempted = false;

            string pivotName = ReadBoundedRequiredString(
                () => (object?)pivot.Name,
                MaximumNameCharacters,
                "PivotTable name");
            object calculatedMembersObject = ReadRequiredCollectionMember(
                () => (object?)pivot.CalculatedMembers,
                () => (object?)pivot.CalculatedMembers(),
                "Excel did not expose PivotTable CalculatedMembers.");
            var calculatedSets = new Dictionary<string, CalculatedMemberHandle>(
                StringComparer.OrdinalIgnoreCase);
            var references = new List<PivotCalculatedMemberReferenceSnapshot>();
            var calculatedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (object calculatedMemberObject in ReadCollection(
                         calculatedMembersObject,
                         MaximumCalculatedMembers,
                         "PivotTable CalculatedMembers"))
            {
                dynamic calculatedMember = calculatedMemberObject;
                string name = ReadBoundedRequiredString(
                    () => (object?)calculatedMember.Name,
                    MaximumNameCharacters,
                    "calculated-member name");
                if (!calculatedNames.Add(name))
                {
                    throw new InvalidOperationException(
                        "Excel exposed duplicate calculated-member names.");
                }

                int type = ReadRequiredInt(
                    () => (object?)calculatedMember.Type,
                    "calculated-member type");
                string readback = ReadBoundedRequiredString(
                    () => (object?)calculatedMember.Formula,
                    MaximumFormulaCharacters,
                    "calculated-member formula");
                bool decoded = PivotNamedSetFormulaTransport.TryDecodeReadback(
                    readback,
                    out string rawFormula);
                bool scanComplete = false;
                if (decoded)
                {
                    scanComplete = MdxNamedSetReferenceScanner.Scan(rawFormula).IsComplete;
                }

                references.Add(new PivotCalculatedMemberReferenceSnapshot(
                    worksheetName,
                    pivotName,
                    name,
                    type,
                    decoded ? rawFormula : string.Empty,
                    scanComplete));
                if (type != CalculatedMemberTypeSet)
                {
                    continue;
                }

                if (!decoded)
                {
                    throw new InvalidOperationException(
                        "Excel exposed an unsupported named-set formula readback.");
                }

                if (!connectionAttempted)
                {
                    connectionRefreshed = TryMakeConnection(pivotTable);
                    connectionAttempted = true;
                    if (requireConnectionRefresh && !connectionRefreshed)
                    {
                        throw new InvalidOperationException(
                            "Excel could not refresh a Data Model PivotCache before named-set validation.");
                    }
                }

                var handle = new CalculatedMemberHandle(
                    calculatedMemberObject,
                    name,
                    type,
                    rawFormula,
                    ReadBoundedOptionalString(
                        () => (object?)calculatedMember.DisplayFolder,
                        MaximumNameCharacters,
                        "calculated-set DisplayFolder"),
                    ReadRequiredBoolean(
                        () => (object?)calculatedMember.Dynamic,
                        "calculated-set Dynamic"),
                    ReadRequiredBoolean(
                        () => (object?)calculatedMember.FlattenHierarchies,
                        "calculated-set FlattenHierarchies"),
                    ReadRequiredBoolean(
                        () => (object?)calculatedMember.HierarchizeDistinct,
                        "calculated-set HierarchizeDistinct"),
                    connectionRefreshed
                        ? ReadOptionalBoolean(
                            () => (object?)calculatedMember.IsValid,
                            "calculated-set IsValid")
                        : null);
                calculatedSets.Add(name, handle);
            }

            object cubeFieldsObject = ReadRequiredCollectionMember(
                () => (object?)pivot.CubeFields,
                () => (object?)pivot.CubeFields(),
                "Excel did not expose PivotTable CubeFields.");
            var cubeSets = new Dictionary<string, CubeSetFieldHandle>(
                StringComparer.OrdinalIgnoreCase);
            foreach (object cubeFieldObject in ReadCollection(
                         cubeFieldsObject,
                         MaximumCubeFields,
                         "PivotTable CubeFields"))
            {
                dynamic cubeField = cubeFieldObject;
                int cubeFieldType = ReadCubeFieldType(cubeField);
                if (cubeFieldType != CubeFieldTypeSet) continue;
                string sourceName = ReadProviderUniqueName(
                    cubeField,
                    "named-set CubeField source name");
                if (cubeSets.ContainsKey(sourceName))
                {
                    throw new InvalidOperationException(
                        "Excel exposed duplicate named-set CubeFields.");
                }

                cubeSets.Add(
                    sourceName,
                    new CubeSetFieldHandle(
                        cubeFieldObject,
                        sourceName,
                        ReadBoundedOptionalString(
                            () => (object?)cubeField.Caption,
                            MaximumCaptionCharacters,
                            "named-set CubeField caption"),
                        cubeFieldType,
                        ReadRequiredBoolean(
                            () => (object?)cubeField.FlattenHierarchies,
                            "named-set CubeField FlattenHierarchies"),
                        ReadRequiredBoolean(
                            () => (object?)cubeField.HierarchizeDistinct,
                            "named-set CubeField HierarchizeDistinct"),
                        ReadRequiredBoolean(
                            () => (object?)cubeField.ShowInFieldList,
                            "named-set CubeField ShowInFieldList"),
                        ReadRequiredInt(
                            () => (object?)cubeField.Orientation,
                            "named-set CubeField orientation")));
            }

            var keys = new HashSet<string>(
                calculatedSets.Keys,
                StringComparer.OrdinalIgnoreCase);
            keys.UnionWith(cubeSets.Keys);
            var artifacts = new List<LivePivotNamedSetSnapshot>();
            foreach (string key in keys.OrderBy(value => value, StringComparer.Ordinal))
            {
                calculatedSets.TryGetValue(key, out CalculatedMemberHandle? calculated);
                cubeSets.TryGetValue(key, out CubeSetFieldHandle? cube);
                PivotNamedSetPairState pairState = calculated != null && cube != null
                    ? PivotNamedSetPairState.Complete
                    : calculated != null
                        ? PivotNamedSetPairState.CalculatedMemberOnly
                        : PivotNamedSetPairState.CubeFieldOnly;
                string name = calculated?.Name ?? cube!.SourceName;
                string formulaFingerprint = calculated == null
                    ? string.Empty
                    : PivotMdxFingerprint.ComputeFormula(calculated.RawFormula);
                string liveFingerprint = PivotNamedSetCanonical.CreateLiveFingerprint(
                    sourceFingerprint,
                    modelLineageFingerprint,
                    name,
                    pairState,
                    formulaFingerprint,
                    calculated?.DisplayFolder ?? string.Empty,
                    cube?.SourceName ?? string.Empty,
                    cube?.Caption ?? string.Empty,
                    calculated?.Type,
                    cube?.Type,
                    calculated?.Dynamic,
                    calculated?.FlattenHierarchies,
                    cube?.FlattenHierarchies,
                    calculated?.HierarchizeDistinct,
                    cube?.HierarchizeDistinct,
                    cube?.ShowInFieldList,
                    cube?.Orientation,
                    calculated?.IsValid);
                artifacts.Add(new LivePivotNamedSetSnapshot(
                    worksheetName,
                    pivotName,
                    isSelected,
                    name,
                    pairState,
                    calculated?.RawFormula ?? string.Empty,
                    formulaFingerprint,
                    calculated?.DisplayFolder ?? string.Empty,
                    cube?.SourceName ?? string.Empty,
                    cube?.Caption ?? string.Empty,
                    calculated?.Type,
                    cube?.Type,
                    calculated?.Dynamic,
                    calculated?.FlattenHierarchies,
                    cube?.FlattenHierarchies,
                    calculated?.HierarchizeDistinct,
                    cube?.HierarchizeDistinct,
                    cube?.ShowInFieldList,
                    cube?.Orientation,
                    calculated?.IsValid,
                    sourceFingerprint,
                    modelLineageFingerprint,
                    liveFingerprint));
            }

            return new PivotNamedSetPivotSnapshot(
                worksheetName,
                pivotName,
                isSelected,
                artifacts,
                references,
                connectionRefreshed,
                PivotNamedSetCanonical.CreatePivotFingerprint(artifacts, references));
        }

        private static bool TryMakeConnection(object pivotTable)
        {
            return PivotLateBound.TryRead(
                       () =>
                       {
                           object cacheObject = ReadPivotCache(pivotTable);
                           dynamic cache = cacheObject;
                           cache.MakeConnection();
                           return (object?)true;
                       },
                       out object? result) &&
                   result is bool refreshed &&
                   refreshed;
        }

        private static void DemandMakeConnection(object pivotTable)
        {
            if (!TryMakeConnection(pivotTable))
            {
                throw new InvalidOperationException(
                    "Excel could not refresh the exact Data Model PivotCache.");
            }
        }

        private static bool IsWorkbookModelPivot(
            object pivotTable,
            object dataModelConnection)
        {
            object cache = ReadPivotCache(pivotTable);
            dynamic nativeCache = cache;
            if (!ReadRequiredBoolean(
                    () => (object?)nativeCache.OLAP,
                    "workbook PivotCache.OLAP"))
            {
                return false;
            }

            object connection = ReadRequired(
                () => (object?)nativeCache.WorkbookConnection,
                "Excel did not expose an OLAP PivotCache connection.");
            return ComObjectIdentity.AreSame(connection, dataModelConnection);
        }

        private static object ReadRequiredPivotTables(dynamic worksheet)
        {
            return ReadRequiredCollectionMember(
                () => (object?)worksheet.PivotTables,
                () => (object?)worksheet.PivotTables(),
                "Excel did not expose worksheet PivotTables.");
        }

        private static object ReadRequiredCollectionMember(
            Func<object?> propertyReader,
            Func<object?> methodReader,
            string message)
        {
            if (TryReadCollectionMember(propertyReader, methodReader, out object? value) &&
                value != null)
            {
                return value;
            }

            throw new InvalidOperationException(message);
        }

        private static bool? ReadOptionalBoolean(Func<object?> reader, string label)
        {
            if (!PivotLateBound.TryRead(reader, out object? value) || value == null)
            {
                return null;
            }

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

        private static string ReadModelLineageFingerprint(
            BoundPivotNamedSetTarget target)
        {
            dynamic model = target.Model;
            object modelTablesObject = ReadRequired(
                () => (object?)model.ModelTables,
                "Excel did not expose Data Model tables.");
            IReadOnlyList<object> tables = ReadCollection(
                modelTablesObject,
                MaximumModelTables,
                "Data Model tables");
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var tokens = new List<string>
            {
                CatalogToken(
                    "data-model-connection-type",
                    DataModelConnectionType.ToString(CultureInfo.InvariantCulture))
            };
            foreach (object tableObject in tables)
            {
                dynamic table = tableObject;
                string tableName = ReadBoundedRequiredString(
                    () => (object?)table.Name,
                    MaximumNameCharacters,
                    "Data Model table name");
                if (!names.Add(tableName))
                {
                    throw new InvalidOperationException(
                        "Excel exposed duplicate Data Model table names.");
                }

                object connectionObject = ReadRequired(
                    () => (object?)table.SourceWorkbookConnection,
                    "Excel did not expose the Data Model table source connection.");
                dynamic connection = connectionObject;
                string connectionName = ReadBoundedRequiredString(
                    () => (object?)connection.Name,
                    MaximumNameCharacters,
                    "Data Model table source connection name");
                int connectionType = ReadRequiredInt(
                    () => (object?)connection.Type,
                    "Data Model table source connection type");
                tokens.Add(CatalogToken(
                    "table",
                    tableName,
                    connectionName,
                    connectionType.ToString(CultureInfo.InvariantCulture)));
            }

            return PivotNamedSetCanonical.CreateModelLineageFingerprint(tokens);
        }

        private static int ReadCubeFieldType(dynamic cubeField)
        {
            if (PivotLateBound.TryRead(
                    () => (object?)cubeField.CubeFieldType,
                    out object? cubeFieldType) &&
                cubeFieldType != null)
            {
                return ConvertRequiredInt(cubeFieldType, "CubeField type");
            }

            return ReadRequiredInt(
                () => (object?)cubeField.Type,
                "CubeField type");
        }

        private static string ReadProviderUniqueName(dynamic value, string label)
        {
            if (PivotLateBound.TryRead(
                    () => (object?)value.SourceName,
                    out object? sourceNameValue) &&
                sourceNameValue is string sourceName &&
                !string.IsNullOrWhiteSpace(sourceName))
            {
                DemandProviderUniqueName(sourceName, label);
                return sourceName;
            }

            string name = ReadBoundedRequiredString(
                () => (object?)value.Name,
                MaximumProviderUniqueNameCharacters,
                label);
            DemandProviderUniqueName(name, label);
            return name;
        }

        private static void DemandProviderUniqueName(string value, string label)
        {
            if (!IsProviderUniqueName(value))
            {
                throw new NotSupportedException(
                    "Excel exposed an invalid " + label + ".");
            }
        }

        private static bool IsProviderUniqueName(string value)
        {
            if (string.IsNullOrWhiteSpace(value) ||
                value.Length > MaximumProviderUniqueNameCharacters ||
                !string.Equals(value, value.Trim(), StringComparison.Ordinal) ||
                value.Any(char.IsControl))
            {
                return false;
            }

            var index = 0;
            var segment = 0;
            while (index < value.Length)
            {
                if (segment > 0)
                {
                    char delimiter = value[index];
                    if (delimiter != '.' && delimiter != '&') return false;
                    index++;
                    if (index >= value.Length) return false;
                    if (delimiter == '.' && value[index] == '&') index++;
                }

                if (index >= value.Length || value[index] != '[') return false;
                index++;
                var content = 0;
                var closed = false;
                while (index < value.Length)
                {
                    if (value[index] != ']')
                    {
                        content++;
                        index++;
                        continue;
                    }

                    if (index + 1 < value.Length && value[index + 1] == ']')
                    {
                        content++;
                        index += 2;
                        continue;
                    }

                    index++;
                    closed = true;
                    break;
                }

                if (!closed || content == 0) return false;
                segment++;
            }

            return segment > 0;
        }

        private static bool TryReadCollectionMember(
            Func<object?> propertyReader,
            Func<object?> methodReader,
            out object? collection)
        {
            if (PivotLateBound.TryRead(propertyReader, out collection) && collection != null)
            {
                return true;
            }

            return PivotLateBound.TryRead(methodReader, out collection) && collection != null;
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
                "Excel did not expose the PivotTable cache.");
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

            if (result.Length > maximum || result.Any(char.IsControl))
            {
                throw new NotSupportedException(
                    "Excel exposed an invalid or unbounded " + label + ".");
            }

            return result;
        }

        private static int ReadRequiredInt(Func<object?> reader, string label)
        {
            return ConvertRequiredInt(
                ReadRequired(reader, "Excel did not expose " + label + "."),
                label);
        }

        private static int ConvertRequiredInt(object value, string label)
        {
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

        private static string CatalogToken(string kind, params string[] values)
        {
            var result = new StringBuilder();
            result.Append(kind.Length.ToString(CultureInfo.InvariantCulture))
                .Append(':')
                .Append(kind);
            foreach (string value in values)
            {
                string actual = value ?? string.Empty;
                result.Append('|')
                    .Append(actual.Length.ToString(CultureInfo.InvariantCulture))
                    .Append(':')
                    .Append(actual);
            }

            return result.ToString();
        }

        private static string? NullIfEmpty(string value)
        {
            return string.IsNullOrEmpty(value) ? null : value;
        }

        private sealed class LiveMeasureDependency
        {
            public LiveMeasureDependency(
                string name,
                string formulaFingerprint,
                string description)
            {
                Name = name;
                FormulaFingerprint = formulaFingerprint;
                Description = description;
            }

            public string Name { get; }
            public string FormulaFingerprint { get; }
            public string Description { get; }
        }

        private sealed class CalculatedMemberHandle
        {
            public CalculatedMemberHandle(
                object native,
                string name,
                int type,
                string rawFormula,
                string displayFolder,
                bool dynamic,
                bool flattenHierarchies,
                bool hierarchizeDistinct,
                bool? isValid)
            {
                Native = native;
                Name = name;
                Type = type;
                RawFormula = rawFormula;
                DisplayFolder = displayFolder;
                Dynamic = dynamic;
                FlattenHierarchies = flattenHierarchies;
                HierarchizeDistinct = hierarchizeDistinct;
                IsValid = isValid;
            }

            public object Native { get; }
            public string Name { get; }
            public int Type { get; }
            public string RawFormula { get; }
            public string DisplayFolder { get; }
            public bool Dynamic { get; }
            public bool FlattenHierarchies { get; }
            public bool HierarchizeDistinct { get; }
            public bool? IsValid { get; }
        }

        private sealed class CubeSetFieldHandle
        {
            public CubeSetFieldHandle(
                object native,
                string sourceName,
                string caption,
                int type,
                bool flattenHierarchies,
                bool hierarchizeDistinct,
                bool showInFieldList,
                int orientation)
            {
                Native = native;
                SourceName = sourceName;
                Caption = caption;
                Type = type;
                FlattenHierarchies = flattenHierarchies;
                HierarchizeDistinct = hierarchizeDistinct;
                ShowInFieldList = showInFieldList;
                Orientation = orientation;
            }

            public object Native { get; }
            public string SourceName { get; }
            public string Caption { get; }
            public int Type { get; }
            public bool FlattenHierarchies { get; }
            public bool HierarchizeDistinct { get; }
            public bool ShowInFieldList { get; }
            public int Orientation { get; }
        }

        private sealed class NativeSetPair
        {
            public NativeSetPair(
                CalculatedMemberHandle? calculatedMember,
                CubeSetFieldHandle? cubeField)
            {
                CalculatedMember = calculatedMember;
                CubeField = cubeField;
            }

            public CalculatedMemberHandle? CalculatedMember { get; }

            public CubeSetFieldHandle? CubeField { get; }
        }
    }
}
