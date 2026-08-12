using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using ExcelReportBuilder.Core.PivotPlus;
using ExcelReportBuilder.Excel.Execution;
using ExcelReportBuilder.Excel.Persistence;
using ExcelReportBuilder.Excel.PivotPlus.Persistence;

namespace ExcelReportBuilder.Excel.PivotPlus.DataModel
{
    /// <summary>
    /// Late-bound Excel implementation of the explicit classic-to-model
    /// conversion boundary. It owns only names derived for the current setup;
    /// the selected worksheet and PivotTable are never registered as owned.
    /// </summary>
    public sealed class LateBoundPivotDataModelEnablementGateway : IPivotDataModelEnablementGateway
    {
        private const int SourceExternal = 2;
        private const int SourceDatabase = 1;
        private const int PivotVersion15 = 6;
        private const int OrientationHidden = 0;
        private const int OrientationRow = 1;
        private const int OrientationColumn = 2;
        private const int OrientationPage = 3;
        private const int OrientationData = 4;
        private const int SortManual = -4135;
        private const int SheetVeryHidden = 2;
        private const int PasteFormats = -4122;
        private const int PasteColumnWidths = 8;
        private const int CellTypeAllValidation = -4174;
        private const int MissingItemsDefault = -1;
        private const int ExcelNoCellsFoundError = unchecked((int)0x800A03EC);
        private const int MaximumWorksheets = 1024;
        private const int MaximumWorkbookObjects = 4096;
        private const int MaximumFields = 512;
        private const int MaximumMembers = 4096;
        private const long MaximumRawRangeCells = 10_000_000L;
        private const int MaximumPivotResultCells = 100_000;
        private const int MaximumPivotResultCharacters = 4 * 1024 * 1024;
        private const int MaximumPivotCellCharacters = 4096;
        private const int ExcelMaximumRows = 1_048_576;
        private const int ExcelMaximumColumns = 16_384;
        private const string TemporaryPurposeProperty = "PivotTablePlusPurpose";
        private const string TemporaryFingerprintProperty = "PivotTablePlusFingerprint";
        private const string TemporaryAnchorProperty = "PivotTablePlusTargetAnchor";
        private const string StagingStateFingerprintProperty =
            "PivotTablePlusStagingStateFingerprint";
        private static readonly Regex A1RangePattern = new Regex(
            @"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})(?::\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6}))?$",
            RegexOptions.CultureInvariant);
        private static readonly Regex R1C1RangePattern = new Regex(
            @"^R([1-9][0-9]{0,6})C([1-9][0-9]{0,4})(?::R([1-9][0-9]{0,6})C([1-9][0-9]{0,4}))?$",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        public void VerifyBoundTarget(
            object workbook,
            object pivotTable,
            PivotTargetIdentity expectedTarget)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            if (expectedTarget == null) throw new ArgumentNullException(nameof(expectedTarget));

            dynamic pivot = pivotTable;
            string pivotName = ReadRequiredString(
                () => (object?)pivot.Name,
                "Excel did not expose the selected PivotTable name.");
            object worksheetObject = ReadRequired(
                () => (object?)pivot.Parent,
                "Excel did not expose the selected PivotTable worksheet.");
            dynamic worksheet = worksheetObject;
            string worksheetName = ReadRequiredString(
                () => (object?)worksheet.Name,
                "Excel did not expose the selected PivotTable worksheet name.");
            object liveWorkbook = ReadRequired(
                () => (object?)worksheet.Parent,
                "Excel did not expose the selected PivotTable workbook.");
            if (!SameNativeObject(workbook, liveWorkbook))
            {
                throw new InvalidOperationException(
                    "The supplied workbook is not the selected PivotTable's workbook.");
            }

            if (!string.Equals(
                    worksheetName,
                    expectedTarget.WorksheetName,
                    StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(
                    pivotName,
                    expectedTarget.PivotTableName,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "The selected PivotTable no longer matches the discovered PivotTable+ target.");
            }

            string workbookId = new StoredWorkbookIdentityResolver().Resolve(workbook);
            if (!string.Equals(
                    workbookId,
                    expectedTarget.WorkbookId,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The selected workbook no longer matches the discovered path-free workbook identity.");
            }
        }

        public void PersistBoundWorkbookIdentity(
            object workbook,
            PivotTargetIdentity expectedTarget)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (expectedTarget == null) throw new ArgumentNullException(nameof(expectedTarget));
            new StoredWorkbookIdentityResolver().Persist(
                workbook,
                expectedTarget.WorkbookId);
        }

        public ClassicPivotSourceDescriptor InspectSupportedSource(
            object workbook,
            object pivotTable,
            PivotSourceDescriptor expectedSource)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            if (expectedSource == null) throw new ArgumentNullException(nameof(expectedSource));

            object cacheObject = ReadPivotCache(pivotTable);
            dynamic cache = cacheObject;
            if (ReadBoolean(() => (object?)cache.OLAP, "PivotCache.OLAP"))
            {
                throw new NotSupportedException(
                    "The selected PivotTable already uses an OLAP or Data Model source.");
            }

            if (ReadRequiredInt(
                    () => (object?)cache.SourceType,
                    "classic PivotCache source type") != SourceDatabase)
            {
                throw new NotSupportedException(
                    "Only an internal worksheet database PivotCache can be upgraded to the Data Model.");
            }

            string sourceToken = ReadClassicSourceToken(cache);
            DemandExpectedClassicSource(expectedSource, sourceToken);
            if (TryResolveTable(workbook, sourceToken, out string? tableName))
            {
                return new ClassicPivotSourceDescriptor(
                    tableName!,
                    Core.PivotPlus.PivotPlusWorkbookObjectKind.Table);
            }

            if (TryResolveWorkbookName(workbook, sourceToken, out string? rangeName))
            {
                return new ClassicPivotSourceDescriptor(
                    rangeName!,
                    Core.PivotPlus.PivotPlusWorkbookObjectKind.NamedRange);
            }

            if (TryResolveWorksheetRange(
                    workbook,
                    sourceToken,
                    out object? nativeRange,
                    out string? canonicalReference))
            {
                return new ClassicPivotSourceDescriptor(
                    nativeRange!,
                    canonicalReference!);
            }

            throw new NotSupportedException(
                "Data Model enablement requires a same-workbook Excel table, workbook-scoped name, or bounded single-area worksheet range.");
        }

        internal ClassicPivotSourceDescriptor InspectSupportedSource(
            object workbook,
            object pivotTable)
        {
            dynamic cache = ReadPivotCache(pivotTable);
            string token = ReadClassicSourceToken(cache);
            return InspectSupportedSource(
                workbook,
                pivotTable,
                new PivotSourceDescriptor(
                    PivotSourceKind.WorksheetRange,
                    token,
                    PivotCapability.UpgradeToDataModel));
        }

        private static void DemandExpectedClassicSource(
            PivotSourceDescriptor expectedSource,
            string liveSourceToken)
        {
            if (expectedSource.Kind != PivotSourceKind.WorksheetRange &&
                expectedSource.Kind != PivotSourceKind.WorksheetTable)
            {
                throw new InvalidOperationException(
                    "The discovered PivotTable source is no longer a classic worksheet source.");
            }

            string expected = NormalizeSourceIdentity(expectedSource.SourceName);
            string live = NormalizeSourceIdentity(liveSourceToken);
            bool exact = string.Equals(expected, live, StringComparison.OrdinalIgnoreCase);
            // A table name is workbook-unique, so Excel may legitimately add
            // or omit its worksheet qualifier between discovery and use. A
            // raw range is not workbook-unique: SheetA!A1:D20 and
            // SheetB!A1:D20 must never be treated as the same live source.
            bool leaf = expectedSource.Kind == PivotSourceKind.WorksheetTable &&
                        (SourceTokenMatches(liveSourceToken, expectedSource.SourceName) ||
                         SourceTokenMatches(expectedSource.SourceName, liveSourceToken));
            if (!exact && !leaf)
            {
                throw new InvalidOperationException(
                    "The classic PivotCache source changed after PivotTable+ discovery. Re-select the PivotTable before enabling the Data Model.");
            }
        }

        private static string NormalizeSourceIdentity(string value)
        {
            string normalized = value.Trim().TrimStart('=').Trim();
            int separator = FindFinalUnquotedSeparator(normalized, '!');
            if (separator <= 0) return normalized;
            string sheet = normalized.Substring(0, separator).Trim();
            if (sheet.Length >= 2 && sheet[0] == '\'' && sheet[sheet.Length - 1] == '\'')
            {
                sheet = sheet.Substring(1, sheet.Length - 2).Replace("''", "'");
            }

            // '$' is an absolute-reference marker only in the address. It is
            // a legal, identity-significant character in a worksheet name.
            string address = normalized.Substring(separator + 1)
                .Trim()
                .Replace("$", string.Empty);
            return sheet + "!" + address;
        }

        private static int FindFinalUnquotedSeparator(string value, char separator)
        {
            bool quoted = false;
            int result = -1;
            for (int index = 0; index < value.Length; index++)
            {
                char current = value[index];
                if (current == '\'')
                {
                    if (quoted && index + 1 < value.Length && value[index + 1] == '\'')
                    {
                        index++;
                        continue;
                    }

                    quoted = !quoted;
                    continue;
                }

                if (!quoted && current == separator)
                {
                    result = index;
                }
            }

            return result;
        }

        public PivotNativeStateSnapshot CaptureReversibleState(object pivotTable)
        {
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            DemandNoConnectedSlicersOrTimelines(pivotTable);
            DemandNoAttachedPivotCharts(pivotTable);
            dynamic pivot = pivotTable;
            DemandCompatibleClassicCachePolicy(pivotTable);
            DemandCompatibleClassicSaveData(pivotTable);
            DemandNoUnsupportedClassicDefinitions(pivotTable);
            string pivotName = ReadRequiredString(
                () => (object?)pivot.Name,
                "Excel did not expose the PivotTable name.");
            object worksheetObject = ReadRequired(
                () => (object?)pivot.Parent,
                "Excel did not expose the PivotTable worksheet.");
            dynamic worksheet = worksheetObject;
            string worksheetName = ReadRequiredString(
                () => (object?)worksheet.Name,
                "Excel did not expose the PivotTable worksheet name.");
            object rangeObject = ReadRequired(
                () => (object?)pivot.TableRange2,
                "Excel did not expose the PivotTable range.");
            DemandNoUnsupportedCustomFormatting(pivotTable, rangeObject);
            DemandNoUnsupportedCellMetadata(rangeObject);
            dynamic range = rangeObject;
            object firstCell = ReadRequired(
                () => (object?)range.Cells[1, 1],
                "Excel did not expose the PivotTable anchor cell.");
            string anchor = ReadAddress(firstCell);

            IReadOnlyList<LateBoundFieldState> fields = ReadFieldStates(pivot);
            LateBoundStyleState style = ReadStyleState(pivot);
            DemandCompatibleOlapInvariants(fields, style);
            var state = new LateBoundPivotState(
                ReadPivotCache(pivotTable),
                fields,
                style,
                ReadResultSignature(rangeObject),
                ReadDataAxisState(pivotTable));
            string fingerprint = PivotPlusFingerprint.Create(
                "pivotplus.native-state.v1",
                state.CanonicalValue());
            return new PivotNativeStateSnapshot(
                worksheetName,
                pivotName,
                anchor,
                fingerprint,
                state);
        }

        public PivotDataModelArtifactPlan PlanOwnedModelArtifacts(
            string setupId,
            ClassicPivotSourceDescriptor source,
            PivotTargetIdentity target,
            PivotNativeStateSnapshot originalState)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (target == null) throw new ArgumentNullException(nameof(target));
            if (originalState == null) throw new ArgumentNullException(nameof(originalState));
            if (!string.Equals(
                    target.WorksheetName,
                    originalState.WorksheetName,
                    StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(
                    target.PivotTableName,
                    originalState.PivotTableName,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "The captured PivotTable target changed before recovery ownership was planned.");
            }

            PivotPlusMetadataValidator.ValidateLocalA1Address(
                originalState.AnchorAddress);
            GeneratedNames names = GeneratedNames.For(setupId);
            string queryObjectName = source.WorkbookObjectName;
            PivotPlusWorkbookObjectKind queryObjectKind = source.ObjectKind;
            string? workbookName = null;
            string? workbookNameFingerprint = null;
            string? requestedReference = null;
            if (source.RequiresOwnedWorkbookName)
            {
                if (source.NativeRange == null ||
                    string.IsNullOrWhiteSpace(source.CanonicalReference))
                {
                    throw new InvalidOperationException(
                        "The inspected raw PivotTable source no longer has a resolvable range receipt.");
                }

                workbookName = names.SourceAliasName;
                requestedReference = source.CanonicalReference;
                workbookNameFingerprint = WorkbookNameFingerprint(
                    requestedReference!);
                queryObjectName = names.SourceAliasName;
                queryObjectKind = PivotPlusWorkbookObjectKind.NamedRange;
            }

            string queryFormula = PivotPlusSourceQueryCompiler.Compile(
                queryObjectName,
                queryObjectKind);
            string connectionString = CanonicalConnectionContract.ConnectionString(
                names.QueryName);
            string commandText = CanonicalConnectionContract.CommandText(
                names.QueryName);
            return new PivotDataModelArtifactPlan(
                names.QueryName,
                names.ConnectionName,
                names.QueryName,
                queryFormula,
                PivotPlusFingerprint.Create("pivotplus.query.v1", queryFormula),
                PivotPlusFingerprint.Create(
                    "pivotplus.connection.v1",
                    connectionString + "\n" + commandText),
                workbookName,
                workbookNameFingerprint,
                requestedReference,
                CreateTemporaryWorksheetReceipts(
                    names,
                    originalState.AnchorAddress),
                CreateTemporaryPivotTableReceipt(
                    setupId,
                    names,
                    target,
                    originalState.AnchorAddress));
        }

        internal PivotDataModelArtifactPlan PlanOwnedModelArtifacts(
            string setupId,
            ClassicPivotSourceDescriptor source)
        {
            var target = new PivotTargetIdentity(
                string.Empty,
                "Sheet1",
                "Pivot1");
            return PlanOwnedModelArtifacts(
                setupId,
                source,
                target,
                new PivotNativeStateSnapshot(
                    target.WorksheetName,
                    target.PivotTableName,
                    "A1",
                    "test-only-plan-snapshot",
                    new object()));
        }

        public void PreflightOwnedModelArtifacts(
            object workbook,
            PivotDataModelArtifactPlan plan,
            PivotPlusWorkbookMetadata? recoveryOwnership)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (plan == null) throw new ArgumentNullException(nameof(plan));
            dynamic book = workbook;
            object queries = ReadRequired(
                () => (object?)book.Queries,
                "Excel did not expose workbook queries during artifact preflight.");
            object connections = ReadRequired(
                () => (object?)book.Connections,
                "Excel did not expose workbook connections during artifact preflight.");
            object? query = FindNamedObject(queries, plan.QueryName, "queries");
            object? connection = FindNamedObject(
                connections,
                plan.ConnectionName,
                "connections");
            object? workbookName = null;
            if (plan.WorkbookName != null)
            {
                workbookName = FindNamedObject(
                    ReadRequired(
                        () => (object?)book.Names,
                        "Excel did not expose workbook names during artifact preflight."),
                    plan.WorkbookName,
                    "names",
                    workbookScopedName: true);
            }

            object worksheets = ReadRequired(
                () => (object?)book.Worksheets,
                "Excel did not expose workbook worksheets during artifact preflight.");
            PivotTemporaryPivotTableArtifact temporaryPivot =
                plan.TemporaryPivotTable ??
                throw new InvalidOperationException(
                    "The deterministic artifact plan has no temporary target PivotTable receipt.");
            int temporaryPivotNameMatches = CountWorkbookPivotTablesByName(
                worksheets,
                temporaryPivot.Name);
            if (recoveryOwnership == null)
            {
                if (query != null || connection != null || workbookName != null ||
                    temporaryPivotNameMatches != 0 ||
                    plan.TemporaryWorksheets.Any(item =>
                        FindNamedObject(
                            worksheets,
                            item.Name,
                            "worksheets") != null))
                {
                    throw new InvalidOperationException(
                        "A generated PivotTable+ artifact name already exists without exact recovery ownership.");
                }

                return;
            }

            DemandExactPlanOwnership(recoveryOwnership, plan);
            if (temporaryPivotNameMatches != 0)
            {
                throw new InvalidOperationException(
                    "A generated recovery PivotTable already exists. Use explicit pending recovery rather than restarting classic enablement.");
            }
            if (query != null) DemandExactPlannedQuery(query, plan);
            if (connection != null) DemandExactPlannedConnection(connection, plan);
            if (workbookName != null) DemandExactPlannedWorkbookName(workbookName, plan);
            foreach (PivotTemporaryWorksheetArtifact temporary in plan.TemporaryWorksheets)
            {
                object? sheet = FindNamedObject(
                    worksheets,
                    temporary.Name,
                    "worksheets");
                if (sheet != null)
                {
                    DemandTemporaryWorksheetMarker((dynamic)sheet, temporary);
                }
            }
        }

        private static int CountWorkbookPivotTablesByName(
            object worksheets,
            string expectedName)
        {
            var count = 0;
            foreach (object worksheetObject in ReadCollection(
                         worksheets,
                         MaximumWorksheets,
                         "worksheets during temporary PivotTable preflight"))
            {
                dynamic worksheet = worksheetObject;
                object pivotTables = ReadRequired(
                    () => (object?)worksheet.PivotTables,
                    "Excel did not expose worksheet PivotTables during temporary-name preflight.");
                foreach (object pivotObject in ReadCollection(
                             pivotTables,
                             MaximumFields,
                             "worksheet PivotTables during temporary-name preflight"))
                {
                    dynamic pivot = pivotObject;
                    string name = ReadRequiredString(
                        () => (object?)pivot.Name,
                        "Excel exposed an unnamed PivotTable during temporary-name preflight.");
                    if (string.Equals(name, expectedName, StringComparison.Ordinal))
                    {
                        count++;
                        if (count > 1)
                        {
                            throw new InvalidOperationException(
                                "More than one PivotTable uses the deterministic recovery name.");
                        }
                    }
                }
            }

            return count;
        }

        public PivotDataModelArtifacts EnsureOwnedModelArtifacts(
            object workbook,
            PivotDataModelArtifactPlan plan,
            PivotPlusWorkbookMetadata ownership)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (plan == null) throw new ArgumentNullException(nameof(plan));
            if (ownership == null) throw new ArgumentNullException(nameof(ownership));
            DemandExactPlanOwnership(ownership, plan);

            dynamic book = workbook;
            object? nativeName = null;
            PivotOwnedWorkbookNameArtifact? ownedWorkbookName = null;
            if (plan.WorkbookName != null)
            {
                object names = ReadRequired(
                    () => (object?)book.Names,
                    "Excel did not expose workbook names while ensuring model artifacts.");
                nativeName = FindNamedObject(
                    names,
                    plan.WorkbookName,
                    "names",
                    workbookScopedName: true);
                if (nativeName == null)
                {
                    nativeName = book.Names.Add(
                        plan.WorkbookName,
                        plan.RequestedWorkbookNameReference,
                        false);
                }

                ownedWorkbookName = CreateOwnedWorkbookNameReceipt(
                    nativeName,
                    plan.WorkbookName,
                    plan.RequestedWorkbookNameReference ?? string.Empty,
                    plan.WorkbookNameFingerprint ?? string.Empty);
            }

            object queries = ReadRequired(
                () => (object?)book.Queries,
                "Excel did not expose workbook queries while ensuring model artifacts.");
            object? query = FindNamedObject(queries, plan.QueryName, "queries");
            if (query == null)
            {
                query = book.Queries.Add(plan.QueryName, plan.QueryFormula);
            }
            DemandExactPlannedQuery(query, plan);

            object connections = ReadRequired(
                () => (object?)book.Connections,
                "Excel did not expose workbook connections while ensuring model artifacts.");
            object? connection = FindNamedObject(
                connections,
                plan.ConnectionName,
                "connections");
            if (connection == null)
            {
                connection = book.Connections.Add2(
                    plan.ConnectionName,
                    "PivotTable+ generated workbook-only Data Model source.",
                    CanonicalConnectionContract.ConnectionString(plan.QueryName),
                    CanonicalConnectionContract.CommandText(plan.QueryName),
                    2,
                    true,
                    false);
            }

            DemandExactPlannedConnection(connection, plan);
            dynamic nativeConnection = connection;
            object oleDbObject = ReadRequired(
                () => (object?)nativeConnection.OLEDBConnection,
                "Excel did not expose the owned OLE DB connection.");
            dynamic oleDb = oleDbObject;
            oleDb.BackgroundQuery = false;
            if (ReadRequiredBoolean(
                    () => (object?)oleDb.BackgroundQuery,
                    "owned OLE DB BackgroundQuery state"))
            {
                throw new InvalidOperationException(
                    "Excel did not disable asynchronous refresh for the owned model source.");
            }

            nativeConnection.Refresh();
            object dataModelConnection = ReadExactDataModelConnection(
                workbook,
                connection,
                plan.ModelTableName);
            return new PivotDataModelArtifacts(
                plan.QueryName,
                plan.ConnectionName,
                plan.ModelTableName,
                plan.QueryFormula,
                plan.QueryFingerprint,
                plan.ConnectionFingerprint,
                connection,
                ownedWorkbookName,
                plan.TemporaryWorksheets,
                dataModelConnection,
                plan.TemporaryPivotTable);
        }

        public PivotDataModelArtifacts ValidatePendingDataModelFinalization(
            object workbook,
            object pivotTable,
            string setupId,
            PivotPlusWorkbookMetadata ownership)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            if (ownership == null) throw new ArgumentNullException(nameof(ownership));
            if (!string.Equals(
                    ownership.SetupId,
                    setupId,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "The pending ownership does not match the requested finalization setup.");
            }

            GeneratedNames names = GeneratedNames.For(setupId);
            PivotPlusOwnedArtifact queryReceipt = DemandSingleOwnedArtifact(
                ownership,
                PivotPlusArtifactKind.Query,
                names.QueryName);
            PivotPlusOwnedArtifact connectionReceipt = DemandSingleOwnedArtifact(
                ownership,
                PivotPlusArtifactKind.Connection,
                names.ConnectionName);
            IReadOnlyList<PivotTemporaryWorksheetArtifact> temporaryReceipts =
                ownership.Artifacts
                    .Where(item => item.Kind == PivotPlusArtifactKind.TemporaryWorksheet)
                    .Select(item => new PivotTemporaryWorksheetArtifact(
                        item.ArtifactId,
                        string.Equals(
                            item.ArtifactId,
                            names.StagingWorksheetName,
                            StringComparison.Ordinal)
                            ? "staging"
                            : "format-backup",
                        item.Fingerprint,
                        ownership.TargetAnchorAddress))
                    .ToList();
            if (temporaryReceipts.Count != 2 ||
                temporaryReceipts.All(item => !string.Equals(
                    item.Name,
                    names.StagingWorksheetName,
                    StringComparison.Ordinal)) ||
                temporaryReceipts.All(item => !string.Equals(
                    item.Name,
                    names.FormatBackupWorksheetName,
                    StringComparison.Ordinal)))
            {
                throw new InvalidOperationException(
                    "The pending conversion does not own the exact temporary worksheets required for finalization.");
            }
            PivotTemporaryPivotTableArtifact? temporaryPivotReceipt = null;
            List<PivotPlusOwnedArtifact> temporaryPivotReceipts = ownership.Artifacts
                .Where(item => item.Kind == PivotPlusArtifactKind.TemporaryPivotTable)
                .ToList();
            if (temporaryPivotReceipts.Count > 1)
            {
                throw new InvalidOperationException(
                    "The pending finalization has ambiguous temporary PivotTable ownership.");
            }
            if (temporaryPivotReceipts.Count == 1)
            {
                temporaryPivotReceipt = CreateTemporaryPivotTableReceipt(
                    setupId,
                    names,
                    new PivotTargetIdentity(
                        string.Empty,
                        ownership.TargetWorksheetName,
                        ownership.TargetPivotTableName),
                    ownership.TargetAnchorAddress);
                PivotPlusOwnedArtifact persistedTemporary = temporaryPivotReceipts[0];
                if (!string.Equals(
                        persistedTemporary.ArtifactId,
                        temporaryPivotReceipt.Name,
                        StringComparison.Ordinal) ||
                    !string.Equals(
                        persistedTemporary.Fingerprint,
                        temporaryPivotReceipt.Fingerprint,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "The pending temporary PivotTable receipt changed before finalization.");
                }
            }
            foreach (PivotTemporaryWorksheetArtifact temporary in temporaryReceipts)
            {
                PivotPlusOwnedArtifact persisted = DemandSingleOwnedArtifact(
                    ownership,
                    temporary.Kind,
                    temporary.Name);
                if (!string.Equals(
                        persisted.Fingerprint,
                        temporary.Fingerprint,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "The pending temporary worksheet receipt changed before finalization.");
                }
            }

            dynamic book = workbook;
            object query = FindNamedObject(
                    ReadRequired(
                        () => (object?)book.Queries,
                        "Excel did not expose workbook queries during finalization."),
                    names.QueryName,
                    "queries") ??
                throw new InvalidOperationException(
                    "The pending owned query is missing during finalization.");
            dynamic nativeQuery = query;
            string queryFormula = ReadRequiredString(
                () => (object?)nativeQuery.Formula,
                "Excel did not expose the pending owned query formula.");
            string sourceObjectName = DemandCanonicalGeneratedQuery(queryFormula);
            if (!PivotPlusFingerprint.Matches(
                    queryReceipt.Fingerprint,
                    "pivotplus.query.v1",
                    queryFormula))
            {
                throw new InvalidOperationException(
                    "The pending owned query changed before finalization.");
            }

            object connection = FindNamedObject(
                    ReadRequired(
                        () => (object?)book.Connections,
                        "Excel did not expose workbook connections during finalization."),
                    names.ConnectionName,
                    "connections") ??
                throw new InvalidOperationException(
                    "The pending owned connection is missing during finalization.");
            var artifacts = new PivotDataModelArtifacts(
                names.QueryName,
                names.ConnectionName,
                names.QueryName,
                queryFormula,
                queryReceipt.Fingerprint,
                connectionReceipt.Fingerprint,
                connection,
                temporaryWorksheets: temporaryReceipts,
                nativeDataModelConnection: ReadExactDataModelConnection(
                    workbook,
                    connection,
                    names.QueryName),
                temporaryPivotTable: temporaryPivotReceipt);
            DemandExactConnection(connection, artifacts);

            List<PivotPlusOwnedArtifact> nameReceipts = ownership.Artifacts
                .Where(item => item.Kind == PivotPlusArtifactKind.WorkbookName)
                .ToList();
            PivotOwnedWorkbookNameArtifact? ownedWorkbookName = null;
            if (nameReceipts.Count > 1)
            {
                throw new InvalidOperationException(
                    "The pending conversion has ambiguous workbook-name ownership.");
            }

            if (nameReceipts.Count == 1)
            {
                PivotPlusOwnedArtifact nameReceipt = nameReceipts[0];
                if (!string.Equals(
                        nameReceipt.ArtifactId,
                        names.SourceAliasName,
                        StringComparison.Ordinal) ||
                    !string.Equals(
                        sourceObjectName,
                        names.SourceAliasName,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "The pending raw-source alias does not match the owned query.");
                }

                object nativeName = FindNamedObject(
                        ReadRequired(
                            () => (object?)book.Names,
                            "Excel did not expose workbook names during finalization."),
                        names.SourceAliasName,
                        "names",
                        workbookScopedName: true) ??
                    throw new InvalidOperationException(
                        "The pending raw-source alias is missing during finalization.");
                dynamic name = nativeName;
                string actualReference = ReadRequiredString(
                    () => (object?)name.RefersTo,
                    "Excel did not expose the pending raw-source alias reference.");
                if (!string.Equals(
                        nameReceipt.Fingerprint,
                        WorkbookNameFingerprint(actualReference),
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "The pending raw-source alias changed before finalization.");
                }

                ownedWorkbookName = new PivotOwnedWorkbookNameArtifact(
                    names.SourceAliasName,
                    nameReceipt.Fingerprint,
                    actualReference,
                    nativeName);
                artifacts = new PivotDataModelArtifacts(
                    names.QueryName,
                    names.ConnectionName,
                    names.QueryName,
                    queryFormula,
                    queryReceipt.Fingerprint,
                    connectionReceipt.Fingerprint,
                    connection,
                    ownedWorkbookName,
                    temporaryReceipts,
                    artifacts.NativeDataModelConnection,
                    temporaryPivotReceipt);
            }

            int expectedArtifactCount = 2 + temporaryReceipts.Count +
                                        (temporaryPivotReceipt == null ? 0 : 1) +
                                        (ownedWorkbookName == null ? 0 : 1);
            if (ownership.Artifacts.Count != expectedArtifactCount)
            {
                throw new InvalidOperationException(
                    "The pending finalization ownership contains unexpected artifacts.");
            }

            object worksheets = ReadRequired(
                () => (object?)book.Worksheets,
                "Excel did not expose workbook worksheets during finalization.");
            foreach (PivotTemporaryWorksheetArtifact temporary in temporaryReceipts)
            {
                object? temporarySheet = FindNamedObject(
                    worksheets,
                    temporary.Name,
                    "worksheets");
                if (temporarySheet != null)
                {
                    DemandTemporaryWorksheetMarker(
                        (dynamic)temporarySheet,
                        temporary);
                    throw new InvalidOperationException(
                        "A pending temporary worksheet still exists; finalization will not hide incomplete cleanup.");
                }
            }

            dynamic pivot = pivotTable;
            dynamic cache = ReadPivotCache(pivotTable);
            if (!ReadBoolean(
                    () => (object?)cache.OLAP,
                    "PivotCache.OLAP during finalization"))
            {
                throw new InvalidOperationException(
                    "The finalization target is not the committed Data Model PivotTable.");
            }

            object pivotConnection = ReadRequired(
                () => (object?)cache.WorkbookConnection,
                "Excel did not expose the finalization target's Data Model connection.");
            if (!SameNativeObject(
                    pivotConnection,
                    artifacts.NativeDataModelConnection))
            {
                throw new InvalidOperationException(
                    "The finalization target does not use the exact owned Data Model connection.");
            }

            return artifacts;
        }

        public PivotStagedDataModelPivot CreateStagedDataModelPivot(
            object workbook,
            string setupId,
            PivotDataModelArtifacts artifacts)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (artifacts == null) throw new ArgumentNullException(nameof(artifacts));

            GeneratedNames names = GeneratedNames.For(setupId);
            PivotTemporaryWorksheetArtifact stagingReceipt =
                FindTemporaryWorksheetReceipt(
                    artifacts,
                    names.StagingWorksheetName,
                    "staging");
            PivotTemporaryWorksheetArtifact formatReceipt =
                FindTemporaryWorksheetReceipt(
                    artifacts,
                    names.FormatBackupWorksheetName,
                    "format-backup");
            dynamic book = workbook;
            ReconcileStaleTemporaryWorksheet(
                workbook,
                stagingReceipt,
                names.StagingPivotTableName,
                isFormatBackup: false,
                artifacts.NativeDataModelConnection);
            DemandCollectionNameAvailable(
                ReadRequired(
                    () => (object?)book.Worksheets,
                    "Excel did not expose workbook worksheets."),
                "worksheet",
                names.StagingWorksheetName);
            object? worksheet = null;
            try
            {
                worksheet = book.Worksheets.Add();
                dynamic stagingSheet = worksheet;
                stagingSheet.Name = names.StagingWorksheetName;
                stagingSheet.Visible = SheetVeryHidden;
                WriteTemporaryWorksheetMarker(stagingSheet, stagingReceipt);
                object cacheCollection = ReadRequired(
                    () => (object?)book.PivotCaches(),
                    "Excel did not expose workbook PivotCaches safely.");
                object? reusableCache = FindReusableDataModelCache(
                    cacheCollection,
                    artifacts.NativeDataModelConnection);
                dynamic nativeCaches = cacheCollection;
                dynamic cache = reusableCache ?? nativeCaches.Create(
                    SourceExternal,
                    artifacts.NativeDataModelConnection,
                    PivotVersion15);
                dynamic destination = stagingSheet.Range["A1"];
                object pivot = cache.CreatePivotTable(destination, names.StagingPivotTableName);
                return new PivotStagedDataModelPivot(
                    names.StagingWorksheetName,
                    names.StagingPivotTableName,
                    worksheet,
                    pivot,
                    cache,
                    stagingReceipt,
                    formatReceipt,
                    artifacts.TemporaryPivotTable ??
                        throw new InvalidOperationException(
                            "Pending ownership has no temporary target PivotTable receipt."));
            }
            catch (Exception failure)
            {
                if (worksheet == null) throw;
                try
                {
                    DeleteWorksheet((dynamic)worksheet);
                }
                catch (Exception cleanupFailure)
                {
                    throw new AggregateException(
                        "PivotTable+ staging PivotTable creation failed and its worksheet cleanup did not complete.",
                        failure,
                        cleanupFailure);
                }

                throw;
            }
        }

        public void RestoreState(
            object pivotTable,
            PivotNativeStateSnapshot snapshot,
            string modelTableName)
        {
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            LateBoundPivotState state = ReadNativeState(snapshot);
            RestoreDataModelState(pivotTable, state, modelTableName);
        }

        public void RefreshPivotTable(object pivotTable)
        {
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            dynamic pivot = pivotTable;
            TryWrite(() => pivot.ManualUpdate = false);
            pivot.RefreshTable();
        }

        public string VerifyDataModelState(
            object pivotTable,
            PivotNativeStateSnapshot expectedState)
        {
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            dynamic cache = ReadPivotCache(pivotTable);
            if (!ReadBoolean(() => (object?)cache.OLAP, "PivotCache.OLAP"))
            {
                throw new InvalidOperationException(
                    "Excel did not create a native Data Model PivotTable.");
            }

            object connection = ReadRequired(
                () => (object?)cache.WorkbookConnection,
                "Excel did not expose the staged PivotTable workbook connection.");
            dynamic workbookConnection = connection;
            if (ReadOptionalInt(() => (object?)workbookConnection.Type, -1) != 7)
            {
                throw new InvalidOperationException(
                    "Excel created an OLAP PivotTable, but it is not backed by this workbook's Data Model.");
            }

            LateBoundPivotState expected = ReadNativeState(expectedState);
            LateBoundPivotState actual = ReadLivePivotState(pivotTable);
            if (!LateBoundPivotState.SemanticallyEquals(
                    expected,
                    actual,
                    normalizeDataModelInvariants: true))
            {
                throw new InvalidOperationException(
                    "The staged Data Model PivotTable did not preserve the original layout, style, or member filters.");
            }

            return FingerprintDataModelState(actual);
        }

        public void MarkStagingVerified(
            PivotStagedDataModelPivot stagedPivot,
            string stagingStateFingerprint)
        {
            if (stagedPivot == null) throw new ArgumentNullException(nameof(stagedPivot));
            PivotPlusMetadataValidator.ValidateFingerprint(
                stagingStateFingerprint,
                "staging state fingerprint");
            PivotTemporaryWorksheetArtifact receipt =
                stagedPivot.StagingWorksheet ??
                throw new InvalidOperationException(
                    "The staging worksheet has no durable ownership receipt.");
            dynamic sheet = stagedPivot.NativeWorksheet;
            DemandTemporaryWorksheetMarker(sheet, receipt);
            Dictionary<string, string> markers =
                ReadTemporaryWorksheetMarkers(sheet);
            if (markers.TryGetValue(
                    StagingStateFingerprintProperty,
                    out string? existing))
            {
                if (!string.Equals(
                        existing,
                        stagingStateFingerprint,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "The staging worksheet verification marker changed.");
                }

                return;
            }

            object properties = ReadRequired(
                () => (object?)sheet.CustomProperties,
                "Excel did not expose staging worksheet CustomProperties.");
            try
            {
                ((dynamic)properties).Add(
                    StagingStateFingerprintProperty,
                    stagingStateFingerprint);
            }
            catch (Exception addFailure)
            {
                Dictionary<string, string> afterFailure =
                    ReadTemporaryWorksheetMarkers(sheet);
                if (!afterFailure.TryGetValue(
                        StagingStateFingerprintProperty,
                        out string? inserted) ||
                    !string.Equals(
                        inserted,
                        stagingStateFingerprint,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "Excel could not durably mark the verified staging PivotTable.",
                        addFailure);
                }
            }

            DemandVerifiedStagingMarker(
                sheet,
                receipt,
                stagingStateFingerprint);
        }

        public IPivotReplacementTransaction PrepareReplacement(
            object workbook,
            object originalPivotTable,
            PivotStagedDataModelPivot stagedPivot,
            PivotNativeStateSnapshot originalState,
            string modelTableName)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (originalPivotTable == null) throw new ArgumentNullException(nameof(originalPivotTable));
            if (stagedPivot == null) throw new ArgumentNullException(nameof(stagedPivot));
            PivotTemporaryWorksheetArtifact formatReceipt =
                stagedPivot.FormatBackupWorksheet ??
                throw new InvalidOperationException(
                    "The staged conversion has no durable format-backup ownership receipt.");
            PivotTemporaryWorksheetArtifact stagingReceipt =
                stagedPivot.StagingWorksheet ??
                throw new InvalidOperationException(
                    "The staged conversion has no durable staging ownership receipt.");
            PivotTemporaryPivotTableArtifact temporaryPivotReceipt =
                stagedPivot.TemporaryTargetPivotTable ??
                throw new InvalidOperationException(
                    "The staged conversion has no durable temporary target PivotTable receipt.");
            Dictionary<string, string> stagingMarkers =
                ReadTemporaryWorksheetMarkers((dynamic)stagedPivot.NativeWorksheet);
            if (!stagingMarkers.TryGetValue(
                    StagingStateFingerprintProperty,
                    out string? verifiedStateFingerprint))
            {
                throw new InvalidOperationException(
                    "The staged conversion was not durably verified before replacement preparation.");
            }

            DemandVerifiedStagingMarker(
                (dynamic)stagedPivot.NativeWorksheet,
                stagingReceipt,
                verifiedStateFingerprint);
            LateBoundPivotState state = ReadNativeState(originalState);
            PivotFormatBackup? formatBackup = null;
            try
            {
                formatBackup = PivotFormatBackup.Create(
                    workbook,
                    originalPivotTable,
                    state.Result,
                    formatReceipt);
                return new LateBoundPivotReplacementTransaction(
                    this,
                    workbook,
                    originalPivotTable,
                    stagedPivot.NativePivotCache,
                    originalState,
                    modelTableName,
                    formatBackup,
                    temporaryPivotReceipt,
                    verifiedStateFingerprint);
            }
            catch
            {
                formatBackup?.Delete();
                throw;
            }
        }

        public void DeleteStagingPivot(object workbook, PivotStagedDataModelPivot stagedPivot)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (stagedPivot == null) throw new ArgumentNullException(nameof(stagedPivot));
            PivotTemporaryWorksheetArtifact receipt = stagedPivot.StagingWorksheet ??
                throw new InvalidOperationException(
                    "The staging worksheet has no durable ownership receipt.");
            dynamic sheet = stagedPivot.NativeWorksheet;
            string currentName = ReadRequiredString(
                () => (object?)sheet.Name,
                "Excel did not expose the staging worksheet name.");
            if (!string.Equals(currentName, stagedPivot.WorksheetName, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The generated staging worksheet was renamed; PivotTable+ will not delete it.");
            }

            DemandTemporaryWorksheetMarker(sheet, receipt);
            DemandTemporaryWorksheetStructure(
                sheet,
                stagedPivot.PivotTableName,
                isFormatBackup: false,
                allowIncomplete: false,
                expectedModelConnection: ReadRequired(
                    () => (object?)((dynamic)stagedPivot.NativePivotCache).WorkbookConnection,
                    "Excel did not expose the staged PivotCache model connection."));
            DeleteOwnedTemporaryWorksheet(workbook, sheet, receipt);
        }

        public void DeleteOwnedModelArtifacts(
            object workbook,
            PivotDataModelArtifacts artifacts)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (artifacts == null) throw new ArgumentNullException(nameof(artifacts));
            dynamic book = workbook;

            object? connection = FindNamedObject(
                ReadRequired(
                    () => (object?)book.Connections,
                    "Excel did not expose workbook connections during cleanup."),
                artifacts.ConnectionName,
                "connections");
            if (connection != null)
            {
                DemandExactConnection(connection, artifacts);
            }

            object? query = FindNamedObject(
                ReadRequired(
                    () => (object?)book.Queries,
                    "Excel did not expose workbook queries during cleanup."),
                artifacts.QueryName,
                "queries");
            if (query != null)
            {
                dynamic nativeQuery = query;
                string formula = ReadRequiredString(
                    () => (object?)nativeQuery.Formula,
                    "Excel did not expose the generated query formula.");
                if (!PivotPlusFingerprint.Matches(
                        artifacts.QueryFingerprint,
                        "pivotplus.query.v1",
                        formula) ||
                    !string.Equals(formula, artifacts.QueryFormula, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "The generated query changed; PivotTable+ will not delete it.");
                }
            }

            object? ownedName = null;
            if (artifacts.OwnedWorkbookName != null)
            {
                PivotOwnedWorkbookNameArtifact receipt = artifacts.OwnedWorkbookName;
                ownedName = FindNamedObject(
                    ReadRequired(
                        () => (object?)book.Names,
                        "Excel did not expose workbook names during cleanup."),
                    receipt.Name,
                    "names",
                    workbookScopedName: true);
                if (ownedName != null)
                {
                    DemandExactOwnedWorkbookName(ownedName, receipt);
                }
            }

            // Validate every surviving object before deleting either one. If a
            // user changed an artifact, cleanup is all-or-nothing at this
            // ownership boundary.
            var failures = new List<Exception>();
            DeleteValidatedArtifact(connection, "connection", failures);
            DeleteValidatedArtifact(query, "query", failures);
            DeleteValidatedArtifact(ownedName, "workbook name", failures);
            if (failures.Count > 0)
            {
                throw new AggregateException(
                    "PivotTable+ validated generated artifacts but one or more deletions failed.",
                    failures);
            }
        }

        public PivotPendingDataModelRecovery RecoverPending(
            object workbook,
            string setupId,
            PivotPlusWorkbookMetadata ownership)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (ownership == null) throw new ArgumentNullException(nameof(ownership));
            if (!string.Equals(
                    ownership.SetupId,
                    setupId,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "The pending recovery ownership does not match the requested setup.");
            }

            if (!string.Equals(
                    ownership.SchemaVersion,
                    PivotPlusWorkbookMetadata.CurrentSchemaVersion,
                    StringComparison.Ordinal))
            {
                throw new NotSupportedException(
                    "Legacy pending conversion metadata cannot drive forward recovery. Re-run enablement only while the exact classic PivotTable survives.");
            }

            PivotPlusMetadataValidator.ValidateLocalA1Address(
                ownership.TargetAnchorAddress);
            GeneratedNames names = GeneratedNames.For(setupId);
            dynamic book = workbook;
            object worksheets = ReadRequired(
                () => (object?)book.Worksheets,
                "Excel did not expose workbook worksheets during pending recovery.");
            object targetWorksheet = FindNamedObject(
                    worksheets,
                    ownership.TargetWorksheetName,
                    "worksheets") ??
                throw new InvalidOperationException(
                    "The pending recovery target worksheet is missing.");
            dynamic targetSheet = targetWorksheet;
            object pivotTables = ReadRequired(
                () => (object?)targetSheet.PivotTables,
                "Excel did not expose target worksheet PivotTables during recovery.");
            object? originalTarget = FindNamedObject(
                pivotTables,
                ownership.TargetPivotTableName,
                "PivotTables");
            object? temporaryTarget = FindNamedObject(
                pivotTables,
                names.ReplacementPivotTableName,
                "PivotTables");
            if (originalTarget != null && temporaryTarget != null)
            {
                throw new InvalidOperationException(
                    "Both the original and generated recovery PivotTable names exist; PivotTable+ will not clear either one.");
            }

            if (ownership.RecoveryPhase == PivotPlusRecoveryPhase.Planned)
            {
                if (temporaryTarget != null || originalTarget == null)
                {
                    throw new InvalidOperationException(
                        "A planned recovery has no verified staging checkpoint and the exact classic target is not intact.");
                }

                DemandExactClassicTarget(
                    originalTarget,
                    targetWorksheet,
                    ownership.TargetPivotTableName,
                    ownership.TargetAnchorAddress);
                throw new InvalidOperationException(
                    "The exact classic PivotTable survived. Re-run Data Model enablement to rebuild its planned artifacts safely.");
            }

            if (ownership.RecoveryPhase != PivotPlusRecoveryPhase.StagingVerified)
            {
                throw new InvalidOperationException(
                    "The setup has no staging-verified recovery checkpoint.");
            }

            PivotTargetIdentity target = new PivotTargetIdentity(
                new StoredWorkbookIdentityResolver().Resolve(workbook),
                ownership.TargetWorksheetName,
                ownership.TargetPivotTableName);
            PivotDataModelArtifacts artifacts =
                RehydratePendingModelArtifacts(
                    workbook,
                    setupId,
                    target,
                    ownership,
                    names);
            PivotTemporaryWorksheetArtifact stagingReceipt =
                FindTemporaryWorksheetReceipt(
                    artifacts,
                    names.StagingWorksheetName,
                    "staging");
            PivotTemporaryWorksheetArtifact formatReceipt =
                FindTemporaryWorksheetReceipt(
                    artifacts,
                    names.FormatBackupWorksheetName,
                    "format-backup");
            PivotTemporaryPivotTableArtifact temporaryReceipt =
                artifacts.TemporaryPivotTable ??
                throw new InvalidOperationException(
                    "The pending recovery has no deterministic temporary PivotTable receipt.");

            object? stagingSheetObject = FindNamedObject(
                worksheets,
                stagingReceipt.Name,
                "worksheets");
            PivotStagedDataModelPivot? staging = null;
            LateBoundPivotState? stagingState = null;
            if (stagingSheetObject != null)
            {
                dynamic stagingSheet = stagingSheetObject;
                DemandVerifiedStagingMarker(
                    stagingSheet,
                    stagingReceipt,
                    ownership.StagingStateFingerprint);
                DemandTemporaryWorksheetStructure(
                    stagingSheet,
                    names.StagingPivotTableName,
                    isFormatBackup: false,
                    allowIncomplete: false,
                    expectedModelConnection: artifacts.NativeDataModelConnection);
                object stagedPivot = FindNamedObject(
                        ReadRequired(
                            () => (object?)stagingSheet.PivotTables,
                            "Excel did not expose the staged recovery PivotTable collection."),
                        names.StagingPivotTableName,
                        "PivotTables") ??
                    throw new InvalidOperationException(
                        "The verified staging PivotTable is missing.");
                DemandStateFingerprint(
                    stagedPivot,
                    ownership.StagingStateFingerprint);
                stagingState = ReadLivePivotState(stagedPivot);
                staging = new PivotStagedDataModelPivot(
                    stagingReceipt.Name,
                    names.StagingPivotTableName,
                    stagingSheetObject,
                    stagedPivot,
                    ReadPivotCache(stagedPivot),
                    stagingReceipt,
                    formatReceipt,
                    temporaryReceipt);
            }

            object? formatSheetObject = FindNamedObject(
                worksheets,
                formatReceipt.Name,
                "worksheets");

            if (originalTarget != null)
            {
                dynamic cache = ReadPivotCache(originalTarget);
                bool isModel = ReadBoolean(
                    () => (object?)cache.OLAP,
                    "pending target PivotCache.OLAP");
                if (!isModel)
                {
                    if (temporaryTarget != null)
                    {
                        throw new InvalidOperationException(
                            "A generated recovery target exists beside the surviving classic PivotTable.");
                    }

                    DemandExactClassicTarget(
                        originalTarget,
                        targetWorksheet,
                        ownership.TargetPivotTableName,
                        ownership.TargetAnchorAddress);
                    if (staging != null)
                    {
                        DemandStateFingerprint(
                            staging.NativePivotTable,
                            ownership.StagingStateFingerprint);
                    }

                    throw new InvalidOperationException(
                        "The exact classic PivotTable survived. Re-run Data Model enablement; forward recovery will not clear it.");
                }

                DemandExactDataModelTarget(
                    originalTarget,
                    targetWorksheet,
                    ownership.TargetPivotTableName,
                    ownership.TargetAnchorAddress,
                    artifacts.NativeDataModelConnection);
                DemandStateFingerprint(
                    originalTarget,
                    ownership.StagingStateFingerprint);
                LateBoundPivotState finalState = ReadLivePivotState(originalTarget);
                if (formatSheetObject != null)
                {
                    PivotFormatBackup backup = PivotFormatBackup.OpenExisting(
                        workbook,
                        formatSheetObject,
                        finalState.Result,
                        formatReceipt);
                    backup.Restore(originalTarget, finalState.Result);
                    DemandStateFingerprint(
                        originalTarget,
                        ownership.StagingStateFingerprint);
                    backup.Delete();
                }

                if (staging != null)
                {
                    DeleteStagingPivot(workbook, staging);
                }

                return new PivotPendingDataModelRecovery(target, artifacts);
            }

            if (staging == null || stagingState == null || formatSheetObject == null)
            {
                throw new InvalidOperationException(
                    "An absent or partial recovery target requires both the exact verified staging PivotTable and format backup.");
            }

            PivotFormatBackup formatBackup = PivotFormatBackup.OpenExisting(
                workbook,
                formatSheetObject,
                stagingState.Result,
                formatReceipt);
            if (temporaryTarget != null)
            {
                DemandExactTemporaryTargetPivot(
                    temporaryTarget,
                    targetWorksheet,
                    temporaryReceipt,
                    artifacts.NativeDataModelConnection);
                ((dynamic)temporaryTarget).TableRange2.Clear();
                temporaryTarget = null;
            }

            DemandDestinationRectangleEmpty(
                targetWorksheet,
                ownership.TargetAnchorAddress,
                stagingState.Result,
                ownership.TargetPivotTableName,
                temporaryReceipt.Name);
            object recoveredTarget = CreateRecoveredTarget(
                targetWorksheet,
                staging,
                stagingState,
                artifacts,
                temporaryReceipt,
                ownership.StagingStateFingerprint,
                formatBackup);
            DemandExactDataModelTarget(
                recoveredTarget,
                targetWorksheet,
                ownership.TargetPivotTableName,
                ownership.TargetAnchorAddress,
                artifacts.NativeDataModelConnection);
            DemandStateFingerprint(
                recoveredTarget,
                ownership.StagingStateFingerprint);

            // The verified final target is now the durable semantic copy.
            // Delete formatting first, staging last, then let the service
            // atomically transition ownership to Active.
            formatBackup.Delete();
            DeleteStagingPivot(workbook, staging);
            return new PivotPendingDataModelRecovery(target, artifacts);
        }

        public void VerifyActiveDataModelOwnership(
            object workbook,
            string setupId,
            PivotPlusWorkbookMetadata ownership)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (ownership == null) throw new ArgumentNullException(nameof(ownership));
            if (!string.Equals(
                    ownership.SetupId,
                    setupId,
                    StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(
                    ownership.SchemaVersion,
                    PivotPlusWorkbookMetadata.CurrentSchemaVersion,
                    StringComparison.Ordinal) ||
                ownership.RecoveryPhase != PivotPlusRecoveryPhase.None ||
                ownership.Artifacts.Any(item =>
                    item.Kind == PivotPlusArtifactKind.TemporaryWorksheet ||
                    item.Kind == PivotPlusArtifactKind.TemporaryPivotTable))
            {
                throw new InvalidOperationException(
                    "The setup is not an exact Active Data Model ownership receipt.");
            }

            GeneratedNames names = GeneratedNames.For(setupId);
            PivotPlusOwnedArtifact queryReceipt = DemandSingleOwnedArtifact(
                ownership,
                PivotPlusArtifactKind.Query,
                names.QueryName);
            PivotPlusOwnedArtifact connectionReceipt = DemandSingleOwnedArtifact(
                ownership,
                PivotPlusArtifactKind.Connection,
                names.ConnectionName);
            dynamic book = workbook;
            object query = FindNamedObject(
                    ReadRequired(
                        () => (object?)book.Queries,
                        "Excel did not expose workbook queries while verifying Active ownership."),
                    names.QueryName,
                    "queries") ??
                throw new InvalidOperationException(
                    "The exact Active owned query is missing.");
            dynamic nativeQuery = query;
            string queryFormula = ReadRequiredString(
                () => (object?)nativeQuery.Formula,
                "Excel did not expose the Active owned query formula.");
            string sourceObjectName = DemandCanonicalGeneratedQuery(queryFormula);
            if (!PivotPlusFingerprint.Matches(
                    queryReceipt.Fingerprint,
                    "pivotplus.query.v1",
                    queryFormula))
            {
                throw new InvalidOperationException(
                    "The Active owned query changed.");
            }

            object connection = FindNamedObject(
                    ReadRequired(
                        () => (object?)book.Connections,
                        "Excel did not expose workbook connections while verifying Active ownership."),
                    names.ConnectionName,
                    "connections") ??
                throw new InvalidOperationException(
                    "The exact Active owned connection is missing.");
            var artifacts = new PivotDataModelArtifacts(
                names.QueryName,
                names.ConnectionName,
                names.QueryName,
                queryFormula,
                queryReceipt.Fingerprint,
                connectionReceipt.Fingerprint,
                connection,
                nativeDataModelConnection: ReadExactDataModelConnection(
                    workbook,
                    connection,
                    names.QueryName));
            DemandExactConnection(connection, artifacts);

            List<PivotPlusOwnedArtifact> nameReceipts = ownership.Artifacts
                .Where(item => item.Kind == PivotPlusArtifactKind.WorkbookName)
                .ToList();
            if (nameReceipts.Count > 1)
            {
                throw new InvalidOperationException(
                    "The Active setup has ambiguous workbook-name ownership.");
            }

            if (nameReceipts.Count == 1)
            {
                PivotPlusOwnedArtifact nameReceipt = nameReceipts[0];
                if (!string.Equals(
                        nameReceipt.ArtifactId,
                        names.SourceAliasName,
                        StringComparison.Ordinal) ||
                    !string.Equals(
                        sourceObjectName,
                        names.SourceAliasName,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "The Active raw-source alias does not match the exact owned query.");
                }

                object nativeName = FindNamedObject(
                        ReadRequired(
                            () => (object?)book.Names,
                            "Excel did not expose workbook names while verifying Active ownership."),
                        names.SourceAliasName,
                        "names",
                        workbookScopedName: true) ??
                    throw new InvalidOperationException(
                        "The exact Active raw-source alias is missing.");
                dynamic name = nativeName;
                string actualReference = ReadRequiredString(
                    () => (object?)name.RefersTo,
                    "Excel did not expose the Active raw-source alias reference.");
                if (!string.Equals(
                        nameReceipt.Fingerprint,
                        WorkbookNameFingerprint(actualReference),
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "The Active raw-source alias changed.");
                }
            }
            else if (string.Equals(
                         sourceObjectName,
                         names.SourceAliasName,
                         StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The Active query references an unowned generated source alias.");
            }

            if (ownership.Artifacts.Count != 2 + nameReceipts.Count)
            {
                throw new InvalidOperationException(
                    "The Active ownership contains unexpected artifacts.");
            }

            object worksheets = ReadRequired(
                () => (object?)book.Worksheets,
                "Excel did not expose workbook worksheets while verifying Active ownership.");
            foreach (string temporarySheetName in new[]
                     {
                         names.StagingWorksheetName,
                         names.FormatBackupWorksheetName
                     })
            {
                if (FindNamedObject(worksheets, temporarySheetName, "worksheets") != null)
                {
                    throw new InvalidOperationException(
                        "An Active setup still has a generated temporary worksheet.");
                }
            }

            if (CountWorkbookPivotTablesByName(
                    worksheets,
                    names.ReplacementPivotTableName) != 0)
            {
                throw new InvalidOperationException(
                    "An Active setup still has a generated temporary target PivotTable.");
            }

            object targetWorksheet = FindNamedObject(
                    worksheets,
                    ownership.TargetWorksheetName,
                    "worksheets") ??
                throw new InvalidOperationException(
                    "The Active target worksheet is missing.");
            dynamic targetSheet = targetWorksheet;
            object targetPivot = FindNamedObject(
                    ReadRequired(
                        () => (object?)targetSheet.PivotTables,
                        "Excel did not expose the Active target PivotTable collection."),
                    ownership.TargetPivotTableName,
                    "PivotTables") ??
                throw new InvalidOperationException(
                    "The Active target PivotTable is missing.");
            dynamic target = targetPivot;
            object parent = ReadRequired(
                () => (object?)target.Parent,
                "Excel did not expose the Active target worksheet.");
            if (!SameNativeObject(parent, targetWorksheet))
            {
                throw new InvalidOperationException(
                    "The Active target PivotTable moved to another worksheet.");
            }

            dynamic cache = ReadPivotCache(targetPivot);
            if (!ReadBoolean(
                    () => (object?)cache.OLAP,
                    "Active target PivotCache.OLAP"))
            {
                throw new InvalidOperationException(
                    "The Active target is no longer a native Data Model PivotTable.");
            }

            object targetConnection = ReadRequired(
                () => (object?)cache.WorkbookConnection,
                "Excel did not expose the Active target Data Model connection.");
            if (!SameNativeObject(
                    targetConnection,
                    artifacts.NativeDataModelConnection))
            {
                throw new InvalidOperationException(
                    "The Active target no longer uses the exact workbook Data Model connection.");
            }
        }

        private static PivotDataModelArtifacts RehydratePendingModelArtifacts(
            object workbook,
            string setupId,
            PivotTargetIdentity target,
            PivotPlusWorkbookMetadata ownership,
            GeneratedNames names)
        {
            PivotPlusOwnedArtifact queryReceipt = DemandSingleOwnedArtifact(
                ownership,
                PivotPlusArtifactKind.Query,
                names.QueryName);
            PivotPlusOwnedArtifact connectionReceipt = DemandSingleOwnedArtifact(
                ownership,
                PivotPlusArtifactKind.Connection,
                names.ConnectionName);
            IReadOnlyList<PivotTemporaryWorksheetArtifact> temporaryReceipts =
                CreateTemporaryWorksheetReceipts(
                    names,
                    ownership.TargetAnchorAddress);
            foreach (PivotTemporaryWorksheetArtifact receipt in temporaryReceipts)
            {
                PivotPlusOwnedArtifact persisted = DemandSingleOwnedArtifact(
                    ownership,
                    receipt.Kind,
                    receipt.Name);
                if (!string.Equals(
                        persisted.Fingerprint,
                        receipt.Fingerprint,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "A temporary worksheet recovery receipt changed.");
                }
            }

            PivotTemporaryPivotTableArtifact temporaryPivot =
                CreateTemporaryPivotTableReceipt(
                    setupId,
                    names,
                    target,
                    ownership.TargetAnchorAddress);
            PivotPlusOwnedArtifact persistedTemporaryPivot =
                DemandSingleOwnedArtifact(
                    ownership,
                    temporaryPivot.Kind,
                    temporaryPivot.Name);
            if (!string.Equals(
                    persistedTemporaryPivot.Fingerprint,
                    temporaryPivot.Fingerprint,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The generated target PivotTable recovery receipt changed.");
            }

            dynamic book = workbook;
            object query = FindNamedObject(
                    ReadRequired(
                        () => (object?)book.Queries,
                        "Excel did not expose workbook queries during recovery."),
                    names.QueryName,
                    "queries") ??
                throw new InvalidOperationException(
                    "The exact owned recovery query is missing.");
            dynamic nativeQuery = query;
            string queryName = ReadRequiredString(
                () => (object?)nativeQuery.Name,
                "Excel did not expose the recovery query name.");
            string queryFormula = ReadRequiredString(
                () => (object?)nativeQuery.Formula,
                "Excel did not expose the recovery query formula.");
            if (!string.Equals(queryName, names.QueryName, StringComparison.Ordinal) ||
                !PivotPlusFingerprint.Matches(
                    queryReceipt.Fingerprint,
                    "pivotplus.query.v1",
                    queryFormula))
            {
                throw new InvalidOperationException(
                    "The exact owned recovery query changed.");
            }

            string sourceObjectName = DemandCanonicalGeneratedQuery(queryFormula);
            object sourceConnection = FindNamedObject(
                    ReadRequired(
                        () => (object?)book.Connections,
                        "Excel did not expose workbook connections during recovery."),
                    names.ConnectionName,
                    "connections") ??
                throw new InvalidOperationException(
                    "The exact owned recovery source connection is missing.");

            List<PivotPlusOwnedArtifact> nameReceipts = ownership.Artifacts
                .Where(item => item.Kind == PivotPlusArtifactKind.WorkbookName)
                .ToList();
            if (nameReceipts.Count > 1)
            {
                throw new InvalidOperationException(
                    "The pending recovery has ambiguous workbook-name ownership.");
            }

            PivotOwnedWorkbookNameArtifact? ownedWorkbookName = null;
            if (nameReceipts.Count == 1)
            {
                PivotPlusOwnedArtifact nameReceipt = nameReceipts[0];
                if (!string.Equals(
                        nameReceipt.ArtifactId,
                        names.SourceAliasName,
                        StringComparison.Ordinal) ||
                    !string.Equals(
                        sourceObjectName,
                        names.SourceAliasName,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "The pending raw-range source alias does not match the exact owned query.");
                }

                object nativeName = FindNamedObject(
                        ReadRequired(
                            () => (object?)book.Names,
                            "Excel did not expose workbook names during recovery."),
                        names.SourceAliasName,
                        "names",
                        workbookScopedName: true) ??
                    throw new InvalidOperationException(
                        "The exact owned raw-range source alias is missing.");
                dynamic name = nativeName;
                string actualReference = ReadRequiredString(
                    () => (object?)name.RefersTo,
                    "Excel did not expose the recovery source alias reference.");
                string actualFingerprint = WorkbookNameFingerprint(actualReference);
                if (!string.Equals(
                        nameReceipt.Fingerprint,
                        actualFingerprint,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "The exact owned raw-range source alias changed.");
                }

                ownedWorkbookName = new PivotOwnedWorkbookNameArtifact(
                    names.SourceAliasName,
                    actualFingerprint,
                    actualReference,
                    nativeName);
            }
            else if (string.Equals(
                         sourceObjectName,
                         names.SourceAliasName,
                         StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The recovery query references an unowned generated source alias.");
            }

            int expectedArtifactCount = 2 + temporaryReceipts.Count + 1 +
                                        (ownedWorkbookName == null ? 0 : 1);
            if (ownership.Artifacts.Count != expectedArtifactCount)
            {
                throw new InvalidOperationException(
                    "The pending recovery ownership contains unexpected artifacts.");
            }

            var provisional = new PivotDataModelArtifacts(
                names.QueryName,
                names.ConnectionName,
                names.QueryName,
                queryFormula,
                queryReceipt.Fingerprint,
                connectionReceipt.Fingerprint,
                sourceConnection,
                ownedWorkbookName,
                temporaryReceipts,
                sourceConnection,
                temporaryPivot);
            DemandExactConnection(sourceConnection, provisional);
            object modelConnection = ReadExactDataModelConnection(
                workbook,
                sourceConnection,
                names.QueryName);
            return new PivotDataModelArtifacts(
                names.QueryName,
                names.ConnectionName,
                names.QueryName,
                queryFormula,
                queryReceipt.Fingerprint,
                connectionReceipt.Fingerprint,
                sourceConnection,
                ownedWorkbookName,
                temporaryReceipts,
                modelConnection,
                temporaryPivot);
        }

        private static void DemandExactClassicTarget(
            object pivotTable,
            object worksheet,
            string expectedName,
            string expectedAnchor)
        {
            DemandExactPivotIdentity(
                pivotTable,
                worksheet,
                expectedName,
                expectedAnchor);
            dynamic cache = ReadPivotCache(pivotTable);
            if (ReadBoolean(
                    () => (object?)cache.OLAP,
                    "classic recovery target PivotCache.OLAP"))
            {
                throw new InvalidOperationException(
                    "The pending target is no longer the exact classic PivotTable.");
            }
        }

        private static void DemandExactDataModelTarget(
            object pivotTable,
            object worksheet,
            string expectedName,
            string expectedAnchor,
            object expectedModelConnection)
        {
            DemandExactPivotIdentity(
                pivotTable,
                worksheet,
                expectedName,
                expectedAnchor);
            dynamic cache = ReadPivotCache(pivotTable);
            if (!ReadBoolean(
                    () => (object?)cache.OLAP,
                    "Data Model recovery target PivotCache.OLAP"))
            {
                throw new InvalidOperationException(
                    "The pending target is not a native Data Model PivotTable.");
            }

            object connection = ReadRequired(
                () => (object?)cache.WorkbookConnection,
                "Excel did not expose the recovery target Data Model connection.");
            if (!SameNativeObject(connection, expectedModelConnection))
            {
                throw new InvalidOperationException(
                    "The pending target does not use the exact workbook Data Model connection.");
            }
        }

        private static void DemandExactPivotIdentity(
            object pivotTable,
            object worksheet,
            string expectedName,
            string expectedAnchor)
        {
            dynamic pivot = pivotTable;
            string name = ReadRequiredString(
                () => (object?)pivot.Name,
                "Excel did not expose the recovery target PivotTable name.");
            object parent = ReadRequired(
                () => (object?)pivot.Parent,
                "Excel did not expose the recovery target worksheet.");
            object range = ReadRequired(
                () => (object?)pivot.TableRange2,
                "Excel did not expose the recovery target range.");
            dynamic nativeRange = range;
            object firstCell = ReadRequired(
                () => (object?)nativeRange.Cells[1, 1],
                "Excel did not expose the recovery target anchor.");
            if (!string.Equals(name, expectedName, StringComparison.Ordinal) ||
                !SameNativeObject(parent, worksheet) ||
                !string.Equals(
                    ReadAddress(firstCell),
                    expectedAnchor,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "The pending recovery target identity or anchor changed.");
            }
        }

        private object CreateRecoveredTarget(
            object targetWorksheet,
            PivotStagedDataModelPivot staging,
            LateBoundPivotState stagedState,
            PivotDataModelArtifacts artifacts,
            PivotTemporaryPivotTableArtifact temporaryReceipt,
            string checkpointFingerprint,
            PivotFormatBackup formatBackup)
        {
            dynamic worksheet = targetWorksheet;
            dynamic cache = staging.NativePivotCache;
            object destination = ReadRequired(
                () => (object?)worksheet.Range[temporaryReceipt.TargetAnchorAddress],
                "Excel did not expose the recovery target anchor cell.");
            object recovered = cache.CreatePivotTable(
                destination,
                temporaryReceipt.Name);
            DemandExactTemporaryTargetPivot(
                recovered,
                targetWorksheet,
                temporaryReceipt,
                artifacts.NativeDataModelConnection);
            RestoreDataModelState(
                recovered,
                stagedState,
                artifacts.ModelTableName);
            RefreshPivotTable(recovered);
            formatBackup.Restore(recovered, stagedState.Result);
            DemandStateFingerprint(recovered, checkpointFingerprint);
            dynamic nativeRecovered = recovered;
            nativeRecovered.Name = temporaryReceipt.TargetPivotTableName;
            string promotedName = ReadRequiredString(
                () => (object?)nativeRecovered.Name,
                "Excel did not expose the promoted recovery target name.");
            if (!string.Equals(
                    promotedName,
                    temporaryReceipt.TargetPivotTableName,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "Excel did not promote the verified recovery target to the original PivotTable name.");
            }

            DemandStateFingerprint(recovered, checkpointFingerprint);
            return recovered;
        }

        private static void DemandDestinationRectangleEmpty(
            object worksheetObject,
            string anchorAddress,
            PivotResultSignature result,
            string originalPivotName,
            string temporaryPivotName)
        {
            if (result.Rows <= 0 || result.Columns <= 0 ||
                (long)result.Rows * result.Columns > MaximumPivotResultCells)
            {
                throw new NotSupportedException(
                    "The recovery destination extent is outside the bounded PivotTable result limit.");
            }

            dynamic worksheet = worksheetObject;
            object anchor = ReadRequired(
                () => (object?)worksheet.Range[anchorAddress],
                "Excel did not expose the recovery destination anchor.");
            dynamic nativeAnchor = anchor;
            object destination = ReadRequired(
                () => (object?)nativeAnchor.Resize[result.Rows, result.Columns],
                "Excel did not expose the bounded recovery destination rectangle.");
            dynamic range = destination;
            if (ContainsWorksheetPayload(ReadRequiredOptional(
                    () => (object?)range.Value2,
                    "recovery destination values")) ||
                ContainsWorksheetPayload(ReadRequiredOptional(
                    () => (object?)range.Formula,
                    "recovery destination formulas")))
            {
                throw new InvalidOperationException(
                    "The recovery destination contains values or formulas.");
            }

            DemandNoUnsupportedCellMetadata(destination);
            object mergeState = ReadRequired(
                () => (object?)range.MergeCells,
                "Excel did not expose the recovery destination merge state.");
            if (!(mergeState is bool merged) || merged)
            {
                throw new InvalidOperationException(
                    "The recovery destination contains merged cells or has an ambiguous merge state.");
            }

            object listObjects = ReadRequired(
                () => (object?)worksheet.ListObjects,
                "Excel did not expose target worksheet tables during recovery.");
            foreach (object listObject in ReadCollection(
                         listObjects,
                         MaximumWorkbookObjects,
                         "target worksheet tables"))
            {
                dynamic table = listObject;
                object tableRange = ReadRequired(
                    () => (object?)table.Range,
                    "Excel did not expose a target worksheet table range.");
                if (RangesOverlap(destination, tableRange))
                {
                    throw new InvalidOperationException(
                        "The recovery destination overlaps an Excel table.");
                }
            }

            object pivotTables = ReadRequired(
                () => (object?)worksheet.PivotTables,
                "Excel did not expose target worksheet PivotTables during destination recovery.");
            foreach (object pivotObject in ReadCollection(
                         pivotTables,
                         MaximumFields,
                         "target worksheet PivotTables"))
            {
                dynamic pivot = pivotObject;
                string name = ReadRequiredString(
                    () => (object?)pivot.Name,
                    "Excel exposed an unnamed PivotTable during destination recovery.");
                if (string.Equals(name, originalPivotName, StringComparison.Ordinal) ||
                    string.Equals(name, temporaryPivotName, StringComparison.Ordinal))
                {
                    // The caller resolves these identities before asking for an
                    // empty destination. Finding either here is inconsistent.
                    throw new InvalidOperationException(
                        "A recovery target still occupies the destination.");
                }

                object pivotRange = ReadRequired(
                    () => (object?)pivot.TableRange2,
                    "Excel did not expose a neighboring PivotTable range.");
                if (RangesOverlap(destination, pivotRange))
                {
                    throw new InvalidOperationException(
                        "The recovery destination overlaps another PivotTable.");
                }
            }
        }

        private static bool RangesOverlap(object leftObject, object rightObject)
        {
            dynamic left = leftObject;
            dynamic right = rightObject;
            int leftRow = ReadRequiredPositiveInt(
                () => (object?)left.Row,
                "range first row");
            int leftColumn = ReadRequiredPositiveInt(
                () => (object?)left.Column,
                "range first column");
            int leftRows = ReadRequiredCollectionCount(
                ReadRequired(() => (object?)left.Rows, "Excel did not expose range rows."),
                MaximumPivotResultCells,
                "range rows");
            int leftColumns = ReadRequiredCollectionCount(
                ReadRequired(() => (object?)left.Columns, "Excel did not expose range columns."),
                MaximumPivotResultCells,
                "range columns");
            int rightRow = ReadRequiredPositiveInt(
                () => (object?)right.Row,
                "neighbor range first row");
            int rightColumn = ReadRequiredPositiveInt(
                () => (object?)right.Column,
                "neighbor range first column");
            int rightRows = ReadRequiredCollectionCount(
                ReadRequired(() => (object?)right.Rows, "Excel did not expose neighbor range rows."),
                ExcelMaximumRows,
                "neighbor range rows");
            int rightColumns = ReadRequiredCollectionCount(
                ReadRequired(() => (object?)right.Columns, "Excel did not expose neighbor range columns."),
                ExcelMaximumColumns,
                "neighbor range columns");
            long leftLastRow = (long)leftRow + leftRows - 1;
            long leftLastColumn = (long)leftColumn + leftColumns - 1;
            long rightLastRow = (long)rightRow + rightRows - 1;
            long rightLastColumn = (long)rightColumn + rightColumns - 1;
            return leftRow <= rightLastRow && rightRow <= leftLastRow &&
                   leftColumn <= rightLastColumn && rightColumn <= leftLastColumn;
        }

        internal static GeneratedNames CompileGeneratedNames(string setupId)
        {
            return GeneratedNames.For(setupId);
        }

        private static IReadOnlyList<PivotTemporaryWorksheetArtifact>
            CreateTemporaryWorksheetReceipts(
                GeneratedNames names,
                string targetAnchorAddress)
        {
            return new[]
            {
                CreateTemporaryWorksheetReceipt(
                    names.StagingWorksheetName,
                    "staging",
                    targetAnchorAddress),
                CreateTemporaryWorksheetReceipt(
                    names.FormatBackupWorksheetName,
                    "format-backup",
                    targetAnchorAddress)
            };
        }

        private static PivotTemporaryWorksheetArtifact
            CreateTemporaryWorksheetReceipt(
                string name,
                string purpose,
                string targetAnchorAddress)
        {
            return new PivotTemporaryWorksheetArtifact(
                name,
                purpose,
                PivotPlusFingerprint.Create(
                    "pivotplus.temporary-worksheet.v2",
                    purpose + "\n" + name + "\n" + targetAnchorAddress),
                targetAnchorAddress);
        }

        private static PivotTemporaryPivotTableArtifact
            CreateTemporaryPivotTableReceipt(
                string setupId,
                GeneratedNames names,
                PivotTargetIdentity target,
                string targetAnchorAddress)
        {
            string canonical = setupId + "\n" +
                               names.ReplacementPivotTableName + "\n" +
                               target.WorksheetName + "\n" +
                               target.PivotTableName + "\n" +
                               targetAnchorAddress + "\n" +
                               names.ConnectionName + "\n" +
                               names.QueryName;
            return new PivotTemporaryPivotTableArtifact(
                setupId,
                names.ReplacementPivotTableName,
                PivotPlusFingerprint.Create(
                    "pivotplus.temporary-pivot-table.v1",
                    canonical),
                target.WorksheetName,
                target.PivotTableName,
                targetAnchorAddress,
                names.ConnectionName,
                names.QueryName);
        }

        private static PivotTemporaryWorksheetArtifact
            FindTemporaryWorksheetReceipt(
                PivotDataModelArtifacts artifacts,
                string expectedName,
                string expectedPurpose)
        {
            List<PivotTemporaryWorksheetArtifact> matches =
                artifacts.TemporaryWorksheets
                    .Where(item =>
                        string.Equals(
                            item.Name,
                            expectedName,
                            StringComparison.Ordinal) &&
                        string.Equals(
                            item.Purpose,
                            expectedPurpose,
                            StringComparison.Ordinal))
                    .ToList();
            if (matches.Count != 1)
            {
                throw new InvalidOperationException(
                    "PivotTable+ is missing an exact durable " +
                    expectedPurpose + " worksheet receipt.");
            }

            PivotTemporaryWorksheetArtifact receipt = matches[0];
            string expectedFingerprint = PivotPlusFingerprint.Create(
                "pivotplus.temporary-worksheet.v2",
                expectedPurpose + "\n" + expectedName + "\n" +
                receipt.TargetAnchorAddress);
            if (!string.Equals(
                    receipt.Fingerprint,
                    expectedFingerprint,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The durable temporary worksheet receipt changed.");
            }

            return receipt;
        }

        private static PivotPlusOwnedArtifact DemandSingleOwnedArtifact(
            PivotPlusWorkbookMetadata ownership,
            PivotPlusArtifactKind kind,
            string expectedId)
        {
            List<PivotPlusOwnedArtifact> matches = ownership.Artifacts
                .Where(item =>
                    item.Kind == kind &&
                    string.Equals(
                        item.ArtifactId,
                        expectedId,
                        StringComparison.Ordinal))
                .ToList();
            if (matches.Count != 1)
            {
                throw new InvalidOperationException(
                    "The pending ownership is missing an exact " +
                    kind + " receipt.");
            }

            return matches[0];
        }

        private static string DemandCanonicalGeneratedQuery(string formula)
        {
            Match match = Regex.Match(
                formula,
                "Excel\\.CurrentWorkbook\\(\\)\\{\\[Name=\\\"" +
                "(?<name>[A-Za-z_][A-Za-z0-9_.]*)\\\"\\]\\}\\[Content\\]",
                RegexOptions.CultureInvariant);
            if (!match.Success || match.NextMatch().Success)
            {
                throw new InvalidOperationException(
                    "The pending query is not a canonical PivotTable+ workbook-only query.");
            }

            string name = match.Groups["name"].Value;
            string table = PivotPlusSourceQueryCompiler.Compile(
                name,
                PivotPlusWorkbookObjectKind.Table);
            string namedRange = PivotPlusSourceQueryCompiler.Compile(
                name,
                PivotPlusWorkbookObjectKind.NamedRange);
            if (!string.Equals(formula, table, StringComparison.Ordinal) &&
                !string.Equals(formula, namedRange, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The pending query is not a canonical PivotTable+ workbook-only query.");
            }

            return name;
        }

        private static object? FindReusableDataModelCache(
            object cacheCollection,
            object nativeConnection)
        {
            var matches = new List<object>();
            foreach (object cacheObject in ReadCollection(
                         cacheCollection,
                         MaximumWorkbookObjects,
                         "PivotCaches"))
            {
                dynamic cache = cacheObject;
                bool olap = ReadBoolean(
                    () => (object?)cache.OLAP,
                    "PivotCache.OLAP during recovery");
                if (!olap) continue;
                object connection = ReadRequired(
                    () => (object?)cache.WorkbookConnection,
                    "Excel did not expose an OLAP PivotCache connection during recovery.");
                if (SameNativeObject(connection, nativeConnection))
                {
                    matches.Add(cacheObject);
                }
            }

            // Workbook.Model.DataModelConnection is intentionally shared by
            // every native Data Model PivotTable. Reuse the first cache in
            // Excel's stable collection order; multiple matches are normal.
            return matches.FirstOrDefault();
        }

        private static object ReadExactDataModelConnection(
            object workbook,
            object sourceConnection,
            string expectedModelTableName)
        {
            dynamic source = sourceConnection;
            if (!ReadRequiredBoolean(
                    () => (object?)source.InModel,
                    "owned source connection InModel state"))
            {
                throw new InvalidOperationException(
                    "Excel did not load the owned source connection into the workbook Data Model.");
            }

            dynamic book = workbook;
            object modelObject = ReadRequired(
                () => (object?)book.Model,
                "Excel did not expose the workbook Data Model after loading the source connection.");
            dynamic model = modelObject;
            object modelTables = ReadRequired(
                () => (object?)model.ModelTables,
                "Excel did not expose Data Model tables after loading the source connection.");
            List<object> exactTables = ReadCollection(
                    modelTables,
                    MaximumWorkbookObjects,
                    "Data Model tables")
                .Where(item =>
                {
                    dynamic table = item;
                    string name = ReadRequiredString(
                        () => (object?)table.Name,
                        "Excel exposed an unnamed Data Model table.");
                    return string.Equals(
                        name,
                        expectedModelTableName,
                        StringComparison.Ordinal);
                })
                .ToList();
            if (exactTables.Count != 1)
            {
                throw new InvalidOperationException(
                    "Excel did not expose exactly one Data Model table for the owned query.");
            }

            dynamic exactTable = exactTables[0];
            object tableSourceConnection = ReadRequired(
                () => (object?)exactTable.SourceWorkbookConnection,
                "Excel did not expose the Data Model table's source workbook connection.");
            if (!SameNativeObject(tableSourceConnection, sourceConnection))
            {
                throw new InvalidOperationException(
                    "The generated Data Model table is not sourced by the exact owned query connection.");
            }

            object modelConnection = ReadRequired(
                () => (object?)model.DataModelConnection,
                "Excel did not expose Workbook.Model.DataModelConnection.");
            dynamic nativeModelConnection = modelConnection;
            if (ReadRequiredInt(
                    () => (object?)nativeModelConnection.Type,
                    "Data Model connection type") != 7)
            {
                throw new InvalidOperationException(
                    "Excel exposed a non-model connection as Workbook.Model.DataModelConnection.");
            }

            return modelConnection;
        }

        private static void WriteTemporaryWorksheetMarker(
            dynamic worksheet,
            PivotTemporaryWorksheetArtifact receipt)
        {
            object properties = ReadRequired(
                () => (object?)worksheet.CustomProperties,
                "Excel did not expose worksheet CustomProperties for durable temporary ownership.");
            if (ReadCollectionCount(
                    properties,
                    MaximumFields,
                    "worksheet CustomProperties") != 0)
            {
                throw new InvalidOperationException(
                    "A newly created temporary worksheet unexpectedly already has CustomProperties.");
            }

            dynamic nativeProperties = properties;
            nativeProperties.Add(TemporaryPurposeProperty, receipt.Purpose);
            nativeProperties.Add(TemporaryFingerprintProperty, receipt.Fingerprint);
            nativeProperties.Add(
                TemporaryAnchorProperty,
                receipt.TargetAnchorAddress);
            DemandTemporaryWorksheetMarker(worksheet, receipt);
        }

        private static void DemandTemporaryWorksheetMarker(
            dynamic worksheet,
            PivotTemporaryWorksheetArtifact receipt)
        {
            Dictionary<string, string> markers =
                ReadTemporaryWorksheetMarkers(worksheet);
            bool hasVerifiedState = markers.ContainsKey(
                StagingStateFingerprintProperty);
            int expectedCount = hasVerifiedState ? 4 : 3;
            if (markers.Count != expectedCount ||
                !markers.TryGetValue(TemporaryPurposeProperty, out string? purpose) ||
                !markers.TryGetValue(TemporaryFingerprintProperty, out string? fingerprint) ||
                !markers.TryGetValue(TemporaryAnchorProperty, out string? anchor) ||
                !string.Equals(purpose, receipt.Purpose, StringComparison.Ordinal) ||
                !string.Equals(fingerprint, receipt.Fingerprint, StringComparison.Ordinal) ||
                !string.Equals(
                    anchor,
                    receipt.TargetAnchorAddress,
                    StringComparison.Ordinal) ||
                (hasVerifiedState && !string.Equals(
                    receipt.Purpose,
                    "staging",
                    StringComparison.Ordinal)))
            {
                throw new InvalidOperationException(
                    "The temporary worksheet ownership marker is missing or changed.");
            }
        }

        private static Dictionary<string, string> ReadTemporaryWorksheetMarkers(
            dynamic worksheet)
        {
            object properties = ReadRequired(
                () => (object?)worksheet.CustomProperties,
                "Excel did not expose the temporary worksheet ownership marker.");
            var markers = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (object propertyObject in ReadCollection(
                         properties,
                         MaximumFields,
                         "worksheet CustomProperties"))
            {
                dynamic property = propertyObject;
                string name = ReadRequiredString(
                    () => (object?)property.Name,
                    "Excel exposed an unnamed worksheet CustomProperty.");
                string value = ReadRequiredString(
                    () => (object?)property.Value,
                    "Excel exposed an empty worksheet CustomProperty value.");
                if (markers.ContainsKey(name))
                {
                    throw new InvalidOperationException(
                        "The temporary worksheet has duplicate ownership markers.");
                }

                markers.Add(name, value);
            }

            return markers;
        }

        private static void DemandVerifiedStagingMarker(
            dynamic worksheet,
            PivotTemporaryWorksheetArtifact receipt,
            string expectedStateFingerprint)
        {
            DemandTemporaryWorksheetMarker(worksheet, receipt);
            Dictionary<string, string> markers =
                ReadTemporaryWorksheetMarkers(worksheet);
            if (!markers.TryGetValue(
                    StagingStateFingerprintProperty,
                    out string? actual) ||
                !string.Equals(
                    actual,
                    expectedStateFingerprint,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The staging PivotTable is not durably checkpointed with the expected verified state.");
            }
        }

        private static string FingerprintDataModelState(LateBoundPivotState state)
        {
            return PivotPlusFingerprint.Create(
                "pivotplus.staging-state.v1",
                state.CanonicalValue());
        }

        private static LateBoundPivotState ReadLivePivotState(object pivotTable)
        {
            dynamic pivot = pivotTable;
            return new LateBoundPivotState(
                ReadPivotCache(pivotTable),
                ReadFieldStates(pivot),
                ReadStyleState(pivot),
                ReadResultSignature(ReadRequired(
                    () => (object?)pivot.TableRange2,
                    "Excel did not expose the Data Model PivotTable result range.")),
                ReadDataAxisState(pivotTable));
        }

        private static void DemandStateFingerprint(
            object pivotTable,
            string expectedFingerprint)
        {
            string actual = FingerprintDataModelState(
                ReadLivePivotState(pivotTable));
            if (!string.Equals(
                    actual,
                    expectedFingerprint,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The Data Model PivotTable state does not match the durable staging checkpoint.");
            }
        }

        private static void DemandTemporaryPivotTableReceipt(
            PivotTemporaryPivotTableArtifact receipt,
            string targetWorksheetName,
            string targetPivotTableName,
            string targetAnchorAddress,
            string modelTableName)
        {
            if (!string.Equals(
                    receipt.TargetWorksheetName,
                    targetWorksheetName,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    receipt.TargetPivotTableName,
                    targetPivotTableName,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    receipt.TargetAnchorAddress,
                    targetAnchorAddress,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    receipt.ModelTableName,
                    modelTableName,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The temporary target PivotTable receipt no longer matches the captured target.");
            }

            string canonical = receipt.SetupId + "\n" +
                               receipt.Name + "\n" +
                               receipt.TargetWorksheetName + "\n" +
                               receipt.TargetPivotTableName + "\n" +
                               receipt.TargetAnchorAddress + "\n" +
                               receipt.ConnectionName + "\n" +
                               receipt.ModelTableName;
            if (!PivotPlusFingerprint.Matches(
                    receipt.Fingerprint,
                    "pivotplus.temporary-pivot-table.v1",
                    canonical))
            {
                throw new InvalidOperationException(
                    "The temporary target PivotTable ownership receipt changed.");
            }
        }

        private static void DemandExactTemporaryTargetPivot(
            object pivotTable,
            object targetWorksheet,
            PivotTemporaryPivotTableArtifact receipt,
            object expectedModelConnection)
        {
            dynamic pivot = pivotTable;
            string name = ReadRequiredString(
                () => (object?)pivot.Name,
                "Excel did not expose the temporary target PivotTable name.");
            if (!string.Equals(name, receipt.Name, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The generated temporary target PivotTable was renamed.");
            }

            object parent = ReadRequired(
                () => (object?)pivot.Parent,
                "Excel did not expose the temporary target PivotTable worksheet.");
            if (!SameNativeObject(parent, targetWorksheet))
            {
                throw new InvalidOperationException(
                    "The generated temporary target PivotTable moved to another worksheet.");
            }

            object range = ReadRequired(
                () => (object?)pivot.TableRange2,
                "Excel did not expose the temporary target PivotTable range.");
            dynamic nativeRange = range;
            object firstCell = ReadRequired(
                () => (object?)nativeRange.Cells[1, 1],
                "Excel did not expose the temporary target PivotTable anchor.");
            if (!string.Equals(
                    ReadAddress(firstCell),
                    receipt.TargetAnchorAddress,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "The generated temporary target PivotTable moved from its owned anchor.");
            }

            dynamic cache = ReadPivotCache(pivotTable);
            if (!ReadBoolean(
                    () => (object?)cache.OLAP,
                    "temporary target PivotCache.OLAP"))
            {
                throw new InvalidOperationException(
                    "The generated temporary target is not a Data Model PivotTable.");
            }

            object connection = ReadRequired(
                () => (object?)cache.WorkbookConnection,
                "Excel did not expose the temporary target Data Model connection.");
            if (!SameNativeObject(connection, expectedModelConnection))
            {
                throw new InvalidOperationException(
                    "The generated temporary target does not use the exact workbook Data Model connection.");
            }
        }

        private static void ReconcileStaleTemporaryWorksheet(
            object workbook,
            PivotTemporaryWorksheetArtifact receipt,
            string expectedPivotName,
            bool isFormatBackup,
            object? expectedModelConnection)
        {
            dynamic book = workbook;
            object worksheets = ReadRequired(
                () => (object?)book.Worksheets,
                "Excel did not expose workbook worksheets during recovery.");
            object? existing = FindNamedObject(
                worksheets,
                receipt.Name,
                "worksheets");
            if (existing == null) return;
            DemandTemporaryWorksheetMarker((dynamic)existing, receipt);
            DemandTemporaryWorksheetStructure(
                (dynamic)existing,
                expectedPivotName,
                isFormatBackup,
                allowIncomplete: true,
                expectedModelConnection: expectedModelConnection);
            DeleteOwnedTemporaryWorksheet(workbook, (dynamic)existing, receipt);
        }

        internal static void DemandTemporaryWorksheetStructure(
            object worksheetObject,
            string expectedPivotName,
            bool isFormatBackup,
            bool allowIncomplete,
            object? expectedModelConnection)
        {
            dynamic worksheet = worksheetObject;
            if (ReadRequiredInt(
                    () => (object?)worksheet.Visible,
                    "temporary worksheet visibility") != SheetVeryHidden)
            {
                throw new InvalidOperationException(
                    "The owned temporary worksheet is no longer VeryHidden.");
            }

            object pivotTables = ReadRequired(
                () => (object?)worksheet.PivotTables,
                "Excel did not expose temporary worksheet PivotTables.");
            int pivotCount = ReadCollectionCount(
                pivotTables,
                MaximumFields,
                "temporary worksheet PivotTables");
            if (isFormatBackup)
            {
                if (pivotCount != 0)
                {
                    throw new InvalidOperationException(
                        "The format-backup worksheet unexpectedly contains a PivotTable.");
                }

                DemandNoWorksheetPayload(worksheet);
            }
            else
            {
                if (pivotCount > 1 || (!allowIncomplete && pivotCount != 1))
                {
                    throw new InvalidOperationException(
                        "The staging worksheet has an unexpected PivotTable structure.");
                }

                if (pivotCount == 1)
                {
                    object pivotObject = ReadCollection(
                        pivotTables,
                        MaximumFields,
                        "temporary worksheet PivotTables")[0];
                    dynamic pivot = pivotObject;
                    string name = ReadRequiredString(
                        () => (object?)pivot.Name,
                        "Excel did not expose the staged PivotTable name.");
                    if (!string.Equals(name, expectedPivotName, StringComparison.Ordinal))
                    {
                        throw new InvalidOperationException(
                            "The staging worksheet contains an unexpected PivotTable.");
                    }

                    DemandExactStagingPivotStructure(
                        worksheet,
                        pivotObject,
                        expectedModelConnection);
                }
                else
                {
                    DemandNoWorksheetPayload(worksheet);
                }
            }

            DemandEmptyWorksheetCollection(worksheet, "ListObjects");
            DemandEmptyWorksheetCollection(worksheet, "Shapes");
        }

        private static void DemandEmptyWorksheetCollection(
            dynamic worksheet,
            string member)
        {
            object collection;
            if (string.Equals(member, "ListObjects", StringComparison.Ordinal))
            {
                collection = ReadRequired(
                    () => (object?)worksheet.ListObjects,
                    "Excel did not expose temporary worksheet ListObjects.");
            }
            else
            {
                collection = ReadRequired(
                    () => (object?)worksheet.Shapes,
                    "Excel did not expose temporary worksheet Shapes.");
            }

            if (ReadCollectionCount(
                    collection,
                    MaximumWorkbookObjects,
                    "temporary worksheet " + member) != 0)
            {
                throw new InvalidOperationException(
                    "The owned temporary worksheet contains unexpected " + member + ".");
            }
        }

        private static void DemandNoWorksheetPayload(dynamic worksheet)
        {
            object usedRange = ReadRequired(
                () => (object?)worksheet.UsedRange,
                "Excel did not expose the format-backup UsedRange.");
            dynamic range = usedRange;
            object cells = ReadRequired(
                () => (object?)range.Cells,
                "Excel did not expose the format-backup cell count.");
            long cellCount = ReadRequiredLong(
                () => (object?)((dynamic)cells).CountLarge,
                "format-backup cell count");
            if (cellCount <= 0 || cellCount > MaximumPivotResultCells)
            {
                throw new NotSupportedException(
                    "The format-backup worksheet exceeds its bounded extent.");
            }

            object? values = ReadRequiredOptional(
                () => (object?)range.Value2,
                "format-backup values");
            object? formulas = ReadRequiredOptional(
                () => (object?)range.Formula,
                "format-backup formulas");
            if (ContainsWorksheetPayload(values) || ContainsWorksheetPayload(formulas))
            {
                throw new InvalidOperationException(
                    "The format-backup worksheet contains values or formulas and will not be adopted or deleted.");
            }
        }

        private static void DemandExactStagingPivotStructure(
            dynamic worksheet,
            object pivotObject,
            object? expectedModelConnection)
        {
            if (expectedModelConnection == null)
            {
                throw new InvalidOperationException(
                    "The staging worksheet cannot be reconciled without the exact Data Model connection.");
            }

            dynamic cache = ReadPivotCache(pivotObject);
            if (!ReadBoolean(
                    () => (object?)cache.OLAP,
                    "staged PivotCache.OLAP"))
            {
                throw new InvalidOperationException(
                    "The owned staging PivotTable is not a Data Model PivotTable.");
            }

            object connection = ReadRequired(
                () => (object?)cache.WorkbookConnection,
                "Excel did not expose the staged PivotCache connection.");
            if (!SameNativeObject(connection, expectedModelConnection))
            {
                throw new InvalidOperationException(
                    "The owned staging PivotTable does not use the exact workbook Data Model connection.");
            }

            dynamic pivot = pivotObject;
            object tableRange = ReadRequired(
                () => (object?)pivot.TableRange2,
                "Excel did not expose the staged PivotTable range.");
            object usedRange = ReadRequired(
                () => (object?)worksheet.UsedRange,
                "Excel did not expose the staging worksheet UsedRange.");
            string tableAddress = ReadAddress(tableRange);
            string usedAddress = ReadAddress(usedRange);
            if (!string.Equals(
                    tableAddress,
                    usedAddress,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "The staging worksheet contains cells outside the exact staged PivotTable range.");
            }

            dynamic nativeTableRange = tableRange;
            dynamic nativeUsedRange = usedRange;
            int tableRows = ReadRequiredCollectionCount(
                ReadRequired(
                    () => (object?)nativeTableRange.Rows,
                    "Excel did not expose staged PivotTable rows."),
                MaximumPivotResultCells,
                "staged PivotTable rows");
            int tableColumns = ReadRequiredCollectionCount(
                ReadRequired(
                    () => (object?)nativeTableRange.Columns,
                    "Excel did not expose staged PivotTable columns."),
                MaximumPivotResultCells,
                "staged PivotTable columns");
            int usedRows = ReadRequiredCollectionCount(
                ReadRequired(
                    () => (object?)nativeUsedRange.Rows,
                    "Excel did not expose staging UsedRange rows."),
                MaximumPivotResultCells,
                "staging UsedRange rows");
            int usedColumns = ReadRequiredCollectionCount(
                ReadRequired(
                    () => (object?)nativeUsedRange.Columns,
                    "Excel did not expose staging UsedRange columns."),
                MaximumPivotResultCells,
                "staging UsedRange columns");
            if (tableRows != usedRows || tableColumns != usedColumns)
            {
                throw new InvalidOperationException(
                    "The staging worksheet UsedRange is not exactly the staged PivotTable extent.");
            }
        }

        private static bool ContainsWorksheetPayload(object? value)
        {
            if (value == null) return false;
            if (value is Array array)
            {
                if (array.Length > MaximumPivotResultCells)
                {
                    throw new NotSupportedException(
                        "The temporary worksheet payload exceeds its bounded inspection limit.");
                }

                foreach (object? item in array)
                {
                    if (ContainsWorksheetPayload(item)) return true;
                }

                return false;
            }

            return !(value is string text) || text.Length != 0;
        }

        private static object? ReadRequiredOptional(
            Func<object?> reader,
            string label)
        {
            if (!PivotLateBound.TryRead(reader, out object? value))
            {
                throw new NotSupportedException(
                    "Excel did not expose the " + label + " safely.");
            }

            return value;
        }

        private static long ReadRequiredLong(Func<object?> reader, string label)
        {
            object value = ReadRequired(
                reader,
                "Excel did not expose the " + label + ".");
            try
            {
                return Convert.ToInt64(value, CultureInfo.InvariantCulture);
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

        private static void DeleteOwnedTemporaryWorksheet(
            object workbook,
            dynamic worksheet,
            PivotTemporaryWorksheetArtifact receipt)
        {
            try
            {
                DeleteWorksheet(worksheet);
            }
            catch (Exception deleteFailure)
            {
                try
                {
                    dynamic book = workbook;
                    object worksheets = ReadRequired(
                        () => (object?)book.Worksheets,
                        "Excel did not expose workbook worksheets after a reported delete failure.");
                    object? survivor = FindNamedObject(
                        worksheets,
                        receipt.Name,
                        "worksheets");
                    if (survivor == null)
                    {
                        // Excel can report a COM failure after Delete committed.
                        return;
                    }

                    DemandTemporaryWorksheetMarker((dynamic)survivor, receipt);
                }
                catch (Exception inspectionFailure)
                {
                    throw new AggregateException(
                        "Excel reported a temporary worksheet delete failure and its outcome is ambiguous.",
                        deleteFailure,
                        inspectionFailure);
                }

                throw new InvalidOperationException(
                    "Excel did not delete the exact owned temporary worksheet.",
                    deleteFailure);
            }
        }

        internal static void DemandNoConnectedSlicersOrTimelines(object pivotTable)
        {
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            dynamic pivot = pivotTable;
            string pivotName = ReadRequiredString(
                () => (object?)pivot.Name,
                "Excel did not expose the PivotTable name while inspecting slicers.");
            object worksheet = ReadRequired(
                () => (object?)pivot.Parent,
                "Excel did not expose the PivotTable worksheet while inspecting slicers.");
            dynamic nativeWorksheet = worksheet;
            object workbook = ReadRequired(
                () => (object?)nativeWorksheet.Parent,
                "Excel did not expose the PivotTable workbook while inspecting slicers.");
            dynamic book = workbook;
            object? slicerCaches = TryGet(() => (object?)book.SlicerCaches);
            if (slicerCaches == null)
            {
                throw new NotSupportedException(
                    "Excel did not expose SlicerCaches needed for a reversible relationship preflight.");
            }

            foreach (object cacheObject in ReadCollection(
                         slicerCaches,
                         MaximumWorkbookObjects,
                         "SlicerCaches"))
            {
                dynamic slicerCache = cacheObject;
                object? connectedPivots = TryGet(() => (object?)slicerCache.PivotTables);
                if (connectedPivots == null)
                {
                    throw new NotSupportedException(
                        "Excel did not expose a SlicerCache PivotTables collection safely.");
                }

                foreach (object connectedObject in ReadCollection(
                             connectedPivots,
                             MaximumFields,
                             "slicer PivotTables"))
                {
                    dynamic connected = connectedObject;
                    object identityObject = connectedObject;
                    string connectedName;
                    if (PivotLateBound.TryRead(
                            () => (object?)connected.Name,
                            out object? directName) &&
                        directName != null &&
                        !string.IsNullOrWhiteSpace(Convert.ToString(
                            directName,
                            CultureInfo.InvariantCulture)))
                    {
                        connectedName = Convert.ToString(
                            directName,
                            CultureInfo.InvariantCulture)!;
                    }
                    else if (PivotLateBound.TryRead(
                                 () => (object?)connected.PivotTable,
                                 out object? nestedPivot) &&
                             nestedPivot != null)
                    {
                        identityObject = nestedPivot;
                        dynamic nested = nestedPivot;
                        connectedName = ReadRequiredString(
                            () => (object?)nested.Name,
                            "Excel exposed an unnamed PivotTable through a SlicerCache.");
                    }
                    else
                    {
                        throw new NotSupportedException(
                            "Excel did not expose a SlicerCache PivotTable identity safely.");
                    }

                    if (SameNativeObject(identityObject, pivotTable) ||
                        string.Equals(
                            connectedName,
                            pivotName,
                            StringComparison.OrdinalIgnoreCase))
                    {
                        throw new NotSupportedException(
                            "The selected PivotTable is connected to a slicer or timeline. PivotTable+ will not replace it until that relationship can be restored transactionally.");
                    }
                }
            }
        }

        internal static void DemandNoAttachedPivotCharts(object pivotTable)
        {
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            dynamic pivot = pivotTable;
            string pivotName = ReadRequiredString(
                () => (object?)pivot.Name,
                "Excel did not expose the PivotTable name while inspecting PivotCharts.");
            object worksheet = ReadRequired(
                () => (object?)pivot.Parent,
                "Excel did not expose the PivotTable worksheet while inspecting PivotCharts.");
            dynamic nativeWorksheet = worksheet;
            object workbook = ReadRequired(
                () => (object?)nativeWorksheet.Parent,
                "Excel did not expose the PivotTable workbook while inspecting PivotCharts.");

            dynamic book = workbook;
            object worksheets = ReadRequired(
                () => (object?)book.Worksheets,
                "Excel did not expose workbook worksheets while inspecting PivotCharts.");
            foreach (object sheetObject in ReadCollection(
                         worksheets,
                         MaximumWorksheets,
                         "worksheets for PivotChart inspection"))
            {
                dynamic sheet = sheetObject;
                object chartObjects = ReadRequired(
                    () => (object?)sheet.ChartObjects,
                    "Excel did not expose worksheet ChartObjects while inspecting PivotCharts.");
                foreach (object chartObject in ReadCollection(
                             chartObjects,
                             MaximumWorkbookObjects,
                             "worksheet charts"))
                {
                    dynamic wrapper = chartObject;
                    object chart = ReadRequired(
                        () => (object?)wrapper.Chart,
                        "Excel did not expose a worksheet chart safely.");
                    if (ChartTargetsPivot(chart, pivotName))
                    {
                        ThrowAttachedPivotChart();
                    }
                }
            }

            object? chartSheets = TryGet(() => (object?)book.Charts);
            if (chartSheets == null)
            {
                throw new NotSupportedException(
                    "Excel did not expose workbook chart sheets while inspecting PivotCharts.");
            }

            foreach (object chart in ReadCollection(
                         chartSheets,
                         MaximumWorkbookObjects,
                         "chart sheets"))
            {
                if (ChartTargetsPivot(chart, pivotName))
                {
                    ThrowAttachedPivotChart();
                }
            }
        }

        private static bool ChartTargetsPivot(object chartObject, string pivotName)
        {
            dynamic chart = chartObject;
            if (!PivotLateBound.TryRead(
                    () => (object?)chart.PivotLayout,
                    out object? layout))
            {
                throw new NotSupportedException(
                    "Excel did not expose a chart PivotLayout safely.");
            }

            if (layout == null) return false;
            dynamic pivotLayout = layout;
            if (!PivotLateBound.TryRead(
                    () => (object?)pivotLayout.PivotTable,
                    out object? chartPivot) ||
                chartPivot == null)
            {
                throw new NotSupportedException(
                    "Excel did not expose the PivotTable behind a chart PivotLayout.");
            }

            dynamic nativePivot = chartPivot;
            string chartPivotName = ReadRequiredString(
                () => (object?)nativePivot.Name,
                "Excel exposed an unnamed PivotTable behind a PivotChart.");
            return string.Equals(
                chartPivotName,
                pivotName,
                StringComparison.OrdinalIgnoreCase);
        }

        private static void ThrowAttachedPivotChart()
        {
            throw new NotSupportedException(
                "The selected PivotTable has an attached PivotChart. PivotTable+ will not replace it until that relationship can be restored transactionally.");
        }

        private static void DemandExactPlanOwnership(
            PivotPlusWorkbookMetadata ownership,
            PivotDataModelArtifactPlan plan)
        {
            if (ownership.Artifacts == null)
            {
                throw new InvalidOperationException(
                    "The pending PivotTable+ ownership record has no artifact set.");
            }

            var expected = new List<PivotPlusOwnedArtifact>
            {
                new PivotPlusOwnedArtifact
                {
                    Kind = PivotPlusArtifactKind.Query,
                    ArtifactId = plan.QueryName,
                    Fingerprint = plan.QueryFingerprint
                },
                new PivotPlusOwnedArtifact
                {
                    Kind = PivotPlusArtifactKind.Connection,
                    ArtifactId = plan.ConnectionName,
                    Fingerprint = plan.ConnectionFingerprint
                }
            };
            if (plan.WorkbookName != null && plan.WorkbookNameFingerprint != null)
            {
                expected.Add(new PivotPlusOwnedArtifact
                {
                    Kind = PivotPlusArtifactKind.WorkbookName,
                    ArtifactId = plan.WorkbookName,
                    Fingerprint = plan.WorkbookNameFingerprint
                });
            }
            else if (plan.WorkbookName != null ||
                     plan.WorkbookNameFingerprint != null ||
                     plan.RequestedWorkbookNameReference != null)
            {
                throw new InvalidOperationException(
                    "The planned raw-source alias receipt is incomplete.");
            }

            expected.AddRange(plan.TemporaryWorksheets.Select(item =>
                new PivotPlusOwnedArtifact
                {
                    Kind = item.Kind,
                    ArtifactId = item.Name,
                    Fingerprint = item.Fingerprint
                }));
            if (plan.TemporaryPivotTable == null)
            {
                throw new InvalidOperationException(
                    "The deterministic artifact plan has no temporary target PivotTable receipt.");
            }

            expected.Add(new PivotPlusOwnedArtifact
            {
                Kind = plan.TemporaryPivotTable.Kind,
                ArtifactId = plan.TemporaryPivotTable.Name,
                Fingerprint = plan.TemporaryPivotTable.Fingerprint
            });
            if (ownership.Artifacts.Count != expected.Count ||
                expected.Any(receipt => ownership.Artifacts.Count(recorded =>
                    recorded.Kind == receipt.Kind &&
                    string.Equals(
                        recorded.ArtifactId,
                        receipt.ArtifactId,
                        StringComparison.Ordinal) &&
                    string.Equals(
                        recorded.Fingerprint,
                        receipt.Fingerprint,
                        StringComparison.Ordinal)) != 1))
            {
                throw new InvalidOperationException(
                    "The pending PivotTable+ ownership record does not exactly match the deterministic artifact plan.");
            }
        }

        private static void DemandExactPlannedQuery(
            object queryObject,
            PivotDataModelArtifactPlan plan)
        {
            dynamic query = queryObject;
            string name = ReadRequiredString(
                () => (object?)query.Name,
                "Excel did not expose the planned query name.");
            string formula = ReadRequiredString(
                () => (object?)query.Formula,
                "Excel did not expose the planned query formula.");
            if (!string.Equals(name, plan.QueryName, StringComparison.Ordinal) ||
                !string.Equals(formula, plan.QueryFormula, StringComparison.Ordinal) ||
                !PivotPlusFingerprint.Matches(
                    plan.QueryFingerprint,
                    "pivotplus.query.v1",
                    formula))
            {
                throw new InvalidOperationException(
                    "The live query does not exactly match the write-ahead PivotTable+ plan.");
            }
        }

        private static void DemandExactPlannedConnection(
            object connectionObject,
            PivotDataModelArtifactPlan plan)
        {
            dynamic connection = connectionObject;
            string name = ReadRequiredString(
                () => (object?)connection.Name,
                "Excel did not expose the planned connection name.");
            object oleDbObject = ReadRequired(
                () => (object?)connection.OLEDBConnection,
                "Excel did not expose the planned OLE DB connection.");
            dynamic oleDb = oleDbObject;
            string connectionString = ReadRequiredString(
                () => (object?)oleDb.Connection,
                "Excel did not expose the planned connection string.");
            string commandText = ReadCommandText(oleDb);
            if (!string.Equals(name, plan.ConnectionName, StringComparison.Ordinal) ||
                !string.Equals(
                    connectionString,
                    CanonicalConnectionContract.ConnectionString(plan.QueryName),
                    StringComparison.Ordinal) ||
                !string.Equals(
                    commandText,
                    CanonicalConnectionContract.CommandText(plan.QueryName),
                    StringComparison.Ordinal) ||
                !PivotPlusFingerprint.Matches(
                    plan.ConnectionFingerprint,
                    "pivotplus.connection.v1",
                    connectionString + "\n" + commandText))
            {
                throw new InvalidOperationException(
                    "The live connection does not exactly match the write-ahead PivotTable+ plan.");
            }
        }

        private static void DemandExactPlannedWorkbookName(
            object nativeName,
            PivotDataModelArtifactPlan plan)
        {
            if (plan.WorkbookName == null ||
                plan.WorkbookNameFingerprint == null ||
                plan.RequestedWorkbookNameReference == null)
            {
                throw new InvalidOperationException(
                    "The planned raw-source alias receipt is incomplete.");
            }

            CreateOwnedWorkbookNameReceipt(
                nativeName,
                plan.WorkbookName,
                plan.RequestedWorkbookNameReference,
                plan.WorkbookNameFingerprint);
        }

        private static void DemandExactConnection(
            object connectionObject,
            PivotDataModelArtifacts artifacts)
        {
            dynamic connection = connectionObject;
            string name = ReadRequiredString(
                () => (object?)connection.Name,
                "Excel did not expose the generated connection name.");
            if (!string.Equals(name, artifacts.ConnectionName, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The generated connection was renamed; PivotTable+ will not delete it.");
            }

            object oleDbObject = ReadRequired(
                () => (object?)connection.OLEDBConnection,
                "Excel did not expose the generated OLE DB connection.");
            dynamic oleDb = oleDbObject;
            string connectionString = ReadRequiredString(
                () => (object?)oleDb.Connection,
                "Excel did not expose the generated connection string.");
            string commandText = ReadCommandText(oleDb);
            string exactContract = connectionString + "\n" + commandText;
            if (!PivotPlusFingerprint.Matches(
                    artifacts.ConnectionFingerprint,
                    "pivotplus.connection.v1",
                    exactContract) ||
                !string.Equals(
                    connectionString,
                    CanonicalConnectionContract.ConnectionString(artifacts.QueryName),
                    StringComparison.Ordinal) ||
                !string.Equals(
                    commandText,
                    CanonicalConnectionContract.CommandText(artifacts.QueryName),
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The generated connection changed; PivotTable+ will not delete it.");
            }
        }

        private static PivotOwnedWorkbookNameArtifact CreateOwnedWorkbookNameReceipt(
            object nativeName,
            string expectedName,
            string requestedReference,
            string? expectedFingerprint = null)
        {
            dynamic name = nativeName;
            string actualName = ReadRequiredString(
                () => (object?)name.Name,
                "Excel did not expose the generated workbook name.");
            string leafName = WorkbookScopedName(actualName);
            if (!string.Equals(
                    leafName,
                    expectedName,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "Excel created the raw-source alias under an unexpected name.");
            }

            string actualReference = ReadRequiredString(
                () => (object?)name.RefersTo,
                "Excel did not expose the generated workbook name reference.");
            if (!string.Equals(
                    NormalizeLocalReference(actualReference),
                    NormalizeLocalReference(requestedReference),
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "Excel created the raw-source alias with an unexpected reference.");
            }

            string fingerprint = WorkbookNameFingerprint(actualReference);
            if (expectedFingerprint != null &&
                !string.Equals(
                    fingerprint,
                    expectedFingerprint,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The raw-source alias does not match its write-ahead ownership fingerprint.");
            }

            return new PivotOwnedWorkbookNameArtifact(
                expectedName,
                fingerprint,
                actualReference,
                nativeName);
        }

        private static void DemandExactOwnedWorkbookName(
            object nativeName,
            PivotOwnedWorkbookNameArtifact artifact)
        {
            dynamic name = nativeName;
            string actualName = ReadRequiredString(
                () => (object?)name.Name,
                "Excel did not expose the generated workbook name.");
            if (!string.Equals(
                    WorkbookScopedName(actualName),
                    artifact.Name,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "The generated raw-source alias was renamed; PivotTable+ will not delete it.");
            }

            string actualReference = ReadRequiredString(
                () => (object?)name.RefersTo,
                "Excel did not expose the generated workbook name reference.");
            if (!string.Equals(
                    artifact.ReferenceFingerprint,
                    WorkbookNameFingerprint(actualReference),
                    StringComparison.Ordinal) ||
                !string.Equals(
                    NormalizeLocalReference(actualReference),
                    NormalizeLocalReference(artifact.CanonicalReference),
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "The generated raw-source alias changed; PivotTable+ will not delete it.");
            }
        }

        private static string NormalizeLocalReference(string value)
        {
            return NormalizeSourceIdentity(value);
        }

        private static string WorkbookNameFingerprint(string reference)
        {
            return PivotPlusFingerprint.Create(
                "pivotplus.workbook-name.v2",
                NormalizeLocalReference(reference).ToUpperInvariant());
        }

        private static string ReadCommandText(dynamic oleDb)
        {
            object value = ReadRequired(
                () => (object?)oleDb.CommandText,
                "Excel did not expose the generated connection command.");
            if (value is string text)
            {
                return text;
            }

            if (value is Array array && array.Length == 1)
            {
                return Convert.ToString(
                           array.GetValue(array.GetLowerBound(0)),
                           CultureInfo.InvariantCulture) ??
                       string.Empty;
            }

            throw new InvalidOperationException(
                "Excel exposed an unsupported generated connection command.");
        }

        private static bool TryResolveWorksheetRange(
            object workbook,
            string sourceToken,
            out object? nativeRange,
            out string? canonicalReference)
        {
            nativeRange = null;
            canonicalReference = null;
            string token = sourceToken.Trim();
            int separator = token.LastIndexOf('!');
            if (separator <= 0 || separator == token.Length - 1)
            {
                return false;
            }

            if (token.IndexOf(',') >= 0 ||
                token.IndexOf(';') >= 0 ||
                token.IndexOf('!') != separator)
            {
                throw new NotSupportedException(
                    "Multi-area and 3D PivotTable source ranges cannot be upgraded transactionally.");
            }

            string sheetToken = token.Substring(0, separator).Trim();
            string addressToken = token.Substring(separator + 1).Trim();
            string worksheetName = ParseLocalWorksheetName(sheetToken);
            if (!TryParseBoundedRangeAddress(
                    addressToken,
                    out string absoluteA1,
                    out int expectedRows,
                    out int expectedColumns))
            {
                throw new NotSupportedException(
                    "The classic PivotTable source is not a bounded absolute A1 or R1C1 worksheet range.");
            }

            dynamic book = workbook;
            object worksheets = ReadRequired(
                () => (object?)book.Worksheets,
                "Excel did not expose workbook worksheets while resolving the raw source.");
            object worksheetObject = FindNamedObject(
                worksheets,
                worksheetName,
                "worksheets") ??
                throw new NotSupportedException(
                    "The classic PivotTable source worksheet could not be resolved in this workbook.");
            dynamic worksheet = worksheetObject;
            object rangeObject = TryGet(() => (object?)worksheet.Range[absoluteA1]) ??
                TryGet(() => (object?)worksheet.Range(absoluteA1)) ??
                throw new NotSupportedException(
                    "Excel could not resolve the classic PivotTable source range.");
            DemandBoundedResolvedRange(
                rangeObject,
                worksheetName,
                expectedRows,
                expectedColumns);

            nativeRange = rangeObject;
            canonicalReference = "='" + worksheetName.Replace("'", "''") +
                                 "'!" + absoluteA1;
            return true;
        }

        private static string ParseLocalWorksheetName(string token)
        {
            if (string.IsNullOrWhiteSpace(token) ||
                token.IndexOf('[') >= 0 ||
                token.IndexOf(']') >= 0 ||
                token.IndexOf(':') >= 0 ||
                token.IndexOf('\\') >= 0 ||
                token.IndexOf('/') >= 0)
            {
                throw new NotSupportedException(
                    "External and 3D PivotTable source ranges cannot be upgraded transactionally.");
            }

            string worksheetName;
            if (token[0] == '\'')
            {
                if (token.Length < 2 || token[token.Length - 1] != '\'')
                {
                    throw new NotSupportedException(
                        "The classic PivotTable source has an invalid worksheet reference.");
                }

                worksheetName = token.Substring(1, token.Length - 2).Replace("''", "'");
            }
            else
            {
                if (token.IndexOf('\'') >= 0)
                {
                    throw new NotSupportedException(
                        "The classic PivotTable source has an invalid worksheet reference.");
                }

                worksheetName = token;
            }

            if (string.IsNullOrWhiteSpace(worksheetName) || worksheetName.Length > 31)
            {
                throw new NotSupportedException(
                    "The classic PivotTable source worksheet name is invalid.");
            }

            return worksheetName;
        }

        private static bool TryParseBoundedRangeAddress(
            string address,
            out string absoluteA1,
            out int rowCount,
            out int columnCount)
        {
            absoluteA1 = string.Empty;
            rowCount = 0;
            columnCount = 0;
            Match a1 = A1RangePattern.Match(address);
            int firstRow;
            int firstColumn;
            int lastRow;
            int lastColumn;
            if (a1.Success)
            {
                firstColumn = ColumnNumber(a1.Groups[1].Value);
                firstRow = ParsePositiveInt(a1.Groups[2].Value);
                lastColumn = a1.Groups[3].Success
                    ? ColumnNumber(a1.Groups[3].Value)
                    : firstColumn;
                lastRow = a1.Groups[4].Success
                    ? ParsePositiveInt(a1.Groups[4].Value)
                    : firstRow;
            }
            else
            {
                Match r1c1 = R1C1RangePattern.Match(address);
                if (!r1c1.Success) return false;
                firstRow = ParsePositiveInt(r1c1.Groups[1].Value);
                firstColumn = ParsePositiveInt(r1c1.Groups[2].Value);
                lastRow = r1c1.Groups[3].Success
                    ? ParsePositiveInt(r1c1.Groups[3].Value)
                    : firstRow;
                lastColumn = r1c1.Groups[4].Success
                    ? ParsePositiveInt(r1c1.Groups[4].Value)
                    : firstColumn;
            }

            if (firstRow > ExcelMaximumRows || lastRow > ExcelMaximumRows ||
                firstColumn > ExcelMaximumColumns || lastColumn > ExcelMaximumColumns)
            {
                throw new NotSupportedException(
                    "The classic PivotTable source exceeds Excel's worksheet bounds.");
            }

            int top = Math.Min(firstRow, lastRow);
            int bottom = Math.Max(firstRow, lastRow);
            int left = Math.Min(firstColumn, lastColumn);
            int right = Math.Max(firstColumn, lastColumn);
            rowCount = bottom - top + 1;
            columnCount = right - left + 1;
            long cells = (long)rowCount * columnCount;
            if (rowCount == ExcelMaximumRows ||
                columnCount == ExcelMaximumColumns ||
                cells > MaximumRawRangeCells)
            {
                throw new NotSupportedException(
                    "Entire-row, entire-column, and oversized PivotTable source ranges are not supported.");
            }

            absoluteA1 = AbsoluteA1(top, left) +
                         (top == bottom && left == right
                             ? string.Empty
                             : ":" + AbsoluteA1(bottom, right));
            return true;
        }

        private static void DemandBoundedResolvedRange(
            object rangeObject,
            string expectedWorksheetName,
            int expectedRows,
            int expectedColumns)
        {
            dynamic range = rangeObject;
            object parent = ReadRequired(
                () => (object?)range.Parent,
                "Excel did not expose the raw source range worksheet.");
            dynamic worksheet = parent;
            string actualWorksheetName = ReadRequiredString(
                () => (object?)worksheet.Name,
                "Excel did not expose the raw source range worksheet name.");
            if (!string.Equals(
                    actualWorksheetName,
                    expectedWorksheetName,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new NotSupportedException(
                    "The raw PivotTable source resolved outside its expected worksheet.");
            }

            object? areas = TryGet(() => (object?)range.Areas);
            if (areas == null || ReadCollectionCount(areas, 2, "range areas") != 1)
            {
                throw new NotSupportedException(
                    "Only a single contiguous raw PivotTable source range is supported.");
            }

            int rows = ReadNestedCollectionCount(range, "Rows");
            int columns = ReadNestedCollectionCount(range, "Columns");
            long cells = ReadOptionalLong(() => (object?)range.Cells.CountLarge, -1L);
            if (rows != expectedRows ||
                columns != expectedColumns ||
                rows == ExcelMaximumRows ||
                columns == ExcelMaximumColumns ||
                cells != (long)rows * columns ||
                cells > MaximumRawRangeCells)
            {
                throw new NotSupportedException(
                    "Excel resolved an oversized or inconsistent raw PivotTable source range.");
            }
        }

        private static int ReadNestedCollectionCount(dynamic parent, string member)
        {
            object? collection = member == "Rows"
                ? TryGet(() => (object?)parent.Rows)
                : TryGet(() => (object?)parent.Columns);
            if (collection == null)
            {
                throw new NotSupportedException(
                    "Excel did not expose the raw source range dimensions.");
            }

            return ReadCollectionCount(
                collection,
                member == "Rows" ? ExcelMaximumRows : ExcelMaximumColumns,
                "range " + member.ToLowerInvariant());
        }

        private static int ParsePositiveInt(string value)
        {
            return int.Parse(value, NumberStyles.None, CultureInfo.InvariantCulture);
        }

        private static int ColumnNumber(string letters)
        {
            var value = 0;
            foreach (char character in letters.ToUpperInvariant())
            {
                value = checked(value * 26 + (character - 'A' + 1));
            }

            return value;
        }

        private static string AbsoluteA1(int row, int column)
        {
            var letters = string.Empty;
            var remaining = column;
            while (remaining > 0)
            {
                remaining--;
                letters = (char)('A' + remaining % 26) + letters;
                remaining /= 26;
            }

            return "$" + letters + "$" + row.ToString(CultureInfo.InvariantCulture);
        }

        private static bool TryResolveTable(
            object workbook,
            string sourceToken,
            out string? tableName)
        {
            tableName = null;
            dynamic book = workbook;
            object worksheetsObject = ReadRequired(
                () => (object?)book.Worksheets,
                "Excel did not expose the workbook worksheets.");
            foreach (object worksheetObject in ReadCollection(
                         worksheetsObject,
                         MaximumWorksheets,
                         "worksheets"))
            {
                dynamic worksheet = worksheetObject;
                object? listObjects = TryGet(() => (object?)worksheet.ListObjects);
                if (listObjects == null)
                {
                    continue;
                }

                foreach (object tableObject in ReadCollection(
                             listObjects,
                             MaximumWorkbookObjects,
                             "tables"))
                {
                    dynamic table = tableObject;
                    string name = ReadOptionalString(() => (object?)table.Name);
                    string displayName = ReadOptionalString(() => (object?)table.DisplayName);
                    if (SourceTokenMatches(sourceToken, name) ||
                        SourceTokenMatches(sourceToken, displayName))
                    {
                        tableName = !string.IsNullOrWhiteSpace(name) ? name : displayName;
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool TryResolveWorkbookName(
            object workbook,
            string sourceToken,
            out string? rangeName)
        {
            rangeName = null;
            dynamic book = workbook;
            object namesObject = ReadRequired(
                () => (object?)book.Names,
                "Excel did not expose the workbook names.");
            foreach (object nameObject in ReadCollection(
                         namesObject,
                         MaximumWorkbookObjects,
                         "names"))
            {
                dynamic name = nameObject;
                string fullName = ReadOptionalString(() => (object?)name.Name);
                string shortName = WorkbookScopedName(fullName);
                if (string.IsNullOrWhiteSpace(shortName) ||
                    (!SourceTokenMatches(sourceToken, fullName) &&
                     !SourceTokenMatches(sourceToken, shortName)))
                {
                    continue;
                }

                object? refersToRange = TryGet(() => (object?)name.RefersToRange);
                if (refersToRange == null)
                {
                    continue;
                }

                DemandSafeNamedRange(workbook, refersToRange);
                rangeName = shortName;
                return true;
            }

            return false;
        }

        private static void DemandSafeNamedRange(
            object workbook,
            object rangeObject)
        {
            dynamic range = rangeObject;
            object worksheetObject = ReadRequired(
                () => (object?)range.Parent,
                "Excel did not expose the named-range worksheet.");
            dynamic worksheet = worksheetObject;
            object rangeWorkbook = ReadRequired(
                () => (object?)worksheet.Parent,
                "Excel did not expose the named-range workbook.");
            if (!SameNativeObject(workbook, rangeWorkbook))
            {
                throw new NotSupportedException(
                    "External workbook names cannot be upgraded to the Data Model.");
            }

            object areas = ReadRequired(
                () => (object?)range.Areas,
                "Excel did not expose named-range areas.");
            if (ReadCollectionCount(areas, 2, "named-range areas") != 1)
            {
                throw new NotSupportedException(
                    "Only a single contiguous workbook-scoped named range can be upgraded to the Data Model.");
            }

            int rows = ReadNestedCollectionCount(range, "Rows");
            int columns = ReadNestedCollectionCount(range, "Columns");
            long cells = ReadRequiredLong(
                () => (object?)range.Cells.CountLarge,
                "named-range cell count");
            if (rows <= 0 || columns <= 0 ||
                rows == ExcelMaximumRows ||
                columns == ExcelMaximumColumns ||
                cells != (long)rows * columns ||
                cells > MaximumRawRangeCells)
            {
                throw new NotSupportedException(
                    "Entire-row, entire-column, oversized, or inconsistent workbook names cannot be upgraded to the Data Model.");
            }
        }

        private static string WorkbookScopedName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            string unquoted = value.Trim().Trim('\'');
            int separator = unquoted.LastIndexOf('!');
            if (separator < 0)
            {
                return unquoted;
            }

            string prefix = unquoted.Substring(0, separator).Trim('\'');
            string candidate = unquoted.Substring(separator + 1);
            // A sheet-scoped name cannot be addressed unambiguously by
            // Excel.CurrentWorkbook using only the leaf name.
            if (!prefix.StartsWith("[", StringComparison.Ordinal) &&
                prefix.IndexOf('.') < 0)
            {
                return string.Empty;
            }

            return candidate;
        }

        private static bool SourceTokenMatches(string sourceToken, string candidate)
        {
            if (string.IsNullOrWhiteSpace(candidate))
            {
                return false;
            }

            string normalized = sourceToken.Trim().Trim('\'');
            if (string.Equals(normalized, candidate, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            int separator = normalized.LastIndexOf('!');
            return separator >= 0 &&
                   string.Equals(
                       normalized.Substring(separator + 1),
                       candidate,
                       StringComparison.OrdinalIgnoreCase);
        }

        private static string ReadClassicSourceToken(dynamic cache)
        {
            object value = ReadRequired(
                () => (object?)cache.SourceData,
                "Excel did not expose the classic PivotTable source.");
            if (value is string text && !string.IsNullOrWhiteSpace(text))
            {
                return text;
            }

            if (value is Array array && array.Length == 1)
            {
                string converted = Convert.ToString(
                    array.GetValue(array.GetLowerBound(0)),
                    CultureInfo.InvariantCulture) ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(converted))
                {
                    return converted;
                }
            }

            throw new NotSupportedException(
                "Excel exposed an unsupported classic PivotTable source contract.");
        }

        internal static IReadOnlyList<LateBoundFieldState> ReadFieldStates(object pivotObject)
        {
            dynamic pivot = pivotObject;
            var fields = new List<LateBoundFieldState>();
            dynamic cache = ReadPivotCache((object)pivot);
            bool isOlap = ReadBoolean(() => (object?)cache.OLAP, "PivotCache.OLAP");
            object dataFields = ReadRequiredPivotCollection(pivot, "DataFields");
            int dataFieldCount = ReadCollectionCount(
                dataFields,
                MaximumFields,
                "DataFields");
            object? dataPivotField = dataFieldCount > 1
                ? ReadRequired(
                    () => (object?)pivot.DataPivotField,
                    "Excel did not expose the Values pseudo-axis field for a multi-value PivotTable.")
                : null;
            ReadArea(fields, pivot, "RowFields", PivotNativeFieldArea.Row, isOlap, dataPivotField);
            ReadArea(fields, pivot, "ColumnFields", PivotNativeFieldArea.Column, isOlap, dataPivotField);
            ReadArea(fields, pivot, "PageFields", PivotNativeFieldArea.Filter, isOlap, null);
            ReadArea(fields, pivot, "DataFields", PivotNativeFieldArea.Values, isOlap, null);
            if (fields.Count > MaximumFields)
            {
                throw new NotSupportedException(
                    "The PivotTable layout exceeds the bounded conversion limit.");
            }

            DemandNoDuplicateImplicitValues(fields);

            return fields;
        }

        internal static LateBoundDataAxisState ReadDataAxisState(object pivotObject)
        {
            dynamic pivot = pivotObject;
            object dataFields = ReadRequiredPivotCollection(pivot, "DataFields");
            int count = ReadCollectionCount(
                dataFields,
                MaximumFields,
                "DataFields");
            if (count <= 1)
            {
                return LateBoundDataAxisState.Hidden;
            }

            object fieldObject = ReadRequired(
                () => (object?)pivot.DataPivotField,
                "Excel did not expose the Values pseudo-axis field for a multi-value PivotTable.");
            dynamic field = fieldObject;
            int orientation = ReadRequiredInt(
                () => (object?)field.Orientation,
                "Values pseudo-axis orientation");
            PivotValuesAxis axis;
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
                    "A multi-value PivotTable must expose its Values pseudo-axis on rows or columns.");
            }

            return new LateBoundDataAxisState(
                axis,
                ReadRequiredPositiveInt(
                    () => (object?)field.Position,
                    "Values pseudo-axis position"));
        }

        private static object ReadRequiredPivotCollection(
            dynamic pivot,
            string collectionName)
        {
            object? collection = null;
            if (string.Equals(collectionName, "DataFields", StringComparison.Ordinal))
            {
                collection = TryGet(() => (object?)pivot.DataFields) ??
                    TryGet(() => (object?)pivot.DataFields());
            }

            return collection ?? throw new NotSupportedException(
                "Excel did not expose the required " + collectionName +
                " collection for a reversible conversion.");
        }

        internal static void DemandNoDuplicateImplicitValues(
            IEnumerable<LateBoundFieldState> fields)
        {
            if (fields == null) throw new ArgumentNullException(nameof(fields));
            bool duplicate = fields
                .Where(field => field.Area == PivotNativeFieldArea.Values)
                .GroupBy(
                    field => FieldLeaf(field.SourceName) + "\u001f" +
                             (field.Function?.ToString(CultureInfo.InvariantCulture) ?? string.Empty),
                    StringComparer.OrdinalIgnoreCase)
                .Any(group => group.Count() > 1);
            if (duplicate)
            {
                throw new NotSupportedException(
                    "The classic PivotTable places the same source field and aggregation more than once. Data Model conversion requires separately authored measures for independent captions.");
            }
        }

        private static void ReadArea(
            ICollection<LateBoundFieldState> output,
            dynamic pivot,
            string collectionName,
            PivotNativeFieldArea area,
            bool isOlap,
            object? excludedField)
        {
            object? collection = null;
            switch (collectionName)
            {
                case "RowFields":
                    collection = TryGet(() => (object?)pivot.RowFields) ??
                        TryGet(() => (object?)pivot.RowFields());
                    break;
                case "ColumnFields":
                    collection = TryGet(() => (object?)pivot.ColumnFields) ??
                        TryGet(() => (object?)pivot.ColumnFields());
                    break;
                case "PageFields":
                    collection = TryGet(() => (object?)pivot.PageFields) ??
                        TryGet(() => (object?)pivot.PageFields());
                    break;
                case "DataFields":
                    collection = TryGet(() => (object?)pivot.DataFields) ??
                        TryGet(() => (object?)pivot.DataFields());
                    break;
            }

            if (collection == null)
            {
                throw new NotSupportedException(
                    "Excel did not expose the required " + collectionName +
                    " collection for a reversible conversion.");
            }

            int fallbackPosition = 1;
            foreach (object fieldObject in ReadCollection(
                         collection,
                         MaximumFields,
                         collectionName))
            {
                if (excludedField != null &&
                    SameNativeObject(fieldObject, excludedField))
                {
                    continue;
                }

                dynamic field = fieldObject;
                DemandNoNativePivotFilters(field);
                DemandNoCalculatedOrGroupedField((object)field);
                if (area == PivotNativeFieldArea.Values)
                {
                    DemandPlainValueCalculation((object)field);
                    DemandSupportedModelAggregation((object)field);
                }
                else if (isOlap)
                {
                    DemandDataModelManualSort(field);
                }
                else
                {
                    DemandClassicManualSort(field);
                    DemandClassicNoShowAllItems(field);
                }
                string name = ReadRequiredString(
                    () => (object?)field.Name,
                    "Excel exposed an unnamed PivotTable field.");
                string sourceName = area == PivotNativeFieldArea.Values
                    ? ReadValueSourceName(field)
                    : ReadRequiredString(
                        () => (object?)field.SourceName,
                        "Excel did not expose a PivotField source name.");
                if (string.IsNullOrWhiteSpace(sourceName))
                {
                    sourceName = ReadNestedSourceName(field);
                }

                if (string.IsNullOrWhiteSpace(sourceName)) sourceName = name;
                string caption = ReadRequiredString(
                    () => (object?)field.Caption,
                    "Excel did not expose a PivotField caption.");
                int position = ReadRequiredPositiveInt(
                    () => (object?)field.Position,
                    "PivotField position");
                int? function = area == PivotNativeFieldArea.Values
                    ? ReadRequiredInt(
                        () => (object?)field.Function,
                        "value aggregation function")
                    : null;
                string numberFormat = area == PivotNativeFieldArea.Values
                    ? ReadRequiredOptionalString(
                        () => (object?)field.NumberFormat,
                        "value number format")
                    : string.Empty;
                bool repeatLabels = area == PivotNativeFieldArea.Row &&
                    ReadRequiredBoolean(
                        () => (object?)field.RepeatLabels,
                        "row-field RepeatLabels");
                IReadOnlyList<bool> subtotals = area == PivotNativeFieldArea.Values
                    ? Array.Empty<bool>()
                    : ReadSubtotals(field);
                IReadOnlyList<LateBoundMemberState> members =
                    area == PivotNativeFieldArea.Values
                        ? Array.Empty<LateBoundMemberState>()
                        : ReadMemberStates(field);
                bool layoutBlankLine =
                    area == PivotNativeFieldArea.Row ||
                    area == PivotNativeFieldArea.Column
                        ? ReadRequiredBoolean(
                            () => (object?)field.LayoutBlankLine,
                            "PivotField blank-line layout state")
                        : false;
                bool layoutPageBreak =
                    area == PivotNativeFieldArea.Row ||
                    area == PivotNativeFieldArea.Column
                        ? ReadRequiredBoolean(
                            () => (object?)field.LayoutPageBreak,
                            "PivotField page-break layout state")
                        : false;
                output.Add(new LateBoundFieldState(
                    area,
                    sourceName,
                    name,
                    string.IsNullOrWhiteSpace(caption) ? name : caption,
                    position,
                    function,
                    numberFormat,
                    repeatLabels,
                    subtotals,
                    members,
                    area == PivotNativeFieldArea.Filter
                        ? ReadCurrentPage(field, isOlap)
                        : string.Empty,
                    area == PivotNativeFieldArea.Filter &&
                        ReadRequiredBoolean(
                            () => (object?)field.EnableMultiplePageItems,
                            "page-field multiple-selection state"),
                    layoutBlankLine,
                    layoutPageBreak));
                fallbackPosition++;
            }
        }

        internal static void DemandNoNativePivotFilters(object fieldObject)
        {
            dynamic field = fieldObject;
            object? filters;
            if (!PivotLateBound.TryRead(
                    () => (object?)field.PivotFilters,
                    out filters) ||
                filters == null)
            {
                throw new NotSupportedException(
                    "Excel did not expose PivotFilters needed for a reversible conversion preflight.");
            }

            int count = ReadCollectionCount(filters, MaximumWorkbookObjects, "PivotFilters");
            if (count > 0)
            {
                throw new NotSupportedException(
                    "The selected PivotTable uses label, value, or date filters that cannot yet be restored transactionally. Remove those filters before enabling the Data Model.");
            }
        }

        internal static void DemandClassicManualSort(object fieldObject)
        {
            dynamic field = fieldObject;
            int order = ReadRequiredInt(
                () => (object?)field.AutoSortOrder,
                "classic PivotField AutoSortOrder");
            if (order != SortManual)
            {
                throw new NotSupportedException(
                    "Automatically sorted classic PivotFields cannot yet be converted transactionally. Set the field to manual order before enabling the Data Model.");
            }
        }

        internal static void DemandClassicNoShowAllItems(object fieldObject)
        {
            dynamic field = fieldObject;
            if (ReadRequiredBoolean(
                    () => (object?)field.ShowAllItems,
                    "classic PivotField ShowAllItems"))
            {
                throw new NotSupportedException(
                    "A classic PivotField with Show items with no data enabled cannot be preserved because Data Model PivotFields always disable ShowAllItems.");
            }
        }

        internal static void DemandClassicDefaultIncludeNewItemsInFilter(
            object fieldObject)
        {
            if (fieldObject == null) throw new ArgumentNullException(nameof(fieldObject));
            dynamic field = fieldObject;
            if (ReadRequiredBoolean(
                    () => (object?)field.IncludeNewItemsInFilter,
                    "classic PivotField IncludeNewItemsInFilter"))
            {
                throw new NotSupportedException(
                    "A classic PivotField that automatically includes new source items in its manual filter cannot yet be converted transactionally. Data Model filtering tracks this policy on the CubeField and the replacement does not mutate that shared hierarchy setting.");
            }
        }

        internal static void DemandCompatibleClassicCachePolicy(object pivotTable)
        {
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            dynamic cache = ReadPivotCache(pivotTable);
            if (ReadRequiredBoolean(
                    () => (object?)cache.RefreshOnFileOpen,
                    "classic PivotCache RefreshOnFileOpen"))
            {
                throw new NotSupportedException(
                    "A classic PivotCache that refreshes when the workbook opens cannot yet be converted without changing refresh policy.");
            }

            if (!ReadRequiredBoolean(
                    () => (object?)cache.EnableRefresh,
                    "classic PivotCache EnableRefresh"))
            {
                throw new NotSupportedException(
                    "A classic PivotCache with user refresh disabled cannot yet be converted without changing refresh policy.");
            }

            int missingItemsLimit = ReadRequiredInt(
                () => (object?)cache.MissingItemsLimit,
                "classic PivotCache MissingItemsLimit");
            if (missingItemsLimit != MissingItemsDefault)
            {
                throw new NotSupportedException(
                    "A classic PivotCache with a nondefault missing-item retention limit cannot be preserved by a Data Model PivotCache.");
            }
        }

        internal static void DemandCompatibleClassicSaveData(object pivotTable)
        {
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            dynamic pivot = pivotTable;
            // Read fail-closed so the classic definition is known before any
            // mutation. OLAP PivotTables always report SaveData=false, but the
            // embedded workbook Data Model remains their offline saved source;
            // this documented category invariant is therefore normalized rather
            // than treated as classic cache-data loss.
            _ = ReadRequiredBoolean(
                () => (object?)pivot.SaveData,
                "classic PivotTable SaveData");
        }

        private static void DemandDataModelManualSort(dynamic field)
        {
            if (ReadRequiredBoolean(
                    () => (object?)field.DatabaseSort,
                    "Data Model PivotField DatabaseSort"))
            {
                throw new InvalidOperationException(
                    "The restored Data Model PivotField did not retain manual member ordering.");
            }
        }

        internal static void DemandNoCalculatedOrGroupedField(object fieldObject)
        {
            dynamic field = fieldObject;
            if (!PivotLateBound.TryRead(
                    () => (object?)field.IsCalculated,
                    out object? calculatedValue) ||
                calculatedValue == null)
            {
                throw new NotSupportedException(
                    "Excel did not expose PivotField.IsCalculated for a reversible conversion preflight.");
            }

            bool isCalculated;
            try
            {
                isCalculated = Convert.ToBoolean(
                    calculatedValue,
                    CultureInfo.InvariantCulture);
            }
            catch (Exception exception) when (
                exception is FormatException ||
                exception is InvalidCastException)
            {
                throw new NotSupportedException(
                    "Excel exposed an invalid PivotField.IsCalculated value.",
                    exception);
            }

            if (isCalculated)
            {
                throw new NotSupportedException(
                    "Calculated PivotFields cannot yet be translated to a Data Model PivotTable transactionally.");
            }

            object? calculatedItems = TryGet(() => (object?)field.CalculatedItems());
            calculatedItems = calculatedItems ?? TryGet(() => (object?)field.CalculatedItems);
            if (calculatedItems != null &&
                ReadCollectionCount(
                    calculatedItems,
                    MaximumMembers,
                    "CalculatedItems") > 0)
            {
                throw new NotSupportedException(
                    "Calculated PivotItems cannot yet be translated to a Data Model PivotTable transactionally.");
            }

            if (!PivotLateBound.TryRead(
                    () => (object?)field.ParentField,
                    out object? parentField) ||
                !PivotLateBound.TryRead(
                    () => (object?)field.ChildField,
                    out object? childField))
            {
                throw new NotSupportedException(
                    "Excel did not expose PivotField grouping relationships for a reversible conversion preflight.");
            }

            if (parentField != null || childField != null)
            {
                throw new NotSupportedException(
                    "Grouped PivotFields cannot yet be translated to a Data Model PivotTable transactionally.");
            }
        }

        internal static void DemandNoUnsupportedClassicDefinitions(object pivotTable)
        {
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            dynamic pivot = pivotTable;
            object? calculatedFields;
            bool calculatedFieldsRead = PivotLateBound.TryRead(
                () => (object?)pivot.CalculatedFields(),
                out calculatedFields);
            if (!calculatedFieldsRead || calculatedFields == null)
            {
                calculatedFieldsRead = PivotLateBound.TryRead(
                    () => (object?)pivot.CalculatedFields,
                    out calculatedFields);
            }

            if (!calculatedFieldsRead || calculatedFields == null)
            {
                throw new NotSupportedException(
                    "Excel did not expose CalculatedFields needed for a reversible conversion preflight.");
            }

            if (ReadCollectionCount(
                    calculatedFields,
                    MaximumFields,
                    "CalculatedFields") > 0)
            {
                throw new NotSupportedException(
                    "Calculated PivotFields cannot yet be translated to a Data Model PivotTable transactionally.");
            }

            object? allFields = TryGet(() => (object?)pivot.PivotFields);
            allFields = allFields ?? TryGet(() => (object?)pivot.PivotFields());
            if (allFields == null)
            {
                throw new NotSupportedException(
                    "Excel did not expose the classic PivotFields needed for a reversible conversion.");
            }

            foreach (object field in ReadCollection(
                         allFields,
                         MaximumFields,
                         "PivotFields"))
            {
                DemandNoCalculatedOrGroupedField(field);
                DemandClassicDefaultIncludeNewItemsInFilter(field);
                dynamic nativeField = field;
                object? calculatedItems = TryGet(
                    () => (object?)nativeField.CalculatedItems());
                calculatedItems = calculatedItems ??
                    TryGet(() => (object?)nativeField.CalculatedItems);
                if (calculatedItems == null)
                {
                    throw new NotSupportedException(
                        "Excel did not expose CalculatedItems needed for a reversible conversion preflight.");
                }
            }
        }

        internal static void DemandNoUnsupportedCustomFormatting(
            object pivotTable,
            object tableRange)
        {
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            if (tableRange == null) throw new ArgumentNullException(nameof(tableRange));
            dynamic range = tableRange;
            object conditions = ReadRequired(
                () => (object?)range.FormatConditions,
                "Excel did not expose PivotTable conditional formats safely.");
            if (ReadRequiredCollectionCount(
                    conditions,
                    MaximumFields,
                    "PivotTable conditional formats") > 0)
            {
                throw new NotSupportedException(
                    "PivotTable conditional formatting cannot yet be restored transactionally. Remove conditional-format rules before Data Model conversion.");
            }
        }

        internal static void DemandNoUnsupportedCellMetadata(object tableRange)
        {
            if (tableRange == null) throw new ArgumentNullException(nameof(tableRange));
            dynamic range = tableRange;
            int firstRow = ReadRequiredPositiveInt(
                () => (object?)range.Row,
                "PivotTable result first row");
            int firstColumn = ReadRequiredPositiveInt(
                () => (object?)range.Column,
                "PivotTable result first column");
            int rowCount = ReadRequiredCollectionCount(
                ReadRequired(
                    () => (object?)range.Rows,
                    "Excel did not expose the PivotTable result rows while checking cell metadata."),
                MaximumPivotResultCells,
                "PivotTable result rows");
            int columnCount = ReadRequiredCollectionCount(
                ReadRequired(
                    () => (object?)range.Columns,
                    "Excel did not expose the PivotTable result columns while checking cell metadata."),
                MaximumPivotResultCells,
                "PivotTable result columns");
            if (rowCount <= 0 || columnCount <= 0 ||
                (long)rowCount * columnCount > MaximumPivotResultCells)
            {
                throw new NotSupportedException(
                    "The PivotTable result exceeds the bounded cell-metadata preflight limit.");
            }

            object hyperlinks = ReadRequired(
                () => (object?)range.Hyperlinks,
                "Excel did not expose PivotTable hyperlinks safely.");
            if (ReadRequiredCollectionCount(
                    hyperlinks,
                    MaximumPivotResultCells,
                    "PivotTable hyperlinks") > 0)
            {
                throw new NotSupportedException(
                    "PivotTable result cells containing hyperlinks cannot yet be restored transactionally. Remove them before Data Model conversion.");
            }

            object worksheetObject = ReadRequired(
                () => (object?)range.Parent,
                "Excel did not expose the PivotTable worksheet while checking cell metadata.");
            dynamic worksheet = worksheetObject;
            object legacyComments = ReadRequired(
                () => (object?)worksheet.Comments,
                "Excel did not expose the worksheet notes collection safely.");
            DemandNoCommentsInRange(
                legacyComments,
                "worksheet notes",
                firstRow,
                firstColumn,
                rowCount,
                columnCount);

            object threadedComments = ReadRequired(
                () => (object?)worksheet.CommentsThreaded,
                "Excel did not expose the worksheet comments collection safely.");
            DemandNoCommentsInRange(
                threadedComments,
                "worksheet comments",
                firstRow,
                firstColumn,
                rowCount,
                columnCount);

            DemandNoDataValidation(range);
        }

        private static void DemandNoCommentsInRange(
            object comments,
            string label,
            int firstRow,
            int firstColumn,
            int rowCount,
            int columnCount)
        {
            int lastRow = checked(firstRow + rowCount - 1);
            int lastColumn = checked(firstColumn + columnCount - 1);
            foreach (object commentObject in ReadCollection(
                         comments,
                         MaximumPivotResultCells,
                         label))
            {
                dynamic comment = commentObject;
                object parentRange = ReadRequired(
                    () => (object?)comment.Parent,
                    "Excel did not expose a cell for an entry in the " + label + " collection.");
                dynamic cell = parentRange;
                int row = ReadRequiredPositiveInt(
                    () => (object?)cell.Row,
                    label + " row");
                int column = ReadRequiredPositiveInt(
                    () => (object?)cell.Column,
                    label + " column");
                if (row >= firstRow && row <= lastRow &&
                    column >= firstColumn && column <= lastColumn)
                {
                    throw new NotSupportedException(
                        "PivotTable result cells containing notes or comments cannot yet be restored transactionally. Remove them before Data Model conversion.");
                }
            }
        }

        private static void DemandNoDataValidation(dynamic range)
        {
            try
            {
                object? validationCells = range.SpecialCells(CellTypeAllValidation);
                if (validationCells == null)
                {
                    throw new NotSupportedException(
                        "Excel returned an unreadable data-validation range during conversion preflight.");
                }

                throw new NotSupportedException(
                    "PivotTable result cells containing data validation cannot yet be restored transactionally. Remove validation before Data Model conversion.");
            }
            catch (Exception exception) when (IsNoCellsFound(exception))
            {
                // Excel reports an empty SpecialCells result with this precise
                // COM error. Other dispatch/RPC failures remain fail-closed.
            }
            catch (NotSupportedException)
            {
                throw;
            }
            catch (Exception exception)
            {
                throw new NotSupportedException(
                    "Excel did not expose PivotTable data-validation state safely.",
                    exception);
            }
        }

        private static bool IsNoCellsFound(Exception exception)
        {
            var comException = exception as COMException;
            if (comException == null ||
                comException.ErrorCode != ExcelNoCellsFoundError)
            {
                return false;
            }

            string message = (comException.Message ?? string.Empty).Trim();
            return string.Equals(
                       message,
                       "No cells were found.",
                       StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(
                       message,
                       "No cells were found",
                       StringComparison.OrdinalIgnoreCase);
        }

        internal static void DemandPlainValueCalculation(object fieldObject)
        {
            dynamic field = fieldObject;
            object value = ReadRequired(
                () => (object?)field.Calculation,
                "Excel did not expose the value field's Show Values As calculation.");
            int calculation;
            try
            {
                calculation = Convert.ToInt32(value, CultureInfo.InvariantCulture);
            }
            catch (Exception exception) when (
                exception is FormatException ||
                exception is InvalidCastException ||
                exception is OverflowException)
            {
                throw new NotSupportedException(
                    "Excel exposed an unsupported Show Values As calculation.",
                    exception);
            }

            if (calculation != -4143)
            {
                throw new NotSupportedException(
                    "Show Values As fields cannot yet be translated to a Data Model PivotTable transactionally.");
            }
        }

        internal static void DemandSupportedModelAggregation(object fieldObject)
        {
            if (fieldObject == null) throw new ArgumentNullException(nameof(fieldObject));
            dynamic field = fieldObject;
            int function = ReadRequiredInt(
                () => (object?)field.Function,
                "value aggregation function");
            if (function != -4157 && // Sum
                function != -4112 && // Count
                function != -4106 && // Average
                function != -4139 && // Minimum
                function != -4136)   // Maximum
            {
                throw new NotSupportedException(
                    "This classic value aggregation cannot be recreated by CubeFields.GetMeasure. Use Sum, Count, Average, Min, or Max before enabling the Data Model.");
            }
        }

        private static string ReadValueSourceName(dynamic field)
        {
            string sourceName = ReadOptionalString(() => (object?)field.SourceNameStandard);
            if (!string.IsNullOrWhiteSpace(sourceName))
            {
                return sourceName;
            }

            object? sourceField = TryGet(() => (object?)field.PivotField);
            if (sourceField != null)
            {
                dynamic nested = sourceField;
                string nestedName = ReadOptionalString(
                    () => (object?)nested.SourceNameStandard);
                if (string.IsNullOrWhiteSpace(nestedName))
                {
                    nestedName = ReadOptionalString(() => (object?)nested.SourceName);
                }

                if (!string.IsNullOrWhiteSpace(nestedName))
                {
                    return nestedName;
                }
            }

            return ReadOptionalString(() => (object?)field.SourceName);
        }

        private static string ReadNestedSourceName(dynamic dataField)
        {
            object? sourceField = TryGet(() => (object?)dataField.PivotField);
            if (sourceField == null)
            {
                return string.Empty;
            }

            dynamic nested = sourceField;
            return ReadOptionalString(() => (object?)nested.SourceName);
        }

        private static string ReadCurrentPage(dynamic field, bool isOlap)
        {
            if (isOlap)
            {
                return ReadRequiredOptionalString(
                    () => (object?)field.CurrentPageName,
                    "OLAP page-field CurrentPageName");
            }

            object currentPage = ReadRequired(
                () => (object?)field.CurrentPage,
                "Excel did not expose the classic page-field CurrentPage PivotItem.");
            if (currentPage is string text)
            {
                return text;
            }

            dynamic item = currentPage;
            string sourceName = ReadOptionalString(() => (object?)item.SourceName);
            if (!string.IsNullOrWhiteSpace(sourceName)) return sourceName;
            string name = ReadOptionalString(() => (object?)item.Name);
            if (!string.IsNullOrWhiteSpace(name)) return name;
            string caption = ReadOptionalString(() => (object?)item.Caption);
            if (!string.IsNullOrWhiteSpace(caption)) return caption;
            throw new NotSupportedException(
                "Excel exposed an unnamed classic page-field PivotItem.");
        }

        private static IReadOnlyList<bool> ReadSubtotals(dynamic field)
        {
            var result = new bool[12];
            for (var index = 1; index <= result.Length; index++)
            {
                int capturedIndex = index;
                result[index - 1] = ReadRequiredBoolean(
                    () => (object?)field.Subtotals[capturedIndex],
                    "PivotField subtotal state");
            }

            return result;
        }

        private static IReadOnlyList<LateBoundMemberState> ReadMemberStates(dynamic field)
        {
            object? items = TryGet(() => (object?)field.PivotItems);
            if (items == null)
            {
                items = TryGet(() => (object?)field.PivotItems());
            }

            if (items == null)
            {
                throw new NotSupportedException(
                    "Excel did not expose PivotItems needed for a reversible member-order and filter snapshot.");
            }

            var result = new List<LateBoundMemberState>();
            int position = 1;
            foreach (object itemObject in ReadCollection(items, MaximumMembers, "PivotItems"))
            {
                dynamic item = itemObject;
                if (ReadOptionalBoolean(() => (object?)item.IsCalculated, false))
                {
                    throw new NotSupportedException(
                        "Calculated PivotItems cannot yet be translated to a Data Model PivotTable transactionally.");
                }
                string name = ReadOptionalString(
                    () => (object?)item.SourceNameStandard);
                if (string.IsNullOrWhiteSpace(name))
                {
                    name = ReadOptionalString(() => (object?)item.SourceName);
                }

                if (string.IsNullOrWhiteSpace(name))
                {
                    name = ReadOptionalString(() => (object?)item.Name);
                }
                string caption = ReadOptionalString(() => (object?)item.Caption);
                if (string.IsNullOrWhiteSpace(name) && string.IsNullOrWhiteSpace(caption))
                {
                    throw new NotSupportedException(
                        "Excel exposed an unnamed PivotTable member that cannot be restored safely.");
                }

                result.Add(new LateBoundMemberState(
                    name,
                    string.IsNullOrWhiteSpace(caption) ? name : caption,
                    ReadRequiredBoolean(
                        () => (object?)item.Visible,
                        "PivotItem visibility"),
                    ReadRequiredPositiveInt(
                        () => (object?)item.Position,
                        "PivotItem position")));
                position++;
            }

            return result;
        }

        private static LateBoundStyleState ReadStyleState(dynamic pivot)
        {
            return new LateBoundStyleState(
                ReadRequiredInt(
                    () => (object?)pivot.LayoutRowDefault,
                    "PivotTable row layout"),
                ReadRequiredBoolean(() => (object?)pivot.RowGrand, "row grand-total state"),
                ReadRequiredBoolean(() => (object?)pivot.ColumnGrand, "column grand-total state"),
                ReadRequiredBoolean(
                    () => (object?)pivot.DisplayFieldCaptions,
                    "field-caption state"),
                ReadRequiredOptionalString(
                    () => (object?)pivot.TableStyle2,
                    "PivotTable style"),
                ReadRequiredBoolean(
                    () => (object?)pivot.PreserveFormatting,
                    "preserve-formatting state"),
                ReadRequiredBoolean(
                    () => (object?)pivot.ShowTableStyleRowStripes,
                    "row-stripe state"),
                ReadRequiredBoolean(
                    () => (object?)pivot.ShowTableStyleColumnStripes,
                    "column-stripe state"),
                ReadRequiredBoolean(
                    () => (object?)pivot.DisplayNullString,
                    "null-string display state"),
                ReadRequiredOptionalString(
                    () => (object?)pivot.NullString,
                    "null-string text"),
                ReadRequiredBoolean(
                    () => (object?)pivot.DisplayErrorString,
                    "error-string display state"),
                ReadRequiredOptionalString(
                    () => (object?)pivot.ErrorString,
                    "error-string text"),
                ReadRequiredBoolean(
                    () => (object?)pivot.ShowDrillIndicators,
                    "drill-indicator state"),
                ReadRequiredBoolean(
                    () => (object?)pivot.EnableDrilldown,
                    "drilldown state"),
                ReadRequiredBoolean(
                    () => (object?)pivot.VisualTotals,
                    "visual-totals state"),
                ReadRequiredBoolean(
                    () => (object?)pivot.SubtotalHiddenPageItems,
                    "hidden-page-item subtotal state"),
                ReadRequiredInt(
                    () => (object?)pivot.PageFieldOrder,
                    "page-field order"),
                ReadRequiredInt(
                    () => (object?)pivot.PageFieldWrapCount,
                    "page-field wrap count"),
                ReadRequiredInt(
                    () => (object?)pivot.CompactRowIndent,
                    "compact-row indentation"),
                ReadRequiredBoolean(
                    () => (object?)pivot.MergeLabels,
                    "merged-label state"));
        }

        internal static void DemandCompatibleOlapInvariants(
            IReadOnlyList<LateBoundFieldState> fields,
            LateBoundStyleState style)
        {
            if (fields == null) throw new ArgumentNullException(nameof(fields));
            if (style == null) throw new ArgumentNullException(nameof(style));
            if (!style.EnableDrilldown)
            {
                throw new NotSupportedException(
                    "Excel Data Model PivotTables always enable drilldown. Enable drilldown on the classic PivotTable before conversion so its behavior can be preserved.");
            }

            if (!style.SubtotalHiddenPageItems &&
                fields.Any(field => field.Area == PivotNativeFieldArea.Filter))
            {
                throw new NotSupportedException(
                    "Excel Data Model PivotTables always include hidden page items in subtotals. This classic PivotTable has a page field with incompatible subtotal semantics.");
            }

            if (fields.Any(field => field.Subtotals.Skip(1).Any(value => value)))
            {
                throw new NotSupportedException(
                    "Excel Data Model PivotFields support only Automatic subtotals. Remove custom Sum, Count, Average, or other subtotal functions before conversion.");
            }
        }

        private static void RestoreDataModelState(
            object pivotTable,
            LateBoundPivotState state,
            string modelTableName)
        {
            if (string.IsNullOrWhiteSpace(modelTableName))
            {
                throw new ArgumentException("A model table name is required.", nameof(modelTableName));
            }

            dynamic pivot = pivotTable;
            pivot.ManualUpdate = true;
            ClearLayout(pivot, useCubeFields: true);
            var placed = new Dictionary<LateBoundFieldState, object>();
            foreach (LateBoundFieldState field in state.Fields
                         .Where(item => item.Area != PivotNativeFieldArea.Values)
                         .OrderBy(item => item.Area)
                         .ThenBy(item => item.Position))
            {
                dynamic cubeField = ResolveCubeField(pivot, modelTableName, field.SourceName);
                cubeField.Orientation = OrientationFor(field.Area);
                cubeField.Position = field.Position;
                object targetField = ResolvePlacedPivotField(pivot, cubeField, field);
                ApplyRegularFieldState(targetField, field, isDataModel: true);
                placed[field] = targetField;
            }

            foreach (LateBoundFieldState field in state.Fields
                         .Where(item => item.Area == PivotNativeFieldArea.Values)
                         .OrderBy(item => item.Position))
            {
                dynamic sourceField = ResolveCubeField(pivot, modelTableName, field.SourceName);
                int function = field.Function ?? -4157;
                dynamic measure = pivot.CubeFields.GetMeasure(
                    sourceField,
                    function,
                    field.Caption);
                measure.Orientation = OrientationData;
                dynamic dataField = ResolveDataField(pivot, field.Caption, field.Position);
                TryWrite(() => dataField.Position = field.Position);
                if (!string.IsNullOrWhiteSpace(field.NumberFormat))
                {
                    dataField.NumberFormat = field.NumberFormat;
                }
            }

            RestoreDataAxisState(pivot, state.DataAxis);

            ApplyStyleState(pivot, state.Style, isDataModel: true);
        }

        private static void RestoreClassicState(
            object pivotTable,
            LateBoundPivotState state)
        {
            dynamic pivot = pivotTable;
            pivot.ManualUpdate = true;
            ClearLayout(pivot, useCubeFields: false);
            foreach (LateBoundFieldState field in state.Fields
                         .Where(item => item.Area != PivotNativeFieldArea.Values)
                         .OrderBy(item => item.Area)
                         .ThenBy(item => item.Position))
            {
                dynamic target = ResolveClassicField(pivot, field.SourceName, field.Name);
                target.Orientation = OrientationFor(field.Area);
                target.Position = field.Position;
                ApplyRegularFieldState(target, field, isDataModel: false);
            }

            foreach (LateBoundFieldState field in state.Fields
                         .Where(item => item.Area == PivotNativeFieldArea.Values)
                         .OrderBy(item => item.Position))
            {
                dynamic source = ResolveClassicField(pivot, field.SourceName, field.Name);
                dynamic dataField = pivot.AddDataField(
                    source,
                    field.Caption,
                    field.Function ?? -4157);
                TryWrite(() => dataField.Position = field.Position);
                if (!string.IsNullOrWhiteSpace(field.NumberFormat))
                {
                    dataField.NumberFormat = field.NumberFormat;
                }
            }

            RestoreDataAxisState(pivot, state.DataAxis);

            ApplyStyleState(pivot, state.Style);
        }

        private static void RestoreDataAxisState(
            dynamic pivot,
            LateBoundDataAxisState dataAxis)
        {
            if (!dataAxis.IsVisible) return;
            object field = ReadRequired(
                () => (object?)pivot.DataPivotField,
                "Excel did not expose the restored Values pseudo-axis field.");
            dynamic nativeField = field;
            nativeField.Orientation = dataAxis.Axis == PivotValuesAxis.Rows
                ? OrientationRow
                : OrientationColumn;
            nativeField.Position = dataAxis.Position;
        }

        private static void ClearLayout(dynamic pivot, bool useCubeFields)
        {
            TryWrite(() => pivot.ClearAllFilters());
            object? fields = useCubeFields
                ? TryGet(() => (object?)pivot.CubeFields)
                : TryGet(() => (object?)pivot.PivotFields);
            if (fields == null)
            {
                return;
            }

            foreach (object fieldObject in ReadCollection(fields, MaximumFields, "fields"))
            {
                dynamic field = fieldObject;
                TryWrite(() => field.Orientation = OrientationHidden);
            }
        }

        private static dynamic ResolveCubeField(
            dynamic pivot,
            string modelTableName,
            string sourceName)
        {
            string leaf = FieldLeaf(sourceName);
            string escapedTable = modelTableName.Replace("]", "]]" );
            string escapedField = leaf.Replace("]", "]]" );
            string[] candidates =
            {
                "[" + escapedTable + "].[" + escapedField + "]",
                "[" + escapedTable + "].[" + escapedField + "].[" + escapedField + "]",
                sourceName
            };
            foreach (string candidate in candidates.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                object? found = TryGet(() => (object?)pivot.CubeFields.Item(candidate));
                if (found != null)
                {
                    return found;
                }
            }

            object fields = ReadRequired(
                () => (object?)pivot.CubeFields,
                "Excel did not expose the Data Model fields.");
            List<object> matches = ReadCollection(fields, MaximumFields, "CubeFields")
                .Where(item => CubeFieldMatches(item, leaf))
                .ToList();
            if (matches.Count == 1)
            {
                return matches[0];
            }

            throw new InvalidOperationException(
                "The Data Model did not expose one unambiguous field for '" + leaf + "'.");
        }

        private static bool CubeFieldMatches(object fieldObject, string leaf)
        {
            dynamic field = fieldObject;
            return string.Equals(
                       FieldLeaf(ReadOptionalString(() => (object?)field.Name)),
                       leaf,
                       StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(
                       ReadOptionalString(() => (object?)field.Caption),
                       leaf,
                       StringComparison.OrdinalIgnoreCase);
        }

        private static dynamic ResolvePlacedPivotField(
            dynamic pivot,
            dynamic cubeField,
            LateBoundFieldState expected)
        {
            string cubeName = ReadOptionalString(() => (object?)cubeField.Name);
            foreach (string candidate in new[] { cubeName, expected.SourceName, expected.Name })
            {
                if (string.IsNullOrWhiteSpace(candidate)) continue;
                object? found = TryGet(() => (object?)pivot.PivotFields.Item(candidate));
                if (found != null) return found;
                found = TryGet(() => (object?)pivot.PivotFields(candidate));
                if (found != null) return found;
            }

            throw new InvalidOperationException(
                "Excel placed a Data Model field but did not expose its PivotField.");
        }

        private static dynamic ResolveClassicField(
            dynamic pivot,
            string sourceName,
            string name)
        {
            foreach (string candidate in new[] { sourceName, name })
            {
                if (string.IsNullOrWhiteSpace(candidate)) continue;
                object? found = TryGet(() => (object?)pivot.PivotFields.Item(candidate));
                if (found != null) return found;
                found = TryGet(() => (object?)pivot.PivotFields(candidate));
                if (found != null) return found;
            }

            throw new InvalidOperationException(
                "The classic PivotCache no longer exposes the field '" + sourceName + "'.");
        }

        private static dynamic ResolveDataField(
            dynamic pivot,
            string caption,
            int position)
        {
            object? byCaption = TryGet(() => (object?)pivot.DataFields.Item(caption));
            if (byCaption != null) return byCaption;
            object dataFields = ReadRequired(
                () => (object?)pivot.DataFields,
                "Excel did not expose the restored value fields.");
            IReadOnlyList<object> fields = ReadCollection(
                dataFields,
                MaximumFields,
                "DataFields");
            if (position > 0 && position <= fields.Count)
            {
                return fields[position - 1];
            }

            throw new InvalidOperationException(
                "Excel did not expose the restored value field '" + caption + "'.");
        }

        internal static void ApplyRegularFieldState(
            object targetFieldObject,
            LateBoundFieldState field,
            bool isDataModel)
        {
            dynamic target = targetFieldObject;
            if (!string.IsNullOrWhiteSpace(field.Caption))
            {
                TryWrite(() => target.Caption = field.Caption);
            }

            for (var index = 0; index < field.Subtotals.Count; index++)
            {
                int capturedIndex = index + 1;
                bool value = field.Subtotals[index];
                TryWrite(() => target.Subtotals[capturedIndex] = value);
            }

            TryWrite(() => target.RepeatLabels = field.RepeatLabels);
            if (field.Area == PivotNativeFieldArea.Row ||
                field.Area == PivotNativeFieldArea.Column)
            {
                target.LayoutBlankLine = field.LayoutBlankLine;
                target.LayoutPageBreak = field.LayoutPageBreak;
            }
            ApplyMemberVisibility(target, field.Members, isDataModel);
            if (field.Area == PivotNativeFieldArea.Filter)
            {
                target.EnableMultiplePageItems = field.MultiplePageItems;
                if (!field.MultiplePageItems && !string.IsNullOrWhiteSpace(field.CurrentPage))
                {
                    object pageItem = ResolvePageItem(target, field.CurrentPage);
                    if (isDataModel)
                    {
                        dynamic nativePageItem = pageItem;
                        string uniqueName = ReadOptionalString(
                            () => (object?)nativePageItem.SourceNameStandard);
                        if (string.IsNullOrWhiteSpace(uniqueName))
                        {
                            uniqueName = ReadOptionalString(
                                () => (object?)nativePageItem.SourceName);
                        }
                        if (string.IsNullOrWhiteSpace(uniqueName))
                        {
                            uniqueName = ReadRequiredString(
                                () => (object?)nativePageItem.Name,
                                "Excel did not expose the OLAP page item's unique name.");
                        }

                        target.CurrentPageName = uniqueName;
                    }
                    else
                    {
                        target.CurrentPage = pageItem;
                    }
                }
            }
        }

        private static object ResolvePageItem(dynamic targetField, string expectedPage)
        {
            object items = ReadRequired(
                () => (object?)targetField.PivotItems,
                "Excel did not expose page-field PivotItems.");
            List<object> matches = ReadCollection(
                    items,
                    MaximumMembers,
                    "page-field PivotItems")
                .Where(item =>
                {
                    dynamic nativeItem = item;
                    string name = ReadOptionalString(() => (object?)nativeItem.Name);
                    string source = ReadOptionalString(
                        () => (object?)nativeItem.SourceNameStandard);
                    if (string.IsNullOrWhiteSpace(source))
                    {
                        source = ReadOptionalString(() => (object?)nativeItem.SourceName);
                    }
                    string caption = ReadOptionalString(() => (object?)nativeItem.Caption);
                    return string.Equals(
                               FieldLeaf(name),
                               FieldLeaf(expectedPage),
                               StringComparison.OrdinalIgnoreCase) ||
                           string.Equals(
                               FieldLeaf(source),
                               FieldLeaf(expectedPage),
                               StringComparison.OrdinalIgnoreCase) ||
                           string.Equals(caption, expectedPage, StringComparison.OrdinalIgnoreCase);
                })
                .ToList();
            if (matches.Count != 1)
            {
                throw new InvalidOperationException(
                    "Excel did not expose one unambiguous PivotItem for the saved page selection.");
            }

            return matches[0];
        }

        private static void ApplyMemberVisibility(
            dynamic targetField,
            IReadOnlyList<LateBoundMemberState> expectedMembers,
            bool isDataModel)
        {
            if (expectedMembers.Count == 0)
            {
                return;
            }

            object? items = TryGet(() => (object?)targetField.PivotItems);
            if (items == null)
            {
                throw new InvalidOperationException(
                    "Excel did not expose PivotItems needed to restore a member filter.");
            }

            IReadOnlyList<object> actualItems = ReadCollection(
                items,
                MaximumMembers,
                "PivotItems");
            var bindings = new List<KeyValuePair<object, LateBoundMemberState>>();
            foreach (object itemObject in actualItems)
            {
                dynamic item = itemObject;
                string name = ReadOptionalString(() => (object?)item.Name);
                string sourceName = ReadOptionalString(
                    () => (object?)item.SourceNameStandard);
                if (string.IsNullOrWhiteSpace(sourceName))
                {
                    sourceName = ReadOptionalString(() => (object?)item.SourceName);
                }
                string caption = ReadOptionalString(() => (object?)item.Caption);
                List<LateBoundMemberState> matches = expectedMembers
                    .Where(member => MemberMatches(
                        member,
                        name,
                        sourceName,
                        caption))
                    .ToList();
                if (matches.Count > 1)
                {
                    throw new InvalidOperationException(
                        "The Data Model exposed an ambiguous PivotItem while restoring item order.");
                }

                if (matches.Count == 1)
                {
                    bindings.Add(new KeyValuePair<object, LateBoundMemberState>(
                        itemObject,
                        matches[0]));
                }
            }

            if (expectedMembers.Any(expected => !bindings.Any(binding =>
                    ReferenceEquals(binding.Value, expected))))
            {
                throw new InvalidOperationException(
                    "The Data Model did not expose every member required to restore the original item order and filter.");
            }

            if (isDataModel)
            {
                targetField.DatabaseSort = false;
                if (ReadRequiredBoolean(
                        () => (object?)targetField.DatabaseSort,
                        "Data Model PivotField DatabaseSort"))
                {
                    throw new InvalidOperationException(
                        "Excel did not enable manual ordering for the Data Model PivotField.");
                }

                foreach (KeyValuePair<object, LateBoundMemberState> binding in bindings
                             .OrderBy(item => item.Value.Position))
                {
                    dynamic item = binding.Key;
                    item.Position = binding.Value.Position;
                }

                List<string> visibleNames = bindings
                    .Where(binding => binding.Value.Visible)
                    .OrderBy(binding => binding.Value.Position)
                    .Select(binding => ReadOlapMemberUniqueName(binding.Key))
                    .ToList();
                if (visibleNames.Count == 0)
                {
                    throw new NotSupportedException(
                        "A Data Model field cannot restore a manual filter with no visible members.");
                }

                if (visibleNames.Count == expectedMembers.Count)
                {
                    object cubeFieldObject = ReadRequired(
                        () => (object?)targetField.CubeField,
                        "Excel did not expose the CubeField needed to clear an OLAP manual filter.");
                    dynamic cubeField = cubeFieldObject;
                    cubeField.ClearManualFilter();
                    if (!ReadRequiredBoolean(
                            () => (object?)cubeField.AllItemsVisible,
                            "CubeField AllItemsVisible state"))
                    {
                        throw new InvalidOperationException(
                            "Excel did not clear the Data Model manual member filter.");
                    }
                }
                else
                {
                    string[] safeArray = visibleNames.ToArray();
                    targetField.VisibleItemsList = safeArray;
                    IReadOnlyList<string> confirmed = ReadRequiredStringVector(
                        () => (object?)targetField.VisibleItemsList,
                        "Data Model PivotField VisibleItemsList");
                    if (!safeArray.SequenceEqual(
                            confirmed,
                            StringComparer.OrdinalIgnoreCase))
                    {
                        throw new InvalidOperationException(
                            "Excel did not retain the exact Data Model manual member filter.");
                    }
                }

                return;
            }

            foreach (KeyValuePair<object, LateBoundMemberState> binding in bindings
                         .OrderBy(item => item.Value.Position))
            {
                dynamic item = binding.Key;
                item.Position = binding.Value.Position;
            }

            // Classic PivotTables support PivotItem.Visible directly. Excel
            // rejects an intermediate state in which every member is hidden,
            // so restore visible members before hiding excluded ones.
            foreach (bool targetVisibility in new[] { true, false })
            {
                foreach (KeyValuePair<object, LateBoundMemberState> binding in bindings
                             .Where(item => item.Value.Visible == targetVisibility))
                {
                    dynamic item = binding.Key;
                    item.Visible = binding.Value.Visible;
                }
            }
        }

        private static string ReadOlapMemberUniqueName(object itemObject)
        {
            dynamic item = itemObject;
            string uniqueName = ReadOptionalString(
                () => (object?)item.SourceNameStandard);
            if (string.IsNullOrWhiteSpace(uniqueName))
            {
                uniqueName = ReadOptionalString(() => (object?)item.SourceName);
            }

            if (string.IsNullOrWhiteSpace(uniqueName))
            {
                throw new InvalidOperationException(
                    "Excel did not expose an OLAP unique member name for manual filtering.");
            }

            return uniqueName;
        }

        private static IReadOnlyList<string> ReadRequiredStringVector(
            Func<object?> reader,
            string label)
        {
            object value = ReadRequired(
                reader,
                "Excel did not expose the " + label + ".");
            if (value is string scalar)
            {
                return new[] { scalar };
            }

            if (!(value is Array array) || array.Rank != 1 ||
                array.Length > MaximumMembers)
            {
                throw new InvalidOperationException(
                    "Excel exposed an unsupported " + label + " contract.");
            }

            var result = new List<string>(array.Length);
            int lower = array.GetLowerBound(0);
            for (var index = 0; index < array.Length; index++)
            {
                string item = Convert.ToString(
                    array.GetValue(lower + index),
                    CultureInfo.InvariantCulture) ?? string.Empty;
                if (string.IsNullOrWhiteSpace(item))
                {
                    throw new InvalidOperationException(
                        "Excel exposed an empty item in " + label + ".");
                }

                result.Add(item);
            }

            return result;
        }

        private static bool MemberMatches(
            LateBoundMemberState expected,
            string name,
            string sourceName,
            string caption)
        {
            return string.Equals(
                       MemberKey(expected.Name, expected.Caption),
                       MemberKey(name, caption),
                       StringComparison.OrdinalIgnoreCase) ||
                   (!string.IsNullOrWhiteSpace(expected.Name) &&
                    string.Equals(
                        expected.Name,
                        sourceName,
                        StringComparison.OrdinalIgnoreCase)) ||
                   (!string.IsNullOrWhiteSpace(expected.Name) &&
                    string.Equals(
                        FieldLeaf(expected.Name),
                        FieldLeaf(sourceName),
                        StringComparison.OrdinalIgnoreCase)) ||
                   (!string.IsNullOrWhiteSpace(expected.Caption) &&
                    string.Equals(expected.Caption, caption, StringComparison.OrdinalIgnoreCase)) ||
                   (!string.IsNullOrWhiteSpace(expected.Name) &&
                    string.Equals(
                        FieldLeaf(expected.Name),
                        FieldLeaf(name),
                        StringComparison.OrdinalIgnoreCase));
        }

        private static string MemberKey(string name, string caption)
        {
            return FieldLeaf(name) + "\u001f" + caption;
        }

        internal static void ApplyStyleState(
            object pivotObject,
            LateBoundStyleState style,
            bool isDataModel = false)
        {
            dynamic pivot = pivotObject;
            pivot.RowGrand = style.RowGrand;
            pivot.ColumnGrand = style.ColumnGrand;
            pivot.DisplayFieldCaptions = style.DisplayFieldCaptions;
            pivot.PreserveFormatting = style.PreserveFormatting;
            pivot.ShowTableStyleRowStripes = style.ShowRowStripes;
            pivot.ShowTableStyleColumnStripes = style.ShowColumnStripes;
            pivot.DisplayNullString = style.DisplayNullString;
            pivot.NullString = style.NullString;
            pivot.DisplayErrorString = style.DisplayErrorString;
            pivot.ErrorString = style.ErrorString;
            pivot.ShowDrillIndicators = style.ShowDrillIndicators;
            pivot.VisualTotals = style.VisualTotals;
            if (isDataModel)
            {
                if (!ReadRequiredBoolean(
                        () => (object?)pivot.EnableDrilldown,
                        "Data Model EnableDrilldown invariant") ||
                    !ReadRequiredBoolean(
                        () => (object?)pivot.SubtotalHiddenPageItems,
                        "Data Model SubtotalHiddenPageItems invariant"))
                {
                    throw new InvalidOperationException(
                        "Excel did not expose the documented Data Model PivotTable invariants.");
                }
            }
            else
            {
                pivot.EnableDrilldown = style.EnableDrilldown;
                pivot.SubtotalHiddenPageItems = style.SubtotalHiddenPageItems;
            }
            pivot.PageFieldOrder = style.PageFieldOrder;
            pivot.PageFieldWrapCount = style.PageFieldWrapCount;
            pivot.CompactRowIndent = style.CompactRowIndent;
            pivot.MergeLabels = style.MergeLabels;
            // An empty TableStyle2 is itself meaningful. A newly-created
            // PivotTable can otherwise retain Excel's default style even when
            // the classic source intentionally had no table style.
            pivot.TableStyle2 = style.TableStyleName;

            pivot.RowAxisLayout(style.RowAxisLayout);
        }

        private static int OrientationFor(PivotNativeFieldArea area)
        {
            switch (area)
            {
                case PivotNativeFieldArea.Row: return OrientationRow;
                case PivotNativeFieldArea.Column: return OrientationColumn;
                case PivotNativeFieldArea.Filter: return OrientationPage;
                case PivotNativeFieldArea.Values: return OrientationData;
                default: throw new ArgumentOutOfRangeException(nameof(area));
            }
        }

        private static string FieldLeaf(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            string normalized = value.Replace("]]", "]");
            string[] parts = normalized.Split(new[] { "].[" }, StringSplitOptions.None);
            string leaf = parts[parts.Length - 1].Trim('[', ']');
            return leaf;
        }

        private static LateBoundPivotState ReadNativeState(PivotNativeStateSnapshot snapshot)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            if (snapshot.NativeState is LateBoundPivotState state)
            {
                return state;
            }

            throw new ArgumentException(
                "The snapshot was not created by this Excel gateway.",
                nameof(snapshot));
        }

        private static object ReadPivotCache(object pivotTable)
        {
            dynamic pivot = pivotTable;
            object? cache = TryGet(() => (object?)pivot.PivotCache());
            cache = cache ?? TryGet(() => (object?)pivot.PivotCache);
            return cache ?? throw new InvalidOperationException(
                "Excel did not expose the PivotTable cache.");
        }

        private static string ReadAddress(object cellObject)
        {
            dynamic cell = cellObject;
            object? value = TryGet(() => (object?)cell.Address(false, false));
            value = value ?? TryGet(() => (object?)cell.Address);
            string address = Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
            if (string.IsNullOrWhiteSpace(address) || address.Length > 64)
            {
                throw new InvalidOperationException(
                    "Excel did not expose a bounded PivotTable anchor address.");
            }

            return address;
        }

        internal static PivotResultSignature ReadResultSignature(object rangeObject)
        {
            if (rangeObject == null) throw new ArgumentNullException(nameof(rangeObject));
            dynamic range = rangeObject;
            int rows = ReadRequiredCollectionCount(
                ReadRequired(
                    () => (object?)range.Rows,
                    "Excel did not expose the PivotTable result rows."),
                MaximumPivotResultCells,
                "PivotTable result rows");
            int columns = ReadRequiredCollectionCount(
                ReadRequired(
                    () => (object?)range.Columns,
                    "Excel did not expose the PivotTable result columns."),
                MaximumPivotResultCells,
                "PivotTable result columns");
            long cells = (long)rows * columns;
            if (rows <= 0 || columns <= 0 || cells > MaximumPivotResultCells)
            {
                throw new NotSupportedException(
                    "The PivotTable result exceeds the bounded reversible-conversion limit.");
            }

            object? values;
            if (!PivotLateBound.TryRead(() => (object?)range.Value2, out values))
            {
                throw new NotSupportedException(
                    "Excel did not expose the PivotTable result values for reversible verification.");
            }

            var canonical = new StringBuilder(
                Math.Min(MaximumPivotResultCharacters, checked((int)cells * 16)));
            canonical.Append(rows.ToString(CultureInfo.InvariantCulture))
                .Append('x')
                .Append(columns.ToString(CultureInfo.InvariantCulture))
                .Append(';');
            if (values is Array array)
            {
                if (array.Rank != 2 ||
                    array.GetLength(0) != rows ||
                    array.GetLength(1) != columns)
                {
                    throw new NotSupportedException(
                        "Excel exposed an inconsistent PivotTable result value matrix.");
                }

                int rowLower = array.GetLowerBound(0);
                int columnLower = array.GetLowerBound(1);
                for (var row = 0; row < rows; row++)
                {
                    for (var column = 0; column < columns; column++)
                    {
                        AppendCanonicalCell(
                            canonical,
                            array.GetValue(rowLower + row, columnLower + column));
                    }
                }
            }
            else
            {
                if (rows != 1 || columns != 1)
                {
                    throw new NotSupportedException(
                        "Excel exposed a scalar value for a multi-cell PivotTable result.");
                }

                AppendCanonicalCell(canonical, values);
            }

            return new PivotResultSignature(
                rows,
                columns,
                PivotPlusFingerprint.Create(
                    "pivotplus.result.v1",
                    canonical.ToString()));
        }

        private static void AppendCanonicalCell(StringBuilder output, object? value)
        {
            string kind;
            string text;
            if (value == null)
            {
                kind = "null";
                text = string.Empty;
            }
            else if (value is string stringValue)
            {
                kind = "string";
                text = stringValue;
            }
            else if (value is bool booleanValue)
            {
                kind = "bool";
                text = booleanValue ? "1" : "0";
            }
            else if (value is ErrorWrapper error)
            {
                kind = "error";
                text = error.ErrorCode.ToString(CultureInfo.InvariantCulture);
            }
            else if (value is DateTime dateTime)
            {
                kind = "date";
                text = dateTime.ToString("O", CultureInfo.InvariantCulture);
            }
            else if (value is IFormattable formattable &&
                     (value is byte || value is sbyte ||
                      value is short || value is ushort ||
                      value is int || value is uint ||
                      value is long || value is ulong ||
                      value is float || value is double || value is decimal))
            {
                kind = value.GetType().Name;
                text = formattable.ToString(null, CultureInfo.InvariantCulture) ?? string.Empty;
            }
            else
            {
                throw new NotSupportedException(
                    "Excel exposed an unsupported PivotTable result value type.");
            }

            if (text.Length > MaximumPivotCellCharacters ||
                output.Length + kind.Length + text.Length + 32 > MaximumPivotResultCharacters)
            {
                throw new NotSupportedException(
                    "The PivotTable result exceeds the bounded verification payload limit.");
            }

            output.Append(kind.Length.ToString(CultureInfo.InvariantCulture))
                .Append(':')
                .Append(kind)
                .Append(':')
                .Append(text.Length.ToString(CultureInfo.InvariantCulture))
                .Append(':')
                .Append(text)
                .Append(';');
        }

        private static int ReadRequiredCollectionCount(
            object collectionObject,
            int maximum,
            string label)
        {
            dynamic collection = collectionObject;
            object value = ReadRequired(
                () => (object?)collection.Count,
                "Excel did not expose the " + label + " count.");
            try
            {
                int count = Convert.ToInt32(value, CultureInfo.InvariantCulture);
                if (count < 0 || count > maximum)
                {
                    throw new NotSupportedException(
                        "The Excel " + label + " collection exceeds its bounded conversion limit.");
                }

                return count;
            }
            catch (NotSupportedException)
            {
                throw;
            }
            catch (Exception exception) when (
                exception is FormatException ||
                exception is InvalidCastException ||
                exception is OverflowException)
            {
                throw new NotSupportedException(
                    "Excel exposed an invalid " + label + " count.",
                    exception);
            }
        }

        private static IReadOnlyList<object> ReadCollection(
            object collectionObject,
            int maximum,
            string label)
        {
            int count = ReadCollectionCount(collectionObject, maximum, label);
            dynamic collection = collectionObject;
            var result = new List<object>(count);
            for (var index = 1; index <= count; index++)
            {
                int capturedIndex = index;
                object? item = TryGet(() => (object?)collection.Item(capturedIndex));
                item = item ?? TryGet(() => (object?)collection[capturedIndex]);
                if (item == null)
                {
                    throw new InvalidOperationException(
                        "Excel did not expose an item in the " + label + " collection.");
                }

                result.Add(item);
            }

            return result;
        }

        private static int ReadCollectionCount(object collectionObject, int maximum, string label)
        {
            dynamic collection = collectionObject;
            int count = ReadOptionalInt(() => (object?)collection.Count, -1);
            if (count < 0 || count > maximum)
            {
                throw new NotSupportedException(
                    "The Excel " + label + " collection exceeds its bounded conversion limit.");
            }

            return count;
        }

        private static object ReadRequired(Func<object?> reader, string message)
        {
            return TryGet(reader) ?? throw new InvalidOperationException(message);
        }

        private static string ReadRequiredString(Func<object?> reader, string message)
        {
            string value = ReadOptionalString(reader);
            return !string.IsNullOrWhiteSpace(value)
                ? value
                : throw new InvalidOperationException(message);
        }

        private static string ReadOptionalString(Func<object?> reader)
        {
            object? value = TryGet(reader);
            return value == null
                ? string.Empty
                : Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
        }

        private static string ReadRequiredOptionalString(
            Func<object?> reader,
            string label)
        {
            if (!PivotLateBound.TryRead(reader, out object? value))
            {
                throw new NotSupportedException(
                    "Excel did not expose the " + label + " for a reversible conversion.");
            }

            return value == null
                ? string.Empty
                : Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
        }

        private static int ReadRequiredInt(Func<object?> reader, string label)
        {
            object value = ReadRequired(
                reader,
                "Excel did not expose the " + label + " for a reversible conversion.");
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
            object value = ReadRequired(
                reader,
                "Excel did not expose the " + label + " for a reversible conversion.");
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

        private static int ReadOptionalInt(Func<object?> reader, int fallback)
        {
            object? value = TryGet(reader);
            if (value == null) return fallback;
            try
            {
                return Convert.ToInt32(value, CultureInfo.InvariantCulture);
            }
            catch (Exception exception) when (
                exception is FormatException ||
                exception is InvalidCastException ||
                exception is OverflowException)
            {
                return fallback;
            }
        }

        private static int? ReadOptionalNullableInt(Func<object?> reader)
        {
            object? value = TryGet(reader);
            if (value == null) return null;
            try
            {
                return Convert.ToInt32(value, CultureInfo.InvariantCulture);
            }
            catch (Exception exception) when (
                exception is FormatException ||
                exception is InvalidCastException ||
                exception is OverflowException)
            {
                return null;
            }
        }

        private static long ReadOptionalLong(Func<object?> reader, long fallback)
        {
            object? value = TryGet(reader);
            if (value == null) return fallback;
            try
            {
                return Convert.ToInt64(value, CultureInfo.InvariantCulture);
            }
            catch (Exception exception) when (
                exception is FormatException ||
                exception is InvalidCastException ||
                exception is OverflowException)
            {
                return fallback;
            }
        }

        private static bool ReadBoolean(Func<object?> reader, string label)
        {
            object value = ReadRequired(
                reader,
                "Excel did not expose " + label + ".");
            try
            {
                return Convert.ToBoolean(value, CultureInfo.InvariantCulture);
            }
            catch (Exception exception) when (
                exception is FormatException || exception is InvalidCastException)
            {
                throw new InvalidOperationException(
                    "Excel exposed an invalid " + label + " value.",
                    exception);
            }
        }

        private static bool ReadOptionalBoolean(Func<object?> reader, bool fallback)
        {
            object? value = TryGet(reader);
            if (value == null) return fallback;
            try
            {
                return Convert.ToBoolean(value, CultureInfo.InvariantCulture);
            }
            catch (Exception exception) when (
                exception is FormatException || exception is InvalidCastException)
            {
                return fallback;
            }
        }

        private static object? TryGet(Func<object?> reader)
        {
            return PivotLateBound.TryRead(reader, out object? value) ? value : null;
        }

        private static bool SameNativeObject(object left, object right)
        {
            if (ReferenceEquals(left, right)) return true;
            if (!Marshal.IsComObject(left) || !Marshal.IsComObject(right)) return false;

            IntPtr leftIdentity = IntPtr.Zero;
            IntPtr rightIdentity = IntPtr.Zero;
            try
            {
                leftIdentity = Marshal.GetIUnknownForObject(left);
                rightIdentity = Marshal.GetIUnknownForObject(right);
                return leftIdentity == rightIdentity;
            }
            finally
            {
                if (leftIdentity != IntPtr.Zero) Marshal.Release(leftIdentity);
                if (rightIdentity != IntPtr.Zero) Marshal.Release(rightIdentity);
            }
        }

        private static void TryWrite(Action writer)
        {
            try
            {
                writer();
            }
            catch (Exception)
            {
                // Optional host-version property. Required mutations are never
                // routed through this helper.
            }
        }

        private static void DemandCollectionNameAvailable(
            object collection,
            string objectKind,
            string objectName,
            bool workbookScopedName = false)
        {
            if (FindNamedObject(
                    collection,
                    objectName,
                    objectKind + "s",
                    workbookScopedName) != null)
            {
                throw new InvalidOperationException(
                    "A workbook " + objectKind + " named '" + objectName +
                    "' already exists. PivotTable+ will not reuse or overwrite it.");
            }
        }

        private static object? FindNamedObject(
            object collection,
            string expectedName,
            string label,
            bool workbookScopedName = false)
        {
            List<object> matches = ReadCollection(
                    collection,
                    MaximumWorkbookObjects,
                    label)
                .Where(item =>
                {
                    dynamic nativeItem = item;
                    string actualName = ReadRequiredString(
                        () => (object?)nativeItem.Name,
                        "Excel exposed an unnamed object in the " + label + " collection.");
                    if (workbookScopedName)
                    {
                        actualName = WorkbookScopedName(actualName);
                    }

                    return string.Equals(
                        actualName,
                        expectedName,
                        StringComparison.OrdinalIgnoreCase);
                })
                .ToList();
            if (matches.Count > 1)
            {
                throw new InvalidOperationException(
                    "Excel exposed duplicate objects named '" + expectedName +
                    "' in the " + label + " collection.");
            }

            return matches.SingleOrDefault();
        }

        private static void DeleteValidatedArtifact(
            object? artifact,
            string label,
            ICollection<Exception> failures)
        {
            if (artifact == null) return;
            try
            {
                ((dynamic)artifact).Delete();
            }
            catch (Exception exception)
            {
                failures.Add(new InvalidOperationException(
                    "PivotTable+ could not delete the validated generated " + label + ".",
                    exception));
            }
        }

        private static void CaptureCleanupFailure(
            object? value,
            string label,
            ICollection<Exception> failures)
        {
            if (value == null) return;
            try
            {
                ((dynamic)value).Delete();
            }
            catch (Exception exception)
            {
                failures.Add(new InvalidOperationException(
                    "PivotTable+ could not delete the partially created " + label + ".",
                    exception));
            }
        }

        private static void DeleteWorksheet(dynamic worksheet)
        {
            object? application = TryGet(() => (object?)worksheet.Application);
            if (application == null)
            {
                worksheet.Delete();
                return;
            }

            dynamic app = application;
            bool alerts = ReadOptionalBoolean(() => (object?)app.DisplayAlerts, true);
            try
            {
                app.DisplayAlerts = false;
                worksheet.Delete();
            }
            finally
            {
                app.DisplayAlerts = alerts;
            }
        }

        /// <summary>
        /// Bounded, format-only backup of the original PivotTable result.
        /// The temporary worksheet never receives values or formulas: Excel's
        /// clipboard is pasted with xlPasteFormats/xlPasteColumnWidths only.
        /// Row heights and column widths are also captured explicitly because
        /// xlPasteFormats does not preserve those dimensions reliably.
        /// </summary>
        private sealed class PivotFormatBackup
        {
            private readonly object workbook;
            private readonly object worksheet;
            private readonly object range;
            private readonly string worksheetName;
            private readonly PivotTemporaryWorksheetArtifact receipt;
            private readonly double[] rowHeights;
            private readonly double[] columnWidths;
            private bool deleted;

            private PivotFormatBackup(
                object workbook,
                object worksheet,
                object range,
                string worksheetName,
                PivotTemporaryWorksheetArtifact receipt,
                double[] rowHeights,
                double[] columnWidths)
            {
                this.workbook = workbook;
                this.worksheet = worksheet;
                this.range = range;
                this.worksheetName = worksheetName;
                this.receipt = receipt;
                this.rowHeights = rowHeights;
                this.columnWidths = columnWidths;
            }

            public static PivotFormatBackup Create(
                object workbook,
                object pivotTable,
                PivotResultSignature result,
                PivotTemporaryWorksheetArtifact receipt)
            {
                dynamic pivot = pivotTable;
                object sourceRange = ReadRequired(
                    () => (object?)pivot.TableRange2,
                    "Excel did not expose the original PivotTable range for its format backup.");
                DemandExactExtent(sourceRange, result);
                double[] rowHeights = ReadDimensionValues(
                    sourceRange,
                    rows: true,
                    result.Rows);
                double[] columnWidths = ReadDimensionValues(
                    sourceRange,
                    rows: false,
                    result.Columns);

                if (!string.Equals(
                        receipt.Purpose,
                        "format-backup",
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "The durable temporary worksheet receipt is not a format-backup receipt.");
                }

                string backupName = receipt.Name;
                dynamic book = workbook;
                object worksheets = ReadRequired(
                    () => (object?)book.Worksheets,
                    "Excel did not expose workbook worksheets for the PivotTable format backup.");
                ReconcileStaleTemporaryWorksheet(
                    workbook,
                    receipt,
                    expectedPivotName: string.Empty,
                    isFormatBackup: true,
                    expectedModelConnection: null);
                DemandCollectionNameAvailable(
                    worksheets,
                    "worksheet",
                    backupName);

                object? backupWorksheet = null;
                try
                {
                    backupWorksheet = book.Worksheets.Add();
                    dynamic nativeWorksheet = backupWorksheet;
                    nativeWorksheet.Name = backupName;
                    nativeWorksheet.Visible = SheetVeryHidden;
                    WriteTemporaryWorksheetMarker(nativeWorksheet, receipt);
                    object anchor = ReadRequired(
                        () => (object?)nativeWorksheet.Range["A1"],
                        "Excel did not expose the format-backup anchor cell.");
                    dynamic nativeAnchor = anchor;
                    object backupRange = ReadRequired(
                        () => (object?)nativeAnchor.Resize[result.Rows, result.Columns],
                        "Excel did not expose the bounded format-backup range.");
                    DemandExactExtent(backupRange, result);

                    CopyFormats(sourceRange, backupRange, workbook);
                    WriteDimensionValues(backupRange, rows: true, rowHeights);
                    WriteDimensionValues(backupRange, rows: false, columnWidths);
                    return new PivotFormatBackup(
                        workbook,
                        backupWorksheet,
                        backupRange,
                        backupName,
                        receipt,
                        rowHeights,
                        columnWidths);
                }
                catch (Exception failure)
                {
                    ClearCopyMode(workbook);
                    if (backupWorksheet == null) throw;
                    try
                    {
                        DeleteWorksheet((dynamic)backupWorksheet);
                    }
                    catch (Exception cleanupFailure)
                    {
                        throw new AggregateException(
                            "PivotTable+ could not create or clean up its format-only backup.",
                            failure,
                            cleanupFailure);
                    }

                    throw;
                }
            }

            public static PivotFormatBackup OpenExisting(
                object workbook,
                object worksheet,
                PivotResultSignature result,
                PivotTemporaryWorksheetArtifact receipt)
            {
                if (workbook == null) throw new ArgumentNullException(nameof(workbook));
                if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));
                if (result == null) throw new ArgumentNullException(nameof(result));
                if (receipt == null) throw new ArgumentNullException(nameof(receipt));
                if (!string.Equals(
                        receipt.Purpose,
                        "format-backup",
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "The pending worksheet receipt is not a format-backup receipt.");
                }

                dynamic nativeWorksheet = worksheet;
                string worksheetName = ReadRequiredString(
                    () => (object?)nativeWorksheet.Name,
                    "Excel did not expose the pending format-backup worksheet name.");
                if (!string.Equals(
                        worksheetName,
                        receipt.Name,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "The pending format-backup worksheet was renamed.");
                }

                DemandTemporaryWorksheetMarker(nativeWorksheet, receipt);
                DemandTemporaryWorksheetStructure(
                    nativeWorksheet,
                    expectedPivotName: string.Empty,
                    isFormatBackup: true,
                    allowIncomplete: false,
                    expectedModelConnection: null);
                object anchor = ReadRequired(
                    () => (object?)nativeWorksheet.Range["A1"],
                    "Excel did not expose the pending format-backup anchor.");
                dynamic nativeAnchor = anchor;
                object backupRange = ReadRequired(
                    () => (object?)nativeAnchor.Resize[result.Rows, result.Columns],
                    "Excel did not expose the pending bounded format-backup range.");
                DemandExactExtent(backupRange, result);
                object usedRange = ReadRequired(
                    () => (object?)nativeWorksheet.UsedRange,
                    "Excel did not expose the pending format-backup UsedRange.");
                if (!string.Equals(
                        ReadAddress(usedRange),
                        ReadAddress(backupRange),
                        StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        "The pending format-backup extent does not match the verified staged result.");
                }

                double[] rowHeights = ReadDimensionValues(
                    backupRange,
                    rows: true,
                    result.Rows);
                double[] columnWidths = ReadDimensionValues(
                    backupRange,
                    rows: false,
                    result.Columns);
                return new PivotFormatBackup(
                    workbook,
                    worksheet,
                    backupRange,
                    worksheetName,
                    receipt,
                    rowHeights,
                    columnWidths);
            }

            public void Restore(object pivotTable, PivotResultSignature expectedResult)
            {
                if (deleted)
                {
                    throw new InvalidOperationException(
                        "The PivotTable format backup was already removed.");
                }

                dynamic pivot = pivotTable;
                object targetRange = ReadRequired(
                    () => (object?)pivot.TableRange2,
                    "Excel did not expose the replacement PivotTable range for format restoration.");
                DemandExactExtent(targetRange, expectedResult);
                CopyFormats(range, targetRange, workbook);
                WriteDimensionValues(targetRange, rows: true, rowHeights);
                WriteDimensionValues(targetRange, rows: false, columnWidths);
            }

            public void Delete()
            {
                if (deleted) return;
                dynamic nativeWorksheet = worksheet;
                string currentName = ReadRequiredString(
                    () => (object?)nativeWorksheet.Name,
                    "Excel did not expose the format-backup worksheet name.");
                if (!string.Equals(
                        currentName,
                        worksheetName,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "The generated format-backup worksheet was renamed; PivotTable+ will not delete it.");
                }

                DemandTemporaryWorksheetMarker(nativeWorksheet, receipt);
                DemandTemporaryWorksheetStructure(
                    nativeWorksheet,
                    expectedPivotName: string.Empty,
                    isFormatBackup: true,
                    allowIncomplete: false,
                    expectedModelConnection: null);
                DeleteOwnedTemporaryWorksheet(workbook, nativeWorksheet, receipt);
                deleted = true;
            }

            private static void CopyFormats(
                object sourceRange,
                object targetRange,
                object workbook)
            {
                try
                {
                    dynamic source = sourceRange;
                    dynamic target = targetRange;
                    source.Copy();
                    target.PasteSpecial(PasteFormats);
                    target.PasteSpecial(PasteColumnWidths);
                }
                finally
                {
                    ClearCopyMode(workbook);
                }
            }

            private static void ClearCopyMode(object workbook)
            {
                dynamic book = workbook;
                TryWrite(() => book.Application.CutCopyMode = false);
            }

            private static void DemandExactExtent(
                object rangeObject,
                PivotResultSignature expected)
            {
                dynamic range = rangeObject;
                object rows = ReadRequired(
                    () => (object?)range.Rows,
                    "Excel did not expose the PivotTable format range rows.");
                object columns = ReadRequired(
                    () => (object?)range.Columns,
                    "Excel did not expose the PivotTable format range columns.");
                int rowCount = ReadRequiredCollectionCount(
                    rows,
                    MaximumPivotResultCells,
                    "PivotTable format rows");
                int columnCount = ReadRequiredCollectionCount(
                    columns,
                    MaximumPivotResultCells,
                    "PivotTable format columns");
                if (rowCount != expected.Rows || columnCount != expected.Columns)
                {
                    throw new InvalidOperationException(
                        "The PivotTable result extent changed before its formats could be restored safely.");
                }
            }

            private static double[] ReadDimensionValues(
                object rangeObject,
                bool rows,
                int expectedCount)
            {
                dynamic range = rangeObject;
                object collection = ReadRequired(
                    () => rows ? (object?)range.Rows : (object?)range.Columns,
                    "Excel did not expose the PivotTable " +
                    (rows ? "rows" : "columns") + " for format preservation.");
                int count = ReadRequiredCollectionCount(
                    collection,
                    MaximumPivotResultCells,
                    rows ? "PivotTable rows" : "PivotTable columns");
                if (count != expectedCount)
                {
                    throw new InvalidOperationException(
                        "The PivotTable extent changed while its formatting was being captured.");
                }

                var values = new double[count];
                dynamic nativeCollection = collection;
                for (var index = 1; index <= count; index++)
                {
                    int capturedIndex = index;
                    object dimension = ReadRequired(
                        () => (object?)nativeCollection.Item(capturedIndex),
                        "Excel did not expose a PivotTable " +
                        (rows ? "row" : "column") + " for format preservation.");
                    dynamic nativeDimension = dimension;
                    object value = ReadRequired(
                        () => rows
                            ? (object?)nativeDimension.RowHeight
                            : (object?)nativeDimension.ColumnWidth,
                        "Excel did not expose a PivotTable " +
                        (rows ? "row height" : "column width") + ".");
                    double numeric;
                    try
                    {
                        numeric = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                    }
                    catch (Exception exception) when (
                        exception is FormatException ||
                        exception is InvalidCastException ||
                        exception is OverflowException)
                    {
                        throw new NotSupportedException(
                            "Excel exposed an invalid PivotTable " +
                            (rows ? "row height" : "column width") + ".",
                            exception);
                    }

                    if (double.IsNaN(numeric) || double.IsInfinity(numeric) || numeric <= 0d)
                    {
                        throw new NotSupportedException(
                            "Excel exposed an unsupported PivotTable " +
                            (rows ? "row height" : "column width") + ".");
                    }

                    values[index - 1] = numeric;
                }

                return values;
            }

            private static void WriteDimensionValues(
                object rangeObject,
                bool rows,
                IReadOnlyList<double> values)
            {
                dynamic range = rangeObject;
                object collection = ReadRequired(
                    () => rows ? (object?)range.Rows : (object?)range.Columns,
                    "Excel did not expose the target " +
                    (rows ? "rows" : "columns") + " for format restoration.");
                int count = ReadRequiredCollectionCount(
                    collection,
                    MaximumPivotResultCells,
                    rows ? "target rows" : "target columns");
                if (count != values.Count)
                {
                    throw new InvalidOperationException(
                        "The target PivotTable extent changed while its formatting was being restored.");
                }

                dynamic nativeCollection = collection;
                for (var index = 1; index <= count; index++)
                {
                    int capturedIndex = index;
                    object dimension = ReadRequired(
                        () => (object?)nativeCollection.Item(capturedIndex),
                        "Excel did not expose a target PivotTable " +
                        (rows ? "row" : "column") + " for format restoration.");
                    dynamic nativeDimension = dimension;
                    if (rows)
                    {
                        nativeDimension.RowHeight = values[index - 1];
                    }
                    else
                    {
                        nativeDimension.ColumnWidth = values[index - 1];
                    }
                }
            }
        }

        private sealed class LateBoundPivotReplacementTransaction : IPivotReplacementTransaction
        {
            private readonly LateBoundPivotDataModelEnablementGateway owner;
            private readonly object workbook;
            private readonly object originalPivotTable;
            private readonly object replacementCache;
            private readonly PivotNativeStateSnapshot originalSnapshot;
            private readonly LateBoundPivotState originalState;
            private readonly string modelTableName;
            private readonly PivotFormatBackup formatBackup;
            private readonly PivotTemporaryPivotTableArtifact temporaryPivotReceipt;
            private readonly string stagingStateFingerprint;
            private bool committed;
            private bool forwardTargetVerified;

            public LateBoundPivotReplacementTransaction(
                LateBoundPivotDataModelEnablementGateway owner,
                object workbook,
                object originalPivotTable,
                object replacementCache,
                PivotNativeStateSnapshot originalSnapshot,
                string modelTableName,
                PivotFormatBackup formatBackup,
                PivotTemporaryPivotTableArtifact temporaryPivotReceipt,
                string stagingStateFingerprint)
            {
                this.owner = owner;
                this.workbook = workbook;
                this.originalPivotTable = originalPivotTable;
                this.replacementCache = replacementCache;
                this.originalSnapshot = originalSnapshot;
                originalState = ReadNativeState(originalSnapshot);
                this.modelTableName = modelTableName;
                this.formatBackup = formatBackup;
                this.temporaryPivotReceipt = temporaryPivotReceipt;
                this.stagingStateFingerprint = stagingStateFingerprint;
            }

            public bool ReplacementAttempted { get; private set; }

            public bool IsCommitted => committed || forwardTargetVerified;

            public object? ReplacementPivotTable { get; private set; }

            public void ReplaceAtOriginalLocation()
            {
                if (ReplacementAttempted)
                {
                    throw new InvalidOperationException("The PivotTable replacement was already attempted.");
                }

                dynamic book = workbook;
                dynamic worksheet = book.Worksheets.Item(originalSnapshot.WorksheetName);
                dynamic destination = worksheet.Range[originalSnapshot.AnchorAddress];
                DemandTemporaryPivotTableReceipt(
                    temporaryPivotReceipt,
                    originalSnapshot.WorksheetName,
                    originalSnapshot.PivotTableName,
                    originalSnapshot.AnchorAddress,
                    modelTableName);
                ReplacementAttempted = true;
                dynamic original = originalPivotTable;
                original.TableRange2.Clear();
                dynamic cache = replacementCache;
                ReplacementPivotTable = cache.CreatePivotTable(
                    destination,
                    temporaryPivotReceipt.Name);
                DemandExactTemporaryTargetPivot(
                    ReplacementPivotTable,
                    worksheet,
                    temporaryPivotReceipt,
                    ReadRequired(
                        () => (object?)cache.WorkbookConnection,
                        "Excel did not expose the replacement cache Data Model connection."));
                RestoreDataModelState(
                    ReplacementPivotTable,
                    originalState,
                    modelTableName);
                owner.RefreshPivotTable(ReplacementPivotTable);
                formatBackup.Restore(ReplacementPivotTable, originalState.Result);
                DemandStateFingerprint(
                    ReplacementPivotTable,
                    stagingStateFingerprint);
                dynamic replacement = ReplacementPivotTable;
                replacement.Name = originalSnapshot.PivotTableName;
                string promotedName = ReadRequiredString(
                    () => (object?)replacement.Name,
                    "Excel did not expose the promoted replacement PivotTable name.");
                if (!string.Equals(
                        promotedName,
                        originalSnapshot.PivotTableName,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "Excel did not promote the verified temporary PivotTable to the original name.");
                }

                DemandStateFingerprint(
                    ReplacementPivotTable,
                    stagingStateFingerprint);
                forwardTargetVerified = true;
            }

            public void VerifyReplacement()
            {
                if (ReplacementPivotTable == null)
                {
                    throw new InvalidOperationException("The replacement PivotTable was not created.");
                }

                string fingerprint = owner.VerifyDataModelState(
                    ReplacementPivotTable,
                    originalSnapshot);
                if (!string.Equals(
                        fingerprint,
                        stagingStateFingerprint,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "The promoted Data Model PivotTable no longer matches the durable staging checkpoint.");
                }
            }

            public void RollBack()
            {
                if (!ReplacementAttempted) return;
                dynamic book = workbook;
                dynamic worksheet = book.Worksheets.Item(originalSnapshot.WorksheetName);
                dynamic destination = worksheet.Range[originalSnapshot.AnchorAddress];
                object? originalNamed = TryGet(
                    () => (object?)worksheet.PivotTables.Item(originalSnapshot.PivotTableName));
                object? temporaryNamed = TryGet(
                    () => (object?)worksheet.PivotTables.Item(temporaryPivotReceipt.Name));
                if (originalNamed != null && temporaryNamed != null &&
                    !SameNativeObject(originalNamed, temporaryNamed))
                {
                    throw new InvalidOperationException(
                        "Both the original and generated temporary PivotTable names exist; rollback will not clear either target.");
                }

                object? incumbent = temporaryNamed ?? originalNamed;
                object? restored = null;
                if (incumbent != null)
                {
                    dynamic incumbentCache = ReadPivotCache(incumbent);
                    bool incumbentIsModel = ReadBoolean(
                        () => (object?)incumbentCache.OLAP,
                        "PivotCache.OLAP");
                    if (incumbentIsModel)
                    {
                        if (originalNamed != null && temporaryNamed == null)
                        {
                            DemandStateFingerprint(
                                incumbent,
                                stagingStateFingerprint);
                            // Rename may have committed before Excel surfaced a
                            // COM failure. Preserve the exact verified forward
                            // target; pending recovery will finish cleanup.
                            ReplacementPivotTable = incumbent;
                            forwardTargetVerified = true;
                            return;
                        }

                        dynamic nativeReplacementCache = replacementCache;
                        object modelConnection = ReadRequired(
                            () => (object?)nativeReplacementCache.WorkbookConnection,
                            "Excel did not expose the rollback Data Model connection.");
                        DemandExactTemporaryTargetPivot(
                            incumbent,
                            worksheet,
                            temporaryPivotReceipt,
                            modelConnection);
                        ((dynamic)incumbent).TableRange2.Clear();
                    }
                    else
                    {
                        // Clear can throw before removing the classic table.
                        // In that case restore the survivor in place instead of
                        // trying to create another PivotTable over it.
                        restored = incumbent;
                    }
                }

                ReplacementPivotTable = null;
                if (restored == null)
                {
                    dynamic originalCache = originalState.OriginalCache;
                    restored = originalCache.CreatePivotTable(
                        destination,
                        originalSnapshot.PivotTableName);
                }

                Exception? failure = null;
                try
                {
                    RestoreClassicState(restored, originalState);
                    owner.RefreshPivotTable(restored);
                    formatBackup.Restore(restored, originalState.Result);
                LateBoundPivotState verified = new LateBoundPivotState(
                    ReadPivotCache(restored),
                    ReadFieldStates((dynamic)restored),
                    ReadStyleState((dynamic)restored),
                    ReadResultSignature(ReadRequired(
                        () => (object?)((dynamic)restored).TableRange2,
                        "Excel did not expose the restored PivotTable result range.")),
                    ReadDataAxisState(restored));
                    if (!LateBoundPivotState.SemanticallyEquals(originalState, verified))
                    {
                        throw new InvalidOperationException(
                            "Excel did not restore the original classic PivotTable state.");
                    }
                }
                catch (Exception exception)
                {
                    failure = exception;
                }

                if (failure != null)
                {
                    // Retain the exact marked format-only backup when classic
                    // restoration is incomplete. It is the only durable copy
                    // of the original cell formatting/dimensions and remains
                    // covered by Pending ownership for explicit recovery.
                    throw failure;
                }

                formatBackup.Delete();
            }

            public void Commit()
            {
                if (ReplacementPivotTable == null)
                {
                    throw new InvalidOperationException("There is no replacement PivotTable to commit.");
                }

                formatBackup.Delete();
                committed = true;
            }

            public void Dispose()
            {
                // The orchestrator owns rollback. Dispose intentionally does
                // not perform a second destructive mutation after an error.
                _ = committed;
            }
        }
    }

    internal enum PivotNativeFieldArea
    {
        Row,
        Column,
        Filter,
        Values
    }

    internal sealed class LateBoundDataAxisState
    {
        public static readonly LateBoundDataAxisState Hidden =
            new LateBoundDataAxisState();

        private LateBoundDataAxisState()
        {
            Axis = PivotValuesAxis.Automatic;
        }

        public LateBoundDataAxisState(PivotValuesAxis axis, int position)
        {
            if (axis != PivotValuesAxis.Rows &&
                axis != PivotValuesAxis.Columns)
            {
                throw new ArgumentOutOfRangeException(nameof(axis));
            }

            if (position <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(position));
            }

            Axis = axis;
            Position = position;
            IsVisible = true;
        }

        public bool IsVisible { get; }

        public PivotValuesAxis Axis { get; }

        public int Position { get; }

        public string CanonicalValue()
        {
            return IsVisible
                ? Axis + ":" + Position.ToString(CultureInfo.InvariantCulture)
                : "hidden";
        }
    }

    internal sealed class LateBoundMemberState
    {
        public LateBoundMemberState(string name, string caption, bool visible, int position)
        {
            Name = name ?? string.Empty;
            Caption = caption ?? string.Empty;
            Visible = visible;
            Position = position;
        }

        public string Name { get; }

        public string Caption { get; }

        public bool Visible { get; }

        public int Position { get; }

        public string CanonicalValue()
        {
            return Position.ToString(CultureInfo.InvariantCulture) + ":" +
                   Token(Name) + ":" + Token(Caption) + ":" + (Visible ? "1" : "0");
        }

        private static string Token(string value)
        {
            return value.Length.ToString(CultureInfo.InvariantCulture) + ":" + value;
        }
    }

    internal sealed class LateBoundFieldState
    {
        public LateBoundFieldState(
            PivotNativeFieldArea area,
            string sourceName,
            string name,
            string caption,
            int position,
            int? function,
            string numberFormat,
            bool repeatLabels,
            IReadOnlyList<bool> subtotals,
            IReadOnlyList<LateBoundMemberState> members,
            string currentPage,
            bool multiplePageItems,
            bool layoutBlankLine = false,
            bool layoutPageBreak = false)
        {
            Area = area;
            SourceName = sourceName ?? string.Empty;
            Name = name ?? string.Empty;
            Caption = caption ?? string.Empty;
            Position = position;
            Function = function;
            NumberFormat = numberFormat ?? string.Empty;
            RepeatLabels = repeatLabels;
            Subtotals = subtotals ?? throw new ArgumentNullException(nameof(subtotals));
            Members = members ?? throw new ArgumentNullException(nameof(members));
            CurrentPage = currentPage ?? string.Empty;
            MultiplePageItems = multiplePageItems;
            LayoutBlankLine = layoutBlankLine;
            LayoutPageBreak = layoutPageBreak;
        }

        public PivotNativeFieldArea Area { get; }

        public string SourceName { get; }

        public string Name { get; }

        public string Caption { get; }

        public int Position { get; }

        public int? Function { get; }

        public string NumberFormat { get; }

        public bool RepeatLabels { get; }

        public IReadOnlyList<bool> Subtotals { get; }

        public IReadOnlyList<LateBoundMemberState> Members { get; }

        public string CurrentPage { get; }

        public bool MultiplePageItems { get; }

        public bool LayoutBlankLine { get; }

        public bool LayoutPageBreak { get; }

        public string CanonicalValue()
        {
            return string.Join(
                "|",
                Area.ToString(),
                Position.ToString(CultureInfo.InvariantCulture),
                Token(SourceName),
                Token(Name),
                Token(Caption),
                Function?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
                Token(NumberFormat),
                RepeatLabels ? "1" : "0",
                string.Concat(Subtotals.Select(value => value ? "1" : "0")),
                Token(CurrentPage),
                MultiplePageItems ? "1" : "0",
                LayoutBlankLine ? "1" : "0",
                LayoutPageBreak ? "1" : "0",
                string.Join(";", Members.Select(member => member.CanonicalValue())));
        }

        private static string Token(string value)
        {
            return value.Length.ToString(CultureInfo.InvariantCulture) + ":" + value;
        }
    }

    internal sealed class LateBoundStyleState
    {
        public LateBoundStyleState(
            int rowAxisLayout,
            bool rowGrand,
            bool columnGrand,
            bool displayFieldCaptions,
            string tableStyleName,
            bool preserveFormatting,
            bool showRowStripes,
            bool showColumnStripes,
            bool displayNullString = true,
            string nullString = "",
            bool displayErrorString = false,
            string errorString = "",
            bool showDrillIndicators = true,
            bool enableDrilldown = true,
            bool visualTotals = true,
            bool subtotalHiddenPageItems = false,
            int pageFieldOrder = 1,
            int pageFieldWrapCount = 0,
            int compactRowIndent = 1,
            bool mergeLabels = false)
        {
            RowAxisLayout = rowAxisLayout;
            RowGrand = rowGrand;
            ColumnGrand = columnGrand;
            DisplayFieldCaptions = displayFieldCaptions;
            TableStyleName = tableStyleName ?? string.Empty;
            PreserveFormatting = preserveFormatting;
            ShowRowStripes = showRowStripes;
            ShowColumnStripes = showColumnStripes;
            DisplayNullString = displayNullString;
            NullString = nullString ?? string.Empty;
            DisplayErrorString = displayErrorString;
            ErrorString = errorString ?? string.Empty;
            ShowDrillIndicators = showDrillIndicators;
            EnableDrilldown = enableDrilldown;
            VisualTotals = visualTotals;
            SubtotalHiddenPageItems = subtotalHiddenPageItems;
            PageFieldOrder = pageFieldOrder;
            PageFieldWrapCount = pageFieldWrapCount;
            CompactRowIndent = compactRowIndent;
            MergeLabels = mergeLabels;
        }

        public int RowAxisLayout { get; }

        public bool RowGrand { get; }

        public bool ColumnGrand { get; }

        public bool DisplayFieldCaptions { get; }

        public string TableStyleName { get; }

        public bool PreserveFormatting { get; }

        public bool ShowRowStripes { get; }

        public bool ShowColumnStripes { get; }

        public bool DisplayNullString { get; }

        public string NullString { get; }

        public bool DisplayErrorString { get; }

        public string ErrorString { get; }

        public bool ShowDrillIndicators { get; }

        public bool EnableDrilldown { get; }

        public bool VisualTotals { get; }

        public bool SubtotalHiddenPageItems { get; }

        public int PageFieldOrder { get; }

        public int PageFieldWrapCount { get; }

        public int CompactRowIndent { get; }

        public bool MergeLabels { get; }

        public string CanonicalValue(bool normalizeDataModelInvariants = false)
        {
            return string.Join(
                "|",
                RowAxisLayout.ToString(CultureInfo.InvariantCulture),
                RowGrand ? "1" : "0",
                ColumnGrand ? "1" : "0",
                DisplayFieldCaptions ? "1" : "0",
                TableStyleName,
                PreserveFormatting ? "1" : "0",
                ShowRowStripes ? "1" : "0",
                ShowColumnStripes ? "1" : "0",
                DisplayNullString ? "1" : "0",
                NullString,
                DisplayErrorString ? "1" : "0",
                ErrorString,
                ShowDrillIndicators ? "1" : "0",
                normalizeDataModelInvariants || EnableDrilldown ? "1" : "0",
                VisualTotals ? "1" : "0",
                normalizeDataModelInvariants || SubtotalHiddenPageItems ? "1" : "0",
                PageFieldOrder.ToString(CultureInfo.InvariantCulture),
                PageFieldWrapCount.ToString(CultureInfo.InvariantCulture),
                CompactRowIndent.ToString(CultureInfo.InvariantCulture),
                MergeLabels ? "1" : "0");
        }
    }

    internal sealed class PivotResultSignature
    {
        public static readonly PivotResultSignature Empty =
            new PivotResultSignature(0, 0, string.Empty);

        public PivotResultSignature(int rows, int columns, string valueFingerprint)
        {
            Rows = rows;
            Columns = columns;
            ValueFingerprint = valueFingerprint ?? string.Empty;
        }

        public int Rows { get; }

        public int Columns { get; }

        public string ValueFingerprint { get; }

        public string CanonicalValue()
        {
            return Rows.ToString(CultureInfo.InvariantCulture) + "x" +
                   Columns.ToString(CultureInfo.InvariantCulture) + ":" +
                   ValueFingerprint;
        }
    }

    internal sealed class LateBoundPivotState
    {
        public LateBoundPivotState(
            object originalCache,
            IReadOnlyList<LateBoundFieldState> fields,
            LateBoundStyleState style,
            PivotResultSignature? result = null,
            LateBoundDataAxisState? dataAxis = null)
        {
            OriginalCache = originalCache ?? throw new ArgumentNullException(nameof(originalCache));
            Fields = fields ?? throw new ArgumentNullException(nameof(fields));
            Style = style ?? throw new ArgumentNullException(nameof(style));
            Result = result ?? PivotResultSignature.Empty;
            DataAxis = dataAxis ?? LateBoundDataAxisState.Hidden;
        }

        public object OriginalCache { get; }

        public IReadOnlyList<LateBoundFieldState> Fields { get; }

        public LateBoundStyleState Style { get; }

        public PivotResultSignature Result { get; }

        public LateBoundDataAxisState DataAxis { get; }

        public string CanonicalValue()
        {
            return Style.CanonicalValue() + "\n" + Result.CanonicalValue() + "\n" +
                   DataAxis.CanonicalValue() + "\n" +
                   string.Join("\n", Fields
                       .OrderBy(field => field.Area)
                       .ThenBy(field => field.Position)
                       .Select(field => field.CanonicalValue()));
        }

        public static bool SemanticallyEquals(
            LateBoundPivotState expected,
            LateBoundPivotState actual,
            bool normalizeDataModelInvariants = false)
        {
            if (!string.Equals(
                    expected.Style.CanonicalValue(normalizeDataModelInvariants),
                    actual.Style.CanonicalValue(normalizeDataModelInvariants),
                    StringComparison.Ordinal) ||
                expected.Result.Rows != actual.Result.Rows ||
                expected.Result.Columns != actual.Result.Columns ||
                !string.Equals(
                    expected.Result.ValueFingerprint,
                    actual.Result.ValueFingerprint,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    expected.DataAxis.CanonicalValue(),
                    actual.DataAxis.CanonicalValue(),
                    StringComparison.Ordinal) ||
                expected.Fields.Count != actual.Fields.Count)
            {
                return false;
            }

            LateBoundFieldState[] left = expected.Fields
                .OrderBy(field => field.Area)
                .ThenBy(field => field.Position)
                .ToArray();
            LateBoundFieldState[] right = actual.Fields
                .OrderBy(field => field.Area)
                .ThenBy(field => field.Position)
                .ToArray();
            for (var index = 0; index < left.Length; index++)
            {
                if (!FieldEquals(left[index], right[index])) return false;
            }

            return true;
        }

        private static bool FieldEquals(LateBoundFieldState expected, LateBoundFieldState actual)
        {
            if (expected.Area != actual.Area ||
                expected.Position != actual.Position ||
                !string.Equals(expected.Caption, actual.Caption, StringComparison.OrdinalIgnoreCase) ||
                expected.Function != actual.Function ||
                !string.Equals(expected.NumberFormat, actual.NumberFormat, StringComparison.Ordinal) ||
                expected.RepeatLabels != actual.RepeatLabels ||
                !expected.Subtotals.SequenceEqual(actual.Subtotals) ||
                expected.MultiplePageItems != actual.MultiplePageItems ||
                expected.LayoutBlankLine != actual.LayoutBlankLine ||
                expected.LayoutPageBreak != actual.LayoutPageBreak ||
                !string.Equals(expected.CurrentPage, actual.CurrentPage, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (!string.Equals(
                    Leaf(expected.SourceName),
                    Leaf(actual.SourceName),
                    StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return MemberSequence(expected.Members).SequenceEqual(
                MemberSequence(actual.Members),
                StringComparer.OrdinalIgnoreCase);
        }

        private static IEnumerable<string> MemberSequence(
            IEnumerable<LateBoundMemberState> members)
        {
            return members
                .OrderBy(member => member.Position)
                .Select(member =>
                    member.Position.ToString(CultureInfo.InvariantCulture) + "|" +
                    (string.IsNullOrWhiteSpace(member.Caption)
                        ? Leaf(member.Name)
                        : member.Caption) + "|" +
                    (member.Visible ? "1" : "0"));
        }

        private static string Leaf(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            string normalized = value.Replace("]]", "]");
            string[] parts = normalized.Split(new[] { "].[" }, StringSplitOptions.None);
            return parts[parts.Length - 1].Trim('[', ']');
        }
    }

    internal sealed class GeneratedNames
    {
        private GeneratedNames(
            string queryName,
            string connectionName,
            string sourceAliasName,
            string stagingWorksheetName,
            string stagingPivotTableName,
            string formatBackupWorksheetName,
            string replacementPivotTableName)
        {
            QueryName = queryName;
            ConnectionName = connectionName;
            SourceAliasName = sourceAliasName;
            StagingWorksheetName = stagingWorksheetName;
            StagingPivotTableName = stagingPivotTableName;
            FormatBackupWorksheetName = formatBackupWorksheetName;
            ReplacementPivotTableName = replacementPivotTableName;
        }

        public string QueryName { get; }

        public string ConnectionName { get; }

        public string SourceAliasName { get; }

        public string StagingWorksheetName { get; }

        public string StagingPivotTableName { get; }

        public string FormatBackupWorksheetName { get; }

        public string ReplacementPivotTableName { get; }

        public static GeneratedNames For(string setupId)
        {
            if (string.IsNullOrWhiteSpace(setupId) ||
                setupId.Length > 64 ||
                !setupId.All(character =>
                    char.IsLetterOrDigit(character) ||
                    character == '.' ||
                    character == '_' ||
                    character == '-'))
            {
                throw new ArgumentException(
                    "A bounded path-free setup identifier is required.",
                    nameof(setupId));
            }

            string fingerprint = PivotPlusFingerprint.Create(
                "pivotplus.setup-name.v1",
                setupId);
            string suffix = fingerprint.Substring(fingerprint.Length - 10, 10);
            string readable = new string(setupId
                .Where(char.IsLetterOrDigit)
                .Take(20)
                .ToArray());
            if (string.IsNullOrEmpty(readable)) readable = "Setup";
            string token = readable + "_" + suffix;
            return new GeneratedNames(
                "PivotPlus_" + token + "_Source",
                "PivotPlus_" + token + "_Model",
                "PivotPlus_" + token + "_Range",
                "_PP_" + suffix,
                "PP_Stage_" + suffix,
                "_PPF_" + suffix,
                "PP_Target_" + suffix);
        }
    }
}
