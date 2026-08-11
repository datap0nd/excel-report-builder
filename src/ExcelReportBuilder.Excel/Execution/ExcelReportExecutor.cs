using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using ExcelReportBuilder.Core.Measures;
using ExcelReportBuilder.Core.Planning;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Core.Transforms;
using ExcelReportBuilder.Excel.Ownership;
using ExcelReportBuilder.Excel.Persistence;
using ExcelReportBuilder.Excel.Rendering;
using ExcelReportBuilder.Excel.Validation;

namespace ExcelReportBuilder.Excel.Execution
{
    public sealed class ExcelBuildResult
    {
        public IReadOnlyList<string> DraftWorksheets { get; set; } = Array.Empty<string>();

        public IReadOnlyList<string> PivotTables { get; set; } = Array.Empty<string>();

        public IReadOnlyList<CheckResult> Checks { get; set; } = Array.Empty<CheckResult>();

        public long NormalizedRows { get; set; }

        public CanonicalBackend Backend { get; set; }
    }

    public sealed class ExcelBuildValidationException : InvalidOperationException
    {
        public ExcelBuildValidationException(ExcelBuildResult result)
            : base("The managed draft failed one or more validation checks.")
        {
            Result = result ?? throw new ArgumentNullException(nameof(result));
        }

        public ExcelBuildResult Result { get; }
    }

    /// <summary>
    /// The deterministic Excel mutation boundary. It builds only managed draft,
    /// canonical, pivot, metadata, and check objects and never saves a workbook.
    /// </summary>
    public sealed class ExcelReportExecutor
    {
        internal const int MaximumIndependentPivotFilterPairs = 14;

        private readonly CanonicalDataLoader canonicalDataLoader;
        private readonly ManagedWorksheetService worksheetService;
        private readonly NativePivotTableExecutor pivotExecutor;
        private readonly DenseReportRenderer denseRenderer;
        private readonly WorkbookSpecStore specStore;
        private readonly ReportReconciler reconciler;
        private readonly CanonicalDataAuditor canonicalDataAuditor;
        private readonly SourceTotalLineageResolver sourceTotalLineageResolver;
        private readonly SourceReconciliationAuditor sourceReconciliationAuditor;

        public ExcelReportExecutor(
            CanonicalDataLoader? canonicalDataLoader = null,
            ManagedWorksheetService? worksheetService = null,
            NativePivotTableExecutor? pivotExecutor = null,
            DenseReportRenderer? denseRenderer = null,
            WorkbookSpecStore? specStore = null,
            ReportReconciler? reconciler = null,
            CanonicalDataAuditor? canonicalDataAuditor = null,
            SourceTotalLineageResolver? sourceTotalLineageResolver = null,
            SourceReconciliationAuditor? sourceReconciliationAuditor = null)
        {
            this.canonicalDataLoader = canonicalDataLoader ?? new CanonicalDataLoader();
            this.worksheetService = worksheetService ?? new ManagedWorksheetService();
            this.pivotExecutor = pivotExecutor ?? new NativePivotTableExecutor();
            this.denseRenderer = denseRenderer ?? new DenseReportRenderer();
            this.specStore = specStore ?? new WorkbookSpecStore();
            this.reconciler = reconciler ?? new ReportReconciler();
            this.canonicalDataAuditor = canonicalDataAuditor ?? new CanonicalDataAuditor();
            this.sourceTotalLineageResolver = sourceTotalLineageResolver ?? new SourceTotalLineageResolver();
            this.sourceReconciliationAuditor = sourceReconciliationAuditor ??
                new SourceReconciliationAuditor();
        }

        public ExcelBuildResult BuildManagedDraft(
            dynamic excelApplication,
            ReportSpecV1 specification,
            ReportBuildPlan plan,
            IExcelProgressSink? progressSink = null,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (excelApplication == null)
            {
                throw new ArgumentNullException(nameof(excelApplication));
            }

            if (specification == null)
            {
                throw new ArgumentNullException(nameof(specification));
            }

            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            if (!string.Equals(specification.Id, plan.SpecificationId, StringComparison.Ordinal) ||
                !string.Equals(specification.OwnershipId, plan.OwnershipId, StringComparison.Ordinal) ||
                !string.Equals(specification.SchemaVersion, plan.SchemaVersion, StringComparison.Ordinal) ||
                !string.Equals(
                    ReportSpecDigest.Compute(specification),
                    plan.SpecificationHash,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    ReportBuildPlanDigest.Compute(plan),
                    plan.PlanHash,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The report plan does not match the exact validated specification or was changed after planning. Rebuild the plan before modifying Excel.");
            }

            ExcelExecutionPreflight.DemandSupported(specification, plan);

            progressSink = progressSink ?? NullExcelProgressSink.Instance;
            dynamic workbook = excelApplication.ActiveWorkbook;
            if (workbook == null)
            {
                throw new InvalidOperationException("Open a workbook before building a report.");
            }

            cancellationToken.ThrowIfCancellationRequested();
            var hasRowRemovalTransforms = specification.Transforms.Any(transform =>
                transform is FilterRowsTransform ||
                transform is ExcludeTotalRowsTransform);
            var requiresSourceTransformAudit = plan.Checks.Any(check =>
                check.Kind == ReportCheckKind.TotalPreservation &&
                check.EvaluationScope == CheckEvaluationScope.CanonicalData) ||
                hasRowRemovalTransforms;
            SourceReconciliationAudit? sourceAudit = null;
            if (requiresSourceTransformAudit)
            {
                progressSink.Report(new ExcelProgress
                {
                    Stage = ExcelBuildStage.Checking,
                    Operation = hasRowRemovalTransforms
                        ? "Independently auditing row-removal transforms and additive totals across the complete source. The next bulk Excel reads may temporarily block Excel."
                        : specification.Transforms.Count > 0
                            ? "Independently auditing transformed additive totals across the complete source. The next bulk Excel reads may temporarily block Excel."
                            : "Independently auditing additive totals across the complete source. The next bulk Excel reads may temporarily block Excel.",
                    ManagedObject = specification.Source.WorkbookObjectName,
                    ProjectedRows = plan.Source.ProjectedRows
                });
                dynamic sourceRange = FindSourceRange(
                    workbook,
                    specification.Source.WorkbookObjectName);
                sourceAudit = sourceReconciliationAuditor.AuditRange(
                    (object)sourceRange,
                    specification,
                    plan.Source.SourceRows,
                    cancellationToken);
                progressSink.Report(new ExcelProgress
                {
                    Stage = ExcelBuildStage.Checking,
                    Operation = hasRowRemovalTransforms
                        ? "Exact post-transform row count and filtered source totals are ready."
                        : "Transformed source totals are ready.",
                    ManagedObject = specification.Source.WorkbookObjectName,
                    ProjectedRows = sourceAudit.ExpectedNormalizedRows
                });
            }

            var sourceTotals = plan.Checks.Any(check =>
                    check.Kind == ReportCheckKind.TotalPreservation &&
                    check.EvaluationScope == CheckEvaluationScope.CanonicalData)
                ? sourceAudit != null
                    ? sourceAudit.ExpectedTotals.ToDictionary(
                        total => total.Key,
                        total => total.Value,
                        StringComparer.OrdinalIgnoreCase)
                    : ReadSourceTotals(excelApplication, workbook, specification)
                : new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
            var canonical = canonicalDataLoader.Load(
                workbook,
                specification.Id,
                specification.OwnershipId,
                plan.Source.PowerQueryM,
                plan.Source.ProjectedRows,
                plan.Source.Route == SourceLoadRoute.DataModel
                    ? CanonicalBackend.DataModel
                    : CanonicalBackend.Worksheet,
                progressSink,
                cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();

            var draftNames = new List<string>();
            var pivotNames = new List<string>();
            var scopedPivotTotals = new Dictionary<ManagedBlockMeasureKey, decimal>();
            var scopedOutputTotals = new Dictionary<ManagedBlockMeasureKey, decimal>();
            var outputMinimums = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
            var missingRequiredValues = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
            var checksForDrafts = new List<CheckResult>();
            var renderedDenseBlocks = new List<RenderedDenseBlock>();
            var renderedNativePivots = new List<RenderedNativePivot>();
            var drafts = new Dictionary<string, ManagedDraftContext>(StringComparer.OrdinalIgnoreCase);
            var measures = specification.Measures.ToDictionary(
                measure => measure.Id,
                StringComparer.OrdinalIgnoreCase);
            var styles = specification.Styles.ToDictionary(
                style => style.Id,
                StringComparer.OrdinalIgnoreCase);

            worksheetService.DeleteStaleHiddenPivotWorksheets(
                excelApplication,
                workbook,
                specification.Id,
                plan.Blocks
                    .Where(block => block.OutputMode == ReportOutputMode.DenseGrid)
                    .Select(block => block.OwnershipId + "_pivot_sheet")
                    .ToArray());

            IReadOnlyList<ManagedOutputWorksheetPlan> activeOutputs =
                ManagedOutputLayoutPlanner.Group(specification.Id, plan.Blocks);
            worksheetService.DeleteStaleDraftWorksheets(
                excelApplication,
                workbook,
                specification.Id,
                activeOutputs.Select(output => output.DraftIdentity.ObjectId).ToArray());

            foreach (var output in activeOutputs)
            {
                dynamic managedDraft = worksheetService.GetOrCreateClearedDraft(
                    workbook,
                    output.DraftIdentity,
                    output.LogicalWorksheetName + " draft");
                var draftName = Convert.ToString(managedDraft.Name, CultureInfo.InvariantCulture) ?? string.Empty;
                drafts.Add(
                    ManagedOutputIdentity.LogicalKey(output.LogicalWorksheetName),
                    new ManagedDraftContext
                    {
                        LogicalWorksheetName = output.LogicalWorksheetName,
                        Worksheet = managedDraft,
                        Identity = output.DraftIdentity,
                        WorksheetName = draftName
                    });
                draftNames.Add(draftName);
            }

            foreach (var block in plan.Blocks)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var logicalWorksheetKey = ManagedOutputIdentity.LogicalKey(block.WorksheetName);
                if (!drafts.TryGetValue(logicalWorksheetKey, out var draftContext))
                {
                    throw new InvalidOperationException(
                        "The validated output worksheet layout could not be resolved.");
                }

                dynamic draft = draftContext.Worksheet;
                var draftIdentityForBlock = draftContext.Identity;

                if (block.OutputMode == ReportOutputMode.StandardMatrix ||
                    block.OutputMode == ReportOutputMode.MetricStack)
                {
                    if (block.Pivot.Values.Any(value => value.RequiresPostAggregationCalculation))
                    {
                        throw new InvalidOperationException(
                            "Calculated measures require a dense output block; a native matrix can display aggregate measures only.");
                    }

                    var pivot = pivotExecutor.Build(
                        workbook,
                        draft,
                        block.AnchorCell,
                        specification.Id,
                        block,
                        canonical,
                        progressSink);
                    pivotNames.Add(pivot.PivotTableName);
                    if (block.OutputMode == ReportOutputMode.MetricStack)
                    {
                        dynamic nativePivot = draft.PivotTables(pivot.PivotTableName);
                        if (Convert.ToInt32(nativePivot.DataFields.Count, CultureInfo.InvariantCulture) > 1)
                        {
                            nativePivot.DataPivotField.Orientation = 1;
                            nativePivot.DataPivotField.Position = 1;
                        }
                        nativePivot.RefreshTable();
                    }

                    dynamic builtPivot = draft.PivotTables(pivot.PivotTableName);
                    ReadPivotTotals(
                        block,
                        builtPivot,
                        pivot,
                        scopedPivotTotals,
                        scopedOutputTotals);
                    renderedNativePivots.Add(new RenderedNativePivot { Pivot = builtPivot, Result = pivot });

                    continue;
                }

                var hiddenIdentity = new ManagedObjectIdentity(
                    specification.Id,
                    block.OwnershipId + "_pivot_sheet",
                    ManagedObjectKind.PivotTable);
                dynamic hidden = worksheetService.GetOrCreateHidden(workbook, hiddenIdentity);
                worksheetService.ClearOwned(hidden, hiddenIdentity);
                var pivotResult = pivotExecutor.Build(
                    workbook,
                    hidden,
                    "$A$3",
                    specification.Id,
                    block,
                    canonical,
                    progressSink);
                pivotNames.Add(pivotResult.PivotTableName);
                dynamic nativeHiddenPivot = hidden.PivotTables(pivotResult.PivotTableName);
                var densePlan = CreateDensePlan(
                    block,
                    nativeHiddenPivot,
                    pivotResult,
                    measures,
                    styles,
                    PeriodFieldName(specification.PeriodMapping));
                denseRenderer.Render(draft, draftIdentityForBlock, densePlan, progressSink);
                renderedDenseBlocks.Add(new RenderedDenseBlock
                {
                    Worksheet = draft,
                    WorksheetName = draftContext.WorksheetName,
                    Plan = densePlan
                });
            }

            progressSink.Report(new ExcelProgress
            {
                Stage = ExcelBuildStage.Calculating,
                Operation = "Calculating managed report formulas."
            });
            excelApplication.Calculate();
            cancellationToken.ThrowIfCancellationRequested();
            var denseAudit = ReadDenseOutputStatistics(
                renderedDenseBlocks,
                scopedPivotTotals,
                scopedOutputTotals,
                outputMinimums,
                missingRequiredValues);
            ReadNativePivotStatistics(renderedNativePivots, outputMinimums, missingRequiredValues);
            var comparablePivotTotals = scopedPivotTotals
                .Where(total => scopedOutputTotals.ContainsKey(total.Key))
                .ToDictionary(total => total.Key, total => total.Value);
            var pivotTotals = ManagedOutputAuditor.AggregateByMeasure(comparablePivotTotals);
            var outputTotals = ManagedOutputAuditor.AggregateByMeasure(scopedOutputTotals);

            foreach (var draft in drafts.Values)
            {
                var owned = new ManagedOwnershipGuard().IsOwned(draft.Worksheet, draft.Identity);
                var expectedBlocks = plan.Blocks.Count(block => string.Equals(
                    ManagedOutputIdentity.LogicalKey(block.WorksheetName),
                    ManagedOutputIdentity.LogicalKey(draft.LogicalWorksheetName),
                    StringComparison.Ordinal));
                var renderedBlocks = renderedDenseBlocks.Count(block => string.Equals(
                        block.WorksheetName,
                        draft.WorksheetName,
                        StringComparison.OrdinalIgnoreCase)) +
                    renderedNativePivots.Count(pivot => string.Equals(
                        pivot.Result.WorksheetName,
                        draft.WorksheetName,
                        StringComparison.OrdinalIgnoreCase));
                checksForDrafts.Add(new CheckResult
                {
                    CheckId = "mandatory-draft-" + draft.Identity.ObjectId,
                    Outcome = owned && renderedBlocks == expectedBlocks
                        ? CheckOutcome.Passed
                        : CheckOutcome.Failed,
                    Message = owned && renderedBlocks == expectedBlocks
                        ? "Every planned block on managed draft '" + draft.LogicalWorksheetName + "' was rendered and retained its ownership marker."
                        : "Managed draft '" + draft.LogicalWorksheetName + "' is missing ownership or one or more planned blocks."
                });
            }

            var canonicalAudit = canonicalDataAuditor.AuditDataModel(
                workbook,
                specification.Id,
                specification.OwnershipId,
                plan.Source.PowerQueryM,
                specification.Measures,
                plan.Source.ProjectedRows,
                progressSink,
                cancellationToken);
            long normalizedRows = canonical.Backend == CanonicalBackend.DataModel
                ? CountDataModelRows(workbook, canonical.TableOrConnectionName)
                : CountCanonicalRows(workbook, canonical);
            Dictionary<string, decimal> normalizedTotals =
                canonical.Backend == CanonicalBackend.DataModel
                    ? new Dictionary<string, decimal>(
                        canonicalAudit.Totals,
                        StringComparer.OrdinalIgnoreCase)
                    : ReadCanonicalTotals(workbook, canonical, specification);
            progressSink.Report(new ExcelProgress
            {
                Stage = ExcelBuildStage.Checking,
                Operation = "Reconciling normalized row counts and managed outputs.",
                ProjectedRows = plan.Source.ProjectedRows
            });
            var checks = new List<CheckResult>(reconciler.Reconcile(new ReconciliationSnapshot
            {
                SourceRows = plan.Source.SourceRows,
                ProjectedNormalizedRows = plan.Source.ProjectedRows,
                ExpectedPostTransformNormalizedRows = sourceAudit?.ExpectedNormalizedRows,
                ActualNormalizedRows = normalizedRows,
                SourceTotals = sourceTotals,
                NormalizedTotals = normalizedTotals,
                PivotTotals = pivotTotals,
                OutputTotals = outputTotals,
                OutputMinimums = outputMinimums,
                MissingRequiredValues = missingRequiredValues
            }, plan.Checks));
            checks.Add(new CheckResult
            {
                CheckId = "mandatory-canonical-load-row-count",
                Outcome = normalizedRows == canonicalAudit.ActualRows
                    ? CheckOutcome.Passed
                    : CheckOutcome.Failed,
                Message = normalizedRows == canonicalAudit.ActualRows
                    ? "The loaded canonical row count matches the independently refreshed audit query."
                    : "The loaded canonical row count differs from the independently refreshed audit query.",
                Expected = canonicalAudit.ActualRows,
                Actual = normalizedRows
            });
            checks.AddRange(checksForDrafts);
            if (denseAudit.FormulasChecked > 0)
            {
                var passed = denseAudit.FormulaErrors == 0 &&
                    denseAudit.FormulaMismatches == 0 &&
                    denseAudit.ExpectedValueMismatches == 0;
                checks.Add(new CheckResult
                {
                    CheckId = "mandatory-dense-formula-integrity",
                    Outcome = passed ? CheckOutcome.Passed : CheckOutcome.Failed,
                    Message = passed
                        ? "All managed dense formulas match the typed render plan and independently evaluated PivotTable values."
                        : denseAudit.FormulaErrors.ToString(CultureInfo.InvariantCulture) +
                          " formula errors and " +
                          denseAudit.FormulaMismatches.ToString(CultureInfo.InvariantCulture) +
                          " changed formulas and " +
                          denseAudit.ExpectedValueMismatches.ToString(CultureInfo.InvariantCulture) +
                          " independently evaluated value mismatches were found in the managed draft."
                });
            }
            else
            {
                checks.Add(new CheckResult
                {
                    CheckId = "mandatory-formula-errors",
                    Outcome = CheckOutcome.Passed,
                    Message = "The selected native PivotTable layout contains no generated report formulas."
                });
            }

            var requiredBlockTotals = plan.Blocks
                .SelectMany(block => block.Pivot.Values
                    .Where(value => value.Expression is AggregateMeasureExpression &&
                                    value.PeriodSliceIds.Count == 0)
                    .Select(value => new ManagedBlockMeasureKey(block.BlockId, value.MeasureId)))
                .Distinct()
                .ToList();
            checks.AddRange(ManagedOutputAuditor.ReconcileBlockTotals(
                requiredBlockTotals,
                scopedPivotTotals,
                scopedOutputTotals,
                ManagedOutputAuditor.FormulaTolerance));
            var missingGrandTotals = requiredBlockTotals
                .Where(key => !scopedPivotTotals.ContainsKey(key) || !scopedOutputTotals.ContainsKey(key))
                .ToList();
            checks.Add(new CheckResult
            {
                CheckId = "mandatory-managed-grand-totals",
                Outcome = missingGrandTotals.Count == 0 ? CheckOutcome.Passed : CheckOutcome.Failed,
                Message = missingGrandTotals.Count == 0
                    ? requiredBlockTotals.Count == 0
                        ? "Calculated Values are validated through managed formula integrity checks."
                        : "Every unsliced aggregate Value has both a PivotTable total and an independently read output total."
                    : "One or more unsliced aggregate Values are missing a PivotTable or output grand total."
            });

            var buildResult = new ExcelBuildResult
            {
                DraftWorksheets = draftNames,
                PivotTables = pivotNames,
                Checks = checks,
                NormalizedRows = normalizedRows,
                Backend = canonical.Backend
            };
            if (checks.Any(check => check.Outcome == CheckOutcome.Failed))
            {
                throw new ExcelBuildValidationException(buildResult);
            }

            specStore.Save(workbook, specification);
            progressSink.Report(new ExcelProgress
            {
                Stage = ExcelBuildStage.Complete,
                Operation = "Managed draft complete. The workbook has not been saved.",
                ProjectedRows = plan.Source.ProjectedRows,
                CompletedChecks = checks.Count
            });

            return buildResult;
        }

        private static DenseGridPlan CreateDensePlan(
            DenseReportBlockPlan block,
            dynamic nativePivot,
            PivotBuildResult pivot,
            IReadOnlyDictionary<string, MeasureDefinition> measures,
            IReadOnlyDictionary<string, PresentationStyleSpec> styles,
            string? periodFieldName)
        {
            var result = new DenseGridPlan
            {
                BlockId = block.BlockId,
                AnchorCell = block.AnchorCell,
                OwnedRowCount = block.OwnedRange.RowCount,
                OwnedColumnCount = block.OwnedRange.ColumnCount,
                Styles = styles
            };
            var nextRow = 0;
            foreach (var header in block.Presentation.Headers)
            {
                result.Cells.Add(new DenseCellWrite
                {
                    RelativeRow = header.RelativeRow,
                    RelativeColumn = header.RelativeColumn,
                    Kind = DenseCellValueKind.Text,
                    Value = header.Text,
                    ColumnSpan = header.ColumnSpan,
                    StyleId = header.StyleId
                });
                nextRow = Math.Max(nextRow, header.RelativeRow + 1);
            }

            if (block.Presentation.Headers.Count == 0 && !string.IsNullOrWhiteSpace(block.Title))
            {
                result.Cells.Add(new DenseCellWrite
                {
                    RelativeRow = 0,
                    RelativeColumn = 0,
                    Kind = DenseCellValueKind.Text,
                    Value = block.Title,
                    ColumnSpan = Math.Max(2, block.Pivot.Values.Count + 1),
                    StyleId = block.Presentation.HeaderStyleId
                });
                nextRow = 1;
            }

            var rowPaths = DenseAxisPlanner.Build(
                ReadAxisPaths((object)nativePivot, true),
                block.Pivot.Rows,
                block.Presentation.SubtotalStyleId,
                (measureId, filters) => ReadPivotScore(nativePivot, pivot, measureId, filters));
            var columnPaths = DenseAxisPlanner.Build(
                ReadAxisPaths((object)nativePivot, false),
                block.Pivot.Columns,
                block.Presentation.SubtotalStyleId,
                (measureId, filters) => ReadPivotScore(nativePivot, pivot, measureId, filters));
            if (block.Pivot.Columns.Count > 0 && block.Presentation.Options.ShowColumnGrandTotals)
            {
                var grandColumn = new DenseAxisPath
                {
                    MemberFilterSets =
                    {
                        (IReadOnlyList<PivotFilterItem>)Array.Empty<PivotFilterItem>()
                    },
                    StyleId = block.Pivot.GrandTotals.StyleId ?? block.Presentation.GrandTotalStyleId
                };
                if (block.Pivot.GrandTotals.ColumnPlacement == TotalPlacement.BeforeMembers)
                {
                    columnPaths.Insert(0, grandColumn);
                }
                else
                {
                    columnPaths.Add(grandColumn);
                }
            }

            var rowFieldCount = Math.Max(1, block.Pivot.Rows.Count);
            for (var fieldIndex = 0; fieldIndex < block.Pivot.Rows.Count; fieldIndex++)
            {
                result.Cells.Add(new DenseCellWrite
                {
                    RelativeRow = nextRow,
                    RelativeColumn = fieldIndex,
                    Kind = DenseCellValueKind.Text,
                    Value = block.Pivot.Rows[fieldIndex].Caption ?? block.Pivot.Rows[fieldIndex].Field,
                    StyleId = block.Presentation.HeaderStyleId
                });
            }

            var periodFilters = ResolvePeriodFilters(nativePivot, block, periodFieldName);
            var outputColumns = new List<DenseOutputColumn>();
            var nextOutputColumn = rowFieldCount;
            foreach (var columnPath in columnPaths)
            {
                nextOutputColumn += AddColumnSpacers(
                    result,
                    block,
                    columnPath,
                    nextOutputColumn);
                foreach (var value in block.Pivot.Values)
                {
                    var sliceIds = value.PeriodSliceIds.Count == 0
                        ? new string?[] { null }
                        : value.PeriodSliceIds.Select(id => (string?)id).ToArray();
                    foreach (var sliceId in sliceIds)
                    {
                        var columnLabel = string.Join(
                            " / ",
                            columnPath.DisplayItems
                                .Select(item => Convert.ToString(item.Value, CultureInfo.InvariantCulture))
                                .Where(label => !string.IsNullOrWhiteSpace(label)));
                        if (columnPath.DisplayItems.Count == 0 && block.Pivot.Columns.Count > 0)
                        {
                            columnLabel = block.Pivot.GrandTotals.ColumnLabel;
                        }

                        if (!string.IsNullOrWhiteSpace(columnLabel))
                        {
                            columnLabel += " / ";
                        }

                        columnLabel += value.Label;
                        IReadOnlyList<IReadOnlyList<PivotFilterItem>> sliceFilterSets =
                            new[] { (IReadOnlyList<PivotFilterItem>)Array.Empty<PivotFilterItem>() };
                        if (!string.IsNullOrWhiteSpace(sliceId))
                        {
                            if (!periodFilters.TryGetValue(sliceId!, out sliceFilterSets))
                            {
                                throw new InvalidOperationException("A Value references an unresolved period slice.");
                            }

                            var slice = block.Presentation.PeriodSlices.First(item =>
                                string.Equals(item.Id, sliceId, StringComparison.OrdinalIgnoreCase));
                            columnLabel += " / " + slice.Label;
                        }

                        outputColumns.Add(new DenseOutputColumn
                        {
                            ColumnPath = columnPath,
                            Value = value,
                            SliceFilterSets = sliceFilterSets,
                            IsSliced = !string.IsNullOrWhiteSpace(sliceId),
                            RelativeColumn = nextOutputColumn++
                        });
                        result.Cells.Add(new DenseCellWrite
                        {
                            RelativeRow = nextRow,
                            RelativeColumn = outputColumns[outputColumns.Count - 1].RelativeColumn,
                            Kind = DenseCellValueKind.Text,
                            Value = columnLabel,
                            StyleId = columnPath.StyleId ?? block.Presentation.HeaderStyleId
                        });
                    }
                }
            }

            var hasExplicitGrandColumn = block.Pivot.Columns.Count > 0 &&
                outputColumns.Any(output => output.ColumnPath.DisplayItems.Count == 0);

            var formulaCompiler = new MeasureFormulaCompiler();
            var measureEvaluator = new PivotMeasureEvaluator(
                measures,
                pivot,
                (caption, filters) => ReadPivotAggregateExpectation(
                    (object)nativePivot,
                    caption,
                    filters));
            var rowFieldOrder = block.Pivot.Rows.Select(row => row.Field).ToList();
            var detailStride = block.Presentation.Options.InsertBlankRows ? 2 : 1;
            result.FreezeHeaders = block.Presentation.Options.FreezeHeaders;
            result.FreezeRelativeRow = nextRow + 1;
            var detailRow = nextRow + 1 +
                            (block.Presentation.Options.ShowRowGrandTotals &&
                             block.Pivot.GrandTotals.RowPlacement == TotalPlacement.BeforeMembers
                                ? 1
                                : 0);
            for (var rowIndex = 0; rowIndex < rowPaths.Count; rowIndex++)
            {
                var rowPath = rowPaths[rowIndex];
                detailRow += AddRowSpacers(result, block, rowPath, detailRow);
                for (var fieldIndex = 0; fieldIndex < rowPath.DisplayItems.Count && fieldIndex < rowFieldCount; fieldIndex++)
                {
                    result.Cells.Add(new DenseCellWrite
                    {
                        RelativeRow = detailRow,
                        RelativeColumn = fieldIndex,
                        Kind = DenseCellValueKind.Text,
                        Value = rowPath.DisplayItems[fieldIndex].Value,
                        StyleId = rowPath.StyleId ?? block.Presentation.BodyStyleId,
                        IndentLevel = Math.Min(15, fieldIndex * block.Presentation.Options.RowIndent)
                    });
                }

                for (var outputIndex = 0; outputIndex < outputColumns.Count; outputIndex++)
                {
                    var output = outputColumns[outputIndex];
                    var memberSets = new List<IReadOnlyList<PivotFilterItem>>();
                    foreach (var rowFilters in rowPath.MemberFilterSets)
                    {
                        foreach (var columnFilters in output.ColumnPath.MemberFilterSets)
                        {
                            foreach (var sliceFilters in output.SliceFilterSets)
                            {
                                var combined = new List<PivotFilterItem>(rowFilters);
                                combined.AddRange(columnFilters);
                                combined.AddRange(sliceFilters);
                                memberSets.Add(combined);
                            }
                        }
                    }

                    var formula = formulaCompiler.CompileAcrossMemberSets(
                        output.Value.MeasureId,
                        measures,
                        pivot,
                        memberSets,
                        periodFilters,
                        rowFieldOrder);
                    var expectedValue = measureEvaluator.EvaluateAcrossMemberSets(
                        output.Value.MeasureId,
                        memberSets,
                        periodFilters,
                        rowFieldOrder);
                    result.Cells.Add(new DenseCellWrite
                    {
                        RelativeRow = detailRow,
                        RelativeColumn = output.RelativeColumn,
                        Kind = DenseCellValueKind.Formula,
                        Formula = formula,
                        ExpectedFormulaValue = expectedValue,
                        NumberFormat = output.Value.NumberFormat,
                        StyleId = rowPath.StyleId ?? output.ColumnPath.StyleId ?? block.Presentation.BodyStyleId,
                        MeasureId = output.Value.MeasureId,
                        IsOutputTotal = IsDenseOutputTotalContribution(
                            block.Presentation.Options.ShowRowGrandTotals,
                            rowPath.IsSubtotal,
                            output.IsSliced,
                            output.Value.Expression is AggregateMeasureExpression,
                            block.Pivot.Columns.Count,
                            hasExplicitGrandColumn,
                            output.ColumnPath.IsSubtotal,
                            output.ColumnPath.DisplayItems.Count == 0)
                    });
                }

                detailRow += detailStride;
            }

            if (block.Presentation.Options.ShowRowGrandTotals)
            {
                var grandTotalRow = block.Pivot.GrandTotals.RowPlacement == TotalPlacement.BeforeMembers
                    ? nextRow + 1
                    : detailRow;
                result.Cells.Add(new DenseCellWrite
                {
                    RelativeRow = grandTotalRow,
                    RelativeColumn = 0,
                    Kind = DenseCellValueKind.Text,
                    Value = block.Pivot.GrandTotals.RowLabel,
                    StyleId = block.Pivot.GrandTotals.StyleId ?? block.Presentation.GrandTotalStyleId
                });
                for (var outputIndex = 0; outputIndex < outputColumns.Count; outputIndex++)
                {
                    var output = outputColumns[outputIndex];
                    var totalMemberSets = new List<IReadOnlyList<PivotFilterItem>>();
                    foreach (var columnFilters in output.ColumnPath.MemberFilterSets)
                    {
                        foreach (var sliceFilters in output.SliceFilterSets)
                        {
                            var combined = new List<PivotFilterItem>(columnFilters);
                            combined.AddRange(sliceFilters);
                            totalMemberSets.Add(combined);
                        }
                    }

                    var formula = formulaCompiler.CompileAcrossMemberSets(
                        output.Value.MeasureId,
                        measures,
                        pivot,
                        totalMemberSets,
                        periodFilters,
                        rowFieldOrder);
                    var expectedValue = measureEvaluator.EvaluateAcrossMemberSets(
                        output.Value.MeasureId,
                        totalMemberSets,
                        periodFilters,
                        rowFieldOrder);
                    result.Cells.Add(new DenseCellWrite
                    {
                        RelativeRow = grandTotalRow,
                        RelativeColumn = output.RelativeColumn,
                        Kind = DenseCellValueKind.Formula,
                        Formula = formula,
                        ExpectedFormulaValue = expectedValue,
                        NumberFormat = output.Value.NumberFormat,
                        StyleId = block.Pivot.GrandTotals.StyleId ?? block.Presentation.GrandTotalStyleId,
                        MeasureId = output.Value.MeasureId,
                        IsOutputTotal = !output.IsSliced &&
                            (block.Pivot.Columns.Count == 0 ||
                             !hasExplicitGrandColumn ||
                             output.ColumnPath.DisplayItems.Count == 0)
                    });
                }
            }

            return result;
        }

        internal static bool IsDenseOutputTotalContribution(
            bool showRowGrandTotals,
            bool rowIsSubtotal,
            bool outputIsSliced,
            bool isAggregateMeasure,
            int columnFieldCount,
            bool hasExplicitGrandColumn,
            bool columnIsSubtotal,
            bool columnIsGrandTotal)
        {
            if (showRowGrandTotals || rowIsSubtotal || outputIsSliced || !isAggregateMeasure)
            {
                return false;
            }

            if (columnFieldCount == 0)
            {
                return true;
            }

            return hasExplicitGrandColumn
                ? columnIsGrandTotal
                : !columnIsSubtotal;
        }

        private static int AddRowSpacers(
            DenseGridPlan result,
            DenseReportBlockPlan block,
            DenseAxisPath path,
            int relativeRow)
        {
            var level = path.IsSubtotal
                ? path.SubtotalLevel
                : Math.Max(0, path.DisplayItems.Count - 1);
            var added = 0;
            foreach (var spacer in block.Presentation.Spacers.Where(item =>
                         item.Axis == SpacerAxis.Row && item.BeforeLevel == level))
            {
                for (var index = 0; index < spacer.Count; index++)
                {
                    result.RowSizes.Add(new DenseDimensionSize
                    {
                        RelativeIndex = relativeRow + added,
                        Size = spacer.Size ?? 6d
                    });
                    added++;
                }
            }

            return added;
        }

        private static int AddColumnSpacers(
            DenseGridPlan result,
            DenseReportBlockPlan block,
            DenseAxisPath path,
            int relativeColumn)
        {
            var level = Math.Max(0, path.DisplayItems.Count - 1);
            var added = 0;
            foreach (var spacer in block.Presentation.Spacers.Where(item =>
                         item.Axis == SpacerAxis.Column && item.BeforeLevel == level))
            {
                for (var index = 0; index < spacer.Count; index++)
                {
                    result.ColumnSizes.Add(new DenseDimensionSize
                    {
                        RelativeIndex = relativeColumn + added,
                        Size = spacer.Size ?? 2d
                    });
                    added++;
                }
            }

            return added;
        }

        private static IReadOnlyDictionary<string, IReadOnlyList<IReadOnlyList<PivotFilterItem>>> ResolvePeriodFilters(
            dynamic pivot,
            DenseReportBlockPlan block,
            string? periodFieldName)
        {
            var result = new Dictionary<string, IReadOnlyList<IReadOnlyList<PivotFilterItem>>>(
                StringComparer.OrdinalIgnoreCase);
            if (block.Presentation.PeriodSlices.Count == 0)
            {
                return result;
            }

            if (string.IsNullOrWhiteSpace(periodFieldName))
            {
                throw new InvalidOperationException("Period slices require an explicit period mapping.");
            }

            dynamic field = pivot.PivotFields(periodFieldName);
            dynamic items = field.PivotItems();
            var members = new List<PeriodMember>();
            var count = Convert.ToInt32(items.Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                dynamic item = items.Item(index);
                object rawValue;
                try
                {
                    rawValue = item.Value;
                }
                catch (Exception)
                {
                    rawValue = item.Name;
                }

                if (!TryConvertPeriod(rawValue, out var period))
                {
                    continue;
                }

                members.Add(new PeriodMember { Period = period, PivotValue = period });
            }

            var resolved = PeriodSliceResolver.BindResolved(
                block.Presentation.ResolvedPeriodSlices,
                members);
            foreach (var pair in resolved)
            {
                result[pair.Key] = pair.Value
                    .Select(value => (IReadOnlyList<PivotFilterItem>)new[]
                    {
                        new PivotFilterItem { Field = periodFieldName!, Value = value }
                    })
                    .ToList();
            }

            return result;
        }

        private static decimal ReadPivotScore(
            dynamic pivot,
            PivotBuildResult pivotResult,
            string measureId,
            IReadOnlyList<PivotFilterItem> filters)
        {
            var descriptors = pivotResult.DataFields
                .Where(field => string.Equals(
                    field.MeasureId,
                    measureId,
                    StringComparison.OrdinalIgnoreCase))
                .GroupBy(field => field.PivotCaption, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .ToList();
            if (descriptors.Count != 1 || descriptors[0].Filters.Count != 0)
            {
                throw new NotSupportedException(
                    "Dense Top N requires one direct aggregate component for its ranking Value.");
            }

            var caption = descriptors[0].PivotCaption;
            dynamic cell;
            switch (filters.Count)
            {
                case 0:
                    cell = pivot.GetPivotData(caption);
                    break;
                case 1:
                    cell = pivot.GetPivotData(
                        caption,
                        filters[0].Field,
                        filters[0].Value);
                    break;
                case 2:
                    cell = pivot.GetPivotData(
                        caption,
                        filters[0].Field,
                        filters[0].Value,
                        filters[1].Field,
                        filters[1].Value);
                    break;
                case 3:
                    cell = pivot.GetPivotData(
                        caption,
                        filters[0].Field,
                        filters[0].Value,
                        filters[1].Field,
                        filters[1].Value,
                        filters[2].Field,
                        filters[2].Value);
                    break;
                case 4:
                    cell = pivot.GetPivotData(
                        caption,
                        filters[0].Field,
                        filters[0].Value,
                        filters[1].Field,
                        filters[1].Value,
                        filters[2].Field,
                        filters[2].Value,
                        filters[3].Field,
                        filters[3].Value);
                    break;
                case 5:
                    cell = pivot.GetPivotData(
                        caption,
                        filters[0].Field, filters[0].Value,
                        filters[1].Field, filters[1].Value,
                        filters[2].Field, filters[2].Value,
                        filters[3].Field, filters[3].Value,
                        filters[4].Field, filters[4].Value);
                    break;
                case 6:
                    cell = pivot.GetPivotData(
                        caption,
                        filters[0].Field, filters[0].Value,
                        filters[1].Field, filters[1].Value,
                        filters[2].Field, filters[2].Value,
                        filters[3].Field, filters[3].Value,
                        filters[4].Field, filters[4].Value,
                        filters[5].Field, filters[5].Value);
                    break;
                case 7:
                    cell = pivot.GetPivotData(
                        caption,
                        filters[0].Field, filters[0].Value,
                        filters[1].Field, filters[1].Value,
                        filters[2].Field, filters[2].Value,
                        filters[3].Field, filters[3].Value,
                        filters[4].Field, filters[4].Value,
                        filters[5].Field, filters[5].Value,
                        filters[6].Field, filters[6].Value);
                    break;
                case 8:
                    cell = pivot.GetPivotData(
                        caption,
                        filters[0].Field, filters[0].Value,
                        filters[1].Field, filters[1].Value,
                        filters[2].Field, filters[2].Value,
                        filters[3].Field, filters[3].Value,
                        filters[4].Field, filters[4].Value,
                        filters[5].Field, filters[5].Value,
                        filters[6].Field, filters[6].Value,
                        filters[7].Field, filters[7].Value);
                    break;
                default:
                    throw new NotSupportedException(
                        "Dense Top N supports up to eight nested member fields in version 1.");
            }

            return Convert.ToDecimal(cell.Value2, CultureInfo.InvariantCulture);
        }

        internal static DenseFormulaExpectation ReadPivotAggregateExpectation(
            object pivot,
            string dataFieldCaption,
            IReadOnlyList<PivotFilterItem> filters)
        {
            if (filters.Count > MaximumIndependentPivotFilterPairs)
            {
                throw new NotSupportedException(
                    "Independent PivotTable validation supports up to fourteen field filters per aggregate term.");
            }

            var arguments = new object[1 + filters.Count * 2];
            arguments[0] = dataFieldCaption;
            for (var index = 0; index < filters.Count; index++)
            {
                arguments[1 + index * 2] = filters[index].Field;
                arguments[2 + index * 2] = filters[index].Value ?? string.Empty;
            }

            object? pivotCell;
            try
            {
                pivotCell = pivot.GetType().InvokeMember(
                    "GetPivotData",
                    BindingFlags.Instance |
                    BindingFlags.Public |
                    BindingFlags.InvokeMethod |
                    BindingFlags.OptionalParamBinding,
                    null,
                    pivot,
                    arguments,
                    CultureInfo.InvariantCulture);
            }
            catch (TargetInvocationException exception) when (exception.InnerException is COMException)
            {
                throw new InvalidOperationException(
                    "The native PivotTable aggregate could not be read independently.",
                    exception.InnerException);
            }
            catch (COMException exception)
            {
                throw new InvalidOperationException(
                    "The native PivotTable aggregate could not be read independently.",
                    exception);
            }

            if (pivotCell == null)
            {
                throw new InvalidOperationException(
                    "The native PivotTable returned no aggregate cell for independent validation.");
            }

            dynamic cell = pivotCell;
            try
            {
                var displayed = Convert.ToString(cell.Text, CultureInfo.InvariantCulture) ?? string.Empty;
                if (ManagedOutputAuditor.IsExcelErrorDisplay(displayed))
                {
                    throw new InvalidOperationException(
                        "The native PivotTable aggregate contains an Excel error and cannot be independently validated.");
                }
            }
            catch (COMException exception)
            {
                throw new InvalidOperationException(
                    "The native PivotTable aggregate display value could not be read independently.",
                    exception);
            }

            object? value;
            try
            {
                value = cell.Value2;
            }
            catch (COMException exception)
            {
                throw new InvalidOperationException(
                    "The native PivotTable aggregate value could not be read independently.",
                    exception);
            }

            if (value == null || string.IsNullOrWhiteSpace(Convert.ToString(value, CultureInfo.InvariantCulture)))
            {
                return DenseFormulaExpectation.Number(0m);
            }

            try
            {
                return DenseFormulaExpectation.Number(
                    Convert.ToDecimal(value, CultureInfo.InvariantCulture));
            }
            catch (Exception exception) when (
                exception is FormatException ||
                exception is InvalidCastException ||
                exception is OverflowException)
            {
                throw new InvalidOperationException(
                    "A PivotTable aggregate returned a non-numeric value during independent validation.",
                    exception);
            }
        }

        private static bool TryConvertPeriod(object rawValue, out DateTime period)
        {
            if (rawValue is DateTime date)
            {
                period = date.Date;
                return true;
            }

            if (rawValue is double serial)
            {
                try
                {
                    period = DateTime.FromOADate(serial).Date;
                    return true;
                }
                catch (ArgumentException)
                {
                }
            }

            return DateTime.TryParse(
                Convert.ToString(rawValue, CultureInfo.InvariantCulture),
                CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces,
                out period);
        }

        private static string? PeriodFieldName(PeriodMappingSpec? mapping)
        {
            if (mapping == null)
            {
                return null;
            }

            return mapping.Kind == PeriodMappingKind.LongDateColumn
                ? mapping.DateColumn
                : mapping.PeriodColumnName;
        }

        private sealed class DenseOutputColumn
        {
            public DenseAxisPath ColumnPath { get; set; } = new DenseAxisPath();

            public PivotValuePlan Value { get; set; } = new PivotValuePlan();

            public IReadOnlyList<IReadOnlyList<PivotFilterItem>> SliceFilterSets { get; set; } =
                new[] { (IReadOnlyList<PivotFilterItem>)Array.Empty<PivotFilterItem>() };

            public bool IsSliced { get; set; }

            public int RelativeColumn { get; set; }
        }

        private sealed class RenderedDenseBlock
        {
            public dynamic Worksheet { get; set; } = null!;

            public string WorksheetName { get; set; } = string.Empty;

            public DenseGridPlan Plan { get; set; } = new DenseGridPlan();
        }

        private sealed class ManagedDraftContext
        {
            public string LogicalWorksheetName { get; set; } = string.Empty;

            public dynamic Worksheet { get; set; } = null!;

            public ManagedObjectIdentity Identity { get; set; } = null!;

            public string WorksheetName { get; set; } = string.Empty;
        }

        private sealed class RenderedNativePivot
        {
            public dynamic Pivot { get; set; } = null!;

            public PivotBuildResult Result { get; set; } = new PivotBuildResult();
        }

        private sealed class DenseFormulaAudit
        {
            public int FormulasChecked { get; set; }

            public int FormulaErrors { get; set; }

            public int FormulaMismatches { get; set; }

            public int ExpectedValueMismatches { get; set; }
        }

        private static DenseFormulaAudit ReadDenseOutputStatistics(
            IReadOnlyList<RenderedDenseBlock> rendered,
            IDictionary<ManagedBlockMeasureKey, decimal> pivotTotals,
            IDictionary<ManagedBlockMeasureKey, decimal> outputTotals,
            IDictionary<string, decimal> minimums,
            IDictionary<string, long> missing)
        {
            var audit = new DenseFormulaAudit();
            foreach (var block in rendered)
            {
                var anchor = CellAddress.Parse(block.Plan.AnchorCell);
                foreach (var write in block.Plan.Cells.Where(cell =>
                             cell.Kind == DenseCellValueKind.Formula &&
                             !string.IsNullOrWhiteSpace(cell.MeasureId)))
                {
                    dynamic cell = block.Worksheet.Cells[
                        anchor.Row + write.RelativeRow,
                        anchor.Column + write.RelativeColumn];
                    audit.FormulasChecked++;
                    var hasExcelError = false;
                    var formulaReadFailed = false;
                    var actualFormula = string.Empty;
                    var displayed = string.Empty;
                    try
                    {
                        actualFormula = Convert.ToString(cell.Formula, CultureInfo.InvariantCulture) ?? string.Empty;
                        displayed = Convert.ToString(cell.Text, CultureInfo.InvariantCulture) ?? string.Empty;
                    }
                    catch (Exception)
                    {
                        hasExcelError = true;
                        formulaReadFailed = true;
                    }

                    object? value;
                    var valueReadFailed = false;
                    try
                    {
                        value = cell.Value2;
                    }
                    catch (Exception)
                    {
                        value = null;
                        valueReadFailed = true;
                    }

                    if (formulaReadFailed || valueReadFailed ||
                        write.Formula == null || write.ExpectedFormulaValue == null)
                    {
                        if (formulaReadFailed || write.Formula == null)
                        {
                            audit.FormulaMismatches++;
                            audit.FormulaErrors++;
                        }

                        audit.ExpectedValueMismatches++;
                    }
                    else
                    {
                        var cellAudit = ManagedOutputAuditor.AuditFormulaCell(
                            write.Formula,
                            write.ExpectedFormulaValue,
                            actualFormula,
                            displayed,
                            value);
                        hasExcelError = cellAudit.HasExcelError;
                        if (cellAudit.FormulaChanged)
                        {
                            audit.FormulaMismatches++;
                        }

                        if (cellAudit.HasExcelError)
                        {
                            audit.FormulaErrors++;
                        }

                        if (!cellAudit.ValueMatches)
                        {
                            audit.ExpectedValueMismatches++;
                        }
                    }

                    if (write.IsOutputTotal &&
                        write.ExpectedFormulaValue != null &&
                        write.ExpectedFormulaValue.Kind == DenseFormulaExpectationKind.Number &&
                        write.ExpectedFormulaValue.NumericValue.HasValue)
                    {
                        var key = new ManagedBlockMeasureKey(block.Plan.BlockId, write.MeasureId!);
                        pivotTotals[key] = pivotTotals.TryGetValue(key, out var total)
                            ? checked(total + write.ExpectedFormulaValue.NumericValue.Value)
                            : write.ExpectedFormulaValue.NumericValue.Value;
                    }

                    if (hasExcelError ||
                        value == null ||
                        string.IsNullOrWhiteSpace(Convert.ToString(value, CultureInfo.InvariantCulture)))
                    {
                        missing[write.MeasureId!] = missing.TryGetValue(write.MeasureId!, out var count)
                            ? count + 1
                            : 1;
                        continue;
                    }

                    try
                    {
                        var number = Convert.ToDecimal(value, CultureInfo.InvariantCulture);
                        if (write.IsOutputTotal)
                        {
                            var key = new ManagedBlockMeasureKey(block.Plan.BlockId, write.MeasureId!);
                            outputTotals[key] = outputTotals.TryGetValue(key, out var total)
                                ? checked(total + number)
                                : number;
                        }

                        if (!minimums.TryGetValue(write.MeasureId!, out var current) || number < current)
                        {
                            minimums[write.MeasureId!] = number;
                        }

                        if (!missing.ContainsKey(write.MeasureId!))
                        {
                            missing[write.MeasureId!] = 0;
                        }
                    }
                    catch (Exception)
                    {
                        missing[write.MeasureId!] = missing.TryGetValue(write.MeasureId!, out var count)
                            ? count + 1
                            : 1;
                    }
                }
            }

            return audit;
        }

        private static void ReadNativePivotStatistics(
            IReadOnlyList<RenderedNativePivot> rendered,
            IDictionary<string, decimal> minimums,
            IDictionary<string, long> missing)
        {
            foreach (var renderedPivot in rendered)
            {
                dynamic? body = null;
                try
                {
                    body = renderedPivot.Pivot.DataBodyRange;
                }
                catch (Exception)
                {
                }

                if (body == null)
                {
                    continue;
                }

                var rowCount = Convert.ToInt32(body.Rows.Count, CultureInfo.InvariantCulture);
                var columnCount = Convert.ToInt32(body.Columns.Count, CultureInfo.InvariantCulture);
                for (var row = 1; row <= rowCount; row++)
                {
                    for (var column = 1; column <= columnCount; column++)
                    {
                        dynamic cell = body.Cells[row, column];
                        string caption;
                        try
                        {
                            caption = Convert.ToString(cell.PivotCell.DataField.Name, CultureInfo.InvariantCulture) ?? string.Empty;
                        }
                        catch (Exception)
                        {
                            continue;
                        }

                        var descriptor = renderedPivot.Result.DataFields.FirstOrDefault(field =>
                            string.Equals(field.PivotCaption, caption, StringComparison.OrdinalIgnoreCase));
                        if (descriptor == null)
                        {
                            continue;
                        }

                        object? value = cell.Value2;
                        if (value == null || string.IsNullOrWhiteSpace(Convert.ToString(value, CultureInfo.InvariantCulture)))
                        {
                            missing[descriptor.MeasureId] = missing.TryGetValue(descriptor.MeasureId, out var count)
                                ? count + 1
                                : 1;
                            continue;
                        }

                        try
                        {
                            var number = Convert.ToDecimal(value, CultureInfo.InvariantCulture);
                            if (!minimums.TryGetValue(descriptor.MeasureId, out var current) || number < current)
                            {
                                minimums[descriptor.MeasureId] = number;
                            }

                            if (!missing.ContainsKey(descriptor.MeasureId))
                            {
                                missing[descriptor.MeasureId] = 0;
                            }
                        }
                        catch (Exception)
                        {
                            missing[descriptor.MeasureId] = missing.TryGetValue(descriptor.MeasureId, out var count)
                                ? count + 1
                                : 1;
                        }
                    }
                }
            }
        }

        private static List<List<PivotFilterItem>> ReadAxisPaths(dynamic pivot, bool rows)
        {
            var result = new List<List<PivotFilterItem>>();
            dynamic? body = null;
            try
            {
                body = pivot.DataBodyRange;
            }
            catch (Exception)
            {
            }

            if (body == null)
            {
                return result;
            }

            var count = Convert.ToInt32(rows ? body.Rows.Count : body.Columns.Count, CultureInfo.InvariantCulture);
            var seen = new HashSet<string>(StringComparer.Ordinal);
            for (var index = 1; index <= count; index++)
            {
                dynamic cell = rows ? body.Cells[index, 1] : body.Cells[1, index];
                dynamic pivotCell = cell.PivotCell;
                dynamic items = rows ? pivotCell.RowItems : pivotCell.ColumnItems;
                var path = new List<PivotFilterItem>();
                var itemCount = Convert.ToInt32(items.Count, CultureInfo.InvariantCulture);
                for (var itemIndex = 1; itemIndex <= itemCount; itemIndex++)
                {
                    dynamic item = items.Item(itemIndex);
                    string fieldName;
                    try
                    {
                        fieldName = Convert.ToString(item.Parent.Name, CultureInfo.InvariantCulture) ?? string.Empty;
                    }
                    catch (Exception)
                    {
                        continue;
                    }

                    if (string.Equals(fieldName, "Values", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(fieldName, "Data", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(fieldName, "Σ Values", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    path.Add(new PivotFilterItem
                    {
                        Field = fieldName,
                        Value = Convert.ToString(item.Name, CultureInfo.InvariantCulture)
                    });
                }

                var key = string.Join("|", path.Select(item => item.Field + "=" + item.Value));
                if (seen.Add(key))
                {
                    result.Add(path);
                }
            }

            return result;
        }

        private Dictionary<string, decimal> ReadSourceTotals(
            dynamic excelApplication,
            dynamic workbook,
            ReportSpecV1 specification)
        {
            dynamic range = FindSourceRange(workbook, specification.Source.WorkbookObjectName);
            var headers = ReadHeaderIndexes((object)range);
            var result = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
            foreach (var measure in specification.Measures)
            {
                if (!(measure.Expression is AggregateMeasureExpression aggregate) ||
                    aggregate.Function != AggregateFunction.Sum ||
                    !string.IsNullOrWhiteSpace(aggregate.PeriodSliceId))
                {
                    continue;
                }

                IReadOnlyList<string> sourceFields = sourceTotalLineageResolver.Resolve(
                    specification,
                    aggregate);
                if (sourceFields.Count == 0)
                {
                    continue;
                }

                decimal total = 0m;
                foreach (var field in sourceFields)
                {
                    if (!headers.TryGetValue(field, out var columnIndex))
                    {
                        throw new InvalidOperationException("A required source total column was not found.");
                    }

                    dynamic sourceColumn = range.Columns.Item(columnIndex);
                    var value = excelApplication.WorksheetFunction.Sum(sourceColumn);
                    total += Convert.ToDecimal(value, CultureInfo.InvariantCulture);
                }

                result[measure.Id] = total;
            }

            return result;
        }

        private static Dictionary<string, decimal> ReadCanonicalTotals(
            dynamic workbook,
            CanonicalLoadPlan canonical,
            ReportSpecV1 specification)
        {
            var result = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
            if (canonical.Backend != CanonicalBackend.Worksheet)
            {
                throw new InvalidOperationException(
                    "Data Model totals must be read from the independently refreshed canonical audit.");
            }

            dynamic? table = FindTable(workbook, canonical.TableOrConnectionName);
            if (table == null)
            {
                throw new InvalidOperationException("The canonical table was not found for total reconciliation.");
            }

            foreach (var measure in specification.Measures)
            {
                if (!(measure.Expression is AggregateMeasureExpression aggregate) ||
                    aggregate.Function != AggregateFunction.Sum)
                {
                    continue;
                }

                dynamic column = table.ListColumns.Item(aggregate.Field);
                dynamic body = column.DataBodyRange;
                var application = workbook.Application;
                var value = application.WorksheetFunction.Sum(body);
                result[measure.Id] = Convert.ToDecimal(value, CultureInfo.InvariantCulture);
            }

            return result;
        }

        private static void ReadPivotTotals(
            DenseReportBlockPlan block,
            dynamic nativePivot,
            PivotBuildResult pivotResult,
            IDictionary<ManagedBlockMeasureKey, decimal> pivotTotals,
            IDictionary<ManagedBlockMeasureKey, decimal> outputTotals)
        {
            foreach (var value in block.Pivot.Values)
            {
                if (!(value.Expression is AggregateMeasureExpression aggregate))
                {
                    continue;
                }

                var descriptor = pivotResult.DataFields.FirstOrDefault(field =>
                    string.Equals(field.MeasureId, value.MeasureId, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(field.SourceField, aggregate.Field, StringComparison.OrdinalIgnoreCase) &&
                    field.Function == aggregate.Function &&
                    field.Filters.Count == 0 &&
                    string.Equals(field.PeriodSliceId, aggregate.PeriodSliceId, StringComparison.OrdinalIgnoreCase));
                if (descriptor == null)
                {
                    continue;
                }

                dynamic totalCell = nativePivot.GetPivotData(descriptor.PivotCaption);
                var total = Convert.ToDecimal(totalCell.Value2, CultureInfo.InvariantCulture);
                var key = new ManagedBlockMeasureKey(block.BlockId, value.MeasureId);
                pivotTotals[key] = total;
                outputTotals[key] = total;
            }
        }

        private static dynamic FindSourceRange(dynamic workbook, string workbookObjectName)
        {
            dynamic? table = FindTable(workbook, workbookObjectName);
            if (table != null)
            {
                return table.Range;
            }

            try
            {
                return workbook.Names.Item(workbookObjectName).RefersToRange;
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException("The selected workbook source could not be reopened for validation.", exception);
            }
        }

        private static dynamic? FindTable(dynamic workbook, string tableName)
        {
            var sheetCount = Convert.ToInt32(workbook.Worksheets.Count, CultureInfo.InvariantCulture);
            for (var sheetIndex = 1; sheetIndex <= sheetCount; sheetIndex++)
            {
                dynamic sheet = workbook.Worksheets.Item(sheetIndex);
                try
                {
                    return sheet.ListObjects.Item(tableName);
                }
                catch (Exception)
                {
                }
            }

            return null;
        }

        private static Dictionary<string, int> ReadHeaderIndexes(dynamic range)
        {
            var result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var columnCount = Convert.ToInt32(range.Columns.Count, CultureInfo.InvariantCulture);
            for (var columnIndex = 1; columnIndex <= columnCount; columnIndex++)
            {
                var value = range.Cells[1, columnIndex].Value2;
                var header = Convert.ToString(value, CultureInfo.InvariantCulture);
                if (!string.IsNullOrWhiteSpace(header))
                {
                    result[header!] = columnIndex;
                }
            }

            return result;
        }

        private static long CountCanonicalRows(dynamic workbook, CanonicalLoadPlan canonical)
        {
            if (canonical.Backend != CanonicalBackend.Worksheet)
            {
                throw new InvalidOperationException(
                    "Data Model row counts must be read from the managed ModelTable record count.");
            }

            var sheetCount = Convert.ToInt32(workbook.Worksheets.Count, CultureInfo.InvariantCulture);
            for (var sheetIndex = 1; sheetIndex <= sheetCount; sheetIndex++)
            {
                dynamic sheet = workbook.Worksheets.Item(sheetIndex);
                try
                {
                    dynamic table = sheet.ListObjects.Item(canonical.TableOrConnectionName);
                    return Convert.ToInt64(table.ListRows.Count, CultureInfo.InvariantCulture);
                }
                catch (Exception)
                {
                }
            }

            throw new InvalidOperationException("The managed canonical table could not be found after refresh.");
        }

        internal static long CountDataModelRows(dynamic workbook, string managedConnectionName)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            if (string.IsNullOrWhiteSpace(managedConnectionName))
            {
                throw new ArgumentException(
                    "A managed Data Model connection name is required.",
                    nameof(managedConnectionName));
            }

            long? result = null;
            dynamic modelTables = workbook.Model.ModelTables;
            int tableCount = Convert.ToInt32(modelTables.Count, CultureInfo.InvariantCulture);
            for (var tableIndex = 1; tableIndex <= tableCount; tableIndex++)
            {
                dynamic modelTable = modelTables.Item(tableIndex);
                dynamic? sourceConnection;
                try
                {
                    sourceConnection = modelTable.SourceWorkbookConnection;
                }
                catch (Exception)
                {
                    continue;
                }

                if (sourceConnection == null)
                {
                    continue;
                }

                string? sourceName = Convert.ToString(
                    sourceConnection.Name,
                    CultureInfo.InvariantCulture);
                if (!string.Equals(
                        sourceName,
                        managedConnectionName,
                        StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (result.HasValue)
                {
                    throw new InvalidOperationException(
                        "More than one Data Model table is linked to the managed canonical connection.");
                }

                long recordCount = Convert.ToInt64(
                    modelTable.RecordCount,
                    CultureInfo.InvariantCulture);
                if (recordCount < 0L)
                {
                    throw new InvalidOperationException(
                        "The managed Data Model table returned an invalid row count.");
                }

                result = recordCount;
            }

            return result ?? throw new InvalidOperationException(
                "The managed canonical connection is not linked to exactly one Data Model table.");
        }
    }
}
