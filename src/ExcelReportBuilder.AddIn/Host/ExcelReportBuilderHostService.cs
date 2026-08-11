using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Dynamic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;
using ExcelReportBuilder.AddIn.Activity;
using ExcelReportBuilder.Agent.Configuration;
using ExcelReportBuilder.Agent.Execution;
using ExcelReportBuilder.Agent.Models;
using ExcelReportBuilder.Agent.OpenAI;
using ExcelReportBuilder.Agent.Tools;
using ExcelReportBuilder.Core.Periods;
using ExcelReportBuilder.Core.Planning;
using ExcelReportBuilder.Core.Profiling;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Core.Validation;
using ExcelReportBuilder.Core.Transforms;
using ExcelReportBuilder.Excel.Execution;
using ExcelReportBuilder.Excel.Ownership;
using ExcelReportBuilder.Excel.Persistence;
using ExcelReportBuilder.Excel.Publishing;
using ExcelReportBuilder.Excel.Source;
using ExcelReportBuilder.Excel.Validation;

namespace ExcelReportBuilder.AddIn.Host
{
    /// <summary>
    /// Excel-owned implementation of the task-pane boundary. Reads are made
    /// against the chosen range without selecting or activating it. Workbook
    /// writes are delegated only to the managed draft and publish services.
    /// </summary>
    public sealed class ExcelReportBuilderHostService : IReportBuilderHostService, IDisposable
    {
        private const int AgentSampleRowLimit = 50;
        private const string PlaceholderSourceName = "SelectedData";

        private readonly object _excelApplication;
        private readonly Dispatcher _excelDispatcher;
        private readonly SourceSelectionInspector _selectionInspector = new SourceSelectionInspector();
        private readonly ManagedSourceNameService _sourceNameService = new ManagedSourceNameService();
        private readonly ExcelReportExecutor _executor = new ExcelReportExecutor();
        private readonly ManagedPublishService _publishService = new ManagedPublishService();
        private readonly ManagedOwnershipGuard _ownershipGuard = new ManagedOwnershipGuard();
        private readonly WorkbookIdentityStore _workbookIdentityStore = new WorkbookIdentityStore();
        private readonly WorkbookSpecStore _workbookSpecStore = new WorkbookSpecStore();
        private readonly ReportSpecTranslator _translator = new ReportSpecTranslator();
        private readonly ProtectedAgentSettingsStore _settingsStore = new ProtectedAgentSettingsStore();
        private readonly WindowsCurrentUserSecretProtector _secretProtector = new WindowsCurrentUserSecretProtector();
        private readonly ManualResetEventSlim _pauseGate = new ManualResetEventSlim(true);
        private readonly object _operationGate = new object();
        private string _reportId = "report_" + Guid.NewGuid().ToString("N");
        private string _workbookId = string.Empty;

        private CancellationTokenSource? _operationCancellation;
        private SelectedSourceState? _selectedSource;
        private ReportSpecV1? _lastSpecification;
        private ExcelBuildResult? _lastBuild;
        private object? _lastBuildWorkbook;
        private bool _lastChecksPassed;
        private bool _disposed;

        public ExcelReportBuilderHostService(object excelApplication)
        {
            _excelApplication = excelApplication ?? throw new ArgumentNullException(nameof(excelApplication));
            _excelDispatcher = Dispatcher.CurrentDispatcher;
            SavedEndpointSettings = LoadSavedEndpointSettings();
        }

        public event EventHandler<HostActivityEventArgs>? ActivityReported;

        public bool IsSynthetic => false;

        public SavedEndpointSettingsSnapshot? SavedEndpointSettings { get; }

        public Task<SourceSnapshot> SelectCurrentDataAsync(CancellationToken cancellationToken)
        {
            return RunOperationAsync(
                cancellationToken,
                token => InvokeExcelAsync(
                    () =>
                    {
                        Report(
                            ActivityStage.Inspecting,
                            ActivityKind.Progress,
                            "Inspecting the current Excel selection.",
                            "The selected cells are read without changing the selection.");
                        token.ThrowIfCancellationRequested();
                        SelectedSourceState source = ReadCurrentSelection();
                        SavedSetupSelection savedSetup = FindSavedSetup(source);
                        string? existingWorkbookId = _workbookIdentityStore.Load(source.Workbook);
                        _workbookId = existingWorkbookId ??
                            _workbookIdentityStore.GetOrCreate(source.Workbook);
                        if (existingWorkbookId == null)
                        {
                            Report(
                                ActivityStage.Inspecting,
                                ActivityKind.Check,
                                "Anonymous workbook identity created.",
                                "Only a random managed identifier was added to the open workbook. No path, workbook name, sheet name, or cell value was stored, and the workbook was not saved.");
                        }

                        if (savedSetup.Specification != null)
                        {
                            _reportId = savedSetup.Specification.Id;
                            _lastSpecification = savedSetup.Specification;
                            Report(
                                ActivityStage.Inspecting,
                                ActivityKind.Check,
                                "Compatible saved report setup restored.",
                                "The selected workbook object and path-free source fingerprint both match. The workbook remains unsaved.");
                        }
                        else
                        {
                            _reportId = "report_" + Guid.NewGuid().ToString("N");
                            _lastSpecification = null;
                            if (!string.IsNullOrWhiteSpace(savedSetup.Status))
                            {
                                Report(
                                    ActivityStage.Inspecting,
                                    ActivityKind.Check,
                                    "Saved report setup was not restored.",
                                    savedSetup.Status);
                            }
                        }

                        _selectedSource = source;
                        _lastBuild = null;
                        _lastBuildWorkbook = null;
                        _lastChecksPassed = false;
                        Report(
                            ActivityStage.Inspecting,
                            ActivityKind.Result,
                            "Workbook Data is ready.",
                            source.SelectionSnapshot.RowCount.ToString("N0", CultureInfo.CurrentCulture) +
                            " rows and " + source.SelectionSnapshot.ColumnCount.ToString(CultureInfo.CurrentCulture) +
                            " columns were profiled.");
                        return ToSourceSnapshot(
                            source,
                            savedSetup.UiSnapshot,
                            savedSetup.Status);
                    },
                    token));
        }

        public Task ConfirmPeriodMappingAsync(
            PeriodMappingSnapshot periodMapping,
            CancellationToken cancellationToken)
        {
            return RunOperationAsync(
                cancellationToken,
                token => InvokeExcelAsync(
                    () =>
                    {
                        SelectedSourceState source = ReadSelectedSource();
                        token.ThrowIfCancellationRequested();
                        PeriodMappingSpec? resolved = _translator.ResolvePeriodMapping(
                            periodMapping,
                            source.Profile);
                        Report(
                            ActivityStage.Planning,
                            ActivityKind.Check,
                            "Period layout confirmed against the chosen Data.",
                            resolved == null
                                ? "No period columns will be normalized."
                                : resolved.Kind.ToString());
                        return true;
                    },
                    token));
        }

        public Task<WideHeaderMappingPreview> PreviewWideHeaderMappingAsync(
            string headerPattern,
            int? reportingYear,
            CancellationToken cancellationToken)
        {
            return RunOperationAsync(
                cancellationToken,
                token => InvokeExcelAsync(
                    () =>
                    {
                        Report(
                            ActivityStage.Inspecting,
                            ActivityKind.Progress,
                            "Inspecting wide period headers in the current selection.",
                            "The selection and source values remain unchanged.");
                        SelectedSourceState current = ReadCurrentSelection();
                        EnsureSameChosenSource(current);
                        token.ThrowIfCancellationRequested();
                        PeriodDetectionResult detection = PeriodDetector.Detect(
                            current.Profile,
                            reportingYear);
                        if (detection.Kind != PeriodLayoutKind.MonthHeaders &&
                            detection.Kind != PeriodLayoutKind.MetricMonthHeaders)
                        {
                            throw new InvalidOperationException(
                                "The current selection does not contain recognizable wide period headers.");
                        }

                        PeriodDetectionIssue? blockingIssue = detection.Issues.FirstOrDefault(issue =>
                            issue.Severity == PeriodDetectionSeverity.Error &&
                            issue.Code != PeriodDetectionIssueCode.MissingReportingYear);
                        if (blockingIssue != null)
                        {
                            throw new InvalidOperationException(blockingIssue.Message);
                        }

                        _selectedSource = current;
                        IReadOnlyList<WideHeaderMappingRowSnapshot> mappings = detection.HeaderMatches
                            .Select(match => new WideHeaderMappingRowSnapshot(
                                match.SourceColumn,
                                FormatPeriod(match, reportingYear),
                                string.IsNullOrWhiteSpace(match.Metric) ? "Value" : match.Metric!,
                                match.Year.HasValue ? 0.99d : reportingYear.HasValue ? 0.96d : 0.9d))
                            .ToList();
                        long projectedRows = checked(
                            current.SelectionSnapshot.RowCount * detection.HeaderMatches.Count);
                        if (detection.RequiresReportingYear)
                        {
                            return new WideHeaderMappingPreview(
                                mappings,
                                projectedRows,
                                Array.Empty<NormalizedSampleRowSnapshot>(),
                                TotalPreservationState.NotChecked,
                                "A reporting year is required before sample rows and totals can be checked.",
                                true);
                        }

                        PeriodMappingSpec mapping = detection.ToPeriodMapping();
                        IReadOnlyList<NormalizedPeriodValue> normalized = WidePeriodNormalizer.Normalize(
                            ToRowDictionaries(current),
                            mapping);
                        var sampleRows = normalized.Take(12)
                            .Select((value, index) => new NormalizedSampleRowSnapshot(
                                "Row " + (index / mapping.Columns.Count + 2).ToString(CultureInfo.CurrentCulture),
                                value.Period.ToString("yyyy-MM", CultureInfo.InvariantCulture),
                                string.IsNullOrWhiteSpace(value.Metric) ? "Value" : value.Metric!,
                                FormatValue(value.Value)))
                            .ToList();
                        decimal sampleSourceTotal = SumMappedValues(current, mapping);
                        decimal sampleNormalizedTotal = normalized.Sum(value => ToDecimalOrZero(value.Value));
                        bool totalsMatch = sampleSourceTotal == sampleNormalizedTotal;
                        int previewRowCount = current.Rows.Count;
                        string totalDetail = totalsMatch
                            ? "The " + previewRowCount.ToString("N0", CultureInfo.CurrentCulture) +
                              "-row preview preserves mapped values at " +
                              sampleSourceTotal.ToString("N2", CultureInfo.CurrentCulture) +
                              ". Full-source reconciliation runs after the managed draft is built."
                            : "The preview source total " +
                              sampleSourceTotal.ToString("N2", CultureInfo.CurrentCulture) +
                              " differs from its normalized total " +
                              sampleNormalizedTotal.ToString("N2", CultureInfo.CurrentCulture) + ".";
                        Report(
                            ActivityStage.Planning,
                            totalsMatch ? ActivityKind.Check : ActivityKind.Error,
                            totalsMatch
                                ? "Wide period preview values are preserved."
                                : "Wide period preview values do not match.",
                            totalDetail);
                        return new WideHeaderMappingPreview(
                            mappings,
                            projectedRows,
                            sampleRows,
                            totalsMatch ? TotalPreservationState.Pass : TotalPreservationState.Fail,
                            totalDetail,
                            false);
                    },
                    token));
        }

        public Task<BuildDraftResult> BuildManagedDraftAsync(
            ReportSpecificationSnapshot specification,
            CancellationToken cancellationToken)
        {
            return RunOperationAsync(
                cancellationToken,
                token => InvokeExcelAsync(
                    () =>
                    {
                        SelectedSourceState source = ReadSelectedSource();
                        WorkbookSourceKind sourceKind = DetermineSourceKind(source.Selection);
                        ReportSpecV1 coreSpecification;
                        try
                        {
                            coreSpecification = _translator.FromUi(
                                specification,
                                source.Profile,
                                PlaceholderSourceName,
                                sourceKind,
                                _reportId,
                                exclusion => CountObservedTotalRows(source, exclusion));
                        }
                        catch (InvalidOperationException exception)
                        {
                            throw new ReportSetupValidationException(exception.Message, exception);
                        }
                        ExcelBuildResult result = BuildCoreSpecification(
                            coreSpecification,
                            source,
                            token);
                        return new BuildDraftResult(
                            result.DraftWorksheets.FirstOrDefault() ?? "Managed draft",
                            result.NormalizedRows > int.MaxValue
                                ? int.MaxValue
                                : Convert.ToInt32(result.NormalizedRows, CultureInfo.InvariantCulture));
                    },
                    token));
        }

        public Task<ChatRunResult> RunChatAsync(
            string request,
            ReportSpecificationSnapshot specification,
            ModelEndpointSettingsSnapshot endpointSettings,
            SecureString? apiKey,
            CancellationToken cancellationToken)
        {
            return RunOperationAsync(
                cancellationToken,
                async token =>
                {
                    SelectedSourceState source = await InvokeExcelAsync(
                        ReadSelectedSource,
                        token).ConfigureAwait(false);
                    AgentEndpointSettings endpoint = await MaterializeEndpointAsync(
                        endpointSettings,
                        apiKey,
                        token).ConfigureAwait(false);
                    Uri endpointUri = AgentEndpointPolicy.Validate(endpoint);
                    if (!endpointUri.IsLoopback && !endpoint.AllowRemoteWorkbookData)
                    {
                        throw new AgentEndpointPolicyException(
                            "A remote model endpoint can receive workbook column names and up to 50 sample rows only after explicit consent.");
                    }

                    var context = new ChatToolContext(
                        specification.PeriodMapping,
                        ReportSpecTranslator.ParseOutputMode(specification.OutputStyle));
                    AgentJobRequest job = CreateAgentJob(
                        request,
                        specification,
                        source,
                        endpoint);
                    if (!WorkbookAgentJobLease.TryAcquire(_workbookId, out WorkbookAgentJobLease? lease) ||
                        lease == null)
                    {
                        throw new InvalidOperationException(
                            "Another AI report-building job is already active for this workbook.");
                    }

                    Report(
                        ActivityStage.Planning,
                        ActivityKind.Progress,
                        "Starting the guarded AI worker.",
                        "Only bounded Data and report setup snapshots are sent to the worker.");

                    using (lease)
                    using (var worker = new AgentWorkerClient(
                               ReportWorkerProgressAsync,
                               ReportWorkerCheckpointAsync,
                               (toolRequest, toolToken) => InvokeExcelAsync(
                                   () => ExecuteHostTool(toolRequest, context, toolToken),
                                   toolToken)))
                    {
                        AgentRunResult result = await worker.RunAsync(job, token).ConfigureAwait(false);
                        await SaveEndpointAsync(endpoint, token).ConfigureAwait(false);
                        if (context.ValidatedSpecification == null ||
                            context.BuildResult == null ||
                            context.CheckResults.Count == 0 ||
                            !context.ChecksPassed ||
                            context.FinalChanges.Count == 0)
                        {
                            throw new InvalidOperationException(
                                "The guarded workflow ended without a complete applied draft, checks, and change summary.");
                        }

                        ReportSpecificationSnapshot applied =
                            _translator.ToAppliedAgentSnapshot(context.ValidatedSpecification);
                        string draftName = context.BuildResult.DraftWorksheets.FirstOrDefault()
                            ?? "Managed draft";
                        int outputRows = context.BuildResult.NormalizedRows > int.MaxValue
                            ? int.MaxValue
                            : Convert.ToInt32(
                                context.BuildResult.NormalizedRows,
                                CultureInfo.InvariantCulture);
                        string response = "Built and independently checked " +
                            context.ValidatedSpecification.Blocks.Count.ToString(CultureInfo.CurrentCulture) +
                            " managed report block" +
                            (context.ValidatedSpecification.Blocks.Count == 1 ? string.Empty : "s") +
                            ". The draft remains unpublished and the workbook was not saved. Repair cycles used: " +
                            result.RepairCyclesUsed.ToString(CultureInfo.CurrentCulture) + ".";
                        return new ChatRunResult(
                            response,
                            applied,
                            draftName,
                            outputRows,
                            context.CheckResults,
                            context.FinalChanges,
                            allChecksPassed: true,
                            published: false);
                    }
                });
        }

        public Task<IReadOnlyList<string>> DiscoverModelsAsync(
            ModelEndpointSettingsSnapshot endpointSettings,
            SecureString? apiKey,
            CancellationToken cancellationToken)
        {
            return RunOperationAsync(
                cancellationToken,
                async token =>
                {
                    AgentEndpointSettings endpoint = await MaterializeEndpointAsync(
                        endpointSettings,
                        apiKey,
                        token).ConfigureAwait(false);
                    Report(
                        ActivityStage.Ready,
                        ActivityKind.Progress,
                        "Requesting the endpoint model list.",
                        "No workbook Data is included in model discovery.");
                    using (var client = new OpenAiCompatibleClient())
                    {
                        ModelDiscoveryResult result = await AwaitWithHeartbeatAsync(
                            client.DiscoverModelsAsync(endpoint, token),
                            "Still waiting for the endpoint model list.",
                            token).ConfigureAwait(false);
                        await SaveEndpointAsync(endpoint, token).ConfigureAwait(false);
                        return (IReadOnlyList<string>)result.ModelIds;
                    }
                });
        }

        public Task<EndpointCheckResult> CheckEndpointAsync(
            ModelEndpointSettingsSnapshot endpointSettings,
            SecureString? apiKey,
            CancellationToken cancellationToken)
        {
            return RunOperationAsync(
                cancellationToken,
                async token =>
                {
                    AgentEndpointSettings endpoint = await MaterializeEndpointAsync(
                        endpointSettings,
                        apiKey,
                        token).ConfigureAwait(false);
                    Report(
                        ActivityStage.Ready,
                        ActivityKind.Check,
                        "Checking model discovery, structured output, and tool calling.",
                        "The capability check uses synthetic Data only.");
                    using (var client = new OpenAiCompatibleClient())
                    {
                        EndpointProbeResult result = await AwaitWithHeartbeatAsync(
                            client.CheckToolCallingAsync(endpoint, token),
                            "Still waiting for the synthetic endpoint capability checks.",
                            token).ConfigureAwait(false);
                        await SaveEndpointAsync(endpoint, token).ConfigureAwait(false);
                        return new EndpointCheckResult(
                            result.ToolCallingAvailable &&
                            result.StructuredOutputAvailable,
                            result.Summary);
                    }
                });
        }

        public Task<IReadOnlyList<HostCheckResult>> RunChecksAsync(
            CancellationToken cancellationToken)
        {
            return RunOperationAsync(
                cancellationToken,
                token => InvokeExcelAsync(
                    () => RunChecksCore(token),
                    token));
        }

        public Task<PublishResult> PublishManagedDraftAsync(
            CancellationToken cancellationToken)
        {
            return RunOperationAsync(
                cancellationToken,
                token => InvokeExcelAsync(
                    () =>
                    {
                        token.ThrowIfCancellationRequested();
                        if (_lastSpecification == null || _lastBuild == null ||
                            _lastBuildWorkbook == null)
                        {
                            throw new InvalidOperationException(
                                "A checked managed draft is required before publishing.");
                        }

                        object activeWorkbook = GetActiveWorkbook();
                        if (!AreSameComObject(activeWorkbook, _lastBuildWorkbook))
                        {
                            throw new InvalidOperationException(
                                "Activate the workbook that contains the checked managed draft.");
                        }

                        Report(
                            ActivityStage.Checking,
                            ActivityKind.Progress,
                            "Rebuilding and rechecking the managed draft before publish.",
                            "Current source values and managed formulas are checked again. The workbook will not be saved.");
                        IReadOnlyList<HostCheckResult> liveChecks = RunChecksCore(token);
                        if (liveChecks.Any(check => !check.Passed))
                        {
                            throw new InvalidOperationException(
                                "The current managed draft failed live checks and cannot be published.");
                        }

                        dynamic workbook = activeWorkbook;
                        Report(
                            ActivityStage.Checking,
                            ActivityKind.Check,
                            "Rechecking every managed output worksheet before publish.",
                            "The workbook will not be saved by the add-in.");
                        var messages = new List<string>();
                        var logicalOutputs = _lastSpecification.Blocks
                            .GroupBy(
                                block => ManagedOutputIdentity.LogicalKey(block.WorksheetName),
                                StringComparer.Ordinal)
                            .Select(group => group.First().WorksheetName)
                            .ToList();
                        var publishRequests = new List<ManagedPublishRequest>();
                        for (var outputIndex = 0; outputIndex < logicalOutputs.Count; outputIndex++)
                        {
                            var logicalWorksheetName = logicalOutputs[outputIndex];
                            Report(
                                ActivityStage.Checking,
                                ActivityKind.Check,
                                "Validating managed output " +
                                (outputIndex + 1).ToString(CultureInfo.InvariantCulture) +
                                " of " + logicalOutputs.Count.ToString(CultureInfo.InvariantCulture) + ".",
                                logicalWorksheetName);
                            var draftIdentity = ManagedOutputIdentity.Draft(
                                _lastSpecification.Id,
                                logicalWorksheetName);
                            dynamic draft = FindExactlyOneOwnedWorksheet(workbook, draftIdentity);
                            var publishedIdentity = ManagedOutputIdentity.Published(
                                _lastSpecification.Id,
                                logicalWorksheetName);
                            var rollbackIdentity = ManagedOutputIdentity.Rollback(
                                _lastSpecification.Id,
                                logicalWorksheetName);
                            publishRequests.Add(new ManagedPublishRequest
                            {
                                DraftWorksheet = draft,
                                DraftIdentity = draftIdentity,
                                PublishedIdentity = publishedIdentity,
                                RollbackIdentity = rollbackIdentity,
                                FinalWorksheetName = PublishedWorksheetName(
                                    logicalWorksheetName,
                                    publishedIdentity)
                            });
                        }

                        IReadOnlyList<ExcelReportBuilder.Excel.Publishing.PublishResult> publishedResults =
                            _publishService.PublishAll(
                                _excelApplication,
                                (object)workbook,
                                publishRequests,
                                userConfirmed: true,
                                beforePublish: (Action<int, int, string>)((current, total, worksheetName) =>
                                    Report(
                                        ActivityStage.Rendering,
                                        ActivityKind.Progress,
                                        "Publishing managed output " +
                                        current.ToString(CultureInfo.InvariantCulture) +
                                        " of " + total.ToString(CultureInfo.InvariantCulture) + ".",
                                        worksheetName)));
                        foreach (var result in publishedResults)
                        {
                            var outputMessage = "Published worksheet " + result.PublishedWorksheetName + ".";
                            if (!string.IsNullOrWhiteSpace(result.RollbackWorksheetName))
                            {
                                outputMessage += " Previous output retained as " + result.RollbackWorksheetName + ".";
                            }

                            messages.Add(outputMessage);
                        }

                        string message = string.Join(" ", messages) + " The workbook remains unsaved.";
                        Report(
                            ActivityStage.Complete,
                            ActivityKind.Result,
                            "Managed draft published.",
                            message);
                        return new PublishResult(message);
                    },
                    token));
        }

        private static string PublishedWorksheetName(
            string logicalWorksheetName,
            ManagedObjectIdentity publishedIdentity)
        {
            var stableSuffix = publishedIdentity.ObjectId.StartsWith(
                "output_",
                StringComparison.Ordinal)
                ? publishedIdentity.ObjectId.Substring("output_".Length)
                : publishedIdentity.ObjectId;
            return ManagedName.Worksheet(logicalWorksheetName, stableSuffix);
        }

        private object FindExactlyOneOwnedWorksheet(
            dynamic workbook,
            ManagedObjectIdentity identity)
        {
            object? match = null;
            var count = Convert.ToInt32(workbook.Worksheets.Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                dynamic worksheet = workbook.Worksheets.Item(index);
                if (!_ownershipGuard.IsOwned(worksheet, identity))
                {
                    continue;
                }

                if (match != null)
                {
                    throw new InvalidOperationException(
                        "More than one worksheet carries the requested managed output identity.");
                }

                match = worksheet;
            }

            return match ?? throw new InvalidOperationException(
                "A checked managed output worksheet is missing and cannot be published.");
        }

        public void RequestPause()
        {
            _pauseGate.Reset();
        }

        public void RequestResume()
        {
            _pauseGate.Set();
        }

        public void RequestCancel()
        {
            lock (_operationGate)
            {
                _operationCancellation?.Cancel();
            }

            _pauseGate.Set();
        }

        public void Dispose()
        {
            lock (_operationGate)
            {
                if (_disposed)
                {
                    return;
                }

                _disposed = true;
                _operationCancellation?.Cancel();
                _operationCancellation?.Dispose();
                _operationCancellation = null;
            }

            _pauseGate.Set();
            _pauseGate.Dispose();
            ActivityReported = null;
            _selectedSource = null;
            _lastBuildWorkbook = null;
        }

        private ExcelBuildResult BuildCoreSpecification(
            ReportSpecV1 specification,
            SelectedSourceState source,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            Report(
                ActivityStage.Inspecting,
                ActivityKind.Progress,
                "Rechecking the chosen Data before the managed build.",
                "Headers and source shape are verified without changing the chosen range.");
            ReportBuildPlanner.Create(specification, source.Profile, null);
            cancellationToken.ThrowIfCancellationRequested();
            dynamic workbook = source.Workbook;
            dynamic selection = source.Selection;
            string workbookObjectName = _sourceNameService.EnsureWorkbookObject(
                workbook,
                selection,
                specification.Id,
                "source");
            specification.Source.Kind = DetermineSourceKind(source.Selection);
            specification.Source.WorkbookObjectName = workbookObjectName;
            ReportBuildPlan plan = ReportBuildPlanner.Create(
                specification,
                source.Profile,
                null);
            Report(
                ActivityStage.Normalizing,
                ActivityKind.Progress,
                "Preparing normalized managed rows.",
                plan.Source.ProjectedRows.ToString("N0", CultureInfo.CurrentCulture) +
                " normalized rows are planned. The workbook will not be saved.");

            var progress = new HostExcelProgressSink(this, cancellationToken);
            try
            {
                ExcelBuildResult result = _executor.BuildManagedDraft(
                    _excelApplication,
                    specification,
                    plan,
                    progress,
                    cancellationToken);
                RememberBuild(specification, source.Workbook, result);
                return result;
            }
            catch (ExcelBuildValidationException exception)
            {
                RememberBuild(specification, source.Workbook, exception.Result);
                foreach (CheckResult failure in exception.Result.Checks.Where(check =>
                             check.Outcome == CheckOutcome.Failed))
                {
                    Report(
                        ActivityStage.Checking,
                        ActivityKind.Error,
                        failure.CheckId + ": Fail",
                        failure.Message);
                }

                throw;
            }
        }

        private void RememberBuild(
            ReportSpecV1 specification,
            object workbook,
            ExcelBuildResult result)
        {
            _lastSpecification = specification;
            _lastBuild = result;
            _lastBuildWorkbook = workbook;
            _lastChecksPassed = false;
        }

        private IReadOnlyList<HostCheckResult> RunChecksCore(CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (_lastSpecification == null || _lastBuild == null || _lastBuildWorkbook == null)
            {
                throw new InvalidOperationException("Build a managed draft before running checks.");
            }

            object activeWorkbook = GetActiveWorkbook();
            if (!AreSameComObject(activeWorkbook, _lastBuildWorkbook))
            {
                throw new InvalidOperationException(
                    "Activate the workbook that contains the managed draft.");
            }

            Report(
                ActivityStage.Checking,
                ActivityKind.Progress,
                "Rebuilding the managed draft against current Data before checks.",
                "This prevents stale source values or edited report formulas from passing an old result.");
            SelectedSourceState currentSource = ReadSelectedSource();
            try
            {
                BuildCoreSpecification(_lastSpecification, currentSource, cancellationToken);
            }
            catch (ExcelBuildValidationException)
            {
                // BuildCoreSpecification retains the failed managed draft and
                // its bounded check details for display and agent repair.
            }

            var results = new List<HostCheckResult>();
            foreach (CheckResult check in _lastBuild.Checks)
            {
                cancellationToken.ThrowIfCancellationRequested();
                bool passed = check.Outcome != CheckOutcome.Failed;
                results.Add(new HostCheckResult(check.CheckId, passed, check.Message));
                Report(
                    ActivityStage.Checking,
                    passed ? ActivityKind.Check : ActivityKind.Error,
                    check.CheckId + ": " + (passed ? "Pass" : "Fail"),
                    check.Message);
            }

            ReportBlockSpec block = _lastSpecification.Blocks[0];
            string draftName = _lastBuild.DraftWorksheets.First();
            dynamic workbook = activeWorkbook;
            dynamic draft = workbook.Worksheets.Item(draftName);
            var identity = new ManagedObjectIdentity(
                _lastSpecification.Id,
                block.OwnershipId + "_draft",
                ManagedObjectKind.DraftWorksheet);
            bool owned = _ownershipGuard.IsOwned(draft, identity);
            results.Add(new HostCheckResult(
                "managed-ownership",
                owned,
                owned
                    ? "The draft carries the expected managed ownership marker."
                    : "The draft ownership marker does not match the report setup."));
            Report(
                ActivityStage.Checking,
                owned ? ActivityKind.Check : ActivityKind.Error,
                "Managed ownership: " + (owned ? "Pass" : "Fail"),
                results.Last().Detail);
            if (_lastSpecification.PeriodMapping != null)
            {
                results.Add(new HostCheckResult(
                    "period-coverage",
                    true,
                    "Every planned period mapping was rebuilt successfully against the current managed source."));
                Report(
                    ActivityStage.Checking,
                    ActivityKind.Check,
                    "Period coverage: Pass",
                    results.Last().Detail);
            }

            _lastChecksPassed = results.Count > 0 && results.All(result => result.Passed);
            return results;
        }

        private HostToolResultRequest ExecuteHostTool(
            HostToolRequestEvent request,
            ChatToolContext context,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!AgentToolCatalog.IsAllowed(request.ToolName))
            {
                return ToolFailure(request, "host_tool_not_allowed", "The requested host tool is not allowlisted.");
            }

            Report(
                MapToolStage(request.ToolName),
                ActivityKind.Progress,
                "Running allowlisted host tool " + request.ToolName + ".",
                "Excel COM remains in the add-in process.");
            try
            {
                switch (request.ToolName)
                {
                    case AgentToolNames.ProposePeriodMapping:
                        return AcceptPeriodProposal(request, context);
                    case AgentToolNames.ProposeTransforms:
                        return AcceptTransformProposal(request, context);
                    case AgentToolNames.ProposeReportSpec:
                        context.ProposalToolCallId = request.ToolCallId;
                        context.ProposalArgumentsJson = request.ArgumentsJson;
                        context.ValidatedSpecification = null;
                        context.ValidatedSpecificationId = null;
                        context.ManagedDraftId = null;
                        context.BuildResult = null;
                        context.ChecksPassed = false;
                        context.CheckResults = Array.Empty<HostCheckResult>();
                        context.FinalChanges = Array.Empty<ChatChangeSnapshot>();
                        return ToolSuccess(
                            request,
                            "report_spec_proposed",
                            new { proposalToolCallId = request.ToolCallId });
                    case AgentToolNames.ValidateSpec:
                        return ValidateAgentSpecification(request, context);
                    case AgentToolNames.RequestManagedDraftBuild:
                        return BuildAgentDraft(request, context, cancellationToken);
                    case AgentToolNames.RunChecks:
                        return RunAgentChecks(request, context, cancellationToken);
                    case AgentToolNames.FinalChangeSummary:
                        return AcceptFinalSummary(request, context);
                    default:
                        return ToolFailure(request, "host_tool_not_allowed", "The requested host tool is not allowlisted.");
                }
            }
            catch (InvalidReportSpecException exception)
            {
                ValidationIssue? issue = exception.Validation.Issues.FirstOrDefault(value =>
                    value.Severity == ValidationSeverity.Error);
                return ToolFailure(
                    request,
                    "spec_validation_failed",
                    issue?.Message ?? "The report setup did not pass deterministic validation.",
                    issue?.Code ?? "spec_invalid");
            }
            catch (JsonException)
            {
                return ToolFailure(
                    request,
                    "tool_arguments_invalid",
                    "The host tool arguments are not valid bounded JSON.");
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception)
            {
                return ToolFailure(
                    request,
                    "host_tool_rejected",
                    "The deterministic host rejected the requested report workflow step.");
            }
        }

        private HostToolResultRequest AcceptPeriodProposal(
            HostToolRequestEvent request,
            ChatToolContext context)
        {
            using (var document = JsonDocument.Parse(request.ArgumentsJson))
            {
                JsonElement root = document.RootElement;
                string mode = root.GetProperty("mode").GetString() ?? string.Empty;
                if (string.Equals(mode, "unresolved", StringComparison.Ordinal))
                {
                    return ToolFailure(
                        request,
                        "period_mapping_unresolved",
                        "Confirm an explicit reporting year or date column before building the managed draft.");
                }

                int? reportingYear = root.GetProperty("reportingYear").ValueKind == JsonValueKind.Number
                    ? root.GetProperty("reportingYear").GetInt32()
                    : (int?)null;
                context.PeriodMapping = string.Equals(mode, "dateColumn", StringComparison.Ordinal)
                    ? new PeriodMappingSnapshot(
                        "Date column",
                        root.GetProperty("periodField").GetString() ?? string.Empty,
                        string.Empty,
                        reportingYear,
                        Array.Empty<WideHeaderMappingRowSnapshot>())
                    : new PeriodMappingSnapshot(
                        "Wide period headers",
                        string.Empty,
                        "Detected month headers",
                        reportingYear,
                        Array.Empty<WideHeaderMappingRowSnapshot>());
                SelectedSourceState source = ReadSelectedSource();
                _translator.ResolvePeriodMapping(context.PeriodMapping, source.Profile);
                return ToolSuccess(request, "period_mapping_accepted", new { accepted = true });
            }
        }

        private static HostToolResultRequest AcceptTransformProposal(
            HostToolRequestEvent request,
            ChatToolContext context)
        {
            using (var document = JsonDocument.Parse(request.ArgumentsJson))
            {
                var accepted = new List<TransformStep>();
                var transformIndex = 0;
                foreach (var proposed in document.RootElement.GetProperty("transforms").EnumerateArray())
                {
                    transformIndex++;
                    string kind = proposed.GetProperty("kind").GetString() ?? string.Empty;
                    string sourceField = proposed.GetProperty("sourceField").GetString() ?? string.Empty;
                    string outputField = proposed.GetProperty("outputField").GetString() ?? string.Empty;
                    string id = "agent_transform_" + transformIndex.ToString(CultureInfo.InvariantCulture);
                    bool appendRename = !string.Equals(
                        sourceField,
                        outputField,
                        StringComparison.OrdinalIgnoreCase);
                    switch (kind)
                    {
                        case "rename":
                            accepted.Add(new RenameColumnTransform
                            {
                                Id = id,
                                From = sourceField,
                                To = outputField
                            });
                            appendRename = false;
                            break;
                        case "trimText":
                            accepted.Add(new TrimTextTransform
                            {
                                Id = id,
                                Columns = new List<string> { sourceField }
                            });
                            break;
                        case "convertNumber":
                            accepted.Add(new ChangeColumnTypeTransform
                            {
                                Id = id,
                                Column = sourceField,
                                DataType = ColumnDataType.DecimalNumber
                            });
                            break;
                        case "convertDate":
                            accepted.Add(new ChangeColumnTypeTransform
                            {
                                Id = id,
                                Column = sourceField,
                                DataType = ColumnDataType.Date
                            });
                            break;
                        case "replaceBlank":
                        case "normalizeBlanks":
                            accepted.Add(new NormalizeBlanksTransform
                            {
                                Id = id,
                                Columns = new List<string> { sourceField },
                                Replacement = ScalarValue.Null(),
                                TreatWhitespaceAsBlank = true
                            });
                            break;
                        case "normalizeErrors":
                            accepted.Add(new NormalizeErrorsTransform
                            {
                                Id = id,
                                Columns = new List<string> { sourceField },
                                Replacement = ScalarValue.Null()
                            });
                            break;
                        case "fillDown":
                            accepted.Add(new FillDownTransform
                            {
                                Id = id,
                                Columns = new List<string> { sourceField }
                            });
                            break;
                        case "filterRows":
                            string filterOperator = proposed.GetProperty("operator").GetString() ?? string.Empty;
                            accepted.Add(new FilterRowsTransform
                            {
                                Id = id,
                                Column = sourceField,
                                Operator = ParseAgentRowFilterOperator(filterOperator),
                                Value = string.Equals(filterOperator, "isBlank", StringComparison.Ordinal) ||
                                    string.Equals(filterOperator, "isNotBlank", StringComparison.Ordinal)
                                    ? null
                                    : ScalarValue.FromText(
                                        proposed.GetProperty("value").GetString() ?? string.Empty)
                            });
                            appendRename = false;
                            break;
                        case "mapValues":
                            var map = new MapValuesTransform
                            {
                                Id = id,
                                Column = sourceField
                            };
                            foreach (JsonElement mapping in proposed.GetProperty("mappings").EnumerateArray())
                            {
                                map.Entries.Add(new ValueMapEntry
                                {
                                    From = ScalarValue.FromText(
                                        mapping.GetProperty("from").GetString() ?? string.Empty),
                                    To = ScalarValue.FromText(
                                        mapping.GetProperty("to").GetString() ?? string.Empty)
                                });
                            }

                            accepted.Add(map);
                            break;
                        case "excludeTotalRows":
                            var evidence = new TotalRowEvidenceSpec
                            {
                                Column = sourceField,
                                MatchKind = ParseAgentTotalRowMatchKind(
                                    proposed.GetProperty("matchKind").GetString()),
                                Source = ParseAgentEvidenceSource(
                                    proposed.GetProperty("evidenceSource").GetString()),
                                ObservedMatchCount = proposed
                                    .GetProperty("observedMatchCount")
                                    .GetInt64()
                            };
                            foreach (JsonElement value in proposed.GetProperty("values").EnumerateArray())
                            {
                                evidence.Values.Add(ScalarValue.FromText(value.GetString() ?? string.Empty));
                            }

                            accepted.Add(new ExcludeTotalRowsTransform
                            {
                                Id = id,
                                Evidence = new List<TotalRowEvidenceSpec> { evidence },
                                RequireAllEvidence = true
                            });
                            appendRename = false;
                            break;
                        case "derivePeriodPart":
                            accepted.Add(new DerivePeriodPartsTransform
                            {
                                Id = id,
                                DateColumn = sourceField,
                                Columns = new List<DerivedPeriodColumnSpec>
                                {
                                    new DerivedPeriodColumnSpec
                                    {
                                        Part = ParseAgentDerivedPeriodPart(
                                            proposed.GetProperty("part").GetString()),
                                        OutputColumn = outputField
                                    }
                                }
                            });
                            appendRename = false;
                            break;
                        case "addArithmeticColumn":
                            string rightField = proposed.GetProperty("rightField").GetString() ?? string.Empty;
                            JsonElement rightNumber = proposed.GetProperty("rightNumber");
                            accepted.Add(new AddArithmeticColumnTransform
                            {
                                Id = id,
                                OutputColumn = outputField,
                                Operator = ParseAgentArithmeticOperator(
                                    proposed.GetProperty("operator").GetString()),
                                Left = new ArithmeticOperand
                                {
                                    Kind = ArithmeticOperandKind.Column,
                                    Column = sourceField
                                },
                                Right = rightField.Length == 0
                                    ? new ArithmeticOperand
                                    {
                                        Kind = ArithmeticOperandKind.Number,
                                        Number = rightNumber.GetDecimal()
                                    }
                                    : new ArithmeticOperand
                                    {
                                        Kind = ArithmeticOperandKind.Column,
                                        Column = rightField
                                    },
                                ResultType = ColumnDataType.DecimalNumber,
                                ReturnNullOnZeroDenominator = true
                            });
                            appendRename = false;
                            break;
                        default:
                            return ToolFailure(
                                request,
                                "transform_kind_not_supported",
                                "Use only the bounded transform kinds supplied by the host.");
                    }

                    if (appendRename)
                    {
                        accepted.Add(new RenameColumnTransform
                        {
                            Id = id + "_rename",
                            From = sourceField,
                            To = outputField
                        });
                    }
                }

                context.ProposedTransforms = accepted;
                return ToolSuccess(
                    request,
                    "transforms_accepted",
                    new { transformCount = accepted.Count });
            }
        }

        private HostToolResultRequest ValidateAgentSpecification(
            HostToolRequestEvent request,
            ChatToolContext context)
        {
            using (var document = JsonDocument.Parse(request.ArgumentsJson))
            {
                string proposalId = document.RootElement
                    .GetProperty("proposalToolCallId")
                    .GetString() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(context.ProposalToolCallId) ||
                    !string.Equals(proposalId, context.ProposalToolCallId, StringComparison.Ordinal) ||
                    string.IsNullOrWhiteSpace(context.ProposalArgumentsJson))
                {
                    return ToolFailure(
                        request,
                        "proposal_reference_invalid",
                        "Validate the report proposal returned by the current guarded workflow.");
                }

                SelectedSourceState source = ReadSelectedSource();
                ReportSpecV1 specification = _translator.FromAgentProposal(
                    context.ProposalArgumentsJson!,
                    context.PeriodMapping,
                    source.Profile,
                    PlaceholderSourceName,
                    DetermineSourceKind(source.Selection),
                    _reportId,
                    context.OutputMode,
                    context.ProposedTransforms);
                ReportBuildPlan plan = ReportBuildPlanner.Create(specification, source.Profile, null);
                context.ValidatedSpecification = specification;
                context.ValidatedSpecificationId = SafeOutcomeCode(
                    "validated_" + request.ToolCallId);
                return ToolSuccess(
                    request,
                    "spec_validated",
                    new
                    {
                        validatedSpecificationId = context.ValidatedSpecificationId,
                        projectedRows = plan.Source.ProjectedRows,
                        blockCount = plan.Blocks.Count
                    });
            }
        }

        private HostToolResultRequest BuildAgentDraft(
            HostToolRequestEvent request,
            ChatToolContext context,
            CancellationToken cancellationToken)
        {
            using (var document = JsonDocument.Parse(request.ArgumentsJson))
            {
                string validatedId = document.RootElement
                    .GetProperty("validatedSpecificationId")
                    .GetString() ?? string.Empty;
                if (context.ValidatedSpecification == null ||
                    !string.Equals(
                        validatedId,
                        context.ValidatedSpecificationId,
                        StringComparison.Ordinal))
                {
                    return ToolFailure(
                        request,
                        "validated_spec_reference_invalid",
                        "Build only the report setup accepted by deterministic validation.");
                }

                SelectedSourceState source = ReadSelectedSource();
                try
                {
                    ExcelBuildResult result = BuildCoreSpecification(
                        context.ValidatedSpecification,
                        source,
                        cancellationToken);
                    context.BuildResult = result;
                    context.ManagedDraftId = "draft_" + context.ValidatedSpecification.Id;
                    context.ChecksPassed = false;
                    context.CheckResults = Array.Empty<HostCheckResult>();
                    context.FinalChanges = Array.Empty<ChatChangeSnapshot>();
                    return ToolSuccess(
                        request,
                        "managed_draft_built",
                        new
                        {
                            managedDraftId = context.ManagedDraftId,
                            normalizedRows = result.NormalizedRows,
                            draftCount = result.DraftWorksheets.Count
                        });
                }
                catch (ExcelBuildValidationException exception)
                {
                    context.ManagedDraftId = "draft_" + context.ValidatedSpecification.Id;
                    context.ChecksPassed = false;
                    var failures = exception.Result.Checks
                        .Where(check => check.Outcome == CheckOutcome.Failed)
                        .Take(50)
                        .Select(check => new HostCheckFailure
                        {
                            Code = SafeOutcomeCode(check.CheckId),
                            Message = BoundedMessage(check.Message)
                        })
                        .ToList();
                    return new HostToolResultRequest
                    {
                        JobId = request.JobId,
                        ToolCallId = request.ToolCallId,
                        Succeeded = false,
                        OutcomeCode = "managed_draft_checks_failed",
                        ResultJson = JsonSerializer.Serialize(new
                        {
                            managedDraftId = context.ManagedDraftId,
                            completedChecks = exception.Result.Checks.Count
                        }),
                        CheckFailures = failures
                    };
                }
            }
        }

        private HostToolResultRequest RunAgentChecks(
            HostToolRequestEvent request,
            ChatToolContext context,
            CancellationToken cancellationToken)
        {
            using (var document = JsonDocument.Parse(request.ArgumentsJson))
            {
                string draftId = document.RootElement
                    .GetProperty("managedDraftId")
                    .GetString() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(context.ManagedDraftId) ||
                    !string.Equals(draftId, context.ManagedDraftId, StringComparison.Ordinal))
                {
                    return ToolFailure(
                        request,
                        "managed_draft_reference_invalid",
                        "Run checks only against the managed draft built by this workflow.");
                }

                var requested = new HashSet<string>(
                    document.RootElement.GetProperty("checks")
                        .EnumerateArray()
                        .Select(value => value.GetString() ?? string.Empty),
                    StringComparer.Ordinal);
                IReadOnlyList<HostCheckResult> allChecks = RunChecksCore(cancellationToken);
                if (_lastBuild != null)
                {
                    context.BuildResult = _lastBuild;
                }
                var selected = allChecks
                    .Where(check => requested.Any(category => CheckMatchesCategory(check.Name, category)))
                    .ToList();
                foreach (string category in requested)
                {
                    if (!selected.Any(check => CheckMatchesCategory(check.Name, category)))
                    {
                        selected.Add(new HostCheckResult(
                            "requested-" + category,
                            false,
                            "The requested check is not applicable to the current managed report setup."));
                    }
                }

                foreach (HostCheckResult failure in allChecks.Where(check => !check.Passed))
                {
                    if (!selected.Any(check => string.Equals(check.Name, failure.Name, StringComparison.Ordinal)))
                    {
                        selected.Add(failure);
                    }
                }

                List<HostCheckResult> failed = selected.Where(check => !check.Passed).ToList();
                context.CheckResults = allChecks;
                context.ChecksPassed = allChecks.All(check => check.Passed) && failed.Count == 0;
                if (failed.Count != 0)
                {
                    return new HostToolResultRequest
                    {
                        JobId = request.JobId,
                        ToolCallId = request.ToolCallId,
                        Succeeded = false,
                        OutcomeCode = "managed_draft_checks_failed",
                        ResultJson = JsonSerializer.Serialize(new
                        {
                            managedDraftId = context.ManagedDraftId,
                            completedChecks = selected.Count
                        }),
                        CheckFailures = failed.Take(50).Select(check => new HostCheckFailure
                        {
                            Code = SafeOutcomeCode(check.Name),
                            Message = BoundedMessage(check.Detail)
                        }).ToList()
                    };
                }

                return ToolSuccess(
                    request,
                    "managed_draft_checks_passed",
                    new
                    {
                        managedDraftId = context.ManagedDraftId,
                        completedChecks = selected.Count,
                        allChecksPassed = true
                    });
            }
        }

        private static bool CheckMatchesCategory(string checkName, string category)
        {
            switch (category)
            {
                case "sourceTotals":
                    return checkName.IndexOf("source-to-normalized", StringComparison.OrdinalIgnoreCase) >= 0 ||
                           checkName.IndexOf("total-preservation", StringComparison.OrdinalIgnoreCase) >= 0;
                case "grandTotals":
                    return checkName.IndexOf("source-to-pivot", StringComparison.OrdinalIgnoreCase) >= 0 ||
                           checkName.IndexOf("source-to-output", StringComparison.OrdinalIgnoreCase) >= 0 ||
                           checkName.IndexOf("balance", StringComparison.OrdinalIgnoreCase) >= 0;
                case "rowCounts":
                    return checkName.IndexOf("truncation", StringComparison.OrdinalIgnoreCase) >= 0 ||
                           checkName.IndexOf("row", StringComparison.OrdinalIgnoreCase) >= 0;
                case "periodCoverage":
                    return checkName.IndexOf("period", StringComparison.OrdinalIgnoreCase) >= 0;
                case "formulaErrors":
                    return checkName.IndexOf("formula", StringComparison.OrdinalIgnoreCase) >= 0;
                case "managedOwnership":
                    return checkName.IndexOf("ownership", StringComparison.OrdinalIgnoreCase) >= 0;
                default:
                    return false;
            }
        }

        private static HostToolResultRequest AcceptFinalSummary(
            HostToolRequestEvent request,
            ChatToolContext context)
        {
            using (var document = JsonDocument.Parse(request.ArgumentsJson))
            {
                JsonElement root = document.RootElement;
                string draftId = root.GetProperty("managedDraftId").GetString() ?? string.Empty;
                bool allChecksPassed = root.GetProperty("allChecksPassed").GetBoolean();
                if (!context.ChecksPassed || !allChecksPassed ||
                    !string.Equals(draftId, context.ManagedDraftId, StringComparison.Ordinal))
                {
                    return ToolFailure(
                        request,
                        "final_summary_gate_failed",
                        "The final summary must reference the checked managed draft and report passing checks.");
                }

                var changes = new List<ChatChangeSnapshot>();
                if (context.ValidatedSpecification != null)
                {
                    changes.Add(new ChatChangeSnapshot(
                        "data",
                        "Applied " + context.ValidatedSpecification.Transforms.Count.ToString(CultureInfo.InvariantCulture) +
                        " typed transformation step(s) to the managed canonical source."));
                    changes.Add(new ChatChangeSnapshot(
                        "formatting",
                        "Built " + context.ValidatedSpecification.Blocks.Count.ToString(CultureInfo.InvariantCulture) +
                        " managed report block(s) from the validated ReportSpecV1."));
                    changes.Add(new ChatChangeSnapshot(
                        "values",
                        "Configured " + context.ValidatedSpecification.Measures.Count.ToString(CultureInfo.InvariantCulture) +
                        " typed measure(s) with no raw formula or code execution."));
                }

                changes.Add(new ChatChangeSnapshot(
                    "checks",
                    "Passed " + context.CheckResults.Count.ToString(CultureInfo.InvariantCulture) +
                    " independent managed-draft check(s)."));
                foreach (JsonElement change in root.GetProperty("changes").EnumerateArray())
                {
                    changes.Add(new ChatChangeSnapshot(
                        change.GetProperty("category").GetString() ?? "checks",
                        change.GetProperty("description").GetString() ?? string.Empty));
                }

                context.FinalChanges = changes;
                return ToolSuccess(
                    request,
                    "final_summary_accepted",
                    new
                    {
                        managedDraftId = context.ManagedDraftId,
                        allChecksPassed = true,
                        changeCount = root.GetProperty("changes").GetArrayLength(),
                        published = false,
                        workbookSaved = false
                    });
            }
        }

        private static HostToolResultRequest ToolSuccess(
            HostToolRequestEvent request,
            string outcomeCode,
            object result)
        {
            return new HostToolResultRequest
            {
                JobId = request.JobId,
                ToolCallId = request.ToolCallId,
                Succeeded = true,
                OutcomeCode = SafeOutcomeCode(outcomeCode),
                ResultJson = JsonSerializer.Serialize(result)
            };
        }

        private static HostToolResultRequest ToolFailure(
            HostToolRequestEvent request,
            string outcomeCode,
            string message,
            string failureCode = "host_gate_rejected")
        {
            return new HostToolResultRequest
            {
                JobId = request.JobId,
                ToolCallId = request.ToolCallId,
                Succeeded = false,
                OutcomeCode = SafeOutcomeCode(outcomeCode),
                ResultJson = "{}",
                CheckFailures = new List<HostCheckFailure>
                {
                    new HostCheckFailure
                    {
                        Code = SafeOutcomeCode(failureCode),
                        Message = BoundedMessage(message)
                    }
                }
            };
        }

        private AgentJobRequest CreateAgentJob(
            string prompt,
            ReportSpecificationSnapshot specification,
            SelectedSourceState source,
            AgentEndpointSettings endpoint)
        {
            var data = new AgentDataSnapshot
            {
                SourceDisplayName = "Selected Data",
                RowCount = source.SelectionSnapshot.RowCount,
                ReportingYear = specification.PeriodMapping.ReportingYear
            };
            foreach (SourceColumnProfile column in source.Profile.Columns)
            {
                data.Fields.Add(new AgentField
                {
                    Name = column.Name,
                    Type = ToAgentFieldType(column.InferredType),
                    AllowsBlank = column.BlankCount > 0
                });
            }

            foreach (object?[] row in source.Rows.Take(AgentSampleRowLimit))
            {
                var agentRow = new AgentSampleRow();
                for (int index = 0; index < source.SelectionSnapshot.Headers.Count; index++)
                {
                    string? value = row[index] == null ? null : FormatValue(row[index]);
                    if (value != null && value.Length > 1024)
                    {
                        value = value.Substring(0, 1024);
                    }

                    agentRow.Values.Add(new AgentSampleValue
                    {
                        Field = source.SelectionSnapshot.Headers[index],
                        Value = value
                    });
                }

                data.SampleRows.Add(agentRow);
            }

            var current = new AgentSpecificationSnapshot();
            if (specification.HasCanonicalReportSpec)
            {
                // The canonical setup is already host-validated and preserves
                // calculated measures and multiple blocks that the simple
                // Rows/Columns projection cannot represent without data loss.
                current.CanonicalReportSpecJson = specification.CanonicalReportSpecJson;
            }
            else
            {
                foreach (FieldPlacementSnapshot placement in specification.Placements)
                {
                    switch (placement.Bucket)
                    {
                        case PlacementBucket.Rows:
                            current.Rows.Add(placement.ColumnName);
                            break;
                        case PlacementBucket.Columns:
                            current.Columns.Add(placement.ColumnName);
                            break;
                        case PlacementBucket.Values:
                            current.Values.Add(new AgentValuePlacement
                            {
                                Field = placement.ColumnName,
                                Aggregation = ToAgentAggregation(placement.Setting)
                            });
                            break;
                        case PlacementBucket.Filters:
                            current.Filters.Add(new AgentFilterPlacement
                            {
                                Field = placement.ColumnName,
                                Operator = "equals"
                            });
                            break;
                    }
                }
            }

            var job = new AgentJobRequest
            {
                JobId = string.Empty,
                WorkbookId = _workbookId,
                UserPrompt = prompt,
                Data = data,
                CurrentSpecification = current,
                Endpoint = endpoint,
                MaxRepairCycles = AgentDefaults.MaxRepairCycles
            };
            job.JobId = AgentJobIdentity.Create(job);
            return job;
        }

        private SelectedSourceState ReadCurrentSelection()
        {
            dynamic application = _excelApplication;
            object? selection = application.Selection as object;
            object? workbook = application.ActiveWorkbook as object;
            if (selection == null || workbook == null)
            {
                throw new InvalidOperationException(
                    "Open a workbook and select one rectangular table or range.");
            }

            SourceSelectionSnapshot snapshot = _selectionInspector.Inspect(application);
            return CreateSourceState(selection, workbook, snapshot);
        }

        private SelectedSourceState ReadSelectedSource()
        {
            if (_selectedSource == null)
            {
                throw new InvalidOperationException("Choose workbook Data before continuing.");
            }

            object activeWorkbook = GetActiveWorkbook();
            if (!AreSameComObject(activeWorkbook, _selectedSource.Workbook))
            {
                throw new InvalidOperationException("Activate the workbook that contains the chosen Data.");
            }

            dynamic proxy = new ExpandoObject();
            proxy.Selection = _selectedSource.Selection;
            SourceSelectionSnapshot snapshot = _selectionInspector.Inspect(proxy);
            SelectedSourceState refreshed = CreateSourceState(
                _selectedSource.Selection,
                _selectedSource.Workbook,
                snapshot);
            _selectedSource = refreshed;
            return refreshed;
        }

        private SelectedSourceState CreateSourceState(
            object selection,
            object workbook,
            SourceSelectionSnapshot snapshot)
        {
            List<object?[]> rows = snapshot.SampleRows
                .Select(row => row.ToArray())
                .ToList();
            CoerceFormattedDates(selection, rows);
            SourceProfile profile = SourceProfiler.Profile(snapshot.Headers, rows);
            profile.RowCount = snapshot.RowCount;
            return new SelectedSourceState(selection, workbook, snapshot, rows, profile);
        }

        private long CountObservedTotalRows(
            SelectedSourceState source,
            ExcludeTotalRowsTransform exclusion)
        {
            if (exclusion.Evidence == null || exclusion.Evidence.Count != 1)
            {
                throw new InvalidOperationException(
                    "Manual total-row exclusion requires exactly one confirmed source-column rule.");
            }

            TotalRowEvidenceSpec evidence = exclusion.Evidence[0];
            SourceColumnProfile? column = source.Profile.FindColumn(evidence.Column);
            if (column == null)
            {
                throw new InvalidOperationException(
                    "The total-row evidence column is not present in the selected Data.");
            }

            Report(
                ActivityStage.Inspecting,
                ActivityKind.Progress,
                "Confirming total-row evidence across the full source.",
                evidence.Column + " may require a blocking Excel column read.");

            dynamic selection = source.Selection;
            int dataRows = checked(Convert.ToInt32(
                source.SelectionSnapshot.RowCount,
                CultureInfo.InvariantCulture));
            dynamic first = selection.Cells[2, column.Index + 1];
            dynamic last = selection.Cells[dataRows + 1, column.Index + 1];
            dynamic range = selection.Worksheet.Range[first, last];
            object? values = range.Value2;
            long matches = 0;
            if (values is Array array && array.Rank == 2)
            {
                int lowerRow = array.GetLowerBound(0);
                int upperRow = array.GetUpperBound(0);
                int columnIndex = array.GetLowerBound(1);
                for (int row = lowerRow; row <= upperRow; row++)
                {
                    if (MatchesTotalRowEvidence(array.GetValue(row, columnIndex), evidence))
                    {
                        matches++;
                    }
                }
            }
            else if (MatchesTotalRowEvidence(values, evidence))
            {
                matches = 1;
            }

            Report(
                ActivityStage.Inspecting,
                matches > 0 ? ActivityKind.Check : ActivityKind.Error,
                matches > 0
                    ? "Confirmed source total rows."
                    : "No source total rows matched the confirmation.",
                matches.ToString("N0", CultureInfo.InvariantCulture) + " matching rows observed.");
            return matches;
        }

        private static bool MatchesTotalRowEvidence(object? sourceValue, TotalRowEvidenceSpec evidence)
        {
            if (evidence.MatchKind == TotalRowMatchKind.IsBlank)
            {
                return sourceValue == null ||
                    sourceValue is string blankText && string.IsNullOrWhiteSpace(blankText);
            }

            if (evidence.Values == null || evidence.Values.Count == 0)
            {
                return false;
            }

            foreach (ScalarValue expected in evidence.Values)
            {
                if (evidence.MatchKind == TotalRowMatchKind.EqualsAny &&
                    ScalarMatchesExcelValue(sourceValue, expected))
                {
                    return true;
                }

                if (sourceValue is string sourceText && expected.Kind == ScalarValueKind.Text)
                {
                    string expectedText = expected.Text ?? string.Empty;
                    if (evidence.MatchKind == TotalRowMatchKind.StartsWith &&
                        sourceText.StartsWith(expectedText, StringComparison.Ordinal) ||
                        evidence.MatchKind == TotalRowMatchKind.Contains &&
                        sourceText.IndexOf(expectedText, StringComparison.Ordinal) >= 0)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool ScalarMatchesExcelValue(object? sourceValue, ScalarValue expected)
        {
            switch (expected.Kind)
            {
                case ScalarValueKind.Null:
                    return sourceValue == null;
                case ScalarValueKind.Text:
                    return sourceValue is string text &&
                        string.Equals(text, expected.Text ?? string.Empty, StringComparison.Ordinal);
                case ScalarValueKind.Boolean:
                    return sourceValue is bool boolean && expected.Boolean == boolean;
                case ScalarValueKind.Number:
                    return TryReadExcelNumber(sourceValue, out decimal number) &&
                        expected.Number == number;
                case ScalarValueKind.Date:
                    return expected.Temporal.HasValue &&
                        TryReadExcelDate(sourceValue, out DateTime date) &&
                        expected.Temporal.Value.Date == date.Date;
                case ScalarValueKind.DateTime:
                    return expected.Temporal.HasValue &&
                        TryReadExcelDate(sourceValue, out DateTime dateTime) &&
                        Math.Abs((expected.Temporal.Value - dateTime).TotalMilliseconds) < 1d;
                default:
                    return false;
            }
        }

        private static bool TryReadExcelNumber(object? value, out decimal number)
        {
            number = 0m;
            if (value == null || value is bool || value is string || value is DateTime)
            {
                return false;
            }

            try
            {
                number = Convert.ToDecimal(value, CultureInfo.InvariantCulture);
                return true;
            }
            catch (Exception exception) when (
                exception is FormatException ||
                exception is InvalidCastException ||
                exception is OverflowException)
            {
                return false;
            }
        }

        private static bool TryReadExcelDate(object? value, out DateTime date)
        {
            if (value is DateTime temporal)
            {
                date = temporal;
                return true;
            }

            if (TryReadExcelNumber(value, out decimal serial))
            {
                try
                {
                    date = DateTime.FromOADate(Convert.ToDouble(serial, CultureInfo.InvariantCulture));
                    return true;
                }
                catch (ArgumentException)
                {
                }
            }

            date = default;
            return false;
        }

        private static void CoerceFormattedDates(object selectionObject, IList<object?[]> rows)
        {
            dynamic selection = selectionObject;
            for (int column = 1; column <= Convert.ToInt32(selection.Columns.Count, CultureInfo.InvariantCulture); column++)
            {
                string format;
                try
                {
                    format = Convert.ToString(
                        selection.Cells[2, column].NumberFormat,
                        CultureInfo.InvariantCulture) ?? string.Empty;
                }
                catch (Exception)
                {
                    continue;
                }

                if (!LooksLikeDateFormat(format))
                {
                    continue;
                }

                foreach (object?[] row in rows)
                {
                    object? value = row[column - 1];
                    try
                    {
                        if (value != null &&
                            !(value is DateTime) &&
                            double.TryParse(
                                Convert.ToString(value, CultureInfo.InvariantCulture),
                                NumberStyles.Float,
                                CultureInfo.InvariantCulture,
                                out double serial))
                        {
                            row[column - 1] = DateTime.FromOADate(serial);
                        }
                    }
                    catch (ArgumentException)
                    {
                    }
                }
            }
        }

        private static bool LooksLikeDateFormat(string format)
        {
            if (string.IsNullOrWhiteSpace(format))
            {
                return false;
            }

            string normalized = format.ToLowerInvariant();
            return normalized.Contains("yy") ||
                normalized.Contains("dd") ||
                (normalized.Contains("mmm") &&
                 (normalized.Contains("d") || normalized.Contains("y")));
        }

        private static SourceSnapshot ToSourceSnapshot(
            SelectedSourceState source,
            ReportSpecificationSnapshot? savedReportSetup = null,
            string savedReportSetupStatus = "")
        {
            var columns = new List<SourceColumnSnapshot>();
            foreach (SourceColumnProfile column in source.Profile.Columns)
            {
                string sample = source.Rows
                    .Select(row => row[column.Index])
                    .Where(value => value != null &&
                        !string.IsNullOrWhiteSpace(Convert.ToString(value, CultureInfo.InvariantCulture)))
                    .Select(FormatValue)
                    .FirstOrDefault() ?? string.Empty;
                columns.Add(new SourceColumnSnapshot(
                    column.Name,
                    ToDisplayType(column.InferredType),
                    sample));
            }

            bool table = DetermineSourceKind(source.Selection) == WorkbookSourceKind.Table;
            return new SourceSnapshot(
                table ? source.SelectionSnapshot.WorkbookObjectName : "Selected range",
                source.SelectionSnapshot.WorkbookObjectName,
                Convert.ToInt32(source.SelectionSnapshot.RowCount, CultureInfo.InvariantCulture),
                columns,
                isSynthetic: false,
                savedReportSetup: savedReportSetup,
                savedReportSetupStatus: savedReportSetupStatus);
        }

        private SavedSetupSelection FindSavedSetup(SelectedSourceState source)
        {
            SourceFingerprintSpec currentFingerprint = SourceFingerprint.FromHeaders(
                source.Profile.Columns
                    .OrderBy(column => column.Index)
                    .Select(column => column.Name));
            WorkbookSourceKind currentKind = DetermineSourceKind(source.Selection);
            List<ReportSpecV1> compatible = _workbookSpecStore.LoadAll(source.Workbook)
                .Where(specification => SourceMatches(
                    specification.Source,
                    currentKind,
                    currentFingerprint,
                    source))
                .ToList();
            if (compatible.Count == 0)
            {
                return new SavedSetupSelection(
                    null,
                    null,
                    "No saved report setup matches this exact workbook Data object.");
            }

            if (compatible.Count > 1)
            {
                return new SavedSetupSelection(
                    null,
                    null,
                    "More than one saved report setup matches this Data object, so none was restored automatically.");
            }

            ReportSpecV1 specification = compatible[0];
            try
            {
                ReportSpecificationSnapshot snapshot = _translator.ToAppliedAgentSnapshot(specification);
                PeriodMappingSpec? resolvedMapping = _translator.ResolvePeriodMapping(
                    snapshot.PeriodMapping,
                    source.Profile);
                if (!PeriodMappingsEqual(specification.PeriodMapping, resolvedMapping))
                {
                    return new SavedSetupSelection(
                        null,
                        null,
                        "A matching saved setup exists, but its period mapping no longer resolves identically against the selected Data.");
                }

                return new SavedSetupSelection(
                    specification,
                    snapshot,
                    "Compatible saved report setup restored. Build a managed draft to use current source values.");
            }
            catch (NotSupportedException)
            {
                throw;
            }
            catch (InvalidOperationException exception)
            {
                return new SavedSetupSelection(
                    null,
                    null,
                    "A matching saved setup exists but was not restored: " + exception.Message);
            }
        }

        private static bool PeriodMappingsEqual(
            PeriodMappingSpec? saved,
            PeriodMappingSpec? current)
        {
            if (saved == null || current == null)
            {
                return saved == null && current == null;
            }

            if (saved.Kind != current.Kind ||
                !string.Equals(saved.DateColumn, current.DateColumn, StringComparison.OrdinalIgnoreCase) ||
                saved.ReportingYear != current.ReportingYear ||
                !string.Equals(saved.PeriodColumnName, current.PeriodColumnName, StringComparison.Ordinal) ||
                !string.Equals(saved.ValueColumnName, current.ValueColumnName, StringComparison.Ordinal) ||
                !string.Equals(saved.MetricColumnName, current.MetricColumnName, StringComparison.Ordinal) ||
                !saved.KeyColumns.SequenceEqual(
                    current.KeyColumns,
                    StringComparer.OrdinalIgnoreCase) ||
                saved.Columns.Count != current.Columns.Count)
            {
                return false;
            }

            for (int index = 0; index < saved.Columns.Count; index++)
            {
                PeriodColumnMapping expected = saved.Columns[index];
                PeriodColumnMapping actual = current.Columns[index];
                if (!string.Equals(
                        expected.SourceColumn,
                        actual.SourceColumn,
                        StringComparison.OrdinalIgnoreCase) ||
                    expected.Month != actual.Month ||
                    expected.Year != actual.Year ||
                    !string.Equals(expected.Metric, actual.Metric, StringComparison.Ordinal))
                {
                    return false;
                }
            }

            return true;
        }

        private bool SourceMatches(
            WorkbookSourceSpec savedSource,
            WorkbookSourceKind currentKind,
            SourceFingerprintSpec currentFingerprint,
            SelectedSourceState currentSource)
        {
            bool objectMatches = currentKind == WorkbookSourceKind.Table
                ? string.Equals(
                    savedSource.WorkbookObjectName,
                    currentSource.SelectionSnapshot.WorkbookObjectName,
                    StringComparison.OrdinalIgnoreCase)
                : NamedRangeRefersToSelection(
                    currentSource.Workbook,
                    savedSource.WorkbookObjectName,
                    currentSource.Selection);
            return SavedSetupCompatibility.Matches(
                savedSource,
                currentKind,
                currentFingerprint,
                objectMatches);
        }

        private static bool NamedRangeRefersToSelection(
            object workbookObject,
            string workbookObjectName,
            object selectionObject)
        {
            if (string.IsNullOrWhiteSpace(workbookObjectName))
            {
                return false;
            }

            try
            {
                dynamic workbook = workbookObject;
                dynamic savedRange = workbook.Names.Item(workbookObjectName).RefersToRange;
                dynamic selection = selectionObject;
                object? savedWorksheet = savedRange.Worksheet as object;
                object? selectedWorksheet = selection.Worksheet as object;
                if (savedWorksheet == null || selectedWorksheet == null ||
                    !AreSameComObject(savedWorksheet, selectedWorksheet))
                {
                    return false;
                }

                string savedAddress = Convert.ToString(
                    savedRange.Address[true, true, 1, false],
                    CultureInfo.InvariantCulture) ?? string.Empty;
                string selectedAddress = Convert.ToString(
                    selection.Address[true, true, 1, false],
                    CultureInfo.InvariantCulture) ?? string.Empty;
                return string.Equals(savedAddress, selectedAddress, StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private void EnsureSameChosenSource(SelectedSourceState current)
        {
            if (_selectedSource == null)
            {
                throw new InvalidOperationException("Choose workbook Data before previewing period headers.");
            }

            bool sameHeaders = current.SelectionSnapshot.Headers.SequenceEqual(
                _selectedSource.SelectionSnapshot.Headers,
                StringComparer.OrdinalIgnoreCase);
            if (!AreSameComObject(current.Workbook, _selectedSource.Workbook) ||
                current.SelectionSnapshot.RowCount != _selectedSource.SelectionSnapshot.RowCount ||
                current.SelectionSnapshot.ColumnCount != _selectedSource.SelectionSnapshot.ColumnCount ||
                !sameHeaders ||
                !string.Equals(
                    current.SelectionSnapshot.WorkbookObjectName,
                    _selectedSource.SelectionSnapshot.WorkbookObjectName,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "The Excel selection changed. Choose the intended Data again before previewing.");
            }
        }

        private static IReadOnlyList<IReadOnlyDictionary<string, object?>> ToRowDictionaries(
            SelectedSourceState source)
        {
            var result = new List<IReadOnlyDictionary<string, object?>>();
            foreach (object?[] row in source.Rows)
            {
                var values = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                for (int index = 0; index < source.SelectionSnapshot.Headers.Count; index++)
                {
                    values[source.SelectionSnapshot.Headers[index]] = row[index];
                }

                result.Add(values);
            }

            return result;
        }

        private static decimal SumMappedValues(
            SelectedSourceState source,
            PeriodMappingSpec mapping)
        {
            var indexes = source.SelectionSnapshot.Headers
                .Select((name, index) => new { name, index })
                .ToDictionary(item => item.name, item => item.index, StringComparer.OrdinalIgnoreCase);
            decimal total = 0m;
            foreach (object?[] row in source.Rows)
            {
                foreach (PeriodColumnMapping column in mapping.Columns)
                {
                    total += ToDecimalOrZero(row[indexes[column.SourceColumn]]);
                }
            }

            return total;
        }

        private static decimal ToDecimalOrZero(object? value)
        {
            if (value == null || value == DBNull.Value)
            {
                return 0m;
            }

            try
            {
                return Convert.ToDecimal(value, CultureInfo.InvariantCulture);
            }
            catch (Exception exception) when (
                exception is FormatException ||
                exception is InvalidCastException ||
                exception is OverflowException)
            {
                return 0m;
            }
        }

        private static string FormatPeriod(PeriodHeaderMatch match, int? reportingYear)
        {
            int? year = match.Year ?? reportingYear;
            return year.HasValue
                ? year.Value.ToString("0000", CultureInfo.InvariantCulture) + "-" +
                  match.Month.ToString("00", CultureInfo.InvariantCulture)
                : "????-" + match.Month.ToString("00", CultureInfo.InvariantCulture);
        }

        private static string FormatValue(object? value)
        {
            if (value == null || value == DBNull.Value)
            {
                return string.Empty;
            }

            if (value is DateTime date)
            {
                return date.TimeOfDay == TimeSpan.Zero
                    ? date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                    : date.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            }

            if (value is double number)
            {
                return number.ToString("G15", CultureInfo.InvariantCulture);
            }

            return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
        }

        private async Task<AgentEndpointSettings> MaterializeEndpointAsync(
            ModelEndpointSettingsSnapshot snapshot,
            SecureString? apiKey,
            CancellationToken cancellationToken)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            string? clearApiKey = SecureStringToString(apiKey);
            if (string.IsNullOrWhiteSpace(clearApiKey))
            {
                PersistedAgentSettings? persisted = await _settingsStore.LoadAsync(cancellationToken)
                    .ConfigureAwait(false);
                if (persisted != null &&
                    string.Equals(
                        persisted.BaseUrl?.TrimEnd('/'),
                        snapshot.BaseUrl.TrimEnd('/'),
                        StringComparison.OrdinalIgnoreCase) &&
                    persisted.AllowRemoteHttp == snapshot.AllowRemoteHttp)
                {
                    AgentEndpointSettings saved = AgentSettingsMaterializer.Unprotect(
                        persisted,
                        _secretProtector);
                    clearApiKey = saved.ApiKey;
                }
            }

            return new AgentEndpointSettings
            {
                BaseUrl = snapshot.BaseUrl,
                Model = string.IsNullOrWhiteSpace(snapshot.ModelId)
                    ? AgentDefaults.Model
                    : snapshot.ModelId,
                AllowRemoteHttp = snapshot.AllowRemoteHttp,
                AllowRemoteWorkbookData = snapshot.AllowRemoteWorkbookData,
                ApiKey = clearApiKey
            };
        }

        private async Task SaveEndpointAsync(
            AgentEndpointSettings endpoint,
            CancellationToken cancellationToken)
        {
            PersistedAgentSettings persisted = AgentSettingsMaterializer.Protect(
                endpoint,
                _secretProtector);
            await _settingsStore.SaveAsync(persisted, cancellationToken).ConfigureAwait(false);
        }

        private SavedEndpointSettingsSnapshot? LoadSavedEndpointSettings()
        {
            PersistedAgentSettings? settings = _settingsStore.TryLoad();
            return settings == null
                ? null
                : new SavedEndpointSettingsSnapshot(
                    settings.BaseUrl,
                    settings.Model,
                    settings.AllowRemoteHttp,
                    !string.IsNullOrWhiteSpace(settings.ProtectedApiKey),
                    settings.AllowRemoteWorkbookData);
        }

        private async Task<T> AwaitWithHeartbeatAsync<T>(
            Task<T> operation,
            string heartbeatMessage,
            CancellationToken cancellationToken)
        {
            while (!operation.IsCompleted)
            {
                Task delay = Task.Delay(TimeSpan.FromSeconds(5), cancellationToken);
                if (await Task.WhenAny(operation, delay).ConfigureAwait(false) == operation)
                {
                    break;
                }

                cancellationToken.ThrowIfCancellationRequested();
                Report(
                    ActivityStage.Ready,
                    ActivityKind.Heartbeat,
                    "Endpoint activity heartbeat.",
                    heartbeatMessage);
            }

            return await operation.ConfigureAwait(false);
        }

        private Task ReportWorkerProgressAsync(
            AgentProgressEvent progress,
            CancellationToken cancellationToken)
        {
            WaitForResume(cancellationToken, pumpDispatcher: false);
            Report(
                MapAgentStage(progress.Stage),
                ActivityKind.Progress,
                progress.Message,
                FormatUnits(progress.CompletedUnits, progress.TotalUnits));
            return Task.CompletedTask;
        }

        private Task ReportWorkerCheckpointAsync(
            AgentCheckpointEvent checkpoint,
            CancellationToken cancellationToken)
        {
            WaitForResume(cancellationToken, pumpDispatcher: false);
            Report(
                MapAgentStage(checkpoint.Stage),
                ActivityKind.Check,
                "Worker checkpoint recorded.",
                checkpoint.LastCompletedStep);
            return Task.CompletedTask;
        }

        private void ReportExcelProgress(
            ExcelProgress progress,
            CancellationToken cancellationToken)
        {
            WaitForResume(cancellationToken, pumpDispatcher: true);
            ActivityStage stage = MapExcelStage(progress.Stage);
            var details = new List<string>();
            if (!string.IsNullOrWhiteSpace(progress.ManagedObject))
            {
                details.Add("Managed object: " + progress.ManagedObject);
            }

            if (progress.SourceRows.HasValue)
            {
                details.Add("Source rows: " + progress.SourceRows.Value.ToString("N0", CultureInfo.CurrentCulture));
            }

            if (progress.ProjectedRows.HasValue)
            {
                details.Add("Projected rows: " + progress.ProjectedRows.Value.ToString("N0", CultureInfo.CurrentCulture));
            }

            if (progress.CompletedChecks > 0)
            {
                details.Add("Checks: " + progress.CompletedChecks.ToString(CultureInfo.CurrentCulture));
            }

            if (progress.Elapsed > TimeSpan.Zero)
            {
                details.Add("Elapsed: " + progress.Elapsed.ToString(@"mm\:ss", CultureInfo.InvariantCulture));
            }

            Report(
                stage,
                progress.IsHeartbeat
                    ? ActivityKind.Heartbeat
                    : progress.Stage == ExcelBuildStage.Checking
                        ? ActivityKind.Check
                        : progress.Stage == ExcelBuildStage.Complete
                            ? ActivityKind.Result
                            : ActivityKind.Progress,
                progress.Operation,
                string.Join(". ", details));
            PumpDispatcherOnce();
        }

        private void WaitForResume(CancellationToken cancellationToken, bool pumpDispatcher)
        {
            while (!_pauseGate.IsSet)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (pumpDispatcher && _excelDispatcher.CheckAccess())
                {
                    PumpDispatcherOnce();
                    Thread.Sleep(25);
                }
                else
                {
                    _pauseGate.Wait(100, cancellationToken);
                }
            }

            cancellationToken.ThrowIfCancellationRequested();
        }

        private void PumpDispatcherOnce()
        {
            if (!_excelDispatcher.CheckAccess())
            {
                return;
            }

            var frame = new DispatcherFrame();
            _excelDispatcher.BeginInvoke(
                DispatcherPriority.Background,
                new Action(() => frame.Continue = false));
            Dispatcher.PushFrame(frame);
        }

        private void Report(
            ActivityStage stage,
            ActivityKind kind,
            string message,
            string detail)
        {
            if (_disposed)
            {
                return;
            }

            Action publish = () => ActivityReported?.Invoke(
                this,
                new HostActivityEventArgs(stage, kind, message, detail));
            if (_excelDispatcher.CheckAccess())
            {
                publish();
            }
            else
            {
                _excelDispatcher.Invoke(publish, DispatcherPriority.Background);
            }
        }

        private Task<T> InvokeExcelAsync<T>(Func<T> action, CancellationToken cancellationToken)
        {
            if (action == null) throw new ArgumentNullException(nameof(action));
            cancellationToken.ThrowIfCancellationRequested();
            if (_excelDispatcher.CheckAccess())
            {
                return Task.FromResult(action());
            }

            return _excelDispatcher.InvokeAsync(
                action,
                DispatcherPriority.Send,
                cancellationToken).Task;
        }

        private async Task<T> RunOperationAsync<T>(
            CancellationToken cancellationToken,
            Func<CancellationToken, Task<T>> operation)
        {
            CancellationTokenSource linked;
            lock (_operationGate)
            {
                ThrowIfDisposed();
                if (_operationCancellation != null)
                {
                    throw new InvalidOperationException("Another report operation is already running.");
                }

                linked = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                _operationCancellation = linked;
                _pauseGate.Set();
            }

            try
            {
                return await operation(linked.Token).ConfigureAwait(false);
            }
            finally
            {
                lock (_operationGate)
                {
                    if (ReferenceEquals(_operationCancellation, linked))
                    {
                        _operationCancellation = null;
                    }
                }

                linked.Dispose();
                _pauseGate.Set();
            }
        }

        private async Task RunOperationAsync(
            CancellationToken cancellationToken,
            Func<CancellationToken, Task<bool>> operation)
        {
            await RunOperationAsync<bool>(cancellationToken, operation).ConfigureAwait(false);
        }

        private object GetActiveWorkbook()
        {
            dynamic application = _excelApplication;
            object? workbook = application.ActiveWorkbook as object;
            return workbook ?? throw new InvalidOperationException("Open a workbook before continuing.");
        }

        private static WorkbookSourceKind DetermineSourceKind(object selectionObject)
        {
            dynamic selection = selectionObject;
            try
            {
                dynamic table = selection.ListObject;
                return table == null ? WorkbookSourceKind.NamedRange : WorkbookSourceKind.Table;
            }
            catch (Exception)
            {
                return WorkbookSourceKind.NamedRange;
            }
        }

        private static bool AreSameComObject(object left, object right)
        {
            if (ReferenceEquals(left, right))
            {
                return true;
            }

            if (!Marshal.IsComObject(left) || !Marshal.IsComObject(right))
            {
                return false;
            }

            IntPtr leftPointer = IntPtr.Zero;
            IntPtr rightPointer = IntPtr.Zero;
            try
            {
                leftPointer = Marshal.GetIUnknownForObject(left);
                rightPointer = Marshal.GetIUnknownForObject(right);
                return leftPointer == rightPointer;
            }
            finally
            {
                if (leftPointer != IntPtr.Zero) Marshal.Release(leftPointer);
                if (rightPointer != IntPtr.Zero) Marshal.Release(rightPointer);
            }
        }

        private static string? SecureStringToString(SecureString? secureString)
        {
            if (secureString == null || secureString.Length == 0)
            {
                return null;
            }

            IntPtr value = IntPtr.Zero;
            try
            {
                value = Marshal.SecureStringToBSTR(secureString);
                return Marshal.PtrToStringBSTR(value);
            }
            finally
            {
                if (value != IntPtr.Zero)
                {
                    Marshal.ZeroFreeBSTR(value);
                }
            }
        }

        private static AgentFieldType ToAgentFieldType(SourceValueType type)
        {
            switch (type)
            {
                case SourceValueType.WholeNumber:
                case SourceValueType.DecimalNumber:
                    return AgentFieldType.Number;
                case SourceValueType.Boolean:
                    return AgentFieldType.Boolean;
                case SourceValueType.Date:
                case SourceValueType.DateTime:
                    return AgentFieldType.Date;
                default:
                    return AgentFieldType.Text;
            }
        }

        private static string ToDisplayType(SourceValueType type)
        {
            switch (type)
            {
                case SourceValueType.WholeNumber: return "Whole number";
                case SourceValueType.DecimalNumber: return "Number";
                case SourceValueType.Boolean: return "Yes/No";
                case SourceValueType.Date: return "Date";
                case SourceValueType.DateTime: return "Date and time";
                case SourceValueType.Empty: return "Blank";
                case SourceValueType.Mixed: return "Mixed";
                default: return "Text";
            }
        }

        private static string ToAgentAggregation(string setting)
        {
            switch ((setting ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "count": return "count";
                case "average": return "average";
                case "min":
                case "minimum": return "min";
                case "max":
                case "maximum": return "max";
                default: return "sum";
            }
        }

        private static RowFilterOperator ParseAgentRowFilterOperator(string? value)
        {
            switch (value)
            {
                case "notEqual": return RowFilterOperator.NotEqual;
                case "contains": return RowFilterOperator.Contains;
                case "startsWith": return RowFilterOperator.StartsWith;
                case "endsWith": return RowFilterOperator.EndsWith;
                case "isBlank": return RowFilterOperator.IsBlank;
                case "isNotBlank": return RowFilterOperator.IsNotBlank;
                default: return RowFilterOperator.Equal;
            }
        }

        private static DerivedPeriodPart ParseAgentDerivedPeriodPart(string? value)
        {
            switch (value)
            {
                case "half": return DerivedPeriodPart.Half;
                case "quarter": return DerivedPeriodPart.Quarter;
                case "monthNumber": return DerivedPeriodPart.MonthNumber;
                case "monthName": return DerivedPeriodPart.MonthName;
                case "yearMonth": return DerivedPeriodPart.YearMonth;
                default: return DerivedPeriodPart.Year;
            }
        }

        private static ArithmeticOperator ParseAgentArithmeticOperator(string? value)
        {
            switch (value)
            {
                case "subtract": return ArithmeticOperator.Subtract;
                case "multiply": return ArithmeticOperator.Multiply;
                case "divide": return ArithmeticOperator.Divide;
                default: return ArithmeticOperator.Add;
            }
        }

        private static TotalRowMatchKind ParseAgentTotalRowMatchKind(string? value)
        {
            switch (value)
            {
                case "startsWith": return TotalRowMatchKind.StartsWith;
                case "contains": return TotalRowMatchKind.Contains;
                case "isBlank": return TotalRowMatchKind.IsBlank;
                default: return TotalRowMatchKind.EqualsAny;
            }
        }

        private static EvidenceSource ParseAgentEvidenceSource(string? value)
        {
            switch (value)
            {
                case "preview": return EvidenceSource.Preview;
                case "userConfirmation": return EvidenceSource.UserConfirmation;
                default: return EvidenceSource.Profile;
            }
        }

        private static ActivityStage MapExcelStage(ExcelBuildStage stage)
        {
            switch (stage)
            {
                case ExcelBuildStage.Inspecting: return ActivityStage.Inspecting;
                case ExcelBuildStage.Normalizing: return ActivityStage.Normalizing;
                case ExcelBuildStage.Planning: return ActivityStage.Planning;
                case ExcelBuildStage.BuildingPivots: return ActivityStage.BuildingPivots;
                case ExcelBuildStage.Rendering: return ActivityStage.Rendering;
                case ExcelBuildStage.Calculating: return ActivityStage.Calculating;
                case ExcelBuildStage.Checking: return ActivityStage.Checking;
                case ExcelBuildStage.Repairing: return ActivityStage.Repairing;
                case ExcelBuildStage.Complete: return ActivityStage.Complete;
                default: return ActivityStage.Planning;
            }
        }

        private static ActivityStage MapAgentStage(AgentProgressStage stage)
        {
            switch (stage)
            {
                case AgentProgressStage.Accepted:
                case AgentProgressStage.ValidatingInput:
                case AgentProgressStage.DiscoveringModels:
                case AgentProgressStage.RequestingProposal:
                case AgentProgressStage.ValidatingProposal:
                    return ActivityStage.Planning;
                case AgentProgressStage.AwaitingHostTool:
                    return ActivityStage.BuildingPivots;
                case AgentProgressStage.ProcessingHostResult:
                    return ActivityStage.Checking;
                case AgentProgressStage.RepairingProposal:
                    return ActivityStage.Repairing;
                case AgentProgressStage.Completed:
                case AgentProgressStage.Cancelled:
                case AgentProgressStage.Failed:
                    return ActivityStage.Complete;
                default:
                    return ActivityStage.Planning;
            }
        }

        private static ActivityStage MapToolStage(string toolName)
        {
            switch (toolName)
            {
                case AgentToolNames.RequestManagedDraftBuild:
                    return ActivityStage.BuildingPivots;
                case AgentToolNames.RunChecks:
                case AgentToolNames.FinalChangeSummary:
                    return ActivityStage.Checking;
                default:
                    return ActivityStage.Planning;
            }
        }

        private static string FormatUnits(int? completed, int? total)
        {
            if (!completed.HasValue && !total.HasValue)
            {
                return string.Empty;
            }

            return (completed ?? 0).ToString(CultureInfo.CurrentCulture) + " of " +
                (total ?? 0).ToString(CultureInfo.CurrentCulture) + " bounded steps.";
        }

        private static string SafeOutcomeCode(string value)
        {
            var builder = new StringBuilder();
            foreach (char character in value ?? string.Empty)
            {
                if (char.IsLetterOrDigit(character) || character == '_' || character == '-' || character == '.')
                {
                    builder.Append(character);
                }
                else
                {
                    builder.Append('_');
                }
            }

            string result = builder.ToString().Trim('_');
            if (string.IsNullOrWhiteSpace(result)) result = "host_gate_rejected";
            return result.Length <= 128 ? result : result.Substring(0, 128);
        }

        private static string BoundedMessage(string value)
        {
            string result = new string((value ?? string.Empty)
                .Where(character => !char.IsControl(character))
                .ToArray()).Trim();
            if (string.IsNullOrWhiteSpace(result)) result = "The deterministic host rejected this workflow step.";
            return result.Length <= 256 ? result : result.Substring(0, 256);
        }

        private void ThrowIfDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(ExcelReportBuilderHostService));
            }
        }

        private sealed class SelectedSourceState
        {
            public SelectedSourceState(
                object selection,
                object workbook,
                SourceSelectionSnapshot selectionSnapshot,
                IReadOnlyList<object?[]> rows,
                SourceProfile profile)
            {
                Selection = selection;
                Workbook = workbook;
                SelectionSnapshot = selectionSnapshot;
                Rows = rows;
                Profile = profile;
            }

            public object Selection { get; }

            public object Workbook { get; }

            public SourceSelectionSnapshot SelectionSnapshot { get; }

            public IReadOnlyList<object?[]> Rows { get; }

            public SourceProfile Profile { get; }
        }

        private sealed class SavedSetupSelection
        {
            public SavedSetupSelection(
                ReportSpecV1? specification,
                ReportSpecificationSnapshot? uiSnapshot,
                string status)
            {
                Specification = specification;
                UiSnapshot = uiSnapshot;
                Status = status ?? string.Empty;
            }

            public ReportSpecV1? Specification { get; }

            public ReportSpecificationSnapshot? UiSnapshot { get; }

            public string Status { get; }
        }

        private sealed class ChatToolContext
        {
            public ChatToolContext(
                PeriodMappingSnapshot periodMapping,
                ReportOutputMode outputMode)
            {
                PeriodMapping = periodMapping;
                OutputMode = outputMode;
            }

            public PeriodMappingSnapshot PeriodMapping { get; set; }

            public ReportOutputMode OutputMode { get; }

            public string? ProposalToolCallId { get; set; }

            public IReadOnlyList<TransformStep> ProposedTransforms { get; set; } =
                Array.Empty<TransformStep>();

            public string? ProposalArgumentsJson { get; set; }

            public string? ValidatedSpecificationId { get; set; }

            public ReportSpecV1? ValidatedSpecification { get; set; }

            public string? ManagedDraftId { get; set; }

            public bool ChecksPassed { get; set; }

            public ExcelBuildResult? BuildResult { get; set; }

            public IReadOnlyList<HostCheckResult> CheckResults { get; set; } =
                Array.Empty<HostCheckResult>();

            public IReadOnlyList<ChatChangeSnapshot> FinalChanges { get; set; } =
                Array.Empty<ChatChangeSnapshot>();
        }

        private sealed class HostExcelProgressSink : IExcelProgressSink
        {
            private readonly ExcelReportBuilderHostService _owner;
            private readonly CancellationToken _cancellationToken;
            private bool _planningReported;

            public HostExcelProgressSink(
                ExcelReportBuilderHostService owner,
                CancellationToken cancellationToken)
            {
                _owner = owner;
                _cancellationToken = cancellationToken;
            }

            public void Report(ExcelProgress progress)
            {
                if (progress == null) throw new ArgumentNullException(nameof(progress));
                if (!_planningReported &&
                    progress.Stage != ExcelBuildStage.Inspecting &&
                    progress.Stage != ExcelBuildStage.Normalizing)
                {
                    _planningReported = true;
                    _owner.Report(
                        ActivityStage.Planning,
                        ActivityKind.Check,
                        "Validated the managed report plan.",
                        "Rows, Columns, Values, Filters, checks, and owned ranges are bounded.");
                    _owner.PumpDispatcherOnce();
                }

                _owner.ReportExcelProgress(progress, _cancellationToken);
            }
        }
    }
}
