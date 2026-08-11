using System;
using System.Collections.Generic;
using System.Security;
using System.Threading;
using System.Threading.Tasks;
using ExcelReportBuilder.AddIn.Activity;

namespace ExcelReportBuilder.AddIn.Host
{
    /// <summary>
    /// Exercisable shell service for development and installer smoke tests. Every
    /// value is synthetic and no workbook object is read or changed.
    /// </summary>
    public sealed class SyntheticReportBuilderHostService : IReportBuilderHostService
    {
        private volatile bool _isPaused;
        private volatile bool _isCancelled;

        public event EventHandler<HostActivityEventArgs>? ActivityReported;

        public bool IsSynthetic => true;

        public SavedEndpointSettingsSnapshot? SavedEndpointSettings => null;

        public async Task<SourceSnapshot> SelectCurrentDataAsync(CancellationToken cancellationToken)
        {
            ResetControlState();
            Report(
                ActivityStage.Inspecting,
                ActivityKind.Progress,
                "Inspecting the selected workbook data.",
                "Synthetic shell service. No workbook data is read.");
            await DelayWithControlAsync(220, cancellationToken);

            return CreateSampleSource();
        }

        public async Task ConfirmPeriodMappingAsync(
            PeriodMappingSnapshot periodMapping,
            CancellationToken cancellationToken)
        {
            Report(
                ActivityStage.Planning,
                ActivityKind.Progress,
                "Confirming the period layout.",
                $"Mode: {periodMapping.Mode}.");
            await DelayWithControlAsync(160, cancellationToken);
        }

        public async Task<WideHeaderMappingPreview> PreviewWideHeaderMappingAsync(
            string headerPattern,
            int? reportingYear,
            CancellationToken cancellationToken)
        {
            ResetControlState();
            Report(
                ActivityStage.Inspecting,
                ActivityKind.Progress,
                "Mapping wide period headers.",
                "Each source header is classified as a Period and Metric.");
            await DelayWithControlAsync(260, cancellationToken);

            bool requiresYear = !reportingYear.HasValue;
            string yearLabel = reportingYear?.ToString() ?? "year required";
            WideHeaderMappingRowSnapshot[] mappings =
            {
                new WideHeaderMappingRowSnapshot("Jan Amount", $"Jan {yearLabel}", "Amount", 0.98),
                new WideHeaderMappingRowSnapshot("Feb Amount", $"Feb {yearLabel}", "Amount", 0.98),
                new WideHeaderMappingRowSnapshot("Jan Units", $"Jan {yearLabel}", "Units", 0.94),
                new WideHeaderMappingRowSnapshot("Feb Units", $"Feb {yearLabel}", "Units", 0.94)
            };
            NormalizedSampleRowSnapshot[] samples =
            {
                new NormalizedSampleRowSnapshot("Row 2", $"Jan {yearLabel}", "Amount", "1,250.00"),
                new NormalizedSampleRowSnapshot("Row 2", $"Feb {yearLabel}", "Amount", "1,410.00"),
                new NormalizedSampleRowSnapshot("Row 3", $"Jan {yearLabel}", "Units", "18")
            };

            return new WideHeaderMappingPreview(
                mappings,
                projectedNormalizedRowCount: 480,
                samples,
                requiresYear ? TotalPreservationState.NotChecked : TotalPreservationState.Pass,
                requiresYear
                    ? "A reporting year is required before totals can be confirmed."
                    : "Synthetic input and projected normalized totals match.",
                requiresYear);
        }

        public async Task<BuildDraftResult> BuildManagedDraftAsync(
            ReportSpecificationSnapshot specification,
            CancellationToken cancellationToken)
        {
            ResetControlState();
            Report(
                ActivityStage.Inspecting,
                ActivityKind.Progress,
                "Rechecking the synthetic Data shape.",
                "Synthetic shell output only. No workbook object is read.");
            await DelayWithControlAsync(120, cancellationToken);

            Report(
                ActivityStage.Normalizing,
                ActivityKind.Progress,
                "Preparing normalized synthetic rows.",
                "Synthetic shell output only. No workbook object is changed.");
            await DelayWithControlAsync(180, cancellationToken);

            Report(
                ActivityStage.Planning,
                ActivityKind.Progress,
                "Planning rows, columns, values, and filters.",
                $"{specification.Placements.Count} placements in the saved report setup.");
            await DelayWithControlAsync(300, cancellationToken);

            Report(
                ActivityStage.BuildingPivots,
                ActivityKind.Progress,
                "Building synthetic pivot structures.",
                specification.OutputStyle);
            await DelayWithControlAsync(220, cancellationToken);

            Report(
                ActivityStage.Rendering,
                ActivityKind.Progress,
                "Rendering the synthetic managed draft.",
                "Synthetic shell output only. No workbook object is changed.");
            await DelayWithControlAsync(420, cancellationToken);

            Report(
                ActivityStage.Calculating,
                ActivityKind.Progress,
                "Calculating synthetic report values.",
                "No workbook calculation is requested in synthetic mode.");
            await DelayWithControlAsync(160, cancellationToken);

            Report(
                ActivityStage.Checking,
                ActivityKind.Check,
                "Checking the draft structure.",
                "Source preservation and total checks are queued.");
            await DelayWithControlAsync(280, cancellationToken);

            return new BuildDraftResult("Synthetic managed draft", 24);
        }

        public async Task<ChatRunResult> RunChatAsync(
            string request,
            ReportSpecificationSnapshot specification,
            ModelEndpointSettingsSnapshot endpointSettings,
            SecureString? apiKey,
            CancellationToken cancellationToken)
        {
            ResetControlState();
            Report(
                ActivityStage.Planning,
                ActivityKind.Progress,
                "Reviewing the request against the report setup.",
                "Only bounded report-specification actions are available.");
            await DelayWithControlAsync(260, cancellationToken);

            return new ChatRunResult(
                "I prepared a bounded change proposal for the current Rows, Columns, Values, and Filters. Review the Build surface before creating a managed draft.");
        }

        public async Task<IReadOnlyList<string>> DiscoverModelsAsync(
            ModelEndpointSettingsSnapshot endpointSettings,
            SecureString? apiKey,
            CancellationToken cancellationToken)
        {
            ResetControlState();
            Report(
                ActivityStage.Ready,
                ActivityKind.Progress,
                "Discovering available models.",
                "Synthetic endpoint response. No network request is sent.");
            await DelayWithControlAsync(180, cancellationToken);
            return new[] { "sample-small", "sample-balanced" };
        }

        public async Task<EndpointCheckResult> CheckEndpointAsync(
            ModelEndpointSettingsSnapshot endpointSettings,
            SecureString? apiKey,
            CancellationToken cancellationToken)
        {
            ResetControlState();
            Report(
                ActivityStage.Ready,
                ActivityKind.Check,
                "Checking the model endpoint.",
                "Synthetic endpoint response. No network request is sent.");
            await DelayWithControlAsync(180, cancellationToken);
            return new EndpointCheckResult(true, "Synthetic endpoint settings are valid for shell testing.");
        }

        public Task PersistEndpointSettingsAsync(
            ModelEndpointSettingsSnapshot endpointSettings,
            SecureString? apiKey,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.CompletedTask;
        }

        public async Task<IReadOnlyList<HostCheckResult>> RunChecksAsync(CancellationToken cancellationToken)
        {
            ResetControlState();
            Report(
                ActivityStage.Checking,
                ActivityKind.Check,
                "Running independent checks.",
                "Synthetic shell results are clearly identified below.");
            await DelayWithControlAsync(320, cancellationToken);

            return new[]
            {
                new HostCheckResult("Source preserved", true, "Synthetic check passed."),
                new HostCheckResult("Totals reconcile", true, "Synthetic check passed."),
                new HostCheckResult("Managed objects only", true, "Synthetic check passed."),
                new HostCheckResult("Period coverage", true, "Synthetic check passed.")
            };
        }

        public async Task<PublishResult> PublishManagedDraftAsync(CancellationToken cancellationToken)
        {
            ResetControlState();
            Report(
                ActivityStage.Complete,
                ActivityKind.Progress,
                "Publishing the checked managed draft.",
                "Synthetic shell service. No workbook object is changed.");
            await DelayWithControlAsync(180, cancellationToken);
            return new PublishResult("Synthetic managed draft marked as published.");
        }

        public void RequestPause()
        {
            _isPaused = true;
        }

        public void RequestResume()
        {
            _isPaused = false;
        }

        public void RequestCancel()
        {
            _isCancelled = true;
        }

        public static SourceSnapshot CreateSampleSource()
        {
            return new SourceSnapshot(
                "Synthetic preview data",
                "SampleData!A1:D121",
                120,
                new[]
                {
                    new SourceColumnSnapshot("Date", "Date", "2026-01-15"),
                    new SourceColumnSnapshot("Region", "Text", "North"),
                    new SourceColumnSnapshot("Category", "Text", "Service"),
                    new SourceColumnSnapshot("Amount", "Number", "1,250.00")
                },
                isSynthetic: true);
        }

        private async Task DelayWithControlAsync(int milliseconds, CancellationToken cancellationToken)
        {
            int remaining = milliseconds;
            while (remaining > 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (_isCancelled)
                {
                    throw new OperationCanceledException(cancellationToken);
                }

                if (_isPaused)
                {
                    await Task.Delay(50, cancellationToken);
                    continue;
                }

                int slice = Math.Min(50, remaining);
                await Task.Delay(slice, cancellationToken);
                remaining -= slice;
            }
        }

        private void ResetControlState()
        {
            _isCancelled = false;
            _isPaused = false;
        }

        private void Report(
            ActivityStage stage,
            ActivityKind kind,
            string message,
            string detail)
        {
            ActivityReported?.Invoke(this, new HostActivityEventArgs(stage, kind, message, detail));
        }
    }
}
