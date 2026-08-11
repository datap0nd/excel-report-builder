using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using ExcelReportBuilder.Agent.Configuration;
using ExcelReportBuilder.Agent.Diagnostics;
using ExcelReportBuilder.Agent.Execution;
using ExcelReportBuilder.Agent.Models;
using ExcelReportBuilder.Agent.OpenAI;
using ExcelReportBuilder.Agent.Protocol;
using ExcelReportBuilder.Agent.Validation;

namespace ExcelReportBuilder.Worker;

internal sealed class AgentWorkerServer
{
    private const int MaximumConcurrentJobs = 2;
    private readonly string _pipeName;
    private readonly IJobCheckpointStore _checkpointStore;
    private readonly ActiveWorkbookJobRegistry _activeWorkbooks;

    public AgentWorkerServer(
        string pipeName,
        IJobCheckpointStore? checkpointStore = null,
        ActiveWorkbookJobRegistry? activeWorkbooks = null)
    {
        _pipeName = PipeNamePolicy.Validate(pipeName);
        _checkpointStore = checkpointStore ?? new LocalAppDataJobCheckpointStore();
        _activeWorkbooks = activeWorkbooks ?? new ActiveWorkbookJobRegistry();
    }

    public async Task RunAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            bool currentUserOnly;
            using var server = CreateServerStream(out currentUserOnly);
            await server.WaitForConnectionAsync(cancellationToken).ConfigureAwait(false);
            await HandleConnectionAsync(server, currentUserOnly, cancellationToken).ConfigureAwait(false);
        }
    }

    private NamedPipeServerStream CreateServerStream(out bool currentUserOnly)
    {
        try
        {
            currentUserOnly = true;
            return new NamedPipeServerStream(
                _pipeName,
                PipeDirection.InOut,
                1,
                PipeTransmissionMode.Byte,
                PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly,
                AgentProtocol.MaximumFrameBytes + 4,
                AgentProtocol.MaximumFrameBytes + 4);
        }
        catch (PlatformNotSupportedException) when (!OperatingSystem.IsWindows())
        {
            currentUserOnly = false;
            return new NamedPipeServerStream(
                _pipeName,
                PipeDirection.InOut,
                1,
                PipeTransmissionMode.Byte,
                PipeOptions.Asynchronous,
                AgentProtocol.MaximumFrameBytes + 4,
                AgentProtocol.MaximumFrameBytes + 4);
        }
    }

    private async Task HandleConnectionAsync(
        NamedPipeServerStream pipe,
        bool currentUserOnly,
        CancellationToken workerCancellation)
    {
        using var connectionCancellation = CancellationTokenSource.CreateLinkedTokenSource(workerCancellation);
        using var writer = new PipeConnectionWriter(pipe, connectionCancellation.Token);
        var hostToolBridge = new ProtocolHostToolBridge(writer);
        using var jobSlots = new SemaphoreSlim(MaximumConcurrentJobs, MaximumConcurrentJobs);
        var jobs = new ConcurrentDictionary<string, JobHandle>(StringComparer.Ordinal);
        var backgroundTasks = new ConcurrentBag<Task>();
        var handshakeComplete = false;

        try
        {
            while (!connectionCancellation.IsCancellationRequested && pipe.IsConnected)
            {
                AgentProtocolEnvelope? envelope;
                try
                {
                    envelope = await PipeJsonProtocol.ReadAsync(pipe, connectionCancellation.Token).ConfigureAwait(false);
                }
                catch (AgentProtocolException exception)
                {
                    await TrySendProtocolErrorAsync(
                        writer,
                        "protocol_error",
                        exception.Message,
                        connectionCancellation.Token).ConfigureAwait(false);
                    break;
                }

                if (envelope == null) break;

                try
                {
                    if (!handshakeComplete)
                    {
                        if (envelope.MessageType != AgentMessageType.Hello)
                        {
                            await TrySendProtocolErrorAsync(
                                writer,
                                "handshake_required",
                                "Complete the protocol handshake before sending worker commands.",
                                connectionCancellation.Token).ConfigureAwait(false);
                            continue;
                        }

                        var hello = AgentProtocol.ReadPayload<HelloRequest>(envelope);
                        if (hello.SupportedProtocolVersions == null ||
                            !hello.SupportedProtocolVersions.Contains(AgentProtocol.Version, StringComparer.Ordinal))
                        {
                            await TrySendProtocolErrorAsync(
                                writer,
                                "protocol_version_unsupported",
                                "The client does not support this worker protocol version.",
                                connectionCancellation.Token).ConfigureAwait(false);
                            break;
                        }

                        handshakeComplete = true;
                        var version = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "1.0.0";
                        await writer.SendAsync(
                            AgentMessageType.HelloAcknowledged,
                            envelope.CorrelationId,
                            new HelloAcknowledgement
                            {
                                ProtocolVersion = AgentProtocol.Version,
                                WorkerVersion = version,
                                CurrentUserOnlyPipe = currentUserOnly,
                            },
                            connectionCancellation.Token).ConfigureAwait(false);
                        continue;
                    }

                    switch (envelope.MessageType)
                    {
                        case AgentMessageType.StartJob:
                        {
                            var start = AgentProtocol.ReadPayload<StartJobRequest>(envelope);
                            if (start.Job == null ||
                                string.IsNullOrWhiteSpace(start.Job.JobId) ||
                                string.IsNullOrWhiteSpace(start.Job.WorkbookId))
                            {
                                await TrySendProtocolErrorAsync(writer, "job_invalid", "The job request is invalid.", connectionCancellation.Token).ConfigureAwait(false);
                                break;
                            }

                            try
                            {
                                AgentRequestValidator.ValidateIdentity(start.Job);
                            }
                            catch (AgentInputValidationException exception)
                            {
                                await TrySendProtocolErrorAsync(
                                    writer,
                                    exception.Code,
                                    exception.Message,
                                    connectionCancellation.Token).ConfigureAwait(false);
                                break;
                            }

                            if (!jobSlots.Wait(0))
                            {
                                await TrySendProtocolErrorAsync(writer, "worker_busy", "The worker is already running the maximum number of jobs.", connectionCancellation.Token).ConfigureAwait(false);
                                break;
                            }

                            var jobCancellation = CancellationTokenSource.CreateLinkedTokenSource(connectionCancellation.Token);
                            var handle = new JobHandle(jobCancellation);
                            if (!jobs.TryAdd(start.Job.JobId, handle))
                            {
                                jobCancellation.Dispose();
                                jobSlots.Release();
                                await TrySendProtocolErrorAsync(writer, "job_already_running", "A job with this ID is already running.", connectionCancellation.Token).ConfigureAwait(false);
                                break;
                            }

                            if (!_activeWorkbooks.TryStart(start.Job.WorkbookId, start.Job.JobId))
                            {
                                jobs.TryRemove(start.Job.JobId, out _);
                                jobCancellation.Dispose();
                                jobSlots.Release();
                                await TrySendProtocolErrorAsync(
                                    writer,
                                    "workbook_job_already_running",
                                    "Only one AI job can run for a workbook at a time.",
                                    connectionCancellation.Token).ConfigureAwait(false);
                                break;
                            }

                            handle.Task = RunJobAsync(
                                start.Job,
                                writer,
                                hostToolBridge,
                                handle,
                                jobs,
                                jobSlots,
                                connectionCancellation.Token);
                            backgroundTasks.Add(handle.Task);
                            break;
                        }
                        case AgentMessageType.CancelJob:
                        {
                            var cancel = AgentProtocol.ReadPayload<CancelJobRequest>(envelope);
                            if (string.IsNullOrWhiteSpace(cancel.JobId))
                            {
                                await TrySendProtocolErrorAsync(writer, "job_id_invalid", "A job ID is required for cancellation.", connectionCancellation.Token).ConfigureAwait(false);
                                break;
                            }

                            JobHandle? handle;
                            var requested = jobs.TryGetValue(cancel.JobId, out handle);
                            if (requested)
                            {
                                handle!.Cancellation.Cancel();
                            }

                            await writer.SendAsync(
                                AgentMessageType.CancelAcknowledged,
                                envelope.CorrelationId,
                                new CancelJobAcknowledgement
                                {
                                    JobId = cancel.JobId,
                                    CancellationRequested = requested,
                                },
                                connectionCancellation.Token).ConfigureAwait(false);
                            break;
                        }
                        case AgentMessageType.ProbeEndpoint:
                        {
                            var probe = AgentProtocol.ReadPayload<EndpointProbeRequest>(envelope);
                            var task = RunProbeAsync(
                                probe,
                                envelope.CorrelationId,
                                writer,
                                connectionCancellation.Token);
                            backgroundTasks.Add(task);
                            break;
                        }
                        case AgentMessageType.HostToolResult:
                        {
                            var result = AgentProtocol.ReadPayload<HostToolResultRequest>(envelope);
                            if (!hostToolBridge.TryComplete(result))
                            {
                                await TrySendProtocolErrorAsync(
                                    writer,
                                    "host_tool_result_unmatched",
                                    "The host tool result does not match a pending request.",
                                    connectionCancellation.Token).ConfigureAwait(false);
                            }

                            break;
                        }
                        case AgentMessageType.ListResumeMetadata:
                        {
                            var request = AgentProtocol.ReadPayload<ResumeMetadataRequest>(envelope);
                            var metadata = await _checkpointStore.ListAsync(
                                request.WorkbookId,
                                connectionCancellation.Token).ConfigureAwait(false);
                            await writer.SendAsync(
                                AgentMessageType.ResumeMetadata,
                                envelope.CorrelationId,
                                new ResumeMetadataResponse { Jobs = metadata.ToList() },
                                connectionCancellation.Token).ConfigureAwait(false);
                            break;
                        }
                        default:
                            await TrySendProtocolErrorAsync(
                                writer,
                                "message_not_allowed",
                                "This message type is not accepted from the worker client.",
                                connectionCancellation.Token).ConfigureAwait(false);
                            break;
                    }
                }
                catch (AgentProtocolException exception)
                {
                    await TrySendProtocolErrorAsync(
                            writer,
                            "payload_invalid",
                            exception.Message,
                            connectionCancellation.Token).ConfigureAwait(false);
                }
            }
        }
        finally
        {
            connectionCancellation.Cancel();
            hostToolBridge.CancelAll(connectionCancellation.Token);
            foreach (var job in jobs.Values)
            {
                job.Cancellation.Cancel();
            }

            try
            {
                await Task.WhenAll(backgroundTasks.ToArray()).ConfigureAwait(false);
            }
            catch (Exception)
            {
                // Individual background operations already emit typed,
                // redacted outcomes when the connection remains available.
            }
        }
    }

    private async Task RunJobAsync(
        AgentJobRequest request,
        PipeConnectionWriter writer,
        ProtocolHostToolBridge hostToolBridge,
        JobHandle handle,
        ConcurrentDictionary<string, JobHandle> jobs,
        SemaphoreSlim jobSlots,
        CancellationToken connectionCancellation)
    {
        try
        {
            if (!string.IsNullOrWhiteSpace(request.ResumeFromCheckpointId))
            {
                var resume = await _checkpointStore.LoadAsync(
                    request.JobId,
                    request.WorkbookId,
                    handle.Cancellation.Token).ConfigureAwait(false);
                if (resume == null || !resume.CanResume ||
                    !string.Equals(resume.CheckpointId, request.ResumeFromCheckpointId, StringComparison.Ordinal))
                {
                    throw new AgentInputValidationException(
                        "resume_checkpoint_unavailable",
                        "The requested resume checkpoint is not available for this job and workbook.");
                }
            }

            using var client = new OpenAiCompatibleClient();
            var eventSink = new PersistingAgentEventSink(writer, _checkpointStore);
            var runner = new GuardedAgentRunner(
                client,
                eventSink,
                hostToolBridge: hostToolBridge);
            var result = await runner.RunAsync(request, handle.Cancellation.Token).ConfigureAwait(false);
            await _checkpointStore.DeleteAsync(
                request.JobId,
                request.WorkbookId,
                connectionCancellation).ConfigureAwait(false);
            await TrySendAsync(
                writer,
                AgentMessageType.JobCompleted,
                request.JobId,
                new JobCompletedEvent { Result = result },
                connectionCancellation).ConfigureAwait(false);
        }
        catch (OperationCanceledException) when (handle.Cancellation.IsCancellationRequested)
        {
            await TryMarkFailureAsync(
                request.JobId,
                request.WorkbookId,
                "job_cancelled",
                CancellationToken.None).ConfigureAwait(false);
            await TrySendAsync(
                writer,
                AgentMessageType.JobCancelled,
                request.JobId,
                new JobCancelledEvent { JobId = request.JobId },
                connectionCancellation).ConfigureAwait(false);
        }
        catch (AgentInputValidationException exception)
        {
            await SendFailureAsync(writer, request, exception.Code, exception, false, connectionCancellation).ConfigureAwait(false);
        }
        catch (AgentEndpointPolicyException exception)
        {
            await SendFailureAsync(writer, request, "endpoint_policy_rejected", exception, false, connectionCancellation).ConfigureAwait(false);
        }
        catch (AgentEndpointException exception)
        {
            await SendFailureAsync(writer, request, exception.Code, exception, exception.Retryable, connectionCancellation).ConfigureAwait(false);
        }
        catch (AgentRunException exception)
        {
            await SendFailureAsync(writer, request, exception.Code, exception, exception.Retryable, connectionCancellation).ConfigureAwait(false);
        }
        catch (Exception)
        {
            var resume = await TryMarkFailureAsync(
                request.JobId,
                request.WorkbookId,
                "unexpected_worker_error",
                CancellationToken.None).ConfigureAwait(false);
            await TrySendAsync(
                writer,
                AgentMessageType.JobFailed,
                request.JobId,
                new JobFailedEvent
                {
                    JobId = request.JobId,
                    Diagnostic = new RedactedDiagnostic
                    {
                        Code = "unexpected_worker_error",
                        Message = "The worker could not complete the guarded request.",
                        Retryable = false,
                    },
                    ResumeMetadata = resume,
                },
                connectionCancellation).ConfigureAwait(false);
        }
        finally
        {
            jobs.TryRemove(request.JobId, out _);
            _activeWorkbooks.Complete(request.WorkbookId, request.JobId);
            handle.Cancellation.Dispose();
            jobSlots.Release();
        }
    }

    private static async Task RunProbeAsync(
        EndpointProbeRequest request,
        string correlationId,
        PipeConnectionWriter writer,
        CancellationToken cancellationToken)
    {
        try
        {
            using var client = new OpenAiCompatibleClient();
            long sequence = 1;
            await writer.PublishProgressAsync(
                new AgentProgressEvent
                {
                    JobId = correlationId,
                    Sequence = sequence,
                    Stage = AgentProgressStage.DiscoveringModels,
                    Message = "Checking model discovery, structured output, and tool calling with synthetic data.",
                    OccurredAtUtc = DateTimeOffset.UtcNow,
                },
                cancellationToken).ConfigureAwait(false);
            var probe = client.CheckToolCallingAsync(request.Endpoint, cancellationToken);
            while (!probe.IsCompleted)
            {
                var delay = Task.Delay(TimeSpan.FromSeconds(5), cancellationToken);
                if (await Task.WhenAny(probe, delay).ConfigureAwait(false) == probe) break;
                cancellationToken.ThrowIfCancellationRequested();
                await writer.PublishProgressAsync(
                    new AgentProgressEvent
                    {
                        JobId = correlationId,
                        Sequence = ++sequence,
                        Stage = AgentProgressStage.DiscoveringModels,
                        Message = "Still waiting for the synthetic endpoint capability checks.",
                        OccurredAtUtc = DateTimeOffset.UtcNow,
                    },
                    cancellationToken).ConfigureAwait(false);
            }

            var result = await probe.ConfigureAwait(false);
            await TrySendAsync(
                writer,
                AgentMessageType.ProbeCompleted,
                correlationId,
                result,
                cancellationToken).ConfigureAwait(false);
        }
        catch (AgentEndpointException exception)
        {
            var diagnostic = DiagnosticRedactor.FromException(
                exception.Code,
                exception,
                exception.Retryable,
                new[] { request.Endpoint.ApiKey });
            await TrySendAsync(
                writer,
                AgentMessageType.Error,
                correlationId,
                new ProtocolErrorEvent { Code = diagnostic.Code, Message = diagnostic.Message },
                cancellationToken).ConfigureAwait(false);
        }
        catch (Exception)
        {
            await TrySendProtocolErrorAsync(
                writer,
                "endpoint_probe_failed",
                "The endpoint check could not be completed.",
                cancellationToken).ConfigureAwait(false);
        }
    }

    private async Task SendFailureAsync(
        PipeConnectionWriter writer,
        AgentJobRequest request,
        string code,
        Exception exception,
        bool retryable,
        CancellationToken cancellationToken)
    {
        var diagnostic = DiagnosticRedactor.FromException(
            code,
            exception,
            retryable,
            new[] { request.Endpoint.ApiKey, request.UserPrompt });
        AgentResumeMetadata? resume = null;
        if (!string.Equals(code, "job_id_invalid", StringComparison.Ordinal) &&
            !string.Equals(code, "workbook_id_invalid", StringComparison.Ordinal))
        {
            resume = await TryMarkFailureAsync(
                request.JobId,
                request.WorkbookId,
                code,
                CancellationToken.None).ConfigureAwait(false);
        }
        await TrySendAsync(
            writer,
            AgentMessageType.JobFailed,
            request.JobId,
            new JobFailedEvent
            {
                JobId = request.JobId,
                Diagnostic = diagnostic,
                ResumeMetadata = resume,
            },
            cancellationToken).ConfigureAwait(false);
    }

    private async Task<AgentResumeMetadata?> TryMarkFailureAsync(
        string jobId,
        string workbookId,
        string failureCode,
        CancellationToken cancellationToken)
    {
        try
        {
            return await _checkpointStore.MarkFailureAsync(
                jobId,
                workbookId,
                failureCode,
                cancellationToken).ConfigureAwait(false);
        }
        catch (IOException)
        {
            return null;
        }
        catch (UnauthorizedAccessException)
        {
            return null;
        }
    }

    private static Task TrySendProtocolErrorAsync(
        PipeConnectionWriter writer,
        string code,
        string message,
        CancellationToken cancellationToken)
    {
        return TrySendAsync(
            writer,
            AgentMessageType.Error,
            Guid.NewGuid().ToString("N"),
            new ProtocolErrorEvent
            {
                Code = code,
                Message = DiagnosticRedactor.Redact(message),
            },
            cancellationToken);
    }

    private static async Task TrySendAsync<TPayload>(
        PipeConnectionWriter writer,
        AgentMessageType messageType,
        string correlationId,
        TPayload payload,
        CancellationToken cancellationToken)
    {
        try
        {
            await writer.SendAsync(messageType, correlationId, payload, cancellationToken).ConfigureAwait(false);
        }
        catch (IOException)
        {
        }
        catch (OperationCanceledException)
        {
        }
    }

    private sealed class JobHandle
    {
        public JobHandle(CancellationTokenSource cancellation)
        {
            Cancellation = cancellation;
            Task = Task.CompletedTask;
        }

        public CancellationTokenSource Cancellation { get; }

        public Task Task { get; set; }
    }
}
