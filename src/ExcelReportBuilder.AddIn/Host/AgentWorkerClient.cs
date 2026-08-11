using System;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExcelReportBuilder.Agent.Models;
using ExcelReportBuilder.Agent.Protocol;
using ExcelReportBuilder.Agent.Security;
using ExcelReportBuilder.Agent.Tools;
using ExcelReportBuilder.Agent.Validation;

namespace ExcelReportBuilder.AddIn.Host
{
    /// <summary>
    /// A single-job client for the out-of-process worker. The worker receives
    /// bounded serializable snapshots only. Excel COM objects remain in host code.
    /// </summary>
    internal sealed class AgentWorkerClient : IDisposable
    {
        private readonly Func<AgentProgressEvent, CancellationToken, Task> _progress;
        private readonly Func<AgentCheckpointEvent, CancellationToken, Task> _checkpoint;
        private readonly Func<HostToolRequestEvent, CancellationToken, Task<HostToolResultRequest>> _hostTool;
        private Process? _ownedWorker;
        private bool _disposed;

        public AgentWorkerClient(
            Func<AgentProgressEvent, CancellationToken, Task> progress,
            Func<AgentCheckpointEvent, CancellationToken, Task> checkpoint,
            Func<HostToolRequestEvent, CancellationToken, Task<HostToolResultRequest>> hostTool)
        {
            _progress = progress ?? throw new ArgumentNullException(nameof(progress));
            _checkpoint = checkpoint ?? throw new ArgumentNullException(nameof(checkpoint));
            _hostTool = hostTool ?? throw new ArgumentNullException(nameof(hostTool));
        }

        public async Task<AgentRunResult> RunAsync(
            AgentJobRequest request,
            CancellationToken cancellationToken)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            ThrowIfDisposed();
            using (var pipe = await ConnectAuthenticatedAsync(
                cancellationToken).ConfigureAwait(false))
            {
                await ApplyResumeMetadataAsync(pipe, request, cancellationToken).ConfigureAwait(false);
                await PipeJsonProtocol.WriteAsync(
                    pipe,
                    AgentProtocol.Create(
                        AgentMessageType.StartJob,
                        request.JobId,
                        new StartJobRequest { Job = request }),
                    cancellationToken).ConfigureAwait(false);

                try
                {
                    return await ReadJobAsync(pipe, request, cancellationToken).ConfigureAwait(false);
                }
                catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
                {
                    await TryCancelAsync(pipe, request.JobId).ConfigureAwait(false);
                    throw;
                }
            }
        }

        private async Task<NamedPipeClientStream> ConnectAuthenticatedAsync(
            CancellationToken cancellationToken)
        {
            string pipeName = WorkerHandshakeAuthenticator.CreatePipeName();
            string handshakeSecret = WorkerHandshakeAuthenticator.CreateSecret();
            NamedPipeClientStream? pipe = null;
            try
            {
                pipe = await ConnectAsync(
                    pipeName,
                    handshakeSecret,
                    cancellationToken).ConfigureAwait(false);
                await HandshakeAsync(
                    pipe,
                    pipeName,
                    handshakeSecret,
                    cancellationToken).ConfigureAwait(false);
                return pipe;
            }
            catch
            {
                pipe?.Dispose();
                throw;
            }
            finally
            {
                handshakeSecret = string.Empty;
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            if (_ownedWorker != null)
            {
                try
                {
                    if (!_ownedWorker.HasExited)
                    {
                        _ownedWorker.Kill();
                        _ownedWorker.WaitForExit(2000);
                    }
                }
                catch (InvalidOperationException)
                {
                }
                catch (System.ComponentModel.Win32Exception)
                {
                }
                finally
                {
                    _ownedWorker.Dispose();
                    _ownedWorker = null;
                }
            }
        }

        private async Task<NamedPipeClientStream> ConnectAsync(
            string pipeName,
            string handshakeSecret,
            CancellationToken cancellationToken)
        {
            string workerPath = FindWorkerPath();
            var startInfo = new ProcessStartInfo
            {
                FileName = workerPath,
                Arguments = "--pipe \"" + pipeName + "\"",
                WorkingDirectory = Path.GetDirectoryName(workerPath) ?? AppDomain.CurrentDomain.BaseDirectory,
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };
            startInfo.EnvironmentVariables[
                WorkerHandshakeAuthenticator.SecretEnvironmentVariable] = handshakeSecret;
            try
            {
                _ownedWorker = Process.Start(startInfo)
                    ?? throw new InvalidOperationException("The guarded AI worker could not be started.");
            }
            finally
            {
                startInfo.EnvironmentVariables.Remove(
                    WorkerHandshakeAuthenticator.SecretEnvironmentVariable);
            }

            var pipe = CreatePipe(pipeName);
            try
            {
                await pipe.ConnectAsync(10000, cancellationToken).ConfigureAwait(false);
                return pipe;
            }
            catch
            {
                pipe.Dispose();
                throw;
            }
        }

        private static NamedPipeClientStream CreatePipe(string pipeName)
        {
            return new NamedPipeClientStream(
                ".",
                pipeName,
                PipeDirection.InOut,
                PipeOptions.Asynchronous);
        }

        private static async Task HandshakeAsync(
            Stream pipe,
            string pipeName,
            string handshakeSecret,
            CancellationToken cancellationToken)
        {
            string correlationId = "hello_" + Guid.NewGuid().ToString("N");
            string clientNonce = WorkerHandshakeAuthenticator.CreateNonce();
            await PipeJsonProtocol.WriteAsync(
                pipe,
                AgentProtocol.Create(
                    AgentMessageType.Hello,
                    correlationId,
                    new HelloRequest
                    {
                        ClientName = "ExcelReportBuilder.AddIn",
                        SupportedProtocolVersions = new System.Collections.Generic.List<string>
                        {
                            AgentProtocol.Version
                        },
                        ClientNonce = clientNonce
                    }),
                cancellationToken).ConfigureAwait(false);

            AgentProtocolEnvelope? response = await PipeJsonProtocol.ReadAsync(
                pipe,
                cancellationToken).ConfigureAwait(false);
            if (response == null ||
                response.MessageType != AgentMessageType.HelloAcknowledged ||
                !string.Equals(response.CorrelationId, correlationId, StringComparison.Ordinal))
            {
                throw new AgentProtocolException("The worker handshake was not acknowledged.");
            }

            HelloAcknowledgement acknowledgement =
                AgentProtocol.ReadPayload<HelloAcknowledgement>(response);
            if (!string.Equals(
                acknowledgement.ProtocolVersion,
                AgentProtocol.Version,
                StringComparison.Ordinal))
            {
                throw new AgentProtocolException("The worker selected an unsupported protocol version.");
            }

            if (Environment.OSVersion.Platform == PlatformID.Win32NT &&
                !acknowledgement.CurrentUserOnlyPipe)
            {
                throw new AgentProtocolException(
                    "The worker did not create a current-user-only connection.");
            }

            if (!WorkerHandshakeAuthenticator.VerifyAuthenticationTag(
                    handshakeSecret,
                    pipeName,
                    clientNonce,
                    acknowledgement.ProtocolVersion,
                    acknowledgement.AuthenticationTag))
            {
                throw new AgentProtocolException(
                    "The worker did not prove that it was launched by this add-in session.");
            }
        }

        private async Task ApplyResumeMetadataAsync(
            Stream pipe,
            AgentJobRequest request,
            CancellationToken cancellationToken)
        {
            string correlationId = "resume_" + Guid.NewGuid().ToString("N");
            await PipeJsonProtocol.WriteAsync(
                pipe,
                AgentProtocol.Create(
                    AgentMessageType.ListResumeMetadata,
                    correlationId,
                    new ResumeMetadataRequest { WorkbookId = request.WorkbookId }),
                cancellationToken).ConfigureAwait(false);

            AgentProtocolEnvelope? response = await PipeJsonProtocol.ReadAsync(
                pipe,
                cancellationToken).ConfigureAwait(false);
            if (response == null ||
                response.MessageType != AgentMessageType.ResumeMetadata ||
                !string.Equals(response.CorrelationId, correlationId, StringComparison.Ordinal))
            {
                throw new AgentProtocolException("The worker did not return valid resume metadata.");
            }

            ResumeMetadataResponse metadata =
                AgentProtocol.ReadPayload<ResumeMetadataResponse>(response);
            AgentResumeMetadata? resumable = metadata.Jobs
                .Where(item => item.CanResume &&
                    string.Equals(item.JobId, request.JobId, StringComparison.Ordinal) &&
                    string.Equals(item.WorkbookId, request.WorkbookId, StringComparison.Ordinal))
                .OrderByDescending(item => item.UpdatedAtUtc)
                .FirstOrDefault();
            if (resumable == null)
            {
                return;
            }

            request.ResumeFromCheckpointId = resumable.CheckpointId;
            await _progress(
                new AgentProgressEvent
                {
                    JobId = request.JobId,
                    Sequence = 0,
                    Stage = AgentProgressStage.Accepted,
                    Message = "Resuming the last checkpoint for this exact workbook request.",
                    OccurredAtUtc = DateTimeOffset.UtcNow
                },
                cancellationToken).ConfigureAwait(false);
        }

        private async Task<AgentRunResult> ReadJobAsync(
            Stream pipe,
            AgentJobRequest request,
            CancellationToken cancellationToken)
        {
            while (true)
            {
                AgentProtocolEnvelope? envelope = await PipeJsonProtocol.ReadAsync(
                    pipe,
                    cancellationToken).ConfigureAwait(false);
                if (envelope == null)
                {
                    throw new AgentProtocolException("The worker connection ended before the job completed.");
                }

                switch (envelope.MessageType)
                {
                    case AgentMessageType.Progress:
                        await _progress(
                            AgentProtocol.ReadPayload<AgentProgressEvent>(envelope),
                            cancellationToken).ConfigureAwait(false);
                        break;
                    case AgentMessageType.Checkpoint:
                        await _checkpoint(
                            AgentProtocol.ReadPayload<AgentCheckpointEvent>(envelope),
                            cancellationToken).ConfigureAwait(false);
                        break;
                    case AgentMessageType.HostToolRequest:
                        HostToolRequestEvent hostRequest =
                            AgentProtocol.ReadPayload<HostToolRequestEvent>(envelope);
                        HostToolResultRequest hostResult = await InvokeHostToolAsync(
                            request,
                            hostRequest,
                            cancellationToken).ConfigureAwait(false);
                        HostToolResultValidator.Validate(hostRequest, hostResult);
                        await PipeJsonProtocol.WriteAsync(
                            pipe,
                            AgentProtocol.Create(
                                AgentMessageType.HostToolResult,
                                envelope.CorrelationId,
                                hostResult),
                            cancellationToken).ConfigureAwait(false);
                        break;
                    case AgentMessageType.JobCompleted:
                        JobCompletedEvent completed =
                            AgentProtocol.ReadPayload<JobCompletedEvent>(envelope);
                        if (!string.Equals(completed.Result.JobId, request.JobId, StringComparison.Ordinal) ||
                            !string.Equals(completed.Result.WorkbookId, request.WorkbookId, StringComparison.Ordinal))
                        {
                            throw new AgentProtocolException("The worker returned a mismatched job result.");
                        }

                        return completed.Result;
                    case AgentMessageType.JobFailed:
                        JobFailedEvent failed = AgentProtocol.ReadPayload<JobFailedEvent>(envelope);
                        throw new AgentWorkerException(
                            failed.Diagnostic.Code,
                            failed.Diagnostic.Message,
                            failed.Diagnostic.Retryable);
                    case AgentMessageType.JobCancelled:
                        throw new OperationCanceledException(
                            AgentProtocol.ReadPayload<JobCancelledEvent>(envelope).Message,
                            cancellationToken);
                    case AgentMessageType.Error:
                        ProtocolErrorEvent error = AgentProtocol.ReadPayload<ProtocolErrorEvent>(envelope);
                        throw new AgentWorkerException(error.Code, error.Message, false);
                    default:
                        throw new AgentProtocolException(
                            "The worker returned a message that is not valid during a job.");
                }
            }
        }

        private async Task<HostToolResultRequest> InvokeHostToolAsync(
            AgentJobRequest job,
            HostToolRequestEvent request,
            CancellationToken cancellationToken)
        {
            if (!string.Equals(request.JobId, job.JobId, StringComparison.Ordinal) ||
                !string.Equals(request.WorkbookId, job.WorkbookId, StringComparison.Ordinal))
            {
                return Failure(request, "host_tool_identity_mismatch");
            }

            if (!AgentToolCatalog.IsAllowed(request.ToolName))
            {
                return Failure(request, "host_tool_not_allowed");
            }

            HostToolResultRequest result = await _hostTool(request, cancellationToken)
                .ConfigureAwait(false);
            return result ?? Failure(request, "host_tool_result_missing");
        }

        private static HostToolResultRequest Failure(
            HostToolRequestEvent request,
            string outcomeCode)
        {
            return new HostToolResultRequest
            {
                JobId = request.JobId,
                ToolCallId = request.ToolCallId,
                Succeeded = false,
                OutcomeCode = outcomeCode,
                ResultJson = "{}"
            };
        }

        private static async Task TryCancelAsync(Stream pipe, string jobId)
        {
            using (var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(1)))
            {
                try
                {
                    await PipeJsonProtocol.WriteAsync(
                        pipe,
                        AgentProtocol.Create(
                            AgentMessageType.CancelJob,
                            "cancel_" + Guid.NewGuid().ToString("N"),
                            new CancelJobRequest { JobId = jobId }),
                        timeout.Token).ConfigureAwait(false);
                }
                catch (Exception exception) when (
                    exception is IOException ||
                    exception is OperationCanceledException ||
                    exception is ObjectDisposedException)
                {
                }
            }
        }

        private static string FindWorkerPath()
        {
            string assemblyDirectory = Path.GetDirectoryName(
                typeof(AgentWorkerClient).Assembly.Location)
                ?? AppDomain.CurrentDomain.BaseDirectory;
            string[] candidates =
            {
                Path.Combine(assemblyDirectory, "worker", "ExcelReportBuilder.Worker.exe"),
                Path.Combine(assemblyDirectory, "ExcelReportBuilder.Worker.exe")
            };
            foreach (string candidate in candidates)
            {
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }

            throw new FileNotFoundException(
                "The guarded AI worker is not installed beside the add-in.");
        }

        private void ThrowIfDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(AgentWorkerClient));
            }
        }
    }

    internal sealed class AgentWorkerException : Exception
    {
        public AgentWorkerException(string code, string message, bool retryable)
            : base(message)
        {
            Code = code ?? "worker_error";
            Retryable = retryable;
        }

        public string Code { get; }

        public bool Retryable { get; }
    }
}
