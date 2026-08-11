using System;
using System.Threading;
using System.Threading.Tasks;
using ExcelReportBuilder.Agent.Security;

namespace ExcelReportBuilder.Worker;

internal static class Program
{
    public static async Task<int> Main(string[] args)
    {
        WorkerOptions options;
        try
        {
            options = WorkerOptions.Parse(args);
        }
        catch (ArgumentException exception)
        {
            Console.Error.WriteLine(exception.Message);
            return 2;
        }

        using var shutdown = new CancellationTokenSource();
        Console.CancelKeyPress += (_, eventArgs) =>
        {
            eventArgs.Cancel = true;
            shutdown.Cancel();
        };

        try
        {
            string handshakeSecret = Environment.GetEnvironmentVariable(
                WorkerHandshakeAuthenticator.SecretEnvironmentVariable) ?? string.Empty;
            Environment.SetEnvironmentVariable(
                WorkerHandshakeAuthenticator.SecretEnvironmentVariable,
                null,
                EnvironmentVariableTarget.Process);
            var server = new AgentWorkerServer(options.PipeName, handshakeSecret);
            handshakeSecret = string.Empty;
            await server.RunAsync(shutdown.Token).ConfigureAwait(false);
            return 0;
        }
        catch (OperationCanceledException) when (shutdown.IsCancellationRequested)
        {
            return 0;
        }
        catch (Exception)
        {
            // Never print exception details. They can contain endpoint or
            // machine information and the host receives typed diagnostics.
            Console.Error.WriteLine("The guarded AI worker stopped unexpectedly.");
            return 1;
        }
    }
}

internal sealed class WorkerOptions
{
    private WorkerOptions(string pipeName)
    {
        PipeName = pipeName;
    }

    public string PipeName { get; }

    public static WorkerOptions Parse(string[] args)
    {
        if (args == null) throw new ArgumentNullException(nameof(args));

        string? pipeName = null;
        for (var index = 0; index < args.Length; index++)
        {
            if (!string.Equals(args[index], "--pipe", StringComparison.Ordinal))
            {
                throw new ArgumentException("Usage: ExcelReportBuilder.Worker [--pipe <pipe-name>]");
            }

            if (++index >= args.Length)
            {
                throw new ArgumentException("A pipe name must follow --pipe.");
            }

            pipeName = args[index];
        }

        if (pipeName == null)
        {
            throw new ArgumentException(
                "Usage: ExcelReportBuilder.Worker --pipe <pipe-name>");
        }

        return new WorkerOptions(PipeNamePolicy.Validate(pipeName));
    }
}
