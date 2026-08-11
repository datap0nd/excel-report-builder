using System.Text;
using ExcelReportBuilder.Agent.Models;
using ExcelReportBuilder.Agent.Protocol;

namespace ExcelReportBuilder.Agent.Tests;

public sealed class ProtocolTests
{
    [Fact]
    public async Task LengthPrefixedProtocol_RoundTripsTypedPayload()
    {
        var envelope = AgentProtocol.Create(
            AgentMessageType.CancelJob,
            "correlation-1",
            new CancelJobRequest { JobId = "job-1" });
        using var stream = new MemoryStream();

        await PipeJsonProtocol.WriteAsync(stream, envelope, CancellationToken.None);
        stream.Position = 0;
        var restored = await PipeJsonProtocol.ReadAsync(stream, CancellationToken.None);

        Assert.NotNull(restored);
        Assert.Equal(AgentProtocol.Version, restored.ProtocolVersion);
        Assert.Equal(AgentMessageType.CancelJob, restored.MessageType);
        Assert.Equal("job-1", AgentProtocol.ReadPayload<CancelJobRequest>(restored).JobId);
    }

    [Fact]
    public void Deserialize_RejectsOtherProtocolVersion()
    {
        var bytes = Encoding.UTF8.GetBytes(
            "{\"protocolVersion\":\"2.0\",\"messageType\":\"hello\",\"correlationId\":\"c1\",\"payload\":{}}");

        var error = Assert.Throws<AgentProtocolException>(() => AgentProtocol.Deserialize(bytes));

        Assert.Contains(AgentProtocol.Version, error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Deserialize_RejectsMissingMessageType()
    {
        var bytes = Encoding.UTF8.GetBytes(
            "{\"protocolVersion\":\"" + AgentProtocol.Version +
            "\",\"correlationId\":\"c1\",\"payload\":{}}");

        Assert.Throws<AgentProtocolException>(() => AgentProtocol.Deserialize(bytes));
    }

    [Fact]
    public void Deserialize_RejectsNumericMessageType()
    {
        var bytes = Encoding.UTF8.GetBytes(
            "{\"protocolVersion\":\"" + AgentProtocol.Version +
            "\",\"messageType\":1,\"correlationId\":\"c1\",\"payload\":{}}");

        Assert.Throws<AgentProtocolException>(() => AgentProtocol.Deserialize(bytes));
    }

    [Fact]
    public void Protocol_RoundTripsHostToolResultForPendingJob()
    {
        var envelope = AgentProtocol.Create(
            AgentMessageType.HostToolResult,
            "host-result-1",
            new HostToolResultRequest
            {
                JobId = "job-1",
                ToolCallId = "checks-1",
                Succeeded = false,
                OutcomeCode = "checks_failed",
                ResultJson = "{}",
                CheckFailures =
                {
                    new HostCheckFailure { Code = "totals_failed", Message = "Synthetic totals differed." },
                },
            });

        var restored = AgentProtocol.ReadPayload<HostToolResultRequest>(
            AgentProtocol.Deserialize(AgentProtocol.Serialize(envelope)));

        Assert.Equal("checks-1", restored.ToolCallId);
        Assert.Single(restored.CheckFailures);
    }

    [Fact]
    public void Protocol_RoundTripsAuthenticatedHello()
    {
        var envelope = AgentProtocol.Create(
            AgentMessageType.Hello,
            "hello-1",
            new HelloRequest
            {
                ClientName = "test-client",
                SupportedProtocolVersions = { AgentProtocol.Version },
                ClientNonce = "nonce-value"
            });

        var restored = AgentProtocol.ReadPayload<HelloRequest>(
            AgentProtocol.Deserialize(AgentProtocol.Serialize(envelope)));

        Assert.Equal("test-client", restored.ClientName);
        Assert.Equal(AgentProtocol.Version, Assert.Single(restored.SupportedProtocolVersions));
        Assert.Equal("nonce-value", restored.ClientNonce);
    }

    [Fact]
    public async Task ReadAsync_RejectsOversizedFrameBeforeAllocation()
    {
        var length = AgentProtocol.MaximumFrameBytes + 1;
        using var stream = new MemoryStream(new[]
        {
            (byte)length,
            (byte)(length >> 8),
            (byte)(length >> 16),
            (byte)(length >> 24),
        });

        await Assert.ThrowsAsync<AgentProtocolException>(() =>
            PipeJsonProtocol.ReadAsync(stream, CancellationToken.None));
    }

    [Fact]
    public async Task ReadAsync_HandlesFragmentedReads()
    {
        var envelope = AgentProtocol.Create(
            AgentMessageType.CancelJob,
            "correlation-fragmented",
            new CancelJobRequest { JobId = "job-fragmented" });
        using var encoded = new MemoryStream();
        await PipeJsonProtocol.WriteAsync(encoded, envelope, CancellationToken.None);
        using var fragmented = new OneByteReadStream(encoded.ToArray());

        var restored = await PipeJsonProtocol.ReadAsync(fragmented, CancellationToken.None);

        Assert.Equal("job-fragmented", AgentProtocol.ReadPayload<CancelJobRequest>(restored!).JobId);
    }

    private sealed class OneByteReadStream : MemoryStream
    {
        public OneByteReadStream(byte[] bytes)
            : base(bytes)
        {
        }

        public override Task<int> ReadAsync(
            byte[] buffer,
            int offset,
            int count,
            CancellationToken cancellationToken)
        {
            return base.ReadAsync(buffer, offset, Math.Min(1, count), cancellationToken);
        }
    }
}
