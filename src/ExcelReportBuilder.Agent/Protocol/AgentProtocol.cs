using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ExcelReportBuilder.Agent.Protocol;

public static class AgentProtocol
{
    public const string Version = "1.0";
    public const int MaximumFrameBytes = 1024 * 1024;

    internal static JsonSerializerOptions JsonOptions { get; } = CreateJsonOptions();

    public static AgentProtocolEnvelope Create<TPayload>(
        AgentMessageType messageType,
        string correlationId,
        TPayload payload)
    {
        if (string.IsNullOrWhiteSpace(correlationId))
        {
            throw new ArgumentException("A correlation ID is required.", nameof(correlationId));
        }

        var payloadJson = JsonSerializer.Serialize(payload, JsonOptions);
        using (var document = JsonDocument.Parse(payloadJson))
        {
            return new AgentProtocolEnvelope
            {
                ProtocolVersion = Version,
                MessageType = messageType,
                CorrelationId = correlationId,
                Payload = document.RootElement.Clone(),
            };
        }
    }

    public static byte[] Serialize(AgentProtocolEnvelope envelope)
    {
        if (envelope == null) throw new ArgumentNullException(nameof(envelope));
        ValidateEnvelope(envelope);

        var result = JsonSerializer.SerializeToUtf8Bytes(envelope, JsonOptions);
        if (result.Length > MaximumFrameBytes)
        {
            throw new AgentProtocolException("The protocol message exceeds the maximum frame size.");
        }

        return result;
    }

    public static AgentProtocolEnvelope Deserialize(byte[] json)
    {
        if (json == null) throw new ArgumentNullException(nameof(json));
        if (json.Length == 0 || json.Length > MaximumFrameBytes)
        {
            throw new AgentProtocolException("The protocol frame size is invalid.");
        }

        AgentProtocolEnvelope? envelope;
        try
        {
            envelope = JsonSerializer.Deserialize<AgentProtocolEnvelope>(json, JsonOptions);
        }
        catch (JsonException exception)
        {
            throw new AgentProtocolException("The protocol frame is not valid JSON.", exception);
        }

        if (envelope == null)
        {
            throw new AgentProtocolException("The protocol frame is empty.");
        }

        ValidateEnvelope(envelope);
        return envelope;
    }

    public static TPayload ReadPayload<TPayload>(AgentProtocolEnvelope envelope)
    {
        if (envelope == null) throw new ArgumentNullException(nameof(envelope));
        ValidateEnvelope(envelope);

        try
        {
            var value = JsonSerializer.Deserialize<TPayload>(envelope.Payload.GetRawText(), JsonOptions);
            if (value == null)
            {
                throw new AgentProtocolException("The protocol payload is empty.");
            }

            return value;
        }
        catch (JsonException exception)
        {
            throw new AgentProtocolException("The protocol payload is invalid.", exception);
        }
    }

    private static void ValidateEnvelope(AgentProtocolEnvelope envelope)
    {
        if (!string.Equals(envelope.ProtocolVersion, Version, StringComparison.Ordinal))
        {
            throw new AgentProtocolException(
                "Unsupported protocol version. Expected " + Version + ".");
        }

        if (string.IsNullOrWhiteSpace(envelope.CorrelationId) || envelope.CorrelationId.Length > 128)
        {
            throw new AgentProtocolException("The correlation ID is invalid.");
        }

        if (envelope.MessageType == AgentMessageType.Unknown ||
            !Enum.IsDefined(typeof(AgentMessageType), envelope.MessageType))
        {
            throw new AgentProtocolException("The protocol message type is invalid.");
        }

        if (envelope.Payload.ValueKind == JsonValueKind.Undefined)
        {
            throw new AgentProtocolException("The protocol payload is missing.");
        }
    }

    private static JsonSerializerOptions CreateJsonOptions()
    {
        var options = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = false,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
        };
        options.Converters.Add(new JsonStringEnumConverter(JsonNamingPolicy.CamelCase, false));
        return options;
    }
}

public sealed class AgentProtocolEnvelope
{
    public string ProtocolVersion { get; set; } = AgentProtocol.Version;

    public AgentMessageType MessageType { get; set; }

    public string CorrelationId { get; set; } = string.Empty;

    public JsonElement Payload { get; set; }
}

public enum AgentMessageType
{
    Unknown,
    Hello,
    HelloAcknowledged,
    StartJob,
    CancelJob,
    CancelAcknowledged,
    ProbeEndpoint,
    ProbeCompleted,
    HostToolRequest,
    HostToolResult,
    ListResumeMetadata,
    ResumeMetadata,
    Progress,
    Checkpoint,
    JobCompleted,
    JobFailed,
    JobCancelled,
    Error,
}

public sealed class AgentProtocolException : IOException
{
    public AgentProtocolException(string message, Exception? innerException = null)
        : base(message, innerException)
    {
    }
}
