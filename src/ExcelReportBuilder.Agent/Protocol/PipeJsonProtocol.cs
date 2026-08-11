using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelReportBuilder.Agent.Protocol;

/// <summary>
/// Length-prefixed UTF-8 JSON framing. The four-byte length is little endian.
/// </summary>
public static class PipeJsonProtocol
{
    public static async Task WriteAsync(
        Stream stream,
        AgentProtocolEnvelope envelope,
        CancellationToken cancellationToken)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));

        var frame = AgentProtocol.Serialize(envelope);
        var header = new byte[4];
        header[0] = (byte)frame.Length;
        header[1] = (byte)(frame.Length >> 8);
        header[2] = (byte)(frame.Length >> 16);
        header[3] = (byte)(frame.Length >> 24);

        await stream.WriteAsync(header, 0, header.Length, cancellationToken).ConfigureAwait(false);
        await stream.WriteAsync(frame, 0, frame.Length, cancellationToken).ConfigureAwait(false);
        await stream.FlushAsync(cancellationToken).ConfigureAwait(false);
    }

    public static async Task<AgentProtocolEnvelope?> ReadAsync(
        Stream stream,
        CancellationToken cancellationToken)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));

        var header = new byte[4];
        var headerBytes = await ReadAtMostAsync(stream, header, cancellationToken).ConfigureAwait(false);
        if (headerBytes == 0)
        {
            return null;
        }

        if (headerBytes != header.Length)
        {
            throw new AgentProtocolException("The protocol frame header ended early.");
        }

        var length = header[0] |
                     (header[1] << 8) |
                     (header[2] << 16) |
                     (header[3] << 24);
        if (length <= 0 || length > AgentProtocol.MaximumFrameBytes)
        {
            throw new AgentProtocolException("The protocol frame length is invalid.");
        }

        var frame = new byte[length];
        var frameBytes = await ReadAtMostAsync(stream, frame, cancellationToken).ConfigureAwait(false);
        if (frameBytes != length)
        {
            throw new AgentProtocolException("The protocol frame ended early.");
        }

        return AgentProtocol.Deserialize(frame);
    }

    private static async Task<int> ReadAtMostAsync(
        Stream stream,
        byte[] buffer,
        CancellationToken cancellationToken)
    {
        var offset = 0;
        while (offset < buffer.Length)
        {
            var read = await stream.ReadAsync(
                buffer,
                offset,
                buffer.Length - offset,
                cancellationToken).ConfigureAwait(false);
            if (read == 0)
            {
                break;
            }

            offset += read;
        }

        return offset;
    }
}
