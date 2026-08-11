using System;
using System.Security.Cryptography;
using System.Text;

namespace ExcelReportBuilder.Agent.Security;

/// <summary>
/// Creates and verifies the one-time proof used between the add-in and the
/// worker process it launched. The secret is inherited through the child
/// process environment and is never sent over the named pipe.
/// </summary>
public static class WorkerHandshakeAuthenticator
{
    public const string PipePrefix = "excel-report-builder-";
    public const int MaximumPipeNameLength = 128;

    public const string SecretEnvironmentVariable =
        "EXCEL_REPORT_BUILDER_WORKER_HANDSHAKE_SECRET";

    private const int SecretByteCount = 32;
    private const int NonceByteCount = 32;

    public static string CreateSecret()
    {
        return CreateRandomBase64(SecretByteCount);
    }

    public static string CreatePipeName()
    {
        var bytes = new byte[24];
        using (var random = RandomNumberGenerator.Create())
        {
            random.GetBytes(bytes);
        }

        try
        {
            return PipePrefix + BitConverter.ToString(bytes).Replace("-", string.Empty).ToLowerInvariant();
        }
        finally
        {
            Array.Clear(bytes, 0, bytes.Length);
        }
    }

    public static string CreateNonce()
    {
        return CreateRandomBase64(NonceByteCount);
    }

    public static bool IsValidSecret(string secret)
    {
        byte[] bytes;
        try
        {
            bytes = DecodeExact(secret, SecretByteCount, "worker secret");
        }
        catch (ArgumentException)
        {
            return false;
        }

        Array.Clear(bytes, 0, bytes.Length);
        return true;
    }

    public static string ComputeAuthenticationTag(
        string secret,
        string pipeName,
        string clientNonce,
        string protocolVersion)
    {
        byte[] secretBytes = DecodeExact(secret, SecretByteCount, "worker secret");
        try
        {
            byte[] nonceBytes = DecodeExact(clientNonce, NonceByteCount, "client nonce");
            try
            {
                byte[] messageBytes = Encoding.UTF8.GetBytes(
                    "excel-report-builder-worker-handshake-v1\n" +
                    (pipeName ?? string.Empty) + "\n" +
                    clientNonce + "\n" +
                    (protocolVersion ?? string.Empty));
                try
                {
                    using (var hmac = new HMACSHA256(secretBytes))
                    {
                        return Convert.ToBase64String(hmac.ComputeHash(messageBytes));
                    }
                }
                finally
                {
                    Array.Clear(messageBytes, 0, messageBytes.Length);
                }
            }
            finally
            {
                Array.Clear(nonceBytes, 0, nonceBytes.Length);
            }
        }
        finally
        {
            Array.Clear(secretBytes, 0, secretBytes.Length);
        }
    }

    public static bool VerifyAuthenticationTag(
        string secret,
        string pipeName,
        string clientNonce,
        string protocolVersion,
        string authenticationTag)
    {
        byte[]? expected = null;
        byte[]? actual = null;
        try
        {
            expected = Convert.FromBase64String(ComputeAuthenticationTag(
                secret,
                pipeName,
                clientNonce,
                protocolVersion));
            actual = Convert.FromBase64String(authenticationTag ?? string.Empty);
            if (expected.Length != actual.Length)
            {
                return false;
            }

            var difference = 0;
            for (var index = 0; index < expected.Length; index++)
            {
                difference |= expected[index] ^ actual[index];
            }

            return difference == 0;
        }
        catch (FormatException)
        {
            return false;
        }
        catch (ArgumentException)
        {
            return false;
        }
        finally
        {
            if (expected != null)
            {
                Array.Clear(expected, 0, expected.Length);
            }

            if (actual != null)
            {
                Array.Clear(actual, 0, actual.Length);
            }
        }
    }

    private static string CreateRandomBase64(int byteCount)
    {
        var bytes = new byte[byteCount];
        using (var random = RandomNumberGenerator.Create())
        {
            random.GetBytes(bytes);
        }

        try
        {
            return Convert.ToBase64String(bytes);
        }
        finally
        {
            Array.Clear(bytes, 0, bytes.Length);
        }
    }

    private static byte[] DecodeExact(string value, int expectedLength, string fieldName)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            throw new ArgumentException("The " + fieldName + " is missing.", nameof(value));
        }

        byte[] bytes;
        try
        {
            bytes = Convert.FromBase64String(value);
        }
        catch (FormatException exception)
        {
            throw new ArgumentException("The " + fieldName + " is invalid.", nameof(value), exception);
        }

        if (bytes.Length != expectedLength)
        {
            Array.Clear(bytes, 0, bytes.Length);
            throw new ArgumentException("The " + fieldName + " is invalid.", nameof(value));
        }

        if (!string.Equals(Convert.ToBase64String(bytes), value, StringComparison.Ordinal))
        {
            Array.Clear(bytes, 0, bytes.Length);
            throw new ArgumentException("The " + fieldName + " is invalid.", nameof(value));
        }

        return bytes;
    }
}
