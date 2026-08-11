using ExcelReportBuilder.Agent.Protocol;
using ExcelReportBuilder.Agent.Security;

namespace ExcelReportBuilder.Agent.Tests;

public sealed class WorkerHandshakeAuthenticatorTests
{
    [Fact]
    public void AuthenticationTag_VerifiesOnlyForExactLaunchContext()
    {
        string secret = WorkerHandshakeAuthenticator.CreateSecret();
        string pipeName = WorkerHandshakeAuthenticator.CreatePipeName();
        string nonce = WorkerHandshakeAuthenticator.CreateNonce();
        string tag = WorkerHandshakeAuthenticator.ComputeAuthenticationTag(
            secret,
            pipeName,
            nonce,
            AgentProtocol.Version);

        Assert.True(WorkerHandshakeAuthenticator.VerifyAuthenticationTag(
            secret,
            pipeName,
            nonce,
            AgentProtocol.Version,
            tag));
        Assert.False(WorkerHandshakeAuthenticator.VerifyAuthenticationTag(
            WorkerHandshakeAuthenticator.CreateSecret(),
            pipeName,
            nonce,
            AgentProtocol.Version,
            tag));
        Assert.False(WorkerHandshakeAuthenticator.VerifyAuthenticationTag(
            secret,
            WorkerHandshakeAuthenticator.CreatePipeName(),
            nonce,
            AgentProtocol.Version,
            tag));
        Assert.False(WorkerHandshakeAuthenticator.VerifyAuthenticationTag(
            secret,
            pipeName,
            WorkerHandshakeAuthenticator.CreateNonce(),
            AgentProtocol.Version,
            tag));
        Assert.False(WorkerHandshakeAuthenticator.VerifyAuthenticationTag(
            secret,
            pipeName,
            nonce,
            "9.9",
            tag));
        Assert.False(WorkerHandshakeAuthenticator.VerifyAuthenticationTag(
            secret,
            pipeName,
            nonce,
            AgentProtocol.Version,
            Convert.ToBase64String(new byte[32])));
        Assert.False(WorkerHandshakeAuthenticator.VerifyAuthenticationTag(
            secret,
            pipeName,
            nonce,
            AgentProtocol.Version,
            "not-base64"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("not-base64")]
    [InlineData("YQ==")]
    [InlineData("YWFhYWFhYWFhYWFhYWFhYWFhYWFhYWFhYWFhYWFhYWE=\n")]
    public void InvalidSecret_IsRejected(string secret)
    {
        Assert.False(WorkerHandshakeAuthenticator.IsValidSecret(secret));
    }

    [Theory]
    [InlineData("")]
    [InlineData("not-base64")]
    [InlineData("YQ==")]
    public void InvalidNonce_CannotProduceAProof(string nonce)
    {
        Assert.Throws<ArgumentException>(() =>
            WorkerHandshakeAuthenticator.ComputeAuthenticationTag(
                WorkerHandshakeAuthenticator.CreateSecret(),
                WorkerHandshakeAuthenticator.CreatePipeName(),
                nonce,
                AgentProtocol.Version));
    }

    [Fact]
    public void RandomCredentials_AreFreshAndBounded()
    {
        string firstSecret = WorkerHandshakeAuthenticator.CreateSecret();
        string secondSecret = WorkerHandshakeAuthenticator.CreateSecret();
        string firstNonce = WorkerHandshakeAuthenticator.CreateNonce();
        string secondNonce = WorkerHandshakeAuthenticator.CreateNonce();

        Assert.NotEqual(firstSecret, secondSecret);
        Assert.NotEqual(firstNonce, secondNonce);
        Assert.True(WorkerHandshakeAuthenticator.IsValidSecret(firstSecret));
        Assert.Equal(32, Convert.FromBase64String(firstNonce).Length);
    }
}
