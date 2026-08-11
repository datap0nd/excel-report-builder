using ExcelReportBuilder.Worker;
using ExcelReportBuilder.Agent.Security;

namespace ExcelReportBuilder.Agent.Tests;

public sealed class PipeNamePolicyTests
{
    [Fact]
    public void RandomPipeName_IsBoundedUniqueAndDoesNotExposeUserName()
    {
        var pipeName = WorkerHandshakeAuthenticator.CreatePipeName();
        var secondPipeName = WorkerHandshakeAuthenticator.CreatePipeName();

        Assert.StartsWith(PipeNamePolicy.Prefix, pipeName, StringComparison.Ordinal);
        Assert.True(pipeName.Length <= PipeNamePolicy.MaximumLength);
        Assert.NotEqual(pipeName, secondPipeName);
        Assert.Equal(pipeName, PipeNamePolicy.Validate(pipeName));
        Assert.Equal(secondPipeName, PipeNamePolicy.Validate(secondPipeName));
        if (!string.IsNullOrEmpty(Environment.UserName))
        {
            Assert.DoesNotContain(Environment.UserName, pipeName, StringComparison.OrdinalIgnoreCase);
        }
    }

    [Theory]
    [InlineData("other-product-pipe")]
    [InlineData("excel-report-builder-invalid/name")]
    [InlineData("excel-report-builder-invalid\\name")]
    public void Validate_RejectsNamesOutsideWorkerNamespace(string pipeName)
    {
        Assert.Throws<ArgumentException>(() => PipeNamePolicy.Validate(pipeName));
    }
}
