using ExcelReportBuilder.Worker;

namespace ExcelReportBuilder.Agent.Tests;

public sealed class PipeNamePolicyTests
{
    [Fact]
    public void DefaultPipeName_IsBoundedAndDoesNotExposeUserName()
    {
        var pipeName = PipeNamePolicy.CreateDefaultForCurrentUser();

        Assert.StartsWith(PipeNamePolicy.Prefix, pipeName, StringComparison.Ordinal);
        Assert.True(pipeName.Length <= PipeNamePolicy.MaximumLength);
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
