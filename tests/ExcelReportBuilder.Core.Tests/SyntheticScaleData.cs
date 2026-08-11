using ExcelReportBuilder.Core.Planning;

namespace ExcelReportBuilder.Core.Tests;

internal static class SyntheticScaleData
{
    public const string FullScaleEnvironmentVariable = "ERB_RUN_FULL_WORKSHEET_SCALE";

    public static IReadOnlyList<string> Headers { get; } = new[]
    {
        "Period", "Region", "Amount"
    };

    public static IReadOnlyList<object?[]> CreateRows(int rowCount)
    {
        if (rowCount < 0 || rowCount > RowProjection.MaximumWorksheetDataRows)
        {
            throw new ArgumentOutOfRangeException(nameof(rowCount));
        }

        var rows = new object?[rowCount][];
        for (var index = 0; index < rowCount; index++)
        {
            rows[index] = new object?[]
            {
                new DateTime(2026, index % 12 + 1, 1),
                "Region " + (index % 20),
                index % 1000
            };
        }

        return rows;
    }
}
