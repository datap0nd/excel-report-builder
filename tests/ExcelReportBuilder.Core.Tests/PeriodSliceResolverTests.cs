using ExcelReportBuilder.Core.Periods;
using ExcelReportBuilder.Core.Specifications;

namespace ExcelReportBuilder.Core.Tests;

public sealed class PeriodSliceResolverTests
{
    [Fact]
    public void Resolves_current_prior_and_same_period_prior_year_to_explicit_dates()
    {
        var slices = new[]
        {
            new PeriodSliceSpec
            {
                Id = "current",
                Label = "Current",
                Kind = PeriodSliceKind.Current,
                SelectedStart = new DateTime(2026, 3, 1),
                SelectedEnd = new DateTime(2026, 3, 31)
            },
            new PeriodSliceSpec
            {
                Id = "prior",
                Label = "Prior",
                Kind = PeriodSliceKind.Prior,
                BasedOnSliceId = "current"
            },
            new PeriodSliceSpec
            {
                Id = "prior_year",
                Label = "Prior year",
                Kind = PeriodSliceKind.SamePeriodPriorYear,
                BasedOnSliceId = "current"
            }
        };

        var resolved = PeriodSliceResolver.Resolve(slices);

        Assert.Equal(new DateTime(2026, 3, 1), resolved[0].StartInclusive);
        Assert.Equal(new DateTime(2026, 2, 1), resolved[1].StartInclusive);
        Assert.Equal(new DateTime(2026, 2, 28), resolved[1].EndInclusive);
        Assert.Equal(new DateTime(2025, 3, 1), resolved[2].StartInclusive);
        Assert.Equal(new DateTime(2025, 3, 31), resolved[2].EndInclusive);
    }

    [Fact]
    public void Current_is_never_reinterpreted_from_available_data()
    {
        var resolved = PeriodSliceResolver.Resolve(new[]
        {
            new PeriodSliceSpec
            {
                Id = "current",
                Label = "Current",
                Kind = PeriodSliceKind.Current,
                SelectedStart = new DateTime(2024, 7, 1),
                SelectedEnd = new DateTime(2024, 7, 31)
            }
        }).Single();

        Assert.Equal(new DateTime(2024, 7, 1), resolved.StartInclusive);
        Assert.Equal(new DateTime(2024, 7, 31), resolved.EndInclusive);
    }
}
