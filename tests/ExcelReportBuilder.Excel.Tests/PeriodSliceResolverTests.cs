using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Excel.Rendering;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class PeriodSliceResolverTests
{
    [Fact]
    public void Binds_only_the_explicit_ranges_resolved_by_the_core_plan()
    {
        var periods = Enumerable.Range(0, 5)
            .Select(offset => new DateTime(2026, 1, 1).AddMonths(offset))
            .Select(date => new PeriodMember { Period = date, PivotValue = date })
            .ToList();
        var slices = new[]
        {
            new ExcelReportBuilder.Core.Periods.ResolvedPeriodSlice
            {
                Id = "selected",
                Label = "Selected",
                Kind = PeriodSliceKind.Selected,
                StartInclusive = new DateTime(2026, 2, 1),
                EndInclusive = new DateTime(2026, 3, 31)
            },
            new ExcelReportBuilder.Core.Periods.ResolvedPeriodSlice
            {
                Id = "prior",
                Label = "Prior",
                Kind = PeriodSliceKind.Prior,
                StartInclusive = new DateTime(2025, 12, 1),
                EndInclusive = new DateTime(2026, 1, 31),
                BasedOnSliceId = "selected"
            }
        };

        var result = PeriodSliceResolver.BindResolved(slices, periods);

        Assert.Equal(
            new[] { new DateTime(2026, 2, 1), new DateTime(2026, 3, 1) },
            result["selected"].Cast<DateTime>());
        Assert.Equal(new DateTime(2026, 1, 1), Assert.IsType<DateTime>(Assert.Single(result["prior"])));
    }

    [Fact]
    public void Resolves_current_selected_prior_and_same_period_prior_year_from_available_data()
    {
        var periods = Enumerable.Range(0, 15)
            .Select(offset => new DateTime(2025, 1, 1).AddMonths(offset))
            .Select(date => new PeriodMember { Period = date, PivotValue = date })
            .ToList();
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
                Id = "selected",
                Label = "Selected",
                Kind = PeriodSliceKind.Selected,
                SelectedStart = new DateTime(2026, 1, 1),
                SelectedEnd = new DateTime(2026, 3, 31)
            },
            new PeriodSliceSpec
            {
                Id = "same_prior_year",
                Label = "Same period prior year",
                Kind = PeriodSliceKind.SamePeriodPriorYear,
                BasedOnSliceId = "selected"
            }
        };

        var result = PeriodSliceResolver.Resolve(slices, periods);

        Assert.Equal(new DateTime(2026, 3, 1), Assert.IsType<DateTime>(Assert.Single(result["current"])));
        Assert.Equal(new DateTime(2026, 2, 1), Assert.IsType<DateTime>(Assert.Single(result["prior"])));
        Assert.Equal(3, result["selected"].Count);
        Assert.Equal(
            new[] { new DateTime(2025, 1, 1), new DateTime(2025, 2, 1), new DateTime(2025, 3, 1) },
            result["same_prior_year"].Cast<DateTime>());
    }

    [Fact]
    public void Rejects_a_slice_that_would_invent_a_missing_period()
    {
        var periods = new[]
        {
            new PeriodMember { Period = new DateTime(2026, 1, 1), PivotValue = new DateTime(2026, 1, 1) }
        };
        var slices = new[]
        {
            new PeriodSliceSpec
            {
                Id = "current",
                Label = "Current",
                Kind = PeriodSliceKind.Current,
                SelectedStart = new DateTime(2026, 1, 1),
                SelectedEnd = new DateTime(2026, 1, 31)
            },
            new PeriodSliceSpec
            {
                Id = "prior",
                Label = "Prior",
                Kind = PeriodSliceKind.Prior,
                BasedOnSliceId = "current"
            }
        };

        var exception = Assert.Throws<InvalidOperationException>(() => PeriodSliceResolver.Resolve(slices, periods));

        Assert.Contains("no matching periods", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}
