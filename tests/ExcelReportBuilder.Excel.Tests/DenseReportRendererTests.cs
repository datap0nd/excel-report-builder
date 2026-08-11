using ExcelReportBuilder.Excel.Rendering;
using ExcelReportBuilder.Excel.Ownership;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class DenseReportRendererTests
{
    [Fact]
    public void Rejects_a_write_outside_the_managed_block_extent()
    {
        var plan = new DenseGridPlan
        {
            AnchorCell = "A1",
            OwnedRowCount = 2,
            OwnedColumnCount = 2,
            Cells =
            {
                new DenseCellWrite
                {
                    RelativeRow = 0,
                    RelativeColumn = 1,
                    ColumnSpan = 2,
                    Kind = DenseCellValueKind.Text,
                    Value = "Too wide"
                }
            }
        };

        var exception = Assert.Throws<InvalidOperationException>(() =>
            DenseReportRenderer.ValidatePlan(plan));

        Assert.Contains("owned extent", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Rejects_a_generated_formula_beyond_excels_limit()
    {
        var formula = "=" + new string('1', SafeExcelFormula.MaximumFormulaCharacters);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            SafeFormulaFactory.FromTypedMeasure(formula));

        Assert.Contains("formula length", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Clears_only_the_owned_block_range_before_an_idempotent_render()
    {
        var identity = new ManagedObjectIdentity(
            "report",
            "output",
            ManagedObjectKind.DraftWorksheet);
        var worksheet = new FakeWorksheet();
        new ManagedOwnershipGuard().MarkOwned(worksheet, identity);
        var plan = new DenseGridPlan
        {
            BlockId = "summary",
            AnchorCell = "C4",
            OwnedRowCount = 5,
            OwnedColumnCount = 7
        };

        new DenseReportRenderer().Render(worksheet, identity, plan);
        new DenseReportRenderer().Render(worksheet, identity, plan);

        Assert.Equal(
            new[] { "clear:C4:I8", "clear:C4:I8" },
            worksheet.Log);
        Assert.DoesNotContain(worksheet.Log, entry => entry.Contains("A1", StringComparison.Ordinal));
    }

    public sealed class FakeWorksheet
    {
        public FakeWorksheet()
        {
            Cells = new FakeCells();
            Range = new FakeRanges(Log);
        }

        public List<string> Log { get; } = new();

        public ManagedWorksheetServiceTests.FakeCustomProperties CustomProperties { get; } = new();

        public FakeCells Cells { get; }

        public FakeRanges Range { get; }
    }

    public sealed class FakeCells
    {
        public FakeCell this[int row, int column] => new(row, column);
    }

    public sealed class FakeRanges
    {
        private readonly List<string> log;

        public FakeRanges(List<string> log)
        {
            this.log = log;
        }

        public FakeRange this[FakeCell first, FakeCell last] => new(first, last, log);
    }

    public sealed class FakeRange
    {
        private readonly FakeCell first;
        private readonly FakeCell last;
        private readonly List<string> log;

        public FakeRange(FakeCell first, FakeCell last, List<string> log)
        {
            this.first = first;
            this.last = last;
            this.log = log;
        }

        public void Clear()
        {
            log.Add("clear:" + first.Address + ":" + last.Address);
        }

        public void UnMerge()
        {
        }
    }

    public sealed class FakeCell
    {
        public FakeCell(int row, int column)
        {
            Row = row;
            Column = column;
        }

        public int Row { get; }

        public int Column { get; }

        public string Address => ColumnName(Column) + Row;

        private static string ColumnName(int column)
        {
            var result = string.Empty;
            while (column > 0)
            {
                column--;
                result = (char)('A' + column % 26) + result;
                column /= 26;
            }

            return result;
        }
    }
}
