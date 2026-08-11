using System.Text.Json;
using ExcelReportBuilder.Core.Measures;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Core.Transforms;
using ExcelReportBuilder.Core.Validation;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelReportBuilder.Core.Tests;

public sealed class ReportSpecJsonTests
{
    [Fact]
    public void Round_trips_every_polymorphic_family_without_type_names()
    {
        var original = SyntheticReportFactory.CreateValidLongSpec();
        original.Transforms.Add(new AddArithmeticColumnTransform
        {
            Id = "arithmetic",
            OutputColumn = "UnitAmount",
            Operator = ArithmeticOperator.Divide,
            Left = new ArithmeticOperand { Kind = ArithmeticOperandKind.Column, Column = "Amount" },
            Right = new ArithmeticOperand { Kind = ArithmeticOperandKind.Column, Column = "Units" }
        });
        original.Measures.Add(new MeasureDefinition
        {
            Id = "filtered_amount_other",
            Label = "Filtered amount",
            ValueType = MeasureValueType.Currency,
            Expression = new FilteredAggregateMeasureExpression
            {
                ResultType = MeasureValueType.Currency,
                Field = "Amount",
                Function = AggregateFunction.Sum,
                Filters =
                {
                    new MeasureFilterSpec
                    {
                        Field = "Category",
                        Operator = MeasureFilterOperator.In,
                        Values = { ScalarValue.FromText("Core"), ScalarValue.FromText("Other") }
                    }
                }
            }
        });

        var json = ReportSpecJson.Serialize(original);
        var roundTripped = ReportSpecJson.Deserialize(json);

        Assert.DoesNotContain("$type", json, StringComparison.Ordinal);
        Assert.Contains("\"kind\": \"derivePeriodParts\"", json);
        Assert.Contains("\"kind\": \"safeDivide\"", json);
        Assert.IsType<DerivePeriodPartsTransform>(roundTripped.Transforms[1]);
        Assert.IsType<SafeDivideMeasureExpression>(roundTripped.Measures[2].Expression);
        Assert.IsType<FilteredAggregateMeasureExpression>(roundTripped.Measures[^1].Expression);
        Assert.Equal(2, roundTripped.Blocks.Count);
    }

    [Fact]
    public void Schema_is_valid_json_and_declares_the_complete_closed_unions()
    {
        var schemaPath = FindSchema();
        using var document = JsonDocument.Parse(File.ReadAllText(schemaPath));
        var definitions = document.RootElement.GetProperty("$defs");
        var transformAlternatives = definitions.GetProperty("transform").GetProperty("oneOf");
        var expressionAlternatives = definitions.GetProperty("measureExpression").GetProperty("oneOf");

        Assert.Equal(17, transformAlternatives.GetArrayLength());
        Assert.Equal(10, expressionAlternatives.GetArrayLength());
        Assert.True(definitions.TryGetProperty("periodMapping", out _));
        Assert.True(definitions.TryGetProperty("reportBlock", out _));
        Assert.True(definitions.TryGetProperty("topN", out _));
        Assert.True(definitions.TryGetProperty("periodSlice", out _));
        Assert.True(definitions.TryGetProperty("sourceFingerprint", out _));
    }

    [Fact]
    public void Rejects_unsupported_versions_on_read_and_write()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        spec.SchemaVersion = "2.0";

        Assert.Throws<UnsupportedReportSpecVersionException>(() => ReportSpecJson.Serialize(spec));
        Assert.Throws<UnsupportedReportSpecVersionException>(() =>
            ReportSpecJson.Deserialize("{\"schemaVersion\":\"2.0\"}"));
        Assert.Throws<UnsupportedReportSpecVersionException>(() =>
            ReportSpecJson.Deserialize("{}"));
    }

    [Fact]
    public void Rejects_numeric_enum_tokens_instead_of_bypassing_the_string_schema()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        var document = JObject.Parse(ReportSpecJson.Serialize(spec));
        document["source"]!["kind"] = 0;

        Assert.Throws<JsonSerializationException>(() => ReportSpecJson.Deserialize(document.ToString()));
    }

    [Fact]
    public void Rejects_numeric_strings_in_custom_union_discriminators()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        var document = JObject.Parse(ReportSpecJson.Serialize(spec));
        document["transforms"]![0]!["kind"] = "1";

        Assert.Throws<JsonSerializationException>(() => ReportSpecJson.Deserialize(document.ToString()));
    }

    [Fact]
    public void Rejects_numeric_strings_for_regular_enum_members()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        var document = JObject.Parse(ReportSpecJson.Serialize(spec));
        document["source"]!["kind"] = "0";

        Assert.Throws<JsonSerializationException>(() => ReportSpecJson.Deserialize(document.ToString()));
    }

    [Fact]
    public void Rejects_explicit_nulls_for_schema_required_members()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        var document = JObject.Parse(ReportSpecJson.Serialize(spec));
        document["transforms"] = JValue.CreateNull();

        Assert.Throws<JsonSerializationException>(() => ReportSpecJson.Deserialize(document.ToString()));
    }


    [Fact]
    public void Rejects_omitted_nested_required_collections()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        var document = JObject.Parse(ReportSpecJson.Serialize(spec));
        ((JObject)document["blocks"]![0]!).Remove("spacers");

        Assert.Throws<JsonSerializationException>(() => ReportSpecJson.Deserialize(document.ToString()));
    }

    [Fact]
    public void Rejects_omitted_required_members_before_defaults_can_hide_them()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        var document = JObject.Parse(ReportSpecJson.Serialize(spec));
        document.Remove("checks");

        Assert.Throws<JsonSerializationException>(() => ReportSpecJson.Deserialize(document.ToString()));
    }

    [Fact]
    public void Round_trips_explicit_period_grain_and_schema_allows_it()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        spec.PeriodMapping!.Grain = PeriodGrain.Month;

        var json = ReportSpecJson.Serialize(spec);
        var roundTripped = ReportSpecJson.Deserialize(json);

        Assert.Contains("\"grain\": \"month\"", json, StringComparison.Ordinal);
        Assert.Equal(PeriodGrain.Month, roundTripped.PeriodMapping!.Grain);

        using var document = JsonDocument.Parse(File.ReadAllText(FindSchema()));
        var grain = document.RootElement.GetProperty("$defs")
            .GetProperty("periodMapping")
            .GetProperty("properties")
            .GetProperty("grain");
        Assert.Equal(3, grain.GetProperty("enum").GetArrayLength());
    }

    private static string FindSchema()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null)
        {
            var candidate = Path.Combine(directory.FullName, "schemas", "report-spec-v1.schema.json");
            if (File.Exists(candidate))
            {
                return candidate;
            }

            directory = directory.Parent;
        }

        throw new FileNotFoundException("Could not find schemas/report-spec-v1.schema.json.");
    }
}
