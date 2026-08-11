using System.Collections.Generic;

namespace ExcelReportBuilder.Agent.Tools;

internal static class AgentToolSchemaFactory
{
    public static object PeriodMapping()
    {
        return Object(
            new Dictionary<string, object>
            {
                ["mode"] = EnumString("dateColumn", "widePeriods", "unresolved"),
                ["periodField"] = BoundedString(128),
                ["reportingYear"] = new Dictionary<string, object>
                {
                    ["type"] = new[] { "integer", "null" },
                    ["minimum"] = 1900,
                    ["maximum"] = 9999
                },
                ["mappings"] = Array(
                    Object(
                        new Dictionary<string, object>
                        {
                            ["sourceField"] = BoundedString(128),
                            ["periodLabel"] = BoundedString(32),
                            ["metric"] = BoundedString(128)
                        },
                        "sourceField", "periodLabel", "metric"),
                    120)
            },
            "mode", "periodField", "reportingYear", "mappings");
    }

    public static object AdvancedTransforms()
    {
        var simpleKinds = new[]
        {
            "rename", "trimText", "convertNumber", "convertDate", "replaceBlank",
            "normalizeBlanks", "normalizeErrors", "fillDown"
        };
        var variants = new List<object>();
        foreach (var kind in simpleKinds)
        {
            variants.Add(Object(
                new Dictionary<string, object>
                {
                    ["kind"] = ConstString(kind),
                    ["sourceField"] = BoundedString(128),
                    ["outputField"] = BoundedString(128)
                },
                "kind", "sourceField", "outputField"));
        }

        variants.Add(Object(
            new Dictionary<string, object>
            {
                ["kind"] = ConstString("filterRows"),
                ["sourceField"] = BoundedString(128),
                ["outputField"] = BoundedString(128),
                ["operator"] = EnumString(
                    "equal", "notEqual", "contains", "startsWith", "endsWith", "isBlank", "isNotBlank"),
                ["value"] = BoundedString(1024)
            },
            "kind", "sourceField", "outputField", "operator", "value"));
        variants.Add(Object(
            new Dictionary<string, object>
            {
                ["kind"] = ConstString("mapValues"),
                ["sourceField"] = BoundedString(128),
                ["outputField"] = BoundedString(128),
                ["mappings"] = Array(
                    Object(
                        new Dictionary<string, object>
                        {
                            ["from"] = BoundedString(1024),
                            ["to"] = BoundedString(1024)
                        },
                        "from", "to"),
                    200,
                    1)
            },
            "kind", "sourceField", "outputField", "mappings"));
        variants.Add(Object(
            new Dictionary<string, object>
            {
                ["kind"] = ConstString("excludeTotalRows"),
                ["sourceField"] = BoundedString(128),
                ["outputField"] = BoundedString(128),
                ["matchKind"] = EnumString("equalsAny", "startsWith", "contains", "isBlank"),
                ["values"] = Array(BoundedString(1024), 50),
                ["evidenceSource"] = EnumString("profile", "preview", "userConfirmation"),
                ["observedMatchCount"] = Integer(0, int.MaxValue)
            },
            "kind", "sourceField", "outputField", "matchKind", "values",
            "evidenceSource", "observedMatchCount"));
        variants.Add(Object(
            new Dictionary<string, object>
            {
                ["kind"] = ConstString("derivePeriodPart"),
                ["sourceField"] = BoundedString(128),
                ["outputField"] = BoundedString(128),
                ["part"] = EnumString("year", "half", "quarter", "monthNumber", "monthName", "yearMonth")
            },
            "kind", "sourceField", "outputField", "part"));
        variants.Add(Object(
            new Dictionary<string, object>
            {
                ["kind"] = ConstString("addArithmeticColumn"),
                ["sourceField"] = BoundedString(128),
                ["outputField"] = BoundedString(128),
                ["operator"] = EnumString("add", "subtract", "multiply", "divide"),
                ["rightField"] = BoundedString(128),
                ["rightNumber"] = NullableNumber()
            },
            "kind", "sourceField", "outputField", "operator", "rightField", "rightNumber"));

        return Object(
            new Dictionary<string, object>
            {
                ["transforms"] = Array(
                    new Dictionary<string, object> { ["oneOf"] = variants.ToArray() },
                    32)
            },
            "transforms");
    }

    public static object AdvancedReportSpecification()
    {
        var definitions = new Dictionary<string, object>
        {
            ["measureFilter"] = Object(
                new Dictionary<string, object>
                {
                    ["field"] = BoundedString(128),
                    ["operator"] = EnumString(
                        "equal", "notEqual", "greaterThan", "greaterThanOrEqual",
                        "lessThan", "lessThanOrEqual", "in", "notIn", "isBlank", "isNotBlank"),
                    ["values"] = Array(BoundedString(1024), 50)
                },
                "field", "operator", "values"),
            ["expression"] = new Dictionary<string, object>
            {
                ["oneOf"] = new object[]
                {
                    Object(
                        new Dictionary<string, object>
                        {
                            ["kind"] = ConstString("aggregate"),
                            ["field"] = BoundedString(128),
                            ["aggregation"] = AggregateEnum(),
                            ["periodSliceId"] = BoundedString(128)
                        },
                        "kind", "field", "aggregation", "periodSliceId"),
                    Object(
                        new Dictionary<string, object>
                        {
                            ["kind"] = ConstString("filteredAggregate"),
                            ["field"] = BoundedString(128),
                            ["aggregation"] = AggregateEnum(),
                            ["periodSliceId"] = BoundedString(128),
                            ["filters"] = Array(Reference("measureFilter"), 16)
                        },
                        "kind", "field", "aggregation", "periodSliceId", "filters"),
                    Object(
                        new Dictionary<string, object>
                        {
                            ["kind"] = ConstString("reference"),
                            ["measureId"] = BoundedString(128)
                        },
                        "kind", "measureId"),
                    Object(
                        new Dictionary<string, object>
                        {
                            ["kind"] = ConstString("constant"),
                            ["value"] = Number()
                        },
                        "kind", "value"),
                    Object(
                        new Dictionary<string, object>
                        {
                            ["kind"] = ConstString("binary"),
                            ["operator"] = EnumString("add", "subtract", "multiply", "divide"),
                            ["leftMeasureId"] = BoundedString(128),
                            ["rightMeasureId"] = BoundedString(128),
                            ["returnBlankOnZeroDenominator"] = Boolean()
                        },
                        "kind", "operator", "leftMeasureId", "rightMeasureId",
                        "returnBlankOnZeroDenominator"),
                    RatioExpression("safeDivide"),
                    RatioExpression("ratio"),
                    RatioExpression("share"),
                    Object(
                        new Dictionary<string, object>
                        {
                            ["kind"] = ConstString("difference"),
                            ["differenceKind"] = EnumString("absolute", "percentage", "percentagePoints"),
                            ["currentMeasureId"] = BoundedString(128),
                            ["baselineMeasureId"] = BoundedString(128),
                            ["onZero"] = ZeroEnum()
                        },
                        "kind", "differenceKind", "currentMeasureId", "baselineMeasureId", "onZero")
                }
            },
            ["measure"] = Object(
                new Dictionary<string, object>
                {
                    ["id"] = BoundedString(128),
                    ["label"] = BoundedString(128),
                    ["valueType"] = EnumString("wholeNumber", "number", "currency", "percentage"),
                    ["numberFormat"] = BoundedString(128),
                    ["expression"] = Reference("expression")
                },
                "id", "label", "valueType", "numberFormat", "expression"),
            ["fieldPlacement"] = Object(
                new Dictionary<string, object>
                {
                    ["field"] = BoundedString(128),
                    ["caption"] = BoundedString(128),
                    ["subtotalMode"] = EnumString("none", "automatic"),
                    ["subtotalPlacement"] = EnumString("beforeMembers", "afterMembers"),
                    ["subtotalLabel"] = BoundedString(128),
                    ["sort"] = EnumString("sourceOrder", "ascending", "descending"),
                    ["memberOrder"] = Array(BoundedString(1024), 200)
                },
                "field", "caption", "subtotalMode", "subtotalPlacement", "subtotalLabel", "sort", "memberOrder"),
            ["valuePlacement"] = Object(
                new Dictionary<string, object>
                {
                    ["measureId"] = BoundedString(128),
                    ["caption"] = BoundedString(128),
                    ["numberFormat"] = BoundedString(128),
                    ["periodSliceIds"] = Array(BoundedString(128), 64),
                    ["styleId"] = BoundedString(128)
                },
                "measureId", "caption", "numberFormat", "periodSliceIds", "styleId"),
            ["blockFilter"] = Object(
                new Dictionary<string, object>
                {
                    ["field"] = BoundedString(128),
                    ["selectedValues"] = Array(BoundedString(1024), 200),
                    ["includeBlank"] = Boolean()
                },
                "field", "selectedValues", "includeBlank"),
            ["periodSlice"] = Object(
                new Dictionary<string, object>
                {
                    ["id"] = BoundedString(128),
                    ["label"] = BoundedString(128),
                    ["kind"] = EnumString("current", "selected", "prior", "samePeriodPriorYear"),
                    ["selectedStart"] = BoundedString(10),
                    ["selectedEnd"] = BoundedString(10),
                    ["basedOnSliceId"] = BoundedString(128)
                },
                "id", "label", "kind", "selectedStart", "selectedEnd", "basedOnSliceId"),
            ["denseLayout"] = Object(
                new Dictionary<string, object>
                {
                    ["repeatRowLabels"] = Boolean(),
                    ["showRowGrandTotals"] = Boolean(),
                    ["showColumnGrandTotals"] = Boolean(),
                    ["insertBlankRows"] = Boolean(),
                    ["rowIndent"] = Integer(0, 15),
                    ["freezeHeaders"] = Boolean()
                },
                "repeatRowLabels", "showRowGrandTotals", "showColumnGrandTotals",
                "insertBlankRows", "rowIndent", "freezeHeaders"),
            ["grandTotals"] = Object(
                new Dictionary<string, object>
                {
                    ["showRows"] = Boolean(),
                    ["showColumns"] = Boolean(),
                    ["rowPlacement"] = EnumString("beforeMembers", "afterMembers"),
                    ["columnPlacement"] = EnumString("beforeMembers", "afterMembers"),
                    ["rowLabel"] = BoundedString(128),
                    ["columnLabel"] = BoundedString(128),
                    ["styleId"] = BoundedString(128)
                },
                "showRows", "showColumns", "rowPlacement", "columnPlacement",
                "rowLabel", "columnLabel", "styleId"),
            ["block"] = Object(
                new Dictionary<string, object>
                {
                    ["id"] = BoundedString(128),
                    ["title"] = BoundedString(128),
                    ["worksheetName"] = BoundedString(31),
                    ["anchorCell"] = BoundedString(10),
                    ["outputMode"] = EnumString("standardMatrix", "metricStack", "denseGrid"),
                    ["rows"] = Array(Reference("fieldPlacement"), 32),
                    ["columns"] = Array(Reference("fieldPlacement"), 16),
                    ["values"] = Array(Reference("valuePlacement"), 32),
                    ["filters"] = Array(Reference("blockFilter"), 32),
                    ["periodSlices"] = Array(Reference("periodSlice"), 64),
                    ["denseLayout"] = Reference("denseLayout"),
                    ["grandTotals"] = Reference("grandTotals"),
                    ["headerStyleId"] = BoundedString(128),
                    ["bodyStyleId"] = BoundedString(128),
                    ["subtotalStyleId"] = BoundedString(128),
                    ["grandTotalStyleId"] = BoundedString(128)
                },
                "id", "title", "worksheetName", "anchorCell", "outputMode",
                "rows", "columns", "values", "filters", "periodSlices",
                "denseLayout", "grandTotals", "headerStyleId", "bodyStyleId",
                "subtotalStyleId", "grandTotalStyleId"),
            ["style"] = Object(
                new Dictionary<string, object>
                {
                    ["id"] = BoundedString(128),
                    ["bold"] = Boolean(),
                    ["italic"] = Boolean(),
                    ["fontColor"] = BoundedString(7),
                    ["fillColor"] = BoundedString(7),
                    ["horizontalAlignment"] = EnumString("general", "left", "center", "right"),
                    ["numberFormat"] = BoundedString(128),
                    ["decimalPlaces"] = Integer(-1, 12),
                    ["topBorder"] = Boolean(),
                    ["bottomBorder"] = Boolean()
                },
                "id", "bold", "italic", "fontColor", "fillColor", "horizontalAlignment",
                "numberFormat", "decimalPlaces", "topBorder", "bottomBorder"),
            ["check"] = Object(
                new Dictionary<string, object>
                {
                    ["id"] = BoundedString(128),
                    ["kind"] = EnumString(
                        "totalPreservation", "noTruncation", "requiredValues", "nonNegative", "balance"),
                    ["measureId"] = BoundedString(128),
                    ["comparedMeasureId"] = BoundedString(128),
                    ["tolerance"] = MinimumNumber(0)
                },
                "id", "kind", "measureId", "comparedMeasureId", "tolerance")
        };

        var root = Object(
            new Dictionary<string, object>
            {
                ["version"] = ConstString("1.0"),
                ["measures"] = Array(Reference("measure"), 64, 1),
                ["blocks"] = Array(Reference("block"), 8, 1),
                ["styles"] = Array(Reference("style"), 32),
                ["checks"] = Array(Reference("check"), 32)
            },
            "version", "measures", "blocks", "styles", "checks");
        root["$defs"] = definitions;
        return root;
    }

    private static Dictionary<string, object> RatioExpression(string kind)
    {
        return Object(
            new Dictionary<string, object>
            {
                ["kind"] = ConstString(kind),
                ["numeratorMeasureId"] = BoundedString(128),
                ["denominatorMeasureId"] = BoundedString(128),
                ["onZero"] = ZeroEnum()
            },
            "kind", "numeratorMeasureId", "denominatorMeasureId", "onZero");
    }

    private static Dictionary<string, object> AggregateEnum()
    {
        return EnumString("sum", "count", "distinctCount", "average", "minimum", "maximum");
    }

    private static Dictionary<string, object> ZeroEnum()
    {
        return EnumString("blank", "zero", "error");
    }

    private static Dictionary<string, object> Object(
        Dictionary<string, object> properties,
        params string[] required)
    {
        return new Dictionary<string, object>
        {
            ["type"] = "object",
            ["properties"] = properties,
            ["required"] = required,
            ["additionalProperties"] = false
        };
    }

    private static Dictionary<string, object> Array(object items, int maximum, int? minimum = null)
    {
        var result = new Dictionary<string, object>
        {
            ["type"] = "array",
            ["items"] = items,
            ["maxItems"] = maximum
        };
        if (minimum.HasValue) result["minItems"] = minimum.Value;
        return result;
    }

    private static Dictionary<string, object> Reference(string name)
    {
        return new Dictionary<string, object> { ["$ref"] = "#/$defs/" + name };
    }

    private static Dictionary<string, object> BoundedString(int maximum)
    {
        return new Dictionary<string, object>
        {
            ["type"] = "string",
            ["maxLength"] = maximum
        };
    }

    private static Dictionary<string, object> EnumString(params string[] values)
    {
        return new Dictionary<string, object>
        {
            ["type"] = "string",
            ["enum"] = values
        };
    }

    private static Dictionary<string, object> ConstString(string value)
    {
        return new Dictionary<string, object>
        {
            ["type"] = "string",
            ["const"] = value
        };
    }

    private static Dictionary<string, object> Boolean()
    {
        return new Dictionary<string, object> { ["type"] = "boolean" };
    }

    private static Dictionary<string, object> Number()
    {
        return new Dictionary<string, object> { ["type"] = "number" };
    }

    private static Dictionary<string, object> NullableNumber()
    {
        return new Dictionary<string, object>
        {
            ["type"] = new[] { "number", "null" }
        };
    }

    private static Dictionary<string, object> MinimumNumber(decimal minimum)
    {
        return new Dictionary<string, object>
        {
            ["type"] = "number",
            ["minimum"] = minimum
        };
    }

    private static Dictionary<string, object> Integer(int minimum, int maximum)
    {
        return new Dictionary<string, object>
        {
            ["type"] = "integer",
            ["minimum"] = minimum,
            ["maximum"] = maximum
        };
    }
}
