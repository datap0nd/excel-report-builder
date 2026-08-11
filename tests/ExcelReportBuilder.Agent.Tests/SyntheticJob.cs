using ExcelReportBuilder.Agent.Configuration;
using ExcelReportBuilder.Agent.Models;

namespace ExcelReportBuilder.Agent.Tests;

internal static class SyntheticJob
{
    public static AgentJobRequest Create(int maxRepairCycles = 2)
    {
        return new AgentJobRequest
        {
            JobId = "job-synthetic-001",
            WorkbookId = "workbook-synthetic-001",
            UserPrompt = "Place Department in Rows and Net Amount in Values.",
            MaxRepairCycles = maxRepairCycles,
            Endpoint = new AgentEndpointSettings
            {
                BaseUrl = "http://127.0.0.1:1234",
                Model = AgentDefaults.Model,
            },
            Data = new AgentDataSnapshot
            {
                SourceDisplayName = "Synthetic extract",
                RowCount = 2,
                ReportingYear = 2026,
                Fields =
                {
                    new AgentField { Name = "Department", Type = AgentFieldType.Text },
                    new AgentField { Name = "Region", Type = AgentFieldType.Text },
                    new AgentField { Name = "Period", Type = AgentFieldType.Date },
                    new AgentField { Name = "Net Amount", Type = AgentFieldType.Number },
                },
                SampleRows =
                {
                    new AgentSampleRow
                    {
                        Values =
                        {
                            new AgentSampleValue { Field = "Department", Value = "Operations" },
                            new AgentSampleValue { Field = "Region", Value = "North" },
                            new AgentSampleValue { Field = "Period", Value = "2026-01-31" },
                            new AgentSampleValue { Field = "Net Amount", Value = "125.50" },
                        },
                    },
                    new AgentSampleRow
                    {
                        Values =
                        {
                            new AgentSampleValue { Field = "Department", Value = "Sales" },
                            new AgentSampleValue { Field = "Region", Value = "South" },
                            new AgentSampleValue { Field = "Period", Value = "2026-02-28" },
                            new AgentSampleValue { Field = "Net Amount", Value = "200.00" },
                        },
                    },
                },
            },
        };
    }

    public static AgentToolCall ValidReportSpecCall(string id = "call-1")
    {
        return new AgentToolCall
        {
            Id = id,
            Name = "propose_report_spec",
            ArgumentsJson =
                "{\"rows\":[\"Department\"],\"columns\":[\"Period\"]," +
                "\"values\":[{\"field\":\"Net Amount\",\"aggregation\":\"sum\"}]," +
                "\"filters\":[],\"subtotals\":[{\"field\":\"Department\",\"mode\":\"show\"}]," +
                "\"formatting\":[{\"field\":\"Net Amount\",\"numberStyle\":\"currency\",\"decimalPlaces\":2}]," +
                "\"ordering\":[{\"field\":\"Department\",\"direction\":\"ascending\",\"by\":\"label\"}]}",
        };
    }

    public static AgentToolCall ValidAdvancedReportSpecCall(string id = "advanced-spec-1")
    {
        return new AgentToolCall
        {
            Id = id,
            Name = "propose_report_spec",
            ArgumentsJson =
                "{\"version\":\"1.0\",\"measures\":[" +
                "{\"id\":\"net\",\"label\":\"Net amount\",\"valueType\":\"currency\",\"numberFormat\":\"$#,##0.00\",\"expression\":{\"kind\":\"aggregate\",\"field\":\"Net Amount\",\"aggregation\":\"sum\",\"periodSliceId\":\"\"}}," +
                "{\"id\":\"target\",\"label\":\"Target amount\",\"valueType\":\"currency\",\"numberFormat\":\"$#,##0.00\",\"expression\":{\"kind\":\"filteredAggregate\",\"field\":\"Net Amount\",\"aggregation\":\"sum\",\"periodSliceId\":\"\",\"filters\":[{\"field\":\"Region\",\"operator\":\"equal\",\"values\":[\"North\"]}]}}," +
                "{\"id\":\"achievement\",\"label\":\"Achievement\",\"valueType\":\"percentage\",\"numberFormat\":\"0.0%\",\"expression\":{\"kind\":\"safeDivide\",\"numeratorMeasureId\":\"net\",\"denominatorMeasureId\":\"target\",\"onZero\":\"blank\"}}]," +
                "\"blocks\":[" + AdvancedBlock("summary", "Summary", true) + "," +
                AdvancedBlock("detail", "Details", false) + "]," +
                "\"styles\":[{\"id\":\"header\",\"bold\":true,\"italic\":false,\"fontColor\":\"#FFFFFF\",\"fillColor\":\"#1F5D50\",\"horizontalAlignment\":\"center\",\"numberFormat\":\"\",\"decimalPlaces\":-1,\"topBorder\":false,\"bottomBorder\":true}]," +
                "\"checks\":[{\"id\":\"totals\",\"kind\":\"totalPreservation\",\"measureId\":\"net\",\"comparedMeasureId\":\"\",\"tolerance\":0.01},{\"id\":\"rows\",\"kind\":\"noTruncation\",\"measureId\":\"\",\"comparedMeasureId\":\"\",\"tolerance\":0}]}"
        };
    }

    private static string AdvancedBlock(string id, string sheet, bool includePrior)
    {
        string slices = includePrior
            ? "[{\"id\":\"current\",\"label\":\"Current\",\"kind\":\"current\",\"selectedStart\":\"2026-01-01\",\"selectedEnd\":\"2026-03-31\",\"basedOnSliceId\":\"\"},{\"id\":\"prior\",\"label\":\"Prior\",\"kind\":\"prior\",\"selectedStart\":\"\",\"selectedEnd\":\"\",\"basedOnSliceId\":\"current\"}]"
            : "[{\"id\":\"current\",\"label\":\"Current\",\"kind\":\"current\",\"selectedStart\":\"2026-01-01\",\"selectedEnd\":\"2026-03-31\",\"basedOnSliceId\":\"\"}]";
        string valueSlices = includePrior ? "[\"current\",\"prior\"]" : "[\"current\"]";
        return "{\"id\":\"" + id + "\",\"title\":\"Managed report\",\"worksheetName\":\"" + sheet +
            "\",\"anchorCell\":\"A1\",\"outputMode\":\"denseGrid\"," +
            "\"rows\":[{\"field\":\"Department\",\"caption\":\"Department\",\"subtotalMode\":\"automatic\",\"subtotalPlacement\":\"afterMembers\",\"subtotalLabel\":\"Total\",\"sort\":\"ascending\",\"memberOrder\":[]}]," +
            "\"columns\":[{\"field\":\"Period\",\"caption\":\"Period\",\"subtotalMode\":\"none\",\"subtotalPlacement\":\"afterMembers\",\"subtotalLabel\":\"\",\"sort\":\"ascending\",\"memberOrder\":[]}]," +
            "\"values\":[{\"measureId\":\"net\",\"caption\":\"Net amount\",\"numberFormat\":\"$#,##0.00\",\"periodSliceIds\":" + valueSlices + ",\"styleId\":\"\"},{\"measureId\":\"achievement\",\"caption\":\"Achievement\",\"numberFormat\":\"0.0%\",\"periodSliceIds\":" + valueSlices + ",\"styleId\":\"\"}]," +
            "\"filters\":[],\"periodSlices\":" + slices + "," +
            "\"denseLayout\":{\"repeatRowLabels\":true,\"showRowGrandTotals\":true,\"showColumnGrandTotals\":true,\"insertBlankRows\":false,\"rowIndent\":1,\"freezeHeaders\":true}," +
            "\"grandTotals\":{\"showRows\":true,\"showColumns\":true,\"rowPlacement\":\"afterMembers\",\"columnPlacement\":\"afterMembers\",\"rowLabel\":\"Grand Total\",\"columnLabel\":\"Grand Total\",\"styleId\":\"\"}," +
            "\"headerStyleId\":\"header\",\"bodyStyleId\":\"\",\"subtotalStyleId\":\"\",\"grandTotalStyleId\":\"\"}";
    }
}
