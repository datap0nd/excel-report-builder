using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.Json;
using ExcelReportBuilder.Agent.Models;

namespace ExcelReportBuilder.Agent.Tools;

internal static class AgentAdvancedReportProposalValidator
{
    private const int MaximumMeasures = 64;
    private const int MaximumBlocks = 8;

    public static bool Validate(
        JsonElement value,
        IReadOnlyDictionary<string, AgentField> fields)
    {
        if (!HasOnlyProperties(value, "version", "measures", "blocks", "styles", "checks") ||
            !TryString(value, "version", out var version) ||
            !string.Equals(version, "1.0", StringComparison.Ordinal) ||
            !TryArray(value, "measures", MaximumMeasures, out var measures) ||
            measures.GetArrayLength() == 0 ||
            !TryArray(value, "blocks", MaximumBlocks, out var blocks) ||
            blocks.GetArrayLength() == 0 ||
            !TryArray(value, "styles", 32, out var styles) ||
            !TryArray(value, "checks", 32, out var checks))
        {
            return false;
        }

        var measureIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var measure in measures.EnumerateArray())
        {
            if (!ValidateMeasure(measure, fields, measureIds)) return false;
        }

        var styleIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var style in styles.EnumerateArray())
        {
            if (!ValidateStyle(style, styleIds)) return false;
        }

        var blockIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var block in blocks.EnumerateArray())
        {
            if (!ValidateBlock(block, fields, measureIds, styleIds, blockIds)) return false;
        }

        var checkIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var check in checks.EnumerateArray())
        {
            if (!ValidateCheck(check, measureIds, checkIds)) return false;
        }

        return true;
    }

    private static bool ValidateMeasure(
        JsonElement value,
        IReadOnlyDictionary<string, AgentField> fields,
        ISet<string> measureIds)
    {
        if (!HasOnlyProperties(value, "id", "label", "valueType", "numberFormat", "expression") ||
            !TryIdentifier(value, "id", out var id) || !measureIds.Add(id!) ||
            !TryBoundedText(value, "label", 128, allowEmpty: false) ||
            !TryEnum(value, "valueType", "wholeNumber", "number", "currency", "percentage") ||
            !TryBoundedText(value, "numberFormat", 128, allowEmpty: true) ||
            !value.TryGetProperty("expression", out var expression) ||
            expression.ValueKind != JsonValueKind.Object)
        {
            return false;
        }

        return ValidateExpression(expression, fields);
    }

    private static bool ValidateExpression(
        JsonElement value,
        IReadOnlyDictionary<string, AgentField> fields)
    {
        if (!TryString(value, "kind", out var kind)) return false;
        switch (kind)
        {
            case "aggregate":
                return HasOnlyProperties(value, "kind", "field", "aggregation", "periodSliceId") &&
                    ValidateAggregateParts(value, fields);
            case "filteredAggregate":
                return HasOnlyProperties(value, "kind", "field", "aggregation", "periodSliceId", "filters") &&
                    ValidateAggregateParts(value, fields) &&
                    TryArray(value, "filters", 16, out var filters) &&
                    ValidateMeasureFilters(filters, fields);
            case "reference":
                return HasOnlyProperties(value, "kind", "measureId") &&
                    TryIdentifier(value, "measureId", out _);
            case "constant":
                return HasOnlyProperties(value, "kind", "value") &&
                    TryDecimal(value, "value", out _);
            case "binary":
                return HasOnlyProperties(
                           value,
                           "kind",
                           "operator",
                           "leftMeasureId",
                           "rightMeasureId",
                           "returnBlankOnZeroDenominator") &&
                    TryEnum(value, "operator", "add", "subtract", "multiply", "divide") &&
                    TryIdentifier(value, "leftMeasureId", out _) &&
                    TryIdentifier(value, "rightMeasureId", out _) &&
                    TryBoolean(value, "returnBlankOnZeroDenominator", out _);
            case "safeDivide":
            case "ratio":
            case "share":
                return HasOnlyProperties(
                           value,
                           "kind",
                           "numeratorMeasureId",
                           "denominatorMeasureId",
                           "onZero") &&
                    TryIdentifier(value, "numeratorMeasureId", out _) &&
                    TryIdentifier(value, "denominatorMeasureId", out _) &&
                    TryEnum(value, "onZero", "blank", "zero", "error");
            case "difference":
                return HasOnlyProperties(
                           value,
                           "kind",
                           "differenceKind",
                           "currentMeasureId",
                           "baselineMeasureId",
                           "onZero") &&
                    TryEnum(value, "differenceKind", "absolute", "percentage", "percentagePoints") &&
                    TryIdentifier(value, "currentMeasureId", out _) &&
                    TryIdentifier(value, "baselineMeasureId", out _) &&
                    TryEnum(value, "onZero", "blank", "zero", "error");
            default:
                return false;
        }
    }

    private static bool ValidateAggregateParts(
        JsonElement value,
        IReadOnlyDictionary<string, AgentField> fields)
    {
        if (!TryString(value, "field", out var fieldName) ||
            !fields.TryGetValue(fieldName!, out var field) ||
            !TryString(value, "aggregation", out var aggregation) ||
            !IsOneOf(aggregation, "sum", "count", "distinctCount", "average", "minimum", "maximum") ||
            !TryBoundedText(value, "periodSliceId", 128, allowEmpty: true))
        {
            return false;
        }

        return string.Equals(aggregation, "count", StringComparison.Ordinal) ||
            string.Equals(aggregation, "distinctCount", StringComparison.Ordinal) ||
            field.Type == AgentFieldType.Number;
    }

    private static bool ValidateMeasureFilters(
        JsonElement filters,
        IReadOnlyDictionary<string, AgentField> fields)
    {
        var filterFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var filter in filters.EnumerateArray())
        {
            if (!HasOnlyProperties(filter, "field", "operator", "values") ||
                !TryString(filter, "field", out var field) ||
                !fields.ContainsKey(field!) || !filterFields.Add(field!) ||
                !TryString(filter, "operator", out var filterOperator) ||
                !IsOneOf(
                    filterOperator,
                    "equal",
                    "notEqual",
                    "greaterThan",
                    "greaterThanOrEqual",
                    "lessThan",
                    "lessThanOrEqual",
                    "in",
                    "notIn",
                    "isBlank",
                    "isNotBlank") ||
                !TryArray(filter, "values", 50, out var values) ||
                !ValidateStringScalars(values))
            {
                return false;
            }

            int count = values.GetArrayLength();
            if ((IsOneOf(filterOperator, "isBlank", "isNotBlank") && count != 0) ||
                (!IsOneOf(filterOperator, "isBlank", "isNotBlank") && count == 0))
            {
                return false;
            }
        }

        return true;
    }

    private static bool ValidateBlock(
        JsonElement value,
        IReadOnlyDictionary<string, AgentField> fields,
        ISet<string> measureIds,
        ISet<string> styleIds,
        ISet<string> blockIds)
    {
        if (!HasOnlyProperties(
                value,
                "id",
                "title",
                "worksheetName",
                "anchorCell",
                "outputMode",
                "rows",
                "columns",
                "values",
                "filters",
                "periodSlices",
                "denseLayout",
                "grandTotals",
                "headerStyleId",
                "bodyStyleId",
                "subtotalStyleId",
                "grandTotalStyleId") ||
            !TryIdentifier(value, "id", out var id) || !blockIds.Add(id!) ||
            !TryBoundedText(value, "title", 128, allowEmpty: true) ||
            !TryString(value, "worksheetName", out var worksheetName) ||
            !IsWorksheetName(worksheetName) ||
            !TryString(value, "anchorCell", out var anchorCell) ||
            !IsCellAddress(anchorCell) ||
            !TryEnum(value, "outputMode", "standardMatrix", "metricStack", "denseGrid") ||
            !TryArray(value, "rows", 32, out var rows) || rows.GetArrayLength() == 0 ||
            !TryArray(value, "columns", 16, out var columns) ||
            !TryArray(value, "values", 32, out var values) || values.GetArrayLength() == 0 ||
            !TryArray(value, "filters", 32, out var filters) ||
            !TryArray(value, "periodSlices", 64, out var slices) ||
            !ValidateFieldPlacements(rows, fields) ||
            !ValidateFieldPlacements(columns, fields) ||
            !ValidateValuePlacements(values, measureIds) ||
            !ValidateBlockFilters(filters, fields) ||
            !ValidatePeriodSlices(slices) ||
            !value.TryGetProperty("denseLayout", out var denseLayout) ||
            !ValidateDenseLayout(denseLayout) ||
            !value.TryGetProperty("grandTotals", out var grandTotals) ||
            !ValidateGrandTotals(grandTotals) ||
            !ValidateStyleReference(value, "headerStyleId", styleIds) ||
            !ValidateStyleReference(value, "bodyStyleId", styleIds) ||
            !ValidateStyleReference(value, "subtotalStyleId", styleIds) ||
            !ValidateStyleReference(value, "grandTotalStyleId", styleIds))
        {
            return false;
        }

        return true;
    }

    private static bool ValidateFieldPlacements(
        JsonElement values,
        IReadOnlyDictionary<string, AgentField> fields)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var value in values.EnumerateArray())
        {
            if (!HasOnlyProperties(
                    value,
                    "field",
                    "caption",
                    "subtotalMode",
                    "subtotalPlacement",
                    "subtotalLabel",
                    "sort",
                    "memberOrder") ||
                !TryString(value, "field", out var field) ||
                !fields.ContainsKey(field!) || !seen.Add(field!) ||
                !TryBoundedText(value, "caption", 128, allowEmpty: true) ||
                !TryEnum(value, "subtotalMode", "none", "automatic") ||
                !TryEnum(value, "subtotalPlacement", "beforeMembers", "afterMembers") ||
                !TryBoundedText(value, "subtotalLabel", 128, allowEmpty: true) ||
                !TryEnum(value, "sort", "sourceOrder", "ascending", "descending") ||
                !TryArray(value, "memberOrder", 200, out var memberOrder) ||
                !ValidateStringScalars(memberOrder))
            {
                return false;
            }
        }

        return true;
    }

    private static bool ValidateValuePlacements(JsonElement values, ISet<string> measureIds)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var value in values.EnumerateArray())
        {
            if (!HasOnlyProperties(
                    value,
                    "measureId",
                    "caption",
                    "numberFormat",
                    "periodSliceIds",
                    "styleId") ||
                !TryIdentifier(value, "measureId", out var measureId) ||
                !measureIds.Contains(measureId!) || !seen.Add(measureId!) ||
                !TryBoundedText(value, "caption", 128, allowEmpty: true) ||
                !TryBoundedText(value, "numberFormat", 128, allowEmpty: true) ||
                !TryArray(value, "periodSliceIds", 64, out var slices) ||
                !ValidateIdentifiers(slices) ||
                !TryBoundedText(value, "styleId", 128, allowEmpty: true))
            {
                return false;
            }
        }

        return true;
    }

    private static bool ValidateBlockFilters(
        JsonElement filters,
        IReadOnlyDictionary<string, AgentField> fields)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var filter in filters.EnumerateArray())
        {
            if (!HasOnlyProperties(filter, "field", "selectedValues", "includeBlank") ||
                !TryString(filter, "field", out var field) ||
                !fields.ContainsKey(field!) || !seen.Add(field!) ||
                !TryArray(filter, "selectedValues", 200, out var selected) ||
                !ValidateStringScalars(selected) ||
                !TryBoolean(filter, "includeBlank", out _) ||
                (selected.GetArrayLength() == 0 && !filter.GetProperty("includeBlank").GetBoolean()))
            {
                return false;
            }
        }

        return true;
    }

    private static bool ValidatePeriodSlices(JsonElement slices)
    {
        var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var slice in slices.EnumerateArray())
        {
            if (!HasOnlyProperties(
                    slice,
                    "id",
                    "label",
                    "kind",
                    "selectedStart",
                    "selectedEnd",
                    "basedOnSliceId") ||
                !TryIdentifier(slice, "id", out var id) || !ids.Add(id!) ||
                !TryBoundedText(slice, "label", 128, allowEmpty: false) ||
                !TryString(slice, "kind", out var kind) ||
                !IsOneOf(kind, "current", "selected", "prior", "samePeriodPriorYear") ||
                !TryString(slice, "selectedStart", out var start) ||
                !TryString(slice, "selectedEnd", out var end) ||
                !TryString(slice, "basedOnSliceId", out var basedOn))
            {
                return false;
            }

            bool absolute = IsOneOf(kind, "current", "selected");
            if (absolute)
            {
                if (!TryDate(start, out var startDate) || !TryDate(end, out var endDate) ||
                    startDate > endDate || basedOn!.Length != 0)
                {
                    return false;
                }
            }
            else if (!IsIdentifier(basedOn) || start!.Length != 0 || end!.Length != 0)
            {
                return false;
            }
        }

        return true;
    }

    private static bool ValidateDenseLayout(JsonElement value)
    {
        return HasOnlyProperties(
                   value,
                   "repeatRowLabels",
                   "showRowGrandTotals",
                   "showColumnGrandTotals",
                   "insertBlankRows",
                   "rowIndent",
                   "freezeHeaders") &&
            TryBoolean(value, "repeatRowLabels", out _) &&
            TryBoolean(value, "showRowGrandTotals", out _) &&
            TryBoolean(value, "showColumnGrandTotals", out _) &&
            TryBoolean(value, "insertBlankRows", out _) &&
            TryInteger(value, "rowIndent", 0, 15, out _) &&
            TryBoolean(value, "freezeHeaders", out _);
    }

    private static bool ValidateGrandTotals(JsonElement value)
    {
        return HasOnlyProperties(
                   value,
                   "showRows",
                   "showColumns",
                   "rowPlacement",
                   "columnPlacement",
                   "rowLabel",
                   "columnLabel",
                   "styleId") &&
            TryBoolean(value, "showRows", out _) &&
            TryBoolean(value, "showColumns", out _) &&
            TryEnum(value, "rowPlacement", "beforeMembers", "afterMembers") &&
            TryEnum(value, "columnPlacement", "beforeMembers", "afterMembers") &&
            TryBoundedText(value, "rowLabel", 128, allowEmpty: false) &&
            TryBoundedText(value, "columnLabel", 128, allowEmpty: false) &&
            TryBoundedText(value, "styleId", 128, allowEmpty: true);
    }

    private static bool ValidateStyle(JsonElement value, ISet<string> styleIds)
    {
        if (!HasOnlyProperties(
                value,
                "id",
                "bold",
                "italic",
                "fontColor",
                "fillColor",
                "horizontalAlignment",
                "numberFormat",
                "decimalPlaces",
                "topBorder",
                "bottomBorder") ||
            !TryIdentifier(value, "id", out var id) || !styleIds.Add(id!) ||
            !TryBoolean(value, "bold", out _) ||
            !TryBoolean(value, "italic", out _) ||
            !TryColor(value, "fontColor") ||
            !TryColor(value, "fillColor") ||
            !TryEnum(value, "horizontalAlignment", "general", "left", "center", "right") ||
            !TryBoundedText(value, "numberFormat", 128, allowEmpty: true) ||
            !TryInteger(value, "decimalPlaces", -1, 12, out _) ||
            !TryBoolean(value, "topBorder", out _) ||
            !TryBoolean(value, "bottomBorder", out _))
        {
            return false;
        }

        return true;
    }

    private static bool ValidateCheck(
        JsonElement value,
        ISet<string> measureIds,
        ISet<string> checkIds)
    {
        if (!HasOnlyProperties(value, "id", "kind", "measureId", "comparedMeasureId", "tolerance") ||
            !TryIdentifier(value, "id", out var id) || !checkIds.Add(id!) ||
            !TryString(value, "kind", out var kind) ||
            !IsOneOf(kind, "totalPreservation", "noTruncation", "requiredValues", "nonNegative", "balance") ||
            !TryString(value, "measureId", out var measureId) ||
            !TryString(value, "comparedMeasureId", out var comparedMeasureId) ||
            !TryDecimal(value, "tolerance", out var tolerance) || tolerance < 0m)
        {
            return false;
        }

        bool needsMeasure = IsOneOf(kind, "totalPreservation", "requiredValues", "nonNegative", "balance");
        bool needsCompared = string.Equals(kind, "balance", StringComparison.Ordinal);
        return (needsMeasure ? IsIdentifier(measureId) && measureIds.Contains(measureId!) : measureId!.Length == 0) &&
            (needsCompared
                ? IsIdentifier(comparedMeasureId) && measureIds.Contains(comparedMeasureId!)
                : comparedMeasureId!.Length == 0);
    }

    private static bool ValidateStyleReference(
        JsonElement value,
        string property,
        ISet<string> styleIds)
    {
        return TryString(value, property, out var styleId) &&
            (styleId!.Length == 0 || (IsIdentifier(styleId) && styleIds.Contains(styleId)));
    }

    private static bool ValidateStringScalars(JsonElement values)
    {
        foreach (var value in values.EnumerateArray())
        {
            if (value.ValueKind != JsonValueKind.String ||
                !IsBoundedText(value.GetString(), 1024, allowEmpty: true))
            {
                return false;
            }
        }

        return true;
    }

    private static bool ValidateIdentifiers(JsonElement values)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var value in values.EnumerateArray())
        {
            if (value.ValueKind != JsonValueKind.String ||
                !IsIdentifier(value.GetString()) ||
                !seen.Add(value.GetString()!))
            {
                return false;
            }
        }

        return true;
    }

    private static bool TryArray(
        JsonElement value,
        string property,
        int maximum,
        out JsonElement result)
    {
        result = default;
        return value.TryGetProperty(property, out result) &&
            result.ValueKind == JsonValueKind.Array &&
            result.GetArrayLength() <= maximum;
    }

    private static bool TryString(JsonElement value, string property, out string? result)
    {
        result = null;
        if (!value.TryGetProperty(property, out var item) || item.ValueKind != JsonValueKind.String)
        {
            return false;
        }

        result = item.GetString();
        return result != null;
    }

    private static bool TryIdentifier(JsonElement value, string property, out string? result)
    {
        return TryString(value, property, out result) && IsIdentifier(result);
    }

    private static bool TryBoundedText(
        JsonElement value,
        string property,
        int maximum,
        bool allowEmpty)
    {
        return TryString(value, property, out var result) &&
            IsBoundedText(result, maximum, allowEmpty);
    }

    private static bool TryBoolean(JsonElement value, string property, out bool result)
    {
        result = false;
        if (!value.TryGetProperty(property, out var item) ||
            (item.ValueKind != JsonValueKind.True && item.ValueKind != JsonValueKind.False))
        {
            return false;
        }

        result = item.GetBoolean();
        return true;
    }

    private static bool TryInteger(
        JsonElement value,
        string property,
        int minimum,
        int maximum,
        out int result)
    {
        result = 0;
        return value.TryGetProperty(property, out var item) &&
            item.ValueKind == JsonValueKind.Number &&
            item.TryGetInt32(out result) &&
            result >= minimum && result <= maximum;
    }

    private static bool TryDecimal(JsonElement value, string property, out decimal result)
    {
        result = 0m;
        return value.TryGetProperty(property, out var item) &&
            item.ValueKind == JsonValueKind.Number &&
            item.TryGetDecimal(out result);
    }

    private static bool TryEnum(JsonElement value, string property, params string[] choices)
    {
        return TryString(value, property, out var result) && IsOneOf(result, choices);
    }

    private static bool TryColor(JsonElement value, string property)
    {
        if (!TryString(value, property, out var color) || color == null) return false;
        if (color.Length == 0) return true;
        if (color.Length != 7 || color[0] != '#') return false;
        for (var index = 1; index < color.Length; index++)
        {
            if (!Uri.IsHexDigit(color[index])) return false;
        }

        return true;
    }

    private static bool TryDate(string? value, out DateTime result)
    {
        return DateTime.TryParseExact(
            value,
            "yyyy-MM-dd",
            CultureInfo.InvariantCulture,
            DateTimeStyles.None,
            out result);
    }

    private static bool IsWorksheetName(string? value)
    {
        if (!IsBoundedText(value, 31, allowEmpty: false)) return false;
        return value!.IndexOfAny(new[] { '[', ']', ':', '*', '?', '/', '\\' }) < 0;
    }

    private static bool IsCellAddress(string? value)
    {
        if (string.IsNullOrWhiteSpace(value) || value!.Length > 10) return false;
        var index = 0;
        while (index < value.Length && value[index] >= 'A' && value[index] <= 'Z') index++;
        if (index == 0 || index > 3 || index == value.Length) return false;
        if (value[index] == '0') return false;
        for (; index < value.Length; index++)
        {
            if (value[index] < '0' || value[index] > '9') return false;
        }

        return true;
    }

    private static bool IsIdentifier(string? value)
    {
        if (!IsBoundedText(value, 128, allowEmpty: false)) return false;
        foreach (var character in value!)
        {
            var valid = (character >= 'a' && character <= 'z') ||
                        (character >= 'A' && character <= 'Z') ||
                        (character >= '0' && character <= '9') ||
                        character == '-' || character == '_' || character == '.';
            if (!valid) return false;
        }

        return true;
    }

    private static bool IsBoundedText(string? value, int maximum, bool allowEmpty)
    {
        if (value == null || value.Length > maximum || (!allowEmpty && string.IsNullOrWhiteSpace(value)))
        {
            return false;
        }

        foreach (var character in value)
        {
            if (char.IsControl(character)) return false;
        }

        return value.Length == 0 || value[0] != '=';
    }

    private static bool IsOneOf(string? value, params string[] choices)
    {
        if (value == null) return false;
        foreach (var choice in choices)
        {
            if (string.Equals(value, choice, StringComparison.Ordinal)) return true;
        }

        return false;
    }

    private static bool HasOnlyProperties(JsonElement value, params string[] properties)
    {
        if (value.ValueKind != JsonValueKind.Object) return false;
        var expected = new HashSet<string>(properties, StringComparer.Ordinal);
        var count = 0;
        foreach (var property in value.EnumerateObject())
        {
            count++;
            if (!expected.Contains(property.Name)) return false;
        }

        return count == expected.Count;
    }
}
