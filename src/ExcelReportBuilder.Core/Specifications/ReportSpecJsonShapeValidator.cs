using System;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelReportBuilder.Core.Specifications
{
    /// <summary>
    /// Enforces presence of schema-required JSON members before CLR property
    /// initializers can replace omitted data with valid-looking defaults.
    /// Semantic and cross-reference checks remain in ReportSpecValidator.
    /// </summary>
    internal static class ReportSpecJsonShapeValidator
    {
        public static void Validate(JObject root)
        {
            Require(root, "$", "schemaVersion", "id", "name", "ownershipId", "source", "transforms", "measures", "blocks", "styles", "checks");

            var source = Object(root, "source", "$.source");
            Require(source, "$.source", "kind", "workbookObjectName", "headerRowCount", "fingerprint");
            var fingerprint = Object(source, "fingerprint", "$.source.fingerprint");
            Require(fingerprint, "$.source.fingerprint", "algorithm", "headerHash", "columnCount");

            var periodMapping = OptionalObject(root, "periodMapping", "$.periodMapping");
            if (periodMapping != null)
            {
                Require(periodMapping, "$.periodMapping", "id", "kind", "keyColumns", "columns", "periodColumnName", "valueColumnName", "metricColumnName");
                var mappingKind = String(periodMapping, "kind", "$.periodMapping.kind");
                if (string.Equals(mappingKind, "longDateColumn", StringComparison.Ordinal))
                {
                    Require(periodMapping, "$.periodMapping", "dateColumn");
                }

                EachObject(Array(periodMapping, "columns", "$.periodMapping.columns"), "$.periodMapping.columns", (column, path) =>
                    Require(column, path, "sourceColumn", "month"));
            }

            EachObject(Array(root, "transforms", "$.transforms"), "$.transforms", ValidateTransform);
            EachObject(Array(root, "measures", "$.measures"), "$.measures", ValidateMeasure);
            EachObject(Array(root, "blocks", "$.blocks"), "$.blocks", ValidateBlock);
            EachObject(Array(root, "styles", "$.styles"), "$.styles", (style, path) =>
                Require(style, path, "id", "bold", "italic", "horizontalAlignment", "topBorder", "bottomBorder"));
            EachObject(Array(root, "checks", "$.checks"), "$.checks", (check, path) =>
                Require(check, path, "id", "kind", "tolerance"));
        }

        private static void ValidateTransform(JObject transform, string path)
        {
            Require(transform, path, "id", "kind");
            var kind = String(transform, "kind", path + ".kind");
            switch (kind)
            {
                case "selectColumns":
                case "keepColumns":
                case "removeColumns":
                case "reorderColumns":
                case "trimText":
                case "fillDown":
                    Require(transform, path, "columns");
                    break;
                case "renameColumn":
                    Require(transform, path, "from", "to");
                    break;
                case "changeColumnType":
                    Require(transform, path, "column", "dataType");
                    break;
                case "replaceValue":
                    Require(transform, path, "column", "find", "replaceWith");
                    ValidateScalar(Object(transform, "find", path + ".find"), path + ".find");
                    ValidateScalar(Object(transform, "replaceWith", path + ".replaceWith"), path + ".replaceWith");
                    break;
                case "normalizeBlanks":
                    Require(transform, path, "columns", "replacement", "treatWhitespaceAsBlank");
                    ValidateScalar(Object(transform, "replacement", path + ".replacement"), path + ".replacement");
                    break;
                case "normalizeErrors":
                    Require(transform, path, "columns", "replacement");
                    ValidateScalar(Object(transform, "replacement", path + ".replacement"), path + ".replacement");
                    break;
                case "mapValues":
                    Require(transform, path, "column", "entries");
                    EachObject(Array(transform, "entries", path + ".entries"), path + ".entries", (entry, entryPath) =>
                    {
                        Require(entry, entryPath, "from", "to");
                        ValidateScalar(Object(entry, "from", entryPath + ".from"), entryPath + ".from");
                        ValidateScalar(Object(entry, "to", entryPath + ".to"), entryPath + ".to");
                    });
                    break;
                case "filterRows":
                    Require(transform, path, "column", "operator");
                    var filterOperator = String(transform, "operator", path + ".operator");
                    if (!string.Equals(filterOperator, "isBlank", StringComparison.Ordinal)
                        && !string.Equals(filterOperator, "isNotBlank", StringComparison.Ordinal))
                    {
                        Require(transform, path, "value");
                        ValidateScalar(Object(transform, "value", path + ".value"), path + ".value");
                    }

                    break;
                case "excludeTotalRows":
                    Require(transform, path, "evidence", "requireAllEvidence");
                    EachObject(Array(transform, "evidence", path + ".evidence"), path + ".evidence", (evidence, evidencePath) =>
                    {
                        Require(evidence, evidencePath, "column", "matchKind", "values", "source", "observedMatchCount");
                        EachObject(Array(evidence, "values", evidencePath + ".values"), evidencePath + ".values", ValidateScalar);
                    });
                    break;
                case "derivePeriodParts":
                    Require(transform, path, "dateColumn", "columns");
                    EachObject(Array(transform, "columns", path + ".columns"), path + ".columns", (column, columnPath) =>
                        Require(column, columnPath, "part", "outputColumn"));
                    break;
                case "addArithmeticColumn":
                    Require(transform, path, "outputColumn", "operator", "left", "right", "resultType", "returnNullOnZeroDenominator");
                    ValidateOperand(Object(transform, "left", path + ".left"), path + ".left");
                    ValidateOperand(Object(transform, "right", path + ".right"), path + ".right");
                    break;
                case "normalizePeriods":
                    Require(transform, path, "periodMappingId");
                    break;
            }
        }

        private static void ValidateOperand(JObject operand, string path)
        {
            Require(operand, path, "kind");
            var kind = String(operand, "kind", path + ".kind");
            Require(operand, path, string.Equals(kind, "column", StringComparison.Ordinal) ? "column" : "number");
        }

        private static void ValidateMeasure(JObject measure, string path)
        {
            Require(measure, path, "id", "label", "valueType", "expression");
            ValidateExpression(Object(measure, "expression", path + ".expression"), path + ".expression");
        }

        private static void ValidateExpression(JObject expression, string path)
        {
            Require(expression, path, "kind", "resultType");
            switch (String(expression, "kind", path + ".kind"))
            {
                case "aggregate":
                    Require(expression, path, "field", "function");
                    break;
                case "filteredAggregate":
                    Require(expression, path, "field", "function", "filters");
                    EachObject(Array(expression, "filters", path + ".filters"), path + ".filters", (filter, filterPath) =>
                    {
                        Require(filter, filterPath, "field", "operator", "values");
                        EachObject(Array(filter, "values", filterPath + ".values"), filterPath + ".values", ValidateScalar);
                    });
                    break;
                case "weightedAggregate":
                    Require(expression, path, "numerator", "denominator", "onZero");
                    ValidateExpression(Object(expression, "numerator", path + ".numerator"), path + ".numerator");
                    ValidateExpression(Object(expression, "denominator", path + ".denominator"), path + ".denominator");
                    break;
                case "reference":
                    Require(expression, path, "measureId");
                    break;
                case "constant":
                    Require(expression, path, "value");
                    break;
                case "binary":
                    Require(expression, path, "operator", "left", "right", "returnBlankOnZeroDenominator");
                    ValidateExpression(Object(expression, "left", path + ".left"), path + ".left");
                    ValidateExpression(Object(expression, "right", path + ".right"), path + ".right");
                    break;
                case "safeDivide":
                    Require(expression, path, "numerator", "denominator", "onZero", "asPercentage");
                    ValidateExpression(Object(expression, "numerator", path + ".numerator"), path + ".numerator");
                    ValidateExpression(Object(expression, "denominator", path + ".denominator"), path + ".denominator");
                    break;
                case "ratio":
                    Require(expression, path, "numerator", "denominator", "onZero");
                    ValidateExpression(Object(expression, "numerator", path + ".numerator"), path + ".numerator");
                    ValidateExpression(Object(expression, "denominator", path + ".denominator"), path + ".denominator");
                    break;
                case "difference":
                    Require(expression, path, "differenceKind", "current", "baseline", "onZero");
                    ValidateExpression(Object(expression, "current", path + ".current"), path + ".current");
                    ValidateExpression(Object(expression, "baseline", path + ".baseline"), path + ".baseline");
                    break;
                case "share":
                    Require(expression, path, "part", "whole", "onZero", "scope");
                    ValidateExpression(Object(expression, "part", path + ".part"), path + ".part");
                    ValidateExpression(Object(expression, "whole", path + ".whole"), path + ".whole");
                    break;
            }
        }

        private static void ValidateBlock(JObject block, string path)
        {
            Require(block, path, "id", "ownershipId", "worksheetName", "anchorCell", "outputMode", "ownedExtent", "layout", "periodSlices", "headers", "spacers");
            var extent = Object(block, "ownedExtent", path + ".ownedExtent");
            Require(extent, path + ".ownedExtent", "rowCount", "columnCount");
            ValidateLayout(Object(block, "layout", path + ".layout"), path + ".layout");
            EachObject(Array(block, "periodSlices", path + ".periodSlices"), path + ".periodSlices", (slice, slicePath) =>
            {
                Require(slice, slicePath, "id", "label", "kind");
                var kind = String(slice, "kind", slicePath + ".kind");
                if (string.Equals(kind, "current", StringComparison.Ordinal)
                    || string.Equals(kind, "selected", StringComparison.Ordinal))
                {
                    Require(slice, slicePath, "selectedStart", "selectedEnd");
                }
                else if (string.Equals(kind, "prior", StringComparison.Ordinal)
                    || string.Equals(kind, "samePeriodPriorYear", StringComparison.Ordinal))
                {
                    Require(slice, slicePath, "basedOnSliceId");
                }
            });
            EachObject(Array(block, "headers", path + ".headers"), path + ".headers", (header, headerPath) =>
                Require(header, headerPath, "text", "relativeRow", "relativeColumn", "columnSpan"));
            EachObject(Array(block, "spacers", path + ".spacers"), path + ".spacers", (spacer, spacerPath) =>
                Require(spacer, spacerPath, "axis", "beforeLevel", "count"));
        }

        private static void ValidateLayout(JObject layout, string path)
        {
            Require(layout, path, "rows", "columns", "values", "filters", "denseLayout", "grandTotals");
            EachObject(Array(layout, "rows", path + ".rows"), path + ".rows", ValidateField);
            EachObject(Array(layout, "columns", path + ".columns"), path + ".columns", ValidateField);
            EachObject(Array(layout, "values", path + ".values"), path + ".values", (value, valuePath) =>
                Require(value, valuePath, "measureId", "periodSliceIds"));
            EachObject(Array(layout, "filters", path + ".filters"), path + ".filters", (filter, filterPath) =>
            {
                Require(filter, filterPath, "field", "selectedValues", "includeBlank");
                EachObject(Array(filter, "selectedValues", filterPath + ".selectedValues"), filterPath + ".selectedValues", ValidateScalar);
            });
            var dense = Object(layout, "denseLayout", path + ".denseLayout");
            Require(dense, path + ".denseLayout", "repeatRowLabels", "showRowGrandTotals", "showColumnGrandTotals", "insertBlankRows", "rowIndent", "freezeHeaders");
            var totals = Object(layout, "grandTotals", path + ".grandTotals");
            Require(totals, path + ".grandTotals", "showRows", "showColumns", "rowPlacement", "columnPlacement", "rowLabel", "columnLabel");
        }

        private static void ValidateField(JObject field, string path)
        {
            Require(field, path, "field", "subtotals", "sort", "memberOrder", "groupBuckets");
            var subtotals = Object(field, "subtotals", path + ".subtotals");
            Require(subtotals, path + ".subtotals", "mode", "placement");
            EachObject(Array(field, "memberOrder", path + ".memberOrder"), path + ".memberOrder", ValidateScalar);
            EachObject(Array(field, "groupBuckets", path + ".groupBuckets"), path + ".groupBuckets", (bucket, bucketPath) =>
            {
                Require(bucket, bucketPath, "id", "label", "members", "includeUnmatched");
                EachObject(Array(bucket, "members", bucketPath + ".members"), bucketPath + ".members", ValidateScalar);
            });
            var topN = OptionalObject(field, "topN", path + ".topN");
            if (topN != null)
            {
                Require(topN, path + ".topN", "count", "measureId", "direction", "includeOthers", "othersLabel");
            }
        }

        private static void ValidateScalar(JObject scalar, string path)
        {
            Require(scalar, path, "kind");
            switch (String(scalar, "kind", path + ".kind"))
            {
                case "text":
                    Require(scalar, path, "text");
                    break;
                case "number":
                    Require(scalar, path, "number");
                    break;
                case "boolean":
                    Require(scalar, path, "boolean");
                    break;
                case "date":
                case "dateTime":
                    Require(scalar, path, "temporal");
                    break;
            }
        }

        private static void EachObject(JArray array, string path, Action<JObject, string> action)
        {
            for (var index = 0; index < array.Count; index++)
            {
                var itemPath = path + "[" + index + "]";
                var item = array[index] as JObject;
                if (item == null)
                {
                    throw new JsonSerializationException(itemPath + " must be a non-null object.");
                }

                action(item, itemPath);
            }
        }

        private static void Require(JObject value, string path, params string[] names)
        {
            foreach (var name in names)
            {
                var property = value.Property(name, StringComparison.Ordinal);
                if (property == null || property.Value.Type == JTokenType.Null)
                {
                    throw new JsonSerializationException(path + "." + name + " is required and cannot be null.");
                }
            }
        }

        private static JObject Object(JObject parent, string name, string path)
        {
            Require(parent, path.Substring(0, Math.Max(1, path.LastIndexOf('.'))), name);
            var result = parent.Property(name, StringComparison.Ordinal)!.Value as JObject;
            if (result == null)
            {
                throw new JsonSerializationException(path + " must be an object.");
            }

            return result;
        }

        private static JObject? OptionalObject(JObject parent, string name, string path)
        {
            var property = parent.Property(name, StringComparison.Ordinal);
            if (property == null)
            {
                return null;
            }

            var result = property.Value as JObject;
            if (result == null)
            {
                throw new JsonSerializationException(path + " must be a non-null object.");
            }

            return result;
        }

        private static JArray Array(JObject parent, string name, string path)
        {
            Require(parent, path.Substring(0, Math.Max(1, path.LastIndexOf('.'))), name);
            var result = parent.Property(name, StringComparison.Ordinal)!.Value as JArray;
            if (result == null)
            {
                throw new JsonSerializationException(path + " must be an array.");
            }

            return result;
        }

        private static string String(JObject parent, string name, string path)
        {
            Require(parent, path.Substring(0, Math.Max(1, path.LastIndexOf('.'))), name);
            var token = parent.Property(name, StringComparison.Ordinal)!.Value;
            if (token.Type != JTokenType.String)
            {
                throw new JsonSerializationException(path + " must be a string.");
            }

            return token.Value<string>() ?? string.Empty;
        }
    }
}
