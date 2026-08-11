using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using ExcelReportBuilder.Core.Measures;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Core.Transforms;

namespace ExcelReportBuilder.AddIn.Host
{
    /// <summary>
    /// Converts the bounded manual calculated-metric editor into the same typed
    /// measure graph used by the report planner. The conversion is atomic: all
    /// entries are validated before any part of the report specification changes.
    /// </summary>
    internal sealed class ManualCalculatedMetricTranslator
    {
        private const int MaximumMeasures = 128;
        private const int MaximumValues = 128;
        private const int MaximumTransforms = 100;
        private const int MaximumFiltersPerMetric = 32;
        private const int MaximumDetailsLength = 8192;

        public void Append(
            ReportSpecV1 specification,
            ReportBlockSpec block,
            IReadOnlyList<ManualCalculatedMetricSnapshot> snapshots,
            IReadOnlyCollection<string>? sourceFields = null)
        {
            if (specification == null) throw new ArgumentNullException(nameof(specification));
            if (block == null) throw new ArgumentNullException(nameof(block));
            if (snapshots == null) throw new ArgumentNullException(nameof(snapshots));
            if (specification.Measures == null)
            {
                throw new InvalidOperationException("The report Measures collection is required.");
            }

            if (specification.Transforms == null)
            {
                throw new InvalidOperationException("The report transformation collection is required.");
            }

            if (block.Layout == null || block.Layout.Values == null)
            {
                throw new InvalidOperationException("The report block Values collection is required.");
            }

            if (snapshots.Count == 0)
            {
                return;
            }

            if (block.OutputMode != ReportOutputMode.DenseGrid)
            {
                throw new InvalidOperationException(
                    "Calculated metrics require a dense output block in this executor version.");
            }

            if (specification.Measures.Count + snapshots.Count > MaximumMeasures)
            {
                throw new InvalidOperationException("A report may contain at most 128 measures.");
            }

            if (block.Layout.Values.Count + snapshots.Count > MaximumValues)
            {
                throw new InvalidOperationException("A report block may contain at most 128 Values.");
            }

            var existingMeasures = BuildExistingMeasureIndex(specification.Measures);
            var existingIds = new HashSet<string>(
                specification.Measures.Select(measure => measure.Id),
                StringComparer.OrdinalIgnoreCase);
            var existingLabels = new HashSet<string>(
                specification.Measures.Select(measure => measure.Label),
                StringComparer.OrdinalIgnoreCase);
            var transformIds = new HashSet<string>(
                specification.Transforms.Select(transform => transform.Id),
                StringComparer.OrdinalIgnoreCase);
            var knownFields = BuildKnownFieldIndex(specification, block, sourceFields);
            var fieldTypes = BuildFieldTypeIndex(specification.Measures);
            var layoutFields = new HashSet<string>(
                block.Layout.Rows.Select(item => item.Field)
                    .Concat(block.Layout.Columns.Select(item => item.Field))
                    .Concat(block.Layout.Filters.Select(item => item.Field)),
                StringComparer.OrdinalIgnoreCase);

            var pendingMeasures = new List<MeasureDefinition>(snapshots.Count);
            var pendingValues = new List<ValuePlacementSpec>(snapshots.Count);
            var pendingTransforms = new List<TransformStep>();

            for (var index = 0; index < snapshots.Count; index++)
            {
                ManualCalculatedMetricSnapshot? snapshot = snapshots[index];
                if (snapshot == null)
                {
                    throw new InvalidOperationException("A calculated metric entry cannot be null.");
                }

                string label = ValidateLabel(snapshot.Label);
                if (!existingLabels.Add(label))
                {
                    throw new InvalidOperationException(
                        "Calculated metric labels must be unique. Duplicate label: '" + label + "'.");
                }

                string kind = NormalizeToken(snapshot.Kind);
                if ((kind == "share of parent" || kind == "share of report total") &&
                    block.Layout.Rows.Count == 0)
                {
                    throw new InvalidOperationException(
                        "Share calculations require at least one field in Rows.");
                }

                string measureId = CreateStableIdentifier("manual_metric", kind + "\n" + label);
                if (!existingIds.Add(measureId))
                {
                    throw new InvalidOperationException(
                        "The calculated metric produces an ID already used by this report: '" + measureId + "'.");
                }

                string numberFormat = ValidateNumberFormat(snapshot.NumberFormat);
                MetricTranslation translated = Translate(
                    snapshot,
                    kind,
                    measureId,
                    existingMeasures,
                    knownFields,
                    fieldTypes,
                    layoutFields,
                    transformIds);

                pendingTransforms.AddRange(translated.Transforms);
                var definition = new MeasureDefinition
                {
                    Id = measureId,
                    Label = label,
                    ValueType = translated.ValueType,
                    NumberFormat = numberFormat,
                    Expression = translated.Expression
                };
                pendingMeasures.Add(definition);
                existingMeasures.Add(definition);
                pendingValues.Add(new ValuePlacementSpec
                {
                    MeasureId = measureId,
                    Caption = label,
                    NumberFormat = numberFormat
                });
            }

            if (specification.Transforms.Count + pendingTransforms.Count > MaximumTransforms)
            {
                throw new InvalidOperationException("A report may contain at most 100 transformations.");
            }

            specification.Transforms.AddRange(pendingTransforms);
            specification.Measures.AddRange(pendingMeasures);
            block.Layout.Values.AddRange(pendingValues);
        }

        private static MetricTranslation Translate(
            ManualCalculatedMetricSnapshot snapshot,
            string kind,
            string measureId,
            ExistingMeasureIndex measures,
            Dictionary<string, string> knownFields,
            Dictionary<string, MeasureValueType> fieldTypes,
            HashSet<string> layoutFields,
            HashSet<string> transformIds)
        {
            if (kind == "weighted average")
            {
                DemandNoDetails(snapshot, kind);
                return TranslateWeightedAverage(
                    snapshot,
                    measureId,
                    knownFields,
                    fieldTypes,
                    transformIds);
            }

            if (kind == "filtered aggregate")
            {
                return TranslateFilteredAggregate(snapshot, knownFields, fieldTypes, layoutFields);
            }

            DemandNoDetails(snapshot, kind);
            MeasureDefinition primary = measures.Resolve(snapshot.Primary, "Primary");
            MeasureDefinition? secondary = null;
            if (kind == "share of parent" || kind == "share of report total")
            {
                secondary = string.IsNullOrWhiteSpace(snapshot.Secondary)
                    ? primary
                    : measures.Resolve(snapshot.Secondary, "Secondary");
            }
            else
            {
                secondary = measures.Resolve(snapshot.Secondary, "Secondary");
            }

            switch (kind)
            {
                case "add":
                case "subtract":
                case "multiply":
                    return TranslateBinary(kind, primary, secondary);
                case "safe divide":
                    return TranslateSafeDivide(primary, secondary);
                case "ratio":
                    return TranslateRatio(primary, secondary);
                case "difference":
                    return TranslateDifference(primary, secondary, DifferenceKind.Absolute);
                case "percentage change":
                    return TranslateDifference(primary, secondary, DifferenceKind.Percentage);
                case "percentage point difference":
                    return TranslateDifference(primary, secondary, DifferenceKind.PercentagePoints);
                case "share of parent":
                    return TranslateShare(primary, secondary, ShareDenominatorScope.Parent);
                case "share of report total":
                    return TranslateShare(primary, secondary, ShareDenominatorScope.FilteredReportTotal);
                default:
                    throw new InvalidOperationException(
                        "Calculated metric kind '" + snapshot.Kind + "' is not supported.");
            }
        }

        private static MetricTranslation TranslateBinary(
            string kind,
            MeasureDefinition primary,
            MeasureDefinition secondary)
        {
            BinaryMeasureOperator operation;
            if (kind == "subtract")
            {
                operation = BinaryMeasureOperator.Subtract;
            }
            else if (kind == "multiply")
            {
                operation = BinaryMeasureOperator.Multiply;
            }
            else
            {
                operation = BinaryMeasureOperator.Add;
            }

            MeasureValueType resultType = InferBinaryType(
                operation,
                primary.ValueType,
                secondary.ValueType);
            return MetricTranslation.ForExpression(
                resultType,
                new BinaryMeasureExpression
                {
                    Operator = operation,
                    Left = Reference(primary),
                    Right = Reference(secondary),
                    ReturnBlankOnZeroDenominator = true,
                    ResultType = resultType
                });
        }

        private static MetricTranslation TranslateSafeDivide(
            MeasureDefinition primary,
            MeasureDefinition secondary)
        {
            MeasureValueType resultType = InferDivisionType(primary.ValueType, secondary.ValueType);
            return MetricTranslation.ForExpression(
                resultType,
                new SafeDivideMeasureExpression
                {
                    Numerator = Reference(primary),
                    Denominator = Reference(secondary),
                    OnZero = ZeroDenominatorBehavior.Blank,
                    AsPercentage = false,
                    ResultType = resultType
                });
        }

        private static MetricTranslation TranslateRatio(
            MeasureDefinition primary,
            MeasureDefinition secondary)
        {
            MeasureValueType resultType = InferDivisionType(primary.ValueType, secondary.ValueType);
            return MetricTranslation.ForExpression(
                resultType,
                new RatioMeasureExpression
                {
                    Numerator = Reference(primary),
                    Denominator = Reference(secondary),
                    OnZero = ZeroDenominatorBehavior.Blank,
                    ResultType = resultType
                });
        }

        private static MetricTranslation TranslateDifference(
            MeasureDefinition primary,
            MeasureDefinition secondary,
            DifferenceKind differenceKind)
        {
            MeasureValueType resultType = InferDifferenceType(
                differenceKind,
                primary.ValueType,
                secondary.ValueType);
            return MetricTranslation.ForExpression(
                resultType,
                new DifferenceMeasureExpression
                {
                    DifferenceKind = differenceKind,
                    Current = Reference(primary),
                    Baseline = Reference(secondary),
                    OnZero = ZeroDenominatorBehavior.Blank,
                    ResultType = resultType
                });
        }

        private static MetricTranslation TranslateShare(
            MeasureDefinition primary,
            MeasureDefinition secondary,
            ShareDenominatorScope scope)
        {
            DemandComparable(primary.ValueType, secondary.ValueType, "A share");
            return MetricTranslation.ForExpression(
                MeasureValueType.Percentage,
                new ShareMeasureExpression
                {
                    Part = Reference(primary),
                    Whole = Reference(secondary),
                    OnZero = ZeroDenominatorBehavior.Blank,
                    Scope = scope,
                    ResultType = MeasureValueType.Percentage
                });
        }

        private static MetricTranslation TranslateWeightedAverage(
            ManualCalculatedMetricSnapshot snapshot,
            string measureId,
            Dictionary<string, string> knownFields,
            Dictionary<string, MeasureValueType> fieldTypes,
            HashSet<string> transformIds)
        {
            string valueField = ResolveSourceField(snapshot.Primary, "Primary", knownFields);
            string weightField = ResolveSourceField(snapshot.Secondary, "Secondary", knownFields);
            if (string.Equals(valueField, weightField, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "Weighted average requires different value and weight source fields.");
            }

            string productField = CreateStableIdentifier("ERBWeighted", measureId);
            if (knownFields.ContainsKey(productField))
            {
                throw new InvalidOperationException(
                    "The weighted-average product column conflicts with an existing source field.");
            }

            string transformId = CreateStableIdentifier("weighted_product", measureId);
            if (!transformIds.Add(transformId))
            {
                throw new InvalidOperationException(
                    "The weighted-average transformation ID conflicts with an existing transformation.");
            }

            MeasureValueType resultType = ResolveWeightedResultType(valueField, fieldTypes);
            var numerator = new AggregateMeasureExpression
            {
                Field = productField,
                Function = AggregateFunction.Sum,
                ResultType = resultType
            };
            var denominator = new AggregateMeasureExpression
            {
                Field = weightField,
                Function = AggregateFunction.Sum,
                ResultType = MeasureValueType.Number
            };

            knownFields.Add(productField, productField);
            return new MetricTranslation
            {
                ValueType = resultType,
                Expression = new WeightedAggregateMeasureExpression
                {
                    Numerator = numerator,
                    Denominator = denominator,
                    OnZero = ZeroDenominatorBehavior.Blank,
                    ResultType = resultType
                },
                Transforms =
                {
                    new AddArithmeticColumnTransform
                    {
                        Id = transformId,
                        OutputColumn = productField,
                        Operator = ArithmeticOperator.Multiply,
                        Left = new ArithmeticOperand
                        {
                            Kind = ArithmeticOperandKind.Column,
                            Column = valueField
                        },
                        Right = new ArithmeticOperand
                        {
                            Kind = ArithmeticOperandKind.Column,
                            Column = weightField
                        },
                        ResultType = ColumnDataType.DecimalNumber,
                        ReturnNullOnZeroDenominator = true
                    }
                }
            };
        }

        private static MetricTranslation TranslateFilteredAggregate(
            ManualCalculatedMetricSnapshot snapshot,
            Dictionary<string, string> knownFields,
            Dictionary<string, MeasureValueType> fieldTypes,
            HashSet<string> layoutFields)
        {
            string valueField = ResolveSourceField(snapshot.Primary, "Primary", knownFields);
            AggregateFunction function = ParseAggregate(snapshot.Secondary);
            List<MeasureFilterSpec> filters = ParseExactFilters(snapshot.Details, knownFields);
            foreach (MeasureFilterSpec filter in filters)
            {
                if (layoutFields.Contains(filter.Field))
                {
                    throw new InvalidOperationException(
                        "A filtered aggregate cannot filter a field already placed in Rows, Columns, or Filters in the same block.");
                }
            }

            MeasureValueType resultType = ResolveAggregateType(valueField, function, fieldTypes);
            var expression = new FilteredAggregateMeasureExpression
            {
                Field = valueField,
                Function = function,
                ResultType = resultType
            };
            expression.Filters.AddRange(filters);
            return MetricTranslation.ForExpression(resultType, expression);
        }

        private static List<MeasureFilterSpec> ParseExactFilters(
            string details,
            Dictionary<string, string> knownFields)
        {
            if (string.IsNullOrWhiteSpace(details))
            {
                throw new InvalidOperationException(
                    "Filtered aggregate Details must use Field=Value;Field=Value syntax.");
            }

            if (details.Length > MaximumDetailsLength || details.Any(char.IsControl))
            {
                throw new InvalidOperationException("Filtered aggregate Details exceed the bounded format.");
            }

            string[] clauses = details.Split(new[] { ';' }, StringSplitOptions.None);
            if (clauses.Length > MaximumFiltersPerMetric)
            {
                throw new InvalidOperationException("A filtered aggregate may contain at most 32 filters.");
            }

            var result = new List<MeasureFilterSpec>(clauses.Length);
            var seenFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string rawClause in clauses)
            {
                string clause = rawClause.Trim();
                int separator = clause.IndexOf('=');
                if (separator <= 0 || separator == clause.Length - 1)
                {
                    throw new InvalidOperationException(
                        "Filtered aggregate Details must use Field=Value;Field=Value syntax.");
                }

                string field = ResolveSourceField(
                    clause.Substring(0, separator).Trim(),
                    "Filter field",
                    knownFields);
                string value = clause.Substring(separator + 1).Trim();
                if (value.Length == 0 || value.Length > 1024 || value.Any(char.IsControl))
                {
                    throw new InvalidOperationException(
                        "Each filtered aggregate value must contain 1 to 1,024 non-control characters.");
                }

                if (!seenFields.Add(field))
                {
                    throw new InvalidOperationException(
                        "Filtered aggregate Details cannot repeat the same filter field.");
                }

                result.Add(new MeasureFilterSpec
                {
                    Field = field,
                    Operator = MeasureFilterOperator.Equal,
                    Values = { ScalarValue.FromText(value) }
                });
            }

            return result;
        }

        private static ExistingMeasureIndex BuildExistingMeasureIndex(
            IReadOnlyList<MeasureDefinition> measures)
        {
            var byId = new Dictionary<string, MeasureDefinition>(StringComparer.OrdinalIgnoreCase);
            var byLabel = new Dictionary<string, List<MeasureDefinition>>(StringComparer.OrdinalIgnoreCase);
            foreach (MeasureDefinition? measure in measures)
            {
                if (measure == null || string.IsNullOrWhiteSpace(measure.Id))
                {
                    throw new InvalidOperationException(
                        "Existing measures must have non-blank IDs before calculated metrics are added.");
                }

                if (byId.ContainsKey(measure.Id))
                {
                    throw new InvalidOperationException("Existing measure IDs must be unique.");
                }

                byId.Add(measure.Id, measure);
                string label = measure.Label ?? string.Empty;
                if (!byLabel.TryGetValue(label, out List<MeasureDefinition>? matches))
                {
                    matches = new List<MeasureDefinition>();
                    byLabel.Add(label, matches);
                }

                matches.Add(measure);
            }

            return new ExistingMeasureIndex(byId, byLabel);
        }

        private static Dictionary<string, string> BuildKnownFieldIndex(
            ReportSpecV1 specification,
            ReportBlockSpec block,
            IReadOnlyCollection<string>? sourceFields)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (sourceFields != null)
            {
                foreach (string field in sourceFields)
                {
                    AddKnownField(result, field);
                }
            }

            foreach (FieldPlacementSpec placement in block.Layout.Rows.Concat(block.Layout.Columns))
            {
                AddKnownField(result, placement.Field);
            }

            foreach (FilterPlacementSpec placement in block.Layout.Filters)
            {
                AddKnownField(result, placement.Field);
            }

            foreach (MeasureDefinition measure in specification.Measures)
            {
                CollectExpressionFields(measure.Expression, result);
            }

            CollectPeriodFields(specification.PeriodMapping, result);
            foreach (TransformStep transform in specification.Transforms)
            {
                CollectTransformFields(transform, result);
            }

            return result;
        }

        private static Dictionary<string, MeasureValueType> BuildFieldTypeIndex(
            IReadOnlyList<MeasureDefinition> measures)
        {
            var result = new Dictionary<string, MeasureValueType>(StringComparer.OrdinalIgnoreCase);
            foreach (MeasureDefinition measure in measures)
            {
                if (measure.Expression is AggregateMeasureExpression aggregate &&
                    aggregate.Function != AggregateFunction.Count &&
                    aggregate.Function != AggregateFunction.DistinctCount)
                {
                    AddFieldType(result, aggregate.Field, aggregate.ResultType);
                }
                else if (measure.Expression is FilteredAggregateMeasureExpression filtered &&
                         filtered.Function != AggregateFunction.Count &&
                         filtered.Function != AggregateFunction.DistinctCount)
                {
                    AddFieldType(result, filtered.Field, filtered.ResultType);
                }
            }

            return result;
        }

        private static void AddFieldType(
            Dictionary<string, MeasureValueType> fieldTypes,
            string field,
            MeasureValueType type)
        {
            if (string.IsNullOrWhiteSpace(field))
            {
                return;
            }

            if (!fieldTypes.TryGetValue(field, out MeasureValueType existing))
            {
                fieldTypes.Add(field, type);
            }
            else if (existing != type && IsPlainNumber(existing) && IsPlainNumber(type))
            {
                fieldTypes[field] = MeasureValueType.Number;
            }
        }

        private static void CollectExpressionFields(
            MeasureExpression? expression,
            Dictionary<string, string> fields)
        {
            switch (expression)
            {
                case AggregateMeasureExpression aggregate:
                    AddKnownField(fields, aggregate.Field);
                    break;
                case FilteredAggregateMeasureExpression filtered:
                    AddKnownField(fields, filtered.Field);
                    foreach (MeasureFilterSpec filter in filtered.Filters)
                    {
                        AddKnownField(fields, filter.Field);
                    }

                    break;
                case WeightedAggregateMeasureExpression weighted:
                    CollectExpressionFields(weighted.Numerator, fields);
                    CollectExpressionFields(weighted.Denominator, fields);
                    break;
                case BinaryMeasureExpression binary:
                    CollectExpressionFields(binary.Left, fields);
                    CollectExpressionFields(binary.Right, fields);
                    break;
                case SafeDivideMeasureExpression divide:
                    CollectExpressionFields(divide.Numerator, fields);
                    CollectExpressionFields(divide.Denominator, fields);
                    break;
                case RatioMeasureExpression ratio:
                    CollectExpressionFields(ratio.Numerator, fields);
                    CollectExpressionFields(ratio.Denominator, fields);
                    break;
                case DifferenceMeasureExpression difference:
                    CollectExpressionFields(difference.Current, fields);
                    CollectExpressionFields(difference.Baseline, fields);
                    break;
                case ShareMeasureExpression share:
                    CollectExpressionFields(share.Part, fields);
                    CollectExpressionFields(share.Whole, fields);
                    break;
            }
        }

        private static void CollectPeriodFields(
            PeriodMappingSpec? mapping,
            Dictionary<string, string> fields)
        {
            if (mapping == null)
            {
                return;
            }

            AddKnownField(fields, mapping.DateColumn);
            foreach (string field in mapping.KeyColumns)
            {
                AddKnownField(fields, field);
            }

            foreach (PeriodColumnMapping column in mapping.Columns)
            {
                AddKnownField(fields, column.SourceColumn);
            }

            AddKnownField(fields, mapping.PeriodColumnName);
            AddKnownField(fields, mapping.ValueColumnName);
            if (mapping.Kind == PeriodMappingKind.MetricMonthHeaders)
            {
                AddKnownField(fields, mapping.MetricColumnName);
            }
        }

        private static void CollectTransformFields(
            TransformStep transform,
            Dictionary<string, string> fields)
        {
            switch (transform)
            {
                case KeepColumnsTransform keep:
                    AddKnownFields(fields, keep.Columns);
                    break;
                case SelectColumnsTransform select:
                    AddKnownFields(fields, select.Columns);
                    break;
                case RemoveColumnsTransform remove:
                    AddKnownFields(fields, remove.Columns);
                    break;
                case ReorderColumnsTransform reorder:
                    AddKnownFields(fields, reorder.Columns);
                    break;
                case RenameColumnTransform rename:
                    AddKnownField(fields, rename.From);
                    AddKnownField(fields, rename.To);
                    break;
                case ChangeColumnTypeTransform change:
                    AddKnownField(fields, change.Column);
                    break;
                case TrimTextTransform trim:
                    AddKnownFields(fields, trim.Columns);
                    break;
                case ReplaceValueTransform replace:
                    AddKnownField(fields, replace.Column);
                    break;
                case NormalizeBlanksTransform blanks:
                    AddKnownFields(fields, blanks.Columns);
                    break;
                case NormalizeErrorsTransform errors:
                    AddKnownFields(fields, errors.Columns);
                    break;
                case FillDownTransform fill:
                    AddKnownFields(fields, fill.Columns);
                    break;
                case MapValuesTransform map:
                    AddKnownField(fields, map.Column);
                    break;
                case FilterRowsTransform filter:
                    AddKnownField(fields, filter.Column);
                    break;
                case ExcludeTotalRowsTransform exclude:
                    foreach (TotalRowEvidenceSpec evidence in exclude.Evidence)
                    {
                        AddKnownField(fields, evidence.Column);
                    }

                    break;
                case DerivePeriodPartsTransform derive:
                    AddKnownField(fields, derive.DateColumn);
                    foreach (DerivedPeriodColumnSpec column in derive.Columns)
                    {
                        AddKnownField(fields, column.OutputColumn);
                    }

                    break;
                case AddArithmeticColumnTransform arithmetic:
                    AddKnownField(fields, arithmetic.Left.Column);
                    AddKnownField(fields, arithmetic.Right.Column);
                    AddKnownField(fields, arithmetic.OutputColumn);
                    break;
            }
        }

        private static void AddKnownFields(
            Dictionary<string, string> target,
            IEnumerable<string> fields)
        {
            foreach (string field in fields)
            {
                AddKnownField(target, field);
            }
        }

        private static void AddKnownField(
            Dictionary<string, string> target,
            string? field)
        {
            if (string.IsNullOrWhiteSpace(field))
            {
                return;
            }

            ValidateFieldFormat(field!, "Source field");
            if (!target.ContainsKey(field!))
            {
                target.Add(field!, field!);
            }
        }

        private static string ResolveSourceField(
            string value,
            string role,
            Dictionary<string, string> knownFields)
        {
            string candidate = (value ?? string.Empty).Trim();
            ValidateFieldFormat(candidate, role);
            if (!knownFields.TryGetValue(candidate, out string? canonical))
            {
                throw new InvalidOperationException(
                    role + " references unknown source field '" + candidate + "'.");
            }

            return canonical;
        }

        private static void ValidateFieldFormat(string value, string role)
        {
            if (string.IsNullOrWhiteSpace(value) || value.Length > 255 || value.Any(char.IsControl))
            {
                throw new InvalidOperationException(
                    role + " must contain 1 to 255 non-control characters.");
            }

            if (value.StartsWith("__erb_", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(role + " uses a reserved field prefix.");
            }
        }

        private static string ValidateLabel(string value)
        {
            string label = (value ?? string.Empty).Trim();
            if (label.Length == 0 || label.Length > 120 || label.Any(char.IsControl))
            {
                throw new InvalidOperationException(
                    "A calculated metric label must contain 1 to 120 non-control characters.");
            }

            return label;
        }

        private static string ValidateNumberFormat(string value)
        {
            string format = string.IsNullOrWhiteSpace(value) ? "General" : value.Trim();
            if (format.Length > 128 || format.Any(char.IsControl) ||
                format.StartsWith("=", StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "A calculated metric number format must be a bounded Excel number-format string.");
            }

            return format;
        }

        private static void DemandNoDetails(
            ManualCalculatedMetricSnapshot snapshot,
            string kind)
        {
            if (!string.IsNullOrWhiteSpace(snapshot.Details))
            {
                throw new InvalidOperationException(
                    "Details are not supported for calculated metric kind '" + kind + "'.");
            }
        }

        private static AggregateFunction ParseAggregate(string value)
        {
            switch (NormalizeToken(value))
            {
                case "sum": return AggregateFunction.Sum;
                case "count": return AggregateFunction.Count;
                case "distinct count": return AggregateFunction.DistinctCount;
                case "average": return AggregateFunction.Average;
                case "minimum":
                case "min": return AggregateFunction.Minimum;
                case "maximum":
                case "max": return AggregateFunction.Maximum;
                default:
                    throw new InvalidOperationException(
                        "Filtered aggregate Secondary must be Sum, Count, Distinct count, Average, Minimum, or Maximum.");
            }
        }

        private static MeasureValueType ResolveAggregateType(
            string field,
            AggregateFunction function,
            Dictionary<string, MeasureValueType> fieldTypes)
        {
            if (function == AggregateFunction.Count || function == AggregateFunction.DistinctCount)
            {
                return MeasureValueType.WholeNumber;
            }

            MeasureValueType type = fieldTypes.TryGetValue(field, out MeasureValueType known)
                ? known
                : MeasureValueType.Number;
            if (function == AggregateFunction.Average && type == MeasureValueType.WholeNumber)
            {
                return MeasureValueType.Number;
            }

            return type;
        }

        private static MeasureValueType ResolveWeightedResultType(
            string valueField,
            Dictionary<string, MeasureValueType> fieldTypes)
        {
            MeasureValueType type = fieldTypes.TryGetValue(valueField, out MeasureValueType known)
                ? known
                : MeasureValueType.Number;
            return type == MeasureValueType.WholeNumber ? MeasureValueType.Number : type;
        }

        private static MeasureValueType InferBinaryType(
            BinaryMeasureOperator operation,
            MeasureValueType left,
            MeasureValueType right)
        {
            if (operation == BinaryMeasureOperator.Add || operation == BinaryMeasureOperator.Subtract)
            {
                if (left == right) return left;
                if (IsPlainNumber(left) && IsPlainNumber(right)) return MeasureValueType.Number;
                throw new InvalidOperationException(
                    "Addition and subtraction require compatible measure types.");
            }

            if (left == MeasureValueType.Percentage)
            {
                return right == MeasureValueType.WholeNumber ? MeasureValueType.Number : right;
            }

            if (right == MeasureValueType.Percentage)
            {
                return left == MeasureValueType.WholeNumber ? MeasureValueType.Number : left;
            }

            if (left == MeasureValueType.Currency && IsPlainNumber(right) ||
                right == MeasureValueType.Currency && IsPlainNumber(left))
            {
                return MeasureValueType.Currency;
            }

            if (IsPlainNumber(left) && IsPlainNumber(right))
            {
                return MeasureValueType.Number;
            }

            throw new InvalidOperationException("Multiplication requires compatible measure types.");
        }

        private static MeasureValueType InferDivisionType(
            MeasureValueType numerator,
            MeasureValueType denominator)
        {
            if (numerator == MeasureValueType.Currency && IsPlainNumber(denominator))
            {
                return MeasureValueType.Currency;
            }

            if (numerator == denominator || IsPlainNumber(numerator) && IsPlainNumber(denominator))
            {
                return MeasureValueType.Number;
            }

            if (numerator == MeasureValueType.Percentage && IsPlainNumber(denominator))
            {
                return MeasureValueType.Percentage;
            }

            throw new InvalidOperationException("Division requires compatible measure types.");
        }

        private static MeasureValueType InferDifferenceType(
            DifferenceKind kind,
            MeasureValueType current,
            MeasureValueType baseline)
        {
            if (kind == DifferenceKind.PercentagePoints)
            {
                if (current != MeasureValueType.Percentage || baseline != MeasureValueType.Percentage)
                {
                    throw new InvalidOperationException(
                        "Percentage-point difference requires two percentage measures.");
                }

                return MeasureValueType.Percentage;
            }

            if (kind == DifferenceKind.Percentage)
            {
                DemandComparable(current, baseline, "Percentage change");
                return MeasureValueType.Percentage;
            }

            if (current == baseline) return current;
            if (IsPlainNumber(current) && IsPlainNumber(baseline)) return MeasureValueType.Number;
            throw new InvalidOperationException("Difference requires compatible measure types.");
        }

        private static void DemandComparable(
            MeasureValueType left,
            MeasureValueType right,
            string operation)
        {
            if (left != right && !(IsPlainNumber(left) && IsPlainNumber(right)))
            {
                throw new InvalidOperationException(
                    operation + " requires comparable measure types.");
            }
        }

        private static bool IsPlainNumber(MeasureValueType value)
        {
            return value == MeasureValueType.WholeNumber || value == MeasureValueType.Number;
        }

        private static ReferenceMeasureExpression Reference(MeasureDefinition measure)
        {
            return new ReferenceMeasureExpression
            {
                MeasureId = measure.Id,
                ResultType = measure.ValueType
            };
        }

        private static string NormalizeToken(string value)
        {
            string normalized = (value ?? string.Empty)
                .Trim()
                .ToLowerInvariant()
                .Replace('-', ' ')
                .Replace('_', ' ');
            return string.Join(
                " ",
                normalized.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries));
        }

        private static string CreateStableIdentifier(string prefix, string seed)
        {
            string slug = new string((seed ?? string.Empty)
                .Select(character => char.IsLetterOrDigit(character) ? character : '_')
                .ToArray())
                .Trim('_');
            if (slug.Length == 0)
            {
                slug = "item";
            }

            const int hashLength = 12;
            int maximumSlugLength = 64 - prefix.Length - hashLength - 2;
            if (maximumSlugLength < 1)
            {
                throw new InvalidOperationException("The internal identifier prefix is too long.");
            }

            if (slug.Length > maximumSlugLength)
            {
                slug = slug.Substring(0, maximumSlugLength);
            }

            string hash;
            using (SHA256 algorithm = SHA256.Create())
            {
                byte[] bytes = algorithm.ComputeHash(Encoding.UTF8.GetBytes(prefix + "\n" + seed));
                var builder = new StringBuilder(hashLength);
                for (var index = 0; index < hashLength / 2; index++)
                {
                    builder.Append(bytes[index].ToString("x2"));
                }

                hash = builder.ToString();
            }

            return prefix + "_" + slug + "_" + hash;
        }

        private sealed class ExistingMeasureIndex
        {
            private readonly Dictionary<string, MeasureDefinition> _byId;
            private readonly Dictionary<string, List<MeasureDefinition>> _byLabel;

            public ExistingMeasureIndex(
                Dictionary<string, MeasureDefinition> byId,
                Dictionary<string, List<MeasureDefinition>> byLabel)
            {
                _byId = byId;
                _byLabel = byLabel;
            }

            public void Add(MeasureDefinition measure)
            {
                if (_byId.ContainsKey(measure.Id))
                {
                    throw new InvalidOperationException(
                        "Calculated metric IDs must be unique within the current batch.");
                }

                _byId.Add(measure.Id, measure);
                if (!_byLabel.TryGetValue(measure.Label, out List<MeasureDefinition>? matches))
                {
                    matches = new List<MeasureDefinition>();
                    _byLabel.Add(measure.Label, matches);
                }

                matches.Add(measure);
            }

            public MeasureDefinition Resolve(string value, string role)
            {
                string reference = (value ?? string.Empty).Trim();
                if (reference.Length == 0 || reference.Length > 120 || reference.Any(char.IsControl))
                {
                    throw new InvalidOperationException(
                        role + " must identify an existing measure by label or ID.");
                }

                if (_byId.TryGetValue(reference, out MeasureDefinition? byId))
                {
                    return byId;
                }

                if (!_byLabel.TryGetValue(reference, out List<MeasureDefinition>? byLabel) ||
                    byLabel.Count == 0)
                {
                    throw new InvalidOperationException(
                        role + " references unknown measure '" + reference + "'.");
                }

                if (byLabel.Count != 1)
                {
                    throw new InvalidOperationException(
                        role + " label '" + reference + "' is ambiguous. Use the measure ID.");
                }

                return byLabel[0];
            }
        }

        private sealed class MetricTranslation
        {
            public MeasureValueType ValueType { get; set; }

            public MeasureExpression Expression { get; set; } = new ConstantMeasureExpression();

            public List<TransformStep> Transforms { get; } = new List<TransformStep>();

            public static MetricTranslation ForExpression(
                MeasureValueType valueType,
                MeasureExpression expression)
            {
                return new MetricTranslation
                {
                    ValueType = valueType,
                    Expression = expression
                };
            }
        }
    }
}
