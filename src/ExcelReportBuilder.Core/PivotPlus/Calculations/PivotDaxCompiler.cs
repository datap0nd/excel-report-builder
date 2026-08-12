using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using ExcelReportBuilder.Core.Validation;

namespace ExcelReportBuilder.Core.PivotPlus.Calculations
{
    /// <summary>
    /// Deterministically compiles the closed PivotTable+ calculation union to
    /// Data Model DAX. The input contract contains no raw formula node.
    /// </summary>
    public static class PivotDaxCompiler
    {
        private const int MaximumCompiledFormulaLength = 32768;

        public static PivotDaxCompilation Compile(PivotMeasureSetDefinition definition)
        {
            if (definition == null) throw new ArgumentNullException(nameof(definition));

            ValidationResult validation = PivotCalculationValidator.Validate(definition);
            if (!validation.IsValid)
            {
                throw new InvalidPivotCalculationException(validation);
            }

            var model = new PivotCalculationModelIndex(definition.Schema);
            var measuresById = definition.Measures.ToDictionary(
                measure => measure.Id,
                StringComparer.OrdinalIgnoreCase);
            var displayIndexes = definition.Measures
                .Select((measure, index) => new { measure.Id, Index = index })
                .ToDictionary(item => item.Id, item => item.Index, StringComparer.OrdinalIgnoreCase);
            var slices = (definition.Periods?.Slices ?? Array.Empty<PivotPeriodSlice>())
                .ToDictionary(slice => slice.Id, StringComparer.OrdinalIgnoreCase);
            var context = new CompilerContext(
                model,
                measuresById,
                displayIndexes,
                definition.Periods,
                slices);

            Dictionary<string, IReadOnlyList<string>> dependencies = definition.Measures
                .ToDictionary(
                    measure => measure.Id,
                    measure => ReadDependencies(measure.Expression)
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .OrderBy(id => displayIndexes[id])
                        .ToArray() as IReadOnlyList<string>,
                    StringComparer.OrdinalIgnoreCase);
            Dictionary<string, int> creationOrders = BuildCreationOrders(
                definition.Measures,
                dependencies,
                displayIndexes);

            var compiled = new List<OwnedPivotMeasureDefinition>(definition.Measures.Count);
            for (var index = 0; index < definition.Measures.Count; index++)
            {
                PivotMeasureDefinition measure = definition.Measures[index];
                string dax = CompileExpression(measure.Expression, context);
                if (dax.Length > MaximumCompiledFormulaLength)
                {
                    var formulaValidation = new ValidationResult();
                    formulaValidation.AddError(
                        "PIVOT_CALC_FORMULA_TOO_LONG",
                        "$.measures[" + index.ToString(CultureInfo.InvariantCulture) + "].expression",
                        "The compiled DAX formula exceeds the bounded host-safe length.");
                    throw new InvalidPivotCalculationException(formulaValidation);
                }

                string definitionCanonical = CanonicalDefinition(
                    measure,
                    index + 1,
                    context);
                PivotModelTableSchema homeTable = model.TryGetTable(
                    measure.HomeTableId,
                    out PivotModelTableSchema table)
                    ? table
                    : throw new InvalidOperationException("Validated home table binding was lost.");

                compiled.Add(new OwnedPivotMeasureDefinition(
                    measure.Id,
                    index + 1,
                    creationOrders[measure.Id],
                    homeTable.Name,
                    measure.Caption,
                    dax,
                    measure.Format,
                    dependencies[measure.Id],
                    PivotDaxFingerprint.ComputeDefinition(definitionCanonical),
                    PivotDaxFingerprint.ComputeFormula(dax)));
            }

            return new PivotDaxCompilation(compiled);
        }

        private static Dictionary<string, int> BuildCreationOrders(
            IReadOnlyList<PivotMeasureDefinition> measures,
            IReadOnlyDictionary<string, IReadOnlyList<string>> dependencies,
            IReadOnlyDictionary<string, int> displayIndexes)
        {
            var ordered = new List<string>(measures.Count);
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var active = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (PivotMeasureDefinition measure in measures)
            {
                Visit(measure.Id);
            }

            var result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (var index = 0; index < ordered.Count; index++)
            {
                result.Add(ordered[index], index + 1);
            }

            return result;

            void Visit(string id)
            {
                if (visited.Contains(id))
                {
                    return;
                }

                if (!active.Add(id))
                {
                    throw new InvalidOperationException("Validated calculation graph unexpectedly contains a cycle.");
                }

                foreach (string dependency in dependencies[id].OrderBy(item => displayIndexes[item]))
                {
                    Visit(dependency);
                }

                active.Remove(id);
                visited.Add(id);
                ordered.Add(id);
            }
        }

        private static IEnumerable<string> ReadDependencies(PivotCalculationExpression expression)
        {
            switch (expression)
            {
                case PivotMeasureReferenceExpression reference:
                    yield return reference.MeasureId;
                    yield break;
                case PivotDifferenceExpression difference:
                    foreach (string value in ReadDependencies(difference.Left)) yield return value;
                    foreach (string value in ReadDependencies(difference.Right)) yield return value;
                    yield break;
                case PivotSafeRatioExpression ratio:
                    foreach (string value in ReadDependencies(ratio.Numerator)) yield return value;
                    foreach (string value in ReadDependencies(ratio.Denominator)) yield return value;
                    yield break;
                case PivotShareExpression share:
                    foreach (string value in ReadDependencies(share.Part)) yield return value;
                    if (share.Denominator is PivotExplicitShareDenominator explicitDenominator)
                    {
                        foreach (string value in ReadDependencies(explicitDenominator.Expression))
                        {
                            yield return value;
                        }
                    }

                    yield break;
                case PivotGrowthExpression growth:
                    foreach (string value in ReadDependencies(growth.Current)) yield return value;
                    foreach (string value in ReadDependencies(growth.Prior)) yield return value;
                    yield break;
                case PivotAchievementExpression achievement:
                    foreach (string value in ReadDependencies(achievement.Actual)) yield return value;
                    foreach (string value in ReadDependencies(achievement.Target)) yield return value;
                    yield break;
                case PivotVarianceExpression variance:
                    foreach (string value in ReadDependencies(variance.Actual)) yield return value;
                    foreach (string value in ReadDependencies(variance.Plan)) yield return value;
                    yield break;
                case PivotVariancePercentageExpression variancePercentage:
                    foreach (string value in ReadDependencies(variancePercentage.Actual)) yield return value;
                    foreach (string value in ReadDependencies(variancePercentage.Plan)) yield return value;
                    yield break;
                case PivotPercentagePointDeltaExpression delta:
                    foreach (string value in ReadDependencies(delta.CurrentRatio)) yield return value;
                    foreach (string value in ReadDependencies(delta.BaselineRatio)) yield return value;
                    yield break;
                default:
                    yield break;
            }
        }

        private static string CompileExpression(
            PivotCalculationExpression expression,
            CompilerContext context)
        {
            switch (expression)
            {
                case PivotAggregateExpression aggregate:
                    return CompileAggregate(
                        aggregate.FieldId,
                        aggregate.Function,
                        Array.Empty<PivotCalculationFilter>(),
                        aggregate.PeriodSliceId,
                        context);
                case PivotFilteredAggregateExpression aggregate:
                    return CompileAggregate(
                        aggregate.FieldId,
                        aggregate.Function,
                        aggregate.Filters,
                        aggregate.PeriodSliceId,
                        context);
                case PivotWeightedResultExpression weighted:
                    return CompileWeighted(weighted, context);
                case PivotMeasureReferenceExpression reference:
                    return MeasureReference(context.Measures[reference.MeasureId].Caption);
                case PivotDifferenceExpression difference:
                    return "(" + CompileExpression(difference.Left, context) + " - " +
                           CompileExpression(difference.Right, context) + ")";
                case PivotSafeRatioExpression ratio:
                    return Divide(
                        CompileExpression(ratio.Numerator, context),
                        CompileExpression(ratio.Denominator, context),
                        ratio.OnZero);
                case PivotShareExpression share:
                    return CompileShare(share, context);
                case PivotGrowthExpression growth:
                {
                    string current = CompileExpression(growth.Current, context);
                    string prior = CompileExpression(growth.Prior, context);
                    return Divide("(" + current + " - " + prior + ")", prior, growth.OnZero);
                }
                case PivotAchievementExpression achievement:
                    return Divide(
                        CompileExpression(achievement.Actual, context),
                        CompileExpression(achievement.Target, context),
                        achievement.OnZero);
                case PivotVarianceExpression variance:
                    return CompileVariance(
                        CompileExpression(variance.Actual, context),
                        CompileExpression(variance.Plan, context),
                        variance.Convention);
                case PivotVariancePercentageExpression variancePercentage:
                {
                    string actual = CompileExpression(variancePercentage.Actual, context);
                    string plan = CompileExpression(variancePercentage.Plan, context);
                    string variance = CompileVariance(actual, plan, variancePercentage.Convention);
                    return Divide(variance, plan, variancePercentage.OnZero);
                }
                case PivotPercentagePointDeltaExpression delta:
                {
                    string current = CompileExpression(delta.CurrentRatio, context);
                    string baseline = CompileExpression(delta.BaselineRatio, context);
                    return "IF(OR(ISBLANK(" + current + "), ISBLANK(" + baseline +
                           ")), BLANK(), 100 * (" + current + " - " + baseline + "))";
                }
                default:
                    throw new InvalidOperationException("Validated expression kind was not recognized.");
            }
        }

        private static string CompileAggregate(
            string fieldId,
            PivotCalculationAggregateFunction function,
            IReadOnlyList<PivotCalculationFilter> filters,
            string? periodSliceId,
            CompilerContext context)
        {
            PivotBoundField field = BoundField(fieldId, context);
            string column = ColumnReference(field);
            string aggregate;
            switch (function)
            {
                case PivotCalculationAggregateFunction.Sum:
                    aggregate = "SUM(" + column + ")";
                    break;
                case PivotCalculationAggregateFunction.Count:
                    aggregate = "COUNTA(" + column + ")";
                    break;
                case PivotCalculationAggregateFunction.DistinctCount:
                    aggregate = "DISTINCTCOUNT(" + column + ")";
                    break;
                case PivotCalculationAggregateFunction.Average:
                    aggregate = "AVERAGE(" + column + ")";
                    break;
                case PivotCalculationAggregateFunction.Minimum:
                    aggregate = "MIN(" + column + ")";
                    break;
                case PivotCalculationAggregateFunction.Maximum:
                    aggregate = "MAX(" + column + ")";
                    break;
                default:
                    throw new InvalidOperationException("Validated aggregate function was not recognized.");
            }

            List<string> arguments = CompileContextArguments(filters, periodSliceId, context);
            return arguments.Count == 0
                ? aggregate
                : "CALCULATE(" + aggregate + ", " + string.Join(", ", arguments) + ")";
        }

        private static string CompileWeighted(
            PivotWeightedResultExpression weighted,
            CompilerContext context)
        {
            PivotBoundField value = BoundField(weighted.ValueFieldId, context);
            PivotBoundField weight = BoundField(weighted.WeightFieldId, context);
            string table = TableReference(value.Table.Name);
            string valueColumn = ColumnReference(value);
            string weightColumn = ColumnReference(weight);
            string rows = "FILTER(" + table + ", NOT(ISBLANK(" + valueColumn +
                          ")) && NOT(ISBLANK(" + weightColumn + ")))";
            List<string> arguments = CompileContextArguments(
                weighted.Filters,
                weighted.PeriodSliceId,
                context);
            if (arguments.Count > 0)
            {
                rows = "CALCULATETABLE(" + rows + ", " + string.Join(", ", arguments) + ")";
            }

            return "VAR __PivotPlusRows = " + rows +
                   " RETURN DIVIDE(SUMX(__PivotPlusRows, " + valueColumn + " * " + weightColumn +
                   "), SUMX(__PivotPlusRows, " + weightColumn + "), " +
                   Alternate(weighted.OnZero) + ")";
        }

        private static string CompileShare(
            PivotShareExpression share,
            CompilerContext context)
        {
            string part = CompileExpression(share.Part, context);
            string denominator;
            switch (share.Denominator)
            {
                case PivotExplicitShareDenominator explicitDenominator:
                    denominator = CompileExpression(explicitDenominator.Expression, context);
                    break;
                case PivotParentShareDenominator parent:
                    denominator = "CALCULATE(" + part + ", " + string.Join(", ",
                        SortBoundFields(parent.ClearedFieldIds, context)
                            .Select(field => "REMOVEFILTERS(" + ColumnReference(field) + ")")) + ")";
                    break;
                case PivotFilteredTotalShareDenominator filteredTotal:
                    denominator = "CALCULATE(" + part + ", " + string.Join(", ",
                        GroupAllSelected(filteredTotal.ClearedFieldIds, context)) + ")";
                    break;
                default:
                    throw new InvalidOperationException("Validated share denominator was not recognized.");
            }

            return Divide(part, denominator, share.OnZero);
        }

        private static IEnumerable<string> GroupAllSelected(
            IReadOnlyList<string> fieldIds,
            CompilerContext context)
        {
            return SortBoundFields(fieldIds, context)
                .GroupBy(field => field.Table.Id, StringComparer.OrdinalIgnoreCase)
                .OrderBy(group => group.First().Table.Name, StringComparer.Ordinal)
                .Select(group => "ALLSELECTED(" + string.Join(", ",
                    group.Select(ColumnReference)) + ")");
        }

        private static IEnumerable<PivotBoundField> SortBoundFields(
            IReadOnlyList<string> fieldIds,
            CompilerContext context)
        {
            return fieldIds
                .Select(id => BoundField(id, context))
                .OrderBy(field => field.Table.Name, StringComparer.Ordinal)
                .ThenBy(field => field.Field.Name, StringComparer.Ordinal);
        }

        private static List<string> CompileContextArguments(
            IReadOnlyList<PivotCalculationFilter> filters,
            string? periodSliceId,
            CompilerContext context)
        {
            var arguments = filters
                .OrderBy(filter => filter.FieldId, StringComparer.OrdinalIgnoreCase)
                .Select(filter => "KEEPFILTERS(" + CompileFilterPredicate(filter, context) + ")")
                .ToList();

            if (periodSliceId == null)
            {
                return arguments;
            }

            if (context.Periods == null)
            {
                throw new InvalidOperationException("Validated period binding was lost.");
            }

            PivotPeriodSlice slice = context.Slices[periodSliceId];
            PivotPeriodSource source = context.Periods.Source;
            if (slice.FilterMode == PivotSliceFilterMode.ReplaceAxisContext)
            {
                IEnumerable<PivotBoundField> cleared = SortBoundFields(
                    source.PeriodContextFieldIds
                        .Concat(source.ScenarioContextFieldIds)
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToArray(),
                    context);
                arguments.AddRange(cleared.Select(field =>
                    "REMOVEFILTERS(" + ColumnReference(field) + ")"));
                arguments.Add(CompilePeriodPredicate(slice, context));
                string? scenario = CompileScenarioPredicate(slice, context);
                if (scenario != null)
                {
                    arguments.Add(scenario);
                }
            }
            else
            {
                arguments.Add("KEEPFILTERS(" + CompilePeriodPredicate(slice, context) + ")");
                string? scenario = CompileScenarioPredicate(slice, context);
                if (scenario != null)
                {
                    arguments.Add("KEEPFILTERS(" + scenario + ")");
                }
            }

            return arguments;
        }

        private static string CompilePeriodPredicate(
            PivotPeriodSlice slice,
            CompilerContext context)
        {
            if (context.Periods == null)
            {
                throw new InvalidOperationException("Validated period binding was lost.");
            }

            PivotPeriodSource source = context.Periods.Source;
            PivotBoundField field = BoundField(source.PeriodFieldId, context);
            string column = ColumnReference(field);
            if (source.DateCoverageMode == PivotDateCoverageMode.ContinuousRange)
            {
                if (!PivotPeriodRules.TryGetDateRange(slice.Point, out DateTime start, out DateTime end))
                {
                    throw new InvalidOperationException("Validated continuous date slice could not be resolved.");
                }

                string startLiteral = DateLiteral(start);
                if (field.Field.DataType == PivotModelDataType.DateTime)
                {
                    string exclusiveEnd = "(" + DateLiteral(end) + " + 1)";
                    return column + " >= " + startLiteral + " && " + column + " < " + exclusiveEnd;
                }

                string endLiteral = DateLiteral(end);
                return start == end
                    ? column + " = " + startLiteral
                    : column + " >= " + startLiteral + " && " + column + " <= " + endLiteral;
            }

            List<PivotScalarValue> values = PivotPeriodRules.ResolveCoverage(context.Periods, slice)
                .Select(member => ResolveValue(source.PeriodFieldId, member.SourceValue, context))
                .OrderBy(PivotCalculationCanonical.ScalarKey, StringComparer.Ordinal)
                .ToList();
            return CompileSetPredicate(column, values);
        }

        private static string? CompileScenarioPredicate(
            PivotPeriodSlice slice,
            CompilerContext context)
        {
            if (context.Periods?.Source.ScenarioFieldId == null)
            {
                return null;
            }

            string fieldId = context.Periods.Source.ScenarioFieldId;
            var member = PivotFilterValue.FromMember(slice.ScenarioMemberId!);
            PivotScalarValue value = ResolveValue(fieldId, member, context);
            return CompileSetPredicate(
                ColumnReference(BoundField(fieldId, context)),
                new[] { value });
        }

        private static string CompileFilterPredicate(
            PivotCalculationFilter filter,
            CompilerContext context)
        {
            string column = ColumnReference(BoundField(filter.FieldId, context));
            List<PivotScalarValue> values = filter.Values
                .Select(value => ResolveValue(filter.FieldId, value, context))
                .OrderBy(PivotCalculationCanonical.ScalarKey, StringComparer.Ordinal)
                .ToList();
            switch (filter.Operator)
            {
                case PivotCalculationFilterOperator.Equal:
                    return CompileSetPredicate(column, values);
                case PivotCalculationFilterOperator.NotEqual:
                    return "NOT (" + CompileSetPredicate(column, values) + ")";
                case PivotCalculationFilterOperator.GreaterThan:
                    return NonBlankComparison(column, ">", values[0]);
                case PivotCalculationFilterOperator.GreaterThanOrEqual:
                    return NonBlankComparison(column, ">=", values[0]);
                case PivotCalculationFilterOperator.LessThan:
                    return NonBlankComparison(column, "<", values[0]);
                case PivotCalculationFilterOperator.LessThanOrEqual:
                    return NonBlankComparison(column, "<=", values[0]);
                case PivotCalculationFilterOperator.In:
                    return CompileSetPredicate(column, values);
                case PivotCalculationFilterOperator.NotIn:
                    return "NOT (" + CompileSetPredicate(column, values) + ")";
                case PivotCalculationFilterOperator.IsBlank:
                    return "ISBLANK(" + column + ")";
                case PivotCalculationFilterOperator.IsNotBlank:
                    return "NOT(ISBLANK(" + column + "))";
                default:
                    throw new InvalidOperationException("Validated filter operator was not recognized.");
            }
        }

        private static string CompileSetPredicate(
            string column,
            IReadOnlyList<PivotScalarValue> values)
        {
            return column + " IN { " + string.Join(", ", values.Select(ScalarLiteral)) + " }";
        }

        private static string NonBlankComparison(
            string column,
            string comparisonOperator,
            PivotScalarValue value)
        {
            return "NOT(ISBLANK(" + column + ")) && " + column + " " +
                   comparisonOperator + " " + ScalarLiteral(value);
        }

        private static PivotScalarValue ResolveValue(
            string fieldId,
            PivotFilterValue value,
            CompilerContext context)
        {
            if (!context.Model.TryResolveValue(fieldId, value, out PivotScalarValue scalar))
            {
                throw new InvalidOperationException("Validated filter value binding was lost.");
            }

            return scalar;
        }

        private static string CompileVariance(
            string actual,
            string plan,
            PivotVarianceConvention convention)
        {
            return convention == PivotVarianceConvention.ActualMinusPlan
                ? "(" + actual + " - " + plan + ")"
                : "(" + plan + " - " + actual + ")";
        }

        private static string Divide(
            string numerator,
            string denominator,
            PivotDenominatorBehavior behavior)
        {
            return "DIVIDE(" + numerator + ", " + denominator + ", " + Alternate(behavior) + ")";
        }

        private static string Alternate(PivotDenominatorBehavior behavior)
        {
            return behavior == PivotDenominatorBehavior.Blank ? "BLANK()" : "0";
        }

        private static string TableReference(string tableName)
        {
            return "'" + tableName.Replace("'", "''") + "'";
        }

        private static string ColumnReference(PivotBoundField field)
        {
            return TableReference(field.Table.Name) + "[" +
                   field.Field.Name.Replace("]", "]]" ) + "]";
        }

        private static string MeasureReference(string measureName)
        {
            return "[" + measureName.Replace("]", "]]" ) + "]";
        }

        private static string ScalarLiteral(PivotScalarValue value)
        {
            switch (value.Kind)
            {
                case PivotScalarKind.Blank:
                    return "BLANK()";
                case PivotScalarKind.Text:
                    return "\"" + (value.TextValue ?? string.Empty).Replace("\"", "\"\"") + "\"";
                case PivotScalarKind.WholeNumber:
                    return value.WholeNumberValue.GetValueOrDefault()
                        .ToString(CultureInfo.InvariantCulture);
                case PivotScalarKind.DecimalNumber:
                    return value.DecimalNumberValue.GetValueOrDefault()
                        .ToString("0.############################", CultureInfo.InvariantCulture);
                case PivotScalarKind.Boolean:
                    return value.BooleanValue.GetValueOrDefault() ? "TRUE()" : "FALSE()";
                case PivotScalarKind.Date:
                    return DateLiteral(value.TemporalValue.GetValueOrDefault());
                case PivotScalarKind.DateTime:
                {
                    DateTime temporal = value.TemporalValue.GetValueOrDefault();
                    decimal seconds = temporal.Second + (temporal.Ticks % TimeSpan.TicksPerSecond) /
                        (decimal)TimeSpan.TicksPerSecond;
                    return DateLiteral(temporal) + " + TIME(" +
                           temporal.Hour.ToString(CultureInfo.InvariantCulture) + ", " +
                           temporal.Minute.ToString(CultureInfo.InvariantCulture) + ", " +
                           seconds.ToString("0.#######", CultureInfo.InvariantCulture) + ")";
                }
                default:
                    throw new InvalidOperationException("Validated scalar kind was not recognized.");
            }
        }

        private static string DateLiteral(DateTime value)
        {
            return "DATE(" + value.Year.ToString(CultureInfo.InvariantCulture) + ", " +
                   value.Month.ToString(CultureInfo.InvariantCulture) + ", " +
                   value.Day.ToString(CultureInfo.InvariantCulture) + ")";
        }

        private static PivotBoundField BoundField(string fieldId, CompilerContext context)
        {
            return context.Model.TryGetField(fieldId, out PivotBoundField field)
                ? field
                : throw new InvalidOperationException("Validated model field binding was lost.");
        }

        private static string CanonicalDefinition(
            PivotMeasureDefinition measure,
            int displayOrder,
            CompilerContext context)
        {
            PivotModelTableSchema table = context.Model.TryGetTable(
                measure.HomeTableId,
                out PivotModelTableSchema boundTable)
                ? boundTable
                : throw new InvalidOperationException("Validated home table binding was lost.");
            var writer = new PivotFingerprintCanonicalWriter();
            writer.Add("version", "1");
            writer.Add("id", measure.Id);
            writer.Add("displayOrder", displayOrder.ToString(CultureInfo.InvariantCulture));
            writer.Add("homeTableId", table.Id);
            writer.Add("homeTableName", table.Name);
            writer.Add("measureName", measure.Caption);
            PivotDaxFingerprint.AddFormat(writer, measure.Format);
            writer.Add("expression", CanonicalExpression(measure.Expression, context));
            return writer.ToString();
        }

        private static string CanonicalExpression(
            PivotCalculationExpression expression,
            CompilerContext context)
        {
            var writer = new PivotFingerprintCanonicalWriter();
            writer.Add("kind", ((int)expression.Kind).ToString(CultureInfo.InvariantCulture));
            switch (expression)
            {
                case PivotAggregateExpression aggregate:
                    AddField(writer, "field", BoundField(aggregate.FieldId, context));
                    writer.Add("function", ((int)aggregate.Function).ToString(CultureInfo.InvariantCulture));
                    AddPeriod(writer, aggregate.PeriodSliceId, context);
                    break;
                case PivotFilteredAggregateExpression aggregate:
                    AddField(writer, "field", BoundField(aggregate.FieldId, context));
                    writer.Add("function", ((int)aggregate.Function).ToString(CultureInfo.InvariantCulture));
                    AddFilters(writer, aggregate.Filters, context);
                    AddPeriod(writer, aggregate.PeriodSliceId, context);
                    break;
                case PivotWeightedResultExpression weighted:
                    AddField(writer, "valueField", BoundField(weighted.ValueFieldId, context));
                    AddField(writer, "weightField", BoundField(weighted.WeightFieldId, context));
                    writer.Add("onZero", ((int)weighted.OnZero).ToString(CultureInfo.InvariantCulture));
                    AddFilters(writer, weighted.Filters, context);
                    AddPeriod(writer, weighted.PeriodSliceId, context);
                    break;
                case PivotMeasureReferenceExpression reference:
                    writer.Add("measureId", reference.MeasureId);
                    break;
                case PivotDifferenceExpression difference:
                    writer.Add("left", CanonicalExpression(difference.Left, context));
                    writer.Add("right", CanonicalExpression(difference.Right, context));
                    break;
                case PivotSafeRatioExpression ratio:
                    writer.Add("numerator", CanonicalExpression(ratio.Numerator, context));
                    writer.Add("denominator", CanonicalExpression(ratio.Denominator, context));
                    writer.Add("onZero", ((int)ratio.OnZero).ToString(CultureInfo.InvariantCulture));
                    break;
                case PivotShareExpression share:
                    writer.Add("part", CanonicalExpression(share.Part, context));
                    writer.Add("onZero", ((int)share.OnZero).ToString(CultureInfo.InvariantCulture));
                    AddShareDenominator(writer, share.Denominator, context);
                    break;
                case PivotGrowthExpression growth:
                    writer.Add("current", CanonicalExpression(growth.Current, context));
                    writer.Add("prior", CanonicalExpression(growth.Prior, context));
                    writer.Add("onZero", ((int)growth.OnZero).ToString(CultureInfo.InvariantCulture));
                    break;
                case PivotAchievementExpression achievement:
                    writer.Add("actual", CanonicalExpression(achievement.Actual, context));
                    writer.Add("target", CanonicalExpression(achievement.Target, context));
                    writer.Add("onZero", ((int)achievement.OnZero).ToString(CultureInfo.InvariantCulture));
                    break;
                case PivotVarianceExpression variance:
                    writer.Add("actual", CanonicalExpression(variance.Actual, context));
                    writer.Add("plan", CanonicalExpression(variance.Plan, context));
                    writer.Add("convention", ((int)variance.Convention).ToString(CultureInfo.InvariantCulture));
                    break;
                case PivotVariancePercentageExpression variancePercentage:
                    writer.Add("actual", CanonicalExpression(variancePercentage.Actual, context));
                    writer.Add("plan", CanonicalExpression(variancePercentage.Plan, context));
                    writer.Add("convention", ((int)variancePercentage.Convention).ToString(CultureInfo.InvariantCulture));
                    writer.Add("onZero", ((int)variancePercentage.OnZero).ToString(CultureInfo.InvariantCulture));
                    break;
                case PivotPercentagePointDeltaExpression delta:
                    writer.Add("currentRatio", CanonicalExpression(delta.CurrentRatio, context));
                    writer.Add("baselineRatio", CanonicalExpression(delta.BaselineRatio, context));
                    break;
            }

            return writer.ToString();
        }

        private static void AddField(
            PivotFingerprintCanonicalWriter writer,
            string prefix,
            PivotBoundField field)
        {
            writer.Add(prefix + ".id", field.Field.Id);
            writer.Add(prefix + ".tableId", field.Table.Id);
            writer.Add(prefix + ".tableName", field.Table.Name);
            writer.Add(prefix + ".name", field.Field.Name);
            writer.Add(prefix + ".type", ((int)field.Field.DataType).ToString(CultureInfo.InvariantCulture));
        }

        private static void AddFilters(
            PivotFingerprintCanonicalWriter writer,
            IReadOnlyList<PivotCalculationFilter> filters,
            CompilerContext context)
        {
            PivotCalculationFilter[] ordered = filters
                .OrderBy(filter => filter.FieldId, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            writer.Add("filterCount", ordered.Length.ToString(CultureInfo.InvariantCulture));
            for (var index = 0; index < ordered.Length; index++)
            {
                PivotCalculationFilter filter = ordered[index];
                var filterWriter = new PivotFingerprintCanonicalWriter();
                AddField(filterWriter, "field", BoundField(filter.FieldId, context));
                filterWriter.Add("operator", ((int)filter.Operator).ToString(CultureInfo.InvariantCulture));
                string[] values = filter.Values
                    .Select(value => PivotCalculationCanonical.ScalarKey(
                        ResolveValue(filter.FieldId, value, context)))
                    .OrderBy(value => value, StringComparer.Ordinal)
                    .ToArray();
                filterWriter.Add("valueCount", values.Length.ToString(CultureInfo.InvariantCulture));
                for (var valueIndex = 0; valueIndex < values.Length; valueIndex++)
                {
                    filterWriter.Add("value", values[valueIndex]);
                }

                writer.Add("filter", filterWriter.ToString());
            }
        }

        private static void AddPeriod(
            PivotFingerprintCanonicalWriter writer,
            string? periodSliceId,
            CompilerContext context)
        {
            if (periodSliceId == null)
            {
                writer.Add("period", string.Empty);
                return;
            }

            if (context.Periods == null)
            {
                throw new InvalidOperationException("Validated period binding was lost.");
            }

            PivotPeriodSlice slice = context.Slices[periodSliceId];
            PivotPeriodSource source = context.Periods.Source;
            var periodWriter = new PivotFingerprintCanonicalWriter();
            periodWriter.Add("sliceId", slice.Id);
            periodWriter.Add("caption", slice.Caption);
            periodWriter.Add("point", PivotCalculationCanonical.PeriodPointKey(slice.Point));
            periodWriter.Add("scenarioMemberId", slice.ScenarioMemberId ?? string.Empty);
            periodWriter.Add("filterMode", ((int)slice.FilterMode).ToString(CultureInfo.InvariantCulture));
            periodWriter.Add("sourceGrain", ((int)source.SourceGrain).ToString(CultureInfo.InvariantCulture));
            periodWriter.Add("dateCoverageMode", ((int)source.DateCoverageMode).ToString(CultureInfo.InvariantCulture));
            AddField(periodWriter, "periodField", BoundField(source.PeriodFieldId, context));
            foreach (PivotBoundField field in SortBoundFields(source.PeriodContextFieldIds, context))
            {
                AddField(periodWriter, "periodContext", field);
            }

            if (source.ScenarioFieldId != null)
            {
                AddField(periodWriter, "scenarioField", BoundField(source.ScenarioFieldId, context));
                periodWriter.Add(
                    "scenarioValue",
                    PivotCalculationCanonical.ScalarKey(ResolveValue(
                        source.ScenarioFieldId,
                        PivotFilterValue.FromMember(slice.ScenarioMemberId!),
                        context)));
                foreach (PivotBoundField field in SortBoundFields(source.ScenarioContextFieldIds, context))
                {
                    AddField(periodWriter, "scenarioContext", field);
                }
            }

            if (source.DateCoverageMode == PivotDateCoverageMode.ContinuousRange)
            {
                periodWriter.Add(
                    "rangeStart",
                    source.ContinuousRangeStart!.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
                periodWriter.Add(
                    "rangeEnd",
                    source.ContinuousRangeEnd!.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
            }
            else
            {
                string[] values = PivotPeriodRules.ResolveCoverage(context.Periods, slice)
                    .Select(member => PivotCalculationCanonical.ScalarKey(
                        ResolveValue(source.PeriodFieldId, member.SourceValue, context)))
                    .OrderBy(value => value, StringComparer.Ordinal)
                    .ToArray();
                foreach (string value in values)
                {
                    periodWriter.Add("sourceValue", value);
                }
            }

            writer.Add("period", periodWriter.ToString());
        }

        private static void AddShareDenominator(
            PivotFingerprintCanonicalWriter writer,
            PivotShareDenominator denominator,
            CompilerContext context)
        {
            writer.Add("denominatorKind", ((int)denominator.Kind).ToString(CultureInfo.InvariantCulture));
            switch (denominator)
            {
                case PivotExplicitShareDenominator explicitDenominator:
                    writer.Add("denominator", CanonicalExpression(explicitDenominator.Expression, context));
                    break;
                case PivotParentShareDenominator parent:
                    foreach (PivotBoundField field in SortBoundFields(parent.ClearedFieldIds, context))
                    {
                        AddField(writer, "clearedField", field);
                    }

                    break;
                case PivotFilteredTotalShareDenominator filteredTotal:
                    foreach (PivotBoundField field in SortBoundFields(filteredTotal.ClearedFieldIds, context))
                    {
                        AddField(writer, "clearedField", field);
                    }

                    break;
            }
        }

        private sealed class CompilerContext
        {
            public CompilerContext(
                PivotCalculationModelIndex model,
                IReadOnlyDictionary<string, PivotMeasureDefinition> measures,
                IReadOnlyDictionary<string, int> displayIndexes,
                PivotPeriodDefinition? periods,
                IReadOnlyDictionary<string, PivotPeriodSlice> slices)
            {
                Model = model;
                Measures = measures;
                DisplayIndexes = displayIndexes;
                Periods = periods;
                Slices = slices;
            }

            public PivotCalculationModelIndex Model { get; }

            public IReadOnlyDictionary<string, PivotMeasureDefinition> Measures { get; }

            public IReadOnlyDictionary<string, int> DisplayIndexes { get; }

            public PivotPeriodDefinition? Periods { get; }

            public IReadOnlyDictionary<string, PivotPeriodSlice> Slices { get; }
        }
    }
}
