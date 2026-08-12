using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using ExcelReportBuilder.Core.Validation;

namespace ExcelReportBuilder.Core.PivotPlus.Calculations
{
    public enum PivotModelDataType
    {
        Unknown,
        Text,
        WholeNumber,
        DecimalNumber,
        Currency,
        Date,
        DateTime,
        Boolean
    }

    public enum PivotScalarKind
    {
        Blank,
        Text,
        WholeNumber,
        DecimalNumber,
        Boolean,
        Date,
        DateTime
    }

    public enum PivotCalculationFilterOperator
    {
        Unknown,
        Equal,
        NotEqual,
        GreaterThan,
        GreaterThanOrEqual,
        LessThan,
        LessThanOrEqual,
        In,
        NotIn,
        IsBlank,
        IsNotBlank
    }

    public enum PivotFilterValueKind
    {
        Scalar,
        Member
    }

    public enum PivotCalculationAggregateFunction
    {
        Unknown,
        Sum,
        Count,
        DistinctCount,
        Average,
        Minimum,
        Maximum
    }

    public enum PivotPeriodGrain
    {
        Unknown,
        Year,
        Half,
        Quarter,
        Month,
        Date
    }

    public enum PivotPeriodCoverageStatus
    {
        Unknown,
        Partial,
        Complete
    }

    public enum PivotDateCoverageMode
    {
        NotApplicable,
        ExplicitCalendarMembers,
        ContinuousRange
    }

    public enum PivotSliceFilterMode
    {
        Unknown,
        ReplaceAxisContext,
        IntersectCurrentContext
    }

    public enum PivotMeasureFormatKind
    {
        Unknown,
        WholeNumber,
        DecimalNumber,
        Currency,
        Percentage,
        PercentagePoints
    }

    /// <summary>
    /// Bounded native formatting semantics. This is intentionally not an
    /// arbitrary DAX or Excel number-format string.
    /// </summary>
    public sealed class PivotMeasureFormat
    {
        public PivotMeasureFormat(
            PivotMeasureFormatKind kind,
            int decimalPlaces,
            bool useThousandsSeparator,
            string? currencySymbolOrCode = null)
        {
            Kind = kind;
            DecimalPlaces = decimalPlaces;
            UseThousandsSeparator = useThousandsSeparator;
            CurrencySymbolOrCode = currencySymbolOrCode;
        }

        public PivotMeasureFormatKind Kind { get; }

        public int DecimalPlaces { get; }

        public bool UseThousandsSeparator { get; }

        public string? CurrencySymbolOrCode { get; }
    }

    public enum PivotDenominatorBehavior
    {
        Unknown,
        Blank,
        Zero
    }

    public enum PivotVarianceConvention
    {
        Unknown,
        ActualMinusPlan,
        PlanMinusActual
    }

    public enum PivotShareDenominatorKind
    {
        Explicit,
        Parent,
        FilteredTotal
    }

    public enum PivotCalculationExpressionKind
    {
        Aggregate,
        FilteredAggregate,
        WeightedResult,
        MeasureReference,
        Difference,
        SafeRatio,
        Share,
        Growth,
        Achievement,
        Variance,
        VariancePercentage,
        PercentagePointDelta
    }

    /// <summary>
    /// A strongly typed DAX literal. No string in this object is interpreted as
    /// DAX source text.
    /// </summary>
    public sealed class PivotScalarValue
    {
        private PivotScalarValue(
            PivotScalarKind kind,
            string? text,
            long? wholeNumber,
            decimal? decimalNumber,
            bool? boolean,
            DateTime? temporal)
        {
            Kind = kind;
            TextValue = text;
            WholeNumberValue = wholeNumber;
            DecimalNumberValue = decimalNumber;
            BooleanValue = boolean;
            TemporalValue = temporal;
        }

        public PivotScalarKind Kind { get; }

        public string? TextValue { get; }

        public long? WholeNumberValue { get; }

        public decimal? DecimalNumberValue { get; }

        public bool? BooleanValue { get; }

        public DateTime? TemporalValue { get; }

        public static PivotScalarValue Blank()
        {
            return new PivotScalarValue(PivotScalarKind.Blank, null, null, null, null, null);
        }

        public static PivotScalarValue Text(string value)
        {
            if (value == null) throw new ArgumentNullException(nameof(value));
            return new PivotScalarValue(PivotScalarKind.Text, value, null, null, null, null);
        }

        public static PivotScalarValue WholeNumber(long value)
        {
            return new PivotScalarValue(PivotScalarKind.WholeNumber, null, value, null, null, null);
        }

        public static PivotScalarValue DecimalNumber(decimal value)
        {
            return new PivotScalarValue(PivotScalarKind.DecimalNumber, null, null, value, null, null);
        }

        public static PivotScalarValue Boolean(bool value)
        {
            return new PivotScalarValue(PivotScalarKind.Boolean, null, null, null, value, null);
        }

        public static PivotScalarValue Date(DateTime value)
        {
            return new PivotScalarValue(PivotScalarKind.Date, null, null, null, null, value.Date);
        }

        public static PivotScalarValue DateTime(DateTime value)
        {
            return new PivotScalarValue(PivotScalarKind.DateTime, null, null, null, null, value);
        }
    }

    public sealed class PivotModelMember
    {
        public PivotModelMember(string id, PivotScalarValue value, string? caption = null)
        {
            Id = id ?? string.Empty;
            Value = value ?? throw new ArgumentNullException(nameof(value));
            Caption = caption;
        }

        public string Id { get; }

        public PivotScalarValue Value { get; }

        public string? Caption { get; }
    }

    public sealed class PivotModelFieldSchema
    {
        public PivotModelFieldSchema(
            string id,
            string name,
            PivotModelDataType dataType,
            IEnumerable<PivotModelMember>? members = null)
        {
            Id = id ?? string.Empty;
            Name = name ?? string.Empty;
            DataType = dataType;
            Members = CalculationCollections.Copy(members);
        }

        public string Id { get; }

        public string Name { get; }

        public PivotModelDataType DataType { get; }

        public IReadOnlyList<PivotModelMember> Members { get; }
    }

    public sealed class PivotModelTableSchema
    {
        public PivotModelTableSchema(
            string id,
            string name,
            IEnumerable<PivotModelFieldSchema>? fields)
        {
            Id = id ?? string.Empty;
            Name = name ?? string.Empty;
            Fields = CalculationCollections.Copy(fields);
        }

        public string Id { get; }

        public string Name { get; }

        public IReadOnlyList<PivotModelFieldSchema> Fields { get; }
    }

    /// <summary>
    /// The complete model binding used by the compiler. Expressions refer to
    /// stable IDs; only this schema contains native table and field names.
    /// </summary>
    public sealed class PivotModelSchema
    {
        public PivotModelSchema(IEnumerable<PivotModelTableSchema>? tables)
        {
            Tables = CalculationCollections.Copy(tables);
        }

        public IReadOnlyList<PivotModelTableSchema> Tables { get; }
    }

    public sealed class PivotFilterValue
    {
        private PivotFilterValue(
            PivotFilterValueKind kind,
            PivotScalarValue? scalar,
            string? memberId)
        {
            Kind = kind;
            Scalar = scalar;
            MemberId = memberId;
        }

        public PivotFilterValueKind Kind { get; }

        public PivotScalarValue? Scalar { get; }

        public string? MemberId { get; }

        public static PivotFilterValue FromScalar(PivotScalarValue value)
        {
            return new PivotFilterValue(
                PivotFilterValueKind.Scalar,
                value ?? throw new ArgumentNullException(nameof(value)),
                null);
        }

        public static PivotFilterValue FromMember(string memberId)
        {
            return new PivotFilterValue(PivotFilterValueKind.Member, null, memberId ?? string.Empty);
        }
    }

    /// <summary>
    /// A bounded filter over one schema-bound field. Ordinary calculation
    /// filters intersect the current PivotTable context through KEEPFILTERS.
    /// </summary>
    public sealed class PivotCalculationFilter
    {
        public PivotCalculationFilter(
            string fieldId,
            PivotCalculationFilterOperator @operator,
            IEnumerable<PivotFilterValue>? values = null)
        {
            FieldId = fieldId ?? string.Empty;
            Operator = @operator;
            Values = CalculationCollections.Copy(values);
        }

        public string FieldId { get; }

        public PivotCalculationFilterOperator Operator { get; }

        public IReadOnlyList<PivotFilterValue> Values { get; }
    }

    public sealed class PivotPeriodPoint
    {
        public PivotPeriodPoint(
            PivotPeriodGrain grain,
            int year,
            int? ordinal = null,
            DateTime? date = null)
        {
            Grain = grain;
            Year = year;
            Ordinal = ordinal;
            Date = date;
        }

        public PivotPeriodGrain Grain { get; }

        public int Year { get; }

        /// <summary>
        /// One-based half, quarter, or month number. It is null for Year and
        /// Date points.
        /// </summary>
        public int? Ordinal { get; }

        public DateTime? Date { get; }
    }

    /// <summary>
    /// Binds one logical period bucket to its exact typed source member and to
    /// the scenario members whose coverage was observed for that bucket.
    /// </summary>
    public sealed class PivotPeriodCoverageMember
    {
        public PivotPeriodCoverageMember(
            PivotPeriodPoint point,
            PivotFilterValue sourceValue,
            IEnumerable<string>? scenarioMemberIds = null)
        {
            Point = point ?? throw new ArgumentNullException(nameof(point));
            SourceValue = sourceValue ?? throw new ArgumentNullException(nameof(sourceValue));
            ScenarioMemberIds = CalculationCollections.Copy(scenarioMemberIds);
        }

        public PivotPeriodPoint Point { get; }

        public PivotFilterValue SourceValue { get; }

        public IReadOnlyList<string> ScenarioMemberIds { get; }
    }

    public sealed class PivotPeriodSource
    {
        public PivotPeriodSource(
            string periodFieldId,
            PivotPeriodGrain sourceGrain,
            PivotPeriodCoverageStatus coverageStatus,
            IEnumerable<PivotPeriodCoverageMember>? coverage,
            IEnumerable<string>? periodContextFieldIds = null,
            string? scenarioFieldId = null,
            IEnumerable<string>? scenarioContextFieldIds = null,
            PivotDateCoverageMode dateCoverageMode = PivotDateCoverageMode.NotApplicable,
            DateTime? continuousRangeStart = null,
            DateTime? continuousRangeEnd = null,
            IEnumerable<string>? continuousRangeScenarioMemberIds = null)
        {
            PeriodFieldId = periodFieldId ?? string.Empty;
            SourceGrain = sourceGrain;
            CoverageStatus = coverageStatus;
            Coverage = CalculationCollections.Copy(coverage);
            PeriodContextFieldIds = CalculationCollections.Copy(
                periodContextFieldIds ?? new[] { PeriodFieldId });
            ScenarioFieldId = scenarioFieldId;
            ScenarioContextFieldIds = CalculationCollections.Copy(
                scenarioContextFieldIds ??
                (scenarioFieldId == null ? Array.Empty<string>() : new[] { scenarioFieldId }));
            DateCoverageMode = dateCoverageMode;
            ContinuousRangeStart = continuousRangeStart;
            ContinuousRangeEnd = continuousRangeEnd;
            ContinuousRangeScenarioMemberIds = CalculationCollections.Copy(
                continuousRangeScenarioMemberIds);
        }

        public string PeriodFieldId { get; }

        public PivotPeriodGrain SourceGrain { get; }

        public PivotPeriodCoverageStatus CoverageStatus { get; }

        public IReadOnlyList<PivotPeriodCoverageMember> Coverage { get; }

        /// <summary>
        /// Exact period hierarchy fields cleared by ReplaceAxisContext. The
        /// compiler never emits an unbounded table-wide ALL operation.
        /// </summary>
        public IReadOnlyList<string> PeriodContextFieldIds { get; }

        public string? ScenarioFieldId { get; }

        public IReadOnlyList<string> ScenarioContextFieldIds { get; }

        public PivotDateCoverageMode DateCoverageMode { get; }

        public DateTime? ContinuousRangeStart { get; }

        public DateTime? ContinuousRangeEnd { get; }

        /// <summary>
        /// Exact scenarios explicitly declared available throughout a complete
        /// continuous date range. This is not inferred from missing fact rows.
        /// </summary>
        public IReadOnlyList<string> ContinuousRangeScenarioMemberIds { get; }
    }

    public sealed class PivotPeriodSlice
    {
        public PivotPeriodSlice(
            string id,
            string caption,
            PivotPeriodPoint point,
            string? scenarioMemberId,
            PivotSliceFilterMode filterMode)
        {
            Id = id ?? string.Empty;
            Caption = caption ?? string.Empty;
            Point = point ?? throw new ArgumentNullException(nameof(point));
            ScenarioMemberId = scenarioMemberId;
            FilterMode = filterMode;
        }

        public string Id { get; }

        public string Caption { get; }

        public PivotPeriodPoint Point { get; }

        public string? ScenarioMemberId { get; }

        public PivotSliceFilterMode FilterMode { get; }
    }

    public sealed class PivotPeriodDefinition
    {
        public PivotPeriodDefinition(
            PivotPeriodSource source,
            IEnumerable<PivotPeriodSlice>? slices)
        {
            Source = source ?? throw new ArgumentNullException(nameof(source));
            Slices = CalculationCollections.Copy(slices);
        }

        public PivotPeriodSource Source { get; }

        public IReadOnlyList<PivotPeriodSlice> Slices { get; }
    }

    public abstract class PivotCalculationExpression
    {
        internal PivotCalculationExpression()
        {
        }

        public abstract PivotCalculationExpressionKind Kind { get; }
    }

    public sealed class PivotAggregateExpression : PivotCalculationExpression
    {
        public PivotAggregateExpression(
            string fieldId,
            PivotCalculationAggregateFunction function,
            string? periodSliceId = null)
        {
            FieldId = fieldId ?? string.Empty;
            Function = function;
            PeriodSliceId = periodSliceId;
        }

        public override PivotCalculationExpressionKind Kind =>
            PivotCalculationExpressionKind.Aggregate;

        public string FieldId { get; }

        public PivotCalculationAggregateFunction Function { get; }

        public string? PeriodSliceId { get; }
    }

    public sealed class PivotFilteredAggregateExpression : PivotCalculationExpression
    {
        public PivotFilteredAggregateExpression(
            string fieldId,
            PivotCalculationAggregateFunction function,
            IEnumerable<PivotCalculationFilter>? filters,
            string? periodSliceId = null)
        {
            FieldId = fieldId ?? string.Empty;
            Function = function;
            Filters = CalculationCollections.Copy(filters);
            PeriodSliceId = periodSliceId;
        }

        public override PivotCalculationExpressionKind Kind =>
            PivotCalculationExpressionKind.FilteredAggregate;

        public string FieldId { get; }

        public PivotCalculationAggregateFunction Function { get; }

        public IReadOnlyList<PivotCalculationFilter> Filters { get; }

        public string? PeriodSliceId { get; }
    }

    public sealed class PivotWeightedResultExpression : PivotCalculationExpression
    {
        public PivotWeightedResultExpression(
            string valueFieldId,
            string weightFieldId,
            PivotDenominatorBehavior onZero,
            IEnumerable<PivotCalculationFilter>? filters = null,
            string? periodSliceId = null)
        {
            ValueFieldId = valueFieldId ?? string.Empty;
            WeightFieldId = weightFieldId ?? string.Empty;
            OnZero = onZero;
            Filters = CalculationCollections.Copy(filters);
            PeriodSliceId = periodSliceId;
        }

        public override PivotCalculationExpressionKind Kind =>
            PivotCalculationExpressionKind.WeightedResult;

        public string ValueFieldId { get; }

        public string WeightFieldId { get; }

        public PivotDenominatorBehavior OnZero { get; }

        public IReadOnlyList<PivotCalculationFilter> Filters { get; }

        public string? PeriodSliceId { get; }
    }

    public sealed class PivotMeasureReferenceExpression : PivotCalculationExpression
    {
        public PivotMeasureReferenceExpression(string measureId)
        {
            MeasureId = measureId ?? string.Empty;
        }

        public override PivotCalculationExpressionKind Kind =>
            PivotCalculationExpressionKind.MeasureReference;

        public string MeasureId { get; }
    }

    public sealed class PivotDifferenceExpression : PivotCalculationExpression
    {
        public PivotDifferenceExpression(
            PivotCalculationExpression left,
            PivotCalculationExpression right)
        {
            Left = left ?? throw new ArgumentNullException(nameof(left));
            Right = right ?? throw new ArgumentNullException(nameof(right));
        }

        public override PivotCalculationExpressionKind Kind =>
            PivotCalculationExpressionKind.Difference;

        public PivotCalculationExpression Left { get; }

        public PivotCalculationExpression Right { get; }
    }

    public sealed class PivotSafeRatioExpression : PivotCalculationExpression
    {
        public PivotSafeRatioExpression(
            PivotCalculationExpression numerator,
            PivotCalculationExpression denominator,
            PivotDenominatorBehavior onZero)
        {
            Numerator = numerator ?? throw new ArgumentNullException(nameof(numerator));
            Denominator = denominator ?? throw new ArgumentNullException(nameof(denominator));
            OnZero = onZero;
        }

        public override PivotCalculationExpressionKind Kind =>
            PivotCalculationExpressionKind.SafeRatio;

        public PivotCalculationExpression Numerator { get; }

        public PivotCalculationExpression Denominator { get; }

        public PivotDenominatorBehavior OnZero { get; }
    }

    public abstract class PivotShareDenominator
    {
        internal PivotShareDenominator(PivotShareDenominatorKind kind)
        {
            Kind = kind;
        }

        public PivotShareDenominatorKind Kind { get; }
    }

    public sealed class PivotExplicitShareDenominator : PivotShareDenominator
    {
        public PivotExplicitShareDenominator(PivotCalculationExpression expression)
            : base(PivotShareDenominatorKind.Explicit)
        {
            Expression = expression ?? throw new ArgumentNullException(nameof(expression));
        }

        public PivotCalculationExpression Expression { get; }
    }

    public sealed class PivotParentShareDenominator : PivotShareDenominator
    {
        public PivotParentShareDenominator(IEnumerable<string>? clearedFieldIds)
            : base(PivotShareDenominatorKind.Parent)
        {
            ClearedFieldIds = CalculationCollections.Copy(clearedFieldIds);
        }

        public IReadOnlyList<string> ClearedFieldIds { get; }
    }

    public sealed class PivotFilteredTotalShareDenominator : PivotShareDenominator
    {
        public PivotFilteredTotalShareDenominator(IEnumerable<string>? clearedFieldIds)
            : base(PivotShareDenominatorKind.FilteredTotal)
        {
            ClearedFieldIds = CalculationCollections.Copy(clearedFieldIds);
        }

        public IReadOnlyList<string> ClearedFieldIds { get; }
    }

    public sealed class PivotShareExpression : PivotCalculationExpression
    {
        public PivotShareExpression(
            PivotCalculationExpression part,
            PivotShareDenominator denominator,
            PivotDenominatorBehavior onZero)
        {
            Part = part ?? throw new ArgumentNullException(nameof(part));
            Denominator = denominator ?? throw new ArgumentNullException(nameof(denominator));
            OnZero = onZero;
        }

        public override PivotCalculationExpressionKind Kind =>
            PivotCalculationExpressionKind.Share;

        public PivotCalculationExpression Part { get; }

        public PivotShareDenominator Denominator { get; }

        public PivotDenominatorBehavior OnZero { get; }
    }

    public sealed class PivotGrowthExpression : PivotCalculationExpression
    {
        public PivotGrowthExpression(
            PivotCalculationExpression current,
            PivotCalculationExpression prior,
            PivotDenominatorBehavior onZero)
        {
            Current = current ?? throw new ArgumentNullException(nameof(current));
            Prior = prior ?? throw new ArgumentNullException(nameof(prior));
            OnZero = onZero;
        }

        public override PivotCalculationExpressionKind Kind =>
            PivotCalculationExpressionKind.Growth;

        public PivotCalculationExpression Current { get; }

        public PivotCalculationExpression Prior { get; }

        public PivotDenominatorBehavior OnZero { get; }
    }

    public sealed class PivotAchievementExpression : PivotCalculationExpression
    {
        public PivotAchievementExpression(
            PivotCalculationExpression actual,
            PivotCalculationExpression target,
            PivotDenominatorBehavior onZero)
        {
            Actual = actual ?? throw new ArgumentNullException(nameof(actual));
            Target = target ?? throw new ArgumentNullException(nameof(target));
            OnZero = onZero;
        }

        public override PivotCalculationExpressionKind Kind =>
            PivotCalculationExpressionKind.Achievement;

        public PivotCalculationExpression Actual { get; }

        public PivotCalculationExpression Target { get; }

        public PivotDenominatorBehavior OnZero { get; }
    }

    public sealed class PivotVarianceExpression : PivotCalculationExpression
    {
        public PivotVarianceExpression(
            PivotCalculationExpression actual,
            PivotCalculationExpression plan,
            PivotVarianceConvention convention)
        {
            Actual = actual ?? throw new ArgumentNullException(nameof(actual));
            Plan = plan ?? throw new ArgumentNullException(nameof(plan));
            Convention = convention;
        }

        public override PivotCalculationExpressionKind Kind =>
            PivotCalculationExpressionKind.Variance;

        public PivotCalculationExpression Actual { get; }

        public PivotCalculationExpression Plan { get; }

        public PivotVarianceConvention Convention { get; }
    }

    public sealed class PivotVariancePercentageExpression : PivotCalculationExpression
    {
        public PivotVariancePercentageExpression(
            PivotCalculationExpression actual,
            PivotCalculationExpression plan,
            PivotVarianceConvention convention,
            PivotDenominatorBehavior onZero)
        {
            Actual = actual ?? throw new ArgumentNullException(nameof(actual));
            Plan = plan ?? throw new ArgumentNullException(nameof(plan));
            Convention = convention;
            OnZero = onZero;
        }

        public override PivotCalculationExpressionKind Kind =>
            PivotCalculationExpressionKind.VariancePercentage;

        public PivotCalculationExpression Actual { get; }

        public PivotCalculationExpression Plan { get; }

        public PivotVarianceConvention Convention { get; }

        public PivotDenominatorBehavior OnZero { get; }
    }

    public sealed class PivotPercentagePointDeltaExpression : PivotCalculationExpression
    {
        public PivotPercentagePointDeltaExpression(
            PivotCalculationExpression currentRatio,
            PivotCalculationExpression baselineRatio)
        {
            CurrentRatio = currentRatio ?? throw new ArgumentNullException(nameof(currentRatio));
            BaselineRatio = baselineRatio ?? throw new ArgumentNullException(nameof(baselineRatio));
        }

        public override PivotCalculationExpressionKind Kind =>
            PivotCalculationExpressionKind.PercentagePointDelta;

        public PivotCalculationExpression CurrentRatio { get; }

        public PivotCalculationExpression BaselineRatio { get; }
    }

    public sealed class PivotMeasureDefinition
    {
        public PivotMeasureDefinition(
            string id,
            string caption,
            string homeTableId,
            PivotMeasureFormat format,
            PivotCalculationExpression expression)
        {
            Id = id ?? string.Empty;
            Caption = caption ?? string.Empty;
            HomeTableId = homeTableId ?? string.Empty;
            Format = format ?? throw new ArgumentNullException(nameof(format));
            Expression = expression ?? throw new ArgumentNullException(nameof(expression));
        }

        public string Id { get; }

        public string Caption { get; }

        public string HomeTableId { get; }

        public PivotMeasureFormat Format { get; }

        public PivotCalculationExpression Expression { get; }
    }

    public sealed class PivotMeasureSetDefinition
    {
        public PivotMeasureSetDefinition(
            PivotModelSchema schema,
            IEnumerable<PivotMeasureDefinition>? measures,
            PivotPeriodDefinition? periods = null)
        {
            Schema = schema ?? throw new ArgumentNullException(nameof(schema));
            Measures = CalculationCollections.Copy(measures);
            Periods = periods;
        }

        public PivotModelSchema Schema { get; }

        /// <summary>
        /// Requested display order. The compiler derives a separate stable
        /// topological creation order from measure references.
        /// </summary>
        public IReadOnlyList<PivotMeasureDefinition> Measures { get; }

        public PivotPeriodDefinition? Periods { get; }
    }

    /// <summary>
    /// Transient compiler output for one owned Data Model measure. Workbook
    /// ownership metadata persists Fingerprint, never DaxFormula.
    /// </summary>
    public sealed class OwnedPivotMeasureDefinition
    {
        internal OwnedPivotMeasureDefinition(
            string definitionId,
            int displayOrder,
            int creationOrder,
            string homeTableName,
            string generatedMeasureName,
            string daxFormula,
            PivotMeasureFormat format,
            IEnumerable<string> directDependencyDefinitionIds,
            string definitionFingerprint,
            string formulaFingerprint)
        {
            DefinitionId = definitionId;
            DisplayOrder = displayOrder;
            CreationOrder = creationOrder;
            HomeTableName = homeTableName;
            GeneratedMeasureName = generatedMeasureName;
            DaxFormula = daxFormula;
            Format = format;
            DirectDependencyDefinitionIds = CalculationCollections.Copy(
                directDependencyDefinitionIds);
            DefinitionFingerprint = definitionFingerprint;
            FormulaFingerprint = formulaFingerprint;
        }

        public string DefinitionId { get; }

        public int DisplayOrder { get; }

        public int CreationOrder { get; }

        public string HomeTableName { get; }

        /// <summary>
        /// The globally unique native ModelMeasure.Name. It is also the stable
        /// caption shown by Excel for the owned measure.
        /// </summary>
        public string GeneratedMeasureName { get; }

        public string DaxFormula { get; }

        public PivotMeasureFormat Format { get; }

        public PivotMeasureFormatKind FormatKind => Format.Kind;

        public IReadOnlyList<string> DirectDependencyDefinitionIds { get; }

        /// <summary>
        /// Stable hash of the typed semantic definition. It does not depend on
        /// how a live Excel host later canonicalizes the DAX text.
        /// </summary>
        public string DefinitionFingerprint { get; }

        /// <summary>
        /// Hash of the compiler's exact transient DAX output. Formula text is
        /// never persisted in PivotTable+ workbook ownership metadata.
        /// </summary>
        public string FormulaFingerprint { get; }
    }

    public sealed class PivotDaxCompilation
    {
        internal PivotDaxCompilation(IEnumerable<OwnedPivotMeasureDefinition> measures)
        {
            Measures = CalculationCollections.Copy(measures);
            CreationSequence = CalculationCollections.Copy(
                Measures.OrderBy(measure => measure.CreationOrder));
        }

        public IReadOnlyList<OwnedPivotMeasureDefinition> Measures { get; }

        public IReadOnlyList<OwnedPivotMeasureDefinition> CreationSequence { get; }
    }

    public sealed class InvalidPivotCalculationException : Exception
    {
        public InvalidPivotCalculationException(ValidationResult validation)
            : base("The PivotTable+ calculation definition is invalid.")
        {
            Validation = validation ?? throw new ArgumentNullException(nameof(validation));
        }

        public ValidationResult Validation { get; }
    }

    internal static class CalculationCollections
    {
        public static IReadOnlyList<T> Copy<T>(IEnumerable<T>? values)
        {
            return new ReadOnlyCollection<T>(
                (values ?? Enumerable.Empty<T>()).ToList());
        }
    }
}
