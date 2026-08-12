using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelReportBuilder.Core.Validation;

namespace ExcelReportBuilder.Core.PivotPlus.Calculations
{
    public static class PivotCalculationValidator
    {
        private const int MaximumTables = 128;
        private const int MaximumFields = 2048;
        private const int MaximumFieldsPerTable = 512;
        private const int MaximumMembersPerField = 10000;
        private const int MaximumMeasures = 128;
        private const int MaximumExpressionDepth = 32;
        private const int MaximumExpressionNodes = 256;
        private const int MaximumDependenciesPerMeasure = 32;
        private const int MaximumFilters = 32;
        private const int MaximumFilterValues = 256;
        private const int MaximumPeriodCoverageMembers = 10000;
        private const int MaximumPeriodSlices = 256;
        private const int MaximumContextFields = 16;
        private const int MaximumIdLength = 128;
        private const int MaximumNativeNameLength = 255;
        private const int MaximumCaptionLength = 255;
        private const int MaximumMemberCaptionLength = 1024;
        private const int MaximumCurrencyMarkerLength = 8;

        public static ValidationResult Validate(PivotMeasureSetDefinition definition)
        {
            if (definition == null) throw new ArgumentNullException(nameof(definition));

            var result = new ValidationResult();
            ValidateSchema(definition.Schema, result);
            var model = new PivotCalculationModelIndex(definition.Schema);
            Dictionary<string, PivotPeriodSlice> slices = ValidatePeriods(
                definition.Periods,
                model,
                result);
            ValidateMeasures(definition, model, slices, result);
            return result;
        }

        private static void ValidateSchema(PivotModelSchema schema, ValidationResult result)
        {
            if (schema.Tables.Count == 0)
            {
                result.AddError(
                    "PIVOT_CALC_SCHEMA_TABLE_REQUIRED",
                    "schema.tables",
                    "At least one Data Model table is required.");
            }
            else if (schema.Tables.Count > MaximumTables)
            {
                result.AddError(
                    "PIVOT_CALC_SCHEMA_TABLE_LIMIT",
                    "schema.tables",
                    "The model schema exceeds the bounded table limit.");
            }

            var tableIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var tableNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var fieldIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var totalFields = 0;
            for (var tableIndex = 0; tableIndex < schema.Tables.Count; tableIndex++)
            {
                PivotModelTableSchema? table = schema.Tables[tableIndex];
                string path = "schema.tables[" + tableIndex.ToString(CultureInfo.InvariantCulture) + "]";
                if (table == null)
                {
                    result.AddError(
                        "PIVOT_CALC_SCHEMA_TABLE_NULL",
                        path,
                        "Model table entries cannot be null.");
                    continue;
                }

                ValidateId(table.Id, path + ".id", "PIVOT_CALC_TABLE_ID_INVALID", result);
                if (!tableIds.Add(table.Id))
                {
                    result.AddError(
                        "PIVOT_CALC_TABLE_ID_DUPLICATE",
                        path + ".id",
                        "Model table IDs must be unique without regard to case.");
                }

                ValidateNativeName(
                    table.Name,
                    path + ".name",
                    "PIVOT_CALC_TABLE_NAME_INVALID",
                    result);
                if (!tableNames.Add(table.Name))
                {
                    result.AddError(
                        "PIVOT_CALC_TABLE_NAME_DUPLICATE",
                        path + ".name",
                        "Native Data Model table names must be unique without regard to case.");
                }

                if (table.Fields.Count == 0)
                {
                    result.AddError(
                        "PIVOT_CALC_SCHEMA_FIELD_REQUIRED",
                        path + ".fields",
                        "Every model table requires at least one field binding.");
                }
                else if (table.Fields.Count > MaximumFieldsPerTable)
                {
                    result.AddError(
                        "PIVOT_CALC_SCHEMA_FIELD_TABLE_LIMIT",
                        path + ".fields",
                        "A model table exceeds the bounded field limit.");
                }

                totalFields += table.Fields.Count;
                var fieldNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                for (var fieldIndex = 0; fieldIndex < table.Fields.Count; fieldIndex++)
                {
                    PivotModelFieldSchema? field = table.Fields[fieldIndex];
                    string fieldPath = path + ".fields[" +
                        fieldIndex.ToString(CultureInfo.InvariantCulture) + "]";
                    if (field == null)
                    {
                        result.AddError(
                            "PIVOT_CALC_SCHEMA_FIELD_NULL",
                            fieldPath,
                            "Model field entries cannot be null.");
                        continue;
                    }

                    ValidateId(
                        field.Id,
                        fieldPath + ".id",
                        "PIVOT_CALC_FIELD_ID_INVALID",
                        result);
                    if (!fieldIds.Add(field.Id))
                    {
                        result.AddError(
                            "PIVOT_CALC_FIELD_ID_DUPLICATE",
                            fieldPath + ".id",
                            "Field IDs must be globally unique without regard to case.");
                    }

                    ValidateNativeName(
                        field.Name,
                        fieldPath + ".name",
                        "PIVOT_CALC_FIELD_NAME_INVALID",
                        result);
                    if (!fieldNames.Add(field.Name))
                    {
                        result.AddError(
                            "PIVOT_CALC_FIELD_NAME_DUPLICATE",
                            fieldPath + ".name",
                            "Native field names must be unique within a model table.");
                    }

                    if (!Enum.IsDefined(typeof(PivotModelDataType), field.DataType) ||
                        field.DataType == PivotModelDataType.Unknown)
                    {
                        result.AddError(
                            "PIVOT_CALC_FIELD_DATA_TYPE_INVALID",
                            fieldPath + ".dataType",
                            "Every field requires a supported model data type.");
                    }

                    ValidateMembers(field, fieldPath + ".members", result);
                }
            }

            if (totalFields > MaximumFields)
            {
                result.AddError(
                    "PIVOT_CALC_SCHEMA_FIELD_LIMIT",
                    "schema.tables",
                    "The model schema exceeds the bounded total field limit.");
            }
        }

        private static void ValidateMembers(
            PivotModelFieldSchema field,
            string path,
            ValidationResult result)
        {
            if (field.Members.Count > MaximumMembersPerField)
            {
                result.AddError(
                    "PIVOT_CALC_MEMBER_LIMIT",
                    path,
                    "The field exceeds the bounded member limit.");
            }

            var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var values = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var index = 0; index < field.Members.Count; index++)
            {
                PivotModelMember? member = field.Members[index];
                string memberPath = path + "[" + index.ToString(CultureInfo.InvariantCulture) + "]";
                if (member == null)
                {
                    result.AddError(
                        "PIVOT_CALC_MEMBER_NULL",
                        memberPath,
                        "Model member entries cannot be null.");
                    continue;
                }

                ValidateId(
                    member.Id,
                    memberPath + ".id",
                    "PIVOT_CALC_MEMBER_ID_INVALID",
                    result);
                if (!ids.Add(member.Id))
                {
                    result.AddError(
                        "PIVOT_CALC_MEMBER_ID_DUPLICATE",
                        memberPath + ".id",
                        "Member IDs must be unique without regard to case.");
                }

                if (member.Caption != null &&
                    !IsBoundedText(member.Caption, MaximumMemberCaptionLength, allowEmpty: false))
                {
                    result.AddError(
                        "PIVOT_CALC_MEMBER_CAPTION_INVALID",
                        memberPath + ".caption",
                        "Member captions must be bounded printable text.");
                }

                if (!IsScalarCompatible(field.DataType, member.Value))
                {
                    result.AddError(
                        "PIVOT_CALC_MEMBER_TYPE_MISMATCH",
                        memberPath + ".value",
                        "The member's typed value does not match its model field.");
                }

                ValidateScalarShape(member.Value, memberPath + ".value", result);

                string key = PivotCalculationCanonical.ScalarKey(member.Value);
                if (!values.Add(key))
                {
                    result.AddError(
                        "PIVOT_CALC_MEMBER_VALUE_DUPLICATE",
                        memberPath + ".value",
                        "Member values must resolve to one exact source member.");
                }
            }
        }

        private static Dictionary<string, PivotPeriodSlice> ValidatePeriods(
            PivotPeriodDefinition? periods,
            PivotCalculationModelIndex model,
            ValidationResult result)
        {
            var slices = new Dictionary<string, PivotPeriodSlice>(StringComparer.OrdinalIgnoreCase);
            if (periods == null)
            {
                return slices;
            }

            PivotPeriodSource source = periods.Source;
            if (!model.TryGetField(source.PeriodFieldId, out PivotBoundField periodField))
            {
                result.AddError(
                    "PIVOT_CALC_PERIOD_FIELD_UNKNOWN",
                    "periods.source.periodFieldId",
                    "The period field is not present in the model schema.");
            }
            else if (!SupportsPeriodValues(periodField.Field.DataType))
            {
                result.AddError(
                    "PIVOT_CALC_PERIOD_FIELD_TYPE_INVALID",
                    "periods.source.periodFieldId",
                    "The period field must be text, whole number, date, or date-time.");
            }

            if (!Enum.IsDefined(typeof(PivotPeriodGrain), source.SourceGrain) ||
                source.SourceGrain == PivotPeriodGrain.Unknown)
            {
                result.AddError(
                    "PIVOT_CALC_PERIOD_SOURCE_GRAIN_INVALID",
                    "periods.source.sourceGrain",
                    "The source period grain must be explicit.");
            }

            if (source.CoverageStatus != PivotPeriodCoverageStatus.Complete)
            {
                result.AddError(
                    "PIVOT_CALC_PERIOD_COVERAGE_INCOMPLETE",
                    "periods.source.coverageStatus",
                    "Period slices require explicitly complete source coverage.");
            }

            ValidateContextFields(
                source.PeriodContextFieldIds,
                source.PeriodFieldId,
                "periods.source.periodContextFieldIds",
                model,
                result);

            PivotBoundField? scenarioField = null;
            if (source.ScenarioFieldId == null)
            {
                if (source.ScenarioContextFieldIds.Count > 0)
                {
                    result.AddError(
                        "PIVOT_CALC_SCENARIO_CONTEXT_WITHOUT_FIELD",
                        "periods.source.scenarioContextFieldIds",
                        "Scenario context fields require a scenario field binding.");
                }
            }
            else
            {
                if (!model.TryGetField(source.ScenarioFieldId, out PivotBoundField boundScenario))
                {
                    result.AddError(
                        "PIVOT_CALC_SCENARIO_FIELD_UNKNOWN",
                        "periods.source.scenarioFieldId",
                        "The scenario field is not present in the model schema.");
                }
                else
                {
                    scenarioField = boundScenario;
                    if (scenarioField.Field.Members.Count == 0)
                    {
                        result.AddError(
                            "PIVOT_CALC_SCENARIO_MEMBERS_REQUIRED",
                            "periods.source.scenarioFieldId",
                            "Scenario slices require exact schema-bound scenario members.");
                    }
                }

                ValidateContextFields(
                    source.ScenarioContextFieldIds,
                    source.ScenarioFieldId,
                    "periods.source.scenarioContextFieldIds",
                    model,
                    result);

                if (string.Equals(
                        source.PeriodFieldId,
                        source.ScenarioFieldId,
                        StringComparison.OrdinalIgnoreCase))
                {
                    result.AddError(
                        "PIVOT_CALC_PERIOD_SCENARIO_FIELD_CONFLICT",
                        "periods.source.scenarioFieldId",
                        "Period and scenario slices require different bound fields.");
                }

                if (source.PeriodContextFieldIds.Intersect(
                        source.ScenarioContextFieldIds,
                        StringComparer.OrdinalIgnoreCase).Any())
                {
                    result.AddError(
                        "PIVOT_CALC_PERIOD_SCENARIO_CONTEXT_OVERLAP",
                        "periods.source.scenarioContextFieldIds",
                        "Period and scenario replacement contexts cannot overlap.");
                }
            }

            ValidatePeriodCoverage(source, periodField, scenarioField, model, result);

            if (periods.Slices.Count == 0)
            {
                result.AddError(
                    "PIVOT_CALC_PERIOD_SLICE_REQUIRED",
                    "periods.slices",
                    "A period definition requires at least one requested slice.");
            }
            else if (periods.Slices.Count > MaximumPeriodSlices)
            {
                result.AddError(
                    "PIVOT_CALC_PERIOD_SLICE_LIMIT",
                    "periods.slices",
                    "The period definition exceeds the bounded slice limit.");
            }

            var captions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var index = 0; index < periods.Slices.Count; index++)
            {
                PivotPeriodSlice? slice = periods.Slices[index];
                string path = "periods.slices[" + index.ToString(CultureInfo.InvariantCulture) + "]";
                if (slice == null)
                {
                    result.AddError(
                        "PIVOT_CALC_PERIOD_SLICE_NULL",
                        path,
                        "Period slice entries cannot be null.");
                    continue;
                }

                ValidateId(slice.Id, path + ".id", "PIVOT_CALC_PERIOD_SLICE_ID_INVALID", result);
                if (slices.ContainsKey(slice.Id))
                {
                    result.AddError(
                        "PIVOT_CALC_PERIOD_SLICE_ID_DUPLICATE",
                        path + ".id",
                        "Period slice IDs must be unique without regard to case.");
                }
                else
                {
                    slices.Add(slice.Id, slice);
                }

                if (!IsBoundedText(slice.Caption, MaximumCaptionLength, allowEmpty: false) ||
                    !PivotPlusPathPolicy.IsPathFree(slice.Caption))
                {
                    result.AddError(
                        "PIVOT_CALC_PERIOD_SLICE_CAPTION_INVALID",
                        path + ".caption",
                        "Period slice captions must be bounded printable text.");
                }
                else if (!captions.Add(slice.Caption))
                {
                    result.AddError(
                        "PIVOT_CALC_PERIOD_SLICE_CAPTION_DUPLICATE",
                        path + ".caption",
                        "Period slice captions must be unique without regard to case.");
                }

                ValidatePeriodPoint(slice.Point, path + ".point", result);
                if (!Enum.IsDefined(typeof(PivotSliceFilterMode), slice.FilterMode) ||
                    slice.FilterMode == PivotSliceFilterMode.Unknown)
                {
                    result.AddError(
                        "PIVOT_CALC_PERIOD_FILTER_MODE_INVALID",
                        path + ".filterMode",
                        "The period slice must explicitly replace or intersect axis context.");
                }

                int sourceRank = PivotPeriodRules.GrainRank(source.SourceGrain);
                int requestedRank = PivotPeriodRules.GrainRank(slice.Point.Grain);
                if (sourceRank >= 0 && requestedRank >= 0 && sourceRank < requestedRank)
                {
                    result.AddError(
                        "PIVOT_CALC_PERIOD_GRAIN_TOO_COARSE",
                        path + ".point.grain",
                        "The source grain is too coarse for the requested period slice.");
                }

                ValidateSliceScenario(slice, path, source, scenarioField, result);
                ValidateSliceCoverage(periods, slice, path, result);
            }

            return slices;
        }

        private static void ValidateContextFields(
            IReadOnlyList<string> fieldIds,
            string requiredFieldId,
            string path,
            PivotCalculationModelIndex model,
            ValidationResult result)
        {
            if (fieldIds.Count == 0)
            {
                result.AddError(
                    "PIVOT_CALC_CONTEXT_FIELD_REQUIRED",
                    path,
                    "A replacement context requires at least its bound axis field.");
                return;
            }

            if (fieldIds.Count > MaximumContextFields)
            {
                result.AddError(
                    "PIVOT_CALC_CONTEXT_FIELD_LIMIT",
                    path,
                    "The replacement context exceeds the bounded field limit.");
            }

            var unique = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var index = 0; index < fieldIds.Count; index++)
            {
                string fieldId = fieldIds[index];
                string fieldPath = path + "[" + index.ToString(CultureInfo.InvariantCulture) + "]";
                ValidateId(fieldId, fieldPath, "PIVOT_CALC_CONTEXT_FIELD_ID_INVALID", result);
                if (string.IsNullOrWhiteSpace(fieldId))
                {
                    continue;
                }

                if (!unique.Add(fieldId))
                {
                    result.AddError(
                        "PIVOT_CALC_CONTEXT_FIELD_DUPLICATE",
                        fieldPath,
                        "Replacement context fields must be distinct.");
                }

                if (!model.TryGetField(fieldId, out _))
                {
                    result.AddError(
                        "PIVOT_CALC_CONTEXT_FIELD_UNKNOWN",
                        fieldPath,
                        "A replacement context field is not present in the model schema.");
                }
            }

            if (!unique.Contains(requiredFieldId))
            {
                result.AddError(
                    "PIVOT_CALC_CONTEXT_BOUND_FIELD_REQUIRED",
                    path,
                    "The replacement context must include its bound period or scenario field.");
            }
        }

        private static void ValidatePeriodCoverage(
            PivotPeriodSource source,
            PivotBoundField? periodField,
            PivotBoundField? scenarioField,
            PivotCalculationModelIndex model,
            ValidationResult result)
        {
            ValidateDateCoverageMode(source, periodField, scenarioField, result);
            if (source.DateCoverageMode == PivotDateCoverageMode.ContinuousRange)
            {
                return;
            }

            if (source.Coverage.Count == 0)
            {
                result.AddError(
                    "PIVOT_CALC_PERIOD_COVERAGE_REQUIRED",
                    "periods.source.coverage",
                    "Complete period coverage requires exact source members.");
            }
            else if (source.Coverage.Count > MaximumPeriodCoverageMembers)
            {
                result.AddError(
                    "PIVOT_CALC_PERIOD_COVERAGE_LIMIT",
                    "periods.source.coverage",
                    "Period coverage exceeds the bounded member limit.");
            }

            var points = new HashSet<string>(StringComparer.Ordinal);
            var sourceValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var index = 0; index < source.Coverage.Count; index++)
            {
                PivotPeriodCoverageMember? member = source.Coverage[index];
                string path = "periods.source.coverage[" +
                    index.ToString(CultureInfo.InvariantCulture) + "]";
                if (member == null)
                {
                    result.AddError(
                        "PIVOT_CALC_PERIOD_COVERAGE_NULL",
                        path,
                        "Period coverage entries cannot be null.");
                    continue;
                }

                ValidatePeriodPoint(member.Point, path + ".point", result);
                if (member.Point.Grain != source.SourceGrain)
                {
                    result.AddError(
                        "PIVOT_CALC_PERIOD_COVERAGE_GRAIN_MISMATCH",
                        path + ".point.grain",
                        "Every coverage point must use the declared source grain.");
                }

                if (!points.Add(PivotCalculationCanonical.PeriodPointKey(member.Point)))
                {
                    result.AddError(
                        "PIVOT_CALC_PERIOD_COVERAGE_POINT_DUPLICATE",
                        path + ".point",
                        "A logical source period can appear only once in coverage.");
                }

                if (periodField != null)
                {
                    PivotScalarValue? scalar = ValidateFilterValue(
                        member.SourceValue,
                        periodField,
                        path + ".sourceValue",
                        model,
                        result);
                    if (scalar != null)
                    {
                        if (scalar.Kind == PivotScalarKind.Blank)
                        {
                            result.AddError(
                                "PIVOT_CALC_PERIOD_SOURCE_VALUE_BLANK",
                                path + ".sourceValue",
                                "A period coverage member cannot bind to blank.");
                        }

                        if (!sourceValues.Add(PivotCalculationCanonical.ScalarKey(scalar)))
                        {
                            result.AddError(
                                "PIVOT_CALC_PERIOD_SOURCE_VALUE_DUPLICATE",
                                path + ".sourceValue",
                                "Period coverage values must identify one logical bucket each.");
                        }

                        if ((periodField.Field.DataType == PivotModelDataType.Date ||
                             periodField.Field.DataType == PivotModelDataType.DateTime) &&
                            scalar.TemporalValue.HasValue &&
                            !TemporalMatchesPoint(scalar.TemporalValue.Value, member.Point))
                        {
                            result.AddError(
                                "PIVOT_CALC_PERIOD_SOURCE_VALUE_MISMATCH",
                                path + ".sourceValue",
                                "The temporal source value does not belong to its declared logical period.");
                        }
                    }
                }

                ValidateCoverageScenarios(member, path, source, scenarioField, result);
            }
        }

        private static void ValidateDateCoverageMode(
            PivotPeriodSource source,
            PivotBoundField? periodField,
            PivotBoundField? scenarioField,
            ValidationResult result)
        {
            if (source.SourceGrain != PivotPeriodGrain.Date)
            {
                if (source.DateCoverageMode != PivotDateCoverageMode.NotApplicable ||
                    source.ContinuousRangeStart.HasValue ||
                    source.ContinuousRangeEnd.HasValue ||
                    source.ContinuousRangeScenarioMemberIds.Count > 0)
                {
                    result.AddError(
                        "PIVOT_CALC_DATE_COVERAGE_UNEXPECTED",
                        "periods.source.dateCoverageMode",
                        "Date coverage settings apply only to a full-date source grain.");
                }

                return;
            }

            if (source.DateCoverageMode != PivotDateCoverageMode.ExplicitCalendarMembers &&
                source.DateCoverageMode != PivotDateCoverageMode.ContinuousRange)
            {
                result.AddError(
                    "PIVOT_CALC_DATE_COVERAGE_MODE_REQUIRED",
                    "periods.source.dateCoverageMode",
                    "Date rollups require explicit calendar members or a declared continuous range.");
                return;
            }

            if (source.DateCoverageMode == PivotDateCoverageMode.ExplicitCalendarMembers)
            {
                if (source.ContinuousRangeStart.HasValue ||
                    source.ContinuousRangeEnd.HasValue ||
                    source.ContinuousRangeScenarioMemberIds.Count > 0)
                {
                    result.AddError(
                        "PIVOT_CALC_DATE_RANGE_UNEXPECTED",
                        "periods.source",
                        "Explicit calendar-member coverage cannot also declare a continuous range.");
                }

                return;
            }

            if (source.Coverage.Count > 0)
            {
                result.AddError(
                    "PIVOT_CALC_DATE_RANGE_MEMBER_CONFLICT",
                    "periods.source.coverage",
                    "Continuous date coverage cannot also enumerate calendar members.");
            }

            if (periodField != null &&
                periodField.Field.DataType != PivotModelDataType.Date &&
                periodField.Field.DataType != PivotModelDataType.DateTime)
            {
                result.AddError(
                    "PIVOT_CALC_DATE_RANGE_FIELD_TYPE_INVALID",
                    "periods.source.periodFieldId",
                    "Continuous date coverage requires a Date or DateTime model field.");
            }

            if (!source.ContinuousRangeStart.HasValue ||
                !source.ContinuousRangeEnd.HasValue ||
                source.ContinuousRangeStart.Value.TimeOfDay != TimeSpan.Zero ||
                source.ContinuousRangeEnd.Value.TimeOfDay != TimeSpan.Zero ||
                source.ContinuousRangeStart.Value > source.ContinuousRangeEnd.Value ||
                source.ContinuousRangeStart.Value.Year < 1900 ||
                source.ContinuousRangeEnd.Value.Year > 9999)
            {
                result.AddError(
                    "PIVOT_CALC_DATE_RANGE_INVALID",
                    "periods.source.continuousRangeStart",
                    "Continuous date coverage requires an ordered inclusive date-only range.");
            }

            var scenarios = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var index = 0; index < source.ContinuousRangeScenarioMemberIds.Count; index++)
            {
                string memberId = source.ContinuousRangeScenarioMemberIds[index];
                string path = "periods.source.continuousRangeScenarioMemberIds[" +
                    index.ToString(CultureInfo.InvariantCulture) + "]";
                ValidateId(
                    memberId,
                    path,
                    "PIVOT_CALC_SCENARIO_MEMBER_ID_INVALID",
                    result);
                if (string.IsNullOrWhiteSpace(memberId))
                {
                    continue;
                }

                if (!scenarios.Add(memberId))
                {
                    result.AddError(
                        "PIVOT_CALC_DATE_RANGE_SCENARIO_DUPLICATE",
                        path,
                        "Continuous-range scenario members must be distinct.");
                }

                if (scenarioField != null && !scenarioField.Field.Members.Any(candidate =>
                        candidate != null &&
                        string.Equals(candidate.Id, memberId, StringComparison.OrdinalIgnoreCase)))
                {
                    result.AddError(
                        "PIVOT_CALC_SCENARIO_MEMBER_UNKNOWN",
                        path,
                        "The continuous-range scenario member is not present in the model schema.");
                }
            }

            if (source.ScenarioFieldId == null &&
                source.ContinuousRangeScenarioMemberIds.Count > 0)
            {
                result.AddError(
                    "PIVOT_CALC_DATE_RANGE_SCENARIO_UNBOUND",
                    "periods.source.continuousRangeScenarioMemberIds",
                    "Continuous-range scenarios require a bound scenario field.");
            }
            else if (source.ScenarioFieldId != null &&
                     source.ContinuousRangeScenarioMemberIds.Count == 0)
            {
                result.AddError(
                    "PIVOT_CALC_DATE_RANGE_SCENARIO_REQUIRED",
                    "periods.source.continuousRangeScenarioMemberIds",
                    "A scenario-bound continuous range requires exact covered scenario members.");
            }
        }

        private static void ValidateCoverageScenarios(
            PivotPeriodCoverageMember member,
            string path,
            PivotPeriodSource source,
            PivotBoundField? scenarioField,
            ValidationResult result)
        {
            if (source.ScenarioFieldId == null)
            {
                if (member.ScenarioMemberIds.Count > 0)
                {
                    result.AddError(
                        "PIVOT_CALC_PERIOD_SCENARIO_WITHOUT_FIELD",
                        path + ".scenarioMemberIds",
                        "Scenario coverage requires a bound scenario field.");
                }

                return;
            }

            if (member.ScenarioMemberIds.Count == 0)
            {
                result.AddError(
                    "PIVOT_CALC_PERIOD_SCENARIO_COVERAGE_REQUIRED",
                    path + ".scenarioMemberIds",
                    "Complete scenario coverage is required for every period bucket.");
                return;
            }

            var unique = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var index = 0; index < member.ScenarioMemberIds.Count; index++)
            {
                string memberId = member.ScenarioMemberIds[index];
                string memberPath = path + ".scenarioMemberIds[" +
                    index.ToString(CultureInfo.InvariantCulture) + "]";
                ValidateId(
                    memberId,
                    memberPath,
                    "PIVOT_CALC_SCENARIO_MEMBER_ID_INVALID",
                    result);
                if (string.IsNullOrWhiteSpace(memberId))
                {
                    continue;
                }

                if (!unique.Add(memberId))
                {
                    result.AddError(
                        "PIVOT_CALC_SCENARIO_MEMBER_DUPLICATE",
                        memberPath,
                        "Scenario coverage members must be distinct.");
                }

                if (scenarioField != null && !scenarioField.Field.Members.Any(candidate =>
                        candidate != null &&
                        string.Equals(candidate.Id, memberId, StringComparison.OrdinalIgnoreCase)))
                {
                    result.AddError(
                        "PIVOT_CALC_SCENARIO_MEMBER_UNKNOWN",
                        memberPath,
                        "The scenario member is not present in the model schema.");
                }
            }
        }

        private static void ValidateSliceScenario(
            PivotPeriodSlice slice,
            string path,
            PivotPeriodSource source,
            PivotBoundField? scenarioField,
            ValidationResult result)
        {
            if (source.ScenarioFieldId == null)
            {
                if (slice.ScenarioMemberId != null)
                {
                    result.AddError(
                        "PIVOT_CALC_PERIOD_SLICE_SCENARIO_UNBOUND",
                        path + ".scenarioMemberId",
                        "A scenario slice requires a bound scenario field.");
                }

                return;
            }

            if (string.IsNullOrWhiteSpace(slice.ScenarioMemberId))
            {
                result.AddError(
                    "PIVOT_CALC_PERIOD_SLICE_SCENARIO_REQUIRED",
                    path + ".scenarioMemberId",
                    "Every slice over a scenario source must name one exact scenario member.");
                return;
            }

            string scenarioMemberId = slice.ScenarioMemberId!;
            ValidateId(
                scenarioMemberId,
                path + ".scenarioMemberId",
                "PIVOT_CALC_SCENARIO_MEMBER_ID_INVALID",
                result);
            if (scenarioField != null && !scenarioField.Field.Members.Any(candidate =>
                    candidate != null &&
                    string.Equals(candidate.Id, scenarioMemberId, StringComparison.OrdinalIgnoreCase)))
            {
                result.AddError(
                    "PIVOT_CALC_SCENARIO_MEMBER_UNKNOWN",
                    path + ".scenarioMemberId",
                    "The requested scenario member is not present in the model schema.");
            }
        }

        private static void ValidateSliceCoverage(
            PivotPeriodDefinition periods,
            PivotPeriodSlice slice,
            string path,
            ValidationResult result)
        {
            if (periods.Source.DateCoverageMode == PivotDateCoverageMode.ContinuousRange)
            {
                if (!PivotPeriodRules.TryGetDateRange(slice.Point, out DateTime start, out DateTime end) ||
                    !periods.Source.ContinuousRangeStart.HasValue ||
                    !periods.Source.ContinuousRangeEnd.HasValue ||
                    start < periods.Source.ContinuousRangeStart.Value ||
                    end > periods.Source.ContinuousRangeEnd.Value)
                {
                    result.AddError(
                        "PIVOT_CALC_PERIOD_SLICE_COVERAGE_MISSING",
                        path + ".point",
                        "The requested slice is outside the declared continuous date coverage.");
                }

                if (!string.IsNullOrWhiteSpace(slice.ScenarioMemberId) &&
                    !periods.Source.ContinuousRangeScenarioMemberIds.Contains(
                        slice.ScenarioMemberId,
                        StringComparer.OrdinalIgnoreCase))
                {
                    result.AddError(
                        "PIVOT_CALC_PERIOD_SCENARIO_COVERAGE_MISSING",
                        path + ".scenarioMemberId",
                        "The requested scenario is not covered by the continuous date range.");
                }

                return;
            }

            List<PivotPeriodCoverageMember> coverage =
                PivotPeriodRules.ResolveCoverage(periods, slice).ToList();
            int expected = PivotPeriodRules.ExpectedBucketCount(
                slice.Point,
                periods.Source.SourceGrain);
            if (expected > 0 && coverage.Count != expected)
            {
                result.AddError(
                    "PIVOT_CALC_PERIOD_SLICE_COVERAGE_MISSING",
                    path + ".point",
                    "The requested slice is not fully represented by the declared source coverage.");
            }

            if (!string.IsNullOrWhiteSpace(slice.ScenarioMemberId) &&
                coverage.Any(member => !member.ScenarioMemberIds.Contains(
                    slice.ScenarioMemberId,
                    StringComparer.OrdinalIgnoreCase)))
            {
                result.AddError(
                    "PIVOT_CALC_PERIOD_SCENARIO_COVERAGE_MISSING",
                    path + ".scenarioMemberId",
                    "The requested scenario is not covered for every source period bucket.");
            }
        }

        private static void ValidateMeasures(
            PivotMeasureSetDefinition definition,
            PivotCalculationModelIndex model,
            IReadOnlyDictionary<string, PivotPeriodSlice> slices,
            ValidationResult result)
        {
            if (definition.Measures.Count == 0)
            {
                result.AddError(
                    "PIVOT_CALC_MEASURE_REQUIRED",
                    "measures",
                    "At least one typed measure definition is required.");
                return;
            }

            if (definition.Measures.Count > MaximumMeasures)
            {
                result.AddError(
                    "PIVOT_CALC_MEASURE_LIMIT",
                    "measures",
                    "The calculation graph exceeds the bounded measure limit.");
            }

            var measures = new Dictionary<string, PivotMeasureDefinition>(StringComparer.OrdinalIgnoreCase);
            var measureIndexes = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var captions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var index = 0; index < definition.Measures.Count; index++)
            {
                PivotMeasureDefinition? measure = definition.Measures[index];
                string path = "measures[" + index.ToString(CultureInfo.InvariantCulture) + "]";
                if (measure == null)
                {
                    result.AddError(
                        "PIVOT_CALC_MEASURE_NULL",
                        path,
                        "Measure definitions cannot be null.");
                    continue;
                }

                ValidateId(measure.Id, path + ".id", "PIVOT_CALC_MEASURE_ID_INVALID", result);
                if (measures.ContainsKey(measure.Id))
                {
                    result.AddError(
                        "PIVOT_CALC_MEASURE_ID_DUPLICATE",
                        path + ".id",
                        "Measure IDs must be unique without regard to case.");
                }
                else
                {
                    measures.Add(measure.Id, measure);
                    measureIndexes.Add(measure.Id, index);
                }

                if (!IsBoundedText(measure.Caption, MaximumCaptionLength, allowEmpty: false) ||
                    !PivotPlusPathPolicy.IsPathFree(measure.Caption))
                {
                    result.AddError(
                        "PIVOT_CALC_MEASURE_CAPTION_INVALID",
                        path + ".caption",
                        "Native measure names must be bounded printable text.");
                }
                else if (!captions.Add(measure.Caption))
                {
                    result.AddError(
                        "PIVOT_CALC_MEASURE_CAPTION_DUPLICATE",
                        path + ".caption",
                        "Native measure names must be globally unique without regard to case.");
                }

                if (!model.TryGetTable(measure.HomeTableId, out _))
                {
                    result.AddError(
                        "PIVOT_CALC_MEASURE_HOME_TABLE_UNKNOWN",
                        path + ".homeTableId",
                        "The measure home table is not present in the model schema.");
                }

                ValidateFormat(measure.Format, path + ".format", result);
            }

            var dependencies = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
            for (var index = 0; index < definition.Measures.Count; index++)
            {
                PivotMeasureDefinition? measure = definition.Measures[index];
                if (measure == null)
                {
                    continue;
                }

                string path = "measures[" + index.ToString(CultureInfo.InvariantCulture) + "].expression";
                var state = new ExpressionValidationState(measure.Id);
                PivotCalculationSemanticKind kind = ValidateExpression(
                    measure.Expression,
                    path,
                    1,
                    state,
                    model,
                    measures,
                    slices,
                    definition.Periods,
                    result);
                dependencies[measure.Id] = state.Dependencies;

                if (state.NodeCount > MaximumExpressionNodes)
                {
                    result.AddError(
                        "PIVOT_CALC_EXPRESSION_NODE_LIMIT",
                        path,
                        "A measure expression exceeds the bounded node limit.");
                }

                if (state.Dependencies.Count > MaximumDependenciesPerMeasure)
                {
                    result.AddError(
                        "PIVOT_CALC_DEPENDENCY_LIMIT",
                        path,
                        "A measure exceeds the bounded direct dependency limit.");
                }

                ValidateOutputKind(kind, measure.Format.Kind, path, result);
            }

            ValidateDependencyGraph(measureIndexes, dependencies, result);
        }

        private static PivotCalculationSemanticKind ValidateExpression(
            PivotCalculationExpression expression,
            string path,
            int depth,
            ExpressionValidationState state,
            PivotCalculationModelIndex model,
            IReadOnlyDictionary<string, PivotMeasureDefinition> measures,
            IReadOnlyDictionary<string, PivotPeriodSlice> slices,
            PivotPeriodDefinition? periods,
            ValidationResult result)
        {
            state.NodeCount++;
            if (depth > MaximumExpressionDepth)
            {
                result.AddError(
                    "PIVOT_CALC_EXPRESSION_DEPTH_LIMIT",
                    path,
                    "A measure expression exceeds the bounded nesting depth.");
                return PivotCalculationSemanticKind.Unknown;
            }

            switch (expression)
            {
                case PivotAggregateExpression aggregate:
                    ValidateAggregate(
                        aggregate.FieldId,
                        aggregate.Function,
                        aggregate.PeriodSliceId,
                        Array.Empty<PivotCalculationFilter>(),
                        path,
                        model,
                        slices,
                        periods,
                        result);
                    return PivotCalculationSemanticKind.Numeric;

                case PivotFilteredAggregateExpression aggregate:
                    if (aggregate.Filters.Count == 0)
                    {
                        result.AddError(
                            "PIVOT_CALC_FILTERED_AGGREGATE_FILTER_REQUIRED",
                            path + ".filters",
                            "Use Aggregate when no calculation filters are required.");
                    }

                    ValidateAggregate(
                        aggregate.FieldId,
                        aggregate.Function,
                        aggregate.PeriodSliceId,
                        aggregate.Filters,
                        path,
                        model,
                        slices,
                        periods,
                        result);
                    return PivotCalculationSemanticKind.Numeric;

                case PivotWeightedResultExpression weighted:
                    ValidateWeighted(weighted, path, model, slices, periods, result);
                    return PivotCalculationSemanticKind.Numeric;

                case PivotMeasureReferenceExpression reference:
                    ValidateId(
                        reference.MeasureId,
                        path + ".measureId",
                        "PIVOT_CALC_REFERENCE_ID_INVALID",
                        result);
                    if (!measures.TryGetValue(reference.MeasureId, out PivotMeasureDefinition? referenced))
                    {
                        result.AddError(
                            "PIVOT_CALC_REFERENCE_UNKNOWN",
                            path + ".measureId",
                            "The referenced measure is not defined in this bounded graph.");
                        return PivotCalculationSemanticKind.Unknown;
                    }

                    state.Dependencies.Add(referenced.Id);
                    return SemanticFromFormat(referenced.Format.Kind);

                case PivotDifferenceExpression difference:
                {
                    PivotCalculationSemanticKind left = ValidateExpression(
                        difference.Left, path + ".left", depth + 1, state,
                        model, measures, slices, periods, result);
                    PivotCalculationSemanticKind right = ValidateExpression(
                        difference.Right, path + ".right", depth + 1, state,
                        model, measures, slices, periods, result);
                    RequireNumericPair(left, right, path, "difference", result);
                    return PivotCalculationSemanticKind.Numeric;
                }

                case PivotSafeRatioExpression ratio:
                {
                    ValidateDenominatorBehavior(ratio.OnZero, path + ".onZero", result);
                    PivotCalculationSemanticKind numerator = ValidateExpression(
                        ratio.Numerator, path + ".numerator", depth + 1, state,
                        model, measures, slices, periods, result);
                    PivotCalculationSemanticKind denominator = ValidateExpression(
                        ratio.Denominator, path + ".denominator", depth + 1, state,
                        model, measures, slices, periods, result);
                    RequireRatioOperands(numerator, denominator, path, result);
                    return PivotCalculationSemanticKind.Ratio;
                }

                case PivotShareExpression share:
                {
                    ValidateDenominatorBehavior(share.OnZero, path + ".onZero", result);
                    PivotCalculationSemanticKind part = ValidateExpression(
                        share.Part, path + ".part", depth + 1, state,
                        model, measures, slices, periods, result);
                    if (part != PivotCalculationSemanticKind.Numeric &&
                        part != PivotCalculationSemanticKind.Unknown)
                    {
                        result.AddError(
                            "PIVOT_CALC_SHARE_PART_TYPE_INVALID",
                            path + ".part",
                            "Share requires a numeric part expression.");
                    }

                    ValidateShareDenominator(
                        share,
                        path + ".denominator",
                        depth,
                        state,
                        model,
                        measures,
                        slices,
                        periods,
                        result);
                    return PivotCalculationSemanticKind.Ratio;
                }

                case PivotGrowthExpression growth:
                {
                    ValidateDenominatorBehavior(growth.OnZero, path + ".onZero", result);
                    PivotCalculationSemanticKind current = ValidateExpression(
                        growth.Current, path + ".current", depth + 1, state,
                        model, measures, slices, periods, result);
                    PivotCalculationSemanticKind prior = ValidateExpression(
                        growth.Prior, path + ".prior", depth + 1, state,
                        model, measures, slices, periods, result);
                    RequireNumericPair(current, prior, path, "growth", result);
                    return PivotCalculationSemanticKind.Ratio;
                }

                case PivotAchievementExpression achievement:
                {
                    ValidateDenominatorBehavior(achievement.OnZero, path + ".onZero", result);
                    PivotCalculationSemanticKind actual = ValidateExpression(
                        achievement.Actual, path + ".actual", depth + 1, state,
                        model, measures, slices, periods, result);
                    PivotCalculationSemanticKind target = ValidateExpression(
                        achievement.Target, path + ".target", depth + 1, state,
                        model, measures, slices, periods, result);
                    RequireNumericPair(actual, target, path, "achievement", result);
                    return PivotCalculationSemanticKind.Ratio;
                }

                case PivotVarianceExpression variance:
                {
                    ValidateVarianceConvention(variance.Convention, path + ".convention", result);
                    PivotCalculationSemanticKind actual = ValidateExpression(
                        variance.Actual, path + ".actual", depth + 1, state,
                        model, measures, slices, periods, result);
                    PivotCalculationSemanticKind plan = ValidateExpression(
                        variance.Plan, path + ".plan", depth + 1, state,
                        model, measures, slices, periods, result);
                    RequireNumericPair(actual, plan, path, "variance", result);
                    return PivotCalculationSemanticKind.Numeric;
                }

                case PivotVariancePercentageExpression variancePercentage:
                {
                    ValidateVarianceConvention(
                        variancePercentage.Convention,
                        path + ".convention",
                        result);
                    ValidateDenominatorBehavior(
                        variancePercentage.OnZero,
                        path + ".onZero",
                        result);
                    PivotCalculationSemanticKind actual = ValidateExpression(
                        variancePercentage.Actual, path + ".actual", depth + 1, state,
                        model, measures, slices, periods, result);
                    PivotCalculationSemanticKind plan = ValidateExpression(
                        variancePercentage.Plan, path + ".plan", depth + 1, state,
                        model, measures, slices, periods, result);
                    RequireNumericPair(actual, plan, path, "variance percentage", result);
                    return PivotCalculationSemanticKind.Ratio;
                }

                case PivotPercentagePointDeltaExpression delta:
                {
                    PivotCalculationSemanticKind current = ValidateExpression(
                        delta.CurrentRatio, path + ".currentRatio", depth + 1, state,
                        model, measures, slices, periods, result);
                    PivotCalculationSemanticKind baseline = ValidateExpression(
                        delta.BaselineRatio, path + ".baselineRatio", depth + 1, state,
                        model, measures, slices, periods, result);
                    if ((current != PivotCalculationSemanticKind.Ratio &&
                         current != PivotCalculationSemanticKind.Unknown) ||
                        (baseline != PivotCalculationSemanticKind.Ratio &&
                         baseline != PivotCalculationSemanticKind.Unknown))
                    {
                        result.AddError(
                            "PIVOT_CALC_PERCENTAGE_POINT_OPERAND_INVALID",
                            path,
                            "Percentage-point delta requires two percentage ratio expressions.");
                    }

                    return PivotCalculationSemanticKind.PercentagePoints;
                }

                default:
                    result.AddError(
                        "PIVOT_CALC_EXPRESSION_KIND_UNSUPPORTED",
                        path,
                        "The calculation expression kind is not supported.");
                    return PivotCalculationSemanticKind.Unknown;
            }
        }

        private static void ValidateAggregate(
            string fieldId,
            PivotCalculationAggregateFunction function,
            string? periodSliceId,
            IReadOnlyList<PivotCalculationFilter> filters,
            string path,
            PivotCalculationModelIndex model,
            IReadOnlyDictionary<string, PivotPeriodSlice> slices,
            PivotPeriodDefinition? periods,
            ValidationResult result)
        {
            if (!model.TryGetField(fieldId, out PivotBoundField field))
            {
                result.AddError(
                    "PIVOT_CALC_AGGREGATE_FIELD_UNKNOWN",
                    path + ".fieldId",
                    "The aggregate field is not present in the model schema.");
            }

            if (!Enum.IsDefined(typeof(PivotCalculationAggregateFunction), function) ||
                function == PivotCalculationAggregateFunction.Unknown)
            {
                result.AddError(
                    "PIVOT_CALC_AGGREGATE_FUNCTION_INVALID",
                    path + ".function",
                    "The aggregate function is not supported.");
            }
            else if (field != null && RequiresNumericField(function) &&
                     !IsNumeric(field.Field.DataType))
            {
                result.AddError(
                    "PIVOT_CALC_AGGREGATE_FIELD_TYPE_INVALID",
                    path + ".fieldId",
                    "This aggregate function requires a numeric model field.");
            }

            ValidateFilters(filters, path + ".filters", model, result);
            ValidatePeriodReference(periodSliceId, filters, path, slices, periods, result);
        }

        private static void ValidateWeighted(
            PivotWeightedResultExpression weighted,
            string path,
            PivotCalculationModelIndex model,
            IReadOnlyDictionary<string, PivotPeriodSlice> slices,
            PivotPeriodDefinition? periods,
            ValidationResult result)
        {
            PivotBoundField? value = null;
            PivotBoundField? weight = null;
            if (!model.TryGetField(weighted.ValueFieldId, out PivotBoundField valueField))
            {
                result.AddError(
                    "PIVOT_CALC_WEIGHTED_VALUE_FIELD_UNKNOWN",
                    path + ".valueFieldId",
                    "The weighted value field is not present in the model schema.");
            }
            else
            {
                value = valueField;
                if (!IsNumeric(value.Field.DataType))
                {
                    result.AddError(
                        "PIVOT_CALC_WEIGHTED_VALUE_FIELD_TYPE_INVALID",
                        path + ".valueFieldId",
                        "Weighted values require a numeric field.");
                }
            }

            if (!model.TryGetField(weighted.WeightFieldId, out PivotBoundField weightField))
            {
                result.AddError(
                    "PIVOT_CALC_WEIGHTED_WEIGHT_FIELD_UNKNOWN",
                    path + ".weightFieldId",
                    "The weight field is not present in the model schema.");
            }
            else
            {
                weight = weightField;
                if (!IsNumeric(weight.Field.DataType))
                {
                    result.AddError(
                        "PIVOT_CALC_WEIGHTED_WEIGHT_FIELD_TYPE_INVALID",
                        path + ".weightFieldId",
                        "Weights require a numeric field.");
                }
            }

            if (value != null && weight != null &&
                !string.Equals(value.Table.Id, weight.Table.Id, StringComparison.OrdinalIgnoreCase))
            {
                result.AddError(
                    "PIVOT_CALC_WEIGHTED_TABLE_MISMATCH",
                    path,
                    "Weighted value and weight fields must belong to the same model table.");
            }

            ValidateDenominatorBehavior(weighted.OnZero, path + ".onZero", result);
            ValidateFilters(weighted.Filters, path + ".filters", model, result);
            ValidatePeriodReference(
                weighted.PeriodSliceId,
                weighted.Filters,
                path,
                slices,
                periods,
                result);
        }

        private static void ValidateShareDenominator(
            PivotShareExpression share,
            string path,
            int depth,
            ExpressionValidationState state,
            PivotCalculationModelIndex model,
            IReadOnlyDictionary<string, PivotMeasureDefinition> measures,
            IReadOnlyDictionary<string, PivotPeriodSlice> slices,
            PivotPeriodDefinition? periods,
            ValidationResult result)
        {
            switch (share.Denominator)
            {
                case PivotExplicitShareDenominator explicitDenominator:
                {
                    PivotCalculationSemanticKind denominator = ValidateExpression(
                        explicitDenominator.Expression,
                        path + ".expression",
                        depth + 1,
                        state,
                        model,
                        measures,
                        slices,
                        periods,
                        result);
                    if (denominator != PivotCalculationSemanticKind.Numeric &&
                        denominator != PivotCalculationSemanticKind.Unknown)
                    {
                        result.AddError(
                            "PIVOT_CALC_SHARE_DENOMINATOR_TYPE_INVALID",
                            path,
                            "Explicit share denominator must be numeric.");
                    }

                    break;
                }

                case PivotParentShareDenominator parent:
                    ValidateShareFields(
                        parent.ClearedFieldIds,
                        path + ".clearedFieldIds",
                        model,
                        result);
                    break;

                case PivotFilteredTotalShareDenominator filteredTotal:
                    ValidateShareFields(
                        filteredTotal.ClearedFieldIds,
                        path + ".clearedFieldIds",
                        model,
                        result);
                    break;

                default:
                    result.AddError(
                        "PIVOT_CALC_SHARE_DENOMINATOR_INVALID",
                        path,
                        "Share denominator scope is not supported.");
                    break;
            }
        }

        private static void ValidateShareFields(
            IReadOnlyList<string> fieldIds,
            string path,
            PivotCalculationModelIndex model,
            ValidationResult result)
        {
            if (fieldIds.Count == 0)
            {
                result.AddError(
                    "PIVOT_CALC_SHARE_SCOPE_FIELD_REQUIRED",
                    path,
                    "Parent and filtered-total shares require exact fields to clear.");
                return;
            }

            if (fieldIds.Count > MaximumContextFields)
            {
                result.AddError(
                    "PIVOT_CALC_SHARE_SCOPE_FIELD_LIMIT",
                    path,
                    "The share scope exceeds the bounded field limit.");
            }

            var unique = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var index = 0; index < fieldIds.Count; index++)
            {
                string fieldId = fieldIds[index];
                string fieldPath = path + "[" + index.ToString(CultureInfo.InvariantCulture) + "]";
                ValidateId(
                    fieldId,
                    fieldPath,
                    "PIVOT_CALC_SHARE_SCOPE_FIELD_ID_INVALID",
                    result);
                if (string.IsNullOrWhiteSpace(fieldId))
                {
                    continue;
                }

                if (!unique.Add(fieldId))
                {
                    result.AddError(
                        "PIVOT_CALC_SHARE_SCOPE_FIELD_DUPLICATE",
                        fieldPath,
                        "Share scope fields must be distinct.");
                }

                if (!model.TryGetField(fieldId, out PivotBoundField field))
                {
                    result.AddError(
                        "PIVOT_CALC_SHARE_SCOPE_FIELD_UNKNOWN",
                        fieldPath,
                        "A share scope field is not present in the model schema.");
                    continue;
                }

            }
        }

        private static void ValidateFilters(
            IReadOnlyList<PivotCalculationFilter> filters,
            string path,
            PivotCalculationModelIndex model,
            ValidationResult result)
        {
            if (filters.Count > MaximumFilters)
            {
                result.AddError(
                    "PIVOT_CALC_FILTER_LIMIT",
                    path,
                    "The expression exceeds the bounded calculation-filter limit.");
            }

            var fields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var index = 0; index < filters.Count; index++)
            {
                PivotCalculationFilter? filter = filters[index];
                string filterPath = path + "[" + index.ToString(CultureInfo.InvariantCulture) + "]";
                if (filter == null)
                {
                    result.AddError(
                        "PIVOT_CALC_FILTER_NULL",
                        filterPath,
                        "Calculation filter entries cannot be null.");
                    continue;
                }

                if (!fields.Add(filter.FieldId))
                {
                    result.AddError(
                        "PIVOT_CALC_FILTER_FIELD_DUPLICATE",
                        filterPath + ".fieldId",
                        "An expression can declare only one filter per model field.");
                }

                if (!model.TryGetField(filter.FieldId, out PivotBoundField field))
                {
                    result.AddError(
                        "PIVOT_CALC_FILTER_FIELD_UNKNOWN",
                        filterPath + ".fieldId",
                        "The calculation filter field is not present in the model schema.");
                    continue;
                }

                if (!Enum.IsDefined(typeof(PivotCalculationFilterOperator), filter.Operator) ||
                    filter.Operator == PivotCalculationFilterOperator.Unknown)
                {
                    result.AddError(
                        "PIVOT_CALC_FILTER_OPERATOR_INVALID",
                        filterPath + ".operator",
                        "The calculation filter operator is not supported.");
                }

                int expected = FilterArity(filter.Operator);
                if (expected >= 0 && filter.Values.Count != expected)
                {
                    result.AddError(
                        "PIVOT_CALC_FILTER_ARITY_INVALID",
                        filterPath + ".values",
                        "The calculation filter has the wrong number of typed values.");
                }
                else if ((filter.Operator == PivotCalculationFilterOperator.In ||
                          filter.Operator == PivotCalculationFilterOperator.NotIn) &&
                         filter.Values.Count == 0)
                {
                    result.AddError(
                        "PIVOT_CALC_FILTER_VALUE_REQUIRED",
                        filterPath + ".values",
                        "IN and NOT IN filters require at least one typed value.");
                }

                if (filter.Values.Count > MaximumFilterValues)
                {
                    result.AddError(
                        "PIVOT_CALC_FILTER_VALUE_LIMIT",
                        filterPath + ".values",
                        "The calculation filter exceeds the bounded value limit.");
                }

                if (IsComparison(filter.Operator) && !SupportsComparison(field.Field.DataType))
                {
                    result.AddError(
                        "PIVOT_CALC_FILTER_COMPARISON_TYPE_INVALID",
                        filterPath + ".fieldId",
                        "Ordered comparisons require a numeric or temporal field.");
                }

                var values = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                for (var valueIndex = 0; valueIndex < filter.Values.Count; valueIndex++)
                {
                    PivotFilterValue? value = filter.Values[valueIndex];
                    string valuePath = filterPath + ".values[" +
                        valueIndex.ToString(CultureInfo.InvariantCulture) + "]";
                    if (value == null)
                    {
                        result.AddError(
                            "PIVOT_CALC_FILTER_VALUE_NULL",
                            valuePath,
                            "Calculation filter values cannot be null.");
                        continue;
                    }

                    PivotScalarValue? scalar = ValidateFilterValue(
                        value,
                        field,
                        valuePath,
                        model,
                        result);
                    if (scalar == null)
                    {
                        continue;
                    }

                    if (scalar.Kind == PivotScalarKind.Blank &&
                        IsComparison(filter.Operator))
                    {
                        result.AddError(
                            "PIVOT_CALC_FILTER_BLANK_COMPARISON_INVALID",
                            valuePath,
                            "Blank cannot be used in an ordered comparison.");
                    }

                    if (!values.Add(PivotCalculationCanonical.ScalarKey(scalar)))
                    {
                        result.AddError(
                            "PIVOT_CALC_FILTER_VALUE_DUPLICATE",
                            valuePath,
                            "Calculation filter values must resolve to a distinct semantic set.");
                    }
                }
            }
        }

        private static PivotScalarValue? ValidateFilterValue(
            PivotFilterValue value,
            PivotBoundField field,
            string path,
            PivotCalculationModelIndex model,
            ValidationResult result)
        {
            if (!Enum.IsDefined(typeof(PivotFilterValueKind), value.Kind))
            {
                result.AddError(
                    "PIVOT_CALC_FILTER_VALUE_KIND_INVALID",
                    path,
                    "The filter value kind is not supported.");
                return null;
            }

            if (!model.TryResolveValue(field.Field.Id, value, out PivotScalarValue scalar))
            {
                result.AddError(
                    value.Kind == PivotFilterValueKind.Member
                        ? "PIVOT_CALC_FILTER_MEMBER_UNKNOWN"
                        : "PIVOT_CALC_FILTER_SCALAR_INVALID",
                    path,
                    value.Kind == PivotFilterValueKind.Member
                        ? "The filter member is not present in the bound field schema."
                        : "The typed scalar filter value is invalid.");
                return null;
            }

            if (!IsScalarCompatible(field.Field.DataType, scalar))
            {
                result.AddError(
                    "PIVOT_CALC_FILTER_VALUE_TYPE_MISMATCH",
                    path,
                    "The typed filter value does not match its bound model field.");
            }

            ValidateScalarShape(scalar, path, result);

            return scalar;
        }

        private static void ValidateScalarShape(
            PivotScalarValue value,
            string path,
            ValidationResult result)
        {
            if (!Enum.IsDefined(typeof(PivotScalarKind), value.Kind))
            {
                result.AddError(
                    "PIVOT_CALC_SCALAR_KIND_INVALID",
                    path,
                    "The typed scalar kind is not supported.");
                return;
            }

            if (value.Kind == PivotScalarKind.Text &&
                !IsBoundedText(value.TextValue ?? string.Empty, 32767, allowEmpty: true))
            {
                result.AddError(
                    "PIVOT_CALC_SCALAR_TEXT_INVALID",
                    path,
                    "Text DAX literals must be bounded and cannot contain control characters.");
            }

            if ((value.Kind == PivotScalarKind.Date ||
                 value.Kind == PivotScalarKind.DateTime) &&
                (!value.TemporalValue.HasValue ||
                 value.TemporalValue.Value.Year < 1900 ||
                 value.TemporalValue.Value.Year > 9999))
            {
                result.AddError(
                    "PIVOT_CALC_SCALAR_DATE_RANGE_INVALID",
                    path,
                    "Temporal DAX literals require a year from 1900 through 9999.");
            }
        }

        private static void ValidatePeriodReference(
            string? periodSliceId,
            IReadOnlyList<PivotCalculationFilter> filters,
            string path,
            IReadOnlyDictionary<string, PivotPeriodSlice> slices,
            PivotPeriodDefinition? periods,
            ValidationResult result)
        {
            if (periodSliceId == null)
            {
                return;
            }

            if (!slices.ContainsKey(periodSliceId))
            {
                result.AddError(
                    "PIVOT_CALC_PERIOD_SLICE_UNKNOWN",
                    path + ".periodSliceId",
                    "The expression references an unknown period slice.");
                return;
            }

            if (periods == null)
            {
                return;
            }

            var reserved = new HashSet<string>(
                periods.Source.PeriodContextFieldIds,
                StringComparer.OrdinalIgnoreCase);
            reserved.UnionWith(periods.Source.ScenarioContextFieldIds);
            foreach (PivotCalculationFilter filter in filters.Where(filter => filter != null))
            {
                if (reserved.Contains(filter.FieldId))
                {
                    result.AddError(
                        "PIVOT_CALC_PERIOD_FILTER_CONTEXT_CONFLICT",
                        path + ".filters",
                        "Explicit filters cannot also target a period or scenario slice context field.");
                }
            }
        }

        private static void ValidateFormat(
            PivotMeasureFormat format,
            string path,
            ValidationResult result)
        {
            if (!Enum.IsDefined(typeof(PivotMeasureFormatKind), format.Kind) ||
                format.Kind == PivotMeasureFormatKind.Unknown)
            {
                result.AddError(
                    "PIVOT_CALC_FORMAT_KIND_INVALID",
                    path + ".kind",
                    "The measure format kind is not supported.");
            }

            if (format.DecimalPlaces < 0 || format.DecimalPlaces > 6)
            {
                result.AddError(
                    "PIVOT_CALC_FORMAT_DECIMALS_INVALID",
                    path + ".decimalPlaces",
                    "Decimal places must be between zero and six.");
            }

            if (format.Kind == PivotMeasureFormatKind.WholeNumber && format.DecimalPlaces != 0)
            {
                result.AddError(
                    "PIVOT_CALC_FORMAT_WHOLE_DECIMALS_INVALID",
                    path + ".decimalPlaces",
                    "Whole-number format cannot declare decimal places.");
            }

            if ((format.Kind == PivotMeasureFormatKind.Percentage ||
                 format.Kind == PivotMeasureFormatKind.PercentagePoints) &&
                format.UseThousandsSeparator)
            {
                result.AddError(
                    "PIVOT_CALC_FORMAT_PERCENT_SEPARATOR_INVALID",
                    path + ".useThousandsSeparator",
                    "Percentage formats do not use a thousands separator.");
            }

            if (format.Kind == PivotMeasureFormatKind.Currency)
            {
                if (!IsCurrencyMarker(format.CurrencySymbolOrCode))
                {
                    result.AddError(
                        "PIVOT_CALC_FORMAT_CURRENCY_INVALID",
                        path + ".currencySymbolOrCode",
                        "Currency format requires a bounded currency symbol or alphabetic code.");
                }
            }
            else if (format.CurrencySymbolOrCode != null)
            {
                result.AddError(
                    "PIVOT_CALC_FORMAT_CURRENCY_UNEXPECTED",
                    path + ".currencySymbolOrCode",
                    "Only Currency format can declare a currency symbol or code.");
            }
        }

        private static void ValidateOutputKind(
            PivotCalculationSemanticKind expressionKind,
            PivotMeasureFormatKind formatKind,
            string path,
            ValidationResult result)
        {
            bool valid;
            switch (expressionKind)
            {
                case PivotCalculationSemanticKind.Numeric:
                    valid = formatKind == PivotMeasureFormatKind.WholeNumber ||
                            formatKind == PivotMeasureFormatKind.DecimalNumber ||
                            formatKind == PivotMeasureFormatKind.Currency;
                    break;
                case PivotCalculationSemanticKind.Ratio:
                    valid = formatKind == PivotMeasureFormatKind.Percentage;
                    break;
                case PivotCalculationSemanticKind.PercentagePoints:
                    valid = formatKind == PivotMeasureFormatKind.PercentagePoints;
                    break;
                default:
                    return;
            }

            if (!valid)
            {
                result.AddError(
                    "PIVOT_CALC_FORMAT_SEMANTIC_MISMATCH",
                    path,
                    "The typed measure format does not match the expression's semantic result.");
            }
        }

        private static void ValidateDependencyGraph(
            IReadOnlyDictionary<string, int> measureIndexes,
            IReadOnlyDictionary<string, HashSet<string>> dependencies,
            ValidationResult result)
        {
            var states = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var cycleReported = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string measureId in measureIndexes.Keys.OrderBy(id => measureIndexes[id]))
            {
                Visit(measureId, 1);
            }

            void Visit(string measureId, int depth)
            {
                if (depth > MaximumExpressionDepth)
                {
                    result.AddError(
                        "PIVOT_CALC_DEPENDENCY_DEPTH_LIMIT",
                        "measures[" + measureIndexes[measureId].ToString(CultureInfo.InvariantCulture) + "]",
                        "The measure dependency graph exceeds the bounded depth limit.");
                    return;
                }

                if (states.TryGetValue(measureId, out int state))
                {
                    if (state == 1 && cycleReported.Add(measureId))
                    {
                        result.AddError(
                            "PIVOT_CALC_REFERENCE_CYCLE",
                            "measures[" + measureIndexes[measureId].ToString(CultureInfo.InvariantCulture) + "]",
                            "The measure dependency graph contains a cycle.");
                    }

                    return;
                }

                states[measureId] = 1;
                if (dependencies.TryGetValue(measureId, out HashSet<string>? direct))
                {
                    foreach (string dependency in direct
                        .Where(measureIndexes.ContainsKey)
                        .OrderBy(id => measureIndexes[id]))
                    {
                        Visit(dependency, depth + 1);
                    }
                }

                states[measureId] = 2;
            }
        }

        private static void ValidatePeriodPoint(
            PivotPeriodPoint point,
            string path,
            ValidationResult result)
        {
            if (!Enum.IsDefined(typeof(PivotPeriodGrain), point.Grain) ||
                point.Grain == PivotPeriodGrain.Unknown)
            {
                result.AddError(
                    "PIVOT_CALC_PERIOD_POINT_GRAIN_INVALID",
                    path + ".grain",
                    "A period point requires a supported grain.");
                return;
            }

            if (point.Year < 1900 || point.Year > 9999)
            {
                result.AddError(
                    "PIVOT_CALC_PERIOD_YEAR_INVALID",
                    path + ".year",
                    "A period point requires an explicit four-digit reporting year.");
            }

            bool validShape;
            switch (point.Grain)
            {
                case PivotPeriodGrain.Year:
                    validShape = !point.Ordinal.HasValue && !point.Date.HasValue;
                    break;
                case PivotPeriodGrain.Half:
                    validShape = (point.Ordinal == 1 || point.Ordinal == 2) && !point.Date.HasValue;
                    break;
                case PivotPeriodGrain.Quarter:
                    validShape = point.Ordinal >= 1 && point.Ordinal <= 4 && !point.Date.HasValue;
                    break;
                case PivotPeriodGrain.Month:
                    validShape = point.Ordinal >= 1 && point.Ordinal <= 12 && !point.Date.HasValue;
                    break;
                case PivotPeriodGrain.Date:
                    validShape = !point.Ordinal.HasValue && point.Date.HasValue &&
                                 point.Date.Value.TimeOfDay == TimeSpan.Zero &&
                                 point.Date.Value.Year == point.Year;
                    break;
                default:
                    validShape = false;
                    break;
            }

            if (!validShape)
            {
                result.AddError(
                    "PIVOT_CALC_PERIOD_POINT_SHAPE_INVALID",
                    path,
                    "The period point does not match its declared grain.");
            }
        }

        private static void ValidateDenominatorBehavior(
            PivotDenominatorBehavior behavior,
            string path,
            ValidationResult result)
        {
            if (behavior != PivotDenominatorBehavior.Blank &&
                behavior != PivotDenominatorBehavior.Zero)
            {
                result.AddError(
                    "PIVOT_CALC_DENOMINATOR_BEHAVIOR_INVALID",
                    path,
                    "Zero denominators must explicitly return Blank or Zero.");
            }
        }

        private static void ValidateVarianceConvention(
            PivotVarianceConvention convention,
            string path,
            ValidationResult result)
        {
            if (convention != PivotVarianceConvention.ActualMinusPlan &&
                convention != PivotVarianceConvention.PlanMinusActual)
            {
                result.AddError(
                    "PIVOT_CALC_VARIANCE_CONVENTION_INVALID",
                    path,
                    "Variance must explicitly declare Actual-minus-Plan or Plan-minus-Actual.");
            }
        }

        private static void RequireNumericPair(
            PivotCalculationSemanticKind left,
            PivotCalculationSemanticKind right,
            string path,
            string operation,
            ValidationResult result)
        {
            if ((left != PivotCalculationSemanticKind.Numeric &&
                 left != PivotCalculationSemanticKind.Unknown) ||
                (right != PivotCalculationSemanticKind.Numeric &&
                 right != PivotCalculationSemanticKind.Unknown))
            {
                result.AddError(
                    "PIVOT_CALC_NUMERIC_OPERAND_REQUIRED",
                    path,
                    "The " + operation + " operation requires numeric operands.");
            }
        }

        private static void RequireRatioOperands(
            PivotCalculationSemanticKind numerator,
            PivotCalculationSemanticKind denominator,
            string path,
            ValidationResult result)
        {
            bool numeratorValid = numerator == PivotCalculationSemanticKind.Numeric ||
                                  numerator == PivotCalculationSemanticKind.Ratio ||
                                  numerator == PivotCalculationSemanticKind.Unknown;
            bool denominatorValid = denominator == PivotCalculationSemanticKind.Numeric ||
                                    denominator == PivotCalculationSemanticKind.Ratio ||
                                    denominator == PivotCalculationSemanticKind.Unknown;
            if (!numeratorValid || !denominatorValid)
            {
                result.AddError(
                    "PIVOT_CALC_RATIO_OPERAND_INVALID",
                    path,
                    "Safe ratio operands must be numeric values or ratios.");
            }
        }

        private static PivotCalculationSemanticKind SemanticFromFormat(PivotMeasureFormatKind kind)
        {
            switch (kind)
            {
                case PivotMeasureFormatKind.Percentage:
                    return PivotCalculationSemanticKind.Ratio;
                case PivotMeasureFormatKind.PercentagePoints:
                    return PivotCalculationSemanticKind.PercentagePoints;
                case PivotMeasureFormatKind.WholeNumber:
                case PivotMeasureFormatKind.DecimalNumber:
                case PivotMeasureFormatKind.Currency:
                    return PivotCalculationSemanticKind.Numeric;
                default:
                    return PivotCalculationSemanticKind.Unknown;
            }
        }

        private static bool RequiresNumericField(PivotCalculationAggregateFunction function)
        {
            return function == PivotCalculationAggregateFunction.Sum ||
                   function == PivotCalculationAggregateFunction.Average ||
                   function == PivotCalculationAggregateFunction.Minimum ||
                   function == PivotCalculationAggregateFunction.Maximum;
        }

        private static bool IsNumeric(PivotModelDataType type)
        {
            return type == PivotModelDataType.WholeNumber ||
                   type == PivotModelDataType.DecimalNumber ||
                   type == PivotModelDataType.Currency;
        }

        private static bool SupportsComparison(PivotModelDataType type)
        {
            return IsNumeric(type) || type == PivotModelDataType.Date ||
                   type == PivotModelDataType.DateTime;
        }

        private static bool SupportsPeriodValues(PivotModelDataType type)
        {
            return type == PivotModelDataType.Text ||
                   type == PivotModelDataType.WholeNumber ||
                   type == PivotModelDataType.Date ||
                   type == PivotModelDataType.DateTime;
        }

        private static bool TemporalMatchesPoint(DateTime value, PivotPeriodPoint point)
        {
            var candidate = new PivotPeriodPoint(
                PivotPeriodGrain.Date,
                value.Year,
                date: value.Date);
            return PivotPeriodRules.IsWithin(candidate, point);
        }

        private static bool IsComparison(PivotCalculationFilterOperator @operator)
        {
            return @operator == PivotCalculationFilterOperator.GreaterThan ||
                   @operator == PivotCalculationFilterOperator.GreaterThanOrEqual ||
                   @operator == PivotCalculationFilterOperator.LessThan ||
                   @operator == PivotCalculationFilterOperator.LessThanOrEqual;
        }

        private static int FilterArity(PivotCalculationFilterOperator @operator)
        {
            switch (@operator)
            {
                case PivotCalculationFilterOperator.Equal:
                case PivotCalculationFilterOperator.NotEqual:
                case PivotCalculationFilterOperator.GreaterThan:
                case PivotCalculationFilterOperator.GreaterThanOrEqual:
                case PivotCalculationFilterOperator.LessThan:
                case PivotCalculationFilterOperator.LessThanOrEqual:
                    return 1;
                case PivotCalculationFilterOperator.IsBlank:
                case PivotCalculationFilterOperator.IsNotBlank:
                    return 0;
                case PivotCalculationFilterOperator.In:
                case PivotCalculationFilterOperator.NotIn:
                    return -1;
                default:
                    return -2;
            }
        }

        private static bool IsScalarCompatible(
            PivotModelDataType type,
            PivotScalarValue value)
        {
            if (value.Kind == PivotScalarKind.Blank)
            {
                return true;
            }

            switch (type)
            {
                case PivotModelDataType.Text:
                    return value.Kind == PivotScalarKind.Text;
                case PivotModelDataType.WholeNumber:
                    return value.Kind == PivotScalarKind.WholeNumber;
                case PivotModelDataType.DecimalNumber:
                case PivotModelDataType.Currency:
                    return value.Kind == PivotScalarKind.WholeNumber ||
                           value.Kind == PivotScalarKind.DecimalNumber;
                case PivotModelDataType.Date:
                    return value.Kind == PivotScalarKind.Date;
                case PivotModelDataType.DateTime:
                    return value.Kind == PivotScalarKind.Date ||
                           value.Kind == PivotScalarKind.DateTime;
                case PivotModelDataType.Boolean:
                    return value.Kind == PivotScalarKind.Boolean;
                default:
                    return false;
            }
        }

        private static bool IsCurrencyMarker(string? value)
        {
            if (value == null ||
                string.IsNullOrWhiteSpace(value) ||
                value.Length > MaximumCurrencyMarkerLength ||
                value != value.Trim())
            {
                return false;
            }

            foreach (char character in value)
            {
                UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(character);
                if (!char.IsLetter(character) && category != UnicodeCategory.CurrencySymbol)
                {
                    return false;
                }
            }

            return true;
        }

        private static void ValidateId(
            string value,
            string path,
            string code,
            ValidationResult result)
        {
            if (!IsId(value))
            {
                result.AddError(
                    code,
                    path,
                    "The identifier must be a bounded path-free token.");
            }
        }

        private static bool IsId(string value)
        {
            if (string.IsNullOrWhiteSpace(value) ||
                value.Length > MaximumIdLength ||
                value != value.Trim() ||
                (!char.IsLetter(value[0]) && value[0] != '_'))
            {
                return false;
            }

            foreach (char character in value)
            {
                if (!char.IsLetterOrDigit(character) &&
                    character != '_' && character != '-' && character != '.')
                {
                    return false;
                }
            }

            return PivotPlusPathPolicy.IsPathFree(value);
        }

        private static void ValidateNativeName(
            string value,
            string path,
            string code,
            ValidationResult result)
        {
            if (!IsBoundedText(value, MaximumNativeNameLength, allowEmpty: false) ||
                !PivotPlusPathPolicy.IsPathFree(value))
            {
                result.AddError(
                    code,
                    path,
                    "The native model identifier must be bounded and path-free.");
            }
        }

        private static bool IsBoundedText(string value, int maximumLength, bool allowEmpty)
        {
            if (value == null || value.Length > maximumLength || value != value.Trim())
            {
                return false;
            }

            if (!allowEmpty && string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            return !value.Any(char.IsControl);
        }

        private sealed class ExpressionValidationState
        {
            public ExpressionValidationState(string measureId)
            {
                MeasureId = measureId;
                Dependencies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }

            public string MeasureId { get; }

            public int NodeCount { get; set; }

            public HashSet<string> Dependencies { get; }
        }
    }
}
