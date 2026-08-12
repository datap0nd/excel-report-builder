using System;
using System.Collections.Generic;
using System.Linq;
using ExcelReportBuilder.Core.Validation;

namespace ExcelReportBuilder.Core.PivotPlus
{
    public static class PivotPlusValidator
    {
        private const int MaximumWorkbookIdLength = 128;
        private const int MaximumWorksheetNameLength = 31;
        private const int MaximumExcelNameLength = 255;
        private const int MaximumMemberCaptionLength = 32767;
        private const int MaximumCapabilityReasonLength = 500;

        private const PivotCapability KnownCapabilities =
            PivotCapability.NativeFieldPlacement |
            PivotCapability.MemberFiltering |
            PivotCapability.LayoutFormatting |
            PivotCapability.ShowValuesAs |
            PivotCapability.DistinctCount |
            PivotCapability.DataModel |
            PivotCapability.ModelMeasures |
            PivotCapability.CalculatedMembers |
            PivotCapability.NamedSets |
            PivotCapability.AsymmetricAxes |
            PivotCapability.Refresh |
            PivotCapability.UpgradeToDataModel;

        private const PivotCapability RequiredMutationCapabilities =
            PivotCapability.NativeFieldPlacement |
            PivotCapability.LayoutFormatting |
            PivotCapability.Refresh;

        private const PivotCapability WorksheetCapabilities =
            RequiredMutationCapabilities |
            PivotCapability.MemberFiltering |
            PivotCapability.ShowValuesAs |
            PivotCapability.UpgradeToDataModel;

        private const PivotCapability DataModelCapabilities =
            RequiredMutationCapabilities |
            PivotCapability.MemberFiltering |
            PivotCapability.ShowValuesAs |
            PivotCapability.DistinctCount |
            PivotCapability.DataModel |
            PivotCapability.ModelMeasures |
            PivotCapability.CalculatedMembers |
            PivotCapability.NamedSets |
            PivotCapability.AsymmetricAxes;

        private const PivotCapability ExternalOlapCapabilities =
            RequiredMutationCapabilities |
            PivotCapability.MemberFiltering |
            PivotCapability.ShowValuesAs |
            PivotCapability.CalculatedMembers |
            PivotCapability.NamedSets |
            PivotCapability.AsymmetricAxes;

        public static ValidationResult Validate(PivotLayoutDefinition definition)
        {
            if (definition == null) throw new ArgumentNullException(nameof(definition));

            var result = new ValidationResult();
            ValidateTarget(definition.Target, result);
            ValidateSource(definition.Source, result);
            ValidateMutationIntent(definition, result);
            Dictionary<string, PivotFieldDescriptor> fields = ValidateFields(definition.Fields, result);
            ValidatePlacements(definition, fields, result);
            ValidateFilters(definition, fields, result);
            ValidateLayout(definition, result);
            ValidateFormat(definition.Format, result);
            ValidateDerivedCapabilities(definition, fields, result);
            ValidateRequirements(definition, result);
            return result;
        }

        private static void ValidateTarget(PivotTargetIdentity target, ValidationResult result)
        {
            if (!IsToken(target.WorkbookId, MaximumWorkbookIdLength))
            {
                result.AddError(
                    "PIVOT_TARGET_WORKBOOK_ID_INVALID",
                    "target.workbookId",
                    "The workbook ID must be a path-free host token containing only letters, numbers, '.', '_' or '-'.");
            }

            if (!IsExcelName(target.WorksheetName, MaximumWorksheetNameLength) ||
                target.WorksheetName.IndexOfAny(new[] { ':', '\\', '/', '?', '*', '[', ']' }) >= 0 ||
                target.WorksheetName.StartsWith("'", StringComparison.Ordinal) ||
                target.WorksheetName.EndsWith("'", StringComparison.Ordinal))
            {
                result.AddError(
                    "PIVOT_TARGET_WORKSHEET_NAME_INVALID",
                    "target.worksheetName",
                    "The worksheet name is not valid for desktop Excel.");
            }

            if (!IsExcelName(target.PivotTableName, MaximumExcelNameLength) ||
                !PivotPlusPathPolicy.IsPathFree(target.PivotTableName))
            {
                result.AddError(
                    "PIVOT_TARGET_NAME_INVALID",
                    "target.pivotTableName",
                    "The PivotTable name must be a non-path Excel object name.");
            }
        }

        private static void ValidateSource(PivotSourceDescriptor source, ValidationResult result)
        {
            if (!Enum.IsDefined(typeof(PivotSourceKind), source.Kind) || source.Kind == PivotSourceKind.Unknown)
            {
                result.AddError(
                    "PIVOT_SOURCE_KIND_UNSUPPORTED",
                    "source.kind",
                    "The selected PivotTable must expose a supported source kind before layout can be applied.");
            }

            if (!IsExcelName(source.SourceName, MaximumExcelNameLength) ||
                !PivotPlusPathPolicy.IsPathFree(source.SourceName))
            {
                result.AddError(
                    "PIVOT_SOURCE_NAME_INVALID",
                    "source.sourceName",
                    "The source must be identified by a workbook object or connection name, not a path.");
            }

            if (source.ModelTableName != null &&
                (!IsExcelName(source.ModelTableName, MaximumExcelNameLength) ||
                 !PivotPlusPathPolicy.IsPathFree(source.ModelTableName)))
            {
                result.AddError(
                    "PIVOT_SOURCE_MODEL_TABLE_NAME_INVALID",
                    "source.modelTableName",
                    "The model table name is invalid.");
            }

            if ((source.Capabilities & ~KnownCapabilities) != 0)
            {
                result.AddError(
                    "PIVOT_SOURCE_CAPABILITY_INVALID",
                    "source.capabilities",
                    "The source advertises an unknown PivotTable+ capability.");
            }

            switch (source.Kind)
            {
                case PivotSourceKind.WorksheetRange:
                case PivotSourceKind.WorksheetTable:
                    ValidateCapabilityTruthTable(
                        source.Capabilities,
                        WorksheetCapabilities,
                        "A worksheet source advertises a Data Model or OLAP-only capability.",
                        result);

                    if (!string.IsNullOrWhiteSpace(source.ModelTableName))
                    {
                        result.AddError(
                            "PIVOT_SOURCE_MODEL_TABLE_UNSUPPORTED",
                            "source.modelTableName",
                            "A worksheet source cannot identify an active model table.");
                    }

                    break;
                case PivotSourceKind.DataModel:
                    ValidateCapabilityTruthTable(
                        source.Capabilities,
                        DataModelCapabilities,
                        "A Data Model source advertises a capability reserved for classic-source upgrade.",
                        result);
                    if ((source.Capabilities & PivotCapability.DataModel) == 0)
                    {
                        result.AddError(
                            "PIVOT_SOURCE_DATA_MODEL_CAPABILITY_REQUIRED",
                            "source.capabilities",
                            "A Data Model source must advertise the DataModel capability.");
                    }

                    break;
                case PivotSourceKind.ExternalOlap:
                    ValidateCapabilityTruthTable(
                        source.Capabilities,
                        ExternalOlapCapabilities,
                        "An external OLAP source advertises a workbook Data Model or classic-upgrade capability.",
                        result);

                    break;
            }
        }

        private static void ValidateMutationIntent(
            PivotLayoutDefinition definition,
            ValidationResult result)
        {
            if (definition.Placements.Count == 0 && !definition.ClearAll)
            {
                result.AddError(
                    "PIVOT_LAYOUT_PLACEMENT_REQUIRED",
                    "placements",
                    "At least one native placement is required unless clearAll is explicitly enabled.");
            }
            else if (definition.Placements.Count > 0 && definition.ClearAll)
            {
                result.AddError(
                    "PIVOT_CLEAR_ALL_PLACEMENT_CONFLICT",
                    "clearAll",
                    "clearAll requires an empty placements collection.");
            }
        }

        private static Dictionary<string, PivotFieldDescriptor> ValidateFields(
            IReadOnlyList<PivotFieldDescriptor> fields,
            ValidationResult result)
        {
            var byName = new Dictionary<string, PivotFieldDescriptor>(StringComparer.OrdinalIgnoreCase);
            for (var index = 0; index < fields.Count; index++)
            {
                PivotFieldDescriptor field = fields[index];
                string path = "fields[" + index + "]";
                if (!IsExcelName(field.Name, MaximumExcelNameLength))
                {
                    result.AddError("PIVOT_FIELD_NAME_INVALID", path + ".name", "The PivotTable field name is invalid.");
                    continue;
                }

                if (byName.ContainsKey(field.Name))
                {
                    result.AddError(
                        "PIVOT_FIELD_DUPLICATE",
                        path + ".name",
                        "Field identifiers must be unique without regard to case.");
                }
                else
                {
                    byName.Add(field.Name, field);
                }

                ValidateOptionalExcelName(field.Caption, path + ".caption", "PIVOT_FIELD_CAPTION_INVALID", result);
                ValidateOptionalPathFreeExcelName(
                    field.TableName,
                    path + ".tableName",
                    "PIVOT_FIELD_TABLE_NAME_INVALID",
                    result);

                if (!Enum.IsDefined(typeof(PivotFieldDataType), field.DataType))
                {
                    result.AddError("PIVOT_FIELD_DATA_TYPE_INVALID", path + ".dataType", "The field data type is invalid.");
                }

                if (field.SupportedAreas == PivotFieldAreaSupport.None ||
                    (field.SupportedAreas & ~PivotFieldAreaSupport.All) != 0)
                {
                    result.AddError(
                        "PIVOT_FIELD_AREA_SUPPORT_INVALID",
                        path + ".supportedAreas",
                        "The field must advertise at least one known placement area.");
                }

                if (field.IsMeasure && field.SupportedAreas != PivotFieldAreaSupport.Values)
                {
                    result.AddError(
                        "PIVOT_MEASURE_AREA_SUPPORT_INVALID",
                        path + ".supportedAreas",
                        "A discovered measure can only be placed in Values.");
                }
            }

            return byName;
        }

        private static void ValidatePlacements(
            PivotLayoutDefinition definition,
            IReadOnlyDictionary<string, PivotFieldDescriptor> fields,
            ValidationResult result)
        {
            var placementKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var nonValueAreas = new Dictionary<string, PivotFieldArea>(StringComparer.OrdinalIgnoreCase);
            var explicitValueCaptions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var resolvedValueCaptions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var valueFieldCounts = definition.Placements
                .Where(item => item.Area == PivotFieldArea.Values)
                .GroupBy(item => item.FieldName, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Count(), StringComparer.OrdinalIgnoreCase);
            var positions = new Dictionary<PivotFieldArea, HashSet<int>>();
            foreach (PivotFieldArea area in Enum.GetValues(typeof(PivotFieldArea)))
            {
                positions[area] = new HashSet<int>();
            }

            for (var index = 0; index < definition.Placements.Count; index++)
            {
                PivotFieldPlacement placement = definition.Placements[index];
                string path = "placements[" + index + "]";
                if (!IsExcelName(placement.FieldName, MaximumExcelNameLength))
                {
                    result.AddError("PIVOT_PLACEMENT_FIELD_NAME_INVALID", path + ".fieldName", "The placement field name is invalid.");
                }

                bool areaDefined = Enum.IsDefined(typeof(PivotFieldArea), placement.Area);
                if (!areaDefined)
                {
                    result.AddError("PIVOT_PLACEMENT_AREA_INVALID", path + ".area", "The field placement area is invalid.");
                }

                string placementKey = placement.Area + "\0" + placement.FieldName;
                if (placement.Area != PivotFieldArea.Values && !placementKeys.Add(placementKey))
                {
                    result.AddError(
                        "PIVOT_PLACEMENT_DUPLICATE",
                        path,
                        "The same field cannot be placed more than once in the same area.");
                }

                if (areaDefined && placement.Area != PivotFieldArea.Values)
                {
                    if (nonValueAreas.TryGetValue(
                            placement.FieldName,
                            out PivotFieldArea existingArea) &&
                        existingArea != placement.Area)
                    {
                        result.AddError(
                            "PIVOT_NONVALUE_FIELD_MULTIPLE_AREAS",
                            path + ".area",
                            "A native PivotField can occupy only one Rows, Columns, or Filters area at a time.");
                    }
                    else
                    {
                        nonValueAreas[placement.FieldName] = placement.Area;
                    }
                }

                if (placement.Area == PivotFieldArea.Values)
                {
                    if (!string.IsNullOrWhiteSpace(placement.Caption) &&
                        !explicitValueCaptions.Add(placement.Caption!))
                    {
                        result.AddError(
                            "PIVOT_VALUE_CAPTION_DUPLICATE",
                            path + ".caption",
                            "Value instances must have unique captions without regard to case.");
                    }

                    if (valueFieldCounts.TryGetValue(placement.FieldName, out var duplicateCount) &&
                        duplicateCount > 1 &&
                        string.IsNullOrWhiteSpace(placement.Caption))
                    {
                        result.AddError(
                            "PIVOT_VALUE_INSTANCE_CAPTION_REQUIRED",
                            path + ".caption",
                            "A source field used more than once in Values requires a distinct caption for each instance.");
                    }
                }

                if (placement.Position < 1)
                {
                    result.AddError(
                        "PIVOT_PLACEMENT_POSITION_UNSUPPORTED",
                        path + ".position",
                        "PivotTable field positions are one-based.");
                }
                else if (areaDefined && !positions[placement.Area].Add(placement.Position))
                {
                    result.AddError(
                        "PIVOT_PLACEMENT_POSITION_DUPLICATE",
                        path + ".position",
                        "Only one field can occupy a position within an area.");
                }

                ValidateOptionalExcelName(placement.Caption, path + ".caption", "PIVOT_PLACEMENT_CAPTION_INVALID", result);

                PivotFieldDescriptor? field = null;
                if (IsExcelName(placement.FieldName, MaximumExcelNameLength) &&
                    !fields.TryGetValue(placement.FieldName, out field))
                {
                    result.AddError(
                        "PIVOT_PLACEMENT_FIELD_UNKNOWN",
                        path + ".fieldName",
                        "The placed field was not present in the discovery snapshot.");
                }

                if (field != null && areaDefined && !Supports(field.SupportedAreas, placement.Area))
                {
                    result.AddError(
                        "PIVOT_PLACEMENT_AREA_UNSUPPORTED",
                        path + ".area",
                        "The discovered field does not support this placement area.");
                }

                if (field != null && placement.Area == PivotFieldArea.Values)
                {
                    string resolvedCaption = PivotPlusValueSemantics.ResolveCaption(field, placement);
                    if (!resolvedValueCaptions.Add(resolvedCaption))
                    {
                        result.AddError(
                            "PIVOT_VALUE_RESOLVED_CAPTION_DUPLICATE",
                            path + ".caption",
                            "Values instances must resolve to unique native captions without regard to case.");
                    }
                }

                ValidatePlacementOptions(definition.Source, field, placement, path, result);
            }

            foreach (KeyValuePair<PivotFieldArea, HashSet<int>> area in positions)
            {
                if (area.Value.Count == 0) continue;
                int maximum = area.Value.Max();
                if (maximum != area.Value.Count)
                {
                    result.AddError(
                        "PIVOT_PLACEMENT_POSITION_GAP",
                        "placements",
                        "Positions in the " + area.Key + " area must be contiguous from 1.");
                }
            }

            ValidateValueInstanceMatrix(definition, fields, result);
        }

        private static void ValidatePlacementOptions(
            PivotSourceDescriptor source,
            PivotFieldDescriptor? field,
            PivotFieldPlacement placement,
            string path,
            ValidationResult result)
        {
            if (placement.Aggregation.HasValue &&
                !Enum.IsDefined(typeof(PivotAggregationFunction), placement.Aggregation.Value))
            {
                result.AddError(
                    "PIVOT_PLACEMENT_AGGREGATION_INVALID",
                    path + ".aggregation",
                    "The value aggregation is invalid.");
            }

            if (!Enum.IsDefined(typeof(PivotSubtotalMode), placement.SubtotalMode))
            {
                result.AddError(
                    "PIVOT_PLACEMENT_SUBTOTAL_MODE_INVALID",
                    path + ".subtotalMode",
                    "The subtotal mode is invalid.");
            }

            if (placement.Area == PivotFieldArea.Values)
            {
                if (field != null && field.IsMeasure && placement.Aggregation.HasValue)
                {
                    result.AddError(
                        "PIVOT_MEASURE_AGGREGATION_UNSUPPORTED",
                        path + ".aggregation",
                        "An existing model measure cannot be aggregated again.");
                }
                else if (field != null && !field.IsMeasure && !placement.Aggregation.HasValue)
                {
                    result.AddError(
                        "PIVOT_VALUE_AGGREGATION_REQUIRED",
                        path + ".aggregation",
                        "A source field placed in Values requires an aggregation.");
                }

                if (placement.Aggregation == PivotAggregationFunction.DistinctCount &&
                    (source.Capabilities & PivotCapability.DistinctCount) == 0)
                {
                    result.AddError(
                        "PIVOT_DISTINCT_COUNT_UNSUPPORTED",
                        path + ".aggregation",
                        "This PivotTable source does not support distinct count.");
                }

                if (field != null)
                {
                    ValidateSourceValuePlacement(source, field, placement, path, result);
                }

                if (placement.SubtotalMode != PivotSubtotalMode.None)
                {
                    result.AddError(
                        "PIVOT_VALUE_SUBTOTAL_UNSUPPORTED",
                        path + ".subtotalMode",
                        "Subtotal settings apply only to row fields.");
                }
            }
            else
            {
                if (placement.Aggregation.HasValue)
                {
                    result.AddError(
                        "PIVOT_NONVALUE_AGGREGATION_UNSUPPORTED",
                        path + ".aggregation",
                        "Aggregation is supported only for Values placements.");
                }

                if (!string.IsNullOrWhiteSpace(placement.NumberFormatCode))
                {
                    result.AddError(
                        "PIVOT_NONVALUE_NUMBER_FORMAT_UNSUPPORTED",
                        path + ".numberFormatCode",
                        "A value number format cannot be applied outside Values.");
                }

                if (placement.Area != PivotFieldArea.Row && placement.SubtotalMode != PivotSubtotalMode.None)
                {
                    result.AddError(
                        "PIVOT_SUBTOTAL_AREA_UNSUPPORTED",
                        path + ".subtotalMode",
                        "Subtotal settings apply only to row fields.");
                }
            }

            if (placement.NumberFormatCode != null &&
                !IsBoundedText(placement.NumberFormatCode, MaximumExcelNameLength, allowEmpty: false))
            {
                result.AddError(
                    "PIVOT_NUMBER_FORMAT_INVALID",
                    path + ".numberFormatCode",
                    "The number format code is invalid.");
            }
        }

        private static void ValidateSourceValuePlacement(
            PivotSourceDescriptor source,
            PivotFieldDescriptor field,
            PivotFieldPlacement placement,
            string path,
            ValidationResult result)
        {
            switch (source.Kind)
            {
                case PivotSourceKind.WorksheetRange:
                case PivotSourceKind.WorksheetTable:
                    if (field.IsMeasure)
                    {
                        result.AddError(
                            "PIVOT_CLASSIC_MEASURE_UNSUPPORTED",
                            path + ".fieldName",
                            "A classic PivotTable cannot place an OLAP measure field in Values.");
                    }

                    if (placement.Aggregation == PivotAggregationFunction.DistinctCount)
                    {
                        result.AddError(
                            "PIVOT_CLASSIC_DISTINCT_COUNT_UNSUPPORTED",
                            path + ".aggregation",
                            "Distinct count requires a Data Model-backed operation.");
                    }

                    break;
                case PivotSourceKind.DataModel:
                    if (!field.IsMeasure && placement.Aggregation.HasValue &&
                        !SupportsDataModelImplicitMeasure(placement.Aggregation.Value))
                    {
                        result.AddError(
                            "PIVOT_DATA_MODEL_AGGREGATION_UNSUPPORTED",
                            path + ".aggregation",
                            "Implicit Data Model values support only Sum, Count, Average, Minimum, and Maximum.");
                    }

                    break;
                case PivotSourceKind.ExternalOlap:
                    if (!field.IsMeasure)
                    {
                        result.AddError(
                            "PIVOT_EXTERNAL_OLAP_VALUE_FIELD_UNSUPPORTED",
                            path + ".fieldName",
                            "External OLAP Values must reference an existing measure.");
                    }

                    break;
            }
        }

        private static void ValidateValueInstanceMatrix(
            PivotLayoutDefinition definition,
            IReadOnlyDictionary<string, PivotFieldDescriptor> fields,
            ValidationResult result)
        {
            var values = definition.Placements
                .Select((placement, index) => new { Placement = placement, Index = index })
                .Where(item => item.Placement.Area == PivotFieldArea.Values &&
                               fields.ContainsKey(item.Placement.FieldName))
                .ToList();

            if (definition.Source.Kind == PivotSourceKind.DataModel)
            {
                foreach (var group in values
                    .Where(item => !fields[item.Placement.FieldName].IsMeasure &&
                                   item.Placement.Aggregation.HasValue)
                    .GroupBy(
                        item => item.Placement.FieldName + "\0" + item.Placement.Aggregation,
                        StringComparer.OrdinalIgnoreCase)
                    .Where(group => group.Count() > 1))
                {
                    foreach (var duplicate in group.Skip(1))
                    {
                        result.AddError(
                            "PIVOT_DATA_MODEL_IMPLICIT_VALUE_DUPLICATE",
                            "placements[" + duplicate.Index + "]",
                            "A Data Model field and aggregation identify one implicit measure and cannot be repeated under another caption.");
                    }
                }
            }

            if (definition.Source.Kind == PivotSourceKind.DataModel ||
                definition.Source.Kind == PivotSourceKind.ExternalOlap)
            {
                foreach (var group in values
                    .Where(item => fields[item.Placement.FieldName].IsMeasure)
                    .GroupBy(item => item.Placement.FieldName, StringComparer.OrdinalIgnoreCase)
                    .Where(group => group.Count() > 1))
                {
                    foreach (var duplicate in group.Skip(1))
                    {
                        result.AddError(
                            "PIVOT_OLAP_MEASURE_INSTANCE_DUPLICATE",
                            "placements[" + duplicate.Index + "]",
                            "An existing OLAP measure can appear only once unless a separately authored measure is used.");
                    }
                }
            }
        }

        private static void ValidateFilters(
            PivotLayoutDefinition definition,
            IReadOnlyDictionary<string, PivotFieldDescriptor> fields,
            ValidationResult result)
        {
            var filteredFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var placedFields = new HashSet<string>(
                definition.Placements
                    .Where(item => item.Area != PivotFieldArea.Values)
                    .Select(item => item.FieldName),
                StringComparer.OrdinalIgnoreCase);

            for (var index = 0; index < definition.Filters.Count; index++)
            {
                PivotFieldFilter filter = definition.Filters[index];
                string path = "filters[" + index + "]";
                if (!IsExcelName(filter.FieldName, MaximumExcelNameLength))
                {
                    result.AddError("PIVOT_FILTER_FIELD_NAME_INVALID", path + ".fieldName", "The filter field name is invalid.");
                }
                else
                {
                    if (!fields.ContainsKey(filter.FieldName))
                    {
                        result.AddError(
                            "PIVOT_FILTER_FIELD_UNKNOWN",
                            path + ".fieldName",
                            "The filtered field was not present in the discovery snapshot.");
                    }

                    if (!filteredFields.Add(filter.FieldName))
                    {
                        result.AddError(
                            "PIVOT_FILTER_DUPLICATE",
                            path + ".fieldName",
                            "A field can have only one bounded member filter in a layout definition.");
                    }

                    if (!placedFields.Contains(filter.FieldName))
                    {
                        result.AddError(
                            "PIVOT_FILTER_FIELD_NOT_PLACED",
                            path + ".fieldName",
                            "A member filter requires the field to be placed in Rows, Columns, or Filters.");
                    }
                }

                if (!Enum.IsDefined(typeof(PivotFilterMode), filter.Mode))
                {
                    result.AddError("PIVOT_FILTER_MODE_INVALID", path + ".mode", "The filter mode is invalid.");
                }

                var members = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                for (var memberIndex = 0; memberIndex < filter.Members.Count; memberIndex++)
                {
                    string member = filter.Members[memberIndex];
                    string memberPath = path + ".members[" + memberIndex + "]";
                    if (!IsBoundedText(member, MaximumMemberCaptionLength, allowEmpty: false))
                    {
                        result.AddError("PIVOT_FILTER_MEMBER_INVALID", memberPath, "The filter member caption is invalid.");
                    }
                    else if (!members.Add(member))
                    {
                        result.AddError(
                            "PIVOT_FILTER_MEMBER_DUPLICATE",
                            memberPath,
                            "Filter members must be unique without regard to case.");
                    }
                }

                if (filter.Mode == PivotFilterMode.All && (filter.Members.Count > 0 || filter.IncludeBlank))
                {
                    result.AddError(
                        "PIVOT_FILTER_ALL_SELECTION_UNSUPPORTED",
                        path,
                        "An All-members filter cannot also contain a bounded selection.");
                }
                else if ((filter.Mode == PivotFilterMode.Include || filter.Mode == PivotFilterMode.Exclude) &&
                         filter.Members.Count == 0 && !filter.IncludeBlank)
                {
                    result.AddError(
                        "PIVOT_FILTER_SELECTION_REQUIRED",
                        path,
                        "Include and Exclude filters require at least one member or the blank member.");
                }
            }
        }

        private static void ValidateLayout(
            PivotLayoutDefinition definition,
            ValidationResult result)
        {
            PivotLayoutMetadata layout = definition.Layout;
            if (!Enum.IsDefined(typeof(PivotLayoutForm), layout.Form))
            {
                result.AddError("PIVOT_LAYOUT_FORM_INVALID", "layout.form", "The PivotTable layout form is invalid.");
            }

            if (layout.RepeatItemLabels && layout.Form != PivotLayoutForm.Tabular)
            {
                result.AddError(
                    "PIVOT_REPEAT_LABELS_UNSUPPORTED",
                    "layout.repeatItemLabels",
                    "Repeated item labels require tabular layout.");
            }

            if (!Enum.IsDefined(typeof(PivotValuesAxis), layout.ValuesAxis))
            {
                result.AddError(
                    "PIVOT_VALUES_AXIS_INVALID",
                    "layout.valuesAxis",
                    "The Values pseudo-field axis is invalid.");
            }

            if (layout.ValuesPosition < 1)
            {
                result.AddError(
                    "PIVOT_VALUES_POSITION_INVALID",
                    "layout.valuesPosition",
                    "The Values pseudo-field position must be one-based.");
            }

            int valueCount = definition.Placements.Count(placement =>
                placement.Area == PivotFieldArea.Values);
            if (valueCount > 1 && layout.ValuesAxis == PivotValuesAxis.Automatic)
            {
                result.AddError(
                    "PIVOT_VALUES_AXIS_REQUIRED",
                    "layout.valuesAxis",
                    "Two or more Values instances require an explicit Rows or Columns Values axis.");
            }

            if (valueCount == 0 && layout.ValuesAxis != PivotValuesAxis.Automatic)
            {
                result.AddError(
                    "PIVOT_VALUES_AXIS_WITHOUT_VALUES",
                    "layout.valuesAxis",
                    "A Rows or Columns Values axis requires at least one Values placement.");
            }

            if (layout.ValuesAxis == PivotValuesAxis.Automatic && layout.ValuesPosition != 1)
            {
                result.AddError(
                    "PIVOT_VALUES_AUTOMATIC_POSITION_INVALID",
                    "layout.valuesPosition",
                    "Automatic Values-axis placement uses the default position 1.");
            }

            int axisFieldCount = layout.ValuesAxis == PivotValuesAxis.Rows
                ? definition.Placements.Count(placement => placement.Area == PivotFieldArea.Row)
                : definition.Placements.Count(placement => placement.Area == PivotFieldArea.Column);
            if (layout.ValuesAxis != PivotValuesAxis.Automatic &&
                layout.ValuesPosition > axisFieldCount + 1)
            {
                result.AddError(
                    "PIVOT_VALUES_POSITION_OUT_OF_RANGE",
                    "layout.valuesPosition",
                    "The Values pseudo-field position exceeds the selected native axis.");
            }
        }

        private static void ValidateFormat(PivotFormatMetadata format, ValidationResult result)
        {
            ValidateOptionalExcelName(
                format.PivotTableStyleName,
                "format.pivotTableStyleName",
                "PIVOT_STYLE_NAME_INVALID",
                result);
        }

        private static void ValidateDerivedCapabilities(
            PivotLayoutDefinition definition,
            IReadOnlyDictionary<string, PivotFieldDescriptor> fields,
            ValidationResult result)
        {
            PivotCapability required = RequiredMutationCapabilities;
            if (definition.Filters.Count > 0)
            {
                required |= PivotCapability.MemberFiltering;
            }

            if (definition.Placements.Any(placement =>
                    placement.Aggregation == PivotAggregationFunction.DistinctCount))
            {
                required |= PivotCapability.DistinctCount;
            }

            if (definition.Source.Kind == PivotSourceKind.DataModel &&
                definition.Placements.Any(placement => placement.Area == PivotFieldArea.Values))
            {
                required |= PivotCapability.ModelMeasures;
            }

            if (definition.Placements.Any(placement =>
                    fields.TryGetValue(placement.FieldName, out PivotFieldDescriptor? field) &&
                    field.IsCalculated))
            {
                required |= PivotCapability.CalculatedMembers;
            }

            foreach (PivotCapability capability in EnumerateCapabilities(required))
            {
                if ((definition.Source.Capabilities & capability) != capability)
                {
                    result.AddError(
                        "PIVOT_OPERATION_CAPABILITY_REQUIRED",
                        "source.capabilities",
                        "The requested native layout requires the " + capability + " capability.");
                }
            }
        }

        private static void ValidateRequirements(PivotLayoutDefinition definition, ValidationResult result)
        {
            var requirements = new HashSet<PivotCapability>();
            for (var index = 0; index < definition.CapabilityRequirements.Count; index++)
            {
                PivotCapabilityRequirement requirement = definition.CapabilityRequirements[index];
                string path = "capabilityRequirements[" + index + "]";
                bool isSingleKnownCapability = requirement.Capability != PivotCapability.None &&
                    (requirement.Capability & ~KnownCapabilities) == 0 &&
                    (((int)requirement.Capability & ((int)requirement.Capability - 1)) == 0);
                if (!isSingleKnownCapability)
                {
                    result.AddError(
                        "PIVOT_CAPABILITY_REQUIREMENT_INVALID",
                        path + ".capability",
                        "Each capability requirement must identify one known capability.");
                }
                else
                {
                    if (!requirements.Add(requirement.Capability))
                    {
                        result.AddError(
                            "PIVOT_CAPABILITY_REQUIREMENT_DUPLICATE",
                            path + ".capability",
                            "A capability can be required only once.");
                    }

                    if ((definition.Source.Capabilities & requirement.Capability) != requirement.Capability)
                    {
                        result.AddError(
                            "PIVOT_CAPABILITY_UNAVAILABLE",
                            path + ".capability",
                            "The selected PivotTable source does not provide the required capability.");
                    }
                }

                if (!IsBoundedText(requirement.Reason, MaximumCapabilityReasonLength, allowEmpty: false))
                {
                    result.AddError(
                        "PIVOT_CAPABILITY_REASON_INVALID",
                        path + ".reason",
                        "A concise reason is required for every capability requirement.");
                }
            }
        }

        private static IEnumerable<PivotCapability> EnumerateCapabilities(PivotCapability capabilities)
        {
            foreach (PivotCapability capability in Enum.GetValues(typeof(PivotCapability)))
            {
                if (capability != PivotCapability.None &&
                    (capabilities & capability) == capability)
                {
                    yield return capability;
                }
            }
        }

        private static void ValidateCapabilityTruthTable(
            PivotCapability actual,
            PivotCapability allowed,
            string message,
            ValidationResult result)
        {
            if ((actual & ~allowed) != 0)
            {
                result.AddError(
                    "PIVOT_SOURCE_CAPABILITY_CONFLICT",
                    "source.capabilities",
                    message);
            }
        }

        private static bool SupportsDataModelImplicitMeasure(
            PivotAggregationFunction function)
        {
            return function == PivotAggregationFunction.Sum ||
                   function == PivotAggregationFunction.Count ||
                   function == PivotAggregationFunction.Average ||
                   function == PivotAggregationFunction.Minimum ||
                   function == PivotAggregationFunction.Maximum;
        }

        private static bool Supports(PivotFieldAreaSupport support, PivotFieldArea area)
        {
            switch (area)
            {
                case PivotFieldArea.Row: return (support & PivotFieldAreaSupport.Row) != 0;
                case PivotFieldArea.Column: return (support & PivotFieldAreaSupport.Column) != 0;
                case PivotFieldArea.Filter: return (support & PivotFieldAreaSupport.Filter) != 0;
                case PivotFieldArea.Values: return (support & PivotFieldAreaSupport.Values) != 0;
                default: return false;
            }
        }

        private static bool IsToken(string value, int maximumLength)
        {
            if (string.IsNullOrWhiteSpace(value) || value.Length > maximumLength || value != value.Trim())
            {
                return false;
            }

            foreach (char character in value)
            {
                if (!char.IsLetterOrDigit(character) && character != '.' && character != '_' && character != '-')
                {
                    return false;
                }
            }

            return true;
        }

        private static bool IsExcelName(string value, int maximumLength)
        {
            return IsBoundedText(value, maximumLength, allowEmpty: false);
        }

        private static bool IsBoundedText(string value, int maximumLength, bool allowEmpty)
        {
            if (value == null || value.Length > maximumLength || value != value.Trim()) return false;
            if (!allowEmpty && string.IsNullOrWhiteSpace(value)) return false;
            return !value.Any(char.IsControl);
        }

        private static void ValidateOptionalExcelName(
            string? value,
            string path,
            string code,
            ValidationResult result)
        {
            if (value != null && !IsExcelName(value, MaximumExcelNameLength))
            {
                result.AddError(code, path, "The optional Excel display name is invalid.");
            }
        }

        private static void ValidateOptionalPathFreeExcelName(
            string? value,
            string path,
            string code,
            ValidationResult result)
        {
            if (value != null &&
                (!IsExcelName(value, MaximumExcelNameLength) ||
                 !PivotPlusPathPolicy.IsPathFree(value)))
            {
                result.AddError(code, path, "The optional Excel source name must be path-free.");
            }
        }
    }
}
