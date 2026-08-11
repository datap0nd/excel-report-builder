using ExcelReportBuilder.Core.PivotPlus;
using ExcelReportBuilder.Core.Validation;

namespace ExcelReportBuilder.Core.Tests;

public sealed class PivotPlusValidationTests
{
    [Fact]
    public void Accepts_a_path_free_native_layout_snapshot()
    {
        PivotLayoutDefinition definition = CreateValidDefinition();

        ValidationResult result = PivotPlusValidator.Validate(definition);

        Assert.True(result.IsValid, Format(result));
        Assert.Equal("workbook-7f0c", definition.Target.WorkbookId);
        Assert.Equal(new[] { "Region", "Period", "Cost" }, definition.Fields.Select(field => field.Name));
    }

    [Fact]
    public void Defensively_copies_discovery_and_layout_collections()
    {
        var fields = ValidFields().ToList();
        var placements = ValidPlacements().ToList();
        var members = new List<string> { "North" };
        var filters = new List<PivotFieldFilter>
        {
            new PivotFieldFilter("Region", PivotFilterMode.Include, members)
        };
        var definition = CreateDefinition(fields, placements, filters: filters);

        fields.Clear();
        placements.Clear();
        members.Add("South");
        filters.Clear();

        Assert.Equal(3, definition.Fields.Count);
        Assert.Equal(3, definition.Placements.Count);
        Assert.Single(definition.Filters);
        Assert.Equal(new[] { "North" }, definition.Filters[0].Members);
    }

    [Fact]
    public void Requires_explicit_clear_all_for_an_empty_native_layout()
    {
        PivotLayoutDefinition omitted = CreateDefinition(
            ValidFields(),
            Array.Empty<PivotFieldPlacement>());
        PivotLayoutDefinition conflicting = CreateDefinition(
            ValidFields(),
            ValidPlacements(),
            clearAll: true);
        PivotLayoutDefinition explicitClear = CreateDefinition(
            ValidFields(),
            Array.Empty<PivotFieldPlacement>(),
            clearAll: true);

        ValidationResult omittedResult = PivotPlusValidator.Validate(omitted);
        ValidationResult conflictingResult = PivotPlusValidator.Validate(conflicting);
        ValidationResult explicitResult = PivotPlusValidator.Validate(explicitClear);

        Assert.Contains(omittedResult.Issues, issue =>
            issue.Code == "PIVOT_LAYOUT_PLACEMENT_REQUIRED");
        Assert.Contains(conflictingResult.Issues, issue =>
            issue.Code == "PIVOT_CLEAR_ALL_PLACEMENT_CONFLICT");
        Assert.True(explicitResult.IsValid, Format(explicitResult));
    }

    [Fact]
    public void Multi_value_layout_requires_an_explicit_bounded_values_axis()
    {
        var fields = ValidFields().Concat(new[]
        {
            new PivotFieldDescriptor(
                "Units",
                "Units",
                PivotFieldDataType.Number,
                PivotFieldAreaSupport.Values)
        });
        var placements = ValidPlacements().Concat(new[]
        {
            new PivotFieldPlacement(
                "Units",
                PivotFieldArea.Values,
                2,
                aggregation: PivotAggregationFunction.Sum)
        });
        PivotLayoutDefinition automatic = CreateDefinition(
            fields,
            placements,
            layout: new PivotLayoutMetadata(PivotLayoutForm.Tabular));
        PivotLayoutDefinition outOfRange = CreateDefinition(
            fields,
            placements,
            layout: new PivotLayoutMetadata(
                PivotLayoutForm.Tabular,
                valuesAxis: PivotValuesAxis.Rows,
                valuesPosition: 3));

        ValidationResult automaticResult = PivotPlusValidator.Validate(automatic);
        ValidationResult rangeResult = PivotPlusValidator.Validate(outOfRange);

        Assert.Contains(automaticResult.Issues, issue =>
            issue.Code == "PIVOT_VALUES_AXIS_REQUIRED");
        Assert.Contains(rangeResult.Issues, issue =>
            issue.Code == "PIVOT_VALUES_POSITION_OUT_OF_RANGE");
    }

    [Theory]
    [InlineData(@"C:\reports\book.xlsx")]
    [InlineData(@"folder\book.xlsx")]
    [InlineData("file:///book.xlsx")]
    [InlineData("workbook id")]
    public void Rejects_non_token_or_path_like_workbook_ids(string workbookId)
    {
        PivotLayoutDefinition definition = CreateDefinition(
            ValidFields(),
            ValidPlacements(),
            target: new PivotTargetIdentity(workbookId, "Analysis", "SalesPivot"));

        ValidationResult result = PivotPlusValidator.Validate(definition);

        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_TARGET_WORKBOOK_ID_INVALID");
    }

    [Theory]
    [InlineData("")]
    [InlineData("Bad/Sheet")]
    [InlineData("'Quoted'")]
    [InlineData("This worksheet name is more than thirty one characters")]
    public void Rejects_invalid_target_worksheet_names(string worksheetName)
    {
        PivotLayoutDefinition definition = CreateDefinition(
            ValidFields(),
            ValidPlacements(),
            target: new PivotTargetIdentity("workbook-7f0c", worksheetName, "SalesPivot"));

        ValidationResult result = PivotPlusValidator.Validate(definition);

        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_TARGET_WORKSHEET_NAME_INVALID");
    }

    [Fact]
    public void Rejects_path_like_source_and_pivot_names_and_blank_field_names()
    {
        var fields = new[]
        {
            new PivotFieldDescriptor(" ", "Invalid", PivotFieldDataType.Text, PivotFieldAreaSupport.Row)
        };
        var placements = new[]
        {
            new PivotFieldPlacement(" ", PivotFieldArea.Row, 1)
        };
        PivotLayoutDefinition definition = CreateDefinition(
            fields,
            placements,
            target: new PivotTargetIdentity("workbook-7f0c", "Analysis", @"folder\SalesPivot"),
            source: new PivotSourceDescriptor(
                PivotSourceKind.WorksheetTable,
                @"C:\data\sales.xlsx",
                PivotCapability.NativeFieldPlacement));

        ValidationResult result = PivotPlusValidator.Validate(definition);

        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_TARGET_NAME_INVALID");
        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_SOURCE_NAME_INVALID");
        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_FIELD_NAME_INVALID");
        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_PLACEMENT_FIELD_NAME_INVALID");
    }

    [Fact]
    public void Rejects_path_like_model_and_field_table_source_identities()
    {
        var fields = ValidFields()
            .Select(field => field.Name == "Cost"
                ? new PivotFieldDescriptor(
                    field.Name,
                    field.Caption,
                    field.DataType,
                    field.SupportedAreas,
                    tableName: "C" + ":" + '\\' + "private" + '\\' + "Sales")
                : field)
            .ToArray();
        var source = new PivotSourceDescriptor(
            PivotSourceKind.DataModel,
            "ThisWorkbookDataModel",
            PivotCapability.NativeFieldPlacement |
            PivotCapability.MemberFiltering |
            PivotCapability.LayoutFormatting |
            PivotCapability.Refresh |
            PivotCapability.DataModel |
            PivotCapability.ModelMeasures,
            modelTableName: new string('\\', 2) + "host" + '\\' + "object");

        ValidationResult result = PivotPlusValidator.Validate(
            CreateDefinition(fields, ValidPlacements(), source: source));

        Assert.Contains(result.Issues, issue =>
            issue.Code == "PIVOT_SOURCE_MODEL_TABLE_NAME_INVALID");
        Assert.Contains(result.Issues, issue =>
            issue.Code == "PIVOT_FIELD_TABLE_NAME_INVALID");
    }

    [Fact]
    public void Rejects_duplicate_field_descriptors_and_placements_without_regard_to_case()
    {
        var fields = ValidFields().Concat(new[]
        {
            new PivotFieldDescriptor("region", "Duplicate", PivotFieldDataType.Text, PivotFieldAreaSupport.Row)
        });
        var placements = ValidPlacements().Concat(new[]
        {
            new PivotFieldPlacement("region", PivotFieldArea.Row, 2)
        });
        PivotLayoutDefinition definition = CreateDefinition(fields, placements);

        ValidationResult result = PivotPlusValidator.Validate(definition);

        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_FIELD_DUPLICATE");
        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_PLACEMENT_DUPLICATE");
    }

    [Fact]
    public void Rejects_the_same_regular_field_on_multiple_nonvalue_axes()
    {
        var fields = ValidFields().Select(field => field.Name == "Region"
            ? new PivotFieldDescriptor(
                field.Name,
                field.Caption,
                field.DataType,
                PivotFieldAreaSupport.Row | PivotFieldAreaSupport.Column)
            : field);
        var placements = new[]
        {
            new PivotFieldPlacement("Region", PivotFieldArea.Row, 1),
            new PivotFieldPlacement("Region", PivotFieldArea.Column, 1)
        };

        ValidationResult result = PivotPlusValidator.Validate(
            CreateDefinition(fields, placements));

        Assert.Contains(result.Issues, issue =>
            issue.Code == "PIVOT_NONVALUE_FIELD_MULTIPLE_AREAS");
    }

    [Fact]
    public void Allows_repeated_value_field_with_distinct_captions()
    {
        var placements = ValidPlacements()
            .Where(item => item.Area != PivotFieldArea.Values)
            .Concat(new[]
            {
                new PivotFieldPlacement(
                    "Cost",
                    PivotFieldArea.Values,
                    1,
                    caption: "Cost",
                    aggregation: PivotAggregationFunction.Sum),
                new PivotFieldPlacement(
                    "Cost",
                    PivotFieldArea.Values,
                    2,
                    caption: "Cost portion",
                    aggregation: PivotAggregationFunction.Sum)
            });

        ValidationResult result = PivotPlusValidator.Validate(
            CreateDefinition(ValidFields(), placements));

        Assert.True(result.IsValid, Format(result));
    }

    [Fact]
    public void Requires_distinct_captions_for_repeated_value_instances()
    {
        var placements = ValidPlacements()
            .Where(item => item.Area != PivotFieldArea.Values)
            .Concat(new[]
            {
                new PivotFieldPlacement(
                    "Cost",
                    PivotFieldArea.Values,
                    1,
                    aggregation: PivotAggregationFunction.Sum),
                new PivotFieldPlacement(
                    "Cost",
                    PivotFieldArea.Values,
                    2,
                    caption: "Cost",
                    aggregation: PivotAggregationFunction.Sum),
                new PivotFieldPlacement(
                    "Units",
                    PivotFieldArea.Values,
                    3,
                    caption: "cost",
                    aggregation: PivotAggregationFunction.Sum)
            });
        var fields = ValidFields().Concat(new[]
        {
            new PivotFieldDescriptor(
                "Units",
                "Units",
                PivotFieldDataType.Number,
                PivotFieldAreaSupport.Values)
        });

        ValidationResult result = PivotPlusValidator.Validate(
            CreateDefinition(fields, placements));

        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_VALUE_INSTANCE_CAPTION_REQUIRED");
        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_VALUE_CAPTION_DUPLICATE");
    }

    [Fact]
    public void Rejects_collisions_between_explicit_and_automatically_resolved_value_captions()
    {
        var fields = ValidFields().Concat(new[]
        {
            new PivotFieldDescriptor(
                "Units",
                "Units",
                PivotFieldDataType.Number,
                PivotFieldAreaSupport.Values)
        });
        var placements = new[]
        {
            new PivotFieldPlacement(
                "Cost",
                PivotFieldArea.Values,
                1,
                aggregation: PivotAggregationFunction.Sum),
            new PivotFieldPlacement(
                "Units",
                PivotFieldArea.Values,
                2,
                caption: "Sum of Cost",
                aggregation: PivotAggregationFunction.Sum)
        };

        ValidationResult result = PivotPlusValidator.Validate(
            CreateDefinition(fields, placements));

        Assert.Contains(result.Issues, issue =>
            issue.Code == "PIVOT_VALUE_RESOLVED_CAPTION_DUPLICATE");
    }

    [Fact]
    public void Enforces_data_model_implicit_measure_function_and_instance_identity()
    {
        var source = new PivotSourceDescriptor(
            PivotSourceKind.DataModel,
            "ThisWorkbookDataModel",
            PivotCapability.NativeFieldPlacement |
            PivotCapability.LayoutFormatting |
            PivotCapability.Refresh |
            PivotCapability.DataModel |
            PivotCapability.ModelMeasures);
        var fields = new[]
        {
            new PivotFieldDescriptor(
                "Cost",
                "Cost",
                PivotFieldDataType.Unknown,
                PivotFieldAreaSupport.Values,
                tableName: "Sales")
        };
        var placements = new[]
        {
            new PivotFieldPlacement(
                "Cost",
                PivotFieldArea.Values,
                1,
                caption: "Cost A",
                aggregation: PivotAggregationFunction.Product),
            new PivotFieldPlacement(
                "Cost",
                PivotFieldArea.Values,
                2,
                caption: "Cost B",
                aggregation: PivotAggregationFunction.Product)
        };

        ValidationResult result = PivotPlusValidator.Validate(
            CreateDefinition(fields, placements, source: source));

        Assert.Contains(result.Issues, issue =>
            issue.Code == "PIVOT_DATA_MODEL_AGGREGATION_UNSUPPORTED");
        Assert.Contains(result.Issues, issue =>
            issue.Code == "PIVOT_DATA_MODEL_IMPLICIT_VALUE_DUPLICATE");
    }

    [Fact]
    public void External_olap_values_require_unique_existing_measures()
    {
        var source = new PivotSourceDescriptor(
            PivotSourceKind.ExternalOlap,
            "ExternalCube",
            PivotCapability.NativeFieldPlacement |
            PivotCapability.LayoutFormatting |
            PivotCapability.Refresh);
        var fields = new[]
        {
            new PivotFieldDescriptor(
                "[Measures].[Revenue]",
                "Revenue",
                PivotFieldDataType.Unknown,
                PivotFieldAreaSupport.Values,
                isMeasure: true),
            new PivotFieldDescriptor(
                "[Sales].[Cost]",
                "Cost",
                PivotFieldDataType.Unknown,
                PivotFieldAreaSupport.Values)
        };
        var placements = new[]
        {
            new PivotFieldPlacement(
                "[Measures].[Revenue]",
                PivotFieldArea.Values,
                1,
                caption: "Revenue A"),
            new PivotFieldPlacement(
                "[Measures].[Revenue]",
                PivotFieldArea.Values,
                2,
                caption: "Revenue B"),
            new PivotFieldPlacement(
                "[Sales].[Cost]",
                PivotFieldArea.Values,
                3,
                caption: "Cost",
                aggregation: PivotAggregationFunction.Sum)
        };

        ValidationResult result = PivotPlusValidator.Validate(
            CreateDefinition(fields, placements, source: source));

        Assert.Contains(result.Issues, issue =>
            issue.Code == "PIVOT_OLAP_MEASURE_INSTANCE_DUPLICATE");
        Assert.Contains(result.Issues, issue =>
            issue.Code == "PIVOT_EXTERNAL_OLAP_VALUE_FIELD_UNSUPPORTED");
    }

    [Fact]
    public void Rejects_unknown_fields_and_areas_not_supported_by_discovery()
    {
        var placements = new[]
        {
            new PivotFieldPlacement("Cost", PivotFieldArea.Row, 1),
            new PivotFieldPlacement("Missing", PivotFieldArea.Column, 1)
        };
        PivotLayoutDefinition definition = CreateDefinition(ValidFields(), placements);

        ValidationResult result = PivotPlusValidator.Validate(definition);

        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_PLACEMENT_AREA_UNSUPPORTED");
        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_PLACEMENT_FIELD_UNKNOWN");
    }

    [Fact]
    public void Rejects_zero_duplicate_and_non_contiguous_area_positions()
    {
        var placements = new[]
        {
            new PivotFieldPlacement("Region", PivotFieldArea.Row, 0),
            new PivotFieldPlacement("Period", PivotFieldArea.Column, 1),
            new PivotFieldPlacement("Cost", PivotFieldArea.Values, 2, aggregation: PivotAggregationFunction.Sum),
            new PivotFieldPlacement("Units", PivotFieldArea.Values, 2, aggregation: PivotAggregationFunction.Sum)
        };
        var fields = ValidFields().Concat(new[]
        {
            new PivotFieldDescriptor("Units", "Units", PivotFieldDataType.Number, PivotFieldAreaSupport.Values)
        });
        PivotLayoutDefinition definition = CreateDefinition(fields, placements);

        ValidationResult result = PivotPlusValidator.Validate(definition);

        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_PLACEMENT_POSITION_UNSUPPORTED");
        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_PLACEMENT_POSITION_DUPLICATE");
        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_PLACEMENT_POSITION_GAP");
    }

    [Fact]
    public void Enforces_value_and_nonvalue_option_boundaries()
    {
        var placements = new[]
        {
            new PivotFieldPlacement(
                "Region",
                PivotFieldArea.Row,
                1,
                aggregation: PivotAggregationFunction.Count,
                numberFormatCode: "0",
                subtotalMode: PivotSubtotalMode.Automatic),
            new PivotFieldPlacement("Cost", PivotFieldArea.Values, 1),
            new PivotFieldPlacement("Existing Measure", PivotFieldArea.Values, 2, aggregation: PivotAggregationFunction.Sum)
        };
        var fields = ValidFields().Concat(new[]
        {
            new PivotFieldDescriptor(
                "Existing Measure",
                "Existing Measure",
                PivotFieldDataType.Number,
                PivotFieldAreaSupport.Values,
                isMeasure: true)
        });
        PivotLayoutDefinition definition = CreateDefinition(fields, placements);

        ValidationResult result = PivotPlusValidator.Validate(definition);

        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_NONVALUE_AGGREGATION_UNSUPPORTED");
        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_NONVALUE_NUMBER_FORMAT_UNSUPPORTED");
        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_VALUE_AGGREGATION_REQUIRED");
        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_MEASURE_AGGREGATION_UNSUPPORTED");
    }

    [Fact]
    public void Rejects_distinct_count_without_source_capability()
    {
        var placements = ValidPlacements()
            .Select(item => item.Area == PivotFieldArea.Values
                ? new PivotFieldPlacement("Cost", PivotFieldArea.Values, 1, aggregation: PivotAggregationFunction.DistinctCount)
                : item);
        PivotLayoutDefinition definition = CreateDefinition(
            ValidFields(),
            placements,
            source: new PivotSourceDescriptor(
                PivotSourceKind.WorksheetTable,
                "SalesTable",
                PivotCapability.NativeFieldPlacement | PivotCapability.MemberFiltering));

        ValidationResult result = PivotPlusValidator.Validate(definition);

        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_DISTINCT_COUNT_UNSUPPORTED");
    }

    [Fact]
    public void Validates_bounded_filters_and_requires_a_nonvalue_placement()
    {
        var filters = new[]
        {
            new PivotFieldFilter("Region", PivotFilterMode.Include, new[] { "North", "north" }),
            new PivotFieldFilter("Region", PivotFilterMode.Exclude, includeBlank: false),
            new PivotFieldFilter("Cost", PivotFilterMode.All, new[] { "10" })
        };
        PivotLayoutDefinition definition = CreateDefinition(ValidFields(), ValidPlacements(), filters: filters);

        ValidationResult result = PivotPlusValidator.Validate(definition);

        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_FILTER_MEMBER_DUPLICATE");
        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_FILTER_DUPLICATE");
        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_FILTER_SELECTION_REQUIRED");
        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_FILTER_FIELD_NOT_PLACED");
        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_FILTER_ALL_SELECTION_UNSUPPORTED");
    }

    [Fact]
    public void Rejects_unavailable_duplicate_and_composite_capability_requirements()
    {
        var requirements = new[]
        {
            new PivotCapabilityRequirement(PivotCapability.NativeFieldPlacement, "Place fields"),
            new PivotCapabilityRequirement(PivotCapability.NativeFieldPlacement, "Duplicate"),
            new PivotCapabilityRequirement(PivotCapability.NamedSets, "Asymmetric columns"),
            new PivotCapabilityRequirement(
                PivotCapability.MemberFiltering | PivotCapability.LayoutFormatting,
                "Composite requirement")
        };
        PivotLayoutDefinition definition = CreateDefinition(
            ValidFields(),
            ValidPlacements(),
            capabilityRequirements: requirements);

        ValidationResult result = PivotPlusValidator.Validate(definition);

        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_CAPABILITY_REQUIREMENT_DUPLICATE");
        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_CAPABILITY_UNAVAILABLE");
        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_CAPABILITY_REQUIREMENT_INVALID");
    }

    [Fact]
    public void Derives_operation_capabilities_and_rejects_source_kind_contradictions()
    {
        PivotLayoutDefinition missing = CreateDefinition(
            ValidFields(),
            ValidPlacements(),
            source: new PivotSourceDescriptor(
                PivotSourceKind.WorksheetTable,
                "SalesTable",
                PivotCapability.None));
        PivotLayoutDefinition contradictory = CreateDefinition(
            ValidFields(),
            ValidPlacements(),
            source: new PivotSourceDescriptor(
                PivotSourceKind.WorksheetTable,
                "SalesTable",
                PivotCapability.NativeFieldPlacement |
                PivotCapability.MemberFiltering |
                PivotCapability.LayoutFormatting |
                PivotCapability.Refresh |
                PivotCapability.DistinctCount));

        ValidationResult missingResult = PivotPlusValidator.Validate(missing);
        ValidationResult contradictoryResult = PivotPlusValidator.Validate(contradictory);

        Assert.Contains(missingResult.Issues, issue =>
            issue.Code == "PIVOT_OPERATION_CAPABILITY_REQUIRED");
        Assert.Contains(contradictoryResult.Issues, issue =>
            issue.Code == "PIVOT_SOURCE_CAPABILITY_CONFLICT");
    }

    [Fact]
    public void Rejects_model_capabilities_on_worksheet_sources_and_missing_model_marker()
    {
        PivotLayoutDefinition worksheet = CreateDefinition(
            ValidFields(),
            ValidPlacements(),
            source: new PivotSourceDescriptor(
                PivotSourceKind.WorksheetTable,
                "SalesTable",
                PivotCapability.NativeFieldPlacement | PivotCapability.NamedSets,
                "SalesModel"));
        PivotLayoutDefinition model = CreateDefinition(
            ValidFields(),
            ValidPlacements(),
            source: new PivotSourceDescriptor(
                PivotSourceKind.DataModel,
                "ModelConnection",
                PivotCapability.NativeFieldPlacement));

        ValidationResult worksheetResult = PivotPlusValidator.Validate(worksheet);
        ValidationResult modelResult = PivotPlusValidator.Validate(model);

        Assert.Contains(worksheetResult.Issues, issue => issue.Code == "PIVOT_SOURCE_CAPABILITY_CONFLICT");
        Assert.Contains(worksheetResult.Issues, issue => issue.Code == "PIVOT_SOURCE_MODEL_TABLE_UNSUPPORTED");
        Assert.Contains(modelResult.Issues, issue => issue.Code == "PIVOT_SOURCE_DATA_MODEL_CAPABILITY_REQUIRED");
    }

    [Fact]
    public void Repeated_labels_require_tabular_layout()
    {
        PivotLayoutDefinition definition = CreateDefinition(
            ValidFields(),
            ValidPlacements(),
            layout: new PivotLayoutMetadata(PivotLayoutForm.Compact, repeatItemLabels: true));

        ValidationResult result = PivotPlusValidator.Validate(definition);

        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_REPEAT_LABELS_UNSUPPORTED");
    }

    private static PivotLayoutDefinition CreateValidDefinition()
    {
        return CreateDefinition(
            ValidFields(),
            ValidPlacements(),
            filters: new[]
            {
                new PivotFieldFilter("Region", PivotFilterMode.Include, new[] { "North", "South" })
            },
            capabilityRequirements: new[]
            {
                new PivotCapabilityRequirement(PivotCapability.NativeFieldPlacement, "Place native fields"),
                new PivotCapabilityRequirement(PivotCapability.MemberFiltering, "Apply a bounded member filter")
            });
    }

    private static PivotLayoutDefinition CreateDefinition(
        IEnumerable<PivotFieldDescriptor> fields,
        IEnumerable<PivotFieldPlacement> placements,
        PivotTargetIdentity? target = null,
        PivotSourceDescriptor? source = null,
        IEnumerable<PivotFieldFilter>? filters = null,
        PivotLayoutMetadata? layout = null,
        PivotFormatMetadata? format = null,
        IEnumerable<PivotCapabilityRequirement>? capabilityRequirements = null,
        bool clearAll = false)
    {
        PivotFieldPlacement[] placementSnapshot = placements.ToArray();
        PivotLayoutMetadata resolvedLayout = layout ?? new PivotLayoutMetadata(
            PivotLayoutForm.Tabular,
            repeatItemLabels: true,
            valuesAxis: placementSnapshot.Count(item => item.Area == PivotFieldArea.Values) > 1
                ? PivotValuesAxis.Columns
                : PivotValuesAxis.Automatic,
            valuesPosition: 1);
        return new PivotLayoutDefinition(
            target ?? new PivotTargetIdentity("workbook-7f0c", "Analysis", "SalesPivot"),
            source ?? new PivotSourceDescriptor(
                PivotSourceKind.WorksheetTable,
                "SalesTable",
                PivotCapability.NativeFieldPlacement |
                PivotCapability.MemberFiltering |
                PivotCapability.LayoutFormatting |
                PivotCapability.Refresh),
            fields,
            placementSnapshot,
            filters,
            resolvedLayout,
            format ?? new PivotFormatMetadata("PivotStyleMedium2", showRowStripes: true),
            capabilityRequirements,
            clearAll);
    }

    private static IEnumerable<PivotFieldDescriptor> ValidFields()
    {
        return new[]
        {
            new PivotFieldDescriptor(
                "Region",
                "Region",
                PivotFieldDataType.Text,
                PivotFieldAreaSupport.Row | PivotFieldAreaSupport.Filter),
            new PivotFieldDescriptor(
                "Period",
                "Period",
                PivotFieldDataType.Date,
                PivotFieldAreaSupport.Column | PivotFieldAreaSupport.Filter),
            new PivotFieldDescriptor(
                "Cost",
                "Cost",
                PivotFieldDataType.Number,
                PivotFieldAreaSupport.Values)
        };
    }

    private static IEnumerable<PivotFieldPlacement> ValidPlacements()
    {
        return new[]
        {
            new PivotFieldPlacement("Region", PivotFieldArea.Row, 1, subtotalMode: PivotSubtotalMode.Automatic),
            new PivotFieldPlacement("Period", PivotFieldArea.Column, 1),
            new PivotFieldPlacement(
                "Cost",
                PivotFieldArea.Values,
                1,
                aggregation: PivotAggregationFunction.Sum,
                numberFormatCode: "#,##0")
        };
    }

    private static string Format(ValidationResult result)
    {
        return string.Join(
            Environment.NewLine,
            result.Issues.Select(issue => issue.Code + " " + issue.Path + ": " + issue.Message));
    }
}
