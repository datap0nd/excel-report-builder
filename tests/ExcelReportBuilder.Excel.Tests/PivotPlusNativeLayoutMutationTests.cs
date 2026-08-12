using ExcelReportBuilder.Core.PivotPlus;
using ExcelReportBuilder.Excel.PivotPlus;
using ExcelReportBuilder.Excel.PivotPlus.Native;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class PivotPlusNativeLayoutMutationTests
{
    [Fact]
    public void Applies_classic_axes_values_layout_totals_subtotals_and_format_metadata()
    {
        PivotSourceDescriptor source = ClassicSource();
        PivotFieldDescriptor[] fields = ClassicFields();
        PivotTableContext context = Context(source, fields);
        PivotLayoutDefinition definition = Definition(
            source,
            fields,
            new[]
            {
                new PivotFieldPlacement(
                    "Region",
                    PivotFieldArea.Row,
                    1,
                    subtotalMode: PivotSubtotalMode.Automatic),
                new PivotFieldPlacement("Month", PivotFieldArea.Column, 1),
                new PivotFieldPlacement("Category", PivotFieldArea.Filter, 1),
                new PivotFieldPlacement(
                    "Cost",
                    PivotFieldArea.Values,
                    1,
                    "Actual Cost",
                    PivotAggregationFunction.Sum,
                    "#,##0")
            },
            new PivotLayoutMetadata(
                PivotLayoutForm.Tabular,
                repeatItemLabels: true,
                showRowGrandTotals: true,
                showColumnGrandTotals: false,
                showFieldHeaders: false),
            new PivotFormatMetadata(
                "PivotStyleMedium2",
                preserveFormatting: true,
                showRowStripes: true,
                showColumnStripes: false));
        var adapter = new RecordingAdapter();
        var pivot = new FakePivot();

        new PivotTableNativeLayoutMutationService(
            adapter,
            new PivotMutationCoordinator()).Apply(pivot, context, definition);

        Assert.Equal(
            new[]
            {
                "bind",
                "bind-source",
                "capture",
                "persist",
                "clear",
                "place:row:0001:Region",
                "place:column:0001:Month",
                "place:filter:0001:Category",
                "place:value:0001:Actual Cost",
                "layout",
                "refresh",
                "verify"
            },
            adapter.Calls);
        Assert.False(pivot.ManualUpdate);

        NativePivotFieldCommand value = adapter.Placed.Single(item =>
            item.Area == PivotFieldArea.Values);
        Assert.Equal(-4157, value.ConsolidationFunction);
        Assert.Equal("#,##0", value.NumberFormatCode);
        Assert.Equal("Actual Cost", value.Caption);
        NativePivotFieldCommand row = adapter.Placed.Single(item =>
            item.Area == PivotFieldArea.Row);
        Assert.Equal(PivotSubtotalMode.Automatic, row.SubtotalMode);

        Assert.NotNull(adapter.Layout);
        Assert.Equal(1, adapter.Layout!.RowAxisLayout);
        Assert.True(adapter.Layout.RepeatItemLabels);
        Assert.True(adapter.Layout.ShowRowGrandTotals);
        Assert.False(adapter.Layout.ShowColumnGrandTotals);
        Assert.False(adapter.Layout.ShowFieldHeaders);
        Assert.Equal("PivotStyleMedium2", adapter.Layout.PivotTableStyleName);
        Assert.True(adapter.Layout.PreserveFormatting);
        Assert.True(adapter.Layout.ShowRowStripes);
        Assert.False(adapter.Layout.ShowColumnStripes);
    }

    [Fact]
    public void Supports_repeated_classic_values_as_ordered_captioned_instances()
    {
        PivotSourceDescriptor source = ClassicSource();
        PivotFieldDescriptor[] fields = ClassicFields();
        PivotLayoutDefinition definition = Definition(
            source,
            fields,
            new[]
            {
                new PivotFieldPlacement(
                    "Cost",
                    PivotFieldArea.Values,
                    1,
                    "Cost total",
                    PivotAggregationFunction.Sum),
                new PivotFieldPlacement(
                    "Cost",
                    PivotFieldArea.Values,
                    2,
                    "Cost total copy",
                    PivotAggregationFunction.Sum)
            });
        var service = new PivotTableNativeLayoutMutationService(
            new RecordingAdapter(),
            new PivotMutationCoordinator());

        NativePivotMutationPlan plan = service.Compile(Context(source, fields), definition);

        Assert.Equal(2, plan.Fields.Count);
        Assert.Equal("value:0001:Cost total", plan.Fields[0].InstanceId);
        Assert.Equal("value:0002:Cost total copy", plan.Fields[1].InstanceId);
        Assert.Equal(-4157, plan.Fields[0].ConsolidationFunction);
        Assert.Equal(-4157, plan.Fields[1].ConsolidationFunction);
        Assert.Equal(PivotValuesAxis.Columns, plan.Layout.ValuesAxis);
        Assert.Equal(1, plan.Layout.ValuesPosition);
    }

    [Fact]
    public void Rejects_repeated_values_without_explicit_unique_instance_captions()
    {
        PivotSourceDescriptor source = ClassicSource();
        PivotFieldDescriptor[] fields = ClassicFields();
        PivotLayoutDefinition definition = Definition(
            source,
            fields,
            new[]
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
                    aggregation: PivotAggregationFunction.Average)
            });
        var adapter = new RecordingAdapter();

        PivotTableNativeMutationValidationException exception =
            Assert.Throws<PivotTableNativeMutationValidationException>(() =>
                new PivotTableNativeLayoutMutationService(
                    adapter,
                    new PivotMutationCoordinator()).Apply(
                    new FakePivot(),
                    Context(source, fields),
                    definition));

        Assert.Contains(
            exception.Issues,
            issue => issue.Code == "PIVOT_VALUE_INSTANCE_CAPTION_REQUIRED");
        Assert.Empty(adapter.Calls);
    }

    [Fact]
    public void Compiles_supported_data_model_implicit_aggregates_and_existing_measures()
    {
        PivotSourceDescriptor source = ModelSource();
        var fields = new[]
        {
            Field("[Sales].[Region]", PivotFieldAreaSupport.Row),
            Field("[Sales].[CustomerKey]", PivotFieldAreaSupport.Values, PivotFieldDataType.Number),
            Field(
                "[Measures].[Gross Margin]",
                PivotFieldAreaSupport.Values,
                PivotFieldDataType.Number,
                isMeasure: true)
        };
        PivotLayoutDefinition definition = Definition(
            source,
            fields,
            new[]
            {
                new PivotFieldPlacement("[Sales].[Region]", PivotFieldArea.Row, 1),
                new PivotFieldPlacement(
                    "[Sales].[CustomerKey]",
                    PivotFieldArea.Values,
                    1,
                    "Average Customer Key",
                    PivotAggregationFunction.Average,
                    "0.0"),
                new PivotFieldPlacement(
                    "[Measures].[Gross Margin]",
                    PivotFieldArea.Values,
                    2,
                    "Gross Margin",
                    numberFormatCode: "0.0%")
            });
        var service = new PivotTableNativeLayoutMutationService(
            new RecordingAdapter(),
            new PivotMutationCoordinator());

        NativePivotMutationPlan plan = service.Compile(Context(source, fields), definition);

        Assert.Equal(PivotSourceKind.DataModel, plan.SourceKind);
        NativePivotFieldCommand implicitMeasure = plan.Fields.Single(item =>
            item.Caption == "Average Customer Key");
        Assert.False(implicitMeasure.IsMeasure);
        Assert.Equal(-4106, implicitMeasure.ConsolidationFunction);
        NativePivotFieldCommand existingMeasure = plan.Fields.Single(item =>
            item.Caption == "Gross Margin");
        Assert.True(existingMeasure.IsMeasure);
        Assert.Null(existingMeasure.ConsolidationFunction);
        Assert.Equal("0.0%", existingMeasure.NumberFormatCode);
    }

    [Theory]
    [InlineData(PivotAggregationFunction.Product)]
    [InlineData(PivotAggregationFunction.CountNumbers)]
    [InlineData(PivotAggregationFunction.StandardDeviation)]
    [InlineData(PivotAggregationFunction.StandardDeviationPopulation)]
    [InlineData(PivotAggregationFunction.Variance)]
    [InlineData(PivotAggregationFunction.VariancePopulation)]
    [InlineData(PivotAggregationFunction.DistinctCount)]
    public void Rejects_unsupported_data_model_implicit_aggregations_before_binding(
        PivotAggregationFunction aggregation)
    {
        PivotSourceDescriptor source = ModelSource();
        PivotFieldDescriptor field = Field(
            "[Sales].[Amount]",
            PivotFieldAreaSupport.Values,
            PivotFieldDataType.Number);
        PivotLayoutDefinition definition = Definition(
            source,
            new[] { field },
            new[]
            {
                new PivotFieldPlacement(
                    field.Name,
                    PivotFieldArea.Values,
                    1,
                    "Amount",
                    aggregation)
            });
        var adapter = new RecordingAdapter();

        PivotTableNativeMutationValidationException exception =
            Assert.Throws<PivotTableNativeMutationValidationException>(() =>
            new PivotTableNativeLayoutMutationService(
                adapter,
                new PivotMutationCoordinator()).Apply(
                new FakePivot(),
                Context(source, new[] { field }),
                definition));
        Assert.Contains(exception.Issues, issue =>
            issue.Code == "PIVOT_DATA_MODEL_AGGREGATION_UNSUPPORTED");

        Assert.Empty(adapter.Calls);
    }

    [Theory]
    [InlineData(PivotAggregationFunction.Sum, -4157)]
    [InlineData(PivotAggregationFunction.Count, -4112)]
    [InlineData(PivotAggregationFunction.Average, -4106)]
    [InlineData(PivotAggregationFunction.Minimum, -4139)]
    [InlineData(PivotAggregationFunction.Maximum, -4136)]
    public void Allows_each_documented_data_model_implicit_measure_function(
        PivotAggregationFunction aggregation,
        int expectedNativeFunction)
    {
        PivotSourceDescriptor source = ModelSource();
        PivotFieldDescriptor field = Field(
            "[Sales].[Amount]",
            PivotFieldAreaSupport.Values,
            PivotFieldDataType.Number);
        PivotLayoutDefinition definition = Definition(
            source,
            new[] { field },
            new[]
            {
                new PivotFieldPlacement(
                    field.Name,
                    PivotFieldArea.Values,
                    1,
                    "Amount",
                    aggregation)
            });

        NativePivotMutationPlan plan = new PivotTableNativeLayoutMutationService(
            new RecordingAdapter(),
            new PivotMutationCoordinator()).Compile(
            Context(source, new[] { field }),
            definition);

        Assert.Equal(
            expectedNativeFunction,
            Assert.Single(plan.Fields).ConsolidationFunction);
    }

    [Fact]
    public void Rejects_repeated_data_model_implicit_measure_source_function_pairs()
    {
        PivotSourceDescriptor source = ModelSource();
        PivotFieldDescriptor field = Field(
            "[Sales].[Amount]",
            PivotFieldAreaSupport.Values,
            PivotFieldDataType.Number);
        PivotLayoutDefinition duplicate = Definition(
            source,
            new[] { field },
            new[]
            {
                new PivotFieldPlacement(
                    field.Name,
                    PivotFieldArea.Values,
                    1,
                    "Amount average A",
                    PivotAggregationFunction.Average),
                new PivotFieldPlacement(
                    field.Name,
                    PivotFieldArea.Values,
                    2,
                    "Amount average B",
                    PivotAggregationFunction.Average)
            });
        var service = new PivotTableNativeLayoutMutationService(
            new RecordingAdapter(),
            new PivotMutationCoordinator());

        PivotTableNativeMutationValidationException exception =
            Assert.Throws<PivotTableNativeMutationValidationException>(() =>
            service.Compile(Context(source, new[] { field }), duplicate));
        Assert.Contains(exception.Issues, issue =>
            issue.Code == "PIVOT_DATA_MODEL_IMPLICIT_VALUE_DUPLICATE");

        PivotLayoutDefinition distinctFunctions = Definition(
            source,
            new[] { field },
            new[]
            {
                new PivotFieldPlacement(
                    field.Name,
                    PivotFieldArea.Values,
                    1,
                    "Amount sum",
                    PivotAggregationFunction.Sum),
                new PivotFieldPlacement(
                    field.Name,
                    PivotFieldArea.Values,
                    2,
                    "Amount average",
                    PivotAggregationFunction.Average)
            });
        Assert.Equal(
            2,
            service.Compile(
                Context(source, new[] { field }),
                distinctFunctions).Fields.Count);
    }

    [Theory]
    [InlineData(PivotAggregationFunction.Product, -4149)]
    [InlineData(PivotAggregationFunction.CountNumbers, -4113)]
    [InlineData(PivotAggregationFunction.StandardDeviation, -4155)]
    [InlineData(PivotAggregationFunction.StandardDeviationPopulation, -4156)]
    [InlineData(PivotAggregationFunction.Variance, -4164)]
    [InlineData(PivotAggregationFunction.VariancePopulation, -4165)]
    public void Keeps_extended_classic_aggregations_supported(
        PivotAggregationFunction aggregation,
        int expectedNativeFunction)
    {
        PivotSourceDescriptor source = ClassicSource();
        PivotFieldDescriptor[] fields = ClassicFields();
        PivotLayoutDefinition definition = Definition(
            source,
            fields,
            new[]
            {
                new PivotFieldPlacement(
                    "Cost",
                    PivotFieldArea.Values,
                    1,
                    "Cost",
                    aggregation)
            });

        NativePivotMutationPlan plan = new PivotTableNativeLayoutMutationService(
            new RecordingAdapter(),
            new PivotMutationCoordinator()).Compile(
            Context(source, fields),
            definition);

        Assert.Equal(
            expectedNativeFunction,
            Assert.Single(plan.Fields).ConsolidationFunction);
    }

    [Fact]
    public void Allows_external_olap_existing_measure_but_rejects_implicit_aggregation_before_capture()
    {
        PivotSourceDescriptor source = ExternalOlapSource();
        PivotFieldDescriptor measure = Field(
            "[Measures].[Revenue]",
            PivotFieldAreaSupport.Values,
            PivotFieldDataType.Number,
            isMeasure: true);
        var service = new PivotTableNativeLayoutMutationService(
            new RecordingAdapter(),
            new PivotMutationCoordinator());
        PivotLayoutDefinition measureDefinition = Definition(
            source,
            new[] { measure },
            new[]
            {
                new PivotFieldPlacement(
                    measure.Name,
                    PivotFieldArea.Values,
                    1,
                    "Revenue")
            });

        NativePivotMutationPlan measurePlan = service.Compile(
            Context(source, new[] { measure }),
            measureDefinition);
        Assert.True(Assert.Single(measurePlan.Fields).IsMeasure);

        PivotFieldDescriptor fact = Field(
            "[Sales].[Amount]",
            PivotFieldAreaSupport.Values,
            PivotFieldDataType.Number);
        PivotLayoutDefinition invalid = Definition(
            source,
            new[] { fact },
            new[]
            {
                new PivotFieldPlacement(
                    fact.Name,
                    PivotFieldArea.Values,
                    1,
                    "Amount",
                    PivotAggregationFunction.Sum)
            });
        var adapter = new RecordingAdapter();
        var applyingService = new PivotTableNativeLayoutMutationService(
            adapter,
            new PivotMutationCoordinator());

        PivotTableNativeMutationValidationException exception =
            Assert.Throws<PivotTableNativeMutationValidationException>(() => applyingService.Apply(
            new FakePivot(),
            Context(source, new[] { fact }),
            invalid));
        Assert.Contains(exception.Issues, issue =>
            issue.Code == "PIVOT_EXTERNAL_OLAP_VALUE_FIELD_UNSUPPORTED");
        Assert.Empty(adapter.Calls);
    }

    [Fact]
    public void Rejects_missing_native_capability_and_target_drift_before_capture()
    {
        PivotFieldDescriptor[] fields = ClassicFields();
        PivotSourceDescriptor discoveredSource = ClassicSource();
        PivotSourceDescriptor missingRefresh = new PivotSourceDescriptor(
            PivotSourceKind.WorksheetTable,
            "SalesTable",
            PivotCapability.NativeFieldPlacement | PivotCapability.LayoutFormatting);
        PivotLayoutDefinition invalidCapabilities = Definition(
            missingRefresh,
            fields,
            new[] { new PivotFieldPlacement("Region", PivotFieldArea.Row, 1) });
        var adapter = new RecordingAdapter();
        var service = new PivotTableNativeLayoutMutationService(
            adapter,
            new PivotMutationCoordinator());

        PivotTableNativeMutationValidationException capabilityException =
            Assert.Throws<PivotTableNativeMutationValidationException>(() => service.Apply(
            new FakePivot(),
            Context(discoveredSource, fields),
            invalidCapabilities));
        Assert.Contains(capabilityException.Issues, issue =>
            issue.Code == "PIVOT_OPERATION_CAPABILITY_REQUIRED");
        Assert.Empty(adapter.Calls);

        PivotLayoutDefinition wrongTarget = Definition(
            discoveredSource,
            fields,
            new[] { new PivotFieldPlacement("Region", PivotFieldArea.Row, 1) },
            target: new PivotTargetIdentity("workbook_2", "Sheet1", "PivotTable1"));
        Assert.Throws<InvalidOperationException>(() => service.Apply(
            new FakePivot(),
            Context(discoveredSource, fields),
            wrongTarget));
        Assert.Empty(adapter.Calls);
    }

    [Fact]
    public void Rejects_a_live_com_target_mismatch_before_snapshot_or_mutation()
    {
        PivotSourceDescriptor source = ClassicSource();
        PivotFieldDescriptor[] fields = ClassicFields();
        PivotLayoutDefinition definition = Definition(
            source,
            fields,
            new[] { new PivotFieldPlacement("Region", PivotFieldArea.Row, 1) });
        var adapter = new RecordingAdapter
        {
            LiveTarget = new PivotTargetIdentity(
                "workbook_2",
                "Sheet1",
                "PivotTable1")
        };

        Assert.Throws<InvalidOperationException>(() =>
            new PivotTableNativeLayoutMutationService(
                adapter,
                new PivotMutationCoordinator()).Apply(
                new FakePivot(),
                Context(source, fields),
                definition));

        Assert.Equal(new[] { "bind" }, adapter.Calls);
    }

    [Fact]
    public void Rejects_a_live_pivot_cache_source_mismatch_before_snapshot_or_mutation()
    {
        PivotSourceDescriptor source = ClassicSource();
        PivotFieldDescriptor[] fields = ClassicFields();
        PivotLayoutDefinition definition = Definition(
            source,
            fields,
            new[] { new PivotFieldPlacement("Region", PivotFieldArea.Row, 1) });
        var adapter = new RecordingAdapter
        {
            LiveSource = new NativePivotSourceIdentity(
                NativePivotCacheKind.ClassicDatabase,
                "DifferentTable")
        };

        Assert.Throws<InvalidOperationException>(() =>
            new PivotTableNativeLayoutMutationService(
                adapter,
                new PivotMutationCoordinator()).Apply(
                new FakePivot(),
                Context(source, fields),
                definition));

        Assert.Equal(new[] { "bind", "bind-source" }, adapter.Calls);
    }

    [Fact]
    public void Late_bound_target_binding_uses_parent_objects_and_path_free_workbook_identity()
    {
        var workbook = new FakeLateBoundWorkbook();
        var worksheet = new FakeLateBoundWorksheet("Sheet1", workbook);
        var pivot = new FakeLateBoundPivot
        {
            Name = "PivotTable1",
            Parent = worksheet
        };
        var identityResolver = new RecordingWorkbookIdentityResolver("workbook_1");

        PivotTargetIdentity target = new LateBoundPivotTableNativeAdapter().ReadTarget(
            pivot,
            identityResolver);

        Assert.Equal("workbook_1", target.WorkbookId);
        Assert.Equal("Sheet1", target.WorksheetName);
        Assert.Equal("PivotTable1", target.PivotTableName);
        Assert.Same(workbook, identityResolver.Workbook);

        new LateBoundPivotTableNativeAdapter().PersistWorkbookIdentity(
            pivot,
            identityResolver,
            "workbook_1");
        Assert.Equal(1, identityResolver.PersistCalls);
    }

    [Fact]
    public void Native_apply_persists_the_resolved_identity_only_after_capture_succeeds()
    {
        PivotSourceDescriptor source = ClassicSource();
        PivotFieldDescriptor[] fields = ClassicFields();
        PivotLayoutDefinition definition = Definition(
            source,
            fields,
            new[] { new PivotFieldPlacement("Region", PivotFieldArea.Row, 1) });
        var identityResolver = new RecordingWorkbookIdentityResolver("workbook_1");
        var adapter = new RecordingAdapter { ForwardIdentityPersistence = true };
        var service = new PivotTableNativeLayoutMutationService(
            adapter,
            new PivotMutationCoordinator(),
            identityResolver);

        service.Apply(new FakePivot(), Context(source, fields), definition);

        Assert.Equal(1, identityResolver.PersistCalls);
        Assert.Equal("workbook_1", identityResolver.PersistedIdentity);
        Assert.True(adapter.Calls.IndexOf("capture") < adapter.Calls.IndexOf("persist"));

        var failingResolver = new RecordingWorkbookIdentityResolver("workbook_1");
        var failingAdapter = new RecordingAdapter
        {
            ForwardIdentityPersistence = true,
            ThrowOnCapture = true
        };
        var failingService = new PivotTableNativeLayoutMutationService(
            failingAdapter,
            new PivotMutationCoordinator(),
            failingResolver);

        Assert.Throws<InvalidOperationException>(() =>
            failingService.Apply(new FakePivot(), Context(source, fields), definition));
        Assert.Equal(0, failingResolver.PersistCalls);
        Assert.Equal(new[] { "bind", "bind-source", "capture" }, failingAdapter.Calls);
    }

    [Fact]
    public void Native_clear_requires_explicit_clear_all_intent()
    {
        PivotSourceDescriptor source = ClassicSource();
        PivotFieldDescriptor[] fields = ClassicFields();
        var rejectedAdapter = new RecordingAdapter();
        PivotLayoutDefinition omitted = Definition(
            source,
            fields,
            Array.Empty<PivotFieldPlacement>());

        Assert.Throws<PivotTableNativeMutationValidationException>(() =>
            new PivotTableNativeLayoutMutationService(
                rejectedAdapter,
                new PivotMutationCoordinator()).Apply(
                new FakePivot(),
                Context(source, fields),
                omitted));
        Assert.Empty(rejectedAdapter.Calls);

        var clearAdapter = new RecordingAdapter();
        PivotLayoutDefinition explicitClear = Definition(
            source,
            fields,
            Array.Empty<PivotFieldPlacement>(),
            clearAll: true);
        new PivotTableNativeLayoutMutationService(
            clearAdapter,
            new PivotMutationCoordinator()).Apply(
            new FakePivot(),
            Context(source, fields),
            explicitClear);

        Assert.Contains("clear", clearAdapter.Calls);
        Assert.DoesNotContain(clearAdapter.Calls, call => call.StartsWith("place:", StringComparison.Ordinal));
    }

    [Fact]
    public void Late_bound_source_binding_classifies_classic_data_model_and_external_olap()
    {
        var adapter = new LateBoundPivotTableNativeAdapter();
        var classic = new FakeLateBoundPivot
        {
            Cache = new FakeLateBoundPivotCache
            {
                OLAP = false,
                SourceType = 1,
                SourceData = "SalesTable"
            }
        };
        var model = new FakeLateBoundPivot
        {
            Cache = new FakeLateBoundPivotCache
            {
                OLAP = true,
                WorkbookConnection = new FakeLateBoundWorkbookConnection(
                    "ThisWorkbookDataModel",
                    7)
            }
        };
        var external = new FakeLateBoundPivot
        {
            Cache = new FakeLateBoundPivotCache
            {
                OLAP = true,
                WorkbookConnection = new FakeLateBoundWorkbookConnection(
                    "Finance Cube",
                    1)
            }
        };

        NativePivotSourceIdentity classicSource = adapter.ReadSource(classic);
        Assert.Equal(NativePivotCacheKind.ClassicDatabase, classicSource.Kind);
        Assert.Equal("SalesTable", classicSource.SourceName);
        NativePivotSourceIdentity modelSource = adapter.ReadSource(model);
        Assert.Equal(NativePivotCacheKind.DataModel, modelSource.Kind);
        Assert.Equal("ThisWorkbookDataModel", modelSource.SourceName);
        NativePivotSourceIdentity externalSource = adapter.ReadSource(external);
        Assert.Equal(NativePivotCacheKind.ExternalOlap, externalSource.Kind);
        Assert.Equal("Finance Cube", externalSource.SourceName);
    }

    [Fact]
    public void Late_bound_source_binding_fails_closed_for_unreadable_or_unsafe_caches()
    {
        var adapter = new LateBoundPivotTableNativeAdapter();

        Assert.Throws<NotSupportedException>(() => adapter.ReadSource(new object()));

        var unsupportedClassic = new FakeLateBoundPivot
        {
            Cache = new FakeLateBoundPivotCache
            {
                OLAP = false,
                SourceType = 2,
                SourceData = "SalesTable"
            }
        };
        Assert.Throws<NotSupportedException>(() => adapter.ReadSource(unsupportedClassic));

        var unsafeClassic = new FakeLateBoundPivot
        {
            Cache = new FakeLateBoundPivotCache
            {
                OLAP = false,
                SourceType = 1,
                SourceData = "C:\\secret\\sales.xlsx"
            }
        };
        Assert.Throws<NotSupportedException>(() => adapter.ReadSource(unsafeClassic));
    }

    [Fact]
    public void Rejects_member_filter_contract_instead_of_silently_ignoring_it()
    {
        PivotSourceDescriptor source = ClassicSource();
        PivotFieldDescriptor[] fields = ClassicFields();
        var definition = new PivotLayoutDefinition(
            Target(),
            source,
            fields,
            new[] { new PivotFieldPlacement("Region", PivotFieldArea.Row, 1) },
            new[]
            {
                new PivotFieldFilter("Region", PivotFilterMode.Include, new[] { "North" })
            });
        var adapter = new RecordingAdapter();

        Assert.Throws<NotSupportedException>(() =>
            new PivotTableNativeLayoutMutationService(
                adapter,
                new PivotMutationCoordinator()).Apply(
                new FakePivot(),
                Context(source, fields),
                definition));
        Assert.Empty(adapter.Calls);
    }

    [Fact]
    public void Restores_snapshot_when_a_later_native_field_operation_fails()
    {
        PivotSourceDescriptor source = ClassicSource();
        PivotFieldDescriptor[] fields = ClassicFields();
        PivotLayoutDefinition definition = Definition(
            source,
            fields,
            new[]
            {
                new PivotFieldPlacement("Region", PivotFieldArea.Row, 1),
                new PivotFieldPlacement("Month", PivotFieldArea.Column, 1)
            });
        var adapter = new RecordingAdapter { FailAtPlacement = 2 };
        var pivot = new FakePivot();

        PivotMutationException exception = Assert.Throws<PivotMutationException>(() =>
            new PivotTableNativeLayoutMutationService(
                adapter,
                new PivotMutationCoordinator()).Apply(
                pivot,
                Context(source, fields),
                definition));

        Assert.True(exception.RollbackCompleted);
        Assert.Equal("place-column:0001:Month", exception.FailedStep);
        Assert.Equal(
            new[]
            {
                "bind",
                "bind-source",
                "capture",
                "persist",
                "clear",
                "place:row:0001:Region",
                "place:column:0001:Month",
                "restore",
                "refresh"
            },
            adapter.Calls);
        Assert.False(pivot.ManualUpdate);
    }

    [Fact]
    public void Late_bound_capture_rejects_existing_show_values_as_before_layout_clear()
    {
        var pivot = new FakeLateBoundPivot();
        pivot.DataFieldItems.Add(new FakeLateBoundField("Cost")
        {
            Orientation = 4,
            Position = 1,
            Calculation = 3
        });
        var adapter = new LateBoundPivotTableNativeAdapter();

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            adapter.CaptureState(pivot, PivotSourceKind.WorksheetTable));

        Assert.Contains("Show Values As", exception.Message, StringComparison.Ordinal);
        Assert.Equal(4, pivot.DataFieldItems[0].Orientation);
    }

    [Fact]
    public void Late_bound_values_pseudo_field_is_captured_applied_restored_and_verified()
    {
        var pivot = new FakeLateBoundPivot();
        pivot.SourceFields.Add(new FakeLateBoundField("Cost"));
        pivot.SourceFields.Add(new FakeLateBoundField("Units"));
        pivot.DataFieldItems.Add(new FakeLateBoundField("Sum of Cost")
        {
            Caption = "Sum of Cost",
            SourceName = "Cost",
            Orientation = 4,
            Position = 1,
            Function = -4157
        });
        pivot.DataFieldItems.Add(new FakeLateBoundField("Sum of Units")
        {
            Caption = "Sum of Units",
            SourceName = "Units",
            Orientation = 4,
            Position = 2,
            Function = -4157
        });
        pivot.DataPivotField = new FakeLateBoundField("Values")
        {
            Orientation = 2,
            Position = 1
        };
        var adapter = new LateBoundPivotTableNativeAdapter();

        object snapshot = adapter.CaptureState(pivot, PivotSourceKind.WorksheetTable);
        adapter.ApplyLayout(pivot, new NativePivotLayoutCommand
        {
            RowAxisLayout = 0,
            ValuesAxis = PivotValuesAxis.Rows,
            ValuesPosition = 1,
            ShowRowGrandTotals = true,
            ShowColumnGrandTotals = true,
            ShowFieldHeaders = true,
            PreserveFormatting = true
        });

        Assert.Equal(1, pivot.DataPivotField.Orientation);
        Assert.Equal(1, pivot.DataPivotField.Position);

        adapter.RestoreState(pivot, snapshot);

        Assert.Equal(2, pivot.DataPivotField.Orientation);
        Assert.Equal(1, pivot.DataPivotField.Position);
        Assert.Equal(2, pivot.DataFields.Count);

        var plan = new NativePivotMutationPlan
        {
            SourceKind = PivotSourceKind.WorksheetTable,
            Fields = new[]
            {
                new NativePivotFieldCommand
                {
                    InstanceId = "value:0001:Sum of Cost",
                    FieldName = "Cost",
                    Caption = "Sum of Cost",
                    SetCaption = true,
                    Area = PivotFieldArea.Values,
                    Position = 1,
                    ConsolidationFunction = -4157
                },
                new NativePivotFieldCommand
                {
                    InstanceId = "value:0002:Sum of Units",
                    FieldName = "Units",
                    Caption = "Sum of Units",
                    SetCaption = true,
                    Area = PivotFieldArea.Values,
                    Position = 2,
                    ConsolidationFunction = -4157
                }
            },
            Layout = new NativePivotLayoutCommand
            {
                RowAxisLayout = 0,
                ValuesAxis = PivotValuesAxis.Columns,
                ValuesPosition = 1,
                ShowRowGrandTotals = true,
                ShowColumnGrandTotals = true,
                ShowFieldHeaders = true,
                PreserveFormatting = true
            }
        };
        adapter.Verify(pivot, plan);

        pivot.DataPivotField.Position = 2;
        Assert.Throws<InvalidOperationException>(() => adapter.Verify(pivot, plan));
    }

    [Fact]
    public void Late_bound_automatic_single_value_leaves_excel_owned_values_axis_untouched()
    {
        var pivot = new FakeLateBoundPivot();
        pivot.SourceFields.Add(new FakeLateBoundField("Cost"));
        pivot.DataFieldItems.Add(new FakeLateBoundField("Sum of Cost")
        {
            Caption = "Sum of Cost",
            SourceName = "Cost",
            Orientation = 4,
            Position = 1,
            Function = -4157
        });
        pivot.DataPivotField = new FakeLateBoundField("Values")
        {
            // Real Excel can expose a non-hidden host value here while still
            // rejecting an explicit xlHidden write for a one-value pivot.
            Orientation = 2,
            Position = 1
        };
        var adapter = new LateBoundPivotTableNativeAdapter();
        var layout = new NativePivotLayoutCommand
        {
            RowAxisLayout = 0,
            ValuesAxis = PivotValuesAxis.Automatic,
            ValuesPosition = 1,
            ShowRowGrandTotals = true,
            ShowColumnGrandTotals = true,
            ShowFieldHeaders = true,
            PreserveFormatting = true
        };

        adapter.ApplyLayout(pivot, layout);

        Assert.Equal(2, pivot.DataPivotField.Orientation);
        Assert.Equal(1, pivot.DataPivotField.Position);
        adapter.Verify(pivot, new NativePivotMutationPlan
        {
            SourceKind = PivotSourceKind.WorksheetTable,
            Fields = new[]
            {
                new NativePivotFieldCommand
                {
                    InstanceId = "value:0001:Sum of Cost",
                    FieldName = "Cost",
                    Caption = "Sum of Cost",
                    SetCaption = true,
                    Area = PivotFieldArea.Values,
                    Position = 1,
                    ConsolidationFunction = -4157
                }
            },
            Layout = layout
        });
    }

    [Fact]
    public void Late_bound_capture_rejects_active_native_filters_and_calculated_objects()
    {
        var filteredPivot = new FakeLateBoundPivot();
        filteredPivot.SourceFields.Add(new FakeLateBoundField("Region")
        {
            Orientation = 1,
            Position = 1,
            AllItemsVisible = false
        });
        var adapter = new LateBoundPivotTableNativeAdapter();

        NotSupportedException filterException = Assert.Throws<NotSupportedException>(() =>
            adapter.CaptureState(filteredPivot, PivotSourceKind.WorksheetTable));
        Assert.Contains("member filter", filterException.Message, StringComparison.Ordinal);

        var calculatedPivot = new FakeLateBoundPivot { CalculatedFieldCount = 1 };
        NotSupportedException calculatedException = Assert.Throws<NotSupportedException>(() =>
            adapter.CaptureState(calculatedPivot, PivotSourceKind.WorksheetTable));
        Assert.Contains("calculated fields", calculatedException.Message, StringComparison.Ordinal);

        var calculatedItemPivot = new FakeLateBoundPivot();
        calculatedItemPivot.SourceFields.Add(new FakeLateBoundField("Region")
        {
            Orientation = 1,
            Position = 1,
            CalculatedItemCount = 1
        });
        NotSupportedException itemException = Assert.Throws<NotSupportedException>(() =>
            adapter.CaptureState(calculatedItemPivot, PivotSourceKind.WorksheetTable));
        Assert.Contains("calculated items", itemException.Message, StringComparison.Ordinal);

        var unreadableFields = new FakeLateBoundPivot { ThrowOnCalculatedFieldsRead = true };
        InvalidOperationException unreadableFieldsException = Assert.Throws<InvalidOperationException>(() =>
            adapter.CaptureState(unreadableFields, PivotSourceKind.WorksheetTable));
        Assert.Contains("calculated-field collection", unreadableFieldsException.Message, StringComparison.Ordinal);

        var unreadableItems = new FakeLateBoundPivot();
        unreadableItems.SourceFields.Add(new FakeLateBoundField("Region")
        {
            Orientation = 1,
            Position = 1,
            ThrowOnCalculatedItemsRead = true
        });
        object unreadableItemsSnapshot = adapter.CaptureState(
            unreadableItems,
            PivotSourceKind.WorksheetTable);
        Assert.NotNull(unreadableItemsSnapshot);
    }

    [Fact]
    public void Late_bound_olap_capture_reads_manual_filter_state_from_the_cube_field()
    {
        var pivot = new FakeLateBoundPivot();
        var cube = new FakeLateBoundCubeField("[Sales].[Region]");
        pivot.SourceFields.Add(new FakeLateBoundField("[Sales].[Region]")
        {
            Orientation = 1,
            Position = 1,
            CubeField = cube,
            ThrowOnAllItemsVisibleRead = true
        });
        var adapter = new LateBoundPivotTableNativeAdapter();

        Assert.NotNull(adapter.CaptureState(pivot, PivotSourceKind.DataModel));

        cube.AllItemsVisible = false;
        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            adapter.CaptureState(pivot, PivotSourceKind.DataModel));
        Assert.Contains("member filter", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Late_bound_capture_rejects_mixed_per_row_repeat_label_state()
    {
        var pivot = new FakeLateBoundPivot();
        pivot.SourceFields.Add(new FakeLateBoundField("Region")
        {
            Orientation = 1,
            Position = 1,
            RepeatLabels = true
        });
        pivot.SourceFields.Add(new FakeLateBoundField("Department")
        {
            Orientation = 1,
            Position = 2,
            RepeatLabels = false
        });

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            new LateBoundPivotTableNativeAdapter().CaptureState(
                pivot,
                PivotSourceKind.WorksheetTable));

        Assert.Contains("mixed per-row RepeatLabels", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Late_bound_capture_fails_closed_when_a_required_collection_count_is_unreadable()
    {
        var adapter = new LateBoundPivotTableNativeAdapter();

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            adapter.CaptureState(
                new UnreadableRowCountPivot(),
                PivotSourceKind.WorksheetTable));

        Assert.Contains("collection count", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Late_bound_clear_fails_closed_when_a_required_collection_item_is_unreadable()
    {
        var adapter = new LateBoundPivotTableNativeAdapter();

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            adapter.ClearLayout(
                new UnreadableDataItemPivot(),
                PivotSourceKind.WorksheetTable));

        Assert.Contains("required DataFields collection", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Late_bound_clear_fails_closed_when_a_required_collection_is_missing()
    {
        var adapter = new LateBoundPivotTableNativeAdapter();

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            adapter.ClearLayout(new object(), PivotSourceKind.WorksheetTable));

        Assert.Contains("required DataFields collection", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Late_bound_restore_round_trips_all_twelve_indexed_subtotal_slots()
    {
        var pivot = new FakeLateBoundPivot();
        var region = new FakeLateBoundField("Region")
        {
            Orientation = 1,
            Position = 1
        };
        bool[] original =
        {
            false, true, false, true, false, true,
            true, false, true, false, true, false
        };
        for (var index = 1; index <= original.Length; index++)
        {
            region.Subtotals[index] = original[index - 1];
        }

        pivot.SourceFields.Add(region);
        var adapter = new LateBoundPivotTableNativeAdapter();
        object snapshot = adapter.CaptureState(pivot, PivotSourceKind.WorksheetTable);

        adapter.ClearLayout(pivot, PivotSourceKind.WorksheetTable);
        for (var index = 1; index <= original.Length; index++)
        {
            region.Subtotals[index] = false;
        }

        adapter.RestoreState(pivot, snapshot);

        Assert.Equal(1, region.Orientation);
        Assert.Equal(
            original,
            Enumerable.Range(1, 12).Select(index => region.Subtotals[index]).ToArray());
    }

    [Fact]
    public void Late_bound_data_model_rollback_deletes_cube_fields_created_after_capture()
    {
        var pivot = new FakeLateBoundPivot();
        var existing = new FakeLateBoundCubeField(
            "[Measures].[Existing]",
            pivot.CubeFieldItems);
        existing.Caption = "Original caption";
        pivot.CubeFieldItems.Add(existing);
        var adapter = new LateBoundPivotTableNativeAdapter();
        object snapshot = adapter.CaptureState(pivot, PivotSourceKind.DataModel);
        existing.Caption = "Caption overwritten by GetMeasure";
        var created = new FakeLateBoundCubeField(
            "[Measures].[Created by failed mutation]",
            pivot.CubeFieldItems);
        pivot.CubeFieldItems.Add(created);

        adapter.RestoreState(pivot, snapshot);

        Assert.True(created.Deleted);
        Assert.Equal(new[] { existing }, pivot.CubeFieldItems);
        Assert.Equal("Original caption", existing.Caption);
    }

    [Fact]
    public void Late_bound_capture_fails_closed_for_unreadable_rollback_state()
    {
        var adapter = new LateBoundPivotTableNativeAdapter();
        var unreadableLayout = new FakeLateBoundPivot
        {
            ThrowOnRowGrandRead = true
        };
        Assert.Throws<InvalidOperationException>(() => adapter.CaptureState(
            unreadableLayout,
            PivotSourceKind.WorksheetTable));

        var unreadablePosition = new FakeLateBoundPivot();
        unreadablePosition.SourceFields.Add(new FakeLateBoundField("Region")
        {
            Orientation = 1,
            Position = 1,
            ThrowOnPositionRead = true
        });
        Assert.Throws<InvalidOperationException>(() => adapter.CaptureState(
            unreadablePosition,
            PivotSourceKind.WorksheetTable));

        var unreadableFunction = PivotWithSingleDataField();
        unreadableFunction.DataFieldItems[0].ThrowOnFunctionRead = true;
        Assert.Throws<InvalidOperationException>(() => adapter.CaptureState(
            unreadableFunction,
            PivotSourceKind.WorksheetTable));

        var unreadableNumberFormat = PivotWithSingleDataField();
        unreadableNumberFormat.DataFieldItems[0].ThrowOnNumberFormatRead = true;
        Assert.Throws<InvalidOperationException>(() => adapter.CaptureState(
            unreadableNumberFormat,
            PivotSourceKind.WorksheetTable));

        var unreadableCubeCaption = new FakeLateBoundPivot();
        unreadableCubeCaption.CubeFieldItems.Add(new FakeLateBoundCubeField(
            "[Measures].[Existing]")
        {
            ThrowOnCaptionRead = true
        });
        Assert.Throws<InvalidOperationException>(() => adapter.CaptureState(
            unreadableCubeCaption,
            PivotSourceKind.DataModel));
    }

    [Fact]
    public void Late_bound_static_fake_ignores_the_special_values_axis_field()
    {
        var pivot = new FakeLateBoundPivot();
        var region = new FakeLateBoundField("Region")
        {
            Orientation = 1,
            Position = 1
        };
        var valuesAxis = new FakeLateBoundField("Values")
        {
            Orientation = 1,
            Position = 2
        };
        pivot.SourceFields.Add(region);
        pivot.SourceFields.Add(valuesAxis);
        pivot.DataPivotField = valuesAxis;
        var adapter = new LateBoundPivotTableNativeAdapter();

        Assert.NotNull(adapter.CaptureState(pivot, PivotSourceKind.WorksheetTable));
        adapter.ClearLayout(pivot, PivotSourceKind.WorksheetTable);

        Assert.Equal(0, region.Orientation);
        Assert.Equal(0, valuesAxis.Orientation);
        // RCW identity and Excel's automatic DataPivotField placement still need a real-host smoke test.
    }

    [Fact]
    public void Late_bound_verify_requires_exact_area_counts_and_explicit_captions()
    {
        var pivot = MatchingVerifyPivot(out FakeLateBoundField region);
        NativePivotMutationPlan plan = MatchingVerifyPlan();
        var adapter = new LateBoundPivotTableNativeAdapter();

        adapter.Verify(pivot, plan);

        region.Caption = "Wrong caption";
        Assert.Throws<InvalidOperationException>(() => adapter.Verify(pivot, plan));

        region.Caption = "Region caption";
        pivot.SourceFields.Add(new FakeLateBoundField("Extra")
        {
            Orientation = 1,
            Position = 2
        });
        Assert.Throws<InvalidOperationException>(() => adapter.Verify(pivot, plan));
    }

    [Fact]
    public void Late_bound_verify_requires_the_exact_twelve_slot_row_subtotal_state()
    {
        var pivot = MatchingVerifyPivot(out FakeLateBoundField region);
        NativePivotMutationPlan plan = MatchingVerifyPlan();
        var adapter = new LateBoundPivotTableNativeAdapter();

        adapter.Verify(pivot, plan);

        region.Subtotals[12] = true;
        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            adapter.Verify(pivot, plan));
        Assert.Contains("12-slot", exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(PivotAggregationFunction.Sum, -4157)]
    [InlineData(PivotAggregationFunction.Count, -4112)]
    [InlineData(PivotAggregationFunction.Average, -4106)]
    [InlineData(PivotAggregationFunction.Minimum, -4139)]
    [InlineData(PivotAggregationFunction.Maximum, -4136)]
    [InlineData(PivotAggregationFunction.Product, -4149)]
    [InlineData(PivotAggregationFunction.CountNumbers, -4113)]
    [InlineData(PivotAggregationFunction.StandardDeviation, -4155)]
    [InlineData(PivotAggregationFunction.StandardDeviationPopulation, -4156)]
    [InlineData(PivotAggregationFunction.Variance, -4164)]
    [InlineData(PivotAggregationFunction.VariancePopulation, -4165)]
    [InlineData(PivotAggregationFunction.DistinctCount, 11)]
    public void Maps_every_core_aggregation_to_the_native_excel_constant(
        PivotAggregationFunction aggregation,
        int expected)
    {
        Assert.Equal(
            expected,
            PivotTableNativeLayoutMutationService.ConsolidationFunction(aggregation));
    }

    private static FakeLateBoundPivot MatchingVerifyPivot(
        out FakeLateBoundField region)
    {
        var pivot = new FakeLateBoundPivot();
        region = new FakeLateBoundField("Region")
        {
            Caption = "Region caption",
            Orientation = 1,
            Position = 1
        };
        region.Subtotals[1] = true;
        pivot.SourceFields.Add(region);
        return pivot;
    }

    private static FakeLateBoundPivot PivotWithSingleDataField()
    {
        var pivot = new FakeLateBoundPivot();
        pivot.DataFieldItems.Add(new FakeLateBoundField("Cost")
        {
            Orientation = 4,
            Position = 1,
            Function = -4157,
            NumberFormat = "#,##0"
        });
        return pivot;
    }

    private static NativePivotMutationPlan MatchingVerifyPlan()
    {
        return new NativePivotMutationPlan
        {
            SourceKind = PivotSourceKind.WorksheetTable,
            Fields = new[]
            {
                new NativePivotFieldCommand
                {
                    InstanceId = "row:0001:Region",
                    FieldName = "Region",
                    Caption = "Region caption",
                    SetCaption = true,
                    Area = PivotFieldArea.Row,
                    Position = 1,
                    SubtotalMode = PivotSubtotalMode.Automatic
                }
            },
            Layout = new NativePivotLayoutCommand
            {
                RowAxisLayout = 0,
                RepeatItemLabels = false,
                ShowRowGrandTotals = true,
                ShowColumnGrandTotals = true,
                ShowFieldHeaders = true,
                PreserveFormatting = true,
                ShowRowStripes = false,
                ShowColumnStripes = false
            }
        };
    }

    private static PivotTableContext Context(
        PivotSourceDescriptor source,
        IReadOnlyList<PivotFieldDescriptor> fields)
    {
        return new PivotTableContext(
            Definition(source, fields, Array.Empty<PivotFieldPlacement>()),
            isConnected: true,
            sourceFieldsComplete: true);
    }

    private static PivotLayoutDefinition Definition(
        PivotSourceDescriptor source,
        IEnumerable<PivotFieldDescriptor> fields,
        IEnumerable<PivotFieldPlacement> placements,
        PivotLayoutMetadata? layout = null,
        PivotFormatMetadata? format = null,
        PivotTargetIdentity? target = null,
        bool clearAll = false)
    {
        PivotFieldPlacement[] placementSnapshot = placements.ToArray();
        PivotLayoutMetadata resolvedLayout = layout ?? new PivotLayoutMetadata(
            valuesAxis: placementSnapshot.Count(item => item.Area == PivotFieldArea.Values) > 1
                ? PivotValuesAxis.Columns
                : PivotValuesAxis.Automatic,
            valuesPosition: 1);
        return new PivotLayoutDefinition(
            target ?? Target(),
            source,
            fields,
            placementSnapshot,
            layout: resolvedLayout,
            format: format,
            clearAll: clearAll);
    }

    private static PivotTargetIdentity Target()
    {
        return new PivotTargetIdentity("workbook_1", "Sheet1", "PivotTable1");
    }

    private static PivotSourceDescriptor ClassicSource()
    {
        return new PivotSourceDescriptor(
            PivotSourceKind.WorksheetTable,
            "SalesTable",
            PivotCapability.NativeFieldPlacement |
            PivotCapability.MemberFiltering |
            PivotCapability.LayoutFormatting |
            PivotCapability.ShowValuesAs |
            PivotCapability.Refresh |
            PivotCapability.UpgradeToDataModel);
    }

    private static PivotSourceDescriptor ModelSource()
    {
        return new PivotSourceDescriptor(
            PivotSourceKind.DataModel,
            "ThisWorkbookDataModel",
            PivotCapability.NativeFieldPlacement |
            PivotCapability.MemberFiltering |
            PivotCapability.LayoutFormatting |
            PivotCapability.ShowValuesAs |
            PivotCapability.Refresh |
            PivotCapability.DistinctCount |
            PivotCapability.DataModel |
            PivotCapability.ModelMeasures,
            "Sales");
    }

    private static PivotSourceDescriptor ExternalOlapSource()
    {
        return new PivotSourceDescriptor(
            PivotSourceKind.ExternalOlap,
            "Finance Cube",
            PivotCapability.NativeFieldPlacement |
            PivotCapability.MemberFiltering |
            PivotCapability.LayoutFormatting |
            PivotCapability.ShowValuesAs |
            PivotCapability.Refresh |
            PivotCapability.CalculatedMembers |
            PivotCapability.NamedSets);
    }

    private static PivotFieldDescriptor[] ClassicFields()
    {
        return new[]
        {
            Field("Region", PivotFieldAreaSupport.Row | PivotFieldAreaSupport.Filter),
            Field("Month", PivotFieldAreaSupport.Column),
            Field("Category", PivotFieldAreaSupport.Filter),
            Field("Cost", PivotFieldAreaSupport.Values, PivotFieldDataType.Number)
        };
    }

    private static PivotFieldDescriptor Field(
        string name,
        PivotFieldAreaSupport areas,
        PivotFieldDataType dataType = PivotFieldDataType.Text,
        bool isMeasure = false)
    {
        return new PivotFieldDescriptor(
            name,
            name,
            dataType,
            areas,
            isMeasure: isMeasure);
    }

    public sealed class FakePivot
    {
        public bool ManualUpdate { get; set; }
    }

    public sealed class FakeLateBoundWorkbook
    {
    }

    public sealed class FakeLateBoundWorkbookConnection
    {
        public FakeLateBoundWorkbookConnection(string name, int type)
        {
            Name = name;
            Type = type;
        }

        public string Name { get; }

        public int Type { get; }
    }

    public sealed class FakeLateBoundPivotCache
    {
        public bool OLAP { get; set; }

        public int SourceType { get; set; } = 1;

        public object SourceData { get; set; } = "SalesTable";

        public FakeLateBoundWorkbookConnection WorkbookConnection { get; set; } =
            new FakeLateBoundWorkbookConnection("ThisWorkbookDataModel", 7);
    }

    public sealed class FakeLateBoundWorksheet
    {
        public FakeLateBoundWorksheet(string name, FakeLateBoundWorkbook parent)
        {
            Name = name;
            Parent = parent;
        }

        public string Name { get; }

        public FakeLateBoundWorkbook Parent { get; }
    }

    public sealed class FakeLateBoundPivot
    {
        private bool rowGrand = true;

        public string Name { get; set; } = "PivotTable1";

        public FakeLateBoundWorksheet Parent { get; set; } =
            new FakeLateBoundWorksheet("Sheet1", new FakeLateBoundWorkbook());

        public FakeLateBoundPivotCache Cache { get; set; } = new();

        public FakeLateBoundPivotCache PivotCache()
        {
            return Cache;
        }

        public List<FakeLateBoundField> SourceFields { get; } = new();

        public List<FakeLateBoundField> DataFieldItems { get; } = new();

        public List<FakeLateBoundCubeField> CubeFieldItems { get; } = new();

        public FakeLateBoundField? DataPivotField { get; set; }

        public int CalculatedFieldCount { get; set; }

        public bool ThrowOnCalculatedFieldsRead { get; set; }

        public FakeLateBoundCollection PivotFields => new FakeLateBoundCollection(SourceFields);

        public FakeLateBoundCollection RowFields => new FakeLateBoundCollection(
            SourceFields.Where(field => field.Orientation == 1)
                .Concat(DataPivotField != null && DataPivotField.Orientation == 1
                    ? new[] { DataPivotField! }
                    : Array.Empty<FakeLateBoundField>()));

        public FakeLateBoundCollection ColumnFields => new FakeLateBoundCollection(
            SourceFields.Where(field => field.Orientation == 2)
                .Concat(DataPivotField != null && DataPivotField.Orientation == 2
                    ? new[] { DataPivotField! }
                    : Array.Empty<FakeLateBoundField>()));

        public FakeLateBoundCollection PageFields => new FakeLateBoundCollection(
            SourceFields.Where(field => field.Orientation == 3));

        public FakeLateBoundCollection DataFields => new FakeLateBoundCollection(
            DataFieldItems.Where(field => field.Orientation == 4));

        public FakeLateBoundCubeCollection CubeFields =>
            new FakeLateBoundCubeCollection(CubeFieldItems);

        public int LayoutRowDefault { get; private set; }

        public bool RowGrand
        {
            get
            {
                if (ThrowOnRowGrandRead)
                {
                    throw new InvalidOperationException("RowGrand unavailable");
                }

                return rowGrand;
            }
            set => rowGrand = value;
        }

        public bool ThrowOnRowGrandRead { get; set; }

        public bool ColumnGrand { get; set; } = true;

        public bool DisplayFieldCaptions { get; set; } = true;

        public string TableStyle2 { get; set; } = string.Empty;

        public bool PreserveFormatting { get; set; } = true;

        public bool ShowTableStyleRowStripes { get; set; }

        public bool ShowTableStyleColumnStripes { get; set; }

        public FakeLateBoundCollection CalculatedFields()
        {
            if (ThrowOnCalculatedFieldsRead)
            {
                throw new InvalidOperationException("CalculatedFields unavailable");
            }

            return new FakeLateBoundCollection(
                Enumerable.Range(0, CalculatedFieldCount)
                    .Select(index => new FakeLateBoundField("Calculated" + index)));
        }

        public FakeLateBoundField AddDataField(
            FakeLateBoundField sourceField,
            string caption,
            int function)
        {
            var dataField = new FakeLateBoundField(caption)
            {
                Caption = caption,
                SourceName = sourceField.SourceName,
                Orientation = 4,
                Position = DataFields.Count + 1,
                Function = function
            };
            DataFieldItems.Add(dataField);
            return dataField;
        }

        public void RowAxisLayout(int layout)
        {
            LayoutRowDefault = layout;
        }

        public void RepeatAllLabels(int mode)
        {
            foreach (FakeLateBoundField field in RowFields.Items)
            {
                field.RepeatLabels = mode == 2;
            }
        }
    }

    public sealed class UnreadableRowCountPivot
    {
        public ThrowingCountCollection RowFields { get; } = new();

        public FakeLateBoundCollection CalculatedFields()
        {
            return new FakeLateBoundCollection(Array.Empty<FakeLateBoundField>());
        }
    }

    public sealed class UnreadableDataItemPivot
    {
        public ThrowingItemCollection DataFields { get; } = new();

        public FakeLateBoundCollection PageFields { get; } =
            new FakeLateBoundCollection(Array.Empty<FakeLateBoundField>());

        public FakeLateBoundCollection ColumnFields { get; } =
            new FakeLateBoundCollection(Array.Empty<FakeLateBoundField>());

        public FakeLateBoundCollection RowFields { get; } =
            new FakeLateBoundCollection(Array.Empty<FakeLateBoundField>());
    }

    public sealed class ThrowingCountCollection
    {
        public int Count => throw new InvalidOperationException("count unavailable");

        public object Item(int index)
        {
            throw new InvalidOperationException("item unavailable");
        }
    }

    public sealed class ThrowingItemCollection
    {
        public int Count => 1;

        public object Item(int index)
        {
            throw new InvalidOperationException("item unavailable");
        }
    }

    public sealed class FakeLateBoundCollection
    {
        public FakeLateBoundCollection(IEnumerable<FakeLateBoundField> items)
        {
            Items = items.ToList();
        }

        public IReadOnlyList<FakeLateBoundField> Items { get; }

        public int Count => Items.Count;

        public FakeLateBoundField Item(object key)
        {
            if (key is int index)
            {
                return Items[index - 1];
            }

            string name = Convert.ToString(key) ?? string.Empty;
            return Items.Single(field =>
                string.Equals(field.Name, name, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(field.Caption, name, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(field.SourceName, name, StringComparison.OrdinalIgnoreCase));
        }
    }

    public sealed class FakeLateBoundCubeCollection
    {
        public FakeLateBoundCubeCollection(IEnumerable<FakeLateBoundCubeField> items)
        {
            Items = items.ToList();
        }

        public IReadOnlyList<FakeLateBoundCubeField> Items { get; }

        public int Count => Items.Count;

        public FakeLateBoundCubeField Item(object key)
        {
            if (key is int index)
            {
                return Items[index - 1];
            }

            string name = Convert.ToString(key) ?? string.Empty;
            return Items.Single(field =>
                string.Equals(field.Name, name, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(field.Caption, name, StringComparison.OrdinalIgnoreCase));
        }
    }

    public sealed class FakeLateBoundField
    {
        private bool allItemsVisible = true;
        private int position;
        private int function = -4157;
        private string numberFormat = "General";

        public FakeLateBoundField(string name)
        {
            Name = name;
            Caption = name;
            SourceName = name;
        }

        public string Name { get; set; }

        public string Caption { get; set; }

        public string SourceName { get; set; }

        public int Orientation { get; set; }

        public int Position
        {
            get
            {
                if (ThrowOnPositionRead)
                {
                    throw new InvalidOperationException("Position unavailable");
                }

                return position;
            }
            set => position = value;
        }

        public bool ThrowOnPositionRead { get; set; }

        public int Function
        {
            get
            {
                if (ThrowOnFunctionRead)
                {
                    throw new InvalidOperationException("Function unavailable");
                }

                return function;
            }
            set => function = value;
        }

        public bool ThrowOnFunctionRead { get; set; }

        public int Calculation { get; set; } = -4143;

        public string NumberFormat
        {
            get
            {
                if (ThrowOnNumberFormatRead)
                {
                    throw new InvalidOperationException("NumberFormat unavailable");
                }

                return numberFormat;
            }
            set => numberFormat = value;
        }

        public bool ThrowOnNumberFormatRead { get; set; }

        public bool RepeatLabels { get; set; }

        public bool AllItemsVisible
        {
            get
            {
                if (ThrowOnAllItemsVisibleRead)
                {
                    throw new InvalidOperationException(
                        "AllItemsVisible is unavailable on an OLAP PivotField.");
                }

                return allItemsVisible;
            }
            set => allItemsVisible = value;
        }

        public bool ThrowOnAllItemsVisibleRead { get; set; }

        public FakeLateBoundCubeField? CubeField { get; set; }

        public bool EnableMultiplePageItems { get; set; }

        public FakeLateBoundCollection PivotFilters { get; } =
            new FakeLateBoundCollection(Array.Empty<FakeLateBoundField>());

        public FakeIndexedSubtotals Subtotals { get; } = new();

        public int CalculatedItemCount { get; set; }

        public bool ThrowOnCalculatedItemsRead { get; set; }

        public FakeLateBoundCollection CalculatedItems()
        {
            if (ThrowOnCalculatedItemsRead)
            {
                throw new InvalidOperationException("CalculatedItems unavailable");
            }

            return new FakeLateBoundCollection(
                Enumerable.Range(0, CalculatedItemCount)
                    .Select(index => new FakeLateBoundField("CalculatedItem" + index)));
        }
    }

    public sealed class FakeLateBoundCubeField
    {
        private readonly ICollection<FakeLateBoundCubeField>? owner;
        private string caption;

        public FakeLateBoundCubeField(
            string name,
            ICollection<FakeLateBoundCubeField>? owner = null)
        {
            Name = name;
            caption = name;
            this.owner = owner;
        }

        public string Name { get; set; }

        public string Caption
        {
            get
            {
                if (ThrowOnCaptionRead)
                {
                    throw new InvalidOperationException("Caption unavailable");
                }

                return caption;
            }
            set => caption = value;
        }

        public bool ThrowOnCaptionRead { get; set; }

        public bool AllItemsVisible { get; set; } = true;

        public bool EnableMultiplePageItems { get; set; }

        public bool Deleted { get; private set; }

        public void Delete()
        {
            Deleted = true;
            owner?.Remove(this);
        }
    }

    public sealed class FakeIndexedSubtotals
    {
        private readonly bool[] values = new bool[12];

        public bool this[int oneBasedIndex]
        {
            get => values[oneBasedIndex - 1];
            set => values[oneBasedIndex - 1] = value;
        }
    }

    private sealed class RecordingAdapter : IPivotTableNativeAdapter
    {
        private int placementCount;

        public List<string> Calls { get; } = new();

        public List<NativePivotFieldCommand> Placed { get; } = new();

        public NativePivotLayoutCommand? Layout { get; private set; }

        public int? FailAtPlacement { get; set; }

        public bool ThrowOnCapture { get; set; }

        public bool ForwardIdentityPersistence { get; set; }

        public object IdentityWorkbook { get; } = new object();

        public PivotTargetIdentity LiveTarget { get; set; } = Target();

        public NativePivotSourceIdentity LiveSource { get; set; } =
            new NativePivotSourceIdentity(
                NativePivotCacheKind.ClassicDatabase,
                "SalesTable");

        public PivotTargetIdentity ReadTarget(
            object pivotTable,
            IWorkbookIdentityResolver workbookIdentityResolver)
        {
            Calls.Add("bind");
            return LiveTarget;
        }

        public NativePivotSourceIdentity ReadSource(object pivotTable)
        {
            Calls.Add("bind-source");
            return LiveSource;
        }

        public void PersistWorkbookIdentity(
            object pivotTable,
            IWorkbookIdentityResolver workbookIdentityResolver,
            string expectedWorkbookId)
        {
            Calls.Add("persist");
            if (ForwardIdentityPersistence)
            {
                workbookIdentityResolver.Persist(IdentityWorkbook, expectedWorkbookId);
            }
        }

        public object CaptureState(object pivotTable, PivotSourceKind sourceKind)
        {
            Calls.Add("capture");
            if (ThrowOnCapture)
            {
                throw new InvalidOperationException("native capture failed");
            }

            return new object();
        }

        public void ClearLayout(object pivotTable, PivotSourceKind sourceKind)
        {
            Calls.Add("clear");
        }

        public void RemoveFieldsNotInPlan(
            object pivotTable,
            PivotSourceKind sourceKind,
            IReadOnlyList<NativePivotFieldCommand> desiredFields)
        {
            Calls.Add("clear");
        }

        public void PlaceField(
            object pivotTable,
            PivotSourceKind sourceKind,
            NativePivotFieldCommand command)
        {
            Calls.Add("place:" + command.InstanceId);
            placementCount++;
            if (placementCount == FailAtPlacement)
            {
                throw new InvalidOperationException("native placement failed");
            }

            Placed.Add(command);
        }

        public void ApplyLayout(object pivotTable, NativePivotLayoutCommand command)
        {
            Calls.Add("layout");
            Layout = command;
        }

        public void RestoreState(object pivotTable, object snapshot)
        {
            Calls.Add("restore");
        }

        public void Refresh(object pivotTable)
        {
            Calls.Add("refresh");
        }

        public void Verify(object pivotTable, NativePivotMutationPlan plan)
        {
            Calls.Add("verify");
        }
    }

    private sealed class RecordingWorkbookIdentityResolver : IWorkbookIdentityResolver
    {
        private readonly string identity;

        public RecordingWorkbookIdentityResolver(string identity)
        {
            this.identity = identity;
        }

        public object? Workbook { get; private set; }

        public int ResolveCalls { get; private set; }

        public int PersistCalls { get; private set; }

        public string? PersistedIdentity { get; private set; }

        public string Resolve(object workbook)
        {
            Workbook = workbook;
            ResolveCalls++;
            return identity;
        }

        public void Persist(object workbook, string expectedWorkbookId)
        {
            Workbook = workbook;
            PersistCalls++;
            PersistedIdentity = expectedWorkbookId;
            if (!string.Equals(identity, expectedWorkbookId, StringComparison.Ordinal))
            {
                throw new InvalidOperationException("identity mismatch");
            }
        }
    }
}
