using ExcelReportBuilder.Excel.Execution;
using ExcelReportBuilder.Excel.Ownership;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class ManagedPivotCachePolicyTests
{
    [Fact]
    public void Source_contract_round_trips_and_matches_supported_backends()
    {
        var worksheet = Contract(CanonicalBackend.Worksheet, "ManagedCanonical");
        var worksheetRoundTrip = PivotCacheSourceContract.Parse(worksheet.Serialized);
        var model = Contract(CanonicalBackend.DataModel, "ManagedModel");
        var modelRoundTrip = PivotCacheSourceContract.Parse(model.Serialized);

        Assert.Equal(worksheet, worksheetRoundTrip);
        Assert.True(worksheetRoundTrip.Matches(new PivotCacheSnapshot
        {
            SourceType = 1,
            WorksheetSource = "='Managed source'!ManagedCanonical[#All]"
        }));
        Assert.False(worksheetRoundTrip.Matches(new PivotCacheSnapshot
        {
            SourceType = 2,
            ConnectionName = "ManagedCanonical"
        }));
        Assert.Equal(model, modelRoundTrip);
        Assert.True(modelRoundTrip.Matches(new PivotCacheSnapshot
        {
            SourceType = 2,
            ConnectionName = "managedmodel"
        }));
        Assert.False(modelRoundTrip.Matches(new PivotCacheSnapshot
        {
            SourceType = 1,
            WorksheetSource = "ManagedModel"
        }));
    }

    [Theory]
    [InlineData("")]
    [InlineData("W|x:Source")]
    [InlineData("X|6:Source")]
    [InlineData("W|5:Source")]
    public void Source_contract_rejects_invalid_serialized_values(string value)
    {
        Assert.Throws<InvalidOperationException>(() => PivotCacheSourceContract.Parse(value));
    }

    [Theory]
    [InlineData("=$A$1:$C$8", "A1:C8")]
    [InlineData("='Managed source'!$A$1:$C$8", "Managed source!A1:C8")]
    [InlineData("='Managed source'!R1C1:R8C3", "'Managed source'!R1C1:R8C3")]
    public void Worksheet_source_references_are_compared_without_excel_decoration(
        string left,
        string right)
    {
        Assert.True(NativePivotTableExecutor.SourceReferencesEqual(left, right));
    }

    [Theory]
    [InlineData("='Managed source'!$A$1:$C$8", "'Managed source'!$A$1:$F$200")]
    [InlineData("R1C1:R8C3", "R1C1:R200C6")]
    public void Worksheet_source_range_can_resize_without_losing_the_managed_table_contract(
        string prior,
        string current)
    {
        Assert.True(NativePivotTableExecutor.SourceReferenceStartsEqual(prior, current));
    }

    [Fact]
    public void Exact_exclusive_registered_cache_is_reused_after_managed_pivot_is_cleared()
    {
        var source = Contract(CanonicalBackend.Worksheet, "ManagedCanonical");

        var plan = ManagedPivotCachePolicy.Plan(
            new[] { Registration(source, 4) },
            Slot(source).Identity,
            Slot(source).RegistryName,
            source,
            managedPivotExists: false,
            Snapshot(4, source, pivotTableCount: 0));

        Assert.Equal(ManagedPivotCacheAction.Reuse, plan.Action);
    }

    [Fact]
    public void Removed_then_readded_block_reuses_its_retained_exclusive_cache_slot()
    {
        var source = Contract(CanonicalBackend.Worksheet, "ManagedCanonical");
        var retainedRecords = new[] { Registration(source, 6) };

        ManagedPivotCachePlan readded = ManagedPivotCachePolicy.Plan(
            retainedRecords,
            Slot(source).Identity,
            Slot(source).RegistryName,
            source,
            managedPivotExists: false,
            Snapshot(6, source, pivotTableCount: 0));

        Assert.Equal(ManagedPivotCacheAction.Reuse, readded.Action);
        Assert.Single(retainedRecords);
    }

    [Fact]
    public void Exact_exclusive_registered_cache_is_reused_with_its_managed_pivot()
    {
        var source = Contract(CanonicalBackend.DataModel, "ManagedModel");

        var plan = ManagedPivotCachePolicy.Plan(
            new[] { Registration(source, 7) },
            Slot(source).Identity,
            Slot(source).RegistryName,
            source,
            managedPivotExists: true,
            Snapshot(7, source, pivotTableCount: 1));

        Assert.Equal(ManagedPivotCacheAction.Reuse, plan.Action);
    }

    [Fact]
    public void Compatible_legacy_cache_is_reused_only_when_exact_managed_pivot_exclusively_identifies_it()
    {
        var source = Contract(CanonicalBackend.Worksheet, "ManagedCanonical");

        var plan = ManagedPivotCachePolicy.Plan(
            Array.Empty<ManagedObjectRecord>(),
            Slot(source).Identity,
            Slot(source).RegistryName,
            source,
            managedPivotExists: true,
            Snapshot(2, source, pivotTableCount: 1));

        Assert.Equal(ManagedPivotCacheAction.ReuseAndRegister, plan.Action);
    }

    [Fact]
    public void Changed_source_contract_retires_only_the_exact_backend_slot()
    {
        var prior = Contract(CanonicalBackend.DataModel, "ManagedModel");
        var requested = Contract(CanonicalBackend.DataModel, "ReplacementModel");

        var plan = ManagedPivotCachePolicy.Plan(
            new[] { Registration(prior, 3) },
            Slot(prior).Identity,
            Slot(prior).RegistryName,
            requested,
            managedPivotExists: false,
            Snapshot(3, prior, pivotTableCount: 0));

        Assert.Equal(ManagedPivotCacheAction.RetireAndCreate, plan.Action);
        Assert.Contains("source changed", plan.Reason, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Worksheet_to_model_to_worksheet_reuses_one_cache_slot_per_backend()
    {
        var worksheet = Contract(CanonicalBackend.Worksheet, "ManagedCanonical");
        var model = Contract(CanonicalBackend.DataModel, "ManagedModel");
        var records = new List<ManagedObjectRecord>();

        ManagedPivotCachePlan firstWorksheet = ManagedPivotCachePolicy.Plan(
            records,
            Slot(worksheet).Identity,
            Slot(worksheet).RegistryName,
            worksheet,
            managedPivotExists: false,
            candidate: null);
        Assert.Equal(ManagedPivotCacheAction.Create, firstWorksheet.Action);
        records.Add(Registration(worksheet, 1));

        ManagedPivotCachePlan firstModel = ManagedPivotCachePolicy.Plan(
            records,
            Slot(model).Identity,
            Slot(model).RegistryName,
            model,
            managedPivotExists: false,
            candidate: null);
        Assert.Equal(ManagedPivotCacheAction.Create, firstModel.Action);
        records.Add(Registration(model, 2));

        ManagedPivotCachePlan secondWorksheet = ManagedPivotCachePolicy.Plan(
            records,
            Slot(worksheet).Identity,
            Slot(worksheet).RegistryName,
            worksheet,
            managedPivotExists: false,
            Snapshot(1, worksheet, pivotTableCount: 0));

        Assert.Equal(ManagedPivotCacheAction.Reuse, secondWorksheet.Action);
        Assert.Equal(2, records.Count);
        Assert.Equal(2, records.Select(record => record.Locator).Distinct().Count());
    }

    [Theory]
    [InlineData(false, 1)]
    [InlineData(true, 2)]
    public void Shared_cache_is_never_reused(bool managedPivotExists, int pivotTableCount)
    {
        var source = Contract(CanonicalBackend.Worksheet, "ManagedCanonical");

        var plan = ManagedPivotCachePolicy.Plan(
            new[] { Registration(source, 3) },
            Slot(source).Identity,
            Slot(source).RegistryName,
            source,
            managedPivotExists,
            Snapshot(3, source, pivotTableCount));

        Assert.Equal(ManagedPivotCacheAction.RetireAndCreate, plan.Action);
        Assert.Contains("shared", plan.Reason, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Missing_orphan_cache_registration_is_retired_but_missing_managed_pivot_cache_fails_closed()
    {
        var source = Contract(CanonicalBackend.Worksheet, "ManagedCanonical");
        var records = new[] { Registration(source, 3) };

        var orphanPlan = ManagedPivotCachePolicy.Plan(
            records,
            Slot(source).Identity,
            Slot(source).RegistryName,
            source,
            managedPivotExists: false,
            candidate: null);
        var exception = Assert.Throws<InvalidOperationException>(() =>
            ManagedPivotCachePolicy.Plan(
                records,
                Slot(source).Identity,
                Slot(source).RegistryName,
                source,
                managedPivotExists: true,
                candidate: null));

        Assert.Equal(ManagedPivotCacheAction.RetireAndCreate, orphanPlan.Action);
        Assert.Contains("cannot be located", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Cache_name_claimed_by_another_identity_fails_closed()
    {
        var source = Contract(CanonicalBackend.Worksheet, "ManagedCanonical");
        var other = Registration(source, 2);
        other.ObjectId = "another_cache_worksheet";

        var exception = Assert.Throws<InvalidOperationException>(() =>
            ManagedPivotCachePolicy.Plan(
                new[] { other },
                Slot(source).Identity,
                Slot(source).RegistryName,
                source,
                managedPivotExists: false,
                candidate: null));

        Assert.Contains("already owned", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("locator")]
    [InlineData("registered-source")]
    [InlineData("managed-name")]
    public void Tampered_registration_or_cache_contract_fails_closed(string scenario)
    {
        var source = Contract(CanonicalBackend.Worksheet, "ManagedCanonical");
        var record = Registration(source, 3);
        var snapshot = Snapshot(3, source, pivotTableCount: 0);
        if (scenario == "locator")
        {
            snapshot.Index = 4;
        }
        else if (scenario == "registered-source")
        {
            snapshot.WorksheetSource = "AnotherTable";
        }
        else
        {
            record.ExcelName = "AnotherManagedCache";
        }

        Assert.Throws<InvalidOperationException>(() =>
            ManagedPivotCachePolicy.Plan(
                new[] { record },
                Slot(source).Identity,
                Slot(source).RegistryName,
                source,
                managedPivotExists: false,
                snapshot));
    }

    [Fact]
    public void Exact_existing_pivot_contract_accepts_only_its_registered_live_cache()
    {
        var source = Contract(CanonicalBackend.Worksheet, "ManagedCanonical");
        var cacheRegistration = Registration(source, 3);
        var pivotRegistration = PivotRegistration(source, 3);

        NativePivotTableExecutor.DemandExactExistingPivotContract(
            pivotRegistration,
            "ManagedPivot",
            cacheRegistration,
            Slot(source),
            3,
            Snapshot(3, source, pivotTableCount: 1));
    }

    [Fact]
    public void Prior_worksheet_pivot_is_validated_when_requested_rebuild_slot_is_data_model()
    {
        var priorSource = Contract(CanonicalBackend.Worksheet, "ManagedCanonical");
        var requestedSource = Contract(CanonicalBackend.DataModel, "ManagedModel");
        var priorRegistration = Registration(priorSource, 3);
        var slots = new[] { Slot(priorSource), Slot(requestedSource) };

        RegisteredPivotCacheBinding binding =
            NativePivotTableExecutor.ResolveExactLivePivotCacheBinding(
                new[] { priorRegistration },
                slots,
                livePivotCacheIndex: 3);

        Assert.Equal(Slot(priorSource).Identity.ObjectId, binding.Slot.Identity.ObjectId);
        Assert.NotEqual(Slot(requestedSource).Identity.ObjectId, binding.Slot.Identity.ObjectId);
        NativePivotTableExecutor.DemandExactExistingPivotContract(
            PivotRegistration(priorSource, 3),
            "ManagedPivot",
            binding.Registration,
            binding.Slot,
            3,
            Snapshot(3, priorSource, pivotTableCount: 1));
    }

    [Theory]
    [InlineData("no-slot")]
    [InlineData("mismatched-index")]
    [InlineData("ambiguous-slots")]
    public void Replacement_pivot_without_one_exact_prior_backend_slot_is_rejected(
        string scenario)
    {
        var worksheetSource = Contract(CanonicalBackend.Worksheet, "ManagedCanonical");
        var modelSource = Contract(CanonicalBackend.DataModel, "ManagedModel");
        var records = new List<ManagedObjectRecord>();
        var liveIndex = 8;
        if (scenario == "mismatched-index")
        {
            records.Add(Registration(worksheetSource, 3));
        }
        else if (scenario == "ambiguous-slots")
        {
            records.Add(Registration(worksheetSource, 8));
            records.Add(Registration(modelSource, 8));
        }

        Assert.Throws<InvalidOperationException>(() =>
            NativePivotTableExecutor.ResolveExactLivePivotCacheBinding(
                records,
                new[] { Slot(worksheetSource), Slot(modelSource) },
                liveIndex));
    }

    [Theory]
    [InlineData("replacement-cache")]
    [InlineData("live-cache-source")]
    [InlineData("registered-cache-slot")]
    [InlineData("pivot-registration")]
    public void Same_named_replacement_pivot_contract_is_rejected_before_range_mutation(
        string scenario)
    {
        var source = Contract(CanonicalBackend.Worksheet, "ManagedCanonical");
        var cacheRegistration = Registration(source, 3);
        var pivotRegistration = PivotRegistration(source, 3);
        var snapshot = Snapshot(3, source, pivotTableCount: 1);
        var livePivotCacheIndex = 3;
        switch (scenario)
        {
            case "replacement-cache":
                livePivotCacheIndex = 8;
                break;
            case "live-cache-source":
                snapshot.WorksheetSource = "UnmanagedTable";
                break;
            case "registered-cache-slot":
                cacheRegistration.ExcelName = "Another cache slot";
                break;
            case "pivot-registration":
                pivotRegistration.SourceContract =
                    Contract(CanonicalBackend.Worksheet, "AnotherTable").Serialized;
                break;
        }

        Assert.Throws<InvalidOperationException>(() =>
            NativePivotTableExecutor.DemandExactExistingPivotContract(
                pivotRegistration,
                "ManagedPivot",
                cacheRegistration,
                Slot(source),
                livePivotCacheIndex,
                snapshot));
    }

    private static ManagedPivotCacheSlot Slot(PivotCacheSourceContract source)
    {
        return ManagedPivotCacheSlot.For(
            "report",
            "block",
            "ManagedCache",
            source.Backend);
    }

    private static ManagedObjectRecord Registration(
        PivotCacheSourceContract source,
        int index)
    {
        ManagedPivotCacheSlot slot = Slot(source);
        return new ManagedObjectRecord
        {
            ReportId = slot.Identity.ReportId,
            ObjectId = slot.Identity.ObjectId,
            Kind = ManagedObjectKind.PivotCache,
            ExcelName = slot.RegistryName,
            Locator = index.ToString(System.Globalization.CultureInfo.InvariantCulture),
            SourceContract = source.Serialized
        };
    }

    private static ManagedObjectRecord PivotRegistration(
        PivotCacheSourceContract source,
        int index)
    {
        return new ManagedObjectRecord
        {
            ReportId = "report",
            ObjectId = "block",
            Kind = ManagedObjectKind.PivotTable,
            ExcelName = "ManagedPivot",
            Locator = index.ToString(System.Globalization.CultureInfo.InvariantCulture),
            SourceContract = source.Serialized
        };
    }

    private static PivotCacheSourceContract Contract(
        CanonicalBackend backend,
        string sourceName)
    {
        return PivotCacheSourceContract.From(new CanonicalLoadPlan
        {
            Backend = backend,
            TableOrConnectionName = sourceName
        });
    }

    private static PivotCacheSnapshot Snapshot(
        int index,
        PivotCacheSourceContract source,
        int pivotTableCount)
    {
        return new PivotCacheSnapshot
        {
            Index = index,
            SourceType = source.Backend == CanonicalBackend.Worksheet ? 1 : 2,
            WorksheetSource = source.Backend == CanonicalBackend.Worksheet
                ? source.SourceName
                : null,
            ConnectionName = source.Backend == CanonicalBackend.DataModel
                ? source.SourceName
                : null,
            PivotTableCount = pivotTableCount
        };
    }
}
