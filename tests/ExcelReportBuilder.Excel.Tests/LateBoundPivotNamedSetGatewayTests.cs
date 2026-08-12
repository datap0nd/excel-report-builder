using ExcelReportBuilder.Core.PivotPlus;
using ExcelReportBuilder.Core.PivotPlus.Calculations;
using ExcelReportBuilder.Core.PivotPlus.NamedSets;
using ExcelReportBuilder.Excel.PivotPlus;
using ExcelReportBuilder.Excel.PivotPlus.NamedSets;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class LateBoundPivotNamedSetGatewayTests
{
    [Fact]
    public void Binds_only_the_exact_selected_data_model_pivot()
    {
        var fixture = new HostFixture();

        BoundPivotNamedSetTarget target = fixture.Bind();

        Assert.Same(fixture.Workbook, target.Workbook);
        Assert.Same(fixture.Pivot, target.PivotTable);
        Assert.Same(fixture.Model, target.Model);
        Assert.Same(fixture.ModelConnection, target.DataModelConnection);

        Assert.Throws<NotSupportedException>(() => fixture.Gateway.Bind(
            fixture.Workbook,
            fixture.Pivot,
            fixture.Context(PivotSourceKind.ExternalOlap)));
        fixture.Pivot.Cache.WorkbookConnection = new FakeConnection("Other", 7);
        Assert.Throws<NotSupportedException>(() => fixture.Bind());
    }

    [Fact]
    public void Discovers_only_existing_cube_fields_pivot_fields_and_pivot_items()
    {
        var fixture = new HostFixture();
        FakeCubeField region = fixture.Pivot.AddHierarchy(
            "[Sales].[Region]",
            "Region");
        FakePivotField level = region.AddLevel(
            "[Sales].[Region].[Region]",
            "Region");
        level.AddItem("[Sales].[Region].&[North]", "North");
        level.AddItem("[Sales].[Region].&[South]", "South");

        PivotNamedSetSchemaDiscoveryResult result = fixture.Gateway.DiscoverSchema(
            fixture.Bind());

        Assert.Equal(PivotNamedSetProviderKind.DataModel, result.Schema.ProviderKind);
        Assert.StartsWith("pivot.source.v1:sha256:", result.Schema.SourceFingerprint);
        PivotNamedSetHierarchySchema hierarchy = Assert.Single(result.Schema.Hierarchies);
        Assert.Equal("[Sales].[Region]", hierarchy.ProviderUniqueName);
        Assert.Null(hierarchy.AllMemberId);
        PivotNamedSetLevelSchema discoveredLevel = Assert.Single(hierarchy.Levels);
        Assert.True(discoveredLevel.MembersComplete);
        Assert.Equal(2, discoveredLevel.Members.Count);
        Assert.All(discoveredLevel.Members, member =>
        {
            Assert.False(member.IsAllMember);
            Assert.Null(member.ParentMemberId);
            Assert.Matches("^[A-Za-z0-9._-]+$", member.Id);
        });
        Assert.Empty(result.Diagnostics);
        Assert.Equal(0, region.CreatePivotFieldsCalls);
        Assert.Equal(0, fixture.Pivot.Cache.AdoConnectionReads);
        Assert.Equal(0, fixture.Workbook.ActiveCellReads);
    }

    [Fact]
    public void Marks_unmaterialized_levels_and_members_incomplete_without_inventing_items()
    {
        var fixture = new HostFixture();
        FakeCubeField missingLevels = fixture.Pivot.AddHierarchy(
            "[Sales].[Department]",
            "Department");
        missingLevels.PivotFieldsAvailable = false;
        FakeCubeField region = fixture.Pivot.AddHierarchy(
            "[Sales].[Region]",
            "Region");
        FakePivotField missingItems = region.AddLevel(
            "[Sales].[Region].[Region]",
            "Region");
        missingItems.PivotItemsAvailable = false;

        PivotNamedSetSchemaDiscoveryResult result = fixture.Gateway.DiscoverSchema(
            fixture.Bind());

        Assert.Equal(2, result.Diagnostics.Count);
        Assert.Contains(result.Diagnostics,
            diagnostic => diagnostic.Code == "PIVOT_SET_DISCOVERY_PIVOTFIELDS_UNAVAILABLE");
        Assert.Contains(result.Diagnostics,
            diagnostic => diagnostic.Code == "PIVOT_SET_DISCOVERY_PIVOTITEMS_UNAVAILABLE");
        Assert.Empty(result.Schema.Hierarchies.Single(hierarchy =>
            hierarchy.ProviderUniqueName == "[Sales].[Department]").Levels);
        Assert.False(result.Schema.Hierarchies.Single(hierarchy =>
                hierarchy.ProviderUniqueName == "[Sales].[Region]")
            .Levels.Single().MembersComplete);
    }

    [Fact]
    public void Discovery_fails_instead_of_truncating_an_over_limit_collection()
    {
        var fixture = new HostFixture();
        fixture.Pivot.CubeFields.CountOverride = 4097;

        Assert.Throws<NotSupportedException>(() => fixture.Gateway.DiscoverSchema(
            fixture.Bind()));
    }

    [Fact]
    public void Capture_pairs_complete_and_orphaned_host_objects_with_fingerprints()
    {
        var fixture = new HostFixture();
        fixture.Pivot.AddExistingSet(
            "[PivotTablePlus_complete]",
            "{[Sales].[Region].DefaultMember}",
            "Complete",
            "PivotTable+|set|owned");
        fixture.Pivot.AddCalculatedOnly(
            "[PivotTablePlus_calculated_only]",
            "{[Sales].[Region].DefaultMember}",
            "PivotTable+|set|orphan");
        fixture.Pivot.AddCubeOnly("[PivotTablePlus_cube_only]", "Cube only");

        PivotNamedSetWorkbookSnapshot snapshot = fixture.Gateway.Capture(fixture.Bind());

        Assert.Equal(3, snapshot.SelectedPivot.Artifacts.Count);
        Assert.Contains(snapshot.SelectedPivot.Artifacts,
            artifact => artifact.PairState == PivotNamedSetPairState.Complete &&
                        artifact.Name == "[PivotTablePlus_complete]");
        Assert.Contains(snapshot.SelectedPivot.Artifacts,
            artifact => artifact.PairState == PivotNamedSetPairState.CalculatedMemberOnly);
        Assert.Contains(snapshot.SelectedPivot.Artifacts,
            artifact => artifact.PairState == PivotNamedSetPairState.CubeFieldOnly);
        Assert.All(snapshot.SelectedPivot.Artifacts,
            artifact => Assert.StartsWith("namedset.host.v1:sha256:", artifact.LiveFingerprint));
        Assert.StartsWith("namedset.pivot.v1:sha256:", snapshot.SelectedPivot.Fingerprint);
    }

    [Fact]
    public void Pivot_fingerprint_covers_the_complete_calculated_member_inventory()
    {
        var fixture = new HostFixture();
        string before = fixture.Gateway.Capture(fixture.Bind()).SelectedPivot.Fingerprint;
        fixture.Pivot.AddCalculatedMember(
            "User Calculation",
            0,
            "{[Measures].[Sales]}");

        string after = fixture.Gateway.Capture(fixture.Bind()).SelectedPivot.Fingerprint;

        Assert.NotEqual(before, after);
    }

    [Fact]
    public void Capture_rejects_an_ambiguous_exact_target_occurrence()
    {
        var fixture = new HostFixture();
        BoundPivotNamedSetTarget target = fixture.Bind();
        fixture.Sheet.PivotItems.Add(fixture.Pivot);

        InvalidOperationException failure = Assert.Throws<InvalidOperationException>(() =>
            fixture.Gateway.Capture(target));

        Assert.Contains("exactly once", failure.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Capture_fails_closed_when_a_formula_read_throws()
    {
        var fixture = new HostFixture();
        FakeCalculatedMember member = fixture.Pivot.AddCalculatedMember(
            "User Calculation",
            0,
            "{[Sales].[Region].DefaultMember}");
        member.ThrowOnFormulaRead = true;

        InvalidOperationException failure = Assert.Throws<InvalidOperationException>(() =>
            fixture.Gateway.Capture(fixture.Bind()));

        Assert.Contains("formula", failure.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Empty(fixture.Pivot.OperationLog);
    }

    [Fact]
    public void Creates_with_exact_seven_argument_add_then_add_set_and_validates_readback()
    {
        var fixture = new HostFixture();
        DesiredPivotNamedSet desired = fixture.Desired();

        LivePivotNamedSetSnapshot created = fixture.Gateway.CreateSet(
            fixture.Bind(),
            desired);

        FakeCalculatedMemberAddCall add = Assert.Single(
            fixture.Pivot.CalculatedMembers.AddCalls);
        Assert.Equal(desired.Name, add.Name);
        Assert.Equal(PivotNamedSetFormulaTransport.EncodeForExcel(desired.RawMdx), add.Formula);
        Assert.Same(Type.Missing, add.SolveOrder);
        Assert.Equal(1, add.Type);
        Assert.Equal(desired.Dynamic, add.Dynamic);
        Assert.Equal(desired.DisplayFolderMarker, add.DisplayFolder);
        Assert.False(add.HierarchizeDistinct);
        Assert.Equal(
            new[]
            {
                "CalculatedMembers.Add",
                "CubeFields.AddSet",
                "PivotCache.MakeConnection"
            },
            fixture.Pivot.OperationLog.Take(3));
        Assert.True(fixture.Pivot.Cache.MakeConnectionCalls >= 2);
        Assert.Equal(desired.RawMdx, created.RawFormula);
        Assert.True(created.IsValid);
    }

    [Fact]
    public void Rejects_a_stale_core_source_fingerprint_before_touching_the_host()
    {
        var fixture = new HostFixture();
        DesiredPivotNamedSet desired = fixture.Desired();
        fixture.Pivot.AddHierarchy("[Sales].[Later]", "Later");

        InvalidOperationException failure = Assert.Throws<InvalidOperationException>(() =>
            fixture.Gateway.CreateSet(fixture.Bind(), desired));

        Assert.Contains("schema", failure.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Empty(fixture.Pivot.CalculatedMembers.AddCalls);
        Assert.Empty(fixture.Pivot.OperationLog);
    }

    [Fact]
    public void Rejects_stale_source_before_replace_delete_or_restore_host_touch()
    {
        var replace = new HostFixture();
        replace.Gateway.CreateSet(replace.Bind(), replace.Desired());
        LivePivotNamedSetSnapshot replaceBefore = replace.Gateway.Capture(replace.Bind())
            .SelectedPivot.Artifacts.Single();
        DesiredPivotNamedSet replacement = replace.Desired(caption: "Changed Rows");
        replace.Pivot.AddHierarchy("[Sales].[Later]", "Later");
        replace.Pivot.OperationLog.Clear();
        Assert.Throws<InvalidOperationException>(() => replace.Gateway.ReplaceSet(
            replace.Bind(),
            replaceBefore,
            replacement));
        Assert.DoesNotContain(replace.Pivot.OperationLog, value =>
            value.EndsWith(".Delete", StringComparison.Ordinal));

        var delete = new HostFixture();
        delete.Gateway.CreateSet(delete.Bind(), delete.Desired());
        LivePivotNamedSetSnapshot deleteBefore = delete.Gateway.Capture(delete.Bind())
            .SelectedPivot.Artifacts.Single();
        delete.Pivot.AddHierarchy("[Sales].[Later]", "Later");
        delete.Pivot.OperationLog.Clear();
        Assert.Throws<InvalidOperationException>(() =>
            delete.Gateway.DeleteSet(delete.Bind(), deleteBefore));
        Assert.DoesNotContain(delete.Pivot.OperationLog, value =>
            value.EndsWith(".Delete", StringComparison.Ordinal));

        var restore = new HostFixture();
        restore.Gateway.CreateSet(restore.Bind(), restore.Desired());
        LivePivotNamedSetSnapshot restoreBefore = restore.Gateway.Capture(restore.Bind())
            .SelectedPivot.Artifacts.Single();
        restore.Pivot.CubeItems.Single(field => field.CubeFieldType == 3).Delete();
        restore.Pivot.AddHierarchy("[Sales].[Later]", "Later");
        restore.Pivot.OperationLog.Clear();
        Assert.Throws<InvalidOperationException>(() =>
            restore.Gateway.RestoreSet(restore.Bind(), restoreBefore));
        Assert.DoesNotContain("CubeFields.AddSet", restore.Pivot.OperationLog);
    }

    [Fact]
    public void Verifies_exact_compiled_measure_dependency_before_named_set_mutation()
    {
        var fixture = new HostFixture();
        string sourceFingerprint = fixture.Gateway.DiscoverSchema(fixture.Bind())
            .Schema.SourceFingerprint;
        (PivotMdxCompilation compilation, PivotDaxCompilation dax) =
            PivotNamedSetCanonicalTests.CompileTopNSet(sourceFingerprint);
        DesiredPivotNamedSet desired = Assert.Single(
            PivotNamedSetCompilationAdapter.CreateDesired("setup_1", compilation));
        DesiredPivotNamedSetMeasureDependency dependency = Assert.Single(
            desired.DirectMeasureDependencies);
        OwnedPivotMeasureDefinition measure = Assert.Single(dax.Measures);
        fixture.Model.ModelMeasureItems.Add(new FakeModelMeasure(
            dependency.GeneratedMeasureName,
            measure.DaxFormula,
            dependency.ExpectedDescriptionMarker));

        LivePivotNamedSetSnapshot created = fixture.Gateway.CreateSet(
            fixture.Bind(),
            desired);

        Assert.Equal(desired.Name, created.Name);

        var stale = new HostFixture();
        string staleSource = stale.Gateway.DiscoverSchema(stale.Bind())
            .Schema.SourceFingerprint;
        (PivotMdxCompilation staleCompilation, PivotDaxCompilation staleDax) =
            PivotNamedSetCanonicalTests.CompileTopNSet(staleSource);
        DesiredPivotNamedSet staleDesired = Assert.Single(
            PivotNamedSetCompilationAdapter.CreateDesired("setup_1", staleCompilation));
        DesiredPivotNamedSetMeasureDependency staleDependency = Assert.Single(
            staleDesired.DirectMeasureDependencies);
        stale.Model.ModelMeasureItems.Add(new FakeModelMeasure(
            staleDependency.GeneratedMeasureName,
            Assert.Single(staleDax.Measures).DaxFormula + " + 0",
            staleDependency.ExpectedDescriptionMarker));

        Assert.Throws<InvalidOperationException>(() =>
            stale.Gateway.CreateSet(stale.Bind(), staleDesired));
        Assert.Empty(stale.Pivot.CalculatedMembers.AddCalls);
        Assert.Empty(stale.Pivot.OperationLog);
    }

    [Fact]
    public void Reconciles_a_calculated_member_add_that_throws_after_commit()
    {
        var fixture = new HostFixture();
        fixture.Pivot.CalculatedMembers.AddFailure = FakeFailure.AfterCommitOnce;

        LivePivotNamedSetSnapshot created = fixture.Gateway.CreateSet(
            fixture.Bind(),
            fixture.Desired());

        Assert.Equal(PivotNamedSetPairState.Complete, created.PairState);
        Assert.Single(fixture.Pivot.CalculatedMembers.AddCalls);
        Assert.Single(fixture.Pivot.CalculatedItems);
        Assert.Single(fixture.Pivot.CubeItems, field => field.CubeFieldType == 3);
    }

    [Fact]
    public void Reconciles_add_set_that_throws_before_commit_without_duplicate_add()
    {
        var fixture = new HostFixture();
        fixture.Pivot.CubeFields.AddFailure = FakeFailure.BeforeCommitOnce;

        LivePivotNamedSetSnapshot created = fixture.Gateway.CreateSet(
            fixture.Bind(),
            fixture.Desired());

        Assert.Equal(PivotNamedSetPairState.Complete, created.PairState);
        Assert.Single(fixture.Pivot.CalculatedMembers.AddCalls);
        Assert.Equal(2, fixture.Pivot.OperationLog.Count(value =>
            value == "CubeFields.AddSet"));
    }

    [Fact]
    public void Reconciles_a_transient_make_connection_failure_after_add_set()
    {
        var fixture = new HostFixture();
        fixture.Pivot.Cache.MakeConnectionFailuresRemaining = 1;

        LivePivotNamedSetSnapshot created = fixture.Gateway.CreateSet(
            fixture.Bind(),
            fixture.Desired());

        Assert.True(created.IsValid == true);
        Assert.Single(fixture.Pivot.CalculatedMembers.AddCalls);
        Assert.True(fixture.Pivot.Cache.MakeConnectionCalls >= 2);
    }

    [Fact]
    public void Reconciles_a_transient_pair_property_failure()
    {
        var fixture = new HostFixture();
        DesiredPivotNamedSet desired = fixture.Desired(flattenHierarchies: true);
        fixture.Pivot.CalculatedFlattenSetFailuresRemaining = 1;

        LivePivotNamedSetSnapshot created = fixture.Gateway.CreateSet(
            fixture.Bind(),
            desired);

        Assert.True(created.CalculatedMemberFlattenHierarchies == true);
        Assert.True(created.CubeFieldFlattenHierarchies == true);
        Assert.Single(fixture.Pivot.CalculatedMembers.AddCalls);
    }

    [Fact]
    public void Reconciles_a_transient_exact_formula_read_failure()
    {
        var fixture = new HostFixture();
        fixture.Pivot.NewCalculatedMemberFormulaReadFailureOnRead = 2;

        LivePivotNamedSetSnapshot created = fixture.Gateway.CreateSet(
            fixture.Bind(),
            fixture.Desired());

        Assert.True(created.IsValid == true);
        Assert.Single(fixture.Pivot.CalculatedMembers.AddCalls);
    }

    [Fact]
    public void Cleanup_failure_is_recoverable_by_the_same_plan_retry()
    {
        var fixture = new HostFixture();
        DesiredPivotNamedSet desired = fixture.Desired();
        fixture.Pivot.Cache.KeepCalculatedMembersInvalid = true;
        fixture.Pivot.NewCalculatedMemberDeleteFailure = FakeFailure.BeforeCommitOnce;

        Assert.Throws<PivotNamedSetRecoveryRequiredException>(() =>
            fixture.Gateway.CreateSet(fixture.Bind(), desired));
        Assert.Single(fixture.Pivot.CalculatedItems);
        Assert.DoesNotContain(fixture.Pivot.CubeItems, field => field.CubeFieldType == 3);

        fixture.Pivot.Cache.KeepCalculatedMembersInvalid = false;
        LivePivotNamedSetSnapshot recovered = fixture.Gateway.CreateSet(
            fixture.Bind(),
            desired);

        Assert.True(recovered.IsValid == true);
        Assert.Single(fixture.Pivot.CalculatedItems);
        Assert.Single(fixture.Pivot.CubeItems, field => field.CubeFieldType == 3);
    }

    [Fact]
    public void Retries_once_when_add_fails_before_commit()
    {
        var fixture = new HostFixture();
        fixture.Pivot.CalculatedMembers.AddFailure = FakeFailure.BeforeCommitOnce;

        LivePivotNamedSetSnapshot created = fixture.Gateway.CreateSet(
            fixture.Bind(),
            fixture.Desired());

        Assert.Equal(PivotNamedSetPairState.Complete, created.PairState);
        Assert.Equal(2, fixture.Pivot.CalculatedMembers.AddCalls.Count);
        Assert.Single(fixture.Pivot.CalculatedItems);
    }

    [Fact]
    public void Reconciles_an_exact_add_set_commit_even_when_com_throws_afterward()
    {
        var fixture = new HostFixture();
        fixture.Pivot.CubeFields.AddFailure = FakeFailure.AfterCommitOnce;
        fixture.Pivot.CubeFields.MarkCalculatedMemberValidAfterAddSet = true;

        LivePivotNamedSetSnapshot created = fixture.Gateway.CreateSet(
            fixture.Bind(),
            fixture.Desired());

        Assert.Equal(PivotNamedSetPairState.Complete, created.PairState);
        Assert.Single(fixture.Pivot.CalculatedMembers.AddCalls);
        Assert.True(fixture.Pivot.Cache.MakeConnectionCalls >= 1);
    }

    [Fact]
    public void Reconciles_add_set_after_commit_when_first_connection_refresh_fails()
    {
        var fixture = new HostFixture();
        fixture.Pivot.CubeFields.AddFailure = FakeFailure.AfterCommitOnce;
        fixture.Pivot.Cache.MakeConnectionFailuresRemaining = 1;

        LivePivotNamedSetSnapshot created = fixture.Gateway.CreateSet(
            fixture.Bind(),
            fixture.Desired());

        Assert.True(created.IsValid == true);
        Assert.Single(fixture.Pivot.CalculatedMembers.AddCalls);
        Assert.Single(fixture.Pivot.CubeItems, field => field.CubeFieldType == 3);
    }

    [Fact]
    public void Fails_recovery_when_excel_rewrites_compiled_formula_text()
    {
        var fixture = new HostFixture();
        fixture.Pivot.FormulaReadbackTransform = encoded =>
            PivotNamedSetFormulaTransport.DecodeRequired(encoded) + " ";

        Assert.Throws<PivotNamedSetRecoveryRequiredException>(() =>
            fixture.Gateway.CreateSet(fixture.Bind(), fixture.Desired()));
    }

    [Fact]
    public void Blocks_unowned_name_collisions_sibling_use_and_formula_references()
    {
        var collision = new HostFixture();
        DesiredPivotNamedSet desired = collision.Desired();
        collision.Pivot.AddExistingSet(
            desired.Name,
            desired.RawMdx,
            desired.Caption,
            "user-owned");
        Assert.Throws<InvalidOperationException>(() =>
            collision.Gateway.CreateSet(collision.Bind(), desired));
        Assert.Empty(collision.Pivot.CalculatedMembers.AddCalls);

        var sibling = new HostFixture();
        FakePivot siblingPivot = sibling.AddPivot(sibling.Sheet, "PivotTable2");
        siblingPivot.AddExistingSet(
            sibling.Desired().Name,
            sibling.Desired().RawMdx,
            "Sibling",
            "user-owned");
        Assert.Throws<InvalidOperationException>(() =>
            sibling.Gateway.CreateSet(sibling.Bind(), sibling.Desired()));

        var reference = new HostFixture();
        reference.Pivot.AddCalculatedMember(
            "User Calculation",
            0,
            "{" + reference.Desired().Name + "}");
        Assert.Throws<InvalidOperationException>(() =>
            reference.Gateway.CreateSet(reference.Bind(), reference.Desired()));
    }

    [Fact]
    public void Unexpected_sibling_inventory_delta_rolls_back_only_the_intended_set()
    {
        var fixture = new HostFixture();
        DesiredPivotNamedSet desired = fixture.Desired();
        FakePivot sibling = fixture.AddPivot(fixture.Sheet, "PivotTable2");
        fixture.Pivot.AfterCalculatedMemberAddCommit = () => sibling.AddCalculatedMember(
            "Unexpected Calculation",
            0,
            "{[Measures].[Unexpected]}");

        Assert.Throws<PivotNamedSetRecoveryRequiredException>(() =>
            fixture.Gateway.CreateSet(fixture.Bind(), desired));

        Assert.DoesNotContain(fixture.Pivot.CalculatedItems, item =>
            string.Equals(item.Name, desired.Name, StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(fixture.Pivot.CubeItems, item =>
            string.Equals(item.SourceName, desired.Name, StringComparison.OrdinalIgnoreCase));
        Assert.Contains(sibling.CalculatedItems, item =>
            item.Name == "Unexpected Calculation");
    }

    [Theory]
    [InlineData("{'[PivotTablePlus_setup_rows]'}")]
    [InlineData("StrToSet('[Unrelated]')")]
    [InlineData("NameToSet(\"[Unrelated]\")")]
    public void Blocks_quoted_and_dynamic_formula_references_before_destructive_use(
        string formula)
    {
        var fixture = new HostFixture();
        fixture.Gateway.CreateSet(fixture.Bind(), fixture.Desired());
        LivePivotNamedSetSnapshot before = fixture.Gateway.Capture(fixture.Bind())
            .SelectedPivot.Artifacts.Single();
        fixture.Pivot.AddCalculatedMember("User Dynamic Calculation", 0, formula);

        Assert.Throws<InvalidOperationException>(() =>
            fixture.Gateway.DeleteSet(fixture.Bind(), before));
        Assert.Contains(fixture.Pivot.CalculatedItems, item =>
            item.Name == before.Name);
    }

    [Fact]
    public void Replaces_hidden_owned_set_by_delete_and_recreate()
    {
        var fixture = new HostFixture();
        DesiredPivotNamedSet beforeDesired = fixture.Desired();
        fixture.Gateway.CreateSet(fixture.Bind(), beforeDesired);
        LivePivotNamedSetSnapshot before = fixture.Gateway.Capture(fixture.Bind())
            .SelectedPivot.Artifacts.Single();
        DesiredPivotNamedSet afterDesired = fixture.Desired(caption: "Changed Rows");

        LivePivotNamedSetSnapshot replaced = fixture.Gateway.ReplaceSet(
            fixture.Bind(),
            before,
            afterDesired);

        Assert.Equal("Changed Rows", replaced.Caption);
        Assert.Equal(afterDesired.DisplayFolderMarker, replaced.DisplayFolder);
        Assert.Contains("CubeField.Delete", fixture.Pivot.OperationLog);
        Assert.Contains("CalculatedMember.Delete", fixture.Pivot.OperationLog);
        Assert.Equal(2, fixture.Pivot.CalculatedMembers.AddCalls.Count);
    }

    [Fact]
    public void Replace_restores_the_exact_prior_snapshot_when_recreation_fails()
    {
        var fixture = new HostFixture();
        fixture.Gateway.CreateSet(fixture.Bind(), fixture.Desired());
        LivePivotNamedSetSnapshot before = fixture.Gateway.Capture(fixture.Bind())
            .SelectedPivot.Artifacts.Single();
        fixture.Pivot.CalculatedMembers.BeforeCommitFailuresRemaining = 2;

        InvalidOperationException failure = Assert.Throws<InvalidOperationException>(() =>
            fixture.Gateway.ReplaceSet(
                fixture.Bind(),
                before,
                fixture.Desired(caption: "Changed Rows")));

        Assert.Contains("restored", failure.Message, StringComparison.OrdinalIgnoreCase);
        LivePivotNamedSetSnapshot restored = fixture.Gateway.Capture(fixture.Bind())
            .SelectedPivot.Artifacts.Single();
        Assert.Equal(before.LiveFingerprint, restored.LiveFingerprint);
    }

    [Fact]
    public void Restore_reconciles_a_matching_calculated_member_only_orphan()
    {
        var fixture = new HostFixture();
        fixture.Gateway.CreateSet(fixture.Bind(), fixture.Desired());
        LivePivotNamedSetSnapshot before = fixture.Gateway.Capture(fixture.Bind())
            .SelectedPivot.Artifacts.Single();
        fixture.Pivot.CubeItems.Single(field => field.CubeFieldType == 3).Delete();

        LivePivotNamedSetSnapshot restored = fixture.Gateway.RestoreSet(
            fixture.Bind(),
            before);

        Assert.Equal(before.LiveFingerprint, restored.LiveFingerprint);
        Assert.Equal(PivotNamedSetPairState.Complete, restored.PairState);
    }

    [Fact]
    public void Deletes_cube_field_before_calculated_member_and_verifies_absence()
    {
        var fixture = new HostFixture();
        fixture.Gateway.CreateSet(fixture.Bind(), fixture.Desired());
        LivePivotNamedSetSnapshot before = fixture.Gateway.Capture(fixture.Bind())
            .SelectedPivot.Artifacts.Single();
        fixture.Pivot.OperationLog.Clear();

        fixture.Gateway.DeleteSet(fixture.Bind(), before);

        Assert.Equal(
            new[] { "CubeField.Delete", "CalculatedMember.Delete" },
            fixture.Pivot.OperationLog.Where(value => value.EndsWith(
                ".Delete",
                StringComparison.Ordinal)));
        Assert.Empty(fixture.Gateway.Capture(fixture.Bind()).SelectedPivot.Artifacts);
    }

    [Fact]
    public void Restores_the_exact_pair_when_delete_fails_after_cube_field_removal()
    {
        var fixture = new HostFixture();
        fixture.Gateway.CreateSet(fixture.Bind(), fixture.Desired());
        LivePivotNamedSetSnapshot before = fixture.Gateway.Capture(fixture.Bind())
            .SelectedPivot.Artifacts.Single();
        fixture.Pivot.CalculatedItems.Single().DeleteFailure = FakeFailure.BeforeCommitOnce;

        InvalidOperationException failure = Assert.Throws<InvalidOperationException>(() =>
            fixture.Gateway.DeleteSet(fixture.Bind(), before));

        Assert.Contains("restored", failure.Message, StringComparison.OrdinalIgnoreCase);
        LivePivotNamedSetSnapshot restored = fixture.Gateway.Capture(fixture.Bind())
            .SelectedPivot.Artifacts.Single();
        Assert.Equal(before.LiveFingerprint, restored.LiveFingerprint);
    }

    [Fact]
    public void Blocks_destructive_use_when_set_is_visible_or_referenced()
    {
        var visible = new HostFixture();
        visible.Gateway.CreateSet(visible.Bind(), visible.Desired());
        visible.Pivot.CubeItems.Single(field => field.CubeFieldType == 3).Orientation = 1;
        LivePivotNamedSetSnapshot visibleSnapshot = visible.Gateway.Capture(visible.Bind())
            .SelectedPivot.Artifacts.Single();
        Assert.Throws<InvalidOperationException>(() =>
            visible.Gateway.DeleteSet(visible.Bind(), visibleSnapshot));

        var referenced = new HostFixture();
        referenced.Gateway.CreateSet(referenced.Bind(), referenced.Desired());
        LivePivotNamedSetSnapshot referencedSnapshot = referenced.Gateway.Capture(
            referenced.Bind()).SelectedPivot.Artifacts.Single();
        referenced.Pivot.AddCalculatedMember(
            "User Calculation",
            0,
            "{" + referencedSnapshot.Name + "}");
        Assert.Throws<InvalidOperationException>(() =>
            referenced.Gateway.DeleteSet(referenced.Bind(), referencedSnapshot));
    }

    [Fact]
    public void Rejects_mismatched_calculated_and_cube_pair_settings_before_mutation()
    {
        var flatten = new HostFixture();
        flatten.Gateway.CreateSet(flatten.Bind(), flatten.Desired());
        flatten.Pivot.CubeItems.Single(field => field.CubeFieldType == 3)
            .FlattenHierarchies = true;
        LivePivotNamedSetSnapshot mismatchedFlatten = flatten.Gateway.Capture(
            flatten.Bind()).SelectedPivot.Artifacts.Single();
        flatten.Pivot.OperationLog.Clear();

        Assert.Throws<InvalidOperationException>(() =>
            flatten.Gateway.DeleteSet(flatten.Bind(), mismatchedFlatten));
        Assert.DoesNotContain(flatten.Pivot.OperationLog, value =>
            value.EndsWith(".Delete", StringComparison.Ordinal));

        var hierarchy = new HostFixture();
        hierarchy.Gateway.CreateSet(hierarchy.Bind(), hierarchy.Desired());
        hierarchy.Pivot.CubeItems.Single(field => field.CubeFieldType == 3)
            .HierarchizeDistinct = true;
        LivePivotNamedSetSnapshot mismatchedHierarchy = hierarchy.Gateway.Capture(
            hierarchy.Bind()).SelectedPivot.Artifacts.Single();
        hierarchy.Pivot.OperationLog.Clear();

        Assert.Throws<InvalidOperationException>(() =>
            hierarchy.Gateway.DeleteSet(hierarchy.Bind(), mismatchedHierarchy));
        Assert.DoesNotContain(hierarchy.Pivot.OperationLog, value =>
            value.EndsWith(".Delete", StringComparison.Ordinal));
    }

    public enum FakeFailure
    {
        None,
        BeforeCommitOnce,
        AfterCommitOnce
    }

    public sealed class HostFixture
    {
        public HostFixture()
        {
            ModelConnection = new FakeConnection("ThisWorkbookDataModel", 7);
            SourceConnection = new FakeConnection("Synthetic source", 1);
            Model = new FakeModel(
                ModelConnection,
                new FakeModelTable("Sales", SourceConnection));
            Workbook = new FakeWorkbook(Model);
            Sheet = AddWorksheet("Sheet1");
            Pivot = AddPivot(Sheet, "PivotTable1");
            Gateway = new LateBoundPivotNamedSetGateway();
        }

        public FakeConnection ModelConnection { get; }
        public FakeConnection SourceConnection { get; }
        public FakeModel Model { get; }
        public FakeWorkbook Workbook { get; }
        public FakeWorksheet Sheet { get; }
        public FakePivot Pivot { get; }
        internal LateBoundPivotNamedSetGateway Gateway { get; }

        internal BoundPivotNamedSetTarget Bind()
        {
            return Gateway.Bind(Workbook, Pivot, Context(PivotSourceKind.DataModel));
        }

        internal DesiredPivotNamedSet Desired(
            string caption = "Management Rows",
            bool flattenHierarchies = false)
        {
            string sourceFingerprint = Gateway.DiscoverSchema(Bind())
                .Schema.SourceFingerprint;
            PivotMdxCompilation compilation =
                PivotNamedSetCanonicalTests.CompileDefaultMemberSet(
                    caption,
                    sourceFingerprint: sourceFingerprint,
                    flattenHierarchies: flattenHierarchies);
            return Assert.Single(PivotNamedSetCompilationAdapter.CreateDesired(
                "setup_1",
                compilation));
        }

        public PivotTableContext Context(PivotSourceKind kind)
        {
            string workbookId = new StoredWorkbookIdentityResolver().Resolve(Workbook);
            PivotCapability capabilities = PivotCapability.NativeFieldPlacement |
                                           PivotCapability.DataModel |
                                           PivotCapability.CalculatedMembers |
                                           PivotCapability.NamedSets;
            return new PivotTableContext(
                new PivotLayoutDefinition(
                    new PivotTargetIdentity(workbookId, Sheet.Name, Pivot.Name),
                    new PivotSourceDescriptor(
                        kind,
                        "ThisWorkbookDataModel",
                        capabilities,
                        "Sales"),
                    fields: Array.Empty<PivotFieldDescriptor>(),
                    placements: Array.Empty<PivotFieldPlacement>(),
                    clearAll: true),
                isConnected: true,
                sourceFieldsComplete: true);
        }

        public FakeWorksheet AddWorksheet(string name)
        {
            var worksheet = new FakeWorksheet(name, Workbook);
            Workbook.WorksheetItems.Add(worksheet);
            return worksheet;
        }

        public FakePivot AddPivot(FakeWorksheet worksheet, string name)
        {
            var pivot = new FakePivot(name, worksheet, ModelConnection);
            worksheet.PivotItems.Add(pivot);
            return pivot;
        }
    }

    public sealed class FakeWorkbook
    {
        private object? activeCell;

        public FakeWorkbook(FakeModel model)
        {
            Model = model;
            Worksheets = new FakeCollection<FakeWorksheet>(() => WorksheetItems);
        }

        public FakeModel Model { get; }
        public List<FakeWorksheet> WorksheetItems { get; } = new();
        public FakeCollection<FakeWorksheet> Worksheets { get; }
        public FakeCustomXmlParts CustomXMLParts { get; } = new();
        public int ActiveCellReads { get; private set; }

        public object? ActiveCell
        {
            get
            {
                ActiveCellReads++;
                return activeCell;
            }
            set => activeCell = value;
        }
    }

    public sealed class FakeCustomXmlParts
    {
        private readonly FakeCollection<object> empty =
            new FakeCollection<object>(() => Array.Empty<object>());

        public FakeCollection<object> SelectByNamespace(string namespaceUri)
        {
            _ = namespaceUri;
            return empty;
        }
    }

    public sealed class FakeWorksheet
    {
        public FakeWorksheet(string name, FakeWorkbook parent)
        {
            Name = name;
            Parent = parent;
            PivotTableCollection = new FakeCollection<FakePivot>(() => PivotItems);
        }

        public string Name { get; }
        public FakeWorkbook Parent { get; }
        public List<FakePivot> PivotItems { get; } = new();
        public FakeCollection<FakePivot> PivotTableCollection { get; }
        public FakeCollection<FakePivot> PivotTables() => PivotTableCollection;
    }

    public sealed class FakePivot
    {
        public FakePivot(string name, FakeWorksheet parent, FakeConnection connection)
        {
            Name = name;
            Parent = parent;
            Cache = new FakePivotCache(this, true, connection);
            CalculatedMembers = new FakeCalculatedMemberCollection(this);
            CubeFields = new FakeCubeFieldCollection(this);
        }

        public string Name { get; }
        public FakeWorksheet Parent { get; }
        public FakePivotCache Cache { get; }
        public List<FakeCalculatedMember> CalculatedItems { get; } = new();
        public List<FakeCubeField> CubeItems { get; } = new();
        public List<string> OperationLog { get; } = new();
        public FakeCalculatedMemberCollection CalculatedMembers { get; }
        public FakeCubeFieldCollection CubeFields { get; }
        public Func<string, string> FormulaReadbackTransform { get; set; } =
            PivotNamedSetFormulaTransport.DecodeRequired;
        public int NewCalculatedMemberFormulaReadFailureOnRead { get; set; }
        public FakeFailure NewCalculatedMemberDeleteFailure { get; set; }
        public Action? AfterCalculatedMemberAddCommit { get; set; }
        public int CalculatedFlattenSetFailuresRemaining { get; set; }
        public int CubeFlattenSetFailuresRemaining { get; set; }

        public FakePivotCache PivotCache() => Cache;

        public FakeCubeField AddHierarchy(string sourceName, string caption)
        {
            var field = new FakeCubeField(this, sourceName, caption, 1);
            CubeItems.Add(field);
            return field;
        }

        public FakeCalculatedMember AddCalculatedMember(
            string name,
            int type,
            string rawFormula,
            string displayFolder = "")
        {
            var member = new FakeCalculatedMember(
                this,
                name,
                type,
                rawFormula,
                displayFolder,
                dynamic: false,
                hierarchizeDistinct: false)
            {
                IsValid = true
            };
            CalculatedItems.Add(member);
            return member;
        }

        public FakeCalculatedMember AddCalculatedOnly(
            string name,
            string rawFormula,
            string displayFolder)
        {
            return AddCalculatedMember(name, 1, rawFormula, displayFolder);
        }

        public FakeCubeField AddCubeOnly(string name, string caption)
        {
            var cube = new FakeCubeField(this, name, caption, 3);
            CubeItems.Add(cube);
            return cube;
        }

        public void AddExistingSet(
            string name,
            string rawFormula,
            string caption,
            string displayFolder,
            int orientation = 0)
        {
            AddCalculatedOnly(name, rawFormula, displayFolder);
            FakeCubeField cube = AddCubeOnly(name, caption);
            cube.Orientation = orientation;
        }
    }

    public sealed class FakePivotCache
    {
        private readonly FakePivot owner;

        public FakePivotCache(
            FakePivot owner,
            bool olap,
            FakeConnection workbookConnection)
        {
            this.owner = owner;
            OLAP = olap;
            WorkbookConnection = workbookConnection;
        }

        public bool OLAP { get; set; }
        public FakeConnection WorkbookConnection { get; set; }
        public int MakeConnectionCalls { get; private set; }
        public int AdoConnectionReads { get; private set; }
        public int MakeConnectionFailuresRemaining { get; set; }
        public bool KeepCalculatedMembersInvalid { get; set; }

        public object ADOConnection
        {
            get
            {
                AdoConnectionReads++;
                throw new InvalidOperationException("ADOConnection must never be read.");
            }
        }

        public void MakeConnection()
        {
            owner.OperationLog.Add("PivotCache.MakeConnection");
            MakeConnectionCalls++;
            if (MakeConnectionFailuresRemaining > 0)
            {
                MakeConnectionFailuresRemaining--;
                throw new InvalidOperationException("simulated MakeConnection failure");
            }

            foreach (FakeCalculatedMember member in owner.CalculatedItems)
            {
                member.IsValid = !KeepCalculatedMembersInvalid;
            }
        }
    }

    public sealed class FakeCalculatedMemberCollection
    {
        private readonly FakePivot owner;

        public FakeCalculatedMemberCollection(FakePivot owner)
        {
            this.owner = owner;
        }

        public int Count => owner.CalculatedItems.Count;
        public List<FakeCalculatedMemberAddCall> AddCalls { get; } = new();
        public FakeFailure AddFailure { get; set; }
        public int BeforeCommitFailuresRemaining { get; set; }

        public FakeCalculatedMember Item(int index)
        {
            return owner.CalculatedItems[index - 1];
        }

        public object Add(
            string name,
            string formula,
            object solveOrder,
            int type,
            bool dynamic,
            string displayFolder,
            bool hierarchizeDistinct)
        {
            owner.OperationLog.Add("CalculatedMembers.Add");
            AddCalls.Add(new FakeCalculatedMemberAddCall(
                name,
                formula,
                solveOrder,
                type,
                dynamic,
                displayFolder,
                hierarchizeDistinct));
            if (BeforeCommitFailuresRemaining > 0)
            {
                BeforeCommitFailuresRemaining--;
                throw new InvalidOperationException("before CalculatedMembers.Add commit");
            }

            if (ConsumeAddFailure(FakeFailure.BeforeCommitOnce))
            {
                throw new InvalidOperationException("before CalculatedMembers.Add commit");
            }

            var created = new FakeCalculatedMember(
                owner,
                name,
                type,
                owner.FormulaReadbackTransform(formula),
                displayFolder,
                dynamic,
                hierarchizeDistinct);
            created.FormulaReadFailureOnRead =
                owner.NewCalculatedMemberFormulaReadFailureOnRead;
            created.DeleteFailure = owner.NewCalculatedMemberDeleteFailure;
            owner.CalculatedItems.Add(created);
            Action? afterCommit = owner.AfterCalculatedMemberAddCommit;
            owner.AfterCalculatedMemberAddCommit = null;
            afterCommit?.Invoke();
            if (ConsumeAddFailure(FakeFailure.AfterCommitOnce))
            {
                throw new InvalidOperationException("after CalculatedMembers.Add commit");
            }

            return created;
        }

        private bool ConsumeAddFailure(FakeFailure expected)
        {
            if (AddFailure != expected) return false;
            AddFailure = FakeFailure.None;
            return true;
        }
    }

    public sealed class FakeCalculatedMemberAddCall
    {
        public FakeCalculatedMemberAddCall(
            string name,
            string formula,
            object solveOrder,
            int type,
            bool dynamic,
            string displayFolder,
            bool hierarchizeDistinct)
        {
            Name = name;
            Formula = formula;
            SolveOrder = solveOrder;
            Type = type;
            Dynamic = dynamic;
            DisplayFolder = displayFolder;
            HierarchizeDistinct = hierarchizeDistinct;
        }

        public string Name { get; }
        public string Formula { get; }
        public object SolveOrder { get; }
        public int Type { get; }
        public bool Dynamic { get; }
        public string DisplayFolder { get; }
        public bool HierarchizeDistinct { get; }
    }

    public sealed class FakeCalculatedMember
    {
        private readonly FakePivot owner;
        private readonly string formula;
        private int formulaReads;
        private bool flattenHierarchies;

        public FakeCalculatedMember(
            FakePivot owner,
            string name,
            int type,
            string formula,
            string displayFolder,
            bool dynamic,
            bool hierarchizeDistinct)
        {
            this.owner = owner;
            Name = name;
            Type = type;
            this.formula = formula;
            DisplayFolder = displayFolder;
            Dynamic = dynamic;
            HierarchizeDistinct = hierarchizeDistinct;
        }

        public string Name { get; }
        public int Type { get; }
        public bool ThrowOnFormulaRead { get; set; }
        public int FormulaReadFailureOnRead { get; set; }
        public string Formula
        {
            get
            {
                formulaReads++;
                if (ThrowOnFormulaRead)
                {
                    throw new InvalidOperationException("simulated formula read failure");
                }

                if (FormulaReadFailureOnRead == formulaReads)
                {
                    throw new InvalidOperationException("simulated transient formula read failure");
                }

                return formula;
            }
        }
        public string DisplayFolder { get; }
        public bool Dynamic { get; }
        public bool FlattenHierarchies
        {
            get => flattenHierarchies;
            set
            {
                if (owner.CalculatedFlattenSetFailuresRemaining > 0)
                {
                    owner.CalculatedFlattenSetFailuresRemaining--;
                    throw new InvalidOperationException(
                        "simulated calculated-member property failure");
                }

                flattenHierarchies = value;
            }
        }
        public bool HierarchizeDistinct { get; }
        public bool IsValid { get; set; }
        public FakeFailure DeleteFailure { get; set; }

        public void Delete()
        {
            owner.OperationLog.Add("CalculatedMember.Delete");
            if (ConsumeDeleteFailure(FakeFailure.BeforeCommitOnce))
            {
                throw new InvalidOperationException("before CalculatedMember.Delete commit");
            }

            owner.CalculatedItems.Remove(this);
            if (ConsumeDeleteFailure(FakeFailure.AfterCommitOnce))
            {
                throw new InvalidOperationException("after CalculatedMember.Delete commit");
            }
        }

        private bool ConsumeDeleteFailure(FakeFailure expected)
        {
            if (DeleteFailure != expected) return false;
            DeleteFailure = FakeFailure.None;
            return true;
        }
    }

    public sealed class FakeCubeFieldCollection
    {
        private readonly FakePivot owner;

        public FakeCubeFieldCollection(FakePivot owner)
        {
            this.owner = owner;
        }

        public int Count => CountOverride ?? owner.CubeItems.Count;
        public int? CountOverride { get; set; }
        public FakeFailure AddFailure { get; set; }
        public bool MarkCalculatedMemberValidAfterAddSet { get; set; }

        public FakeCubeField Item(int index)
        {
            return owner.CubeItems[index - 1];
        }

        public object AddSet(string name, string caption)
        {
            owner.OperationLog.Add("CubeFields.AddSet");
            if (ConsumeAddFailure(FakeFailure.BeforeCommitOnce))
            {
                throw new InvalidOperationException("before CubeFields.AddSet commit");
            }

            FakeCalculatedMember calculated = owner.CalculatedItems.Single(item =>
                item.Type == 1 &&
                string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase));
            var cube = new FakeCubeField(owner, name, caption, 3)
            {
                HierarchizeDistinct = calculated.HierarchizeDistinct
            };
            owner.CubeItems.Add(cube);
            if (MarkCalculatedMemberValidAfterAddSet) calculated.IsValid = true;
            if (ConsumeAddFailure(FakeFailure.AfterCommitOnce))
            {
                throw new InvalidOperationException("after CubeFields.AddSet commit");
            }

            return cube;
        }

        private bool ConsumeAddFailure(FakeFailure expected)
        {
            if (AddFailure != expected) return false;
            AddFailure = FakeFailure.None;
            return true;
        }
    }

    public sealed class FakeCubeField
    {
        private readonly FakePivot owner;
        private bool flattenHierarchies;

        public FakeCubeField(
            FakePivot owner,
            string sourceName,
            string caption,
            int cubeFieldType)
        {
            this.owner = owner;
            SourceName = sourceName;
            Name = sourceName;
            Caption = caption;
            CubeFieldType = cubeFieldType;
        }

        public string SourceName { get; }
        public string Name { get; }
        public string Caption { get; }
        public int CubeFieldType { get; }
        public int Type => CubeFieldType;
        public bool FlattenHierarchies
        {
            get => flattenHierarchies;
            set
            {
                if (owner.CubeFlattenSetFailuresRemaining > 0)
                {
                    owner.CubeFlattenSetFailuresRemaining--;
                    throw new InvalidOperationException("simulated CubeField property failure");
                }

                flattenHierarchies = value;
            }
        }
        public bool HierarchizeDistinct { get; set; }
        public bool ShowInFieldList { get; set; } = true;
        public int Orientation { get; set; }
        public bool PivotFieldsAvailable { get; set; } = true;
        public List<FakePivotField> PivotFieldItems { get; } = new();
        public FakeCollection<FakePivotField>? PivotFields =>
            PivotFieldsAvailable
                ? new FakeCollection<FakePivotField>(() => PivotFieldItems)
                : null;
        public int CreatePivotFieldsCalls { get; private set; }

        public FakePivotField AddLevel(string sourceName, string caption)
        {
            var field = new FakePivotField(sourceName, caption);
            PivotFieldItems.Add(field);
            return field;
        }

        public void CreatePivotFields()
        {
            CreatePivotFieldsCalls++;
            throw new InvalidOperationException("CreatePivotFields must never be called.");
        }

        public void Delete()
        {
            owner.OperationLog.Add("CubeField.Delete");
            owner.CubeItems.Remove(this);
        }
    }

    public sealed class FakePivotField
    {
        public FakePivotField(string sourceName, string caption)
        {
            SourceName = sourceName;
            Name = sourceName;
            Caption = caption;
        }

        public string SourceName { get; }
        public string Name { get; }
        public string Caption { get; }
        public bool PivotItemsAvailable { get; set; } = true;
        public List<FakePivotItem> PivotItemItems { get; } = new();
        public FakeCollection<FakePivotItem>? PivotItems =>
            PivotItemsAvailable
                ? new FakeCollection<FakePivotItem>(() => PivotItemItems)
                : null;

        public void AddItem(string sourceName, string caption)
        {
            PivotItemItems.Add(new FakePivotItem(sourceName, caption));
        }
    }

    public sealed class FakePivotItem
    {
        public FakePivotItem(string sourceName, string caption)
        {
            SourceName = sourceName;
            Name = sourceName;
            Caption = caption;
        }

        public string SourceName { get; }
        public string Name { get; }
        public string Caption { get; }
    }

    public sealed class FakeConnection
    {
        public FakeConnection(string name, int type)
        {
            Name = name;
            Type = type;
        }

        public string Name { get; }
        public int Type { get; }
    }

    public sealed class FakeModelTable
    {
        public FakeModelTable(string name, FakeConnection sourceWorkbookConnection)
        {
            Name = name;
            SourceWorkbookConnection = sourceWorkbookConnection;
        }

        public string Name { get; }
        public FakeConnection SourceWorkbookConnection { get; }
    }

    public sealed class FakeModel
    {
        public FakeModel(FakeConnection dataModelConnection, FakeModelTable table)
        {
            DataModelConnection = dataModelConnection;
            ModelTableItems.Add(table);
            ModelTables = new FakeCollection<FakeModelTable>(() => ModelTableItems);
            ModelMeasures = new FakeCollection<FakeModelMeasure>(() => ModelMeasureItems);
        }

        public FakeConnection DataModelConnection { get; }
        public List<FakeModelTable> ModelTableItems { get; } = new();
        public FakeCollection<FakeModelTable> ModelTables { get; }
        public List<FakeModelMeasure> ModelMeasureItems { get; } = new();
        public FakeCollection<FakeModelMeasure> ModelMeasures { get; }
    }

    public sealed class FakeModelMeasure
    {
        public FakeModelMeasure(string name, string formula, string description)
        {
            Name = name;
            Formula = formula;
            Description = description;
        }

        public string Name { get; }
        public string Formula { get; set; }
        public string Description { get; set; }
    }

    public sealed class FakeCollection<T>
    {
        private readonly Func<IReadOnlyList<T>> source;

        public FakeCollection(Func<IReadOnlyList<T>> source)
        {
            this.source = source;
        }

        public int Count => source().Count;

        public T Item(int index)
        {
            if (index <= 0) throw new ArgumentOutOfRangeException(nameof(index));
            return source()[index - 1];
        }
    }
}
