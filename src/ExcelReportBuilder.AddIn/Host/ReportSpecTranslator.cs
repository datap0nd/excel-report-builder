using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.Json;
using ExcelReportBuilder.Core.Measures;
using ExcelReportBuilder.Core.Periods;
using ExcelReportBuilder.Core.Profiling;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Core.Transforms;

namespace ExcelReportBuilder.AddIn.Host
{
    internal sealed class ReportSpecTranslator
    {
        private const int MaximumCanonicalSnapshotCharacters = 512 * 1024;
        private const int MaximumManualBlocks = 8;

        public ReportSpecificationSnapshot ToAppliedAgentSnapshot(ReportSpecV1 specification)
        {
            if (specification == null) throw new ArgumentNullException(nameof(specification));
            if (!string.Equals(
                    specification.SchemaVersion,
                    ReportSpecV1.CurrentSchemaVersion,
                    StringComparison.Ordinal))
            {
                throw new NotSupportedException("Unknown report specification version.");
            }

            string canonical = ReportSpecJson.Serialize(
                specification,
                Newtonsoft.Json.Formatting.None);
            if (canonical.Length > MaximumCanonicalSnapshotCharacters)
            {
                throw new InvalidOperationException(
                    "The agent report setup exceeds the bounded UI snapshot size.");
            }

            var placements = new List<FieldPlacementSnapshot>();
            ReportBlockSpec? firstBlock = specification.Blocks.FirstOrDefault();
            if (firstBlock != null)
            {
                foreach (FieldPlacementSpec row in firstBlock.Layout.Rows)
                {
                    placements.Add(new FieldPlacementSnapshot(
                        PlacementBucket.Rows,
                        row.Field,
                        ToUiSort(row.Sort),
                        row.Subtotals.Mode != SubtotalMode.None,
                        subtotalPlacement: ToUiTotalPlacement(row.Subtotals.Placement),
                        memberOrder: row.MemberOrder.Select(ToDisplayScalar).Where(value => value != null).Cast<string>().ToArray()));
                }

                foreach (FieldPlacementSpec column in firstBlock.Layout.Columns)
                {
                    placements.Add(new FieldPlacementSnapshot(
                        PlacementBucket.Columns,
                        column.Field,
                        ToUiSort(column.Sort),
                        column.Subtotals.Mode != SubtotalMode.None,
                        subtotalPlacement: ToUiTotalPlacement(column.Subtotals.Placement),
                        memberOrder: column.MemberOrder.Select(ToDisplayScalar).Where(value => value != null).Cast<string>().ToArray()));
                }

                var measures = specification.Measures.ToDictionary(
                    measure => measure.Id,
                    StringComparer.OrdinalIgnoreCase);
                foreach (ValuePlacementSpec value in firstBlock.Layout.Values)
                {
                    if (!measures.TryGetValue(value.MeasureId, out MeasureDefinition? measure))
                    {
                        continue;
                    }

                    if (measure.Expression is AggregateMeasureExpression aggregate)
                    {
                        placements.Add(new FieldPlacementSnapshot(
                            PlacementBucket.Values,
                            aggregate.Field,
                            ToUiAggregate(aggregate.Function),
                            numberFormat: value.NumberFormat ?? measure.NumberFormat ?? "General"));
                    }
                    else
                    {
                        placements.Add(new FieldPlacementSnapshot(
                            PlacementBucket.Values,
                            measure.Label,
                            "Calculated metric"));
                    }
                }

                foreach (FilterPlacementSpec filter in firstBlock.Layout.Filters)
                {
                    string[] selectedValues = filter.SelectedValues
                        .Select(ToDisplayScalar)
                        .Where(value => value != null)
                        .Cast<string>()
                        .ToArray();
                    placements.Add(new FieldPlacementSnapshot(
                        PlacementBucket.Filters,
                        filter.Field,
                        selectedValues.Length == 0 ? "All" : string.Join("; ", selectedValues),
                        selectedValues: selectedValues));
                }
            }

            return new ReportSpecificationSnapshot(
                ToUiPeriodMapping(specification.PeriodMapping),
                placements,
                firstBlock == null
                    ? "Dense management block"
                    : ToUiOutputStyle(firstBlock.OutputMode),
                canonical,
                blocks: ToManualBlocks(specification.Blocks),
                layout: firstBlock == null
                    ? new ManualLayoutSnapshot()
                    : ToManualLayout(firstBlock),
                checks: ToManualChecks(specification.Checks),
                manualProjectionComplete: false);
        }

        public ReportSpecificationSnapshot ToUi(ReportSpecV1 specification)
        {
            if (specification == null) throw new ArgumentNullException(nameof(specification));
            if (!string.Equals(
                    specification.SchemaVersion,
                    ReportSpecV1.CurrentSchemaVersion,
                    StringComparison.Ordinal))
            {
                throw new NotSupportedException("Unknown report specification version.");
            }

            if (specification.Blocks.Count < 1 ||
                specification.Blocks.Count > MaximumManualBlocks)
            {
                throw new InvalidOperationException(
                    "The saved setup cannot be represented by the bounded manual builder because it must contain between one and eight report blocks.");
            }

            EnsureTransformsCanRoundTrip(specification);
            ReportBlockSpec block = specification.Blocks[0];
            var measures = specification.Measures.ToDictionary(
                measure => measure.Id,
                StringComparer.OrdinalIgnoreCase);
            IReadOnlyList<FieldPlacementSnapshot> placements = ProjectManualPlacements(
                specification,
                block,
                measures);
            ManualLayoutSnapshot layout = ToManualLayout(block);
            for (var index = 1; index < specification.Blocks.Count; index++)
            {
                ReportBlockSpec candidate = specification.Blocks[index];
                IReadOnlyList<FieldPlacementSnapshot> candidatePlacements = ProjectManualPlacements(
                    specification,
                    candidate,
                    measures);
                if (candidate.OutputMode != block.OutputMode ||
                    !ManualPlacementsEqual(placements, candidatePlacements) ||
                    !ValueMeasureOrderEqual(block.Layout.Values, candidate.Layout.Values) ||
                    !ManualLayoutsEqual(layout, ToManualLayout(candidate)))
                {
                    throw new InvalidOperationException(
                        "The saved setup contains report blocks with different Rows, Columns, Values, Filters, output styles, or layout settings. The bounded manual builder can edit multiple blocks only when those shared settings are identical.");
                }
            }

            return new ReportSpecificationSnapshot(
                ToUiPeriodMapping(specification.PeriodMapping),
                placements,
                ToUiOutputStyle(block.OutputMode),
                blocks: ToManualBlocks(specification.Blocks),
                layout: layout,
                checks: ToManualChecks(specification.Checks));
        }

        public ReportSpecV1 FromUi(
            ReportSpecificationSnapshot snapshot,
            SourceProfile sourceProfile,
            string workbookObjectName,
            WorkbookSourceKind sourceKind,
            string reportId,
            Func<ExcludeTotalRowsTransform, long>? totalRowEvidenceCounter = null)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            if (snapshot.HasCanonicalReportSpec)
            {
                if (snapshot.CanonicalReportSpecJson.Length > MaximumCanonicalSnapshotCharacters)
                {
                    throw new InvalidOperationException(
                        "The canonical report setup exceeds the bounded snapshot size.");
                }

                ReportSpecV1 canonical = ReportSpecJson.Deserialize(snapshot.CanonicalReportSpecJson);
                SourceFingerprintSpec currentFingerprint = SourceFingerprint.FromHeaders(
                    sourceProfile.Columns
                        .OrderBy(column => column.Index)
                        .Select(column => column.Name));
                if (!string.Equals(
                        canonical.Source.Fingerprint.GetSavedSetupKey(),
                        currentFingerprint.GetSavedSetupKey(),
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "The applied agent report setup does not match the currently selected Data columns.");
                }

                canonical.Source = new WorkbookSourceSpec
                {
                    Kind = sourceKind,
                    WorkbookObjectName = workbookObjectName,
                    HeaderRowCount = 1,
                    Fingerprint = currentFingerprint
                };
                ResolveManualTotalRowEvidence(canonical.Transforms, totalRowEvidenceCounter);
                return canonical;
            }

            var inputs = snapshot.Placements.Select(PlacementInput.FromUi).ToList();
            List<TransformStep> manualTransforms = new ManualTransformTranslator()
                .Translate(snapshot.Transforms);
            ResolveManualTotalRowEvidence(manualTransforms, totalRowEvidenceCounter);
            return Create(
                snapshot.PeriodMapping,
                inputs,
                sourceProfile,
                workbookObjectName,
                sourceKind,
                reportId,
                ParseOutputMode(snapshot.OutputStyle),
                manualTransforms,
                snapshot.CalculatedMetrics,
                snapshot.Blocks,
                snapshot.Layout,
                snapshot.Checks);
        }

        private static void ResolveManualTotalRowEvidence(
            IReadOnlyList<TransformStep> transforms,
            Func<ExcludeTotalRowsTransform, long>? evidenceCounter)
        {
            foreach (ExcludeTotalRowsTransform exclusion in transforms.OfType<ExcludeTotalRowsTransform>())
            {
                if (evidenceCounter == null)
                {
                    throw new InvalidOperationException(
                        "Total-row exclusion requires a full-source evidence check before the report can be built.");
                }

                long observed = evidenceCounter(exclusion);
                if (observed <= 0)
                {
                    throw new InvalidOperationException(
                        "No source rows matched the confirmed total-row values. The source was not changed.");
                }

                foreach (TotalRowEvidenceSpec evidence in exclusion.Evidence)
                {
                    evidence.ObservedMatchCount = observed;
                }
            }
        }

        public ReportSpecV1 FromAgentProposal(
            string argumentsJson,
            PeriodMappingSnapshot periodMapping,
            SourceProfile sourceProfile,
            string workbookObjectName,
            WorkbookSourceKind sourceKind,
            string reportId,
            ReportOutputMode outputMode = ReportOutputMode.DenseGrid,
            IReadOnlyList<TransformStep>? proposedTransforms = null)
        {
            if (string.IsNullOrWhiteSpace(argumentsJson))
            {
                throw new ArgumentException("A report proposal is required.", nameof(argumentsJson));
            }

            using (var document = JsonDocument.Parse(argumentsJson))
            {
                JsonElement root = document.RootElement;
                if (root.TryGetProperty("version", out _))
                {
                    return CreateAdvancedAgentSpecification(
                        root,
                        periodMapping,
                        sourceProfile,
                        workbookObjectName,
                        sourceKind,
                        reportId,
                        proposedTransforms);
                }

                var subtotals = ReadFieldSettings(root, "subtotals", "mode");
                var ordering = ReadFieldSettings(root, "ordering", "direction");
                var formatting = ReadFormatting(root);
                var inputs = new List<PlacementInput>();
                foreach (string field in ReadStrings(root, "rows"))
                {
                    inputs.Add(new PlacementInput
                    {
                        Bucket = PlacementBucket.Rows,
                        Field = field,
                        Subtotals = subtotals.TryGetValue(field, out string? subtotal)
                            ? subtotal
                            : "show",
                        Sort = ordering.TryGetValue(field, out string? rowSort)
                            ? rowSort
                            : "sourceOrder"
                    });
                }

                foreach (string field in ReadStrings(root, "columns"))
                {
                    inputs.Add(new PlacementInput
                    {
                        Bucket = PlacementBucket.Columns,
                        Field = field,
                        Subtotals = subtotals.TryGetValue(field, out string? subtotal)
                            ? subtotal
                            : "show",
                        Sort = ordering.TryGetValue(field, out string? columnSort)
                            ? columnSort
                            : "sourceOrder"
                    });
                }

                foreach (JsonElement value in root.GetProperty("values").EnumerateArray())
                {
                    string field = value.GetProperty("field").GetString() ?? string.Empty;
                    inputs.Add(new PlacementInput
                    {
                        Bucket = PlacementBucket.Values,
                        Field = field,
                        Aggregate = value.GetProperty("aggregation").GetString() ?? "sum",
                        NumberFormat = formatting.TryGetValue(field, out string? numberFormat)
                            ? numberFormat
                            : null
                    });
                }

                foreach (JsonElement filter in root.GetProperty("filters").EnumerateArray())
                {
                    string filterOperator = filter.GetProperty("operator").GetString() ?? string.Empty;
                    if (!string.Equals(filterOperator, "equals", StringComparison.Ordinal) &&
                        !string.Equals(filterOperator, "in", StringComparison.Ordinal))
                    {
                        throw new InvalidOperationException(
                            "The proposed filter operator cannot be represented by the bounded workbook filter contract.");
                    }

                    inputs.Add(new PlacementInput
                    {
                        Bucket = PlacementBucket.Filters,
                        Field = filter.GetProperty("field").GetString() ?? string.Empty,
                        FilterValues = filter.GetProperty("values")
                            .EnumerateArray()
                            .Select(value => value.GetString() ?? string.Empty)
                            .ToList()
                    });
                }

                return Create(
                    periodMapping,
                    inputs,
                    sourceProfile,
                    workbookObjectName,
                    sourceKind,
                    reportId,
                    outputMode,
                    proposedTransforms);
            }
        }

        public PeriodMappingSpec? ResolvePeriodMapping(
            PeriodMappingSnapshot snapshot,
            SourceProfile sourceProfile)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            if (sourceProfile == null) throw new ArgumentNullException(nameof(sourceProfile));

            if (string.Equals(snapshot.Mode, "No period columns", StringComparison.Ordinal))
            {
                return null;
            }

            if (string.Equals(snapshot.Mode, "Date column", StringComparison.Ordinal))
            {
                SourceColumnProfile? column = sourceProfile.FindColumn(snapshot.PeriodColumn);
                if (column == null)
                {
                    throw new InvalidOperationException("The selected period column is not in the chosen Data.");
                }

                if (column.InferredType != SourceValueType.Date &&
                    column.InferredType != SourceValueType.DateTime &&
                    column.PeriodLikeRatio < 0.8d)
                {
                    throw new InvalidOperationException("The selected period column is not period-like.");
                }

                return new PeriodMappingSpec
                {
                    Id = "periods",
                    Kind = PeriodMappingKind.LongDateColumn,
                    DateColumn = column.Name,
                    ReportingYear = snapshot.ReportingYear,
                    Grain = ResolveSelectedPeriodGrain(column),
                    KeyColumns = sourceProfile.Columns
                        .Where(candidate => !string.Equals(
                            candidate.Name,
                            column.Name,
                            StringComparison.OrdinalIgnoreCase))
                        .Select(candidate => candidate.Name)
                        .ToList()
                };
            }

            if (!string.Equals(snapshot.Mode, "Wide period headers", StringComparison.Ordinal))
            {
                throw new InvalidOperationException("The selected period mode is not supported.");
            }

            PeriodDetectionResult detection = PeriodDetector.Detect(
                sourceProfile,
                snapshot.ReportingYear);
            if (detection.Kind != PeriodLayoutKind.MonthHeaders &&
                detection.Kind != PeriodLayoutKind.MetricMonthHeaders)
            {
                throw new InvalidOperationException("Wide period headers were not detected in the chosen Data.");
            }

            if (detection.IsAmbiguous)
            {
                string message = detection.Issues.First(issue =>
                    issue.Severity == PeriodDetectionSeverity.Error).Message;
                throw new InvalidOperationException(message);
            }

            return detection.ToPeriodMapping();
        }

        private ReportSpecV1 CreateAdvancedAgentSpecification(
            JsonElement root,
            PeriodMappingSnapshot periodSnapshot,
            SourceProfile sourceProfile,
            string workbookObjectName,
            WorkbookSourceKind sourceKind,
            string reportId,
            IReadOnlyList<TransformStep>? proposedTransforms)
        {
            PeriodMappingSpec? periodMapping = ResolvePeriodMapping(periodSnapshot, sourceProfile);
            string specificationId = BoundedId(reportId);
            var specification = new ReportSpecV1
            {
                Id = specificationId,
                Name = "AI-assisted report setup",
                OwnershipId = BoundedId("owner_" + reportId),
                Source = new WorkbookSourceSpec
                {
                    Kind = sourceKind,
                    WorkbookObjectName = workbookObjectName,
                    HeaderRowCount = 1,
                    Fingerprint = SourceFingerprint.FromHeaders(
                        sourceProfile.Columns
                            .OrderBy(column => column.Index)
                            .Select(column => column.Name))
                },
                PeriodMapping = periodMapping
            };
            if (proposedTransforms != null)
            {
                specification.Transforms.AddRange(proposedTransforms);
            }

            AddRequiredPeriodPreparation(specification, sourceProfile);
            ValidateTransformTypeSafety(
                sourceProfile,
                periodMapping,
                specification.Transforms);

            var measureTypes = new Dictionary<string, MeasureValueType>(StringComparer.OrdinalIgnoreCase);
            foreach (JsonElement measure in root.GetProperty("measures").EnumerateArray())
            {
                measureTypes.Add(
                    measure.GetProperty("id").GetString() ?? string.Empty,
                    ParseMeasureValueType(measure.GetProperty("valueType").GetString()));
            }

            foreach (JsonElement measure in root.GetProperty("measures").EnumerateArray())
            {
                string id = measure.GetProperty("id").GetString() ?? string.Empty;
                MeasureValueType valueType = measureTypes[id];
                string numberFormat = measure.GetProperty("numberFormat").GetString() ?? string.Empty;
                specification.Measures.Add(new MeasureDefinition
                {
                    Id = id,
                    Label = measure.GetProperty("label").GetString() ?? id,
                    ValueType = valueType,
                    NumberFormat = EmptyToNull(numberFormat),
                    Expression = ParseAdvancedExpression(
                        measure.GetProperty("expression"),
                        valueType,
                        measureTypes)
                });
            }

            foreach (JsonElement style in root.GetProperty("styles").EnumerateArray())
            {
                int decimalPlaces = style.GetProperty("decimalPlaces").GetInt32();
                specification.Styles.Add(new PresentationStyleSpec
                {
                    Id = style.GetProperty("id").GetString() ?? string.Empty,
                    Bold = style.GetProperty("bold").GetBoolean(),
                    Italic = style.GetProperty("italic").GetBoolean(),
                    FontColor = EmptyToNull(style.GetProperty("fontColor").GetString()),
                    FillColor = EmptyToNull(style.GetProperty("fillColor").GetString()),
                    HorizontalAlignment = ParseHorizontalAlignment(
                        style.GetProperty("horizontalAlignment").GetString()),
                    NumberFormat = EmptyToNull(style.GetProperty("numberFormat").GetString()),
                    DecimalPlaces = decimalPlaces < 0 ? (int?)null : decimalPlaces,
                    TopBorder = style.GetProperty("topBorder").GetBoolean(),
                    BottomBorder = style.GetProperty("bottomBorder").GetBoolean()
                });
            }

            var blockOwnershipIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (JsonElement blockElement in root.GetProperty("blocks").EnumerateArray())
            {
                string blockId = blockElement.GetProperty("id").GetString() ?? string.Empty;
                var block = new ReportBlockSpec
                {
                    Id = blockId,
                    OwnershipId = UniqueId(
                        "block_owner_" + reportId + "_" + blockId,
                        specification.Blocks.Count + 1,
                        blockOwnershipIds),
                    Title = EmptyToNull(blockElement.GetProperty("title").GetString()),
                    WorksheetName = blockElement.GetProperty("worksheetName").GetString() ?? "Report",
                    AnchorCell = blockElement.GetProperty("anchorCell").GetString() ?? "A1",
                    OutputMode = ParseAgentOutputMode(
                        blockElement.GetProperty("outputMode").GetString()),
                    OwnedExtent = new OwnedRangeExtentSpec
                    {
                        RowCount = 100000,
                        ColumnCount = 512
                    },
                    HeaderStyleId = EmptyToNull(blockElement.GetProperty("headerStyleId").GetString()),
                    BodyStyleId = EmptyToNull(blockElement.GetProperty("bodyStyleId").GetString()),
                    SubtotalStyleId = EmptyToNull(blockElement.GetProperty("subtotalStyleId").GetString()),
                    GrandTotalStyleId = EmptyToNull(blockElement.GetProperty("grandTotalStyleId").GetString())
                };

                foreach (JsonElement row in blockElement.GetProperty("rows").EnumerateArray())
                {
                    block.Layout.Rows.Add(ParseAdvancedFieldPlacement(row));
                }

                foreach (JsonElement column in blockElement.GetProperty("columns").EnumerateArray())
                {
                    block.Layout.Columns.Add(ParseAdvancedFieldPlacement(column));
                }

                foreach (JsonElement value in blockElement.GetProperty("values").EnumerateArray())
                {
                    var placement = new ValuePlacementSpec
                    {
                        MeasureId = value.GetProperty("measureId").GetString() ?? string.Empty,
                        Caption = EmptyToNull(value.GetProperty("caption").GetString()),
                        NumberFormat = EmptyToNull(value.GetProperty("numberFormat").GetString()),
                        StyleId = EmptyToNull(value.GetProperty("styleId").GetString())
                    };
                    placement.PeriodSliceIds.AddRange(value.GetProperty("periodSliceIds")
                        .EnumerateArray()
                        .Select(item => item.GetString() ?? string.Empty));
                    block.Layout.Values.Add(placement);
                }

                foreach (JsonElement filter in blockElement.GetProperty("filters").EnumerateArray())
                {
                    var placement = new FilterPlacementSpec
                    {
                        Field = filter.GetProperty("field").GetString() ?? string.Empty,
                        IncludeBlank = filter.GetProperty("includeBlank").GetBoolean()
                    };
                    foreach (JsonElement item in filter.GetProperty("selectedValues").EnumerateArray())
                    {
                        placement.SelectedValues.Add(ScalarValue.FromText(item.GetString() ?? string.Empty));
                    }

                    block.Layout.Filters.Add(placement);
                }

                foreach (JsonElement slice in blockElement.GetProperty("periodSlices").EnumerateArray())
                {
                    string start = slice.GetProperty("selectedStart").GetString() ?? string.Empty;
                    string end = slice.GetProperty("selectedEnd").GetString() ?? string.Empty;
                    block.PeriodSlices.Add(new PeriodSliceSpec
                    {
                        Id = slice.GetProperty("id").GetString() ?? string.Empty,
                        Label = slice.GetProperty("label").GetString() ?? string.Empty,
                        Kind = ParsePeriodSliceKind(slice.GetProperty("kind").GetString()),
                        SelectedStart = ParseOptionalDate(start),
                        SelectedEnd = ParseOptionalDate(end),
                        BasedOnSliceId = EmptyToNull(slice.GetProperty("basedOnSliceId").GetString())
                    });
                }

                JsonElement dense = blockElement.GetProperty("denseLayout");
                block.Layout.DenseLayout = new DenseLayoutOptions
                {
                    RepeatRowLabels = dense.GetProperty("repeatRowLabels").GetBoolean(),
                    ShowRowGrandTotals = dense.GetProperty("showRowGrandTotals").GetBoolean(),
                    ShowColumnGrandTotals = dense.GetProperty("showColumnGrandTotals").GetBoolean(),
                    InsertBlankRows = dense.GetProperty("insertBlankRows").GetBoolean(),
                    RowIndent = dense.GetProperty("rowIndent").GetInt32(),
                    FreezeHeaders = dense.GetProperty("freezeHeaders").GetBoolean()
                };

                JsonElement totals = blockElement.GetProperty("grandTotals");
                block.Layout.GrandTotals = new GrandTotalsSpec
                {
                    ShowRows = totals.GetProperty("showRows").GetBoolean(),
                    ShowColumns = totals.GetProperty("showColumns").GetBoolean(),
                    RowPlacement = ParseTotalPlacement(totals.GetProperty("rowPlacement").GetString()),
                    ColumnPlacement = ParseTotalPlacement(totals.GetProperty("columnPlacement").GetString()),
                    RowLabel = totals.GetProperty("rowLabel").GetString() ?? "Grand Total",
                    ColumnLabel = totals.GetProperty("columnLabel").GetString() ?? "Grand Total",
                    StyleId = EmptyToNull(totals.GetProperty("styleId").GetString())
                };
                specification.Blocks.Add(block);
            }

            foreach (JsonElement check in root.GetProperty("checks").EnumerateArray())
            {
                specification.Checks.Add(new ReportCheckSpec
                {
                    Id = check.GetProperty("id").GetString() ?? string.Empty,
                    Kind = ParseCheckKind(check.GetProperty("kind").GetString()),
                    MeasureId = EmptyToNull(check.GetProperty("measureId").GetString()),
                    ComparedMeasureId = EmptyToNull(check.GetProperty("comparedMeasureId").GetString()),
                    Tolerance = check.GetProperty("tolerance").GetDecimal()
                });
            }

            return specification;
        }

        private static FieldPlacementSpec ParseAdvancedFieldPlacement(JsonElement value)
        {
            var result = new FieldPlacementSpec
            {
                Field = value.GetProperty("field").GetString() ?? string.Empty,
                Caption = EmptyToNull(value.GetProperty("caption").GetString()),
                Sort = ParseSort(value.GetProperty("sort").GetString()),
                Subtotals = new SubtotalSpec
                {
                    Mode = string.Equals(
                        value.GetProperty("subtotalMode").GetString(),
                        "none",
                        StringComparison.Ordinal)
                        ? SubtotalMode.None
                        : SubtotalMode.Automatic,
                    Placement = ParseTotalPlacement(
                        value.GetProperty("subtotalPlacement").GetString()),
                    Label = EmptyToNull(value.GetProperty("subtotalLabel").GetString())
                }
            };
            foreach (JsonElement item in value.GetProperty("memberOrder").EnumerateArray())
            {
                result.MemberOrder.Add(ScalarValue.FromText(item.GetString() ?? string.Empty));
            }

            return result;
        }

        private static MeasureExpression ParseAdvancedExpression(
            JsonElement value,
            MeasureValueType resultType,
            IReadOnlyDictionary<string, MeasureValueType> measureTypes)
        {
            string kind = value.GetProperty("kind").GetString() ?? string.Empty;
            switch (kind)
            {
                case "aggregate":
                    return new AggregateMeasureExpression
                    {
                        Field = value.GetProperty("field").GetString() ?? string.Empty,
                        Function = ParseAggregate(value.GetProperty("aggregation").GetString()),
                        PeriodSliceId = EmptyToNull(value.GetProperty("periodSliceId").GetString()),
                        ResultType = resultType
                    };
                case "filteredAggregate":
                    var filtered = new FilteredAggregateMeasureExpression
                    {
                        Field = value.GetProperty("field").GetString() ?? string.Empty,
                        Function = ParseAggregate(value.GetProperty("aggregation").GetString()),
                        PeriodSliceId = EmptyToNull(value.GetProperty("periodSliceId").GetString()),
                        ResultType = resultType
                    };
                    foreach (JsonElement filter in value.GetProperty("filters").EnumerateArray())
                    {
                        var parsedFilter = new MeasureFilterSpec
                        {
                            Field = filter.GetProperty("field").GetString() ?? string.Empty,
                            Operator = ParseMeasureFilterOperator(
                                filter.GetProperty("operator").GetString())
                        };
                        foreach (JsonElement item in filter.GetProperty("values").EnumerateArray())
                        {
                            parsedFilter.Values.Add(ScalarValue.FromText(item.GetString() ?? string.Empty));
                        }

                        filtered.Filters.Add(parsedFilter);
                    }

                    return filtered;
                case "reference":
                    return CreateMeasureReference(
                        value.GetProperty("measureId").GetString(),
                        measureTypes);
                case "constant":
                    return new ConstantMeasureExpression
                    {
                        Value = value.GetProperty("value").GetDecimal(),
                        ResultType = resultType
                    };
                case "binary":
                    return new BinaryMeasureExpression
                    {
                        Operator = ParseBinaryOperator(value.GetProperty("operator").GetString()),
                        Left = CreateMeasureReference(
                            value.GetProperty("leftMeasureId").GetString(),
                            measureTypes),
                        Right = CreateMeasureReference(
                            value.GetProperty("rightMeasureId").GetString(),
                            measureTypes),
                        ReturnBlankOnZeroDenominator = value
                            .GetProperty("returnBlankOnZeroDenominator")
                            .GetBoolean(),
                        ResultType = resultType
                    };
                case "safeDivide":
                    return new SafeDivideMeasureExpression
                    {
                        Numerator = CreateMeasureReference(
                            value.GetProperty("numeratorMeasureId").GetString(),
                            measureTypes),
                        Denominator = CreateMeasureReference(
                            value.GetProperty("denominatorMeasureId").GetString(),
                            measureTypes),
                        OnZero = ParseZeroBehavior(value.GetProperty("onZero").GetString()),
                        AsPercentage = resultType == MeasureValueType.Percentage,
                        ResultType = resultType
                    };
                case "ratio":
                    return new RatioMeasureExpression
                    {
                        Numerator = CreateMeasureReference(
                            value.GetProperty("numeratorMeasureId").GetString(),
                            measureTypes),
                        Denominator = CreateMeasureReference(
                            value.GetProperty("denominatorMeasureId").GetString(),
                            measureTypes),
                        OnZero = ParseZeroBehavior(value.GetProperty("onZero").GetString()),
                        ResultType = resultType
                    };
                case "share":
                    return new ShareMeasureExpression
                    {
                        Part = CreateMeasureReference(
                            value.GetProperty("numeratorMeasureId").GetString(),
                            measureTypes),
                        Whole = CreateMeasureReference(
                            value.GetProperty("denominatorMeasureId").GetString(),
                            measureTypes),
                        OnZero = ParseZeroBehavior(value.GetProperty("onZero").GetString()),
                        ResultType = resultType
                    };
                case "difference":
                    return new DifferenceMeasureExpression
                    {
                        DifferenceKind = ParseDifferenceKind(
                            value.GetProperty("differenceKind").GetString()),
                        Current = CreateMeasureReference(
                            value.GetProperty("currentMeasureId").GetString(),
                            measureTypes),
                        Baseline = CreateMeasureReference(
                            value.GetProperty("baselineMeasureId").GetString(),
                            measureTypes),
                        OnZero = ParseZeroBehavior(value.GetProperty("onZero").GetString()),
                        ResultType = resultType
                    };
                default:
                    throw new InvalidOperationException(
                        "The agent measure expression kind is not supported.");
            }
        }

        private static ReferenceMeasureExpression CreateMeasureReference(
            string? measureId,
            IReadOnlyDictionary<string, MeasureValueType> measureTypes)
        {
            string id = measureId ?? string.Empty;
            if (!measureTypes.TryGetValue(id, out MeasureValueType type))
            {
                throw new InvalidOperationException(
                    "The agent measure expression references an unknown measure.");
            }

            return new ReferenceMeasureExpression
            {
                MeasureId = id,
                ResultType = type
            };
        }

        private ReportSpecV1 Create(
            PeriodMappingSnapshot periodSnapshot,
            IReadOnlyList<PlacementInput> rawInputs,
            SourceProfile sourceProfile,
            string workbookObjectName,
            WorkbookSourceKind sourceKind,
            string reportId,
            ReportOutputMode outputMode,
            IReadOnlyList<TransformStep>? proposedTransforms,
            IReadOnlyList<ManualCalculatedMetricSnapshot>? manualCalculatedMetrics = null,
            IReadOnlyList<ManualReportBlockSnapshot>? manualBlocks = null,
            ManualLayoutSnapshot? manualLayout = null,
            IReadOnlyList<ManualCheckSnapshot>? manualChecks = null)
        {
            if (sourceProfile == null) throw new ArgumentNullException(nameof(sourceProfile));
            if (string.IsNullOrWhiteSpace(workbookObjectName))
            {
                throw new ArgumentException("A workbook object name is required.", nameof(workbookObjectName));
            }

            string ownershipId = BoundedId("owner_" + reportId);
            string blockOwnershipId = BoundedId("block_owner_" + reportId);
            PeriodMappingSpec? periodMapping = ResolvePeriodMapping(periodSnapshot, sourceProfile);
            List<PlacementInput> inputs = NormalizeWidePlacements(rawInputs, periodMapping);
            var specification = new ReportSpecV1
            {
                Id = BoundedId(reportId),
                Name = "Saved report setup",
                OwnershipId = ownershipId,
                Source = new WorkbookSourceSpec
                {
                    Kind = sourceKind,
                    WorkbookObjectName = workbookObjectName,
                    HeaderRowCount = 1,
                    Fingerprint = SourceFingerprint.FromHeaders(
                        sourceProfile.Columns
                            .OrderBy(column => column.Index)
                            .Select(column => column.Name))
                },
                PeriodMapping = periodMapping
            };
            if (proposedTransforms != null)
            {
                specification.Transforms.AddRange(proposedTransforms);
            }

            AddRequiredPeriodPreparation(specification, sourceProfile);
            ValidateTransformTypeSafety(
                sourceProfile,
                periodMapping,
                specification.Transforms);

            specification.Styles.Add(new PresentationStyleSpec
            {
                Id = "report_header",
                Bold = true,
                FontColor = "#FFFFFF",
                FillColor = "#1F5D50",
                BottomBorder = true
            });

            var block = new ReportBlockSpec
            {
                Id = "report_block",
                OwnershipId = blockOwnershipId,
                Title = "Management report",
                WorksheetName = "Report",
                AnchorCell = "A1",
                OutputMode = outputMode,
                OwnedExtent = new OwnedRangeExtentSpec
                {
                    RowCount = 100000,
                    ColumnCount = 512
                },
                HeaderStyleId = "report_header"
            };
            specification.Blocks.Add(block);

            foreach (PlacementInput input in inputs.Where(value => value.Bucket == PlacementBucket.Rows))
            {
                block.Layout.Rows.Add(ToFieldPlacement(input));
            }

            foreach (PlacementInput input in inputs.Where(value => value.Bucket == PlacementBucket.Columns))
            {
                block.Layout.Columns.Add(ToFieldPlacement(input));
            }

            var measureIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int measureIndex = 0;
            foreach (PlacementInput input in inputs.Where(value => value.Bucket == PlacementBucket.Values))
            {
                AggregateFunction aggregate = ParseAggregate(input.Aggregate);
                bool count = aggregate == AggregateFunction.Count ||
                    aggregate == AggregateFunction.DistinctCount;
                if (!count && !IsNumericField(
                        input.Field,
                        sourceProfile,
                        periodMapping,
                        specification.Transforms))
                {
                    throw new InvalidOperationException(
                        "Sum, average, minimum, and maximum Values require a numeric source column.");
                }

                MeasureValueType resultType = count
                    ? MeasureValueType.WholeNumber
                    : MeasureValueType.Number;
                string measureId = UniqueId(
                    "measure_" + input.Field,
                    ++measureIndex,
                    measureIds);
                string label = AggregateLabel(aggregate) + " of " + input.Field;
                string? numberFormat = input.NumberFormat ?? (count ? "#,##0" : "#,##0.00");
                specification.Measures.Add(new MeasureDefinition
                {
                    Id = measureId,
                    Label = label,
                    ValueType = resultType,
                    NumberFormat = numberFormat,
                    Expression = new AggregateMeasureExpression
                    {
                        Field = input.Field,
                        Function = aggregate,
                        ResultType = resultType
                    }
                });
                block.Layout.Values.Add(new ValuePlacementSpec
                {
                    MeasureId = measureId,
                    Caption = label,
                    NumberFormat = numberFormat
                });
            }

            foreach (PlacementInput input in inputs.Where(value => value.Bucket == PlacementBucket.Filters))
            {
                var filter = new FilterPlacementSpec { Field = input.Field };
                foreach (string selectedValue in input.FilterValues)
                {
                    filter.SelectedValues.Add(ScalarValue.FromText(selectedValue));
                }

                block.Layout.Filters.Add(filter);
            }

            if (block.Layout.Rows.Count == 0)
            {
                throw new InvalidOperationException("Place at least one source column in Rows.");
            }

            if (block.Layout.Values.Count == 0)
            {
                throw new InvalidOperationException("Place at least one source column in Values.");
            }

            if (manualCalculatedMetrics != null && manualCalculatedMetrics.Count != 0)
            {
                IReadOnlyList<ReportOutputMode> actualOutputModes = manualBlocks != null && manualBlocks.Count != 0
                    ? manualBlocks.Select(item => ParseOutputMode(item.OutputStyle)).ToArray()
                    : new[] { outputMode };
                if (actualOutputModes.Any(mode => mode != ReportOutputMode.DenseGrid))
                {
                    throw new InvalidOperationException(
                        "Calculated metrics require every report block to use Dense management block.");
                }

                block.OutputMode = ReportOutputMode.DenseGrid;
                foreach (ManualCalculatedMetricSnapshot metric in manualCalculatedMetrics)
                {
                    if (string.Equals(metric.Kind, "Weighted average", StringComparison.OrdinalIgnoreCase) &&
                        (!IsNumericField(metric.Primary, sourceProfile, periodMapping, specification.Transforms) ||
                         !IsNumericField(metric.Secondary, sourceProfile, periodMapping, specification.Transforms)))
                    {
                        throw new InvalidOperationException(
                            "Weighted average requires numeric value and weight columns after preparation.");
                    }
                }

                new ManualCalculatedMetricTranslator().Append(
                    specification,
                    block,
                    manualCalculatedMetrics,
                    sourceProfile.Columns
                        .OrderBy(column => column.Index)
                        .Select(column => column.Name)
                        .ToArray());
            }

            if (manualBlocks != null && manualBlocks.Count != 0)
            {
                new ManualLayoutTranslator().Apply(
                    block,
                    manualBlocks,
                    manualLayout ?? new ManualLayoutSnapshot(),
                    manualChecks ?? Array.Empty<ManualCheckSnapshot>(),
                    specification);
            }
            else if (manualChecks != null && manualChecks.Count != 0)
            {
                new ManualLayoutTranslator().Apply(
                    block,
                    new[]
                    {
                        new ManualReportBlockSnapshot(
                            block.Title ?? "Management report",
                            block.WorksheetName,
                            block.AnchorCell,
                            ToUiOutputStyle(block.OutputMode))
                    },
                    manualLayout ?? new ManualLayoutSnapshot(),
                    manualChecks,
                    specification);
            }

            return specification;
        }

        private static void AddRequiredPeriodPreparation(
            ReportSpecV1 specification,
            SourceProfile sourceProfile)
        {
            PeriodMappingSpec? mapping = specification.PeriodMapping;
            if (mapping == null ||
                specification.Transforms.OfType<NormalizePeriodsTransform>().Any())
            {
                return;
            }

            if (mapping.Kind == PeriodMappingKind.LongDateColumn)
            {
                SourceColumnProfile? column = sourceProfile.FindColumn(mapping.DateColumn ?? string.Empty);
                bool alreadyTyped = specification.Transforms
                    .OfType<ChangeColumnTypeTransform>()
                    .Any(change => string.Equals(
                        change.Column,
                        mapping.DateColumn,
                        StringComparison.OrdinalIgnoreCase) &&
                        (change.DataType == ColumnDataType.Date ||
                         change.DataType == ColumnDataType.DateTime));
                if (alreadyTyped)
                {
                    return;
                }

                if (column != null &&
                    (column.InferredType == SourceValueType.Date ||
                     column.InferredType == SourceValueType.DateTime))
                {
                    specification.Transforms.Insert(0, new ChangeColumnTypeTransform
                    {
                        Id = "type_period_as_date",
                        Column = mapping.DateColumn ?? string.Empty,
                        DataType = ColumnDataType.Date
                    });
                    return;
                }
            }

            var normalization = new NormalizePeriodsTransform
            {
                Id = "normalize_periods",
                PeriodMappingId = mapping.Id
            };
            if (mapping.Kind == PeriodMappingKind.LongDateColumn)
            {
                specification.Transforms.Insert(0, normalization);
            }
            else
            {
                specification.Transforms.Add(normalization);
            }
        }

        private static List<PlacementInput> NormalizeWidePlacements(
            IReadOnlyList<PlacementInput> inputs,
            PeriodMappingSpec? mapping)
        {
            var result = new List<PlacementInput>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            bool isWide = mapping != null && mapping.Kind != PeriodMappingKind.LongDateColumn;
            var mappedColumns = isWide
                ? new HashSet<string>(mapping!.Columns.Select(column => column.SourceColumn), StringComparer.OrdinalIgnoreCase)
                : new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (PlacementInput input in inputs)
            {
                var normalized = input.Clone();
                if (isWide && mappedColumns.Contains(normalized.Field))
                {
                    if (normalized.Bucket != PlacementBucket.Values)
                    {
                        continue;
                    }

                    normalized.Field = mapping!.ValueColumnName;
                }

                string key = normalized.Bucket + "|" + normalized.Field;
                if (seen.Add(key))
                {
                    result.Add(normalized);
                }
            }

            if (isWide)
            {
                AddAutomaticColumn(result, seen, mapping!.PeriodColumnName);
                if (mapping.Kind == PeriodMappingKind.MetricMonthHeaders)
                {
                    AddAutomaticColumn(result, seen, mapping.MetricColumnName);
                }
            }

            return result;
        }

        private static void AddAutomaticColumn(
            ICollection<PlacementInput> placements,
            ISet<string> seen,
            string field)
        {
            bool alreadyPlaced = seen.Contains(PlacementBucket.Rows + "|" + field) ||
                seen.Contains(PlacementBucket.Columns + "|" + field);
            if (alreadyPlaced)
            {
                return;
            }

            seen.Add(PlacementBucket.Columns + "|" + field);
            placements.Add(new PlacementInput
            {
                Bucket = PlacementBucket.Columns,
                Field = field,
                Subtotals = "hide",
                Sort = "ascending"
            });
        }

        private static FieldPlacementSpec ToFieldPlacement(PlacementInput input)
        {
            var result = new FieldPlacementSpec
            {
                Field = input.Field,
                Sort = ParseSort(input.Sort),
                Subtotals = new SubtotalSpec
                {
                    Mode = string.Equals(input.Subtotals, "hide", StringComparison.OrdinalIgnoreCase)
                        ? SubtotalMode.None
                        : SubtotalMode.Automatic,
                    Placement = string.Equals(
                        input.SubtotalPlacement,
                        "Before members",
                        StringComparison.OrdinalIgnoreCase)
                        ? TotalPlacement.BeforeMembers
                        : TotalPlacement.AfterMembers
                }
            };
            foreach (string member in input.MemberOrder)
            {
                result.MemberOrder.Add(ScalarValue.FromText(member));
            }

            return result;
        }

        private static AggregateFunction ParseAggregate(string? aggregate)
        {
            switch ((aggregate ?? "sum").Trim().ToLowerInvariant())
            {
                case "count": return AggregateFunction.Count;
                case "distinct count":
                case "distinctcount": return AggregateFunction.DistinctCount;
                case "average": return AggregateFunction.Average;
                case "min":
                case "minimum": return AggregateFunction.Minimum;
                case "max":
                case "maximum": return AggregateFunction.Maximum;
                case "sum": return AggregateFunction.Sum;
                default: throw new InvalidOperationException("The requested Value calculation is not supported.");
            }
        }

        private static MeasureValueType ParseMeasureValueType(string? value)
        {
            switch (value)
            {
                case "wholeNumber": return MeasureValueType.WholeNumber;
                case "currency": return MeasureValueType.Currency;
                case "percentage": return MeasureValueType.Percentage;
                default: return MeasureValueType.Number;
            }
        }

        private static BinaryMeasureOperator ParseBinaryOperator(string? value)
        {
            switch (value)
            {
                case "subtract": return BinaryMeasureOperator.Subtract;
                case "multiply": return BinaryMeasureOperator.Multiply;
                case "divide": return BinaryMeasureOperator.Divide;
                default: return BinaryMeasureOperator.Add;
            }
        }

        private static DifferenceKind ParseDifferenceKind(string? value)
        {
            switch (value)
            {
                case "percentage": return DifferenceKind.Percentage;
                case "percentagePoints": return DifferenceKind.PercentagePoints;
                default: return DifferenceKind.Absolute;
            }
        }

        private static ZeroDenominatorBehavior ParseZeroBehavior(string? value)
        {
            switch (value)
            {
                case "zero": return ZeroDenominatorBehavior.Zero;
                case "error": return ZeroDenominatorBehavior.Error;
                default: return ZeroDenominatorBehavior.Blank;
            }
        }

        private static MeasureFilterOperator ParseMeasureFilterOperator(string? value)
        {
            switch (value)
            {
                case "notEqual": return MeasureFilterOperator.NotEqual;
                case "greaterThan": return MeasureFilterOperator.GreaterThan;
                case "greaterThanOrEqual": return MeasureFilterOperator.GreaterThanOrEqual;
                case "lessThan": return MeasureFilterOperator.LessThan;
                case "lessThanOrEqual": return MeasureFilterOperator.LessThanOrEqual;
                case "in": return MeasureFilterOperator.In;
                case "notIn": return MeasureFilterOperator.NotIn;
                case "isBlank": return MeasureFilterOperator.IsBlank;
                case "isNotBlank": return MeasureFilterOperator.IsNotBlank;
                default: return MeasureFilterOperator.Equal;
            }
        }

        private static PeriodSliceKind ParsePeriodSliceKind(string? value)
        {
            switch (value)
            {
                case "selected": return PeriodSliceKind.Selected;
                case "prior": return PeriodSliceKind.Prior;
                case "samePeriodPriorYear": return PeriodSliceKind.SamePeriodPriorYear;
                default: return PeriodSliceKind.Current;
            }
        }

        private static TotalPlacement ParseTotalPlacement(string? value)
        {
            return string.Equals(value, "beforeMembers", StringComparison.Ordinal)
                ? TotalPlacement.BeforeMembers
                : TotalPlacement.AfterMembers;
        }

        private static ReportOutputMode ParseAgentOutputMode(string? value)
        {
            switch (value)
            {
                case "standardMatrix": return ReportOutputMode.StandardMatrix;
                case "metricStack": return ReportOutputMode.MetricStack;
                default: return ReportOutputMode.DenseGrid;
            }
        }

        private static HorizontalAlignment ParseHorizontalAlignment(string? value)
        {
            switch (value)
            {
                case "left": return HorizontalAlignment.Left;
                case "center": return HorizontalAlignment.Center;
                case "right": return HorizontalAlignment.Right;
                default: return HorizontalAlignment.General;
            }
        }

        private static ReportCheckKind ParseCheckKind(string? value)
        {
            switch (value)
            {
                case "noTruncation": return ReportCheckKind.NoTruncation;
                case "requiredValues": return ReportCheckKind.RequiredValues;
                case "nonNegative": return ReportCheckKind.NonNegative;
                case "balance": return ReportCheckKind.Balance;
                default: return ReportCheckKind.TotalPreservation;
            }
        }

        private static DateTime? ParseOptionalDate(string? value)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;
            return DateTime.ParseExact(
                value,
                "yyyy-MM-dd",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None);
        }

        private static string? EmptyToNull(string? value)
        {
            return string.IsNullOrEmpty(value) ? null : value;
        }

        internal static ReportOutputMode ParseOutputMode(string? outputStyle)
        {
            if (string.Equals(outputStyle, "Standard matrix", StringComparison.Ordinal))
            {
                return ReportOutputMode.StandardMatrix;
            }

            if (string.Equals(outputStyle, "Metric stack", StringComparison.Ordinal))
            {
                return ReportOutputMode.MetricStack;
            }

            return ReportOutputMode.DenseGrid;
        }

        private static string ToUiOutputStyle(ReportOutputMode outputMode)
        {
            switch (outputMode)
            {
                case ReportOutputMode.StandardMatrix: return "Standard matrix";
                case ReportOutputMode.MetricStack: return "Metric stack";
                default: return "Dense management block";
            }
        }

        private static PeriodMappingSnapshot ToUiPeriodMapping(PeriodMappingSpec? mapping)
        {
            if (mapping == null)
            {
                return new PeriodMappingSnapshot(
                    "No period columns",
                    string.Empty,
                    string.Empty,
                    null,
                    Array.Empty<WideHeaderMappingRowSnapshot>());
            }

            if (mapping.Kind == PeriodMappingKind.LongDateColumn)
            {
                return new PeriodMappingSnapshot(
                    "Date column",
                    mapping.DateColumn ?? string.Empty,
                    string.Empty,
                    mapping.ReportingYear,
                    Array.Empty<WideHeaderMappingRowSnapshot>());
            }

            var rows = mapping.Columns.Select(column =>
                new WideHeaderMappingRowSnapshot(
                    column.SourceColumn,
                    column.Year.HasValue
                        ? new DateTime(column.Year.Value, column.Month, 1)
                            .ToString("yyyy-MM", CultureInfo.InvariantCulture)
                        : column.Month.ToString("00", CultureInfo.InvariantCulture),
                    string.IsNullOrWhiteSpace(column.Metric) ? "Value" : column.Metric!,
                    1d)).ToList();
            return new PeriodMappingSnapshot(
                "Wide period headers",
                string.Empty,
                "Automatic month and metric-month detection",
                mapping.ReportingYear,
                rows);
        }

        private static IReadOnlyList<FieldPlacementSnapshot> ProjectManualPlacements(
            ReportSpecV1 specification,
            ReportBlockSpec block,
            IReadOnlyDictionary<string, MeasureDefinition> measures)
        {
            if (block.PeriodSlices.Count != 0 ||
                block.Headers.Count != 0 ||
                block.Spacers.Count != 0 ||
                !HasSupportedManualPresentation(specification, block))
            {
                throw new InvalidOperationException(
                    "The saved setup uses advanced layout features that the bounded manual builder cannot safely edit.");
            }

            if (block.Layout.Rows.Count == 0 || block.Layout.Values.Count == 0)
            {
                throw new InvalidOperationException(
                    "The saved setup does not contain the Rows and Values required by the bounded manual builder.");
            }

            var placements = new List<FieldPlacementSnapshot>();
            foreach (FieldPlacementSpec row in block.Layout.Rows)
            {
                placements.Add(ToUiFieldPlacement(PlacementBucket.Rows, row));
            }

            foreach (FieldPlacementSpec column in block.Layout.Columns)
            {
                placements.Add(ToUiFieldPlacement(PlacementBucket.Columns, column));
            }

            var usedMeasures = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var expectedMeasureIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var measureIndex = 0;
            foreach (ValuePlacementSpec value in block.Layout.Values)
            {
                if (value.PeriodSliceIds.Count != 0 ||
                    !string.IsNullOrWhiteSpace(value.StyleId) ||
                    !measures.TryGetValue(value.MeasureId, out MeasureDefinition? measure) ||
                    !(measure.Expression is AggregateMeasureExpression aggregate) ||
                    !string.IsNullOrWhiteSpace(aggregate.PeriodSliceId) ||
                    !string.Equals(value.Caption, measure.Label, StringComparison.Ordinal) ||
                    !string.Equals(value.NumberFormat, measure.NumberFormat, StringComparison.Ordinal) ||
                    !IsManualAggregateMeasure(measure, aggregate) ||
                    !usedMeasures.Add(measure.Id))
                {
                    throw new InvalidOperationException(
                        "The saved setup contains a calculated, sliced, duplicated, or otherwise unsupported Value that the bounded manual builder cannot safely edit.");
                }

                string expectedMeasureId = UniqueId(
                    "measure_" + aggregate.Field,
                    ++measureIndex,
                    expectedMeasureIds);
                if (!string.Equals(measure.Id, expectedMeasureId, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "The saved setup contains a Value identity that the bounded manual builder cannot preserve exactly.");
                }

                placements.Add(new FieldPlacementSnapshot(
                    PlacementBucket.Values,
                    aggregate.Field,
                    ToUiAggregate(aggregate.Function),
                    numberFormat: measure.NumberFormat ?? "General"));
            }

            if (usedMeasures.Count != measures.Count)
            {
                throw new InvalidOperationException(
                    "The saved setup contains metrics that are not represented by the bounded manual Values editor.");
            }

            foreach (FilterPlacementSpec filter in block.Layout.Filters)
            {
                if (filter.IncludeBlank || filter.SelectedValues.Any(value =>
                        value.Kind != ScalarValueKind.Text ||
                        value.Text == null ||
                        value.Text.Contains(";")))
                {
                    throw new InvalidOperationException(
                        "The saved setup contains a Filter that the bounded manual builder cannot safely edit.");
                }

                string[] selectedValues = filter.SelectedValues
                    .Select(value => value.Text!)
                    .ToArray();
                placements.Add(new FieldPlacementSnapshot(
                    PlacementBucket.Filters,
                    filter.Field,
                    selectedValues.Length == 0 ? "All" : string.Join("; ", selectedValues),
                    selectedValues: selectedValues));
            }

            return placements;
        }

        private static bool ManualPlacementsEqual(
            IReadOnlyList<FieldPlacementSnapshot> expected,
            IReadOnlyList<FieldPlacementSnapshot> actual)
        {
            if (expected.Count != actual.Count)
            {
                return false;
            }

            for (var index = 0; index < expected.Count; index++)
            {
                FieldPlacementSnapshot left = expected[index];
                FieldPlacementSnapshot right = actual[index];
                if (left.Bucket != right.Bucket ||
                    !string.Equals(left.ColumnName, right.ColumnName, StringComparison.Ordinal) ||
                    !string.Equals(left.Setting, right.Setting, StringComparison.Ordinal) ||
                    left.ShowSubtotals != right.ShowSubtotals ||
                    !left.SelectedValues.SequenceEqual(right.SelectedValues, StringComparer.Ordinal) ||
                    !string.Equals(left.SubtotalPlacement, right.SubtotalPlacement, StringComparison.Ordinal) ||
                    !left.MemberOrder.SequenceEqual(right.MemberOrder, StringComparer.Ordinal) ||
                    !string.Equals(left.NumberFormat, right.NumberFormat, StringComparison.Ordinal))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool ValueMeasureOrderEqual(
            IReadOnlyList<ValuePlacementSpec> expected,
            IReadOnlyList<ValuePlacementSpec> actual)
        {
            return expected.Select(value => value.MeasureId)
                .SequenceEqual(actual.Select(value => value.MeasureId), StringComparer.Ordinal);
        }

        private static bool ManualLayoutsEqual(
            ManualLayoutSnapshot expected,
            ManualLayoutSnapshot actual)
        {
            return expected.RepeatRowLabels == actual.RepeatRowLabels &&
                expected.InsertBlankRows == actual.InsertBlankRows &&
                expected.FreezeHeaders == actual.FreezeHeaders &&
                expected.ShowRowGrandTotals == actual.ShowRowGrandTotals &&
                expected.ShowColumnGrandTotals == actual.ShowColumnGrandTotals &&
                expected.RowIndent == actual.RowIndent &&
                string.Equals(
                    expected.RowGrandTotalLabel,
                    actual.RowGrandTotalLabel,
                    StringComparison.Ordinal) &&
                string.Equals(
                    expected.ColumnGrandTotalLabel,
                    actual.ColumnGrandTotalLabel,
                    StringComparison.Ordinal);
        }

        private static FieldPlacementSnapshot ToUiFieldPlacement(
            PlacementBucket bucket,
            FieldPlacementSpec placement)
        {
            if (!string.IsNullOrWhiteSpace(placement.Caption) ||
                placement.GroupBuckets.Count != 0 ||
                placement.TopN != null ||
                !string.IsNullOrWhiteSpace(placement.Subtotals.Label) ||
                !string.IsNullOrWhiteSpace(placement.Subtotals.StyleId) ||
                placement.MemberOrder.Any(value =>
                    value.Kind != ScalarValueKind.Text || value.Text == null))
            {
                throw new InvalidOperationException(
                    "The saved setup contains advanced field placement settings that the bounded manual builder cannot safely edit.");
            }

            string setting;
            switch (placement.Sort)
            {
                case SortDirection.Ascending:
                    setting = "Ascending";
                    break;
                case SortDirection.Descending:
                    setting = "Descending";
                    break;
                default:
                    setting = "Default order";
                    break;
            }

            return new FieldPlacementSnapshot(
                bucket,
                placement.Field,
                setting,
                placement.Subtotals.Mode != SubtotalMode.None,
                subtotalPlacement: ToUiTotalPlacement(placement.Subtotals.Placement),
                memberOrder: placement.MemberOrder
                    .Select(ToDisplayScalar)
                    .Where(value => value != null)
                    .Cast<string>()
                    .ToArray());
        }

        private static string ToUiAggregate(AggregateFunction aggregate)
        {
            switch (aggregate)
            {
                case AggregateFunction.Sum: return "Sum";
                case AggregateFunction.Count: return "Count";
                case AggregateFunction.DistinctCount: return "Distinct count";
                case AggregateFunction.Average: return "Average";
                case AggregateFunction.Minimum: return "Minimum";
                case AggregateFunction.Maximum: return "Maximum";
                default:
                    throw new InvalidOperationException(
                        "The saved setup uses a Value calculation that the bounded manual builder cannot safely edit.");
            }
        }

        private static string ToUiTotalPlacement(TotalPlacement placement)
        {
            return placement == TotalPlacement.BeforeMembers
                ? "Before members"
                : "After members";
        }

        private static string ToUiSort(SortDirection sort)
        {
            switch (sort)
            {
                case SortDirection.Ascending: return "Ascending";
                case SortDirection.Descending: return "Descending";
                default: return "Default order";
            }
        }

        private static string? ToDisplayScalar(ScalarValue value)
        {
            switch (value.Kind)
            {
                case ScalarValueKind.Text: return value.Text;
                case ScalarValueKind.Number:
                    return value.Number?.ToString(CultureInfo.InvariantCulture);
                case ScalarValueKind.Boolean:
                    return value.Boolean?.ToString();
                case ScalarValueKind.Date:
                    return value.Temporal?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                case ScalarValueKind.DateTime:
                    return value.Temporal?.ToString("O", CultureInfo.InvariantCulture);
                case ScalarValueKind.Null: return null;
                default: return null;
            }
        }

        private static bool IsManualAggregateMeasure(
            MeasureDefinition measure,
            AggregateMeasureExpression aggregate)
        {
            bool isCount = aggregate.Function == AggregateFunction.Count ||
                aggregate.Function == AggregateFunction.DistinctCount;
            string expectedLabel = AggregateLabel(aggregate.Function) + " of " + aggregate.Field;
            MeasureValueType expectedType = isCount
                ? MeasureValueType.WholeNumber
                : MeasureValueType.Number;
            return string.Equals(measure.Label, expectedLabel, StringComparison.Ordinal) &&
                IsSupportedManualNumberFormat(measure.NumberFormat) &&
                measure.ValueType == expectedType;
        }

        private static bool IsSupportedManualNumberFormat(string? numberFormat)
        {
            return string.Equals(numberFormat, "General", StringComparison.Ordinal) ||
                string.Equals(numberFormat, "#,##0", StringComparison.Ordinal) ||
                string.Equals(numberFormat, "#,##0.00", StringComparison.Ordinal) ||
                string.Equals(numberFormat, "0.0%", StringComparison.Ordinal) ||
                string.Equals(numberFormat, "0.00%", StringComparison.Ordinal);
        }

        private static bool HasSupportedManualPresentation(
            ReportSpecV1 specification,
            ReportBlockSpec block)
        {
            if (!string.Equals(block.HeaderStyleId, "report_header", StringComparison.Ordinal) ||
                !string.IsNullOrWhiteSpace(block.BodyStyleId) ||
                !string.IsNullOrWhiteSpace(block.SubtotalStyleId) ||
                !string.IsNullOrWhiteSpace(block.GrandTotalStyleId) ||
                specification.Styles.Count != 1)
            {
                return false;
            }

            DenseLayoutOptions dense = block.Layout.DenseLayout;
            GrandTotalsSpec totals = block.Layout.GrandTotals;
            if (dense.ShowRowGrandTotals != totals.ShowRows ||
                dense.ShowColumnGrandTotals != totals.ShowColumns ||
                totals.RowPlacement != TotalPlacement.AfterMembers ||
                totals.ColumnPlacement != TotalPlacement.AfterMembers ||
                !string.IsNullOrWhiteSpace(totals.StyleId))
            {
                return false;
            }

            PresentationStyleSpec style = specification.Styles[0];
            return string.Equals(style.Id, "report_header", StringComparison.Ordinal) &&
                style.Bold &&
                !style.Italic &&
                string.Equals(style.FontColor, "#FFFFFF", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(style.FillColor, "#1F5D50", StringComparison.OrdinalIgnoreCase) &&
                style.HorizontalAlignment == HorizontalAlignment.General &&
                string.IsNullOrWhiteSpace(style.NumberFormat) &&
                !style.DecimalPlaces.HasValue &&
                !style.TopBorder &&
                style.BottomBorder;
        }

        private static IReadOnlyList<ManualReportBlockSnapshot> ToManualBlocks(
            IReadOnlyList<ReportBlockSpec> blocks)
        {
            return blocks.Select(block => new ManualReportBlockSnapshot(
                block.Title ?? string.Empty,
                block.WorksheetName,
                block.AnchorCell,
                ToUiOutputStyle(block.OutputMode),
                stableId: block.Id,
                ownedRows: block.OwnedExtent.RowCount,
                ownedColumns: block.OwnedExtent.ColumnCount,
                canonicalBlockId: block.Id,
                canonicalOwnershipId: block.OwnershipId)).ToArray();
        }

        private static ManualLayoutSnapshot ToManualLayout(ReportBlockSpec block)
        {
            DenseLayoutOptions dense = block.Layout.DenseLayout;
            GrandTotalsSpec totals = block.Layout.GrandTotals;
            return new ManualLayoutSnapshot
            {
                RepeatRowLabels = dense.RepeatRowLabels,
                InsertBlankRows = dense.InsertBlankRows,
                FreezeHeaders = dense.FreezeHeaders,
                ShowRowGrandTotals = dense.ShowRowGrandTotals,
                ShowColumnGrandTotals = dense.ShowColumnGrandTotals,
                RowIndent = dense.RowIndent,
                RowGrandTotalLabel = totals.RowLabel,
                ColumnGrandTotalLabel = totals.ColumnLabel
            };
        }

        private static IReadOnlyList<ManualCheckSnapshot> ToManualChecks(
            IReadOnlyList<ReportCheckSpec> checks)
        {
            return checks.Select(check => new ManualCheckSnapshot(
                ToManualCheckKind(check),
                check.MeasureId ?? string.Empty,
                check.ComparedMeasureId ?? string.Empty,
                check.Tolerance)).ToArray();
        }

        private static string ToManualCheckKind(ReportCheckSpec check)
        {
            switch (check.Kind)
            {
                case ReportCheckKind.TotalPreservation:
                    return check.Id.IndexOf("rendered_output", StringComparison.OrdinalIgnoreCase) >= 0
                        ? "Rendered output"
                        : "Total preservation";
                case ReportCheckKind.NoTruncation: return "No truncation";
                case ReportCheckKind.RequiredValues: return "Required values";
                case ReportCheckKind.NonNegative: return "Non-negative";
                case ReportCheckKind.Balance: return "Balance";
                default:
                    throw new InvalidOperationException(
                        "The saved setup contains a check that the bounded manual builder cannot safely edit.");
            }
        }

        private static void EnsureTransformsCanRoundTrip(ReportSpecV1 specification)
        {
            if (specification.PeriodMapping == null)
            {
                if (specification.Transforms.Count != 0)
                {
                    throw new InvalidOperationException(
                        "The saved setup contains transformations that the bounded manual builder cannot safely edit.");
                }

                return;
            }

            bool isExpectedNormalize = specification.Transforms.Count == 1 &&
                specification.Transforms[0] is NormalizePeriodsTransform normalize &&
                string.Equals(
                    normalize.PeriodMappingId,
                    specification.PeriodMapping.Id,
                    StringComparison.Ordinal);
            bool isExpectedDateType = specification.PeriodMapping.Kind == PeriodMappingKind.LongDateColumn &&
                specification.Transforms.Count == 1 &&
                specification.Transforms[0] is ChangeColumnTypeTransform change &&
                string.Equals(
                    change.Column,
                    specification.PeriodMapping.DateColumn,
                    StringComparison.OrdinalIgnoreCase) &&
                change.DataType == ColumnDataType.Date;
            if (!isExpectedNormalize && !isExpectedDateType)
            {
                throw new InvalidOperationException(
                    "The saved setup contains transformations that the bounded manual builder cannot safely edit.");
            }
        }

        private static PeriodGrain ResolveSelectedPeriodGrain(SourceColumnProfile column)
        {
            var grains = new List<PeriodGrain>();
            if (column.DayGrainCount > 0) grains.Add(PeriodGrain.Day);
            if (column.MonthGrainCount > 0) grains.Add(PeriodGrain.Month);
            if (column.QuarterGrainCount > 0) grains.Add(PeriodGrain.Quarter);
            if (grains.Count != 1)
            {
                throw new InvalidOperationException(
                    "The selected period column mixes day, month, or quarter values. Normalize it to one grain before building.");
            }

            return grains[0];
        }

        private static string AggregateLabel(AggregateFunction aggregate)
        {
            switch (aggregate)
            {
                case AggregateFunction.Count: return "Count";
                case AggregateFunction.DistinctCount: return "Distinct count";
                case AggregateFunction.Average: return "Average";
                case AggregateFunction.Minimum: return "Minimum";
                case AggregateFunction.Maximum: return "Maximum";
                default: return "Sum";
            }
        }

        private static SortDirection ParseSort(string? sort)
        {
            switch ((sort ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "ascending": return SortDirection.Ascending;
                case "descending": return SortDirection.Descending;
                default: return SortDirection.SourceOrder;
            }
        }

        private static bool IsNumericField(
            string field,
            SourceProfile profile,
            PeriodMappingSpec? mapping,
            IReadOnlyList<TransformStep> transforms)
        {
            Dictionary<string, SourceValueType> fieldTypes = BuildFieldTypeIndex(
                profile,
                mapping,
                transforms,
                validateOperations: false);
            return fieldTypes.TryGetValue(field, out SourceValueType type) && IsNumericType(type);
        }

        private static void ValidateTransformTypeSafety(
            SourceProfile profile,
            PeriodMappingSpec? mapping,
            IReadOnlyList<TransformStep> transforms)
        {
            BuildFieldTypeIndex(profile, mapping, transforms, validateOperations: true);
        }

        private static Dictionary<string, SourceValueType> BuildFieldTypeIndex(
            SourceProfile profile,
            PeriodMappingSpec? mapping,
            IReadOnlyList<TransformStep> transforms,
            bool validateOperations)
        {
            var fieldTypes = profile.Columns.ToDictionary(
                column => column.Name,
                column => column.InferredType,
                StringComparer.OrdinalIgnoreCase);
            foreach (var transform in transforms)
            {
                switch (transform)
                {
                    case SelectColumnsTransform select:
                        RetainFieldTypes(fieldTypes, select.Columns);
                        break;
                    case KeepColumnsTransform keep:
                        RetainFieldTypes(fieldTypes, keep.Columns);
                        break;
                    case RemoveColumnsTransform remove:
                        foreach (string column in remove.Columns) fieldTypes.Remove(column);
                        break;
                    case RenameColumnTransform rename:
                        if (fieldTypes.TryGetValue(rename.From, out var renamedType))
                        {
                            fieldTypes.Remove(rename.From);
                            fieldTypes[rename.To] = renamedType;
                        }

                        break;
                    case ChangeColumnTypeTransform change:
                        fieldTypes[change.Column] = ToSourceValueType(change.DataType);
                        break;
                    case TrimTextTransform trim:
                        foreach (string column in trim.Columns) fieldTypes[column] = SourceValueType.Text;
                        break;
                    case ReplaceValueTransform replace:
                        MergeReplacementType(fieldTypes, replace.Column, replace.ReplaceWith);
                        break;
                    case NormalizeBlanksTransform blanks:
                        foreach (string column in blanks.Columns)
                        {
                            MergeReplacementType(fieldTypes, column, blanks.Replacement);
                        }

                        break;
                    case NormalizeErrorsTransform errors:
                        foreach (string column in errors.Columns)
                        {
                            MergeReplacementType(fieldTypes, column, errors.Replacement);
                        }

                        break;
                    case MapValuesTransform map:
                        foreach (ValueMapEntry entry in map.Entries)
                        {
                            MergeReplacementType(fieldTypes, map.Column, entry.To);
                        }

                        break;
                    case DerivePeriodPartsTransform derive:
                        if (validateOperations &&
                            (!fieldTypes.TryGetValue(derive.DateColumn, out SourceValueType dateType) ||
                             dateType != SourceValueType.Date && dateType != SourceValueType.DateTime))
                        {
                            throw new InvalidOperationException(
                                "Derive period parts requires a Date or DateTime column after earlier preparation steps.");
                        }

                        foreach (DerivedPeriodColumnSpec column in derive.Columns)
                        {
                            fieldTypes[column.OutputColumn] =
                                column.Part == DerivedPeriodPart.Year || column.Part == DerivedPeriodPart.MonthNumber
                                    ? SourceValueType.WholeNumber
                                    : SourceValueType.Text;
                        }

                        break;
                    case AddArithmeticColumnTransform arithmetic:
                        if (validateOperations)
                        {
                            DemandNumericOperand(arithmetic.Left, fieldTypes);
                            DemandNumericOperand(arithmetic.Right, fieldTypes);
                        }

                        fieldTypes[arithmetic.OutputColumn] = arithmetic.ResultType == ColumnDataType.WholeNumber
                            ? SourceValueType.WholeNumber
                            : SourceValueType.DecimalNumber;
                        break;
                    case NormalizePeriodsTransform:
                        ApplyNormalizedPeriodTypes(fieldTypes, mapping);
                        break;
                }
            }

            return fieldTypes;
        }

        private static void RetainFieldTypes(
            Dictionary<string, SourceValueType> fieldTypes,
            IReadOnlyCollection<string> retained)
        {
            foreach (string field in fieldTypes.Keys.ToArray())
            {
                if (!retained.Contains(field, StringComparer.OrdinalIgnoreCase))
                {
                    fieldTypes.Remove(field);
                }
            }
        }

        private static void MergeReplacementType(
            Dictionary<string, SourceValueType> fieldTypes,
            string field,
            ScalarValue replacement)
        {
            if (!fieldTypes.TryGetValue(field, out SourceValueType current)) return;
            SourceValueType replacementType = ToSourceValueType(replacement);
            fieldTypes[field] = MergeSourceTypes(current, replacementType);
        }

        private static SourceValueType ToSourceValueType(ScalarValue value)
        {
            if (value == null || value.Kind == ScalarValueKind.Null) return SourceValueType.Empty;
            switch (value.Kind)
            {
                case ScalarValueKind.Number: return SourceValueType.DecimalNumber;
                case ScalarValueKind.Boolean: return SourceValueType.Boolean;
                case ScalarValueKind.Date: return SourceValueType.Date;
                case ScalarValueKind.DateTime: return SourceValueType.DateTime;
                default: return SourceValueType.Text;
            }
        }

        private static SourceValueType MergeSourceTypes(SourceValueType left, SourceValueType right)
        {
            if (right == SourceValueType.Empty) return left;
            if (left == SourceValueType.Empty) return right;
            if (left == right) return left;
            if (IsNumericType(left) && IsNumericType(right)) return SourceValueType.DecimalNumber;
            if ((left == SourceValueType.Date || left == SourceValueType.DateTime) &&
                (right == SourceValueType.Date || right == SourceValueType.DateTime))
            {
                return SourceValueType.DateTime;
            }

            return SourceValueType.Mixed;
        }

        private static void DemandNumericOperand(
            ArithmeticOperand operand,
            IReadOnlyDictionary<string, SourceValueType> fieldTypes)
        {
            if (operand.Kind != ArithmeticOperandKind.Column) return;
            if (!fieldTypes.TryGetValue(operand.Column ?? string.Empty, out SourceValueType type) || !IsNumericType(type))
            {
                throw new InvalidOperationException(
                    "Arithmetic requires numeric columns after earlier preparation steps.");
            }
        }

        private static void ApplyNormalizedPeriodTypes(
            Dictionary<string, SourceValueType> fieldTypes,
            PeriodMappingSpec? mapping)
        {
            if (mapping == null) return;
            if (mapping.Kind == PeriodMappingKind.LongDateColumn)
            {
                fieldTypes[mapping.DateColumn ?? string.Empty] = SourceValueType.Date;
                return;
            }

            SourceValueType valueType = mapping.Columns
                .Select(column => fieldTypes.TryGetValue(column.SourceColumn, out SourceValueType type)
                    ? type
                    : SourceValueType.Mixed)
                .Aggregate(SourceValueType.Empty, MergeSourceTypes);
            var normalized = new Dictionary<string, SourceValueType>(StringComparer.OrdinalIgnoreCase);
            foreach (string key in mapping.KeyColumns)
            {
                if (fieldTypes.TryGetValue(key, out SourceValueType type)) normalized[key] = type;
            }

            normalized[mapping.PeriodColumnName] = SourceValueType.Date;
            normalized[mapping.ValueColumnName] = valueType;
            if (mapping.Kind == PeriodMappingKind.MetricMonthHeaders)
            {
                normalized[mapping.MetricColumnName] = SourceValueType.Text;
            }

            fieldTypes.Clear();
            foreach (KeyValuePair<string, SourceValueType> item in normalized)
            {
                fieldTypes[item.Key] = item.Value;
            }
        }

        private static bool IsNumericType(SourceValueType type)
        {
            return type == SourceValueType.WholeNumber || type == SourceValueType.DecimalNumber;
        }

        private static SourceValueType ToSourceValueType(ColumnDataType dataType)
        {
            switch (dataType)
            {
                case ColumnDataType.WholeNumber: return SourceValueType.WholeNumber;
                case ColumnDataType.DecimalNumber: return SourceValueType.DecimalNumber;
                case ColumnDataType.Boolean: return SourceValueType.Boolean;
                case ColumnDataType.Date: return SourceValueType.Date;
                case ColumnDataType.DateTime: return SourceValueType.DateTime;
                default: return SourceValueType.Text;
            }
        }

        private static IReadOnlyList<string> ReadStrings(JsonElement root, string property)
        {
            return root.GetProperty(property)
                .EnumerateArray()
                .Select(value => value.GetString() ?? string.Empty)
                .ToList();
        }

        private static Dictionary<string, string> ReadFieldSettings(
            JsonElement root,
            string property,
            string setting)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (JsonElement item in root.GetProperty(property).EnumerateArray())
            {
                string field = item.GetProperty("field").GetString() ?? string.Empty;
                result[field] = item.GetProperty(setting).GetString() ?? string.Empty;
            }

            return result;
        }

        private static Dictionary<string, string> ReadFormatting(JsonElement root)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (JsonElement item in root.GetProperty("formatting").EnumerateArray())
            {
                string field = item.GetProperty("field").GetString() ?? string.Empty;
                string style = item.GetProperty("numberStyle").GetString() ?? "general";
                int decimals = item.GetProperty("decimalPlaces").GetInt32();
                result[field] = ToNumberFormat(style, decimals);
            }

            return result;
        }

        private static string ToNumberFormat(string style, int decimals)
        {
            string decimalPart = decimals == 0 ? string.Empty : "." + new string('0', decimals);
            switch (style)
            {
                case "integer": return "#,##0";
                case "decimal": return "#,##0" + decimalPart;
                case "currency": return "$#,##0" + decimalPart;
                case "percentage": return "0" + decimalPart + "%";
                default: return "General";
            }
        }

        private static string UniqueId(
            string candidate,
            int index,
            ISet<string> existing)
        {
            string value = BoundedId(candidate);
            if (existing.Add(value))
            {
                return value;
            }

            value = BoundedId(candidate + "_" + index.ToString(CultureInfo.InvariantCulture));
            while (!existing.Add(value))
            {
                index++;
                value = BoundedId(candidate + "_" + index.ToString(CultureInfo.InvariantCulture));
            }

            return value;
        }

        internal static string BoundedId(string value)
        {
            var builder = new StringBuilder();
            foreach (char character in value ?? string.Empty)
            {
                if (char.IsLetterOrDigit(character) || character == '_' || character == '-')
                {
                    builder.Append(character);
                }
                else
                {
                    builder.Append('_');
                }
            }

            string result = builder.ToString().Trim('_');
            if (string.IsNullOrWhiteSpace(result) || !char.IsLetter(result[0]))
            {
                result = "item_" + result;
            }

            return result.Length <= 64 ? result : result.Substring(0, 64);
        }

        private sealed class PlacementInput
        {
            public PlacementBucket Bucket { get; set; }

            public string Field { get; set; } = string.Empty;

            public string Aggregate { get; set; } = "sum";

            public string Subtotals { get; set; } = "show";

            public string Sort { get; set; } = "sourceOrder";

            public string? NumberFormat { get; set; }

            public string SubtotalPlacement { get; set; } = "After members";

            public List<string> MemberOrder { get; set; } = new List<string>();

            public List<string> FilterValues { get; set; } = new List<string>();

            public PlacementInput Clone()
            {
                return new PlacementInput
                {
                    Bucket = Bucket,
                    Field = Field,
                    Aggregate = Aggregate,
                    Subtotals = Subtotals,
                    Sort = Sort,
                    NumberFormat = NumberFormat,
                    SubtotalPlacement = SubtotalPlacement,
                    MemberOrder = new List<string>(MemberOrder),
                    FilterValues = new List<string>(FilterValues)
                };
            }

            public static PlacementInput FromUi(FieldPlacementSnapshot snapshot)
            {
                return new PlacementInput
                {
                    Bucket = snapshot.Bucket,
                    Field = snapshot.ColumnName,
                    Aggregate = snapshot.Bucket == PlacementBucket.Values
                        ? snapshot.Setting
                        : "sum",
                    Sort = snapshot.Bucket == PlacementBucket.Rows ||
                        snapshot.Bucket == PlacementBucket.Columns
                        ? snapshot.Setting
                        : "sourceOrder",
                    Subtotals = snapshot.ShowSubtotals ? "show" : "hide",
                    SubtotalPlacement = snapshot.SubtotalPlacement,
                    MemberOrder = new List<string>(snapshot.MemberOrder),
                    NumberFormat = snapshot.Bucket == PlacementBucket.Values
                        ? snapshot.NumberFormat
                        : null,
                    FilterValues = new List<string>(snapshot.SelectedValues)
                };
            }
        }
    }
}
