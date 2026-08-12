using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using ExcelReportBuilder.Core.PivotPlus;
using ExcelReportBuilder.Excel.PivotPlus.Persistence;

namespace ExcelReportBuilder.Excel.PivotPlus.Semantics
{
    internal static class PivotSemanticLayoutCanonical
    {
        private const int MaximumPlacements = 256;
        private const int MaximumDefinitionMappings = 512;
        private const int MaximumUniqueNameCharacters = 2048;

        public static string CreateFilterFingerprint(
            IEnumerable<PivotSemanticFilterFieldSnapshot> filters)
        {
            if (filters == null) throw new ArgumentNullException(nameof(filters));
            var canonical = new StringBuilder("semantic-filters-v1");
            foreach (PivotSemanticFilterFieldSnapshot field in filters
                         .OrderBy(item => item.Position))
            {
                Append(canonical, field.UniqueName);
                Append(canonical, field.Caption);
                Append(canonical, field.Position);
                Append(canonical, field.StateFingerprint);
            }

            return PivotPlusFingerprint.Create(
                "semantic.filters.v1",
                canonical.ToString());
        }

        public static string CreateLayoutFingerprint(
            IEnumerable<PivotSemanticAxisFieldSnapshot> rows,
            IEnumerable<PivotSemanticAxisFieldSnapshot> columns,
            IEnumerable<PivotSemanticValueFieldSnapshot> values,
            PivotValuesAxis valuesAxis,
            int valuesPosition,
            string filterFingerprint)
        {
            if (rows == null) throw new ArgumentNullException(nameof(rows));
            if (columns == null) throw new ArgumentNullException(nameof(columns));
            if (values == null) throw new ArgumentNullException(nameof(values));
            if (filterFingerprint == null)
            {
                throw new ArgumentNullException(nameof(filterFingerprint));
            }

            var canonical = new StringBuilder("semantic-layout-v1");
            AppendAxis(canonical, "rows", rows);
            AppendAxis(canonical, "columns", columns);
            canonical.Append('|').Append("values");
            foreach (PivotSemanticValueFieldSnapshot value in values
                         .OrderBy(item => item.Position))
            {
                Append(canonical, value.UniqueName);
                Append(canonical, value.Caption);
                Append(canonical, value.CaptionFingerprint);
                Append(canonical, value.NumberFormat);
                Append(canonical, value.NumberFormatFingerprint);
                Append(canonical, value.Position);
                Append(canonical, value.CubeFieldType);
            }

            Append(canonical, (int)valuesAxis);
            Append(canonical, valuesPosition);
            Append(canonical, filterFingerprint);
            return PivotPlusFingerprint.Create(
                "semantic.layout.v1",
                canonical.ToString());
        }

        public static void ValidatePlanAndMappings(
            PivotSemanticLayoutPlan plan,
            IReadOnlyDictionary<string, string> namedSets,
            IReadOnlyDictionary<string, string> measures,
            PivotSemanticLayoutSnapshot before)
        {
            if (plan == null) throw new ArgumentNullException(nameof(plan));
            if (namedSets == null) throw new ArgumentNullException(nameof(namedSets));
            if (measures == null) throw new ArgumentNullException(nameof(measures));
            if (before == null) throw new ArgumentNullException(nameof(before));
            if ((long)plan.Rows.Count + plan.Columns.Count + plan.Values.Count >
                    MaximumPlacements ||
                (long)before.Rows.Count + before.Columns.Count + before.Values.Count >
                    MaximumPlacements ||
                namedSets.Count > MaximumDefinitionMappings ||
                measures.Count > MaximumDefinitionMappings)
            {
                throw new NotSupportedException(
                    "The semantic layout exceeds its bounded field limit.");
            }

            DemandExactPositions(plan.Rows.Select(item => item.Position), "Rows");
            DemandExactPositions(plan.Columns.Select(item => item.Position), "Columns");
            DemandExactPositions(plan.Values.Select(item => item.Position), "Values");
            DemandValuesAxis(plan);
            ValidateMappings(namedSets, "named-set");
            ValidateMappings(measures, "measure");

            var existingAxisKeys = new HashSet<string>(StringComparer.Ordinal);
            var generatedSetIds = new HashSet<string>(StringComparer.Ordinal);
            ValidateAxis(
                plan.Rows,
                PivotFieldArea.Row,
                namedSets,
                before,
                existingAxisKeys,
                generatedSetIds);
            ValidateAxis(
                plan.Columns,
                PivotFieldArea.Column,
                namedSets,
                before,
                existingAxisKeys,
                generatedSetIds);

            var existingValueKeys = new HashSet<string>(StringComparer.Ordinal);
            var generatedMeasureIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (PivotSemanticValuePlacement value in plan.Values)
            {
                if (value == null)
                {
                    throw new ArgumentException(
                        "A semantic Values placement cannot be null.",
                        nameof(plan));
                }

                if (value.IsGeneratedMeasure)
                {
                    string definitionId = DemandDefinitionId(
                        value.DefinitionId,
                        "measure definition identifier");
                    if (!measures.ContainsKey(definitionId) ||
                        !generatedMeasureIds.Add(definitionId))
                    {
                        throw new ArgumentException(
                            "Each generated measure placement must resolve once through the trusted map.",
                            nameof(plan));
                    }

                    continue;
                }

                if (!string.IsNullOrEmpty(value.DefinitionId))
                {
                    throw new ArgumentException(
                        "A Values placement cannot be both existing and generated.",
                        nameof(plan));
                }

                PivotExistingSemanticValueIdentity identity = value.ExistingDataField!;
                DemandHostIdentity(
                    identity.UniqueName,
                    identity.CurrentCaptionFingerprint,
                    identity.CurrentPosition,
                    "existing Values field");
                PivotPlusMetadataValidator.ValidateFingerprint(
                    identity.CurrentNumberFormatFingerprint,
                    "existing Values number-format fingerprint");
                string key = ValueKey(identity);
                if (!existingValueKeys.Add(key) ||
                    !before.Values.Any(field => Matches(field, identity)))
                {
                    throw new InvalidOperationException(
                        "An existing Values placement does not match exactly one preview field.");
                }
            }

            var filterNames = new HashSet<string>(
                before.Filters.Select(item => item.UniqueName),
                StringComparer.OrdinalIgnoreCase);
            IEnumerable<string> desiredNames = plan.Rows.Concat(plan.Columns)
                .Where(item => item.IsGeneratedNamedSet)
                .Select(item => namedSets[item.DefinitionId!])
                .Concat(plan.Rows.Concat(plan.Columns)
                    .Where(item => !item.IsGeneratedNamedSet)
                    .Select(item => item.ExistingField!.UniqueName))
                .Concat(plan.Values
                    .Where(item => item.IsGeneratedMeasure)
                    .Select(item => measures[item.DefinitionId!]))
                .Concat(plan.Values
                    .Where(item => !item.IsGeneratedMeasure)
                    .Select(item => item.ExistingDataField!.UniqueName));
            if (desiredNames.Any(filterNames.Contains))
            {
                throw new NotSupportedException(
                    "A field cannot be moved while it remains in the preserved Filters area.");
            }

            if (plan.Rows.Concat(plan.Columns)
                    .Select(item => item.IsGeneratedNamedSet
                        ? namedSets[item.DefinitionId!]
                        : item.ExistingField!.UniqueName)
                    .GroupBy(name => name, StringComparer.OrdinalIgnoreCase)
                    .Any(group => group.Count() != 1))
            {
                throw new NotSupportedException(
                    "A hierarchy or named set can appear only once across Rows and Columns.");
            }

            var generatedValueNames = new HashSet<string>(
                plan.Values
                    .Where(item => item.IsGeneratedMeasure)
                    .Select(item => measures[item.DefinitionId!]),
                StringComparer.OrdinalIgnoreCase);
            if (plan.Values
                .Where(item => !item.IsGeneratedMeasure)
                .Any(item => generatedValueNames.Contains(
                    item.ExistingDataField!.UniqueName)))
            {
                throw new NotSupportedException(
                    "A generated measure and retained existing Values occurrence cannot share one host identity in the same plan.");
            }
        }

        public static bool Matches(
            PivotSemanticAxisFieldSnapshot field,
            PivotExistingAxisFieldIdentity identity)
        {
            return string.Equals(
                       field.UniqueName,
                       identity.UniqueName,
                       StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(
                       field.CaptionFingerprint,
                       identity.CurrentCaptionFingerprint,
                       StringComparison.Ordinal) &&
                   field.Area == identity.CurrentArea &&
                   field.Position == identity.CurrentPosition;
        }

        public static bool Matches(
            PivotSemanticValueFieldSnapshot field,
            PivotExistingSemanticValueIdentity identity)
        {
            return string.Equals(
                       field.UniqueName,
                       identity.UniqueName,
                       StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(
                       field.CaptionFingerprint,
                       identity.CurrentCaptionFingerprint,
                       StringComparison.Ordinal) &&
                   string.Equals(
                       field.NumberFormatFingerprint,
                       identity.CurrentNumberFormatFingerprint,
                       StringComparison.Ordinal) &&
                   field.Position == identity.CurrentPosition;
        }

        public static string AxisKey(PivotExistingAxisFieldIdentity identity)
        {
            return identity.UniqueName + "\u001f" +
                   identity.CurrentCaptionFingerprint + "\u001f" +
                   ((int)identity.CurrentArea).ToString(CultureInfo.InvariantCulture) + "\u001f" +
                   identity.CurrentPosition.ToString(CultureInfo.InvariantCulture);
        }

        public static string ValueKey(PivotExistingSemanticValueIdentity identity)
        {
            return identity.UniqueName + "\u001f" +
                   identity.CurrentCaptionFingerprint + "\u001f" +
                   identity.CurrentNumberFormatFingerprint + "\u001f" +
                   identity.CurrentPosition.ToString(CultureInfo.InvariantCulture);
        }

        private static void ValidateAxis(
            IEnumerable<PivotSemanticAxisPlacement> placements,
            PivotFieldArea desiredArea,
            IReadOnlyDictionary<string, string> namedSets,
            PivotSemanticLayoutSnapshot before,
            ISet<string> existingKeys,
            ISet<string> generatedIds)
        {
            foreach (PivotSemanticAxisPlacement placement in placements)
            {
                if (placement == null)
                {
                    throw new ArgumentException(
                        "A semantic axis placement cannot be null.",
                        nameof(placements));
                }

                if (placement.IsGeneratedNamedSet)
                {
                    string definitionId = DemandDefinitionId(
                        placement.DefinitionId,
                        "named-set definition identifier");
                    if (!namedSets.ContainsKey(definitionId) ||
                        !generatedIds.Add(definitionId))
                    {
                        throw new ArgumentException(
                            "Each generated named-set placement must resolve once through the trusted map.",
                            nameof(placements));
                    }

                    continue;
                }

                if (!string.IsNullOrEmpty(placement.DefinitionId))
                {
                    throw new ArgumentException(
                        "An axis placement cannot be both existing and generated.",
                        nameof(placements));
                }

                PivotExistingAxisFieldIdentity identity = placement.ExistingField!;
                DemandHostIdentity(
                    identity.UniqueName,
                    identity.CurrentCaptionFingerprint,
                    identity.CurrentPosition,
                    "existing axis field");
                if (identity.CurrentArea != PivotFieldArea.Row &&
                    identity.CurrentArea != PivotFieldArea.Column)
                {
                    throw new ArgumentException(
                        "An existing axis identity must come from Rows or Columns.",
                        nameof(placements));
                }

                string key = AxisKey(identity);
                IEnumerable<PivotSemanticAxisFieldSnapshot> preview =
                    before.Rows.Concat(before.Columns);
                if (!existingKeys.Add(key) ||
                    preview.Count(field => Matches(field, identity)) != 1)
                {
                    throw new InvalidOperationException(
                        "An existing axis placement does not match exactly one preview field.");
                }

                _ = desiredArea;
            }
        }

        private static void ValidateMappings(
            IReadOnlyDictionary<string, string> mappings,
            string label)
        {
            var hostNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, string> mapping in mappings)
            {
                DemandDefinitionId(mapping.Key, label + " definition identifier");
                DemandUniqueName(mapping.Value, label + " host unique name");
                if (!hostNames.Add(mapping.Value))
                {
                    throw new ArgumentException(
                        "Trusted " + label + " mappings contain duplicate host identities.",
                        nameof(mappings));
                }
            }
        }

        private static void DemandHostIdentity(
            string uniqueName,
            string captionFingerprint,
            int position,
            string label)
        {
            DemandUniqueName(uniqueName, label + " unique name");
            PivotPlusMetadataValidator.ValidateFingerprint(
                captionFingerprint,
                label + " caption fingerprint");
            if (position <= 0)
            {
                throw new ArgumentException(
                    "A one-based " + label + " position is required.");
            }
        }

        private static string DemandDefinitionId(string? value, string label)
        {
            string result = value ?? string.Empty;
            PivotPlusMetadataValidator.ValidateId(result, label);
            return result;
        }

        private static void DemandUniqueName(string value, string label)
        {
            if (string.IsNullOrWhiteSpace(value) ||
                value.Length > MaximumUniqueNameCharacters ||
                value.Any(char.IsControl))
            {
                throw new ArgumentException("A bounded " + label + " is required.");
            }
        }

        private static void DemandExactPositions(
            IEnumerable<int> positions,
            string label)
        {
            int[] ordered = positions.OrderBy(position => position).ToArray();
            if (ordered.Where((position, index) => position != index + 1).Any())
            {
                throw new ArgumentException(
                    label + " must be a complete, contiguous one-based sequence.");
            }
        }

        private static void DemandValuesAxis(PivotSemanticLayoutPlan plan)
        {
            if (plan.Values.Count <= 1)
            {
                if (plan.ValuesAxis != PivotValuesAxis.Automatic ||
                    plan.ValuesPosition != 1)
                {
                    throw new ArgumentException(
                        "A zero- or single-value semantic layout must use the automatic Values axis at position one.");
                }

                return;
            }

            if (plan.ValuesAxis != PivotValuesAxis.Rows &&
                plan.ValuesAxis != PivotValuesAxis.Columns)
            {
                throw new ArgumentException(
                    "A multi-value semantic layout requires a row or column Values axis.");
            }

            int regularCount = plan.ValuesAxis == PivotValuesAxis.Rows
                ? plan.Rows.Count
                : plan.Columns.Count;
            if (plan.ValuesPosition <= 0 ||
                plan.ValuesPosition > regularCount + 1)
            {
                throw new ArgumentException(
                    "The Values pseudo-field position is outside the chosen axis.");
            }
        }

        private static void AppendAxis(
            StringBuilder canonical,
            string label,
            IEnumerable<PivotSemanticAxisFieldSnapshot> fields)
        {
            canonical.Append('|').Append(label);
            foreach (PivotSemanticAxisFieldSnapshot field in fields
                         .OrderBy(item => item.Position))
            {
                Append(canonical, field.UniqueName);
                Append(canonical, field.Caption);
                Append(canonical, field.CaptionFingerprint);
                Append(canonical, (int)field.Area);
                Append(canonical, field.Position);
                Append(canonical, field.CubeFieldType);
            }
        }

        private static void Append(StringBuilder target, string value)
        {
            string actual = value ?? string.Empty;
            target.Append('|')
                .Append(actual.Length.ToString(CultureInfo.InvariantCulture))
                .Append(':')
                .Append(actual);
        }

        private static void Append(StringBuilder target, int value)
        {
            target.Append('|').Append(value.ToString(CultureInfo.InvariantCulture));
        }
    }
}
