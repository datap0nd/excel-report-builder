using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using ExcelReportBuilder.Core.PivotPlus.Calculations;
using ExcelReportBuilder.Excel.PivotPlus.Persistence;

namespace ExcelReportBuilder.Excel.PivotPlus.Measures
{
    internal static class PivotModelMeasureCanonical
    {
        public static string CreateDescriptionMarker(
            string setupId,
            string definitionId,
            string definitionFingerprint)
        {
            PivotPlusMetadataValidator.ValidateId(setupId, "setup identifier");
            PivotPlusMetadataValidator.ValidateId(definitionId, "measure definition identifier");
            PivotPlusMetadataValidator.ValidateFingerprint(
                definitionFingerprint,
                "measure definition fingerprint");
            var semantic = new StringBuilder("measure-semantic-v1");
            Append(semantic, setupId);
            Append(semantic, definitionId);
            Append(semantic, definitionFingerprint);
            // ModelMeasure.Description is a host marker, not a second copy of
            // the semantic plan. Keep it short and hash-only so maximum-length
            // valid setup/definition identifiers remain safe for Excel.
            return "PivotTable+|" + PivotPlusFingerprint.Create(
                "measure.semantic.v1",
                semantic.ToString());
        }

        public static string CreatePlanFingerprint(
            IEnumerable<DesiredModelMeasure> definitions,
            PivotMeasurePlacementPlan placement)
        {
            if (definitions == null) throw new ArgumentNullException(nameof(definitions));
            if (placement == null) throw new ArgumentNullException(nameof(placement));

            var canonical = new StringBuilder("measure-plan-v1");
            foreach (DesiredModelMeasure definition in definitions
                         .OrderBy(item => item.DefinitionId, StringComparer.Ordinal))
            {
                Append(canonical, definition.DefinitionId);
                Append(canonical, definition.DefinitionFingerprint);
            }

            canonical.Append('|').Append((int)placement.ValuesAxis);
            canonical.Append('|').Append(
                placement.ValuesPosition.ToString(CultureInfo.InvariantCulture));
            foreach (PivotMeasureValuePlacement value in placement.Values
                         .OrderBy(item => item.Position))
            {
                canonical.Append('|').Append(value.Position.ToString(CultureInfo.InvariantCulture));
                if (value.IsGeneratedMeasure)
                {
                    canonical.Append("|generated");
                    Append(canonical, value.DefinitionId ?? string.Empty);
                }
                else
                {
                    canonical.Append("|existing");
                    Append(canonical, value.ExistingDataField!.UniqueName);
                    Append(canonical, value.ExistingDataField.CurrentCaptionFingerprint);
                    Append(canonical, value.ExistingDataField.CurrentNumberFormatFingerprint);
                    canonical.Append('|').Append(
                        value.ExistingDataField.CurrentPosition.ToString(
                            CultureInfo.InvariantCulture));
                }
            }

            return PivotPlusFingerprint.Create("measure.plan.v1", canonical.ToString());
        }

        public static string CreateArtifactPlanFingerprint(
            IEnumerable<DesiredModelMeasure> definitions,
            IEnumerable<PivotPlusSemanticArtifactTransition> transitions)
        {
            if (definitions == null) throw new ArgumentNullException(nameof(definitions));
            if (transitions == null) throw new ArgumentNullException(nameof(transitions));

            var canonical = new StringBuilder("measure-artifact-plan-v1");
            foreach (DesiredModelMeasure definition in definitions
                         .OrderBy(item => item.DefinitionId, StringComparer.Ordinal))
            {
                Append(canonical, definition.DefinitionId);
                Append(canonical, definition.Name);
                Append(canonical, definition.DefinitionFingerprint);
            }

            foreach (PivotPlusSemanticArtifactTransition transition in transitions
                         .OrderBy(item => item.ArtifactId, StringComparer.Ordinal))
            {
                canonical.Append('|').Append((int)transition.Kind);
                canonical.Append('|').Append((int)transition.Operation);
                Append(canonical, transition.ArtifactId);
                Append(canonical, transition.BeforeLiveFingerprint ?? string.Empty);
                Append(canonical, transition.PlannedDefinitionFingerprint);
            }

            return PivotPlusFingerprint.Create(
                "measure.artifact-plan.v1",
                canonical.ToString());
        }

        public static string CreateLiveFingerprint(
            string name,
            string associatedTableName,
            string associatedTableLineageFingerprint,
            string formula,
            string description,
            ModelMeasureFormatSnapshot format)
        {
            if (format == null) throw new ArgumentNullException(nameof(format));
            var canonical = new StringBuilder("measure-host-v1");
            Append(canonical, name);
            Append(canonical, associatedTableName);
            Append(canonical, associatedTableLineageFingerprint);
            Append(canonical, formula);
            Append(canonical, description);
            Append(canonical, FormatKey(format));
            return PivotPlusFingerprint.Create("measure.host.v1", canonical.ToString());
        }

        public static string CreatePivotFingerprint(ModelPivotUsageSnapshot pivot)
        {
            if (pivot == null) throw new ArgumentNullException(nameof(pivot));
            var canonical = new StringBuilder("measure-pivot-v1");
            canonical.Append((int)pivot.ValuesAxis).Append('|')
                .Append(pivot.ValuesPosition.ToString(CultureInfo.InvariantCulture));
            foreach (ModelDataFieldSnapshot field in pivot.DataFields.OrderBy(item => item.Position))
            {
                canonical.Append('|').Append(field.Position.ToString(CultureInfo.InvariantCulture));
                Append(canonical, field.UniqueName);
                Append(canonical, field.CaptionFingerprint);
                Append(canonical, field.NumberFormat);
            }

            return PivotPlusFingerprint.Create("measure.pivot.v1", canonical.ToString());
        }

        public static string CreateExpectedPivotFingerprint(
            PivotMeasurePlacementPlan placement,
            IReadOnlyDictionary<string, DesiredModelMeasure> definitionsById,
            ModelPivotUsageSnapshot before)
        {
            if (placement == null) throw new ArgumentNullException(nameof(placement));
            if (definitionsById == null) throw new ArgumentNullException(nameof(definitionsById));
            if (before == null) throw new ArgumentNullException(nameof(before));

            var existing = before.DataFields.ToDictionary(
                field => ExistingKey(
                    field.UniqueName,
                    field.CaptionFingerprint,
                    PivotMeasurePlacementFingerprint.CreateNumberFormatFingerprint(
                        field.NumberFormat),
                    field.Position),
                StringComparer.OrdinalIgnoreCase);
            var canonical = new StringBuilder("measure-expected-pivot-v1");
            canonical.Append((int)placement.ValuesAxis).Append('|')
                .Append(placement.ValuesPosition.ToString(CultureInfo.InvariantCulture));
            foreach (PivotMeasureValuePlacement value in placement.Values.OrderBy(item => item.Position))
            {
                canonical.Append('|').Append(value.Position.ToString(CultureInfo.InvariantCulture));
                if (value.IsGeneratedMeasure)
                {
                    DesiredModelMeasure definition = definitionsById[value.DefinitionId!];
                    Append(canonical, "generated");
                    Append(canonical, definition.Name);
                    Append(canonical, definition.DefinitionFingerprint);
                }
                else
                {
                    PivotExistingDataFieldIdentity identity = value.ExistingDataField!;
                    ModelDataFieldSnapshot field = existing[
                        ExistingKey(
                            identity.UniqueName,
                            identity.CurrentCaptionFingerprint,
                            identity.CurrentNumberFormatFingerprint,
                            identity.CurrentPosition)];
                    Append(canonical, "existing");
                    Append(canonical, field.UniqueName);
                    Append(canonical, field.CaptionFingerprint);
                    Append(canonical, field.NumberFormat);
                }
            }

            return PivotPlusFingerprint.Create("measure.expected-pivot.v1", canonical.ToString());
        }

        public static bool MatchesObservedExpectedPivotFingerprint(
            PivotMeasurePlacementPlan placement,
            IReadOnlyDictionary<string, DesiredModelMeasure> definitionsById,
            ModelPivotUsageSnapshot observed,
            string expectedFingerprint)
        {
            return TryCreateObservedExpectedPivotFingerprint(
                       placement,
                       definitionsById,
                       observed,
                       out string actualFingerprint) &&
                   string.Equals(
                       actualFingerprint,
                       expectedFingerprint,
                       StringComparison.Ordinal);
        }

        public static bool TryCreateObservedExpectedPivotFingerprint(
            PivotMeasurePlacementPlan placement,
            IReadOnlyDictionary<string, DesiredModelMeasure> definitionsById,
            ModelPivotUsageSnapshot observed,
            out string observedFingerprint)
        {
            if (placement == null) throw new ArgumentNullException(nameof(placement));
            if (definitionsById == null) throw new ArgumentNullException(nameof(definitionsById));
            if (observed == null) throw new ArgumentNullException(nameof(observed));
            observedFingerprint = string.Empty;
            if (placement.ValuesAxis != observed.ValuesAxis ||
                placement.ValuesPosition != observed.ValuesPosition ||
                placement.Values.Count != observed.DataFields.Count)
            {
                return false;
            }

            var observedByPosition = observed.DataFields.ToDictionary(field => field.Position);
            var canonical = new StringBuilder("measure-expected-pivot-v1");
            canonical.Append((int)placement.ValuesAxis).Append('|')
                .Append(placement.ValuesPosition.ToString(CultureInfo.InvariantCulture));
            foreach (PivotMeasureValuePlacement value in placement.Values.OrderBy(item => item.Position))
            {
                if (!observedByPosition.TryGetValue(
                        value.Position,
                        out ModelDataFieldSnapshot? field))
                {
                    return false;
                }

                canonical.Append('|').Append(value.Position.ToString(CultureInfo.InvariantCulture));
                if (value.IsGeneratedMeasure)
                {
                    DesiredModelMeasure definition = definitionsById[value.DefinitionId!];
                    if (!field.IsModelMeasure ||
                        !string.Equals(
                            field.ModelMeasureName,
                            definition.Name,
                            StringComparison.OrdinalIgnoreCase))
                    {
                        return false;
                    }

                    Append(canonical, "generated");
                    Append(canonical, definition.Name);
                    Append(canonical, definition.DefinitionFingerprint);
                    continue;
                }

                PivotExistingDataFieldIdentity identity = value.ExistingDataField!;
                if (!string.Equals(
                        field.UniqueName,
                        identity.UniqueName,
                        StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(
                        field.CaptionFingerprint,
                        identity.CurrentCaptionFingerprint,
                        StringComparison.Ordinal) ||
                    !string.Equals(
                        PivotMeasurePlacementFingerprint.CreateNumberFormatFingerprint(
                            field.NumberFormat),
                        identity.CurrentNumberFormatFingerprint,
                        StringComparison.Ordinal))
                {
                    return false;
                }

                Append(canonical, "existing");
                Append(canonical, field.UniqueName);
                Append(canonical, field.CaptionFingerprint);
                Append(canonical, field.NumberFormat);
            }

            observedFingerprint = PivotPlusFingerprint.Create(
                "measure.expected-pivot.v1",
                canonical.ToString());
            return true;
        }

        public static string CreateDeleteDefinitionFingerprint(
            string artifactId,
            string beforeLiveFingerprint)
        {
            var canonical = new StringBuilder("measure-delete-v1");
            Append(canonical, artifactId);
            Append(canonical, beforeLiveFingerprint);
            return PivotPlusFingerprint.Create("measure.delete.v1", canonical.ToString());
        }

        private static string ExistingKey(
            string uniqueName,
            string captionFingerprint,
            string numberFormatFingerprint,
            int position)
        {
            return uniqueName + "\u001f" + captionFingerprint + "\u001f" +
                   numberFormatFingerprint + "\u001f" +
                   position.ToString(CultureInfo.InvariantCulture);
        }

        private static string FormatKey(ModelMeasureFormatSnapshot format)
        {
            var result = new StringBuilder();
            result.Append((int)format.Kind).Append('|')
                .Append(format.DecimalPlaces.HasValue
                    ? format.DecimalPlaces.Value.ToString(CultureInfo.InvariantCulture)
                    : string.Empty)
                .Append('|')
                .Append(format.UseThousandsSeparator.HasValue
                    ? (format.UseThousandsSeparator.Value ? "true" : "false")
                    : string.Empty);
            Append(result, format.CurrencySymbol ?? string.Empty);
            Append(result, format.DateFormatString ?? string.Empty);
            return result.ToString();
        }

        private static void Append(StringBuilder target, string value)
        {
            string actual = value ?? string.Empty;
            target.Append('|')
                .Append(actual.Length.ToString(CultureInfo.InvariantCulture))
                .Append(':')
                .Append(actual);
        }
    }

    /// <summary>
    /// Conservative reference scanner for live user-authored DAX. Every
    /// bracketed identifier outside strings and comments is considered a
    /// possible measure dependency. This can reject a same-named column, but it
    /// cannot silently overlook a simple measure reference.
    /// </summary>
    internal static class DaxMeasureReferenceScanner
    {
        public static IReadOnlyCollection<string> ReadPossibleReferences(string formula)
        {
            if (formula == null) throw new ArgumentNullException(nameof(formula));
            var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int index = 0;
            while (index < formula.Length)
            {
                char current = formula[index];
                if (current == '"')
                {
                    index = SkipQuoted(formula, index + 1, '"');
                    continue;
                }

                if (current == '\'' )
                {
                    index = SkipQuoted(formula, index + 1, '\'');
                    continue;
                }

                if (current == '/' && index + 1 < formula.Length && formula[index + 1] == '/')
                {
                    index = SkipLine(formula, index + 2);
                    continue;
                }

                if (current == '-' && index + 1 < formula.Length && formula[index + 1] == '-')
                {
                    index = SkipLine(formula, index + 2);
                    continue;
                }

                if (current == '/' && index + 1 < formula.Length && formula[index + 1] == '*')
                {
                    int end = formula.IndexOf("*/", index + 2, StringComparison.Ordinal);
                    index = end < 0 ? formula.Length : end + 2;
                    continue;
                }

                if (current != '[')
                {
                    index++;
                    continue;
                }

                var name = new StringBuilder();
                bool closed = false;
                index++;
                while (index < formula.Length)
                {
                    if (formula[index] != ']')
                    {
                        name.Append(formula[index++]);
                        continue;
                    }

                    if (index + 1 < formula.Length && formula[index + 1] == ']')
                    {
                        name.Append(']');
                        index += 2;
                        continue;
                    }

                    index++;
                    closed = true;
                    break;
                }

                if (closed && name.Length > 0)
                {
                    result.Add(name.ToString());
                }
            }

            return result;
        }

        private static int SkipQuoted(string value, int index, char quote)
        {
            while (index < value.Length)
            {
                if (value[index] != quote)
                {
                    index++;
                    continue;
                }

                if (index + 1 < value.Length && value[index + 1] == quote)
                {
                    index += 2;
                    continue;
                }

                return index + 1;
            }

            return value.Length;
        }

        private static int SkipLine(string value, int index)
        {
            while (index < value.Length && value[index] != '\r' && value[index] != '\n')
            {
                index++;
            }

            return index;
        }
    }
}
