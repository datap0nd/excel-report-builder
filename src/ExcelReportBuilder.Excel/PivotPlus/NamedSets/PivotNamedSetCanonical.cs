using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using ExcelReportBuilder.Excel.PivotPlus.Persistence;

namespace ExcelReportBuilder.Excel.PivotPlus.NamedSets
{
    internal static class PivotNamedSetCanonical
    {
        public static string CreateDisplayFolderMarker(
            string setupId,
            string definitionId,
            string definitionFingerprint)
        {
            PivotPlusMetadataValidator.ValidateId(setupId, "setup identifier");
            PivotPlusMetadataValidator.ValidateId(
                definitionId,
                "named-set definition identifier");
            PivotPlusMetadataValidator.ValidateFingerprint(
                definitionFingerprint,
                "named-set definition fingerprint");
            var canonical = new StringBuilder("namedset-semantic-v1");
            Append(canonical, setupId);
            Append(canonical, definitionId);
            Append(canonical, definitionFingerprint);
            return "PivotTable+|set|" + PivotPlusFingerprint.Create(
                "namedset.semantic.v1",
                canonical.ToString());
        }

        public static string CreateStableCatalogId(string kind, string providerUniqueName)
        {
            if (string.IsNullOrWhiteSpace(kind) ||
                kind.Any(character => !char.IsLetterOrDigit(character) && character != '_'))
            {
                throw new ArgumentException("A safe catalog ID kind is required.", nameof(kind));
            }

            if (string.IsNullOrWhiteSpace(providerUniqueName))
            {
                throw new ArgumentException(
                    "A provider unique name is required.",
                    nameof(providerUniqueName));
            }

            string fingerprint = PivotPlusFingerprint.Create(
                "namedset.catalog-id.v1",
                providerUniqueName);
            int separator = fingerprint.LastIndexOf(':');
            return kind + "_" + fingerprint.Substring(separator + 1);
        }

        public static string CreateModelLineageFingerprint(
            IEnumerable<string> modelLineageTokens)
        {
            if (modelLineageTokens == null)
            {
                throw new ArgumentNullException(nameof(modelLineageTokens));
            }

            var canonical = new StringBuilder("namedset-model-lineage-v1");
            foreach (string token in modelLineageTokens.OrderBy(
                         value => value,
                         StringComparer.Ordinal))
            {
                Append(canonical, token);
            }

            return PivotPlusFingerprint.Create(
                "namedset.model-lineage.v1",
                canonical.ToString());
        }

        public static string CreateSourceFingerprint(
            string modelLineageFingerprint,
            IEnumerable<string> catalogTokens)
        {
            if (modelLineageFingerprint == null)
            {
                throw new ArgumentNullException(nameof(modelLineageFingerprint));
            }

            if (catalogTokens == null) throw new ArgumentNullException(nameof(catalogTokens));
            var canonical = new StringBuilder("namedset-source-v1");
            Append(canonical, modelLineageFingerprint);
            foreach (string token in catalogTokens.OrderBy(
                         value => value,
                         StringComparer.Ordinal))
            {
                Append(canonical, token);
            }

            return PivotPlusFingerprint.Create("pivot.source.v1", canonical.ToString());
        }

        public static string CreateCaptionFingerprint(string caption)
        {
            return PivotPlusFingerprint.Create(
                "namedset.caption.v1",
                caption ?? string.Empty);
        }

        public static string CreateLiveFingerprint(
            string sourceFingerprint,
            string modelLineageFingerprint,
            string name,
            PivotNamedSetPairState pairState,
            string formulaFingerprint,
            string displayFolder,
            string sourceName,
            string caption,
            int? calculatedMemberType,
            int? cubeFieldType,
            bool? dynamic,
            bool? calculatedMemberFlattenHierarchies,
            bool? cubeFieldFlattenHierarchies,
            bool? calculatedMemberHierarchizeDistinct,
            bool? cubeFieldHierarchizeDistinct,
            bool? showInFieldList,
            int? orientation,
            bool? isValid)
        {
            var canonical = new StringBuilder("namedset-host-v1");
            Append(canonical, sourceFingerprint);
            Append(canonical, modelLineageFingerprint);
            Append(canonical, name);
            canonical.Append('|').Append((int)pairState);
            Append(canonical, formulaFingerprint);
            Append(canonical, displayFolder);
            Append(canonical, sourceName);
            Append(canonical, caption);
            AppendNullable(canonical, calculatedMemberType);
            AppendNullable(canonical, cubeFieldType);
            AppendNullable(canonical, dynamic);
            AppendNullable(canonical, calculatedMemberFlattenHierarchies);
            AppendNullable(canonical, cubeFieldFlattenHierarchies);
            AppendNullable(canonical, calculatedMemberHierarchizeDistinct);
            AppendNullable(canonical, cubeFieldHierarchizeDistinct);
            AppendNullable(canonical, showInFieldList);
            AppendNullable(canonical, orientation);
            AppendNullable(canonical, isValid);
            return PivotPlusFingerprint.Create("namedset.host.v1", canonical.ToString());
        }

        public static string CreatePivotFingerprint(
            IEnumerable<LivePivotNamedSetSnapshot> artifacts,
            IEnumerable<PivotCalculatedMemberReferenceSnapshot> calculatedMembers)
        {
            if (artifacts == null) throw new ArgumentNullException(nameof(artifacts));
            if (calculatedMembers == null)
            {
                throw new ArgumentNullException(nameof(calculatedMembers));
            }

            var canonical = new StringBuilder("namedset-pivot-v1");
            foreach (LivePivotNamedSetSnapshot artifact in artifacts
                         .OrderBy(item => item.Name, StringComparer.Ordinal)
                         .ThenBy(item => item.PairState))
            {
                Append(canonical, artifact.Name);
                Append(canonical, artifact.LiveFingerprint);
            }

            foreach (PivotCalculatedMemberReferenceSnapshot member in calculatedMembers
                         .OrderBy(item => item.Name, StringComparer.Ordinal)
                         .ThenBy(item => item.Type))
            {
                Append(canonical, member.WorksheetName);
                Append(canonical, member.PivotTableName);
                Append(canonical, member.Name);
                canonical.Append('|').Append(member.Type.ToString(CultureInfo.InvariantCulture));
                Append(canonical, member.RawFormula);
                canonical.Append('|').Append(member.FormulaScanComplete ? "complete" : "incomplete");
            }

            return PivotPlusFingerprint.Create("namedset.pivot.v1", canonical.ToString());
        }

        private static void Append(StringBuilder target, string value)
        {
            string actual = value ?? string.Empty;
            target.Append('|')
                .Append(actual.Length.ToString(CultureInfo.InvariantCulture))
                .Append(':')
                .Append(actual);
        }

        private static void AppendNullable(StringBuilder target, int? value)
        {
            Append(
                target,
                value.HasValue
                    ? value.Value.ToString(CultureInfo.InvariantCulture)
                    : "missing");
        }

        private static void AppendNullable(StringBuilder target, bool? value)
        {
            Append(
                target,
                value.HasValue ? (value.Value ? "true" : "false") : "missing");
        }
    }

    internal sealed class MdxReferenceScanResult
    {
        public MdxReferenceScanResult(
            IEnumerable<string>? bracketedIdentifiers,
            IEnumerable<string>? quotedLiterals,
            bool hasDynamicNameResolution,
            bool isComplete)
        {
            BracketedIdentifiers = bracketedIdentifiers == null
                ? Array.Empty<string>()
                : bracketedIdentifiers.ToArray();
            QuotedLiterals = quotedLiterals == null
                ? Array.Empty<string>()
                : quotedLiterals.ToArray();
            HasDynamicNameResolution = hasDynamicNameResolution;
            IsComplete = isComplete;
        }

        public IReadOnlyList<string> BracketedIdentifiers { get; }

        public IReadOnlyList<string> QuotedLiterals { get; }

        public bool HasDynamicNameResolution { get; }

        public bool IsComplete { get; }
    }

    /// <summary>
    /// Conservative scanner used only to block destructive host operations.
    /// It never compiles or executes formula text.
    /// </summary>
    internal static class MdxNamedSetReferenceScanner
    {
        private static readonly HashSet<string> DynamicNameResolutionFunctions =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "LookupCube",
                "NameToSet",
                "StrToMember",
                "StrToSet",
                "StrToTuple",
                "StrToValue"
            };

        public static MdxReferenceScanResult Scan(string formula)
        {
            if (formula == null) throw new ArgumentNullException(nameof(formula));
            var identifiers = new List<string>();
            var quotedLiterals = new List<string>();
            var currentIdentifier = new StringBuilder();
            var hasDynamicNameResolution = false;
            int index = 0;
            while (index < formula.Length)
            {
                char current = formula[index];
                if (current == '\'' || current == '"')
                {
                    if (!TryReadQuoted(
                            formula,
                            ref index,
                            current,
                            out string quotedLiteral))
                    {
                        return new MdxReferenceScanResult(
                            identifiers,
                            quotedLiterals,
                            hasDynamicNameResolution,
                            false);
                    }

                    quotedLiterals.Add(quotedLiteral);
                    continue;
                }

                if ((current == '-' && Peek(formula, index, '-')) ||
                    (current == '/' && Peek(formula, index, '/')))
                {
                    index += 2;
                    while (index < formula.Length &&
                           formula[index] != '\r' &&
                           formula[index] != '\n')
                    {
                        index++;
                    }

                    continue;
                }

                if (current == '/' && Peek(formula, index, '*'))
                {
                    int end = formula.IndexOf("*/", index + 2, StringComparison.Ordinal);
                    if (end < 0)
                    {
                        return new MdxReferenceScanResult(
                            identifiers,
                            quotedLiterals,
                            hasDynamicNameResolution,
                            false);
                    }

                    index = end + 2;
                    continue;
                }

                if (char.IsLetter(current) || current == '_')
                {
                    int wordStart = index;
                    index++;
                    while (index < formula.Length &&
                           (char.IsLetterOrDigit(formula[index]) || formula[index] == '_'))
                    {
                        index++;
                    }

                    string word = formula.Substring(wordStart, index - wordStart);
                    int next = index;
                    while (next < formula.Length && char.IsWhiteSpace(formula[next])) next++;
                    if (next < formula.Length &&
                        formula[next] == '(' &&
                        DynamicNameResolutionFunctions.Contains(word))
                    {
                        hasDynamicNameResolution = true;
                    }

                    continue;
                }

                if (current != '[')
                {
                    index++;
                    continue;
                }

                currentIdentifier.Clear();
                int start = index;
                bool hasSegment = false;
                while (index < formula.Length && formula[index] == '[')
                {
                    hasSegment = true;
                    if (!TryReadBracketSegment(formula, ref index, currentIdentifier))
                    {
                        return new MdxReferenceScanResult(
                            identifiers,
                            quotedLiterals,
                            hasDynamicNameResolution,
                            false);
                    }

                    if (index + 1 < formula.Length && formula[index] == '.' &&
                        formula[index + 1] == '[')
                    {
                        currentIdentifier.Append('.');
                        index++;
                        continue;
                    }

                    break;
                }

                if (!hasSegment || index <= start)
                {
                    return new MdxReferenceScanResult(
                        identifiers,
                        quotedLiterals,
                        hasDynamicNameResolution,
                        false);
                }

                identifiers.Add(currentIdentifier.ToString());
            }

            return new MdxReferenceScanResult(
                identifiers,
                quotedLiterals,
                hasDynamicNameResolution,
                true);
        }

        public static bool MightReference(string formula, string generatedSetName)
        {
            MdxReferenceScanResult scan = Scan(formula);
            if (!scan.IsComplete)
            {
                throw new InvalidOperationException(
                    "A live calculated-member formula could not be scanned safely.");
            }

            return scan.HasDynamicNameResolution ||
                   scan.BracketedIdentifiers.Any(identifier => string.Equals(
                       identifier,
                       generatedSetName,
                       StringComparison.OrdinalIgnoreCase)) ||
                   scan.QuotedLiterals.Any(literal => literal.IndexOf(
                       generatedSetName,
                       StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private static bool TryReadQuoted(
            string value,
            ref int index,
            char quote,
            out string literal)
        {
            var result = new StringBuilder();
            index++;
            while (index < value.Length)
            {
                if (value[index] != quote)
                {
                    result.Append(value[index]);
                    index++;
                    continue;
                }

                if (index + 1 < value.Length && value[index + 1] == quote)
                {
                    result.Append(quote);
                    index += 2;
                    continue;
                }

                index++;
                literal = result.ToString();
                return true;
            }

            literal = string.Empty;
            return false;
        }

        private static bool TryReadBracketSegment(
            string value,
            ref int index,
            StringBuilder result)
        {
            result.Append('[');
            index++;
            int content = 0;
            while (index < value.Length)
            {
                char current = value[index];
                if (current != ']')
                {
                    result.Append(current);
                    content++;
                    index++;
                    continue;
                }

                if (index + 1 < value.Length && value[index + 1] == ']')
                {
                    result.Append("]]");
                    content++;
                    index += 2;
                    continue;
                }

                result.Append(']');
                index++;
                return content > 0;
            }

            return false;
        }

        private static bool Peek(string value, int index, char expected)
        {
            return index + 1 < value.Length && value[index + 1] == expected;
        }
    }
}
