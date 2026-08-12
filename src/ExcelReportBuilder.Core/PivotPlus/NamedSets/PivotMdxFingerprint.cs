using System;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace ExcelReportBuilder.Core.PivotPlus.NamedSets
{
    /// <summary>
    /// Versioned hashes for typed named-set definitions and exact transient MDX.
    /// Live Excel readback fingerprints belong to the host layer.
    /// </summary>
    public static class PivotMdxFingerprint
    {
        public const string FormulaAlgorithm = "namedset.formula.v1:sha256";
        public const string DefinitionAlgorithm = "namedset.definition.v2:sha256";
        public const string CompilationAlgorithm = "namedset.compilation.v1:sha256";

        public static string ComputeFormula(string mdxFormula)
        {
            if (mdxFormula == null) throw new ArgumentNullException(nameof(mdxFormula));

            var writer = new PivotNamedSetCanonicalWriter();
            writer.Add("formula", mdxFormula);
            return Hash(FormulaAlgorithm, writer.ToString());
        }

        internal static string ComputeDefinition(string canonicalDefinition)
        {
            if (canonicalDefinition == null)
            {
                throw new ArgumentNullException(nameof(canonicalDefinition));
            }

            return Hash(DefinitionAlgorithm, canonicalDefinition);
        }

        internal static string ComputeCompilation(
            string sourceFingerprint,
            System.Collections.Generic.IEnumerable<OwnedPivotNamedSetDefinition> namedSets)
        {
            if (sourceFingerprint == null)
            {
                throw new ArgumentNullException(nameof(sourceFingerprint));
            }

            if (namedSets == null) throw new ArgumentNullException(nameof(namedSets));

            var writer = new PivotNamedSetCanonicalWriter();
            writer.Add("sourceFingerprint", sourceFingerprint);
            int setIndex = 0;
            foreach (OwnedPivotNamedSetDefinition set in namedSets)
            {
                writer.Add("setIndex", setIndex.ToString(CultureInfo.InvariantCulture));
                writer.Add("definitionId", set.DefinitionId);
                writer.Add("displayOrder", set.DisplayOrder.ToString(CultureInfo.InvariantCulture));
                writer.Add("generatedSetName", set.GeneratedSetName);
                writer.Add("caption", set.Caption);
                writer.Add("axis", ((int)set.Axis).ToString(CultureInfo.InvariantCulture));
                writer.Add("dynamic", set.Dynamic ? "1" : "0");
                writer.Add("flattenHierarchies", set.FlattenHierarchies ? "1" : "0");
                writer.Add("hierarchizeDistinct", set.HierarchizeDistinct ? "1" : "0");
                writer.Add("definitionFingerprint", set.DefinitionFingerprint);
                writer.Add("formulaFingerprint", set.FormulaFingerprint);
                writer.Add(
                    "dependencyCount",
                    set.DirectMeasureDependencies.Count.ToString(CultureInfo.InvariantCulture));
                for (var dependencyIndex = 0;
                     dependencyIndex < set.DirectMeasureDependencies.Count;
                     dependencyIndex++)
                {
                    PivotNamedSetMeasureDependencyBinding dependency =
                        set.DirectMeasureDependencies[dependencyIndex];
                    writer.Add(
                        "dependencyIndex",
                        dependencyIndex.ToString(CultureInfo.InvariantCulture));
                    writer.Add("dependencyDefinitionId", dependency.DefinitionId);
                    writer.Add("dependencyGeneratedName", dependency.GeneratedMeasureName);
                    writer.Add(
                        "dependencyDefinitionFingerprint",
                        dependency.MeasureDefinitionFingerprint);
                    writer.Add(
                        "dependencyFormulaFingerprint",
                        dependency.MeasureFormulaFingerprint);
                }

                setIndex++;
            }

            writer.Add("setCount", setIndex.ToString(CultureInfo.InvariantCulture));
            return Hash(CompilationAlgorithm, writer.ToString());
        }

        private static string Hash(string algorithm, string canonical)
        {
            using (var sha256 = SHA256.Create())
            {
                byte[] bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(canonical));
                var result = new StringBuilder(bytes.Length * 2);
                foreach (byte value in bytes)
                {
                    result.Append(value.ToString("x2", CultureInfo.InvariantCulture));
                }

                return algorithm + ":" + result;
            }
        }
    }

    internal sealed class PivotNamedSetCanonicalWriter
    {
        private readonly StringBuilder builder = new StringBuilder();

        public void Add(string name, string value)
        {
            if (name == null) throw new ArgumentNullException(nameof(name));
            if (value == null) throw new ArgumentNullException(nameof(value));

            builder.Append(name.Length.ToString(CultureInfo.InvariantCulture));
            builder.Append(':');
            builder.Append(name);
            builder.Append('=');
            builder.Append(value.Length.ToString(CultureInfo.InvariantCulture));
            builder.Append(':');
            builder.Append(value);
            builder.Append(';');
        }

        public override string ToString()
        {
            return builder.ToString();
        }
    }
}
