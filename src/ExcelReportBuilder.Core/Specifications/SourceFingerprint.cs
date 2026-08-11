using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace ExcelReportBuilder.Core.Specifications
{
    /// <summary>
    /// A bounded fingerprint of the selected source. Only hashes and shape
    /// counts are persisted. Workbook paths, header text, and sampled values
    /// are never part of the contract.
    /// </summary>
    public sealed class SourceFingerprintSpec
    {
        public const string CurrentAlgorithm = "sha256-v1";

        public string Algorithm { get; set; } = CurrentAlgorithm;

        public string HeaderHash { get; set; } = string.Empty;

        public int ColumnCount { get; set; }

        /// <summary>
        /// Optional hash produced by the host from at most 64 sampled rows.
        /// The raw sampled values must not be persisted in ReportSpec JSON.
        /// </summary>
        public string? SampleHash { get; set; }

        public int? SampleRowCount { get; set; }

        /// <summary>
        /// Stable schema key suitable for namespacing saved report setups.
        /// It intentionally excludes the optional sample so normal data
        /// refreshes do not invalidate a compatible setup.
        /// </summary>
        public string GetSavedSetupKey()
        {
            return Algorithm + ":" + ColumnCount.ToString(CultureInfo.InvariantCulture) + ":" + HeaderHash;
        }
    }

    public static class SourceFingerprint
    {
        public static SourceFingerprintSpec FromHeaders(IEnumerable<string> headers)
        {
            if (headers == null)
            {
                throw new ArgumentNullException(nameof(headers));
            }

            var materialized = headers.ToList();
            if (materialized.Any(header => header == null))
            {
                throw new ArgumentException("Header names cannot be null.", nameof(headers));
            }

            return new SourceFingerprintSpec
            {
                Algorithm = SourceFingerprintSpec.CurrentAlgorithm,
                HeaderHash = ComputeHeaderHash(materialized),
                ColumnCount = materialized.Count
            };
        }

        public static string ComputeHeaderHash(IEnumerable<string> headers)
        {
            if (headers == null)
            {
                throw new ArgumentNullException(nameof(headers));
            }

            var materialized = headers.ToList();
            var canonical = new StringBuilder();
            canonical.Append(materialized.Count.ToString(CultureInfo.InvariantCulture));
            canonical.Append('|');
            foreach (var header in materialized)
            {
                if (header == null)
                {
                    throw new ArgumentException("Header names cannot be null.", nameof(headers));
                }

                var normalized = header.ToUpperInvariant();
                canonical.Append(normalized.Length.ToString(CultureInfo.InvariantCulture));
                canonical.Append(':');
                canonical.Append(normalized);
                canonical.Append('|');
            }

            using (var sha256 = SHA256.Create())
            {
                var bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(canonical.ToString()));
                var result = new StringBuilder(bytes.Length * 2);
                foreach (var value in bytes)
                {
                    result.Append(value.ToString("x2", CultureInfo.InvariantCulture));
                }

                return result.ToString();
            }
        }
    }
}
