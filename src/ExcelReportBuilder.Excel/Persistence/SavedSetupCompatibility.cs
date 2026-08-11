using System;
using ExcelReportBuilder.Core.Specifications;

namespace ExcelReportBuilder.Excel.Persistence
{
    /// <summary>
    /// Pure compatibility gate for restoring a saved setup. The caller must
    /// establish workbook-object equality using Excel, while this gate checks
    /// source kind and the path-free header fingerprint.
    /// </summary>
    public static class SavedSetupCompatibility
    {
        public static bool Matches(
            WorkbookSourceSpec savedSource,
            WorkbookSourceKind currentKind,
            SourceFingerprintSpec currentFingerprint,
            bool workbookObjectMatches)
        {
            if (savedSource == null)
            {
                throw new ArgumentNullException(nameof(savedSource));
            }

            if (currentFingerprint == null)
            {
                throw new ArgumentNullException(nameof(currentFingerprint));
            }

            SourceFingerprintSpec savedFingerprint = savedSource.Fingerprint;
            return workbookObjectMatches &&
                savedSource.Kind == currentKind &&
                savedFingerprint != null &&
                string.Equals(
                    savedFingerprint.Algorithm,
                    SourceFingerprintSpec.CurrentAlgorithm,
                    StringComparison.Ordinal) &&
                string.Equals(
                    currentFingerprint.Algorithm,
                    SourceFingerprintSpec.CurrentAlgorithm,
                    StringComparison.Ordinal) &&
                string.Equals(
                    savedFingerprint.GetSavedSetupKey(),
                    currentFingerprint.GetSavedSetupKey(),
                    StringComparison.Ordinal);
        }
    }
}
