using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelReportBuilder.Excel.PivotPlus.Persistence
{
    /// <summary>
    /// The only generated Excel object kinds PivotTable+ can claim. A target
    /// worksheet and PivotTable are references and are never ownership records.
    /// </summary>
    public enum PivotPlusArtifactKind
    {
        Measure,
        NamedSet,
        Query,
        Connection,
        /// <summary>
        /// A generated workbook-scoped Excel Name used as a path-free source
        /// identity. The Name's RefersTo expression is never persisted here.
        /// </summary>
        WorkbookName,
        /// <summary>
        /// A generated, VeryHidden transaction worksheet. Ownership is valid
        /// only together with an exact worksheet CustomProperties marker.
        /// </summary>
        TemporaryWorksheet,
        /// <summary>
        /// A deterministic, generated PivotTable name used only while a
        /// destructive classic-to-Data-Model replacement is incomplete.
        /// </summary>
        TemporaryPivotTable
    }

    public enum PivotPlusRecoveryPhase
    {
        None,
        Planned,
        StagingVerified
    }

    public enum PivotPlusSemanticArtifactOperation
    {
        Create,
        Update,
        Delete
    }

    public enum PivotPlusFieldArea
    {
        Filter,
        Column,
        Row,
        Data
    }

    public sealed class PivotPlusOwnedArtifact
    {
        public PivotPlusArtifactKind Kind { get; set; }

        public string ArtifactId { get; set; } = string.Empty;

        /// <summary>
        /// Hash of the exact generated definition. The generated definition is
        /// deliberately not persisted in the ownership record.
        /// </summary>
        public string Fingerprint { get; set; } = string.Empty;
    }

    /// <summary>
    /// One hash-only artifact transition in a semantic Apply journal. The
    /// generated definition is deliberately absent; only its fingerprint is
    /// persisted. Active ownership remains in <see cref="PivotPlusWorkbookMetadata.Artifacts" />
    /// until the Apply commits.
    /// </summary>
    public sealed class PivotPlusSemanticArtifactTransition
    {
        public PivotPlusArtifactKind Kind { get; set; }

        public string ArtifactId { get; set; } = string.Empty;

        public PivotPlusSemanticArtifactOperation Operation { get; set; }

        /// <summary>
        /// Exact live fingerprint before an Update or Delete. It must be empty
        /// for Create.
        /// </summary>
        public string BeforeLiveFingerprint { get; set; } = string.Empty;

        public string PlannedDefinitionFingerprint { get; set; } = string.Empty;
    }

    /// <summary>
    /// Bounded write-ahead receipt for a semantic model Apply. All values are
    /// identifiers or canonical hashes; no generated definition, source member,
    /// workbook path, or cell payload is retained.
    /// </summary>
    public sealed class PivotPlusPendingSemanticApplyMetadata
    {
        public string ApplyId { get; set; } = string.Empty;

        public string PlanFingerprint { get; set; } = string.Empty;

        public string BeforePivotFingerprint { get; set; } = string.Empty;

        public string ExpectedPivotFingerprint { get; set; } = string.Empty;

        public IList<PivotPlusSemanticArtifactTransition> Transitions { get; set; } =
            new List<PivotPlusSemanticArtifactTransition>();
    }

    /// <summary>
    /// A hash-only description of one native PivotTable field placement before
    /// the last Apply. No source header, item caption, formula, or cell value is
    /// stored here.
    /// </summary>
    public sealed class PivotPlusUndoFieldPlacement
    {
        public string FieldFingerprint { get; set; } = string.Empty;

        public PivotPlusFieldArea Area { get; set; }

        public int Position { get; set; }
    }

    /// <summary>
    /// One-level undo metadata. The bounded created-artifact list lets cleanup
    /// require an exact id and content-fingerprint match. Previous placements
    /// contain hashes only and are likewise bounded.
    /// </summary>
    public sealed class PivotPlusUndoMetadata
    {
        public string ApplyId { get; set; } = string.Empty;

        public string BeforePivotFingerprint { get; set; } = string.Empty;

        public string AfterPivotFingerprint { get; set; } = string.Empty;

        public IList<PivotPlusOwnedArtifact> CreatedArtifacts { get; set; } =
            new List<PivotPlusOwnedArtifact>();

        public IList<PivotPlusUndoFieldPlacement> PreviousFieldPlacements { get; set; } =
            new List<PivotPlusUndoFieldPlacement>();
    }

    /// <summary>
    /// Path-free workbook metadata for one PivotTable+ setup. It intentionally
    /// has no workbook path, source data, formula, MDX, DAX, query text, or cell
    /// payload property.
    /// </summary>
    public sealed class PivotPlusWorkbookMetadata
    {
        public const string Version1_0 = "1.0";

        public const string Version1_1 = "1.1";

        public const string Version1_2 = "1.2";

        public const string Version1_3 = "1.3";

        public const string CurrentSchemaVersion = "1.4";

        public string SchemaVersion { get; set; } = CurrentSchemaVersion;

        public string SetupId { get; set; } = string.Empty;

        public string TargetWorksheetName { get; set; } = string.Empty;

        public string TargetPivotTableName { get; set; } = string.Empty;

        public PivotPlusRecoveryPhase RecoveryPhase { get; set; }

        /// <summary>
        /// Strict local A1 anchor used only while a conversion is recoverable.
        /// It is never a workbook path or external reference.
        /// </summary>
        public string TargetAnchorAddress { get; set; } = string.Empty;

        /// <summary>
        /// Hash of the verified staged native state. No field caption, formula,
        /// cell value, or source reference is persisted here.
        /// </summary>
        public string StagingStateFingerprint { get; set; } = string.Empty;

        public IList<PivotPlusOwnedArtifact> Artifacts { get; set; } =
            new List<PivotPlusOwnedArtifact>();

        /// <summary>
        /// At most one semantic Apply can be pending. This write-ahead receipt
        /// is mutually exclusive with classic-to-Data-Model recovery.
        /// </summary>
        public PivotPlusPendingSemanticApplyMetadata? PendingSemanticApply { get; set; }

        /// <summary>
        /// At most one Apply can be undone. Assigning a new record replaces the
        /// previous record instead of growing a history in the workbook.
        /// </summary>
        public PivotPlusUndoMetadata? Undo { get; set; }
    }

    /// <summary>
    /// Produces the canonical hash tokens accepted by PivotTable+ metadata.
    /// Persisted fingerprints contain no generated definition or workbook data.
    /// </summary>
    public static class PivotPlusFingerprint
    {
        private static readonly Regex ContractKindPattern = new Regex(
            "^[a-z0-9][a-z0-9._-]{0,63}$",
            RegexOptions.CultureInvariant);

        public static string Create(string contractKind, string value)
        {
            if (contractKind == null)
            {
                throw new ArgumentNullException(nameof(contractKind));
            }

            if (!ContractKindPattern.IsMatch(contractKind))
            {
                throw new ArgumentException(
                    "A canonical lower-case fingerprint contract kind is required.",
                    nameof(contractKind));
            }

            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            using (var sha256 = SHA256.Create())
            {
                var digest = sha256.ComputeHash(Encoding.UTF8.GetBytes(value));
                var result = new StringBuilder(digest.Length * 2);
                foreach (var item in digest)
                {
                    result.Append(item.ToString("x2", CultureInfo.InvariantCulture));
                }

                return contractKind + ":sha256:" + result;
            }
        }

        public static bool Matches(string? persistedFingerprint, string contractKind, string value)
        {
            return !string.IsNullOrWhiteSpace(persistedFingerprint) &&
                   string.Equals(
                       persistedFingerprint,
                       Create(contractKind, value),
                       StringComparison.Ordinal);
        }
    }

    internal static class PivotPlusMetadataValidator
    {
        public const int MaxArtifacts = 128;
        public const int MaxSemanticTransitions = 128;
        public const int MaxUndoArtifacts = 128;
        public const int MaxUndoFieldPlacements = 256;
        public const int MaxSerializedCharacters = 256 * 1024;

        private const int MaxIdLength = 128;
        private const int MaxArtifactNameLength = 255;
        private const int MaxPivotTableNameLength = 255;

        private static readonly Regex IdPattern = new Regex(
            "^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$",
            RegexOptions.CultureInvariant);

        private static readonly Regex FingerprintPattern = new Regex(
            "^[a-z0-9][a-z0-9._-]{0,63}:sha256:[0-9a-f]{64}$",
            RegexOptions.CultureInvariant);

        private static readonly Regex LocalA1Pattern = new Regex(
            "^([A-Z]{1,3})([1-9][0-9]{0,6})$",
            RegexOptions.CultureInvariant);

        public static void Validate(
            PivotPlusWorkbookMetadata metadata,
            bool requireCurrentVersion = true)
        {
            if (metadata == null)
            {
                throw new ArgumentNullException(nameof(metadata));
            }

            if ((requireCurrentVersion && !string.Equals(
                    metadata.SchemaVersion,
                    PivotPlusWorkbookMetadata.CurrentSchemaVersion,
                    StringComparison.Ordinal)) ||
                (!requireCurrentVersion && !IsSupportedSchemaVersion(metadata.SchemaVersion)))
            {
                throw new NotSupportedException("Unknown PivotTable+ metadata version.");
            }

            ValidateId(metadata.SetupId, "setup identifier");
            ValidateWorksheetName(metadata.TargetWorksheetName);
            ValidatePivotTableName(metadata.TargetPivotTableName);

            if (!Enum.IsDefined(typeof(PivotPlusRecoveryPhase), metadata.RecoveryPhase))
            {
                throw new ArgumentException("The recovery phase is invalid.", nameof(metadata));
            }

            if (metadata.Artifacts == null)
            {
                throw new ArgumentException("The artifact list cannot be null.", nameof(metadata));
            }

            if (metadata.Artifacts.Count > MaxArtifacts)
            {
                throw new ArgumentException(
                    "PivotTable+ metadata exceeds the owned-artifact limit.",
                    nameof(metadata));
            }

            var artifactKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var artifact in metadata.Artifacts)
            {
                ValidateArtifact(artifact, metadata.SchemaVersion, nameof(metadata));
                if (!artifactKeys.Add(ArtifactKey(artifact.Kind, artifact.ArtifactId)))
                {
                    throw new InvalidOperationException(
                        "PivotTable+ metadata contains a duplicate artifact identity.");
                }
            }

            ValidateRecovery(metadata);

            ValidatePendingSemanticApply(metadata);

            if (metadata.Undo != null)
            {
                ValidateUndo(metadata.Undo, metadata.Artifacts, metadata.SchemaVersion);
            }
        }

        public static bool IsSupportedSchemaVersion(string version)
        {
            return string.Equals(
                       version,
                       PivotPlusWorkbookMetadata.Version1_0,
                       StringComparison.Ordinal) ||
                   string.Equals(
                       version,
                       PivotPlusWorkbookMetadata.Version1_1,
                       StringComparison.Ordinal) ||
                   string.Equals(
                       version,
                       PivotPlusWorkbookMetadata.Version1_2,
                       StringComparison.Ordinal) ||
                   string.Equals(
                       version,
                       PivotPlusWorkbookMetadata.Version1_3,
                       StringComparison.Ordinal) ||
                   string.Equals(
                       version,
                       PivotPlusWorkbookMetadata.CurrentSchemaVersion,
                       StringComparison.Ordinal);
        }

        public static void ValidateId(string value, string description)
        {
            if (value == null || value.Length > MaxIdLength || !IdPattern.IsMatch(value))
            {
                throw new ArgumentException(
                    "A path-free " + description + " is required.",
                    description.Replace(' ', '_'));
            }
        }

        public static void ValidateFingerprint(string value, string description)
        {
            if (value == null || !FingerprintPattern.IsMatch(value))
            {
                throw new ArgumentException(
                    "A canonical SHA-256 " + description + " is required.",
                    description.Replace(' ', '_'));
            }
        }

        public static void ValidateArtifactName(string value)
        {
            if (string.IsNullOrWhiteSpace(value) ||
                value.Length > MaxArtifactNameLength ||
                !string.Equals(value, value.Trim(), StringComparison.Ordinal) ||
                value.Any(char.IsControl) ||
                value.IndexOf('/') >= 0 ||
                value.IndexOf('\\') >= 0 ||
                value.IndexOf(':') >= 0 ||
                value.IndexOf("file:", StringComparison.OrdinalIgnoreCase) >= 0 ||
                (value.Length >= 2 && char.IsLetter(value[0]) && value[1] == ':'))
            {
                throw new ArgumentException(
                    "A bounded path-free Excel artifact name is required.",
                    nameof(value));
            }
        }

        public static string ArtifactKey(PivotPlusArtifactKind kind, string artifactId)
        {
            return kind.ToString() + "\u001f" + artifactId;
        }

        private static void ValidateArtifact(
            PivotPlusOwnedArtifact artifact,
            string schemaVersion,
            string parameterName)
        {
            if (artifact == null)
            {
                throw new ArgumentException("An artifact cannot be null.", parameterName);
            }

            if (!Enum.IsDefined(typeof(PivotPlusArtifactKind), artifact.Kind))
            {
                throw new ArgumentException("The artifact kind is invalid.", parameterName);
            }

            if (artifact.Kind == PivotPlusArtifactKind.WorkbookName &&
                string.Equals(
                    schemaVersion,
                    PivotPlusWorkbookMetadata.Version1_0,
                    StringComparison.Ordinal))
            {
                throw new ArgumentException(
                    "Workbook-scoped source-name ownership requires PivotTable+ metadata version 1.1.",
                    parameterName);
            }

            if (artifact.Kind == PivotPlusArtifactKind.TemporaryWorksheet &&
                (string.Equals(
                     schemaVersion,
                     PivotPlusWorkbookMetadata.Version1_0,
                     StringComparison.Ordinal) ||
                 string.Equals(
                     schemaVersion,
                     PivotPlusWorkbookMetadata.Version1_1,
                     StringComparison.Ordinal)))
            {
                throw new ArgumentException(
                    "Temporary worksheet ownership requires PivotTable+ metadata version 1.2.",
                    parameterName);
            }

            if (artifact.Kind == PivotPlusArtifactKind.TemporaryPivotTable &&
                !SupportsRecovery(schemaVersion))
            {
                throw new ArgumentException(
                    "Temporary PivotTable ownership requires PivotTable+ metadata version 1.3.",
                    parameterName);
            }

            ValidateArtifactName(artifact.ArtifactId);
            ValidateFingerprint(artifact.Fingerprint, "artifact fingerprint");
        }

        private static void ValidateRecovery(PivotPlusWorkbookMetadata metadata)
        {
            bool supportsRecovery = SupportsRecovery(metadata.SchemaVersion);
            int temporaryWorksheets = metadata.Artifacts.Count(item =>
                item.Kind == PivotPlusArtifactKind.TemporaryWorksheet);
            int temporaryPivots = metadata.Artifacts.Count(item =>
                item.Kind == PivotPlusArtifactKind.TemporaryPivotTable);

            if (!supportsRecovery)
            {
                if (metadata.RecoveryPhase != PivotPlusRecoveryPhase.None ||
                    !string.IsNullOrEmpty(metadata.TargetAnchorAddress) ||
                    !string.IsNullOrEmpty(metadata.StagingStateFingerprint) ||
                    temporaryPivots != 0)
                {
                    throw new ArgumentException(
                        "Recovery state is not valid before PivotTable+ metadata version 1.3.",
                        nameof(metadata));
                }

                return;
            }

            if (metadata.RecoveryPhase == PivotPlusRecoveryPhase.None)
            {
                if (!string.IsNullOrEmpty(metadata.TargetAnchorAddress) ||
                    !string.IsNullOrEmpty(metadata.StagingStateFingerprint) ||
                    temporaryWorksheets != 0 ||
                    temporaryPivots != 0)
                {
                    throw new ArgumentException(
                        "Active PivotTable+ metadata cannot retain recovery-only state or temporary receipts.",
                        nameof(metadata));
                }

                return;
            }

            ValidateLocalA1Address(metadata.TargetAnchorAddress);
            if (temporaryWorksheets != 2 || temporaryPivots != 1)
            {
                throw new ArgumentException(
                    "Pending recovery requires exactly two temporary worksheets and one temporary PivotTable receipt.",
                    nameof(metadata));
            }

            if (metadata.RecoveryPhase == PivotPlusRecoveryPhase.Planned)
            {
                if (!string.IsNullOrEmpty(metadata.StagingStateFingerprint))
                {
                    throw new ArgumentException(
                        "Planned recovery cannot claim a verified staging state.",
                        nameof(metadata));
                }

                return;
            }

            ValidateFingerprint(
                metadata.StagingStateFingerprint,
                "staging state fingerprint");
        }

        private static bool SupportsRecovery(string schemaVersion)
        {
            return string.Equals(
                       schemaVersion,
                       PivotPlusWorkbookMetadata.Version1_3,
                       StringComparison.Ordinal) ||
                   string.Equals(
                       schemaVersion,
                       PivotPlusWorkbookMetadata.CurrentSchemaVersion,
                       StringComparison.Ordinal);
        }

        private static void ValidatePendingSemanticApply(PivotPlusWorkbookMetadata metadata)
        {
            PivotPlusPendingSemanticApplyMetadata? pending = metadata.PendingSemanticApply;
            if (pending == null)
            {
                return;
            }

            if (!string.Equals(
                    metadata.SchemaVersion,
                    PivotPlusWorkbookMetadata.CurrentSchemaVersion,
                    StringComparison.Ordinal))
            {
                throw new ArgumentException(
                    "Pending semantic Apply metadata requires PivotTable+ metadata version 1.4.",
                    nameof(metadata));
            }

            if (metadata.RecoveryPhase != PivotPlusRecoveryPhase.None)
            {
                throw new ArgumentException(
                    "A pending semantic Apply is mutually exclusive with conversion recovery.",
                    nameof(metadata));
            }

            ValidateId(pending.ApplyId, "semantic Apply identifier");
            ValidateFingerprint(pending.PlanFingerprint, "semantic plan fingerprint");
            ValidateFingerprint(
                pending.BeforePivotFingerprint,
                "before-PivotTable fingerprint");
            ValidateFingerprint(
                pending.ExpectedPivotFingerprint,
                "expected-PivotTable fingerprint");

            if (pending.Transitions == null)
            {
                throw new ArgumentException(
                    "The pending semantic transition list cannot be null.",
                    nameof(metadata));
            }

            if (pending.Transitions.Count > MaxSemanticTransitions)
            {
                throw new ArgumentException(
                    "A pending semantic Apply exceeds the transition limit.",
                    nameof(metadata));
            }

            var transitionKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (PivotPlusSemanticArtifactTransition transition in pending.Transitions)
            {
                if (transition == null)
                {
                    throw new ArgumentException(
                        "A pending semantic transition cannot be null.",
                        nameof(metadata));
                }

                if (transition.Kind != PivotPlusArtifactKind.Measure &&
                    transition.Kind != PivotPlusArtifactKind.NamedSet)
                {
                    throw new ArgumentException(
                        "A semantic transition can target only a measure or named set.",
                        nameof(metadata));
                }

                ValidateArtifactName(transition.ArtifactId);
                if (!Enum.IsDefined(
                        typeof(PivotPlusSemanticArtifactOperation),
                        transition.Operation))
                {
                    throw new ArgumentException(
                        "The semantic artifact operation is invalid.",
                        nameof(metadata));
                }

                ValidateFingerprint(
                    transition.PlannedDefinitionFingerprint,
                    "planned-definition fingerprint");

                string transitionKey = ArtifactKey(
                    transition.Kind,
                    transition.ArtifactId);
                if (!transitionKeys.Add(transitionKey))
                {
                    throw new InvalidOperationException(
                        "A pending semantic Apply contains a duplicate artifact identity.");
                }

                PivotPlusOwnedArtifact? prior = metadata.Artifacts.SingleOrDefault(artifact =>
                    artifact.Kind == transition.Kind &&
                    string.Equals(
                        artifact.ArtifactId,
                        transition.ArtifactId,
                        StringComparison.OrdinalIgnoreCase));

                if (transition.Operation == PivotPlusSemanticArtifactOperation.Create)
                {
                    if (!string.IsNullOrEmpty(transition.BeforeLiveFingerprint))
                    {
                        throw new ArgumentException(
                            "A semantic Create transition cannot contain a before-live fingerprint.",
                            nameof(metadata));
                    }

                    if (prior != null)
                    {
                        throw new InvalidOperationException(
                            "A semantic Create transition cannot replace a currently owned artifact.");
                    }

                    continue;
                }

                ValidateFingerprint(
                    transition.BeforeLiveFingerprint,
                    "before-live fingerprint");
                if (prior == null ||
                    !string.Equals(
                        prior.ArtifactId,
                        transition.ArtifactId,
                        StringComparison.Ordinal) ||
                    !string.Equals(
                        prior.Fingerprint,
                        transition.BeforeLiveFingerprint,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "A semantic Update or Delete transition must exactly match current owned-artifact truth.");
                }
            }
        }

        public static void ValidateLocalA1Address(string value)
        {
            Match match = value == null ? Match.Empty : LocalA1Pattern.Match(value);
            if (!match.Success ||
                !int.TryParse(
                    match.Groups[2].Value,
                    NumberStyles.None,
                    CultureInfo.InvariantCulture,
                    out int row) ||
                row > 1_048_576 ||
                ColumnNumber(match.Groups[1].Value) > 16_384)
            {
                throw new ArgumentException(
                    "A strict bounded local A1 target anchor is required.",
                    nameof(value));
            }
        }

        private static int ColumnNumber(string letters)
        {
            var result = 0;
            foreach (char letter in letters)
            {
                result = checked(result * 26 + (letter - 'A' + 1));
            }

            return result;
        }

        private static void ValidateUndo(
            PivotPlusUndoMetadata undo,
            IList<PivotPlusOwnedArtifact> ownedArtifacts,
            string schemaVersion)
        {
            ValidateId(undo.ApplyId, "Apply identifier");
            ValidateFingerprint(undo.BeforePivotFingerprint, "before-PivotTable fingerprint");
            ValidateFingerprint(undo.AfterPivotFingerprint, "after-PivotTable fingerprint");

            if (undo.CreatedArtifacts == null)
            {
                throw new ArgumentException("The undo artifact list cannot be null.", nameof(undo));
            }

            if (undo.CreatedArtifacts.Count > MaxUndoArtifacts)
            {
                throw new ArgumentException(
                    "PivotTable+ undo metadata exceeds the created-artifact limit.",
                    nameof(undo));
            }

            var createdKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var artifact in undo.CreatedArtifacts)
            {
                ValidateArtifact(artifact, schemaVersion, nameof(undo));
                var key = ArtifactKey(artifact.Kind, artifact.ArtifactId);
                if (!createdKeys.Add(key))
                {
                    throw new InvalidOperationException(
                        "PivotTable+ undo metadata contains a duplicate artifact identity.");
                }

                var exactOwnedMatch = ownedArtifacts.Any(owned =>
                    owned.Kind == artifact.Kind &&
                    string.Equals(owned.ArtifactId, artifact.ArtifactId, StringComparison.Ordinal) &&
                    string.Equals(owned.Fingerprint, artifact.Fingerprint, StringComparison.Ordinal));
                if (!exactOwnedMatch)
                {
                    throw new InvalidOperationException(
                        "Undo can reference only an exactly matching owned artifact.");
                }
            }

            if (undo.PreviousFieldPlacements == null)
            {
                throw new ArgumentException("The undo placement list cannot be null.", nameof(undo));
            }

            if (undo.PreviousFieldPlacements.Count > MaxUndoFieldPlacements)
            {
                throw new ArgumentException(
                    "PivotTable+ undo metadata exceeds the field-placement limit.",
                    nameof(undo));
            }

            var positions = new HashSet<string>(StringComparer.Ordinal);
            foreach (var placement in undo.PreviousFieldPlacements)
            {
                if (placement == null)
                {
                    throw new ArgumentException("An undo field placement cannot be null.", nameof(undo));
                }

                ValidateFingerprint(placement.FieldFingerprint, "field fingerprint");
                if (!Enum.IsDefined(typeof(PivotPlusFieldArea), placement.Area))
                {
                    throw new ArgumentException("The PivotTable field area is invalid.", nameof(undo));
                }

                if (placement.Position < 0 || placement.Position >= MaxUndoFieldPlacements)
                {
                    throw new ArgumentOutOfRangeException(
                        nameof(undo),
                        "A PivotTable field position is outside the bounded undo range.");
                }

                var positionKey = placement.Area.ToString() + "\u001f" +
                                  placement.Position.ToString(CultureInfo.InvariantCulture);
                if (!positions.Add(positionKey))
                {
                    throw new InvalidOperationException(
                        "PivotTable+ undo metadata contains a duplicate field position.");
                }
            }
        }

        private static void ValidateWorksheetName(string value)
        {
            ValidateTargetName(value, 31, "worksheet");
            if (value.IndexOfAny(new[] { '[', ']', ':', '*', '?' }) >= 0)
            {
                throw new ArgumentException("The target worksheet name is invalid.", nameof(value));
            }
        }

        private static void ValidatePivotTableName(string value)
        {
            ValidateTargetName(value, MaxPivotTableNameLength, "PivotTable");
        }

        private static void ValidateTargetName(string value, int maxLength, string description)
        {
            if (string.IsNullOrWhiteSpace(value) ||
                value.Length > maxLength ||
                !string.Equals(value, value.Trim(), StringComparison.Ordinal) ||
                value.IndexOf(':') >= 0 ||
                value.IndexOf('/') >= 0 ||
                value.IndexOf('\\') >= 0 ||
                value.IndexOf("//", StringComparison.Ordinal) >= 0 ||
                value.Any(char.IsControl))
            {
                throw new ArgumentException(
                    "A path-free target " + description + " name is required.",
                    nameof(value));
            }
        }
    }
}
