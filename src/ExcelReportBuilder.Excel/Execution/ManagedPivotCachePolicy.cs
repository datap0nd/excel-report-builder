using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelReportBuilder.Excel.Ownership;

namespace ExcelReportBuilder.Excel.Execution
{
    internal enum ManagedPivotCacheAction
    {
        Create,
        Reuse,
        ReuseAndRegister,
        RetireAndCreate
    }

    internal sealed class PivotCacheSnapshot
    {
        public int Index { get; set; }

        public int SourceType { get; set; }

        public string? WorksheetSource { get; set; }

        public string? ConnectionName { get; set; }

        public int PivotTableCount { get; set; }
    }

    internal sealed class ManagedPivotCachePlan
    {
        public ManagedPivotCacheAction Action { get; set; }

        public ManagedObjectRecord? Registration { get; set; }

        public string Reason { get; set; } = string.Empty;
    }

    internal sealed class RegisteredPivotCacheBinding
    {
        public ManagedPivotCacheSlot Slot { get; set; } = null!;

        public ManagedObjectRecord Registration { get; set; } = null!;
    }

    internal sealed class ManagedPivotCacheSlot
    {
        private ManagedPivotCacheSlot(
            ManagedObjectIdentity identity,
            string registryName)
        {
            Identity = identity;
            RegistryName = registryName;
        }

        public ManagedObjectIdentity Identity { get; }

        public string RegistryName { get; }

        public static ManagedPivotCacheSlot For(
            string reportId,
            string blockOwnershipId,
            string logicalCacheName,
            CanonicalBackend backend)
        {
            if (string.IsNullOrWhiteSpace(reportId))
            {
                throw new ArgumentException("A report identifier is required.", nameof(reportId));
            }

            if (string.IsNullOrWhiteSpace(blockOwnershipId))
            {
                throw new ArgumentException("A block ownership identifier is required.", nameof(blockOwnershipId));
            }

            if (string.IsNullOrWhiteSpace(logicalCacheName))
            {
                throw new ArgumentException("A logical managed cache name is required.", nameof(logicalCacheName));
            }

            string suffix;
            switch (backend)
            {
                case CanonicalBackend.Worksheet:
                    suffix = "worksheet";
                    break;
                case CanonicalBackend.DataModel:
                    suffix = "model";
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(backend));
            }
            return new ManagedPivotCacheSlot(
                new ManagedObjectIdentity(
                    reportId,
                    blockOwnershipId + "_cache_" + suffix,
                    ManagedObjectKind.PivotCache),
                logicalCacheName + " [" + suffix + "]");
        }
    }

    internal sealed class PivotCacheSourceContract : IEquatable<PivotCacheSourceContract>
    {
        private const int WorksheetSourceType = 1;
        private const int ExternalSourceType = 2;

        private PivotCacheSourceContract(CanonicalBackend backend, string sourceName)
        {
            if (string.IsNullOrWhiteSpace(sourceName))
            {
                throw new ArgumentException("A PivotCache source name is required.", nameof(sourceName));
            }

            Backend = backend;
            SourceName = sourceName;
        }

        public CanonicalBackend Backend { get; }

        public string SourceName { get; }

        public string Serialized
        {
            get
            {
                var normalized = SourceName.ToUpperInvariant();
                var backend = Backend == CanonicalBackend.Worksheet ? "W" : "M";
                return backend + "|" +
                       normalized.Length.ToString(CultureInfo.InvariantCulture) + ":" +
                       normalized;
            }
        }

        public static PivotCacheSourceContract From(CanonicalLoadPlan source)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            return new PivotCacheSourceContract(source.Backend, source.TableOrConnectionName);
        }

        public static PivotCacheSourceContract Parse(string value)
        {
            if (string.IsNullOrWhiteSpace(value) || value.Length < 4 || value[1] != '|')
            {
                throw new InvalidOperationException("The managed PivotCache source contract is invalid.");
            }

            CanonicalBackend backend;
            switch (value[0])
            {
                case 'W': backend = CanonicalBackend.Worksheet; break;
                case 'M': backend = CanonicalBackend.DataModel; break;
                default:
                    throw new InvalidOperationException("The managed PivotCache source contract is invalid.");
            }

            var separator = value.IndexOf(':', 2);
            if (separator < 3 ||
                !int.TryParse(
                    value.Substring(2, separator - 2),
                    NumberStyles.None,
                    CultureInfo.InvariantCulture,
                    out var length) ||
                length < 1 ||
                value.Length - separator - 1 != length)
            {
                throw new InvalidOperationException("The managed PivotCache source contract is invalid.");
            }

            return new PivotCacheSourceContract(backend, value.Substring(separator + 1));
        }

        public bool Matches(PivotCacheSnapshot snapshot)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            if (Backend == CanonicalBackend.DataModel)
            {
                return snapshot.SourceType == ExternalSourceType &&
                       string.Equals(
                           snapshot.ConnectionName,
                           SourceName,
                           StringComparison.OrdinalIgnoreCase);
            }

            return snapshot.SourceType == WorksheetSourceType &&
                   WorksheetSourceMatches(snapshot.WorksheetSource, SourceName);
        }

        public bool Equals(PivotCacheSourceContract? other)
        {
            return other != null &&
                   Backend == other.Backend &&
                   string.Equals(SourceName, other.SourceName, StringComparison.OrdinalIgnoreCase);
        }

        public override bool Equals(object? obj)
        {
            return Equals(obj as PivotCacheSourceContract);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return ((int)Backend * 397) ^
                       StringComparer.OrdinalIgnoreCase.GetHashCode(SourceName);
            }
        }

        internal static bool WorksheetSourceMatches(string? actual, string expected)
        {
            if (string.IsNullOrWhiteSpace(actual))
            {
                return false;
            }

            var candidate = actual!.Trim();
            if (candidate.StartsWith("=", StringComparison.Ordinal))
            {
                candidate = candidate.Substring(1).Trim();
            }

            if (string.Equals(candidate, expected, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            var bang = candidate.LastIndexOf('!');
            if (bang >= 0 && bang + 1 < candidate.Length)
            {
                candidate = candidate.Substring(bang + 1).Trim();
            }

            candidate = candidate.Trim('\'', '"');
            const string allRowsSuffix = "[#All]";
            if (candidate.EndsWith(allRowsSuffix, StringComparison.OrdinalIgnoreCase))
            {
                candidate = candidate.Substring(0, candidate.Length - allRowsSuffix.Length);
            }

            return string.Equals(candidate, expected, StringComparison.OrdinalIgnoreCase);
        }
    }

    internal static class ManagedPivotCachePolicy
    {
        public static ManagedPivotCachePlan Plan(
            IReadOnlyList<ManagedObjectRecord> records,
            ManagedObjectIdentity identity,
            string managedCacheName,
            PivotCacheSourceContract requestedSource,
            bool managedPivotExists,
            PivotCacheSnapshot? candidate)
        {
            if (records == null) throw new ArgumentNullException(nameof(records));
            if (identity == null) throw new ArgumentNullException(nameof(identity));
            if (identity.Kind != ManagedObjectKind.PivotCache)
            {
                throw new ArgumentException("A PivotCache ownership identity is required.", nameof(identity));
            }

            if (string.IsNullOrWhiteSpace(managedCacheName))
            {
                throw new ArgumentException("A managed PivotCache name is required.", nameof(managedCacheName));
            }

            if (requestedSource == null) throw new ArgumentNullException(nameof(requestedSource));

            var exact = records.Where(record => SameIdentity(record, identity)).ToList();
            if (exact.Count > 1)
            {
                throw new InvalidOperationException(
                    "More than one ownership record claims the managed PivotCache identity.");
            }

            if (records.Any(record =>
                    record.Kind == ManagedObjectKind.PivotCache &&
                    !SameIdentity(record, identity) &&
                    string.Equals(record.ExcelName, managedCacheName, StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidOperationException(
                    "The requested managed PivotCache name is already owned by another object.");
            }

            var registration = exact.SingleOrDefault();
            if (registration == null)
            {
                if (managedPivotExists &&
                    candidate != null &&
                    candidate.Index > 0 &&
                    candidate.PivotTableCount == 1 &&
                    requestedSource.Matches(candidate))
                {
                    return new ManagedPivotCachePlan
                    {
                        Action = ManagedPivotCacheAction.ReuseAndRegister,
                        Reason = "The exact managed PivotTable exclusively identifies a compatible legacy cache."
                    };
                }

                return new ManagedPivotCachePlan
                {
                    Action = ManagedPivotCacheAction.Create,
                    Reason = "No exact managed cache registration can be reused."
                };
            }

            if (!string.Equals(registration.ExcelName, managedCacheName, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The managed PivotCache identity is registered under a different name.");
            }

            if (!int.TryParse(
                    registration.Locator,
                    NumberStyles.None,
                    CultureInfo.InvariantCulture,
                    out var registeredIndex) ||
                registeredIndex < 1)
            {
                throw new InvalidOperationException(
                    "The managed PivotCache ownership locator is invalid.");
            }

            var registeredSource = PivotCacheSourceContract.Parse(
                registration.SourceContract ?? string.Empty);
            if (candidate == null)
            {
                if (managedPivotExists)
                {
                    throw new InvalidOperationException(
                        "The managed PivotTable exists but its registered PivotCache cannot be located.");
                }

                return new ManagedPivotCachePlan
                {
                    Action = ManagedPivotCacheAction.RetireAndCreate,
                    Registration = registration,
                    Reason = "The exact managed cache registration no longer resolves to a workbook cache."
                };
            }

            if (candidate.Index != registeredIndex)
            {
                throw new InvalidOperationException(
                    "The managed PivotTable and PivotCache ownership locator disagree.");
            }

            if (!registeredSource.Matches(candidate))
            {
                throw new InvalidOperationException(
                    "The registered PivotCache no longer matches its owned source contract.");
            }

            if (managedPivotExists && candidate.PivotTableCount < 1)
            {
                throw new InvalidOperationException(
                    "The managed PivotTable is not attached to its registered PivotCache.");
            }

            if ((!managedPivotExists && candidate.PivotTableCount != 0) ||
                (managedPivotExists && candidate.PivotTableCount != 1))
            {
                return new ManagedPivotCachePlan
                {
                    Action = ManagedPivotCacheAction.RetireAndCreate,
                    Registration = registration,
                    Reason = "The registered cache is shared with another PivotTable and cannot be changed autonomously."
                };
            }

            if (!registeredSource.Equals(requestedSource))
            {
                return new ManagedPivotCachePlan
                {
                    Action = ManagedPivotCacheAction.RetireAndCreate,
                    Registration = registration,
                    Reason = "The validated report source changed since the cache was registered."
                };
            }

            return new ManagedPivotCachePlan
            {
                Action = ManagedPivotCacheAction.Reuse,
                Registration = registration,
                Reason = "The exact owned PivotCache and source contract match the validated block."
            };
        }

        private static bool SameIdentity(
            ManagedObjectRecord record,
            ManagedObjectIdentity identity)
        {
            return string.Equals(record.ReportId, identity.ReportId, StringComparison.Ordinal) &&
                   string.Equals(record.ObjectId, identity.ObjectId, StringComparison.Ordinal) &&
                   record.Kind == identity.Kind;
        }
    }
}
