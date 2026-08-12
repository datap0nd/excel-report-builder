using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml.Linq;

namespace ExcelReportBuilder.Excel.PivotPlus.Persistence
{
    /// <summary>
    /// Excel may report a CustomXMLParts.Add failure after inserting the part.
    /// Callers must preserve any newly-created workbook artifacts when the
    /// ownership outcome cannot be established safely.
    /// </summary>
    public sealed class PivotPlusOwnershipAmbiguousException : InvalidOperationException
    {
        internal PivotPlusOwnershipAmbiguousException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }

    /// <summary>
    /// Persists one strict, path-free Custom XML part per PivotTable+ setup.
    /// Ownership is granted only by an exact kind/id/fingerprint match.
    /// </summary>
    public sealed class PivotPlusWorkbookMetadataStore
    {
        public const string NamespaceUri =
            "urn:excel-report-builder:pivot-table-plus:metadata:v1";

        public void Save(dynamic workbook, PivotPlusWorkbookMetadata metadata)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            PivotPlusMetadataValidator.Validate(metadata, requireCurrentVersion: false);
            if (!string.Equals(
                    metadata.SchemaVersion,
                    PivotPlusWorkbookMetadata.CurrentSchemaVersion,
                    StringComparison.Ordinal))
            {
                metadata.SchemaVersion = PivotPlusWorkbookMetadata.CurrentSchemaVersion;
            }

            PivotPlusMetadataValidator.Validate(metadata);

            // Read and validate everything before replacing a part. This avoids
            // silently repairing ambiguous ownership or unknown schema versions.
            IReadOnlyList<PivotPlusWorkbookMetadata> existing = LoadAll((object)workbook);
            PivotPlusWorkbookMetadata? priorForSetup = existing.SingleOrDefault(item =>
                string.Equals(
                    item.SetupId,
                    metadata.SetupId,
                    StringComparison.OrdinalIgnoreCase));
            if (metadata.PendingSemanticApply != null &&
                priorForSetup != null &&
                !HaveExactArtifactTruth(
                    priorForSetup.Artifacts,
                    metadata.Artifacts))
            {
                throw new InvalidOperationException(
                    "A pending semantic Apply must preserve prior active owned-artifact truth until commit.");
            }

            var candidates = existing
                .Where(item => !string.Equals(item.SetupId, metadata.SetupId, StringComparison.OrdinalIgnoreCase))
                .Concat(new[] { metadata })
                .ToList();
            ValidateCollisions(candidates);

            string serialized = Serialize(metadata);
            IReadOnlyList<dynamic> priorParts = EnumerateParts((object)workbook)
                .Where(part => string.Equals(
                    ReadOwnedPart((string)part.XML).SetupId,
                    metadata.SetupId,
                    StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (priorParts.Count > 1)
            {
                throw new InvalidOperationException(
                    "More than one PivotTable+ metadata part claims the setup identifier.");
            }

            dynamic replacement;
            try
            {
                replacement = workbook.CustomXMLParts.Add(serialized);
                if (replacement == null)
                {
                    throw new InvalidOperationException(
                        "Excel did not return the new PivotTable+ metadata part.");
                }
            }
            catch (Exception addFailure)
            {
                IReadOnlyList<dynamic> afterFailure;
                try
                {
                    afterFailure = EnumerateParts((object)workbook)
                        .Where(part => string.Equals(
                            ReadOwnedPart((string)part.XML).SetupId,
                            metadata.SetupId,
                            StringComparison.OrdinalIgnoreCase))
                        .ToList();
                }
                catch (Exception inspectionFailure)
                {
                    throw new PivotPlusOwnershipAmbiguousException(
                        "Excel reported a metadata add failure and PivotTable+ could not determine whether durable ownership was inserted.",
                        new AggregateException(addFailure, inspectionFailure));
                }

                List<dynamic> insertedExact = afterFailure
                    .Where(part =>
                        string.Equals((string)part.XML, serialized, StringComparison.Ordinal) &&
                        !priorParts.Any(prior => SameNativePart(prior, part)))
                    .ToList();
                if (insertedExact.Count == 1)
                {
                    // CustomXMLParts.Add committed before surfacing a COM
                    // failure/null result. Continue the add-before-delete
                    // transaction with the exact inserted part.
                    replacement = insertedExact[0];
                }
                else if (insertedExact.Count == 0 &&
                         afterFailure.All(part =>
                             priorParts.Any(prior => SameNativePart(prior, part))))
                {
                    throw new InvalidOperationException(
                        "Excel could not add the replacement PivotTable+ metadata. The prior metadata was not removed.",
                        addFailure);
                }
                else
                {
                    throw new PivotPlusOwnershipAmbiguousException(
                        "Excel reported a metadata add failure and left an ambiguous set of ownership parts.",
                        addFailure);
                }
            }

            if (priorParts.Count == 0)
            {
                return;
            }

            try
            {
                priorParts[0].Delete();
            }
            catch (Exception deleteFailure)
            {
                IReadOnlyList<dynamic> survivingParts;
                try
                {
                    survivingParts = EnumerateParts((object)workbook)
                        .Where(part => string.Equals(
                            ReadOwnedPart((string)part.XML).SetupId,
                            metadata.SetupId,
                            StringComparison.OrdinalIgnoreCase))
                        .ToList();
                }
                catch (Exception inspectionFailure)
                {
                    // The replacement is deliberately retained when Excel will
                    // not reveal whether the prior part was deleted. Removing
                    // it here could leave the workbook with no ownership record.
                    throw new InvalidOperationException(
                        "Excel reported a metadata replacement failure and the surviving PivotTable+ metadata could not be verified. The replacement was retained.",
                        new AggregateException(deleteFailure, inspectionFailure));
                }

                if (survivingParts.Count == 1 &&
                    string.Equals((string)survivingParts[0].XML, serialized, StringComparison.Ordinal))
                {
                    // Some COM implementations report a failure after Delete
                    // has already committed. The desired replacement is the
                    // sole surviving part, so the transaction succeeded.
                    return;
                }

                Exception cause = deleteFailure;
                if (survivingParts.Count > 1)
                {
                    try
                    {
                        replacement.Delete();
                    }
                    catch (Exception rollbackFailure)
                    {
                        cause = new AggregateException(deleteFailure, rollbackFailure);
                    }
                }

                throw new InvalidOperationException(
                    "Excel could not replace the prior PivotTable+ metadata transactionally.",
                    cause);
            }
        }

        public IReadOnlyList<PivotPlusWorkbookMetadata> LoadAll(dynamic workbook)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            var result = new List<PivotPlusWorkbookMetadata>();
            foreach (var part in EnumerateParts((object)workbook))
            {
                result.Add(ReadOwnedPart((string)part.XML));
            }

            ValidateCollisions(result);
            return result
                .OrderBy(item => item.SetupId, StringComparer.Ordinal)
                .ToList();
        }

        public PivotPlusWorkbookMetadata? Load(dynamic workbook, string setupId)
        {
            PivotPlusMetadataValidator.ValidateId(setupId, "setup identifier");
            return LoadAll((object)workbook).SingleOrDefault(item =>
                string.Equals(item.SetupId, setupId, StringComparison.OrdinalIgnoreCase));
        }

        public PivotPlusWorkbookMetadata? LoadForTarget(
            dynamic workbook,
            string worksheetName,
            string pivotTableName)
        {
            // Reuse full validation without exposing a second, weaker target
            // validation path.
            var targetProbe = new PivotPlusWorkbookMetadata
            {
                SetupId = "target_probe",
                TargetWorksheetName = worksheetName,
                TargetPivotTableName = pivotTableName
            };
            PivotPlusMetadataValidator.Validate(targetProbe);

            return LoadAll((object)workbook).SingleOrDefault(item =>
                string.Equals(item.TargetWorksheetName, worksheetName, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(item.TargetPivotTableName, pivotTableName, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Returns true only for an exact owned-artifact record. The target
        /// PivotTable itself can never satisfy this guard.
        /// </summary>
        public bool IsOwnedArtifact(
            dynamic workbook,
            string setupId,
            PivotPlusArtifactKind kind,
            string artifactId,
            string fingerprint)
        {
            PivotPlusMetadataValidator.ValidateArtifactName(artifactId);
            PivotPlusMetadataValidator.ValidateFingerprint(fingerprint, "artifact fingerprint");

            PivotPlusWorkbookMetadata? metadata = Load((object)workbook, setupId);
            return metadata != null && metadata.Artifacts.Any(artifact =>
                artifact.Kind == kind &&
                string.Equals(artifact.ArtifactId, artifactId, StringComparison.Ordinal) &&
                string.Equals(artifact.Fingerprint, fingerprint, StringComparison.Ordinal));
        }

        public string Serialize(PivotPlusWorkbookMetadata metadata)
        {
            PivotPlusMetadataValidator.Validate(metadata);
            XNamespace ns = NamespaceUri;

            var root = new XElement(
                ns + "pivotTablePlusMetadata",
                new XAttribute("schemaVersion", metadata.SchemaVersion),
                new XAttribute("setupId", metadata.SetupId),
                new XElement(
                    ns + "target",
                    new XAttribute("worksheet", metadata.TargetWorksheetName),
                    new XAttribute("pivotTable", metadata.TargetPivotTableName)),
                new XElement(
                    ns + "artifacts",
                    OrderArtifacts(metadata.Artifacts).Select(artifact =>
                        SerializeArtifact(ns, artifact))));

            if (metadata.RecoveryPhase != PivotPlusRecoveryPhase.None)
            {
                root.Add(SerializeRecovery(ns, metadata));
            }

            if (metadata.PendingSemanticApply != null)
            {
                root.Add(SerializePendingSemanticApply(
                    ns,
                    metadata.PendingSemanticApply));
            }

            if (metadata.Undo != null)
            {
                root.Add(SerializeUndo(ns, metadata.Undo));
            }

            var xml = new XDocument(root).ToString(SaveOptions.DisableFormatting);
            if (xml.Length > PivotPlusMetadataValidator.MaxSerializedCharacters)
            {
                throw new InvalidOperationException(
                    "PivotTable+ metadata exceeds the serialized size limit.");
            }

            return xml;
        }

        private static XElement SerializeRecovery(
            XNamespace ns,
            PivotPlusWorkbookMetadata metadata)
        {
            var recovery = new XElement(
                ns + "recovery",
                new XAttribute("phase", FormatRecoveryPhase(metadata.RecoveryPhase)),
                new XAttribute("targetAnchor", metadata.TargetAnchorAddress));
            if (metadata.RecoveryPhase == PivotPlusRecoveryPhase.StagingVerified)
            {
                recovery.Add(new XAttribute(
                    "stagingStateFingerprint",
                    metadata.StagingStateFingerprint));
            }

            return recovery;
        }

        private static XElement SerializePendingSemanticApply(
            XNamespace ns,
            PivotPlusPendingSemanticApplyMetadata pending)
        {
            return new XElement(
                ns + "pendingSemanticApply",
                new XAttribute("applyId", pending.ApplyId),
                new XAttribute("planFingerprint", pending.PlanFingerprint),
                new XAttribute(
                    "beforePivotFingerprint",
                    pending.BeforePivotFingerprint),
                new XAttribute(
                    "expectedPivotFingerprint",
                    pending.ExpectedPivotFingerprint),
                new XElement(
                    ns + "transitions",
                    pending.Transitions
                        .OrderBy(transition => transition.Kind)
                        .ThenBy(
                            transition => transition.ArtifactId,
                            StringComparer.Ordinal)
                        .ThenBy(transition => transition.Operation)
                        .Select(transition =>
                            SerializeSemanticTransition(ns, transition))));
        }

        private static XElement SerializeSemanticTransition(
            XNamespace ns,
            PivotPlusSemanticArtifactTransition transition)
        {
            var element = new XElement(
                ns + "transition",
                new XAttribute("kind", FormatKind(transition.Kind)),
                new XAttribute("id", transition.ArtifactId),
                new XAttribute(
                    "operation",
                    FormatSemanticArtifactOperation(transition.Operation)),
                new XAttribute(
                    "plannedDefinitionFingerprint",
                    transition.PlannedDefinitionFingerprint));
            if (transition.Operation != PivotPlusSemanticArtifactOperation.Create)
            {
                element.Add(new XAttribute(
                    "beforeLiveFingerprint",
                    transition.BeforeLiveFingerprint));
            }

            return element;
        }

        private static XElement SerializeUndo(XNamespace ns, PivotPlusUndoMetadata undo)
        {
            return new XElement(
                ns + "undo",
                new XAttribute("applyId", undo.ApplyId),
                new XAttribute("beforePivotFingerprint", undo.BeforePivotFingerprint),
                new XAttribute("afterPivotFingerprint", undo.AfterPivotFingerprint),
                new XElement(
                    ns + "createdArtifacts",
                    OrderArtifacts(undo.CreatedArtifacts).Select(artifact =>
                        SerializeArtifact(ns, artifact))),
                new XElement(
                    ns + "previousFieldPlacements",
                    undo.PreviousFieldPlacements
                        .OrderBy(placement => placement.Area)
                        .ThenBy(placement => placement.Position)
                        .ThenBy(placement => placement.FieldFingerprint, StringComparer.Ordinal)
                        .Select(placement => new XElement(
                            ns + "field",
                            new XAttribute("fingerprint", placement.FieldFingerprint),
                            new XAttribute("area", FormatArea(placement.Area)),
                            new XAttribute(
                                "position",
                                placement.Position.ToString(CultureInfo.InvariantCulture))))));
        }

        private static IEnumerable<PivotPlusOwnedArtifact> OrderArtifacts(
            IEnumerable<PivotPlusOwnedArtifact> artifacts)
        {
            return artifacts
                .OrderBy(artifact => artifact.Kind)
                .ThenBy(artifact => artifact.ArtifactId, StringComparer.Ordinal)
                .ThenBy(artifact => artifact.Fingerprint, StringComparer.Ordinal);
        }

        private static bool HaveExactArtifactTruth(
            IList<PivotPlusOwnedArtifact> left,
            IList<PivotPlusOwnedArtifact> right)
        {
            return left.Count == right.Count && left.All(leftArtifact =>
                right.Any(rightArtifact =>
                    rightArtifact.Kind == leftArtifact.Kind &&
                    string.Equals(
                        rightArtifact.ArtifactId,
                        leftArtifact.ArtifactId,
                        StringComparison.Ordinal) &&
                    string.Equals(
                        rightArtifact.Fingerprint,
                        leftArtifact.Fingerprint,
                        StringComparison.Ordinal)));
        }

        private static bool SameNativePart(object left, object right)
        {
            if (ReferenceEquals(left, right)) return true;
            if (!Marshal.IsComObject(left) || !Marshal.IsComObject(right)) return false;

            IntPtr leftIdentity = IntPtr.Zero;
            IntPtr rightIdentity = IntPtr.Zero;
            try
            {
                leftIdentity = Marshal.GetIUnknownForObject(left);
                rightIdentity = Marshal.GetIUnknownForObject(right);
                return leftIdentity == rightIdentity;
            }
            finally
            {
                if (leftIdentity != IntPtr.Zero) Marshal.Release(leftIdentity);
                if (rightIdentity != IntPtr.Zero) Marshal.Release(rightIdentity);
            }
        }

        private static XElement SerializeArtifact(XNamespace ns, PivotPlusOwnedArtifact artifact)
        {
            return new XElement(
                ns + "artifact",
                new XAttribute("kind", FormatKind(artifact.Kind)),
                new XAttribute("id", artifact.ArtifactId),
                new XAttribute("fingerprint", artifact.Fingerprint));
        }

        private static PivotPlusWorkbookMetadata ReadOwnedPart(string xml)
        {
            try
            {
                if (xml == null || xml.Length > PivotPlusMetadataValidator.MaxSerializedCharacters)
                {
                    throw new InvalidOperationException(
                        "PivotTable+ metadata exceeds the serialized size limit.");
                }

                XNamespace ns = NamespaceUri;
                var document = XDocument.Parse(xml, LoadOptions.None);
                var root = document.Root;
                if (root == null || root.Name != ns + "pivotTablePlusMetadata")
                {
                    throw new InvalidOperationException("The managed metadata root is invalid.");
                }

                EnsureAttributes(root, "schemaVersion", "setupId");
                EnsureElementContent(
                    root,
                    ns + "target",
                    ns + "artifacts",
                    ns + "recovery",
                    ns + "pendingSemanticApply",
                    ns + "undo");
                var version = RequiredAttribute(root, "schemaVersion");
                if (!PivotPlusMetadataValidator.IsSupportedSchemaVersion(version))
                {
                    throw new NotSupportedException("Unknown PivotTable+ metadata version.");
                }

                var target = RequiredSingleElement(root, ns + "target");
                EnsureAttributes(target, "worksheet", "pivotTable");
                EnsureElementContent(target);

                var artifactContainer = RequiredSingleElement(root, ns + "artifacts");
                EnsureAttributes(artifactContainer);
                EnsureElementContent(artifactContainer, ns + "artifact");

                var metadata = new PivotPlusWorkbookMetadata
                {
                    SchemaVersion = version,
                    SetupId = RequiredAttribute(root, "setupId"),
                    TargetWorksheetName = RequiredAttribute(target, "worksheet"),
                    TargetPivotTableName = RequiredAttribute(target, "pivotTable"),
                    Artifacts = artifactContainer.Elements(ns + "artifact")
                        .Select(element => ReadArtifact(element, ns, version))
                        .ToList()
                };

                var recoveryElements = root.Elements(ns + "recovery").ToList();
                if (recoveryElements.Count > 1)
                {
                    throw new InvalidOperationException(
                        "PivotTable+ metadata can contain only one recovery checkpoint.");
                }

                if (recoveryElements.Count == 1)
                {
                    if (!SupportsRecovery(version))
                    {
                        throw new InvalidOperationException(
                            "Recovery checkpoints are not valid before metadata version 1.3.");
                    }

                    ReadRecovery(recoveryElements[0], metadata);
                }

                var pendingApplyElements = root.Elements(ns + "pendingSemanticApply").ToList();
                if (pendingApplyElements.Count > 1)
                {
                    throw new InvalidOperationException(
                        "PivotTable+ metadata can contain only one pending semantic Apply.");
                }

                if (pendingApplyElements.Count == 1)
                {
                    if (!string.Equals(
                            version,
                            PivotPlusWorkbookMetadata.CurrentSchemaVersion,
                            StringComparison.Ordinal))
                    {
                        throw new InvalidOperationException(
                            "Pending semantic Apply metadata is not valid before metadata version 1.4.");
                    }

                    metadata.PendingSemanticApply = ReadPendingSemanticApply(
                        pendingApplyElements[0],
                        ns,
                        version);
                }

                var undoElements = root.Elements(ns + "undo").ToList();
                if (undoElements.Count > 1)
                {
                    throw new InvalidOperationException(
                        "PivotTable+ metadata can contain only one undo record.");
                }

                if (undoElements.Count == 1)
                {
                    metadata.Undo = ReadUndo(undoElements[0], ns, version);
                }

                PivotPlusMetadataValidator.Validate(metadata, requireCurrentVersion: false);
                return metadata;
            }
            catch (NotSupportedException)
            {
                throw;
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    "Managed PivotTable+ metadata could not be read.",
                    exception);
            }
        }

        private static void ReadRecovery(
            XElement element,
            PivotPlusWorkbookMetadata metadata)
        {
            string phaseText = RequiredAttribute(element, "phase");
            PivotPlusRecoveryPhase phase = ParseRecoveryPhase(phaseText);
            if (phase == PivotPlusRecoveryPhase.Planned)
            {
                EnsureAttributes(element, "phase", "targetAnchor");
            }
            else if (phase == PivotPlusRecoveryPhase.StagingVerified)
            {
                EnsureAttributes(
                    element,
                    "phase",
                    "targetAnchor",
                    "stagingStateFingerprint");
            }
            else
            {
                throw new InvalidOperationException(
                    "A persisted recovery checkpoint cannot use the active phase.");
            }

            EnsureElementContent(element);
            metadata.RecoveryPhase = phase;
            metadata.TargetAnchorAddress = RequiredAttribute(element, "targetAnchor");
            metadata.StagingStateFingerprint =
                phase == PivotPlusRecoveryPhase.StagingVerified
                    ? RequiredAttribute(element, "stagingStateFingerprint")
                    : string.Empty;
        }

        private static PivotPlusOwnedArtifact ReadArtifact(
            XElement element,
            XNamespace ns,
            string schemaVersion)
        {
            EnsureAttributes(element, "kind", "id", "fingerprint");
            EnsureElementContent(element);
            return new PivotPlusOwnedArtifact
            {
                Kind = ParseKind(RequiredAttribute(element, "kind"), schemaVersion),
                ArtifactId = RequiredAttribute(element, "id"),
                Fingerprint = RequiredAttribute(element, "fingerprint")
            };
        }

        private static PivotPlusPendingSemanticApplyMetadata ReadPendingSemanticApply(
            XElement element,
            XNamespace ns,
            string schemaVersion)
        {
            EnsureAttributes(
                element,
                "applyId",
                "planFingerprint",
                "beforePivotFingerprint",
                "expectedPivotFingerprint");
            EnsureElementContent(element, ns + "transitions");

            XElement transitions = RequiredSingleElement(element, ns + "transitions");
            EnsureAttributes(transitions);
            EnsureElementContent(transitions, ns + "transition");

            return new PivotPlusPendingSemanticApplyMetadata
            {
                ApplyId = RequiredAttribute(element, "applyId"),
                PlanFingerprint = RequiredAttribute(element, "planFingerprint"),
                BeforePivotFingerprint = RequiredAttribute(
                    element,
                    "beforePivotFingerprint"),
                ExpectedPivotFingerprint = RequiredAttribute(
                    element,
                    "expectedPivotFingerprint"),
                Transitions = transitions.Elements(ns + "transition")
                    .Select(item => ReadSemanticTransition(item, schemaVersion))
                    .ToList()
            };
        }

        private static PivotPlusSemanticArtifactTransition ReadSemanticTransition(
            XElement element,
            string schemaVersion)
        {
            PivotPlusSemanticArtifactOperation operation = ParseSemanticArtifactOperation(
                RequiredAttribute(element, "operation"));
            if (operation == PivotPlusSemanticArtifactOperation.Create)
            {
                EnsureAttributes(
                    element,
                    "kind",
                    "id",
                    "operation",
                    "plannedDefinitionFingerprint");
            }
            else
            {
                EnsureAttributes(
                    element,
                    "kind",
                    "id",
                    "operation",
                    "beforeLiveFingerprint",
                    "plannedDefinitionFingerprint");
            }

            EnsureElementContent(element);
            return new PivotPlusSemanticArtifactTransition
            {
                Kind = ParseKind(RequiredAttribute(element, "kind"), schemaVersion),
                ArtifactId = RequiredAttribute(element, "id"),
                Operation = operation,
                BeforeLiveFingerprint =
                    operation == PivotPlusSemanticArtifactOperation.Create
                        ? string.Empty
                        : RequiredAttribute(element, "beforeLiveFingerprint"),
                PlannedDefinitionFingerprint = RequiredAttribute(
                    element,
                    "plannedDefinitionFingerprint")
            };
        }

        private static PivotPlusUndoMetadata ReadUndo(
            XElement element,
            XNamespace ns,
            string schemaVersion)
        {
            EnsureAttributes(
                element,
                "applyId",
                "beforePivotFingerprint",
                "afterPivotFingerprint");
            EnsureElementContent(
                element,
                ns + "createdArtifacts",
                ns + "previousFieldPlacements");

            var created = RequiredSingleElement(element, ns + "createdArtifacts");
            EnsureAttributes(created);
            EnsureElementContent(created, ns + "artifact");

            var previous = RequiredSingleElement(element, ns + "previousFieldPlacements");
            EnsureAttributes(previous);
            EnsureElementContent(previous, ns + "field");

            return new PivotPlusUndoMetadata
            {
                ApplyId = RequiredAttribute(element, "applyId"),
                BeforePivotFingerprint = RequiredAttribute(element, "beforePivotFingerprint"),
                AfterPivotFingerprint = RequiredAttribute(element, "afterPivotFingerprint"),
                CreatedArtifacts = created.Elements(ns + "artifact")
                    .Select(item => ReadArtifact(item, ns, schemaVersion))
                    .ToList(),
                PreviousFieldPlacements = previous.Elements(ns + "field")
                    .Select(ReadFieldPlacement)
                    .ToList()
            };
        }

        private static PivotPlusUndoFieldPlacement ReadFieldPlacement(XElement element)
        {
            EnsureAttributes(element, "fingerprint", "area", "position");
            EnsureElementContent(element);

            if (!int.TryParse(
                    RequiredAttribute(element, "position"),
                    NumberStyles.None,
                    CultureInfo.InvariantCulture,
                    out var position))
            {
                throw new InvalidOperationException("The undo field position is invalid.");
            }

            return new PivotPlusUndoFieldPlacement
            {
                FieldFingerprint = RequiredAttribute(element, "fingerprint"),
                Area = ParseArea(RequiredAttribute(element, "area")),
                Position = position
            };
        }

        private static void ValidateCollisions(
            IReadOnlyCollection<PivotPlusWorkbookMetadata> metadata)
        {
            var setupIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var targets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var artifactOwners = new Dictionary<string, string>(
                StringComparer.OrdinalIgnoreCase);

            foreach (var item in metadata)
            {
                PivotPlusMetadataValidator.Validate(item, requireCurrentVersion: false);
                if (!setupIds.Add(item.SetupId))
                {
                    throw new InvalidOperationException(
                        "Multiple PivotTable+ metadata parts claim the same setup identifier.");
                }

                var targetKey = item.TargetWorksheetName + "\u001f" + item.TargetPivotTableName;
                if (!targets.Add(targetKey))
                {
                    throw new InvalidOperationException(
                        "Multiple PivotTable+ setups reference the same target PivotTable.");
                }

                var reservedBySetup = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var artifact in item.Artifacts)
                {
                    reservedBySetup.Add(PivotPlusMetadataValidator.ArtifactKey(
                        artifact.Kind,
                        artifact.ArtifactId));
                }

                if (item.PendingSemanticApply != null)
                {
                    foreach (PivotPlusSemanticArtifactTransition transition in
                             item.PendingSemanticApply.Transitions)
                    {
                        reservedBySetup.Add(PivotPlusMetadataValidator.ArtifactKey(
                            transition.Kind,
                            transition.ArtifactId));
                    }
                }

                foreach (string artifactKey in reservedBySetup)
                {
                    if (artifactOwners.TryGetValue(artifactKey, out string? ownerSetupId) &&
                        !string.Equals(
                            ownerSetupId,
                            item.SetupId,
                            StringComparison.OrdinalIgnoreCase))
                    {
                        throw new InvalidOperationException(
                            "Multiple PivotTable+ setups claim the same generated artifact.");
                    }

                    artifactOwners[artifactKey] = item.SetupId;
                }
            }
        }

        private static string FormatKind(PivotPlusArtifactKind kind)
        {
            switch (kind)
            {
                case PivotPlusArtifactKind.Measure:
                    return "measure";
                case PivotPlusArtifactKind.NamedSet:
                    return "namedSet";
                case PivotPlusArtifactKind.Query:
                    return "query";
                case PivotPlusArtifactKind.Connection:
                    return "connection";
                case PivotPlusArtifactKind.WorkbookName:
                    return "workbookName";
                case PivotPlusArtifactKind.TemporaryWorksheet:
                    return "temporaryWorksheet";
                case PivotPlusArtifactKind.TemporaryPivotTable:
                    return "temporaryPivotTable";
                default:
                    throw new ArgumentOutOfRangeException(nameof(kind));
            }
        }

        private static PivotPlusArtifactKind ParseKind(string value, string schemaVersion)
        {
            switch (value)
            {
                case "measure":
                    return PivotPlusArtifactKind.Measure;
                case "namedSet":
                    return PivotPlusArtifactKind.NamedSet;
                case "query":
                    return PivotPlusArtifactKind.Query;
                case "connection":
                    return PivotPlusArtifactKind.Connection;
                case "workbookName":
                    if (string.Equals(
                            schemaVersion,
                            PivotPlusWorkbookMetadata.Version1_0,
                            StringComparison.Ordinal))
                    {
                        throw new InvalidOperationException(
                            "Workbook-scoped source-name ownership is not valid in metadata version 1.0.");
                    }

                    return PivotPlusArtifactKind.WorkbookName;
                case "temporaryWorksheet":
                    if (string.Equals(
                            schemaVersion,
                            PivotPlusWorkbookMetadata.Version1_0,
                            StringComparison.Ordinal) ||
                        string.Equals(
                            schemaVersion,
                            PivotPlusWorkbookMetadata.Version1_1,
                            StringComparison.Ordinal))
                    {
                        throw new InvalidOperationException(
                            "Temporary worksheet ownership is not valid before metadata version 1.2.");
                    }

                    return PivotPlusArtifactKind.TemporaryWorksheet;
                case "temporaryPivotTable":
                    if (!SupportsRecovery(schemaVersion))
                    {
                        throw new InvalidOperationException(
                            "Temporary PivotTable ownership is not valid before metadata version 1.3.");
                    }

                    return PivotPlusArtifactKind.TemporaryPivotTable;
                default:
                    throw new InvalidOperationException("The artifact kind is invalid.");
            }
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

        private static string FormatSemanticArtifactOperation(
            PivotPlusSemanticArtifactOperation operation)
        {
            switch (operation)
            {
                case PivotPlusSemanticArtifactOperation.Create:
                    return "create";
                case PivotPlusSemanticArtifactOperation.Update:
                    return "update";
                case PivotPlusSemanticArtifactOperation.Delete:
                    return "delete";
                default:
                    throw new ArgumentOutOfRangeException(nameof(operation));
            }
        }

        private static PivotPlusSemanticArtifactOperation ParseSemanticArtifactOperation(
            string value)
        {
            switch (value)
            {
                case "create":
                    return PivotPlusSemanticArtifactOperation.Create;
                case "update":
                    return PivotPlusSemanticArtifactOperation.Update;
                case "delete":
                    return PivotPlusSemanticArtifactOperation.Delete;
                default:
                    throw new InvalidOperationException(
                        "The semantic artifact operation is invalid.");
            }
        }

        private static string FormatRecoveryPhase(PivotPlusRecoveryPhase phase)
        {
            switch (phase)
            {
                case PivotPlusRecoveryPhase.Planned:
                    return "planned";
                case PivotPlusRecoveryPhase.StagingVerified:
                    return "stagingVerified";
                default:
                    throw new ArgumentOutOfRangeException(nameof(phase));
            }
        }

        private static PivotPlusRecoveryPhase ParseRecoveryPhase(string value)
        {
            switch (value)
            {
                case "planned":
                    return PivotPlusRecoveryPhase.Planned;
                case "stagingVerified":
                    return PivotPlusRecoveryPhase.StagingVerified;
                default:
                    throw new InvalidOperationException(
                        "The PivotTable+ recovery phase is invalid.");
            }
        }

        private static string FormatArea(PivotPlusFieldArea area)
        {
            switch (area)
            {
                case PivotPlusFieldArea.Filter:
                    return "filter";
                case PivotPlusFieldArea.Column:
                    return "column";
                case PivotPlusFieldArea.Row:
                    return "row";
                case PivotPlusFieldArea.Data:
                    return "data";
                default:
                    throw new ArgumentOutOfRangeException(nameof(area));
            }
        }

        private static PivotPlusFieldArea ParseArea(string value)
        {
            switch (value)
            {
                case "filter":
                    return PivotPlusFieldArea.Filter;
                case "column":
                    return PivotPlusFieldArea.Column;
                case "row":
                    return PivotPlusFieldArea.Row;
                case "data":
                    return PivotPlusFieldArea.Data;
                default:
                    throw new InvalidOperationException("The PivotTable field area is invalid.");
            }
        }

        private static XElement RequiredSingleElement(XElement parent, XName name)
        {
            var elements = parent.Elements(name).ToList();
            if (elements.Count != 1)
            {
                throw new InvalidOperationException(
                    "Managed PivotTable+ metadata has a missing or duplicate element.");
            }

            return elements[0];
        }

        private static string RequiredAttribute(XElement element, XName name)
        {
            var attribute = element.Attribute(name);
            if (attribute == null || string.IsNullOrEmpty(attribute.Value))
            {
                throw new InvalidOperationException(
                    "Managed PivotTable+ metadata has a missing attribute.");
            }

            return attribute.Value;
        }

        private static void EnsureAttributes(XElement element, params XName[] allowedNames)
        {
            var allowed = new HashSet<XName>(allowedNames);
            if (element.Attributes().Any(attribute =>
                    !attribute.IsNamespaceDeclaration && !allowed.Contains(attribute.Name)))
            {
                throw new InvalidOperationException(
                    "Managed PivotTable+ metadata contains an unknown attribute.");
            }
        }

        private static void EnsureElementContent(XElement element, params XName[] allowedChildren)
        {
            var allowed = new HashSet<XName>(allowedChildren);
            if (element.Elements().Any(child => !allowed.Contains(child.Name)) ||
                element.Nodes().OfType<XText>().Any(text =>
                    !string.IsNullOrWhiteSpace(text.Value)))
            {
                throw new InvalidOperationException(
                    "Managed PivotTable+ metadata contains an unknown payload.");
            }
        }

        private static IEnumerable<dynamic> EnumerateParts(dynamic workbook)
        {
            dynamic parts;
            try
            {
                parts = workbook.CustomXMLParts.SelectByNamespace(NamespaceUri);
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    "Excel did not expose PivotTable+ workbook metadata safely.",
                    exception);
            }

            int count;
            try
            {
                count = Convert.ToInt32(parts.Count, CultureInfo.InvariantCulture);
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    "Excel did not expose the PivotTable+ metadata part count.",
                    exception);
            }

            if (count < 0 || count > PivotPlusMetadataValidator.MaxArtifacts)
            {
                throw new InvalidOperationException(
                    "The workbook contains an unsupported number of PivotTable+ metadata parts.");
            }

            for (var index = 1; index <= count; index++)
            {
                dynamic part;
                try
                {
                    part = parts.Item(index);
                    if (part == null)
                    {
                        throw new InvalidOperationException(
                            "Excel returned a missing PivotTable+ metadata part.");
                    }
                }
                catch (Exception exception)
                {
                    throw new InvalidOperationException(
                        "Excel did not expose a PivotTable+ metadata part safely.",
                        exception);
                }

                yield return part;
            }
        }
    }
}
