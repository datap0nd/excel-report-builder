[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SbomPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$PayloadPath,

    [ValidateNotNullOrEmpty()]
    [string]$ExpectedProductName = "Excel Report Builder",

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^\d+\.\d+\.\d+$')]
    [string]$ExpectedVersion
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest
. (Join-Path $PSScriptRoot "ReleaseSbom.Common.ps1")

$resolvedSbom = (Resolve-Path $SbomPath).Path
$resolvedPayload = (Resolve-Path $PayloadPath).Path
try {
    $document = Get-Content -LiteralPath $resolvedSbom -Raw | ConvertFrom-Json
}
catch {
    throw "The release SBOM is not valid JSON."
}

if ([string]$document.spdxVersion -notmatch '^SPDX-2\.') {
    throw "The release SBOM must use SPDX 2.x JSON."
}
if ([string]$document.dataLicense -ne "CC0-1.0") {
    throw "The release SBOM must use the SPDX CC0-1.0 document data license."
}
if ([string]$document.name -cne $ExpectedProductName) {
    throw (
        "The release SBOM document name must be $ExpectedProductName; " +
        "found $([string]$document.name).")
}

$documentNamespace = [string]$document.documentNamespace
$namespaceUri = $null
if (-not [Uri]::TryCreate($documentNamespace, [UriKind]::Absolute, [ref]$namespaceUri) -or
    $namespaceUri.Scheme -ne [Uri]::UriSchemeHttps) {
    throw "The release SBOM must contain an absolute HTTPS document namespace."
}

$fileEntries = @($document.files)
$packages = @($document.packages)
$relationships = @($document.relationships)
if ($fileEntries.Count -eq 0 -or
    $packages.Count -eq 0 -or
    $relationships.Count -eq 0) {
    throw "The release SBOM must contain file, package, and relationship inventories."
}

$payloadInventory = Get-ReleasePayloadInventory -ResolvedPayloadPath $resolvedPayload
$sbomInventory = Get-ReleaseSbomFileInventory -FileEntries $fileEntries
Assert-ExactReleaseFileInventory `
    -PayloadInventory $payloadInventory `
    -SbomInventory $sbomInventory

foreach ($relativePath in @($payloadInventory.ByPath.Keys | Sort-Object)) {
    $payloadFile = $payloadInventory.ByPath[$relativePath]
    $entry = $sbomInventory.ByPath[$relativePath]
    $checksumProperty = $entry.PSObject.Properties['checksums']
    $checksums = if ($null -eq $checksumProperty) { @() } else { @($entry.checksums) }
    $sha256Values = @($checksums |
        Where-Object { [string]$_.algorithm -eq "SHA256" } |
        ForEach-Object { [string]$_.checksumValue })
    $actualHash = Get-FileHash -LiteralPath $payloadFile.FullName -Algorithm SHA256
    $actualSha256 = $actualHash.Hash.ToLowerInvariant()
    if ($sha256Values.Count -ne 1) {
        throw (
            "The release SBOM must contain exactly one SHA-256 checksum for " +
            "$relativePath; found $($sha256Values.Count).")
    }

    if (-not [string]::Equals(
            $sha256Values[0],
            $actualSha256,
            [StringComparison]::OrdinalIgnoreCase)) {
        throw (
            "The release SBOM SHA-256 does not match staged payload file " +
            "$relativePath. SBOM=$($sha256Values[0]); payload=$actualSha256.")
    }
}

$documentId = Get-RequiredSbomStringProperty `
    -InputObject $document `
    -PropertyName "SPDXID" `
    -Description "The release SBOM document"
$elementIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
[void]$elementIds.Add($documentId)
foreach ($fileId in $sbomInventory.ById.Keys) {
    if (-not $elementIds.Add($fileId)) {
        throw "The release SBOM contains duplicate SPDXID $fileId."
    }
}

$packageById = [Collections.Generic.Dictionary[string, object]]::new(
    [StringComparer]::Ordinal)
foreach ($package in $packages) {
    if ($null -eq $package) {
        throw "The release SBOM contains a null package entry."
    }

    $packageId = Get-RequiredSbomStringProperty `
        -InputObject $package `
        -PropertyName "SPDXID" `
        -Description "Each release SBOM package"
    [void](Get-RequiredSbomStringProperty `
        -InputObject $package `
        -PropertyName "name" `
        -Description "Each release SBOM package")
    if (-not $elementIds.Add($packageId)) {
        throw "The release SBOM contains duplicate SPDXID $packageId."
    }

    $packageById.Add($packageId, $package)
}

$relationshipRecords = [Collections.Generic.List[object]]::new()
foreach ($relationship in $relationships) {
    if ($null -eq $relationship) {
        throw "The release SBOM contains a null relationship entry."
    }

    $sourceId = Get-RequiredSbomStringProperty `
        -InputObject $relationship `
        -PropertyName "spdxElementId" `
        -Description "Each release SBOM relationship"
    $targetId = Get-RequiredSbomStringProperty `
        -InputObject $relationship `
        -PropertyName "relatedSpdxElement" `
        -Description "Each release SBOM relationship"
    $relationshipType = Get-RequiredSbomStringProperty `
        -InputObject $relationship `
        -PropertyName "relationshipType" `
        -Description "Each release SBOM relationship"
    if (-not $elementIds.Contains($sourceId) -or -not $elementIds.Contains($targetId)) {
        throw (
            "The release SBOM relationship $sourceId $relationshipType $targetId " +
            "references an unknown SPDX element.")
    }

    $relationshipRecords.Add([pscustomobject]@{
        SourceId = $sourceId
        TargetId = $targetId
        Type = $relationshipType
    })
}

$describesRelationships = @($relationshipRecords | Where-Object {
    $_.SourceId -ceq $documentId -and $_.Type -ceq "DESCRIBES"
})
if ($describesRelationships.Count -ne 1) {
    throw "The release SBOM document must have exactly one DESCRIBES relationship."
}

$rootPackageId = $describesRelationships[0].TargetId
if (-not $packageById.ContainsKey($rootPackageId)) {
    throw "The release SBOM DESCRIBES relationship must target its root package."
}
$rootPackage = $packageById[$rootPackageId]
$rootPackageName = Get-RequiredSbomStringProperty `
    -InputObject $rootPackage `
    -PropertyName "name" `
    -Description "The release SBOM root package"
$rootPackageVersion = Get-RequiredSbomStringProperty `
    -InputObject $rootPackage `
    -PropertyName "versionInfo" `
    -Description "The release SBOM root package"
if ($rootPackageName -cne $ExpectedProductName -or
    $rootPackageVersion -cne $ExpectedVersion) {
    throw (
        "The release SBOM root package must be $ExpectedProductName $ExpectedVersion; " +
        "found $rootPackageName $rootPackageVersion.")
}

function Get-FileEvidencePackages {
    param([Parameter(Mandatory = $true)][string]$FileSpdxId)

    $evidence = [Collections.Generic.List[object]]::new()
    foreach ($relationshipRecord in $relationshipRecords) {
        # Syft 1.42.3 represents package-to-binary ownership with OTHER.
        if ($relationshipRecord.TargetId -ceq $FileSpdxId -and
            $relationshipRecord.Type -ceq "OTHER" -and
            $packageById.ContainsKey($relationshipRecord.SourceId)) {
            $evidence.Add($packageById[$relationshipRecord.SourceId])
        }
    }

    return @($evidence)
}

$expectedComponents = @(
    [pscustomobject]@{
        RelativePath = "addin/ExcelReportBuilder.AddIn.dll"
        PackageNames = @("Excel Report Builder")
        RuntimePackageName = $null
    },
    [pscustomobject]@{
        RelativePath = "addin/ExcelReportBuilder.Core.dll"
        PackageNames = @("Excel Report Builder")
        RuntimePackageName = $null
    },
    [pscustomobject]@{
        RelativePath = "addin/ExcelReportBuilder.Excel.dll"
        PackageNames = @("Excel Report Builder")
        RuntimePackageName = $null
    },
    [pscustomobject]@{
        RelativePath = "addin/ExcelReportBuilder.Agent.dll"
        PackageNames = @("Excel Report Builder")
        RuntimePackageName = $null
    },
    [pscustomobject]@{
        RelativePath = "worker-x64/ExcelReportBuilder.Worker.exe"
        PackageNames = @("Excel Report Builder", "ExcelReportBuilder.Worker")
        RuntimePackageName = "runtimepack.Microsoft.NETCore.App.Runtime.win-x64"
    },
    [pscustomobject]@{
        RelativePath = "worker-x86/ExcelReportBuilder.Worker.exe"
        PackageNames = @("Excel Report Builder", "ExcelReportBuilder.Worker")
        RuntimePackageName = "runtimepack.Microsoft.NETCore.App.Runtime.win-x86"
    }
)

$acceptableFirstPartyVersions = @($ExpectedVersion, "$ExpectedVersion.0")
foreach ($component in $expectedComponents) {
    if (-not $sbomInventory.ByPath.ContainsKey($component.RelativePath)) {
        throw "The staged release payload is missing required component $($component.RelativePath)."
    }

    $entry = $sbomInventory.ByPath[$component.RelativePath]
    $evidencePackages = @(Get-FileEvidencePackages -FileSpdxId ([string]$entry.SPDXID))
    $hasExpectedEvidence = @($evidencePackages | Where-Object {
        $name = [string]$_.name
        $versionProperty = $_.PSObject.Properties['versionInfo']
        $version = if ($null -eq $versionProperty) { "" } else { [string]$versionProperty.Value }
        $component.PackageNames -ccontains $name -and
            $acceptableFirstPartyVersions -ccontains $version
    }).Count -gt 0
    if (-not $hasExpectedEvidence) {
        throw (
            "The release SBOM has no correctly versioned package-to-file evidence for " +
            "$($component.RelativePath).")
    }

    if ($null -ne $component.RuntimePackageName) {
        $hasRuntimeEvidence = @($evidencePackages | Where-Object {
            $name = [string]$_.name
            $versionProperty = $_.PSObject.Properties['versionInfo']
            $version = if ($null -eq $versionProperty) { "" } else { [string]$versionProperty.Value }
            $name -ceq $component.RuntimePackageName -and
                $version -match '^\d+\.\d+\.\d+([-.+].*)?$'
        }).Count -gt 0
        if (-not $hasRuntimeEvidence) {
            throw (
                "The release SBOM has no architecture-specific runtime package evidence for " +
                "$($component.RelativePath).")
        }
    }
}

$packageNames = @($packages | ForEach-Object { [string]$_.name })
foreach ($requiredPackage in @(
    "Excel Report Builder",
    "Json.NET",
    "System.Text.Json"
)) {
    if ($packageNames -cnotcontains $requiredPackage) {
        throw "The release SBOM is missing expected package $requiredPackage."
    }
}

Write-Host (
    "Release SBOM for {0} {1} exactly covers {2} staged files, all first-party components, and both worker runtimes." -f
    $ExpectedProductName,
    $ExpectedVersion,
    $payloadInventory.Count)
