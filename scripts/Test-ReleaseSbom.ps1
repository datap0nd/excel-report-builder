[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SbomPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$PayloadPath
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

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

$fileEntries = @($document.files)
$packages = @($document.packages)
$relationships = @($document.relationships)
if ($fileEntries.Count -eq 0 -or $packages.Count -eq 0) {
    throw "The release SBOM must contain file and package inventories."
}

function Get-NormalizedSbomFileName {
    param([Parameter(Mandatory = $true)]$FileEntry)

    return ([string]$FileEntry.fileName).Replace('\', '/').TrimStart(
        [char[]]@('.', '/'))
}

function Get-SbomFileEntry {
    param([Parameter(Mandatory = $true)][string]$RelativePath)

    $normalizedPath = $RelativePath.Replace('\', '/').TrimStart([char[]]@('.', '/'))
    $matches = @($fileEntries | Where-Object {
        $candidate = Get-NormalizedSbomFileName -FileEntry $_
        [string]::Equals(
            $candidate,
            $normalizedPath,
            [StringComparison]::OrdinalIgnoreCase) -or
        $candidate.EndsWith(
            "/" + $normalizedPath,
            [StringComparison]::OrdinalIgnoreCase)
    })
    if ($matches.Count -ne 1) {
        throw "The release SBOM must contain exactly one file entry for $normalizedPath."
    }

    return $matches[0]
}

$payloadPrefix = $resolvedPayload.TrimEnd([char[]]@('\', '/')) + [IO.Path]::DirectorySeparatorChar
$payloadFiles = @(Get-ChildItem -LiteralPath $resolvedPayload -Recurse -File)
if ($payloadFiles.Count -eq 0) {
    throw "The staged release payload is empty."
}

foreach ($payloadFile in $payloadFiles) {
    if (-not $payloadFile.FullName.StartsWith(
            $payloadPrefix,
            [StringComparison]::OrdinalIgnoreCase)) {
        throw "A staged payload file resolved outside the payload root."
    }

    $relativePath = $payloadFile.FullName.Substring($payloadPrefix.Length).Replace('\', '/')
    $entry = Get-SbomFileEntry -RelativePath $relativePath
    $sha256Values = @($entry.checksums |
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

$packageById = @{}
foreach ($package in $packages) {
    $packageById[[string]$package.SPDXID] = $package
}

$expectedComponents = @(
    [pscustomobject]@{
        RelativePath = "addin/ExcelReportBuilder.AddIn.dll"
        PackageNames = @("Excel Report Builder")
    },
    [pscustomobject]@{
        RelativePath = "addin/ExcelReportBuilder.Core.dll"
        PackageNames = @("Excel Report Builder")
    },
    [pscustomobject]@{
        RelativePath = "addin/ExcelReportBuilder.Excel.dll"
        PackageNames = @("Excel Report Builder")
    },
    [pscustomobject]@{
        RelativePath = "addin/ExcelReportBuilder.Agent.dll"
        PackageNames = @("Excel Report Builder")
    },
    [pscustomobject]@{
        RelativePath = "worker-x64/ExcelReportBuilder.Worker.exe"
        PackageNames = @("Excel Report Builder", "ExcelReportBuilder.Worker")
    },
    [pscustomobject]@{
        RelativePath = "worker-x86/ExcelReportBuilder.Worker.exe"
        PackageNames = @("Excel Report Builder", "ExcelReportBuilder.Worker")
    }
)

foreach ($component in $expectedComponents) {
    $entry = Get-SbomFileEntry -RelativePath $component.RelativePath
    $evidencePackages = @($relationships |
        Where-Object { [string]$_.relatedSpdxElement -eq [string]$entry.SPDXID } |
        ForEach-Object {
            $packageId = [string]$_.spdxElementId
            if ($packageById.ContainsKey($packageId)) {
                [string]$packageById[$packageId].name
            }
        })
    $hasExpectedEvidence = @($evidencePackages |
        Where-Object { $component.PackageNames -contains $_ }).Count -gt 0
    if (-not $hasExpectedEvidence) {
        throw "The release SBOM has no package evidence for $($component.RelativePath)."
    }
}

$packageNames = @($packages | ForEach-Object { [string]$_.name })
foreach ($requiredPackage in @(
    "Excel Report Builder",
    "Json.NET",
    "System.Text.Json"
)) {
    if ($packageNames -notcontains $requiredPackage) {
        throw "The release SBOM is missing expected package $requiredPackage."
    }
}

Write-Host (
    "Release SBOM covers {0} staged files, all first-party components, and expected runtime packages." -f
    $payloadFiles.Count)
