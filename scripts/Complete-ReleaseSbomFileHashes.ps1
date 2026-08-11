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

$fileEntries = @($document.files)
if ($fileEntries.Count -eq 0) {
    throw "The release SBOM must contain a file inventory."
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

$addedHashCount = 0
foreach ($payloadFile in $payloadFiles) {
    if (-not $payloadFile.FullName.StartsWith(
            $payloadPrefix,
            [StringComparison]::OrdinalIgnoreCase)) {
        throw "A staged payload file resolved outside the payload root."
    }

    $relativePath = $payloadFile.FullName.Substring($payloadPrefix.Length).Replace('\', '/')
    $entry = Get-SbomFileEntry -RelativePath $relativePath
    $checksumProperty = $entry.PSObject.Properties['checksums']
    $checksums = if ($null -eq $checksumProperty) { @() } else { @($entry.checksums) }
    $sha256Values = @($checksums |
        Where-Object { [string]$_.algorithm -eq "SHA256" } |
        ForEach-Object { [string]$_.checksumValue })
    if ($sha256Values.Count -gt 1) {
        throw "The release SBOM contains duplicate SHA-256 checksums for $relativePath."
    }

    $actualHash = Get-FileHash -LiteralPath $payloadFile.FullName -Algorithm SHA256
    $actualSha256 = $actualHash.Hash.ToLowerInvariant()
    if ($sha256Values.Count -eq 1) {
        if (-not [string]::Equals(
                $sha256Values[0],
                $actualSha256,
                [StringComparison]::OrdinalIgnoreCase)) {
            throw "Syft supplied an incorrect SHA-256 checksum for $relativePath."
        }

        continue
    }

    $completedChecksums = @($checksums) + [pscustomobject]@{
        algorithm = "SHA256"
        checksumValue = $actualSha256
    }
    if ($null -eq $checksumProperty) {
        $entry | Add-Member -NotePropertyName checksums -NotePropertyValue $completedChecksums
    }
    else {
        $entry.checksums = $completedChecksums
    }
    $addedHashCount++
}

$json = $document | ConvertTo-Json -Depth 100
Set-Content -LiteralPath $resolvedSbom -Value $json -Encoding utf8NoBOM
Write-Host (
    "Completed SHA-256 coverage for {0} payload files; added {1} missing hashes." -f
    $payloadFiles.Count,
    $addedHashCount)
