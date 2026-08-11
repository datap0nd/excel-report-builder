Set-StrictMode -Version Latest

function Get-RequiredSbomStringProperty {
    param(
        [Parameter(Mandatory = $true)]$InputObject,
        [Parameter(Mandatory = $true)][string]$PropertyName,
        [Parameter(Mandatory = $true)][string]$Description
    )

    $property = $InputObject.PSObject.Properties[$PropertyName]
    if ($null -eq $property -or [string]::IsNullOrWhiteSpace([string]$property.Value)) {
        throw "$Description must contain a non-empty $PropertyName property."
    }

    return [string]$property.Value
}

function ConvertTo-CanonicalReleaseRelativePath {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Description
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw "$Description must not be empty."
    }

    $normalized = $Path.Replace('\', '/')
    if ([IO.Path]::IsPathRooted($Path) -or
        $normalized.StartsWith('/', [StringComparison]::Ordinal) -or
        $normalized -match '^[A-Za-z]:') {
        throw "$Description must be a root-relative path: $Path"
    }

    $segments = @($normalized.Split(
        [char[]]@('/'),
        [StringSplitOptions]::None))
    foreach ($segment in $segments) {
        if ([string]::IsNullOrEmpty($segment) -or $segment -in @('.', '..')) {
            throw "$Description contains an empty or dot path segment: $Path"
        }
    }

    return [string]::Join('/', $segments)
}

function Get-ReleasePayloadInventory {
    param([Parameter(Mandatory = $true)][string]$ResolvedPayloadPath)

    $payloadPrefix = $ResolvedPayloadPath.TrimEnd(
        [char[]]@('\', '/')) + [IO.Path]::DirectorySeparatorChar
    $payloadFiles = @(Get-ChildItem -LiteralPath $ResolvedPayloadPath -Recurse -File)
    if ($payloadFiles.Count -eq 0) {
        throw "The staged release payload is empty."
    }

    $byPath = [Collections.Generic.Dictionary[string, IO.FileInfo]]::new(
        [StringComparer]::OrdinalIgnoreCase)
    foreach ($payloadFile in $payloadFiles) {
        if (-not $payloadFile.FullName.StartsWith(
                $payloadPrefix,
                [StringComparison]::OrdinalIgnoreCase)) {
            throw "A staged payload file resolved outside the payload root."
        }

        $relativePath = $payloadFile.FullName.Substring($payloadPrefix.Length)
        $canonicalPath = ConvertTo-CanonicalReleaseRelativePath `
            -Path $relativePath `
            -Description "Staged payload file path"
        if ($byPath.ContainsKey($canonicalPath)) {
            throw "The staged release payload contains a duplicate path: $canonicalPath"
        }

        $byPath.Add($canonicalPath, $payloadFile)
    }

    $directoryPaths = [Collections.Generic.HashSet[string]]::new(
        [StringComparer]::OrdinalIgnoreCase)
    foreach ($payloadDirectory in @(
            Get-ChildItem -LiteralPath $ResolvedPayloadPath -Recurse -Directory)) {
        if (-not $payloadDirectory.FullName.StartsWith(
                $payloadPrefix,
                [StringComparison]::OrdinalIgnoreCase)) {
            throw "A staged payload directory resolved outside the payload root."
        }

        $relativePath = $payloadDirectory.FullName.Substring($payloadPrefix.Length)
        $canonicalPath = ConvertTo-CanonicalReleaseRelativePath `
            -Path $relativePath `
            -Description "Staged payload directory path"
        if ($byPath.ContainsKey($canonicalPath) -or -not $directoryPaths.Add($canonicalPath)) {
            throw "The staged release payload contains a duplicate path: $canonicalPath"
        }
    }

    return [pscustomobject]@{
        ByPath = $byPath
        DirectoryPaths = $directoryPaths
        Count = $byPath.Count
        InventoryCount = $byPath.Count + $directoryPaths.Count
    }
}

function Get-ReleaseSbomFileInventory {
    param([Parameter(Mandatory = $true)][object[]]$FileEntries)

    $byPath = [Collections.Generic.Dictionary[string, object]]::new(
        [StringComparer]::OrdinalIgnoreCase)
    $byId = [Collections.Generic.Dictionary[string, object]]::new(
        [StringComparer]::Ordinal)
    $rootEntry = $null

    foreach ($fileEntry in $FileEntries) {
        if ($null -eq $fileEntry) {
            throw "The release SBOM contains a null file entry."
        }

        $spdxId = Get-RequiredSbomStringProperty `
            -InputObject $fileEntry `
            -PropertyName "SPDXID" `
            -Description "Each release SBOM file entry"
        if ($byId.ContainsKey($spdxId)) {
            throw "The release SBOM contains duplicate file SPDXID $spdxId."
        }
        $byId.Add($spdxId, $fileEntry)

        $fileNameProperty = $fileEntry.PSObject.Properties['fileName']
        if ($null -eq $fileNameProperty) {
            throw "Each release SBOM file entry must contain a fileName property."
        }

        $fileName = [string]$fileNameProperty.Value
        if ($fileName.Length -eq 0) {
            if ($null -ne $rootEntry) {
                throw "The release SBOM contains duplicate empty source-root file entries."
            }

            # Syft emits one empty file entry for the scanned directory itself.
            $rootEntry = $fileEntry
            continue
        }

        $canonicalPath = ConvertTo-CanonicalReleaseRelativePath `
            -Path $fileName `
            -Description "Release SBOM fileName"
        if ($byPath.ContainsKey($canonicalPath)) {
            throw "The release SBOM contains a duplicate file path: $canonicalPath"
        }

        $byPath.Add($canonicalPath, $fileEntry)
    }

    return [pscustomobject]@{
        ByPath = $byPath
        ById = $byId
        Count = $byPath.Count
        RootEntry = $rootEntry
    }
}

function Assert-ExactReleaseFileInventory {
    param(
        [Parameter(Mandatory = $true)]$PayloadInventory,
        [Parameter(Mandatory = $true)]$SbomInventory
    )

    $missingFromSbom = [Collections.Generic.List[string]]::new()
    $expectedPaths = [Collections.Generic.HashSet[string]]::new(
        [StringComparer]::OrdinalIgnoreCase)
    foreach ($payloadPath in $PayloadInventory.ByPath.Keys) {
        [void]$expectedPaths.Add($payloadPath)
    }
    foreach ($directoryPath in $PayloadInventory.DirectoryPaths) {
        [void]$expectedPaths.Add($directoryPath)
    }

    foreach ($payloadPath in $expectedPaths) {
        if (-not $SbomInventory.ByPath.ContainsKey($payloadPath)) {
            $missingFromSbom.Add($payloadPath)
        }
    }

    $extraInSbom = [Collections.Generic.List[string]]::new()
    foreach ($sbomPath in $SbomInventory.ByPath.Keys) {
        if (-not $expectedPaths.Contains($sbomPath)) {
            $extraInSbom.Add($sbomPath)
        }
    }

    if ($missingFromSbom.Count -gt 0 -or $extraInSbom.Count -gt 0) {
        $missingText = if ($missingFromSbom.Count -eq 0) {
            "none"
        }
        else {
            (@($missingFromSbom) | Sort-Object) -join ', '
        }
        $extraText = if ($extraInSbom.Count -eq 0) {
            "none"
        }
        else {
            (@($extraInSbom) | Sort-Object) -join ', '
        }
        throw (
            "The release SBOM file inventory must exactly match the staged payload. " +
            "Missing from SBOM: $missingText. Extra in SBOM: $extraText.")
    }

    if ($PayloadInventory.InventoryCount -ne $SbomInventory.Count) {
        throw (
            "The release SBOM file inventory is not one-to-one with the staged payload. " +
            "Payload=$($PayloadInventory.InventoryCount); SBOM=$($SbomInventory.Count).")
    }

    if ($null -eq $SbomInventory.RootEntry) {
        throw "The release SBOM is missing Syft's source-root inventory entry."
    }

    $directoryEntries = @(
        [pscustomobject]@{
            Description = "source root"
            Entry = $SbomInventory.RootEntry
        }
    )
    foreach ($directoryPath in $PayloadInventory.DirectoryPaths) {
        $directoryEntries += [pscustomobject]@{
            Description = "directory $directoryPath"
            Entry = $SbomInventory.ByPath[$directoryPath]
        }
    }

    foreach ($directoryEntry in $directoryEntries) {
        $fileTypesProperty = $directoryEntry.Entry.PSObject.Properties['fileTypes']
        $checksumsProperty = $directoryEntry.Entry.PSObject.Properties['checksums']
        $fileTypes = @(if ($null -eq $fileTypesProperty) {
            @()
        }
        else {
            @($fileTypesProperty.Value)
        })
        $sha1Values = @(if ($null -eq $checksumsProperty) {
            @()
        }
        else {
            @($checksumsProperty.Value |
                Where-Object { [string]$_.algorithm -eq "SHA1" } |
                ForEach-Object { [string]$_.checksumValue })
        })
        if ($fileTypes.Count -ne 1 -or
            [string]$fileTypes[0] -cne "OTHER" -or
            $sha1Values.Count -ne 1 -or
            $sha1Values[0] -cne ('0' * 40)) {
            throw (
                "The release SBOM $($directoryEntry.Description) entry does not use " +
                "Syft's pinned directory sentinel contract.")
        }
    }
}
