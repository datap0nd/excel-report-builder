$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$workingFiles = @(git -c core.quotepath=false ls-files --cached --others --exclude-standard)
if ($LASTEXITCODE -ne 0) {
    throw "Could not enumerate tracked and nonignored untracked files."
}
$stagedFiles = @(git -c core.quotepath=false diff --cached --name-only --diff-filter=ACMR)
if ($LASTEXITCODE -ne 0) {
    throw "Could not enumerate staged files."
}

$reviewFiles = @($workingFiles | Where-Object { $_ } | Sort-Object -Unique)
$stagedFiles = @($stagedFiles | Where-Object { $_ } | Sort-Object -Unique)

$forbiddenExtensions = @(
    ".xls", ".xlsx", ".xlsm", ".xlsb", ".xlam", ".ods",
    ".csv", ".tsv", ".parquet", ".feather", ".arrow",
    ".db", ".sqlite", ".sqlite3",
    ".doc", ".docx", ".ppt", ".pptx", ".pdf",
    ".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp",
    ".tif", ".tiff", ".heic",
    ".log", ".transcript", ".prompt", ".jsonl", ".ndjson",
    ".har", ".pcap", ".pcapng", ".dmp", ".etl", ".sarif", ".trx",
    ".pfx", ".p12", ".pem", ".key"
)
$forbiddenLeafNames = @(
    ".env",
    "appsettings.local.json",
    "settings.local.json",
    "credentials.json",
    "credential.json",
    "secrets.json",
    "secret.json"
)
$forbiddenPathSegments = @(
    "workbooks",
    "exports",
    "screenshots",
    "transcripts",
    "private-data"
)
$maximumReviewBytes = 5MB

function Test-PublicFileName {
    param([Parameter(Mandatory = $true)][string]$File)

    $extension = [IO.Path]::GetExtension($File).ToLowerInvariant()
    if ($forbiddenExtensions -contains $extension) {
        throw "Private-artifact extension is not allowed: $File"
    }

    $leafName = [IO.Path]::GetFileName($File).ToLowerInvariant()
    if ($forbiddenLeafNames -contains $leafName -or
        $leafName -match '^(credentials?|secrets?)\.[^.]+$') {
        throw "Local credential or secret file is not allowed: $File"
    }

    $segments = $File.Replace('\', '/').Split('/')
    foreach ($segment in $segments) {
        if ($forbiddenPathSegments -contains $segment.ToLowerInvariant()) {
            throw "Private-artifact directory is not allowed: $File"
        }
    }
}

$patterns = [ordered]@{
    "GitHub token" = ('(?i)gh' + '[pousr]_[A-Za-z0-9_]{20,}')
    "OpenAI-style token" = ('(?i)\bsk' + '-[A-Za-z0-9_-]{20,}\b')
    "AWS access key" = ('\bAK' + 'IA[0-9A-Z]{16}\b')
    "Google API key" = ('\bAI' + 'za[0-9A-Za-z_-]{30,}\b')
    "private key" = ('-----BEGIN [A-Z0-9 ]*PRIVATE' + ' KEY-----')
    "quoted credential" = ('(?i)(api[_-]?key|client[_-]?secret|access[_-]?token|pass' + 'word)\s*[:=]\s*["''][^"'']{8,}["'']')
    "authorization token" = ('(?i)authoriz' + 'ation\s*:\s*(bearer|basic)\s+[A-Za-z0-9+/_.=-]{20,}')
    "user-local absolute path" = ('(?i)([A-Z]:\\' + 'Users\\|/' + 'Users/|/' + 'home/)[^\s"'']+')
    "network share path" = ('(?i)(?<!\\)\\' + '\\[A-Za-z0-9][A-Za-z0-9.-]{0,62}\\[A-Za-z0-9$][A-Za-z0-9$._ -]{0,127}(?:\\[A-Za-z0-9$._ -]+)*')
}

function Test-PublicText {
    param(
        [Parameter(Mandatory = $true)][string]$Content,
        [Parameter(Mandatory = $true)][string]$Label
    )

    foreach ($entry in $patterns.GetEnumerator()) {
        if ($Content -match $entry.Value) {
            throw "$($entry.Key) pattern matched in $Label."
        }
    }
}

$binaryExtensions = @(
    ".ico", ".exe", ".dll", ".zip", ".nupkg", ".snupkg"
)

foreach ($file in $reviewFiles) {
    Test-PublicFileName -File $file
    if (-not (Test-Path -LiteralPath $file -PathType Leaf)) {
        continue
    }

    $item = Get-Item -LiteralPath $file -Force
    if ($item.Length -gt $maximumReviewBytes) {
        throw "Nonignored file exceeds the 5 MB public-review limit: $file"
    }

    $extension = [IO.Path]::GetExtension($file).ToLowerInvariant()
    if ($binaryExtensions -contains $extension) {
        continue
    }

    $content = Get-Content -LiteralPath $file -Raw
    Test-PublicText -Content $content -Label $file
}

foreach ($file in $stagedFiles) {
    Test-PublicFileName -File $file
    $extension = [IO.Path]::GetExtension($file).ToLowerInvariant()
    if ($binaryExtensions -contains $extension) {
        continue
    }

    $stagedContent = (git show ":$file") -join "`n"
    if ($LASTEXITCODE -ne 0) {
        throw "Could not inspect staged content for $file."
    }
    Test-PublicText -Content $stagedContent -Label "$file (staged)"
}

Write-Host (
    "Public-safety checks passed for {0} tracked/nonignored files and {1} staged files." -f
    $reviewFiles.Count,
    $stagedFiles.Count)
