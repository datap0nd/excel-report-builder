[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$InstallerPath,

    [string[]]$PublishedWorkerPaths = @()
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

function Get-DotNetFrameworkRelease {
    $view = if ([Environment]::Is64BitOperatingSystem) {
        [Microsoft.Win32.RegistryView]::Registry64
    }
    else {
        [Microsoft.Win32.RegistryView]::Registry32
    }

    $root = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
        [Microsoft.Win32.RegistryHive]::LocalMachine,
        $view)
    try {
        $key = $root.OpenSubKey("SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full")
        if ($null -eq $key) { return 0 }
        try {
            return [int]$key.GetValue("Release", 0)
        }
        finally {
            $key.Dispose()
        }
    }
    finally {
        $root.Dispose()
    }
}

function Get-PeMachine {
    param([Parameter(Mandatory = $true)][string]$Path)

    $stream = [IO.File]::OpenRead($Path)
    $reader = [IO.BinaryReader]::new($stream)
    try {
        if ($reader.ReadUInt16() -ne 0x5A4D) {
            throw "Installed worker is not a PE executable."
        }
        $stream.Position = 0x3C
        $peOffset = $reader.ReadInt32()
        $stream.Position = $peOffset
        if ($reader.ReadUInt32() -ne 0x00004550) {
            throw "Installed worker has an invalid PE header."
        }
        return $reader.ReadUInt16()
    }
    finally {
        $reader.Dispose()
        $stream.Dispose()
    }
}

function Invoke-WaitingProcess {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [string[]]$ArgumentList = @()
    )

    $process = $null
    try {
        $process = Start-Process `
            -FilePath $FilePath `
            -ArgumentList $ArgumentList `
            -Wait `
            -PassThru
        return [int]$process.ExitCode
    }
    finally {
        if ($null -ne $process) {
            $process.Dispose()
        }
    }
}

function Assert-MachineComActivationKeysAbsent {
    param(
        [Parameter(Mandatory = $true)][Microsoft.Win32.RegistryView]$View,
        [Parameter(Mandatory = $true)][string]$Context
    )

    $paths = @(
        "Software\Microsoft\Office\Excel\Addins\ExcelReportBuilder.AddIn",
        "Software\Classes\CLSID\{F953480C-A73C-4121-9E21-18676EC34CE8}",
        "Software\Classes\CLSID\{A3F4E10D-0DD1-420E-8B6F-E0A654BBEA16}",
        "Software\Classes\ExcelReportBuilder.AddIn",
        "Software\Classes\ExcelReportBuilder.TaskPaneHost"
    )
    $root = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
        [Microsoft.Win32.RegistryHive]::LocalMachine,
        $View)
    try {
        foreach ($path in $paths) {
            $key = $root.OpenSubKey($path)
            try {
                if ($null -ne $key) {
                    throw "$Context found machine registration $path in $View."
                }
            }
            finally {
                if ($null -ne $key) { $key.Dispose() }
            }
        }
    }
    finally {
        $root.Dispose()
    }
}

function Invoke-ComActivationSmoke {
    param(
        [Parameter(Mandatory = $true)][string]$ScriptPath,
        [Parameter(Mandatory = $true)][string]$AssemblyPath
    )

    @'
$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

foreach ($progId in @(
    "ExcelReportBuilder.AddIn",
    "ExcelReportBuilder.TaskPaneHost"
)) {
    $instance = $null
    try {
        $instance = New-Object -ComObject $progId
        if ($null -eq $instance) {
            throw "COM activation returned no instance for $progId."
        }
    }
    finally {
        if ($null -ne $instance -and $instance -is [IDisposable]) {
            $instance.Dispose()
        }
        if ($null -ne $instance -and [Runtime.InteropServices.Marshal]::IsComObject($instance)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($instance)
        }
    }
}
'@ | Set-Content $ScriptPath -Encoding ascii

    $activationHosts = if ([Environment]::Is64BitOperatingSystem) {
        @(
            [pscustomobject]@{
                PowerShell = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
                RegAsm = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
                View = [Microsoft.Win32.RegistryView]::Registry64
            },
            [pscustomobject]@{
                PowerShell = "$env:WINDIR\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
                RegAsm = "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
                View = [Microsoft.Win32.RegistryView]::Registry32
            }
        )
    }
    else {
        @(
            [pscustomobject]@{
                PowerShell = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
                RegAsm = "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
                View = [Microsoft.Win32.RegistryView]::Registry32
            }
        )
    }

    foreach ($activationHost in $activationHosts) {
        Assert-MachineComActivationKeysAbsent `
            -View $activationHost.View `
            -Context "COM activation preflight"
    }

    foreach ($activationHost in $activationHosts) {
        $primaryError = $null
        $cleanupFailures = @()
        try {
            & $activationHost.RegAsm $AssemblyPath /codebase /nologo
            if ($LASTEXITCODE -ne 0) {
                throw "Temporary COM registration failed in $($activationHost.RegAsm)."
            }

            & $activationHost.PowerShell -NoProfile -STA -File $ScriptPath
            if ($LASTEXITCODE -ne 0) {
                throw "Installed COM activation failed in $($activationHost.PowerShell)."
            }
        }
        catch {
            $primaryError = $_
        }
        finally {
            try {
                & $activationHost.RegAsm $AssemblyPath /unregister /nologo
                if ($LASTEXITCODE -ne 0) {
                    throw "Temporary COM registration cleanup failed in $($activationHost.RegAsm)."
                }
            }
            catch {
                $cleanupFailures += $_.Exception.Message
            }

            try {
                Assert-MachineComActivationKeysAbsent `
                    -View $activationHost.View `
                    -Context "Temporary COM cleanup"
            }
            catch {
                $cleanupFailures += $_.Exception.Message
            }
        }

        if ($null -ne $primaryError) {
            if ($cleanupFailures.Count -gt 0) {
                try {
                    Write-Warning (
                        "Temporary COM cleanup also failed after the primary activation failure: " +
                        ($cleanupFailures -join " "))
                }
                catch {
                    # Diagnostics must never replace the primary activation error.
                }
            }
            throw $primaryError
        }
        if ($cleanupFailures.Count -gt 0) {
            throw ($cleanupFailures -join " ")
        }
    }
}

function Read-ExactBytes {
    param(
        [Parameter(Mandatory = $true)][IO.Stream]$Stream,
        [Parameter(Mandatory = $true)][byte[]]$Buffer,
        [ValidateRange(1, 60000)][int]$TimeoutMilliseconds = 10000
    )

    $offset = 0
    while ($offset -lt $Buffer.Length) {
        $readTask = $Stream.ReadAsync($Buffer, $offset, $Buffer.Length - $offset)
        if (-not $readTask.Wait($TimeoutMilliseconds)) {
            throw "Timed out while reading the worker protocol response."
        }
        $read = $readTask.GetAwaiter().GetResult()
        if ($read -eq 0) {
            throw "The worker closed the pipe before completing its response."
        }
        $offset += $read
    }
}

function Invoke-WorkerHandshakeSmoke {
    param([Parameter(Mandatory = $true)][string]$WorkerPath)

    $pipeName = "excel-report-builder-installer-" + [Guid]::NewGuid().ToString("N")
    $secretBytes = [byte[]]::new(32)
    $random = [Security.Cryptography.RandomNumberGenerator]::Create()
    try {
        $random.GetBytes($secretBytes)
    }
    finally {
        $random.Dispose()
    }
    $handshakeSecret = [Convert]::ToBase64String($secretBytes)
    $startInfo = [Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = $WorkerPath
    $startInfo.Arguments = "--pipe `"$pipeName`""
    $startInfo.WorkingDirectory = [IO.Path]::GetDirectoryName($WorkerPath)
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.EnvironmentVariables[
        "EXCEL_REPORT_BUILDER_WORKER_HANDSHAKE_SECRET"] = $handshakeSecret

    $workerProcess = [Diagnostics.Process]::new()
    $workerProcess.StartInfo = $startInfo
    $workerStarted = $false
    $pipe = $null
    $handshakeCompleted = $false
    $workerPrimaryError = $null
    $workerDiagnostic = ""
    $workerCleanupFailures = @()
    try {
        if (-not $workerProcess.Start()) {
            throw "The installed worker could not be started."
        }
        $workerStarted = $true
        [void]$startInfo.EnvironmentVariables.Remove(
            "EXCEL_REPORT_BUILDER_WORKER_HANDSHAKE_SECRET")

        $pipe = [IO.Pipes.NamedPipeClientStream]::new(
            ".",
            $pipeName,
            [IO.Pipes.PipeDirection]::InOut,
            [IO.Pipes.PipeOptions]::Asynchronous)
        $pipe.Connect(10000)

        $correlationId = "installer-smoke"
        $nonceBytes = [byte[]]::new(32)
        $random = [Security.Cryptography.RandomNumberGenerator]::Create()
        try {
            $random.GetBytes($nonceBytes)
        }
        finally {
            $random.Dispose()
        }
        $clientNonce = [Convert]::ToBase64String($nonceBytes)
        $request = [ordered]@{
            protocolVersion = "1.1"
            messageType = "hello"
            correlationId = $correlationId
            payload = [ordered]@{
                clientName = "installer-smoke"
                supportedProtocolVersions = @("1.1")
                clientNonce = $clientNonce
            }
        }
        $json = $request | ConvertTo-Json -Depth 6 -Compress
        $frame = [Text.Encoding]::UTF8.GetBytes($json)
        $header = [BitConverter]::GetBytes([int]$frame.Length)
        $pipe.Write($header, 0, $header.Length)
        $pipe.Write($frame, 0, $frame.Length)
        $pipe.Flush()

        $responseHeader = [byte[]]::new(4)
        Read-ExactBytes -Stream $pipe -Buffer $responseHeader
        $responseLength = [BitConverter]::ToInt32($responseHeader, 0)
        if ($responseLength -le 0 -or $responseLength -gt (1024 * 1024)) {
            throw "The worker returned an invalid protocol frame length."
        }

        $responseFrame = [byte[]]::new($responseLength)
        Read-ExactBytes -Stream $pipe -Buffer $responseFrame
        $response = [Text.Encoding]::UTF8.GetString($responseFrame) | ConvertFrom-Json

        if ($response.protocolVersion -ne "1.1" -or
            $response.messageType -ne "helloAcknowledged" -or
            $response.correlationId -ne $correlationId) {
            throw "The installed worker returned an invalid handshake response."
        }
        if ($response.payload.protocolVersion -ne "1.1" -or
            -not $response.payload.currentUserOnlyPipe) {
            throw "The installed worker did not confirm a current-user-only pipe."
        }

        $proofInput = "excel-report-builder-worker-handshake-v1`n" +
            $pipeName + "`n" + $clientNonce + "`n1.1"
        $proofBytes = [Text.Encoding]::UTF8.GetBytes($proofInput)
        $hmac = [Security.Cryptography.HMACSHA256]::new($secretBytes)
        try {
            $expectedTag = $hmac.ComputeHash($proofBytes)
        }
        finally {
            $hmac.Dispose()
            [Array]::Clear($proofBytes, 0, $proofBytes.Length)
        }
        try {
            $actualTag = [Convert]::FromBase64String(
                [string]$response.payload.authenticationTag)
        }
        catch {
            throw "The installed worker returned an invalid launch proof."
        }
        $difference = $expectedTag.Length -bxor $actualTag.Length
        if ($expectedTag.Length -eq $actualTag.Length) {
            for ($index = 0; $index -lt $expectedTag.Length; $index++) {
                $difference = $difference -bor ($expectedTag[$index] -bxor $actualTag[$index])
            }
        }
        [Array]::Clear($expectedTag, 0, $expectedTag.Length)
        [Array]::Clear($actualTag, 0, $actualTag.Length)
        if ($difference -ne 0) {
            throw "The installed worker did not prove it was launched by this installer test."
        }
        $handshakeCompleted = $true
    }
    catch {
        $workerPrimaryError = $_
        if ($workerStarted) {
            try {
                if ($workerProcess.HasExited) {
                    $workerDiagnostic = $workerProcess.StandardError.ReadToEnd().Trim()
                }
            }
            catch {
                $workerCleanupFailures += (
                    "Worker diagnostic collection failed: " + $_.Exception.Message)
            }
        }
    }
    finally {
        try {
            [void]$startInfo.EnvironmentVariables.Remove(
                "EXCEL_REPORT_BUILDER_WORKER_HANDSHAKE_SECRET")
        }
        catch {
            $workerCleanupFailures += (
                "Worker cleanup could not remove the launch secret: " + $_.Exception.Message)
        }

        if ($null -ne $pipe) {
            try {
                $pipe.Dispose()
            }
            catch {
                $workerCleanupFailures += (
                    "Worker pipe cleanup failed: " + $_.Exception.Message)
            }
        }

        $workerExited = $false
        $workerExitKnown = $false
        $workerWasKilled = $false
        if ($workerStarted) {
            try {
                $workerExited = $workerProcess.HasExited
                $workerExitKnown = $true
            }
            catch {
                $workerCleanupFailures += (
                    "Worker exit-state check failed: " + $_.Exception.Message)
            }

            if (-not $workerExitKnown -or -not $workerExited) {
                try {
                    $workerExited = $workerProcess.WaitForExit(5000)
                    $workerExitKnown = $true
                }
                catch {
                    $workerCleanupFailures += (
                        "Worker exit wait failed: " + $_.Exception.Message)
                }

                if (-not $workerExited) {
                    if ($handshakeCompleted) {
                        $workerCleanupFailures += (
                            "The installed worker did not exit after its one authenticated connection.")
                    }
                    $workerWasKilled = $true
                    try {
                        $workerProcess.Kill()
                    }
                    catch {
                        $workerCleanupFailures += (
                            "Worker termination failed: " + $_.Exception.Message)
                    }
                    try {
                        $workerExited = $workerProcess.WaitForExit(5000)
                        $workerExitKnown = $true
                        if (-not $workerExited) {
                            $workerCleanupFailures += (
                                "The installed worker remained running after termination.")
                        }
                    }
                    catch {
                        $workerCleanupFailures += (
                            "Worker post-termination wait failed: " + $_.Exception.Message)
                    }
                }
            }

            if ($handshakeCompleted -and
                $workerExitKnown -and
                $workerExited -and
                -not $workerWasKilled) {
                try {
                    if ($workerProcess.ExitCode -ne 0) {
                        $workerCleanupFailures += (
                            "The installed worker exited with code $($workerProcess.ExitCode) " +
                            "after a valid handshake.")
                    }
                }
                catch {
                    $workerCleanupFailures += (
                        "Worker exit-code check failed: " + $_.Exception.Message)
                }
            }
        }

        try {
            $workerProcess.Dispose()
        }
        catch {
            $workerCleanupFailures += (
                "Worker process cleanup failed: " + $_.Exception.Message)
        }
        try {
            [Array]::Clear($secretBytes, 0, $secretBytes.Length)
            if ($null -ne (Get-Variable nonceBytes -ErrorAction SilentlyContinue)) {
                [Array]::Clear($nonceBytes, 0, $nonceBytes.Length)
            }
        }
        catch {
            $workerCleanupFailures += (
                "Worker secret cleanup failed: " + $_.Exception.Message)
        }
    }

    if ($null -ne $workerPrimaryError) {
        if ($workerDiagnostic) {
            try {
                Write-Warning "Worker diagnostic after handshake failure: $workerDiagnostic"
            }
            catch {
                # Diagnostics must never replace the primary handshake error.
            }
        }
        if ($workerCleanupFailures.Count -gt 0) {
            try {
                Write-Warning (
                    "Worker cleanup also failed after the primary handshake failure: " +
                    ($workerCleanupFailures -join " "))
            }
            catch {
                # Diagnostics must never replace the primary handshake error.
            }
        }
        throw $workerPrimaryError
    }
    if ($workerCleanupFailures.Count -gt 0) {
        throw ($workerCleanupFailures -join " ")
    }

    $rejectedPipeName = "excel-report-builder-installer-reject-" +
        [Guid]::NewGuid().ToString("N")
    $rejectedStartInfo = [Diagnostics.ProcessStartInfo]::new()
    $rejectedStartInfo.FileName = $WorkerPath
    $rejectedStartInfo.Arguments = "--pipe `"$rejectedPipeName`""
    $rejectedStartInfo.WorkingDirectory = [IO.Path]::GetDirectoryName($WorkerPath)
    $rejectedStartInfo.UseShellExecute = $false
    $rejectedStartInfo.CreateNoWindow = $true
    $rejectedStartInfo.RedirectStandardError = $true
    [void]$rejectedStartInfo.EnvironmentVariables.Remove(
        "EXCEL_REPORT_BUILDER_WORKER_HANDSHAKE_SECRET")
    $rejectedProcess = [Diagnostics.Process]::new()
    $rejectedProcess.StartInfo = $rejectedStartInfo
    $rejectedStarted = $false
    $rejectedPrimaryError = $null
    $rejectedCleanupFailures = @()
    try {
        if (-not $rejectedProcess.Start()) {
            throw "The installed worker could not be started for the fail-closed test."
        }
        $rejectedStarted = $true
        if (-not $rejectedProcess.WaitForExit(5000)) {
            throw "The worker did not reject a launch without authentication."
        }
        if ($rejectedProcess.ExitCode -eq 0) {
            throw "The worker accepted a launch without authentication."
        }
    }
    catch {
        $rejectedPrimaryError = $_
    }
    finally {
        if ($rejectedStarted) {
            $rejectedExited = $false
            $rejectedExitKnown = $false
            try {
                $rejectedExited = $rejectedProcess.HasExited
                $rejectedExitKnown = $true
            }
            catch {
                $rejectedCleanupFailures += (
                    "Rejected-worker exit-state check failed: " + $_.Exception.Message)
            }

            if (-not $rejectedExitKnown -or -not $rejectedExited) {
                try {
                    $rejectedProcess.Kill()
                }
                catch {
                    $rejectedCleanupFailures += (
                        "Rejected-worker termination failed: " + $_.Exception.Message)
                }
                try {
                    $rejectedExited = $rejectedProcess.WaitForExit(5000)
                    if (-not $rejectedExited) {
                        $rejectedCleanupFailures += (
                            "The rejected worker remained running after termination.")
                    }
                }
                catch {
                    $rejectedCleanupFailures += (
                        "Rejected-worker exit wait failed: " + $_.Exception.Message)
                }
            }
        }

        try {
            $rejectedProcess.Dispose()
        }
        catch {
            $rejectedCleanupFailures += (
                "Rejected-worker process cleanup failed: " + $_.Exception.Message)
        }
    }

    if ($null -ne $rejectedPrimaryError) {
        if ($rejectedCleanupFailures.Count -gt 0) {
            try {
                Write-Warning (
                    "Rejected-worker cleanup also failed after the primary fail-closed error: " +
                    ($rejectedCleanupFailures -join " "))
            }
            catch {
                # Diagnostics must never replace the primary fail-closed error.
            }
        }
        throw $rejectedPrimaryError
    }
    if ($rejectedCleanupFailures.Count -gt 0) {
        throw ($rejectedCleanupFailures -join " ")
    }
}

function Assert-RegistrationRemoved {
    param([Parameter(Mandatory = $true)][Microsoft.Win32.RegistryView]$View)

    $paths = @(
        "Software\Microsoft\Office\Excel\Addins\ExcelReportBuilder.AddIn",
        "Software\Classes\CLSID\{F953480C-A73C-4121-9E21-18676EC34CE8}",
        "Software\Classes\CLSID\{A3F4E10D-0DD1-420E-8B6F-E0A654BBEA16}",
        "Software\Classes\ExcelReportBuilder.AddIn",
        "Software\Classes\ExcelReportBuilder.TaskPaneHost"
    )

    $root = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
        [Microsoft.Win32.RegistryHive]::CurrentUser,
        $View)
    try {
        foreach ($path in $paths) {
            $key = $root.OpenSubKey($path)
            try {
                if ($null -ne $key) {
                    throw "Uninstall left registration $path in $View."
                }
            }
            finally {
                if ($null -ne $key) { $key.Dispose() }
            }
        }
    }
    finally {
        $root.Dispose()
    }
}

function Remove-PerUserComActivationKeys {
    param([Parameter(Mandatory = $true)][Microsoft.Win32.RegistryView[]]$Views)

    $paths = @(
        "Software\Classes\CLSID\{F953480C-A73C-4121-9E21-18676EC34CE8}",
        "Software\Classes\CLSID\{A3F4E10D-0DD1-420E-8B6F-E0A654BBEA16}",
        "Software\Classes\ExcelReportBuilder.AddIn",
        "Software\Classes\ExcelReportBuilder.TaskPaneHost"
    )

    foreach ($view in $Views) {
        $root = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
            [Microsoft.Win32.RegistryHive]::CurrentUser,
            $view)
        try {
            foreach ($path in $paths) {
                $root.DeleteSubKeyTree($path, $false)
            }
        }
        finally {
            $root.Dispose()
        }
    }
}

function Get-RequiredRegistryValue {
    param(
        [Parameter(Mandatory = $true)][Microsoft.Win32.RegistryKey]$Root,
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$ValueName = "",
        [Parameter(Mandatory = $true)]
        [Microsoft.Win32.RegistryValueKind]$ExpectedKind,
        [Parameter(Mandatory = $true)][string]$Label
    )

    $key = $Root.OpenSubKey($Path)
    if ($null -eq $key) {
        throw "$Label key is missing."
    }
    try {
        $value = $key.GetValue(
            $ValueName,
            $null,
            [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
        if ($null -eq $value) {
            $displayName = if ($ValueName) { $ValueName } else { "(Default)" }
            throw "$Label value $displayName is missing."
        }
        $actualKind = $key.GetValueKind($ValueName)
        if ($actualKind -ne $ExpectedKind) {
            $displayName = if ($ValueName) { $ValueName } else { "(Default)" }
            throw (
                "$Label value $displayName has registry kind $actualKind; " +
                "expected $ExpectedKind.")
        }
        return $value
    }
    finally {
        $key.Dispose()
    }
}

function Assert-RegistryValueEquals {
    param(
        [Parameter(Mandatory = $true)][Microsoft.Win32.RegistryKey]$Root,
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$ValueName = "",
        [Parameter(Mandatory = $true)]
        [Microsoft.Win32.RegistryValueKind]$ExpectedKind,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [object]$ExpectedValue,
        [Parameter(Mandatory = $true)][string]$Label
    )

    $actualValue = Get-RequiredRegistryValue `
        -Root $Root `
        -Path $Path `
        -ValueName $ValueName `
        -ExpectedKind $ExpectedKind `
        -Label $Label
    $matches = if ($ExpectedKind -in @(
            [Microsoft.Win32.RegistryValueKind]::String,
            [Microsoft.Win32.RegistryValueKind]::ExpandString)) {
        [string]::Equals(
            [string]$actualValue,
            [string]$ExpectedValue,
            [StringComparison]::Ordinal)
    }
    else {
        [object]::Equals($actualValue, $ExpectedValue)
    }
    if (-not $matches) {
        $displayName = if ($ValueName) { $ValueName } else { "(Default)" }
        throw "$Label value $displayName is wrong."
    }
}

function Assert-CodeBaseContract {
    param(
        [Parameter(Mandatory = $true)][string]$RegisteredCodeBase,
        [Parameter(Mandatory = $true)][string]$AssemblyPath,
        [Parameter(Mandatory = $true)][Microsoft.Win32.RegistryView]$View,
        [Parameter(Mandatory = $true)][string]$Label
    )

    $codeBaseUri = [Uri]::new($RegisteredCodeBase, [UriKind]::Absolute)
    if (-not $codeBaseUri.IsFile -or
        -not [string]::IsNullOrEmpty($codeBaseUri.Query) -or
        -not [string]::IsNullOrEmpty($codeBaseUri.Fragment) -or
        -not [string]::Equals(
            [IO.Path]::GetFullPath($codeBaseUri.LocalPath),
            [IO.Path]::GetFullPath($AssemblyPath),
            [StringComparison]::OrdinalIgnoreCase)) {
        throw "$Label CodeBase is not a safe URL for the installed assembly in $View."
    }
    if ($RegisteredCodeBase.Contains("#") -or
        -not $RegisteredCodeBase.Contains("%23") -or
        -not $RegisteredCodeBase.Contains("%25")) {
        throw "$Label CodeBase did not escape reserved path characters in $View."
    }
}

function Assert-InstalledPerUserRegistration {
    param(
        [Parameter(Mandatory = $true)][Microsoft.Win32.RegistryView[]]$Views,
        [Parameter(Mandatory = $true)][string]$AssemblyPath,
        [Parameter(Mandatory = $true)][string]$AppClsid,
        [Parameter(Mandatory = $true)][string]$PaneClsid
    )

    $appName = "Excel Report Builder"
    $assemblyVersion = "0.1.0.0"
    $assemblyName = (
        "ExcelReportBuilder.AddIn, Version=$assemblyVersion, " +
        "Culture=neutral, PublicKeyToken=null")
    $runtimeVersion = "v4.0.30319"
    $managedCategory = "{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"
    $controlCategory = "{40FC6ED4-2438-11CF-A3DB-080036F12502}"
    $stringKind = [Microsoft.Win32.RegistryValueKind]::String
    $dwordKind = [Microsoft.Win32.RegistryValueKind]::DWord

    $contracts = @(
        [pscustomobject]@{
            ProgId = "ExcelReportBuilder.AddIn"
            Clsid = $AppClsid
            ClassName = "ExcelReportBuilder.AddIn.Com.ExcelReportBuilderAddIn"
            DisplayName = $appName
            IsControl = $false
            Label = "Excel add-in COM"
        },
        [pscustomobject]@{
            ProgId = "ExcelReportBuilder.TaskPaneHost"
            Clsid = $PaneClsid
            ClassName = "ExcelReportBuilder.AddIn.Hosting.TaskPaneHost"
            DisplayName = "$appName task pane"
            IsControl = $true
            Label = "Task-pane COM"
        }
    )

    foreach ($view in $Views) {
        $root = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
            [Microsoft.Win32.RegistryHive]::CurrentUser,
            $view)
        try {
            $addinPath = "Software\Microsoft\Office\Excel\Addins\ExcelReportBuilder.AddIn"
            foreach ($addinValue in @(
                [pscustomobject]@{
                    Name = "FriendlyName"
                    Value = "PivotTable+"
                    Kind = $stringKind
                },
                [pscustomobject]@{
                    Name = "Description"
                    Value = "Enhance a selected native Excel PivotTable."
                    Kind = $stringKind
                },
                [pscustomobject]@{
                    Name = "LoadBehavior"
                    Value = [int]3
                    Kind = $dwordKind
                },
                [pscustomobject]@{
                    Name = "CommandLineSafe"
                    Value = [int]0
                    Kind = $dwordKind
                }
            )) {
                Assert-RegistryValueEquals `
                    -Root $root `
                    -Path $addinPath `
                    -ValueName $addinValue.Name `
                    -ExpectedKind $addinValue.Kind `
                    -ExpectedValue $addinValue.Value `
                    -Label "Excel add-in registration in $view"
            }

            foreach ($contract in $contracts) {
                $clsidPath = "Software\Classes\CLSID\$($contract.Clsid)"
                $inprocPath = "$clsidPath\InprocServer32"
                $versionedPath = "$inprocPath\$assemblyVersion"
                $progIdPath = "Software\Classes\$($contract.ProgId)"

                Assert-RegistryValueEquals `
                    -Root $root `
                    -Path $clsidPath `
                    -ExpectedKind $stringKind `
                    -ExpectedValue $contract.DisplayName `
                    -Label "$($contract.Label) registration in $view"

                foreach ($inprocValue in @(
                    [pscustomobject]@{ Name = ""; Value = "mscoree.dll" },
                    [pscustomobject]@{ Name = "ThreadingModel"; Value = "Both" },
                    [pscustomobject]@{ Name = "Class"; Value = $contract.ClassName },
                    [pscustomobject]@{ Name = "Assembly"; Value = $assemblyName },
                    [pscustomobject]@{ Name = "RuntimeVersion"; Value = $runtimeVersion }
                )) {
                    Assert-RegistryValueEquals `
                        -Root $root `
                        -Path $inprocPath `
                        -ValueName $inprocValue.Name `
                        -ExpectedKind $stringKind `
                        -ExpectedValue $inprocValue.Value `
                        -Label "$($contract.Label) registration in $view"
                }

                $registeredCodeBase = [string](Get-RequiredRegistryValue `
                    -Root $root `
                    -Path $inprocPath `
                    -ValueName "CodeBase" `
                    -ExpectedKind $stringKind `
                    -Label "$($contract.Label) registration in $view")
                Assert-CodeBaseContract `
                    -RegisteredCodeBase $registeredCodeBase `
                    -AssemblyPath $AssemblyPath `
                    -View $view `
                    -Label $contract.Label

                foreach ($versionedValue in @(
                    [pscustomobject]@{ Name = "Class"; Value = $contract.ClassName },
                    [pscustomobject]@{ Name = "Assembly"; Value = $assemblyName },
                    [pscustomobject]@{ Name = "RuntimeVersion"; Value = $runtimeVersion }
                )) {
                    Assert-RegistryValueEquals `
                        -Root $root `
                        -Path $versionedPath `
                        -ValueName $versionedValue.Name `
                        -ExpectedKind $stringKind `
                        -ExpectedValue $versionedValue.Value `
                        -Label "$($contract.Label) versioned registration in $view"
                }
                $versionedCodeBase = [string](Get-RequiredRegistryValue `
                    -Root $root `
                    -Path $versionedPath `
                    -ValueName "CodeBase" `
                    -ExpectedKind $stringKind `
                    -Label "$($contract.Label) versioned registration in $view")
                Assert-CodeBaseContract `
                    -RegisteredCodeBase $versionedCodeBase `
                    -AssemblyPath $AssemblyPath `
                    -View $view `
                    -Label "$($contract.Label) versioned"

                Assert-RegistryValueEquals `
                    -Root $root `
                    -Path "$clsidPath\ProgId" `
                    -ExpectedKind $stringKind `
                    -ExpectedValue $contract.ProgId `
                    -Label "$($contract.Label) ProgId registration in $view"
                Assert-RegistryValueEquals `
                    -Root $root `
                    -Path $progIdPath `
                    -ExpectedKind $stringKind `
                    -ExpectedValue $contract.DisplayName `
                    -Label "$($contract.Label) ProgId registration in $view"
                Assert-RegistryValueEquals `
                    -Root $root `
                    -Path "$progIdPath\CLSID" `
                    -ExpectedKind $stringKind `
                    -ExpectedValue $contract.Clsid `
                    -Label "$($contract.Label) CLSID registration in $view"
                Assert-RegistryValueEquals `
                    -Root $root `
                    -Path "$clsidPath\Implemented Categories\$managedCategory" `
                    -ExpectedKind $stringKind `
                    -ExpectedValue "" `
                    -Label "$($contract.Label) managed category in $view"

                if ($contract.IsControl) {
                    foreach ($controlPath in @(
                        "$clsidPath\Implemented Categories\$controlCategory",
                        "$clsidPath\Programmable",
                        "$clsidPath\Control"
                    )) {
                        Assert-RegistryValueEquals `
                            -Root $root `
                            -Path $controlPath `
                            -ExpectedKind $stringKind `
                            -ExpectedValue "" `
                            -Label "$($contract.Label) control registration in $view"
                    }
                }
            }
        }
        finally {
            $root.Dispose()
        }
    }
}

function Invoke-BestEffortInstalledCleanup {
    param([Parameter(Mandatory = $true)][string]$UninstallerPath)

    if (-not (Test-Path -LiteralPath $UninstallerPath -PathType Leaf)) {
        return
    }

    try {
        $exitCode = Invoke-WaitingProcess -FilePath $UninstallerPath -ArgumentList @(
            "/VERYSILENT",
            "/SUPPRESSMSGBOXES",
            "/NORESTART"
        )
        if ($exitCode -ne 0) {
            try {
                Write-Warning "Best-effort uninstall cleanup exited with code $exitCode."
            }
            catch {
                # Cleanup diagnostics are non-fatal by design.
            }
        }
    }
    catch {
        $cleanupMessage = $_.Exception.Message
        try {
            Write-Warning "Best-effort uninstall cleanup failed: $cleanupMessage"
        }
        catch {
            # Cleanup diagnostics are non-fatal by design.
        }
    }
}

$frameworkRelease = Get-DotNetFrameworkRelease
if ($frameworkRelease -lt 528040) {
    throw ".NET Framework 4.8 or newer is required for the installation smoke test."
}

$publishedWorkerMachines = @()
foreach ($publishedWorkerPath in $PublishedWorkerPaths) {
    $resolvedPublishedWorker = (Resolve-Path $publishedWorkerPath).Path
    $publishedWorkerMachine = Get-PeMachine -Path $resolvedPublishedWorker
    if ($publishedWorkerMachine -notin @(0x014C, 0x8664)) {
        throw ("Published worker has unsupported machine type 0x{0:X4}: {1}" -f
            $publishedWorkerMachine,
            $publishedWorkerPath)
    }
    if ($publishedWorkerMachines -contains $publishedWorkerMachine) {
        throw ("Published worker coverage contains duplicate machine type 0x{0:X4}." -f
            $publishedWorkerMachine)
    }

    $publishedWorkerMachines += $publishedWorkerMachine
    Invoke-WorkerHandshakeSmoke -WorkerPath $resolvedPublishedWorker
}

if ($PublishedWorkerPaths.Count -gt 0) {
    foreach ($requiredMachine in @(0x014C, 0x8664)) {
        if ($publishedWorkerMachines -notcontains $requiredMachine) {
            throw ("Published worker smoke coverage is missing machine type 0x{0:X4}." -f
                $requiredMachine)
        }
    }
}

$resolvedInstaller = (Resolve-Path $InstallerPath).Path
$temporaryRoot = if ($env:RUNNER_TEMP) { $env:RUNNER_TEMP } else { [IO.Path]::GetTempPath() }
$installDirectory = Join-Path $temporaryRoot (
    "Excel Report Builder # percent% " + [Guid]::NewGuid().ToString("N"))
$installerArguments = @(
    "/VERYSILENT",
    "/SUPPRESSMSGBOXES",
    "/NORESTART",
    "/DIR=`"$installDirectory`""
)
$uninstaller = Join-Path $installDirectory "unins000.exe"

try {
    $installExitCode = Invoke-WaitingProcess `
        -FilePath $resolvedInstaller `
        -ArgumentList $installerArguments
    if ($installExitCode -ne 0) {
        throw "Installer exited with code $installExitCode."
    }

    $assemblyPath = Join-Path $installDirectory "ExcelReportBuilder.AddIn.dll"
    $workerPath = Join-Path $installDirectory "worker\ExcelReportBuilder.Worker.exe"
    if (-not (Test-Path -LiteralPath $assemblyPath -PathType Leaf)) {
        throw "Installed add-in assembly was not found."
    }
    if (-not (Test-Path -LiteralPath $workerPath -PathType Leaf)) {
        throw "Installed worker executable was not found."
    }

    $expectedMachine = if ([Environment]::Is64BitOperatingSystem) { 0x8664 } else { 0x014C }
    $actualMachine = Get-PeMachine -Path $workerPath
    if ($actualMachine -ne $expectedMachine) {
        throw ("Installer selected worker machine 0x{0:X4}; expected 0x{1:X4}." -f
            $actualMachine, $expectedMachine)
    }

    $appClsid = "{F953480C-A73C-4121-9E21-18676EC34CE8}"
    $paneClsid = "{A3F4E10D-0DD1-420E-8B6F-E0A654BBEA16}"
    $views = @([Microsoft.Win32.RegistryView]::Registry32)
    if ([Environment]::Is64BitOperatingSystem) {
        $views = @(
            [Microsoft.Win32.RegistryView]::Registry64,
            [Microsoft.Win32.RegistryView]::Registry32
        )
    }

    Assert-InstalledPerUserRegistration `
        -Views $views `
        -AssemblyPath $assemblyPath `
        -AppClsid $appClsid `
        -PaneClsid $paneClsid

    $activationSmoke = Join-Path $temporaryRoot (
        "excel-report-builder-com-activation-" + [Guid]::NewGuid().ToString("N") + ".ps1")
    Remove-PerUserComActivationKeys -Views $views
    $activationError = $null
    $activationScriptCleanupError = $null
    try {
        # GitHub-hosted Windows runners are elevated non-interactive services. In
        # that context Windows ignores per-user COM classes. The registry checks
        # above verify the installed HKCU contract; temporary machine registration
        # exercises both real managed class factories and is removed immediately.
        Invoke-ComActivationSmoke `
            -ScriptPath $activationSmoke `
            -AssemblyPath $assemblyPath
    }
    catch {
        $activationError = $_
    }
    finally {
        try {
            if (Test-Path -LiteralPath $activationSmoke) {
                Remove-Item -LiteralPath $activationSmoke -Force
            }
        }
        catch {
            $activationScriptCleanupError = $_
        }
    }
    if ($null -ne $activationError) {
        if ($null -ne $activationScriptCleanupError) {
            try {
                Write-Warning (
                    "Temporary activation-script cleanup also failed after the primary activation failure: " +
                    $activationScriptCleanupError.Exception.Message)
            }
            catch {
                # Diagnostics must never replace the primary activation error.
            }
        }
        throw $activationError
    }
    if ($null -ne $activationScriptCleanupError) {
        throw $activationScriptCleanupError
    }

    $repairExitCode = Invoke-WaitingProcess `
        -FilePath $resolvedInstaller `
        -ArgumentList $installerArguments
    if ($repairExitCode -ne 0) {
        throw "Installer repair exited with code $repairExitCode."
    }
    Assert-InstalledPerUserRegistration `
        -Views $views `
        -AssemblyPath $assemblyPath `
        -AppClsid $appClsid `
        -PaneClsid $paneClsid

    Invoke-WorkerHandshakeSmoke -WorkerPath $workerPath

    if (-not (Test-Path -LiteralPath $uninstaller -PathType Leaf)) {
        throw "Installed uninstaller was not found."
    }
    $uninstallExitCode = Invoke-WaitingProcess -FilePath $uninstaller -ArgumentList @(
        "/VERYSILENT",
        "/SUPPRESSMSGBOXES",
        "/NORESTART"
    )
    if ($uninstallExitCode -ne 0) {
        throw "Uninstaller exited with code $uninstallExitCode."
    }

    $cleanupDeadline = [DateTime]::UtcNow.AddSeconds(15)
    while ((Test-Path -LiteralPath $installDirectory) -and
        [DateTime]::UtcNow -lt $cleanupDeadline) {
        Start-Sleep -Milliseconds 250
    }
    if (Test-Path -LiteralPath $installDirectory) {
        throw "Uninstall left files in the installation directory."
    }
    foreach ($view in $views) {
        Assert-RegistrationRemoved -View $view
    }

    Write-Host "Per-user x86/x64 registration, real COM activation, worker handshake, and uninstall smoke tests passed."
}
catch {
    $primaryError = $_
    try {
        Invoke-BestEffortInstalledCleanup -UninstallerPath $uninstaller
    }
    catch {
        # Best-effort cleanup must never replace the primary smoke-test error.
    }
    throw $primaryError
}
