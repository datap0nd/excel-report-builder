[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$InstallerPath
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

function Invoke-ComActivationSmoke {
    param([Parameter(Mandatory = $true)][string]$ScriptPath)

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

    $powershellHosts = @("$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe")
    if ([Environment]::Is64BitOperatingSystem) {
        $powershellHosts += "$env:WINDIR\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
    }

    foreach ($powershellHost in $powershellHosts) {
        & $powershellHost -NoProfile -STA -File $ScriptPath
        if ($LASTEXITCODE -ne 0) {
            throw "Installed COM activation failed in $powershellHost."
        }
    }
}

function Read-ExactBytes {
    param(
        [Parameter(Mandatory = $true)][IO.Stream]$Stream,
        [Parameter(Mandatory = $true)][byte[]]$Buffer
    )

    $offset = 0
    while ($offset -lt $Buffer.Length) {
        $read = $Stream.Read($Buffer, $offset, $Buffer.Length - $offset)
        if ($read -eq 0) {
            throw "The worker closed the pipe before completing its response."
        }
        $offset += $read
    }
}

function Invoke-WorkerHandshakeSmoke {
    param([Parameter(Mandatory = $true)][string]$WorkerPath)

    $pipeName = "excel-report-builder-installer-" + [Guid]::NewGuid().ToString("N")
    $startInfo = [Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = $WorkerPath
    $startInfo.Arguments = "--pipe `"$pipeName`""
    $startInfo.WorkingDirectory = [IO.Path]::GetDirectoryName($WorkerPath)
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardError = $true

    $workerProcess = [Diagnostics.Process]::new()
    $workerProcess.StartInfo = $startInfo
    if (-not $workerProcess.Start()) {
        throw "The installed worker could not be started."
    }

    $pipe = $null
    try {
        $pipe = [IO.Pipes.NamedPipeClientStream]::new(
            ".",
            $pipeName,
            [IO.Pipes.PipeDirection]::InOut,
            [IO.Pipes.PipeOptions]::None)
        $pipe.Connect(10000)
        $pipe.ReadTimeout = 10000
        $pipe.WriteTimeout = 10000

        $correlationId = "installer-smoke"
        $request = [ordered]@{
            protocolVersion = "1.0"
            messageType = "hello"
            correlationId = $correlationId
            payload = [ordered]@{
                clientName = "installer-smoke"
                supportedProtocolVersions = @("1.0")
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

        if ($response.protocolVersion -ne "1.0" -or
            $response.messageType -ne "helloAcknowledged" -or
            $response.correlationId -ne $correlationId) {
            throw "The installed worker returned an invalid handshake response."
        }
        if ($response.payload.protocolVersion -ne "1.0" -or
            -not $response.payload.currentUserOnlyPipe) {
            throw "The installed worker did not confirm a current-user-only pipe."
        }
    }
    catch {
        $diagnostic = if ($workerProcess.HasExited) {
            $workerProcess.StandardError.ReadToEnd().Trim()
        }
        else {
            ""
        }
        if ($diagnostic) {
            throw "Worker handshake failed: $diagnostic"
        }
        throw
    }
    finally {
        if ($null -ne $pipe) { $pipe.Dispose() }
        if (-not $workerProcess.HasExited) {
            $workerProcess.Kill()
            $workerProcess.WaitForExit(5000) | Out-Null
        }
        $workerProcess.Dispose()
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

$frameworkRelease = Get-DotNetFrameworkRelease
if ($frameworkRelease -lt 528040) {
    throw ".NET Framework 4.8 or newer is required for the installation smoke test."
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

$installProcess = Start-Process $resolvedInstaller -ArgumentList $installerArguments -Wait -PassThru
if ($installProcess.ExitCode -ne 0) {
    throw "Installer exited with code $($installProcess.ExitCode)."
}

$assemblyPath = Join-Path $installDirectory "ExcelReportBuilder.AddIn.dll"
$workerPath = Join-Path $installDirectory "worker\ExcelReportBuilder.Worker.exe"
if (-not (Test-Path $assemblyPath -PathType Leaf)) {
    throw "Installed add-in assembly was not found."
}
if (-not (Test-Path $workerPath -PathType Leaf)) {
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

foreach ($view in $views) {
    $addin = $null
    $server = $null
    $pane = $null
    $control = $null
    $versionedServer = $null
    $versionedPane = $null
    $root = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
        [Microsoft.Win32.RegistryHive]::CurrentUser,
        $view)
    try {
        $addin = $root.OpenSubKey("Software\Microsoft\Office\Excel\Addins\ExcelReportBuilder.AddIn")
        if ($null -eq $addin -or $addin.GetValue("LoadBehavior") -ne 3) {
            throw "Excel add-in registration is missing in $view."
        }

        $server = $root.OpenSubKey("Software\Classes\CLSID\$appClsid\InprocServer32")
        $pane = $root.OpenSubKey("Software\Classes\CLSID\$paneClsid\InprocServer32")
        if ($null -eq $server -or $null -eq $pane) {
            throw "Managed COM registration is missing in $view."
        }
        if ([string]$server.GetValue("Class") -ne "ExcelReportBuilder.AddIn.Com.ExcelReportBuilderAddIn") {
            throw "Excel add-in COM class is wrong in $view."
        }
        if ([string]$pane.GetValue("Class") -ne "ExcelReportBuilder.AddIn.Hosting.TaskPaneHost") {
            throw "Task-pane COM class is wrong in $view."
        }

        foreach ($registeredCodeBase in @(
            [string]$server.GetValue("CodeBase"),
            [string]$pane.GetValue("CodeBase")
        )) {
            $codeBaseUri = [Uri]::new($registeredCodeBase, [UriKind]::Absolute)
            if (-not $codeBaseUri.IsFile -or
                -not [string]::Equals(
                    [IO.Path]::GetFullPath($codeBaseUri.LocalPath),
                    [IO.Path]::GetFullPath($assemblyPath),
                    [StringComparison]::OrdinalIgnoreCase)) {
                throw "Managed COM CodeBase is not a safe URL for the installed assembly in $view."
            }
            if ($registeredCodeBase.Contains("#") -or
                -not $registeredCodeBase.Contains("%23") -or
                -not $registeredCodeBase.Contains("%25")) {
                throw "Managed COM CodeBase did not escape reserved path characters in $view."
            }
        }

        $versionedServer = $server.OpenSubKey("0.1.0.0")
        $versionedPane = $pane.OpenSubKey("0.1.0.0")
        if ($null -eq $versionedServer -or $null -eq $versionedPane) {
            throw "Versioned managed COM registration is missing in $view."
        }

        $control = $root.OpenSubKey("Software\Classes\CLSID\$paneClsid\Control")
        if ($null -eq $control) {
            throw "Task-pane ActiveX Control registration is missing in $view."
        }
    }
    finally {
        if ($null -ne $addin) { $addin.Dispose() }
        if ($null -ne $versionedServer) { $versionedServer.Dispose() }
        if ($null -ne $versionedPane) { $versionedPane.Dispose() }
        if ($null -ne $server) { $server.Dispose() }
        if ($null -ne $pane) { $pane.Dispose() }
        if ($null -ne $control) { $control.Dispose() }
        $root.Dispose()
    }
}

$activationSmoke = Join-Path $temporaryRoot (
    "excel-report-builder-com-activation-" + [Guid]::NewGuid().ToString("N") + ".ps1")
try {
    Invoke-ComActivationSmoke -ScriptPath $activationSmoke
}
finally {
    if (Test-Path -LiteralPath $activationSmoke) {
        Remove-Item -LiteralPath $activationSmoke -Force
    }
}

Invoke-WorkerHandshakeSmoke -WorkerPath $workerPath

$uninstaller = Join-Path $installDirectory "unins000.exe"
if (-not (Test-Path $uninstaller -PathType Leaf)) {
    throw "Installed uninstaller was not found."
}
$uninstallProcess = Start-Process $uninstaller -ArgumentList @(
    "/VERYSILENT",
    "/SUPPRESSMSGBOXES",
    "/NORESTART"
) -Wait -PassThru
if ($uninstallProcess.ExitCode -ne 0) {
    throw "Uninstaller exited with code $($uninstallProcess.ExitCode)."
}

$cleanupDeadline = [DateTime]::UtcNow.AddSeconds(15)
while ((Test-Path $installDirectory) -and [DateTime]::UtcNow -lt $cleanupDeadline) {
    Start-Sleep -Milliseconds 250
}
if (Test-Path $installDirectory) {
    throw "Uninstall left files in the installation directory."
}
foreach ($view in $views) {
    Assert-RegistrationRemoved -View $view
}

Write-Host "Per-user x86/x64 registration, real COM activation, worker handshake, and uninstall smoke tests passed."
