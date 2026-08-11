$ErrorActionPreference = "Stop"

$assemblyPath = Resolve-Path "src\ExcelReportBuilder.AddIn\bin\Release\net48\ExcelReportBuilder.AddIn.dll"
$assembly = [Reflection.Assembly]::LoadFrom($assemblyPath)

$expected = @(
    @{
        Name = "ExcelReportBuilder.AddIn.Com.ExcelReportBuilderAddIn"
        Guid = "F953480C-A73C-4121-9E21-18676EC34CE8"
        ProgId = "ExcelReportBuilder.AddIn"
    },
    @{
        Name = "ExcelReportBuilder.AddIn.Hosting.TaskPaneHost"
        Guid = "A3F4E10D-0DD1-420E-8B6F-E0A654BBEA16"
        ProgId = "ExcelReportBuilder.TaskPaneHost"
    }
)

foreach ($contract in $expected) {
    $type = $assembly.GetType($contract.Name, $true)
    if ($type.GUID.ToString().ToUpperInvariant() -ne $contract.Guid) {
        throw "COM GUID mismatch for $($contract.Name)."
    }
    $progId = [Runtime.InteropServices.Marshal]::GenerateProgIdForType($type)
    if ($progId -ne $contract.ProgId) {
        throw "COM ProgID mismatch for $($contract.Name)."
    }
    $instance = [Activator]::CreateInstance($type)
    if ($null -eq $instance) {
        throw "Could not instantiate $($contract.Name)."
    }
    if ($instance -is [IDisposable]) {
        $instance.Dispose()
    }
}

$regasm = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
& $regasm $assemblyPath /codebase /regfile:managed-com.reg | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "RegAsm contract export failed."
}
$registration = Get-Content managed-com.reg -Raw
foreach ($contract in $expected) {
    if (-not $registration.Contains($contract.Name)) {
        throw "RegAsm did not export $($contract.Name)."
    }
}

Write-Host "COM contract is valid."
