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

function Get-RequiredAttribute {
    param(
        [Parameter(Mandatory = $true)][Reflection.ICustomAttributeProvider]$Provider,
        [Parameter(Mandatory = $true)][Type]$AttributeType,
        [Parameter(Mandatory = $true)][string]$ContractName
    )

    $attributes = @($Provider.GetCustomAttributes($AttributeType, $false))
    if ($attributes.Count -ne 1) {
        throw "$ContractName must have exactly one $($AttributeType.Name)."
    }

    return $attributes[0]
}

function Assert-InParameter {
    param(
        [Parameter(Mandatory = $true)][Reflection.ParameterInfo]$Parameter,
        [Parameter(Mandatory = $true)][string]$ContractName
    )

    if (-not $Parameter.IsIn -or $Parameter.IsOut) {
        throw "$ContractName must be marshalled as an in-only parameter."
    }
}

function Assert-VariantSafeArray {
    param(
        [Parameter(Mandatory = $true)][Reflection.ParameterInfo]$Parameter,
        [Parameter(Mandatory = $true)][string]$ContractName
    )

    Assert-InParameter -Parameter $Parameter -ContractName $ContractName
    if (-not $Parameter.ParameterType.IsByRef -or
        $Parameter.ParameterType.GetElementType() -ne [Array]) {
        throw "$ContractName must use the System.Array& managed signature."
    }
    $marshal = Get-RequiredAttribute `
        -Provider $Parameter `
        -AttributeType ([Runtime.InteropServices.MarshalAsAttribute]) `
        -ContractName $ContractName
    if ($marshal.Value -ne [Runtime.InteropServices.UnmanagedType]::SafeArray -or
        $marshal.SafeArraySubType -ne [Runtime.InteropServices.VarEnum]::VT_VARIANT) {
        throw "$ContractName must be marshalled as SAFEARRAY(VARIANT)."
    }
}

$extensibility = $assembly.GetType(
    "ExcelReportBuilder.AddIn.Interop.IDTExtensibility2",
    $true)
$ribbonExtensibility = $assembly.GetType(
    "ExcelReportBuilder.AddIn.Interop.IRibbonExtensibility",
    $true)
$taskPaneFactory = $assembly.GetType(
    "ExcelReportBuilder.AddIn.Interop.ICTPFactory",
    $true)
$taskPaneConsumer = $assembly.GetType(
    "ExcelReportBuilder.AddIn.Interop.ICustomTaskPaneConsumer",
    $true)

$expectedInterfaceGuids = @{
    $extensibility.FullName = "B65AD801-ABAF-11D0-BB8B-00A0C90F2744"
    $ribbonExtensibility.FullName = "000C0396-0000-0000-C000-000000000046"
    $taskPaneFactory.FullName = "000C033D-0000-0000-C000-000000000046"
    $taskPaneConsumer.FullName = "000C033E-0000-0000-C000-000000000046"
}
foreach ($interface in @(
    $extensibility,
    $ribbonExtensibility,
    $taskPaneFactory,
    $taskPaneConsumer)) {
    if ($interface.GUID.ToString().ToUpperInvariant() -ne
        $expectedInterfaceGuids[$interface.FullName]) {
        throw "COM GUID mismatch for $($interface.FullName)."
    }

    $interfaceType = Get-RequiredAttribute `
        -Provider $interface `
        -AttributeType ([Runtime.InteropServices.InterfaceTypeAttribute]) `
        -ContractName $interface.FullName
    if ($interfaceType.Value -ne [Runtime.InteropServices.ComInterfaceType]::InterfaceIsDual) {
        throw "$($interface.FullName) must preserve Office's dual-interface ABI."
    }
}

$addInType = $assembly.GetType(
    "ExcelReportBuilder.AddIn.Com.ExcelReportBuilderAddIn",
    $true)
foreach ($requiredInterface in @(
    $extensibility,
    $ribbonExtensibility,
    $taskPaneConsumer)) {
    if (-not $requiredInterface.IsAssignableFrom($addInType)) {
        throw "ExcelReportBuilderAddIn must implement $($requiredInterface.FullName)."
    }
}

foreach ($methodName in @(
    "OnConnection",
    "OnDisconnection",
    "OnAddInsUpdate",
    "OnStartupComplete",
    "OnBeginShutdown")) {
    $method = $extensibility.GetMethod($methodName)
    $dispId = Get-RequiredAttribute `
        -Provider $method `
        -AttributeType ([Runtime.InteropServices.DispIdAttribute]) `
        -ContractName "IDTExtensibility2.$methodName"
    $expectedDispId = [Array]::IndexOf(@(
        "OnConnection",
        "OnDisconnection",
        "OnAddInsUpdate",
        "OnStartupComplete",
        "OnBeginShutdown"), $methodName) + 1
    if ($dispId.Value -ne $expectedDispId) {
        throw "IDTExtensibility2.$methodName must retain DispId $expectedDispId."
    }
    $parameters = @($method.GetParameters())
    Assert-VariantSafeArray `
        -Parameter $parameters[$parameters.Count - 1] `
        -ContractName "IDTExtensibility2.$methodName custom"
    if ($parameters.Count -gt 1) {
        foreach ($parameter in $parameters[0..($parameters.Count - 2)]) {
            Assert-InParameter `
                -Parameter $parameter `
                -ContractName "IDTExtensibility2.$methodName $($parameter.Name)"
        }
    }
}

$onConnectionParameters = $extensibility.GetMethod("OnConnection").GetParameters()
foreach ($parameterIndex in @(0, 2)) {
    $parameter = $onConnectionParameters[$parameterIndex]
    $marshal = Get-RequiredAttribute `
        -Provider $parameter `
        -AttributeType ([Runtime.InteropServices.MarshalAsAttribute]) `
        -ContractName "IDTExtensibility2.OnConnection $($parameter.Name)"
    if ($marshal.Value -ne [Runtime.InteropServices.UnmanagedType]::IDispatch) {
        throw "IDTExtensibility2.OnConnection $($parameter.Name) must marshal as IDispatch."
    }
}

$factoryCallback = $taskPaneConsumer.GetMethod("CTPFactoryAvailable")
$factoryCallbackDispId = Get-RequiredAttribute `
    -Provider $factoryCallback `
    -AttributeType ([Runtime.InteropServices.DispIdAttribute]) `
    -ContractName "ICustomTaskPaneConsumer.CTPFactoryAvailable"
if ($factoryCallbackDispId.Value -ne 1) {
    throw "ICustomTaskPaneConsumer.CTPFactoryAvailable must retain DispId 1."
}
$factoryParameter = $factoryCallback.GetParameters()[0]
Assert-InParameter `
    -Parameter $factoryParameter `
    -ContractName "ICustomTaskPaneConsumer.CTPFactoryAvailable taskPaneFactory"
if ($factoryParameter.ParameterType -ne $taskPaneFactory) {
    throw "ICustomTaskPaneConsumer must accept the exact ICTPFactory contract."
}
$factoryMarshal = Get-RequiredAttribute `
    -Provider $factoryParameter `
    -AttributeType ([Runtime.InteropServices.MarshalAsAttribute]) `
    -ContractName "ICustomTaskPaneConsumer.CTPFactoryAvailable taskPaneFactory"
if ($factoryMarshal.Value -ne [Runtime.InteropServices.UnmanagedType]::Interface) {
    throw "ICustomTaskPaneConsumer must marshal ICTPFactory as an interface."
}

$createTaskPane = $taskPaneFactory.GetMethod("CreateCTP")
$createTaskPaneDispId = Get-RequiredAttribute `
    -Provider $createTaskPane `
    -AttributeType ([Runtime.InteropServices.DispIdAttribute]) `
    -ContractName "ICTPFactory.CreateCTP"
if ($createTaskPaneDispId.Value -ne 1) {
    throw "ICTPFactory.CreateCTP must retain DispId 1."
}
foreach ($parameter in $createTaskPane.GetParameters()) {
    Assert-InParameter `
        -Parameter $parameter `
        -ContractName "ICTPFactory.CreateCTP $($parameter.Name)"
}
$createTaskPaneParameters = $createTaskPane.GetParameters()
foreach ($parameterIndex in @(0, 1)) {
    $parameter = $createTaskPaneParameters[$parameterIndex]
    $marshal = Get-RequiredAttribute `
        -Provider $parameter `
        -AttributeType ([Runtime.InteropServices.MarshalAsAttribute]) `
        -ContractName "ICTPFactory.CreateCTP $($parameter.Name)"
    if ($marshal.Value -ne [Runtime.InteropServices.UnmanagedType]::BStr) {
        throw "ICTPFactory.CreateCTP $($parameter.Name) must marshal as BSTR."
    }
}
$parentWindowParameter = $createTaskPaneParameters[2]
if (-not $parentWindowParameter.IsOptional) {
    throw "ICTPFactory.CreateCTP parentWindow must remain optional."
}
$parentWindowMarshal = Get-RequiredAttribute `
    -Provider $parentWindowParameter `
    -AttributeType ([Runtime.InteropServices.MarshalAsAttribute]) `
    -ContractName "ICTPFactory.CreateCTP parentWindow"
if ($parentWindowMarshal.Value -ne [Runtime.InteropServices.UnmanagedType]::Struct) {
    throw "ICTPFactory.CreateCTP parentWindow must marshal as VARIANT."
}
$createTaskPaneReturn = Get-RequiredAttribute `
    -Provider $createTaskPane.ReturnParameter `
    -AttributeType ([Runtime.InteropServices.MarshalAsAttribute]) `
    -ContractName "ICTPFactory.CreateCTP return"
if ($createTaskPaneReturn.Value -ne [Runtime.InteropServices.UnmanagedType]::Interface) {
    throw "ICTPFactory.CreateCTP must return an interface pointer."
}

$regasm = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
$contractArtifactDirectory = Join-Path (Resolve-Path ".") "artifacts\com-contract"
New-Item -ItemType Directory -Force -Path $contractArtifactDirectory | Out-Null
$registrationPath = Join-Path $contractArtifactDirectory "managed-com.reg"
& $regasm $assemblyPath /codebase "/regfile:$registrationPath" | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "RegAsm contract export failed."
}
$registration = Get-Content -LiteralPath $registrationPath -Raw
foreach ($contract in $expected) {
    if (-not $registration.Contains($contract.Name)) {
        throw "RegAsm did not export $($contract.Name)."
    }
}

Write-Host "COM contract is valid."
