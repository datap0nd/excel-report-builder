$ErrorActionPreference = "Stop"

$xamlPath = "src\ExcelReportBuilder.AddIn\Views\ReportBuilderView.xaml"
$controllerPath = "src\ExcelReportBuilder.AddIn\Activity\OperationActivityController.cs"
$hostInterfacePath = "src\ExcelReportBuilder.AddIn\Host\IReportBuilderHostService.cs"
$hostPath = "src\ExcelReportBuilder.AddIn\Host\ExcelReportBuilderHostService.cs"
$syntheticHostPath = "src\ExcelReportBuilder.AddIn\Host\SyntheticReportBuilderHostService.cs"
$viewModelPath = "src\ExcelReportBuilder.AddIn\Presentation\ShellViewModel.cs"
$xaml = Get-Content -LiteralPath $xamlPath -Raw
$controller = Get-Content -LiteralPath $controllerPath -Raw
$hostInterface = Get-Content -LiteralPath $hostInterfacePath -Raw
$taskPaneHostSource = Get-Content -LiteralPath $hostPath -Raw
$syntheticHost = Get-Content -LiteralPath $syntheticHostPath -Raw
$viewModel = Get-Content -LiteralPath $viewModelPath -Raw

function Get-SourceSection {
    param(
        [Parameter(Mandatory = $true)][string]$Source,
        [Parameter(Mandatory = $true)][string]$StartMarker,
        [Parameter(Mandatory = $true)][string]$EndMarker
    )

    $start = $Source.IndexOf($StartMarker, [System.StringComparison]::Ordinal)
    if ($start -lt 0) {
        throw "Source contract start marker is missing: $StartMarker"
    }

    $end = $Source.IndexOf(
        $EndMarker,
        $start + $StartMarker.Length,
        [System.StringComparison]::Ordinal)
    if ($end -lt 0) {
        throw "Source contract end marker is missing: $EndMarker"
    }

    return $Source.Substring($start, $end - $start)
}

$requiredXaml = @(
    'MinWidth="320"',
    'MinHeight="480"',
    'AutomationProperties.Name="Excel Report Builder workbench"',
    'AutomationProperties.LiveSetting="Polite"',
    'Key="D1"',
    'Key="D2"',
    'Key="D3"',
    'Key="D4"',
    'Key="P"',
    'Key="C"',
    'Key="B"',
    'Key="Enter"',
    'Text="Preparation steps"',
    'Text="Calculated metrics"',
    'Text="Report blocks"',
    'Text="Required checks"',
    'AutomationProperties.Name="Current operation identity"',
    'Header="Column" Binding="{Binding Name}" Width="*" MinWidth="96"',
    'Command="{Binding TogglePauseCommand}"',
    'Command="{Binding CancelCommand}"'
)
foreach ($value in $requiredXaml) {
    if (-not $xaml.Contains($value)) {
        throw "Task-pane accessibility contract is missing: $value"
    }
}

if ($xaml -match 'IsIndeterminate\s*=\s*"True"') {
    throw "The task pane cannot replace continuous feedback with an unexplained indeterminate progress bar."
}
if (-not $controller.Contains('TimeSpan.FromSeconds(15)')) {
    throw "The activity controller must retain the 15-second no-silence heartbeat."
}
if (-not $controller.Contains('MaximumTimelineEntries = 200')) {
    throw "The activity timeline must remain bounded."
}

$persistSignature = 'Task PersistEndpointSettingsAsync('
if (-not $hostInterface.Contains($persistSignature)) {
    throw "The host boundary must expose explicit endpoint-settings persistence."
}
if (-not $syntheticHost.Contains('public Task PersistEndpointSettingsAsync(')) {
    throw "The synthetic host must implement the endpoint-settings persistence contract."
}

$hostDiscovery = Get-SourceSection `
    -Source $taskPaneHostSource `
    -StartMarker 'public Task<IReadOnlyList<string>> DiscoverModelsAsync(' `
    -EndMarker 'public Task<EndpointCheckResult> CheckEndpointAsync('
$hostCheck = Get-SourceSection `
    -Source $taskPaneHostSource `
    -StartMarker 'public Task<EndpointCheckResult> CheckEndpointAsync(' `
    -EndMarker 'public Task PersistEndpointSettingsAsync('
$hostPersist = Get-SourceSection `
    -Source $taskPaneHostSource `
    -StartMarker 'public Task PersistEndpointSettingsAsync(' `
    -EndMarker 'public Task<IReadOnlyList<HostCheckResult>> RunChecksAsync('

foreach ($probe in @($hostDiscovery, $hostCheck)) {
    if ($probe.Contains('SaveEndpointAsync(') -or
        $probe.Contains('PersistEndpointSettingsAsync(')) {
        throw "Endpoint probes must not persist settings before the ViewModel accepts their result."
    }
}
if (-not $hostPersist.Contains('MaterializeEndpointAsync(') -or
    -not $hostPersist.Contains('AgentEndpointPolicy.Validate(endpoint)') -or
    -not $hostPersist.Contains('SaveEndpointAsync(endpoint, token)')) {
    throw "Explicit endpoint persistence must preserve credential materialization, policy validation, and protected storage."
}

$viewModelDiscovery = Get-SourceSection `
    -Source $viewModel `
    -StartMarker 'private async Task DiscoverModelsAsync()' `
    -EndMarker 'private async Task CheckEndpointAsync()'
$viewModelCheck = Get-SourceSection `
    -Source $viewModel `
    -StartMarker 'private async Task CheckEndpointAsync()' `
    -EndMarker 'private async Task RunChecksAsync()'

function Assert-VersionGatedPersistence {
    param(
        [Parameter(Mandatory = $true)][string]$Source,
        [Parameter(Mandatory = $true)][string]$RequestMarker
    )

    $requestIndex = $Source.IndexOf(
        $RequestMarker,
        [System.StringComparison]::Ordinal)
    $acceptIndex = $Source.IndexOf(
        'TryPersistCurrentEndpointSettings(',
        [System.StringComparison]::Ordinal)
    if ($requestIndex -lt 0 -or $acceptIndex -lt $requestIndex) {
        throw "Endpoint results must pass through the version-gated persistence path."
    }
}

Assert-VersionGatedPersistence `
    -Source $viewModelDiscovery `
    -RequestMarker '_hostService.DiscoverModelsAsync('
Assert-VersionGatedPersistence `
    -Source $viewModelCheck `
    -RequestMarker '_hostService.CheckEndpointAsync('

$viewModelPersist = Get-SourceSection `
    -Source $viewModel `
    -StartMarker 'private bool TryPersistCurrentEndpointSettings(' `
    -EndMarker 'private void InvalidateEndpointCheck()'
$versionGuard = 'if (requestedConfigurationVersion != _endpointConfigurationVersion)'
$guardMatches = [regex]::Matches($viewModelPersist, [regex]::Escape($versionGuard))
$persistCallIndex = $viewModelPersist.IndexOf(
    '_hostService.PersistEndpointSettingsAsync(',
    [System.StringComparison]::Ordinal)
$recordIndex = $viewModelPersist.IndexOf(
    'RecordPersistedEndpoint(',
    [System.StringComparison]::Ordinal)
if ($guardMatches.Count -ne 2 -or
    $guardMatches[0].Index -gt $persistCallIndex -or
    $guardMatches[1].Index -lt $persistCallIndex -or
    $recordIndex -lt $guardMatches[1].Index) {
    throw "Endpoint persistence must remain between pre-commit and post-commit configuration-version guards."
}
if (-not $viewModelPersist.Contains('using (SecureString? apiKey = CopyApiKey())') -or
    -not $viewModelPersist.Contains('.GetAwaiter()') -or
    -not $viewModelPersist.Contains('.GetResult()')) {
    throw "Endpoint persistence must copy and dispose the API key and complete without yielding on the UI thread."
}

$canBuildDraft = Get-SourceSection `
    -Source $viewModel `
    -StartMarker 'private bool CanBuildDraft()' `
    -EndMarker 'private bool CanSendChat()'
if (-not $canBuildDraft.Contains('_agentAppliedSpecification.HasCanonicalReportSpec')) {
    throw "An exact canonical Chat or saved setup must remain rebuildable when its read-only manual projection has no Rows."
}

[xml]$parsed = $xaml
if ($null -eq $parsed.DocumentElement) {
    throw "The task-pane XAML is not well formed."
}

Write-Host "Task-pane accessibility, no-silence, and endpoint persistence contracts passed."
