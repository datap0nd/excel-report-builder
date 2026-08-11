$ErrorActionPreference = "Stop"

$xamlPath = "src\ExcelReportBuilder.AddIn\Views\ReportBuilderView.xaml"
$controllerPath = "src\ExcelReportBuilder.AddIn\Activity\OperationActivityController.cs"
$xaml = Get-Content -LiteralPath $xamlPath -Raw
$controller = Get-Content -LiteralPath $controllerPath -Raw

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

[xml]$parsed = $xaml
if ($null -eq $parsed.DocumentElement) {
    throw "The task-pane XAML is not well formed."
}

Write-Host "Task-pane accessibility and no-silence contract passed."
