$ErrorActionPreference = "Stop"

$xamlPath = "src\ExcelReportBuilder.AddIn\Views\PivotPlusView.xaml"
$viewPath = "src\ExcelReportBuilder.AddIn\Views\PivotPlusView.xaml.cs"
$viewModelPath = "src\ExcelReportBuilder.AddIn\Presentation\PivotPlusViewModel.cs"
$hostContractPath = "src\ExcelReportBuilder.AddIn\Host\PivotPlusHostContracts.cs"
$hostPath = "src\ExcelReportBuilder.AddIn\Host\ExcelPivotPlusHostService.cs"
$bootstrapperPath = "src\ExcelReportBuilder.AddIn\Hosting\TaskPaneBootstrapper.cs"
$taskPaneHostPath = "src\ExcelReportBuilder.AddIn\Hosting\TaskPaneHost.cs"
$ribbonPath = "src\ExcelReportBuilder.AddIn\Ribbon\RibbonMarkup.cs"

$xaml = Get-Content -LiteralPath $xamlPath -Raw
$view = Get-Content -LiteralPath $viewPath -Raw
$viewModel = Get-Content -LiteralPath $viewModelPath -Raw
$hostContract = Get-Content -LiteralPath $hostContractPath -Raw
$hostSource = Get-Content -LiteralPath $hostPath -Raw
$bootstrapper = Get-Content -LiteralPath $bootstrapperPath -Raw
$taskPaneHost = Get-Content -LiteralPath $taskPaneHostPath -Raw
$ribbon = Get-Content -LiteralPath $ribbonPath -Raw

$requiredXaml = @(
    'MinWidth="320"',
    'MinHeight="480"',
    'AutomationProperties.Name="PivotTable Plus pane"',
    'AutomationProperties.LiveSetting="Polite"',
    'ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto"',
    '<Border Grid.Row="2"',
    'Text="PivotTable Fields"',
    'Content="Open Excel Field List"',
    'Text="Drag one or more fields into Filters, Columns, Rows, or Values. Drag placed fields to reorder or move them."',
    'AutomationProperties.Name="Available PivotTable fields"',
    'AllowDrop="True"',
    'Text="Preview only — Excel changes after confirmation."',
    'Content="Apply"',
    'Content="Enable PivotTable+"',
    'Header="Dates &amp; periods"',
    'Header="Portion %"',
    'Content="Add Portion %"',
    'Content="Undo"',
    'Key="R"',
    'Key="Enter"',
    'Key="Z"'
)
foreach ($value in $requiredXaml) {
    if (-not $xaml.Contains($value)) {
        throw "PivotTable+ pane contract is missing: $value"
    }
}

[xml]$parsed = $xaml
if ($null -eq $parsed.DocumentElement) {
    throw "PivotTable+ XAML is not well formed."
}

if (-not $view.Contains('Dispatcher.Yield(DispatcherPriority.ContextIdle)') -or
    $view.IndexOf('Dispatcher.Yield(DispatcherPriority.ContextIdle)') -gt
        $view.IndexOf('viewModel.InitializeAsync()')) {
    throw "The PivotTable+ pane must yield out of Office's creation callback before inspecting Excel."
}

foreach ($value in @(
    'Task<PivotPlusPaneSnapshot> InspectAsync(',
    'Task<PivotPlusPaneSnapshot> ApplyLayoutAsync(',
    'Task<PivotPlusPaneSnapshot> EnableDataModelAsync(',
    'Task<PivotPlusPaneSnapshot> AddParentPortionAsync(',
    'Task<PivotPlusPaneSnapshot> UndoLastExtraAsync(',
    'void OpenExcelFieldList()')) {
    if (-not $hostContract.Contains($value)) {
        throw "PivotTable+ host contract is missing: $value"
    }
}

if (-not $hostSource.Contains('PivotTableNativeLayoutMutationService') -or
    -not $hostSource.Contains('PivotDataModelEnablementService') -or
    -not $hostSource.Contains('PivotModelMeasureMutationService') -or
    -not $hostSource.Contains('PivotDaxCompiler.Compile(definition)') -or
    -not $hostSource.Contains('PivotParentShareDenominator')) {
    throw "The pane host must use typed native layout and Portion measure services."
}

foreach ($forbidden in @('Process.Start(', 'File.Write', 'Worksheet.Add', 'Formula =', 'MDX', 'DAX =')) {
    if ($hostSource.Contains($forbidden)) {
        throw "The pane host exposes a forbidden mutation path: $forbidden"
    }
}

if (-not $viewModel.Contains('HasPendingChanges') -or
    -not $viewModel.Contains('BuildPlacementRequests()') -or
    -not $viewModel.Contains('Preview changed. Choose Apply to update Excel.') -or
    -not $viewModel.Contains('catch (Exception exception)')) {
    throw "The ViewModel must preserve preview-before-apply and visible failure feedback."
}

if (-not $bootstrapper.Contains('Func<IPivotPlusHostService>') -or
    -not $taskPaneHost.Contains('new PivotPlusView(') -or
    -not $taskPaneHost.Contains('RenderMode.SoftwareOnly') -or
    $taskPaneHost.Contains('new ReportBuilderView(')) {
    throw "The COM task-pane host must compose a software-rendered PivotTable+, not the retired report workbench."
}

if (-not $ribbon.Contains('label=""PivotTable+""') -or
    -not $ribbon.Contains('PivotFieldListShowHide') -or
    -not $ribbon.Contains('OnOpenExcelFieldList')) {
    throw "The Ribbon must expose PivotTable+ and the native Excel Field List action."
}

Write-Host "PivotTable+ pane, preview, accessibility, typed host, and Ribbon contracts passed."
