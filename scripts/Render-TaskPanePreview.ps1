$ErrorActionPreference = "Stop"

Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework

$assemblyPath = Resolve-Path "src\ExcelReportBuilder.AddIn\bin\Release\net48\ExcelReportBuilder.AddIn.dll"
$outputDirectory = Join-Path (Resolve-Path ".") "artifacts"
New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
$outputPath = Join-Path $outputDirectory "task-pane-preview.png"

$assembly = [Reflection.Assembly]::LoadFrom($assemblyPath)
$viewType = $assembly.GetType("ExcelReportBuilder.AddIn.Views.ReportBuilderView", $true)
$view = [Activator]::CreateInstance($viewType)
try {
    $view.Width = 420
    $view.Height = 900
    $size = [System.Windows.Size]::new(420, 900)
    $view.Measure($size)
    $view.Arrange([System.Windows.Rect]::new(0, 0, 420, 900))
    $view.UpdateLayout()

    $bitmap = [System.Windows.Media.Imaging.RenderTargetBitmap]::new(
        420,
        900,
        96,
        96,
        [System.Windows.Media.PixelFormats]::Pbgra32)
    $bitmap.Render($view)
    $encoder = [System.Windows.Media.Imaging.PngBitmapEncoder]::new()
    $encoder.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
    $stream = [IO.File]::Open($outputPath, [IO.FileMode]::Create)
    try {
        $encoder.Save($stream)
    }
    finally {
        $stream.Dispose()
    }
}
finally {
    if ($view -is [IDisposable]) {
        $view.Dispose()
    }
}

if (-not (Test-Path $outputPath)) {
    throw "The synthetic task-pane preview was not created."
}

Write-Host "Synthetic task-pane preview: $outputPath"
