param(
    [ValidateRange(1, 1048575)]
    [int]$Rows = 100000,

    [ValidateRange(1000, 50000)]
    [int]$ChunkRows = 10000,

    [string]$OutputPath = "",

    [switch]$LeaveOpen
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $artifactDirectory = Join-Path (Get-Location) "artifacts"
    New-Item -ItemType Directory -Path $artifactDirectory -Force | Out-Null
    $OutputPath = Join-Path $artifactDirectory "synthetic-wide-benchmark.xlsx"
}
$OutputPath = [IO.Path]::GetFullPath($OutputPath)

$excel = $null
$workbook = $null
$worksheet = $null
$started = [Diagnostics.Stopwatch]::StartNew()
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = [bool]$LeaveOpen
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.Calculation = -4135

    $workbook = $excel.Workbooks.Add()
    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Name = "Raw Data"
    $headers = @("Category", "Channel", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    for ($column = 0; $column -lt $headers.Count; $column++) {
        $worksheet.Cells.Item(1, $column + 1).Value2 = $headers[$column]
    }

    for ($start = 1; $start -le $Rows; $start += $ChunkRows) {
        $count = [Math]::Min($ChunkRows, $Rows - $start + 1)
        $values = New-Object 'object[,]' $count, $headers.Count
        for ($offset = 0; $offset -lt $count; $offset++) {
            $rowNumber = $start + $offset
            $values[$offset, 0] = "Category " + (($rowNumber % 25) + 1)
            $values[$offset, 1] = if (($rowNumber % 2) -eq 0) { "Direct" } else { "Partner" }
            for ($month = 1; $month -le 12; $month++) {
                $values[$offset, $month + 1] = [double](($rowNumber % 1000) + $month)
            }
        }

        $top = $start + 1
        $bottom = $top + $count - 1
        $target = $worksheet.Range[
            $worksheet.Cells.Item($top, 1),
            $worksheet.Cells.Item($bottom, $headers.Count)]
        $target.Value2 = $values
        Write-Progress -Activity "Generating synthetic benchmark" -Status "$bottom of $($Rows + 1) worksheet rows" -PercentComplete (($start + $count - 1) * 100 / $Rows)
    }

    $sourceRange = $worksheet.Range[
        $worksheet.Cells.Item(1, 1),
        $worksheet.Cells.Item($Rows + 1, $headers.Count)]
    $table = $worksheet.ListObjects.Add(1, $sourceRange, $null, 1)
    $table.Name = "SyntheticWideData"
    $worksheet.Columns.Item(1).ColumnWidth = 16
    $worksheet.Columns.Item(2).ColumnWidth = 12
    $workbook.SaveAs($OutputPath, 51)
    $started.Stop()

    [pscustomobject]@{
        Workbook = $OutputPath
        SourceRows = $Rows
        SourceColumns = $headers.Count
        ProjectedNormalizedRows = [long]$Rows * 12
        ExpectedBackend = if (([long]$Rows * 12) -gt 1048575) { "DataModel" } else { "Worksheet" }
        GenerationSeconds = [Math]::Round($started.Elapsed.TotalSeconds, 2)
    }
}
finally {
    Write-Progress -Activity "Generating synthetic benchmark" -Completed
    if ($null -ne $workbook -and -not $LeaveOpen) {
        $workbook.Close($false)
    }
    if ($null -ne $excel -and -not $LeaveOpen) {
        $excel.Quit()
    }
    foreach ($value in @($worksheet, $workbook, $excel)) {
        if ($null -ne $value) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)
        }
    }
}
