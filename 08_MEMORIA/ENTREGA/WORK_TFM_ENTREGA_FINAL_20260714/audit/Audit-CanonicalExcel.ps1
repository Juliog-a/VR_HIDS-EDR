[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ManifestPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Release-ComObject {
    param([object]$Object)
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Escape-TsvValue {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return '' }
    $text = [string]$Value
    return ($text -replace "`r`n", '\n' -replace "`r", '\n' -replace "`n", '\n' -replace "`t", '\t')
}

function Sanitize-FileName {
    param([string]$Name)
    $invalid = [IO.Path]::GetInvalidFileNameChars()
    $result = $Name
    foreach ($char in $invalid) { $result = $result.Replace([string]$char, '_') }
    return $result
}

$manifest = Get-Content -Raw -LiteralPath (Resolve-Path -LiteralPath $ManifestPath).Path | ConvertFrom-Json
if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir | Out-Null }
$outFull = (Resolve-Path -LiteralPath $OutputDir).Path
$progressPath = Join-Path $outFull 'excel_canonical_audit_progress.log'
[IO.File]::WriteAllText($progressPath, "INICIO $(Get-Date -Format o)`r`n", $utf8NoBom)
function Log([string]$Message) { [IO.File]::AppendAllText($progressPath, "$(Get-Date -Format o) | $Message`r`n", $utf8NoBom) }

$baselinePids = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$fatalError = $null
$excel = $null
$workbooks = $null
$report = [ordered]@{
    generated_at = (Get-Date).ToString('o')
    excel_version = $null
    excel_build = $null
    baseline_excel_pids = $baselinePids
    created_excel_pids = @()
    workbooks = @()
    after_excel_pids = @()
    forced_cleanup_pids = @()
    final_excel_pids = @()
}

try {
    Log 'Creando Excel.Application'
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    try { $excel.AutomationSecurity = 3 } catch {}
    try { $excel.Calculation = -4135 } catch {}
    $report.excel_version = [string]$excel.Version
    $report.excel_build = [string]$excel.Build
    $workbooks = $excel.Workbooks
    $report.created_excel_pids = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id | Where-Object { $baselinePids -notcontains $_ })
    Log "Excel listo: $($report.excel_version) build $($report.excel_build)"

    foreach ($path in @($manifest.workbooks)) {
        if (-not (Test-Path -LiteralPath $path)) { throw "No existe el Excel canonico: $path" }
        $fullPath = (Resolve-Path -LiteralPath $path).Path
        $item = Get-Item -LiteralPath $fullPath
        $hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $fullPath).Hash
        $wb = $null
        $sheets = $null
        try {
            Log "Abriendo solo lectura: $fullPath"
            $wb = $workbooks.Open($fullPath, 0, $true)
            if ($null -eq $wb) { throw "Excel devolvio workbook nulo: $fullPath" }
            $sheetRecords = @()
            $sheets = $wb.Worksheets
            for ($sheetIndex = 1; $sheetIndex -le $sheets.Count; $sheetIndex++) {
                $sheet = $null
                $used = $null
                $rows = $null
                $cols = $null
                $listObjects = $null
                $chartObjects = $null
                try {
                    $sheet = $sheets.Item($sheetIndex)
                    $used = $sheet.UsedRange
                    $rows = $used.Rows
                    $cols = $used.Columns
                    $rowCount = [int]$rows.Count
                    $colCount = [int]$cols.Count
                    $firstRow = [int]$used.Row
                    $firstCol = [int]$used.Column
                    $values = $used.Value2
                    $safeSheet = Sanitize-FileName ([string]$sheet.Name)
                    $tsvName = ([IO.Path]::GetFileNameWithoutExtension($fullPath) + '__' + ('{0:D2}' -f $sheetIndex) + '__' + $safeSheet + '.tsv')
                    $tsvPath = Join-Path $outFull $tsvName
                    $writer = New-Object System.IO.StreamWriter($tsvPath, $false, $utf8NoBom)
                    try {
                        if ($rowCount -eq 1 -and $colCount -eq 1) {
                            $writer.WriteLine((Escape-TsvValue $values))
                        }
                        else {
                            for ($r = 1; $r -le $rowCount; $r++) {
                                $line = New-Object System.Text.StringBuilder
                                for ($c = 1; $c -le $colCount; $c++) {
                                    if ($c -gt 1) { [void]$line.Append("`t") }
                                    [void]$line.Append((Escape-TsvValue $values[$r, $c]))
                                }
                                $writer.WriteLine($line.ToString())
                            }
                        }
                    }
                    finally {
                        $writer.Dispose()
                    }
                    $listObjects = $sheet.ListObjects
                    $chartObjects = $sheet.ChartObjects()
                    $sheetRecords += [ordered]@{
                        index = $sheetIndex
                        name = [string]$sheet.Name
                        visible = [int]$sheet.Visible
                        used_first_row = $firstRow
                        used_first_column = $firstCol
                        used_rows = $rowCount
                        used_columns = $colCount
                        table_count = $(try { [int]$listObjects.Count } catch { $null })
                        chart_count = $(try { [int]$chartObjects.Count } catch { $null })
                        tsv = $tsvPath
                    }
                    Log "Hoja exportada: $($sheet.Name) ${rowCount}x${colCount}"
                }
                finally {
                    Release-ComObject $chartObjects
                    Release-ComObject $listObjects
                    Release-ComObject $cols
                    Release-ComObject $rows
                    Release-ComObject $used
                    Release-ComObject $sheet
                }
            }
            $report.workbooks += [ordered]@{
                name = $item.Name
                path = $fullPath
                size_bytes = [int64]$item.Length
                modified = $item.LastWriteTime.ToString('o')
                sha256 = $hash
                read_only_open = [bool]$wb.ReadOnly
                worksheets = [int]$sheets.Count
                calculation_version = $(try { [int]$wb.CalculationVersion } catch { $null })
                sheets = $sheetRecords
            }
            Log "Workbook completado: $fullPath"
        }
        finally {
            Release-ComObject $sheets
            if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
            Release-ComObject $wb
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }
}
catch {
    $fatalError = $_
    $errorRecord = [ordered]@{
        generated_at = (Get-Date).ToString('o')
        message = $_.Exception.Message
        exception_type = $_.Exception.GetType().FullName
        position = $_.InvocationInfo.PositionMessage
        script_stack = $_.ScriptStackTrace
    }
    [IO.File]::WriteAllText((Join-Path $outFull 'excel_canonical_audit_error.json'), ($errorRecord | ConvertTo-Json -Depth 6), $utf8NoBom)
    Log "ERROR: $($_.Exception.Message)"
}
finally {
    Release-ComObject $workbooks
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 2
}

$afterPids = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$orphanPids = @($afterPids | Where-Object { $report.created_excel_pids -contains $_ })
$report.after_excel_pids = $afterPids
if ($orphanPids.Count -gt 0) {
    foreach ($pidToStop in $orphanPids) { try { Stop-Process -Id $pidToStop -Force -ErrorAction Stop } catch {} }
    Start-Sleep -Seconds 2
    $report.forced_cleanup_pids = $orphanPids
}
$report.final_excel_pids = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$reportPath = Join-Path $outFull 'excel_canonical_audit.json'
[IO.File]::WriteAllText($reportPath, ($report | ConvertTo-Json -Depth 10), $utf8NoBom)

[pscustomobject]@{
    Report = $reportPath
    ExcelVersion = $report.excel_version
    ExcelBuild = $report.excel_build
    Workbooks = $report.workbooks.Count
    ForcedCleanupPids = ($report.forced_cleanup_pids -join ',')
    FinalExcelPids = ($report.final_excel_pids -join ',')
} | Format-List

if (@($report.final_excel_pids | Where-Object { $report.created_excel_pids -contains $_ }).Count -gt 0) {
    throw "Persisten EXCEL creados por la auditoria: $($report.final_excel_pids -join ',')"
}
if ($null -ne $fatalError) { throw "Auditoria Excel fallida tras limpieza COM: $($fatalError.Exception.Message)" }
