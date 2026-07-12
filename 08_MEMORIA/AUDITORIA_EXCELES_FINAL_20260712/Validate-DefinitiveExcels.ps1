[CmdletBinding()]
param(
    [string[]]$WorkbookPaths = @(
        "C:\Users\julio\Desktop\TFM\08_MEMORIA\ENTREGABLE\TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712.xlsx",
        "C:\Users\julio\Desktop\TFM\08_MEMORIA\ENTREGABLE\TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_DEFINITIVO_20260712.xlsx"
    ),
    [string]$OutputPath = "C:\Users\julio\Desktop\TFM\08_MEMORIA\AUDITORIA_EXCELES_FINAL_20260712\EXCEL_COM_VALIDACION_FINAL_20260712.json"
)

$ErrorActionPreference = "Stop"
$xlCalculationAutomatic = -4105
$xlCellTypeConstants = 2
$xlCellTypeFormulas = -4123
$xlErrors = 16

if (-not ("TFMExcelValidationNative" -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class TFMExcelValidationNative {
    [DllImport("user32.dll", SetLastError=true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
"@
}

function Release-ComObject {
    param($Object)
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object) } catch {}
    }
}

function Get-SpecialCellCount {
    param($Range, [int]$CellType, [int]$ValueType)
    $special = $null
    try {
        $special = $Range.SpecialCells($CellType, $ValueType)
        return [int64]$special.CountLarge
    } catch {
        return [int64]0
    } finally {
        Release-ComObject $special
    }
}

function Inspect-Workbook {
    param($Workbook)
    $sheetDetails = [System.Collections.Generic.List[object]]::new()
    $totalCharts = 0
    $totalFormulaErrors = 0
    $totalStoredErrors = 0

    foreach ($worksheet in @($Workbook.Worksheets)) {
        $used = $null
        $chartObjects = $null
        try {
            $used = $worksheet.UsedRange
            $formulaErrors = Get-SpecialCellCount -Range $used -CellType $xlCellTypeFormulas -ValueType $xlErrors
            $storedErrors = Get-SpecialCellCount -Range $used -CellType $xlCellTypeConstants -ValueType $xlErrors
            $chartObjects = $worksheet.ChartObjects()
            $chartCount = [int]$chartObjects.Count
            $totalCharts += $chartCount
            $totalFormulaErrors += $formulaErrors
            $totalStoredErrors += $storedErrors
            $charts = [System.Collections.Generic.List[object]]::new()
            for ($i = 1; $i -le $chartCount; $i++) {
                $chartObject = $null
                $chart = $null
                $seriesCollection = $null
                try {
                    $chartObject = $chartObjects.Item($i)
                    $chart = $chartObject.Chart
                    $seriesCollection = $chart.SeriesCollection()
                    $charts.Add([pscustomobject]@{
                        Name = [string]$chartObject.Name
                        HasTitle = [bool]$chart.HasTitle
                        Title = if ($chart.HasTitle) { [string]$chart.ChartTitle.Text } else { "" }
                        SeriesCount = [int]$seriesCollection.Count
                    })
                } finally {
                    Release-ComObject $seriesCollection
                    Release-ComObject $chart
                    Release-ComObject $chartObject
                }
            }
            $sheetDetails.Add([pscustomobject]@{
                Name = [string]$worksheet.Name
                Visible = [int]$worksheet.Visible
                UsedRows = [int]$used.Rows.Count
                UsedColumns = [int]$used.Columns.Count
                FormulaErrors = $formulaErrors
                StoredErrors = $storedErrors
                ChartCount = $chartCount
                Charts = @($charts)
            })
        } finally {
            Release-ComObject $chartObjects
            Release-ComObject $used
            Release-ComObject $worksheet
        }
    }
    return [pscustomobject]@{
        SheetCount = [int]$Workbook.Worksheets.Count
        ChartCount = $totalCharts
        FormulaErrors = $totalFormulaErrors
        StoredErrors = $totalStoredErrors
        Sheets = @($sheetDetails)
    }
}

$excel = $null
$excelPid = 0
$results = [System.Collections.Generic.List[object]]::new()
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    try { $excel.AutomationSecurity = 3 } catch {}
    [uint32]$pidValue = 0
    [void][TFMExcelValidationNative]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$pidValue)
    $excelPid = [int]$pidValue

    foreach ($path in $WorkbookPaths) {
        $resolved = (Resolve-Path -LiteralPath $path).Path
        $writeWorkbook = $null
        $readWorkbook = $null
        $saveInspection = $null
        $reopenInspection = $null
        $saved = $false
        $reopened = $false
        $repairDetected = $false
        $exception = $null
        try {
            $writeWorkbook = $excel.Workbooks.Open($resolved, 0, $false)
            try { $excel.Calculation = $xlCalculationAutomatic } catch {}
            try { $excel.CalculateBeforeSave = $true } catch {}
            try { $writeWorkbook.ForceFullCalculation = $true } catch {}
            try { $writeWorkbook.FullCalculationOnLoad = $true } catch {}
            $excel.CalculateFullRebuild()
            $saveInspection = Inspect-Workbook -Workbook $writeWorkbook
            $writeWorkbook.Save()
            $saved = $true
            $writeWorkbook.Close($false)
            Release-ComObject $writeWorkbook
            $writeWorkbook = $null

            $readWorkbook = $excel.Workbooks.Open($resolved, 0, $true)
            $reopened = $true
            $reopenInspection = Inspect-Workbook -Workbook $readWorkbook
            $readWorkbook.Close($false)
            Release-ComObject $readWorkbook
            $readWorkbook = $null
        } catch {
            $exception = $_.Exception.Message
            throw
        } finally {
            if ($readWorkbook) { try { $readWorkbook.Close($false) } catch {}; Release-ComObject $readWorkbook }
            if ($writeWorkbook) { try { $writeWorkbook.Close($false) } catch {}; Release-ComObject $writeWorkbook }
        }
        $results.Add([pscustomobject]@{
            Path = $resolved
            Saved = $saved
            ReopenedReadOnly = $reopened
            RepairDetected = $repairDetected
            Exception = $exception
            SaveInspection = $saveInspection
            ReopenInspection = $reopenInspection
            Valid = ($saved -and $reopened -and -not $repairDetected -and $reopenInspection.FormulaErrors -eq 0 -and $reopenInspection.StoredErrors -eq 0 -and (@($reopenInspection.Sheets | ForEach-Object { $_.Charts } | Where-Object { $_.SeriesCount -eq 0 }).Count -eq 0))
            ValidatedAt = (Get-Date).ToString("o")
        })
    }
} finally {
    if ($excel) { try { $excel.Quit() } catch {}; Release-ComObject $excel }
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 2
    if ($excelPid -gt 0 -and (Get-Process -Id $excelPid -ErrorAction SilentlyContinue)) {
        Stop-Process -Id $excelPid -Force
    }
}

$payload = [pscustomobject]@{
    GeneratedAt = (Get-Date).ToString("o")
    ExcelPid = $excelPid
    Results = @($results)
}
$payload | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $OutputPath -Encoding utf8
$payload.Results | Select-Object Path,Saved,ReopenedReadOnly,RepairDetected,Valid,@{N="Sheets";E={$_.ReopenInspection.SheetCount}},@{N="Charts";E={$_.ReopenInspection.ChartCount}},@{N="FormulaErrors";E={$_.ReopenInspection.FormulaErrors}},@{N="StoredErrors";E={$_.ReopenInspection.StoredErrors}}
