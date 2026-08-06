param(
    [string]$Path = "C:\Users\julio\Desktop\TFM\08_MEMORIA\AUDITORIA_EXCELES_FINAL_20260712\TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712_CUSTOM_WORKING.xlsx",
    [string]$Output = "C:\Users\julio\Desktop\TFM\08_MEMORIA\AUDITORIA_EXCELES_FINAL_20260712\CUSTOM_REVIEW_COM_VALIDATION.json"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Release-Com([object]$Object) {
    if ($null -ne $Object) {
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object) } catch { }
    }
}

function Count-Charts($Workbook) {
    $count = 0
    $sheets = $Workbook.Worksheets
    try {
        for ($i = 1; $i -le $sheets.Count; $i++) {
            $sheet = $sheets.Item($i)
            $charts = $sheet.ChartObjects()
            $count += $charts.Count
            Release-Com $charts
            Release-Com $sheet
        }
    } finally { Release-Com $sheets }
    return $count
}

function Count-Pending($Workbook) {
    $count = 0
    $sheets = $Workbook.Worksheets
    try {
        for ($i = 1; $i -le $sheets.Count; $i++) {
            $sheet = $sheets.Item($i)
            $range = $sheet.UsedRange
            $values = $range.Value2
            if ($values -is [System.Array]) {
                foreach ($value in $values) { if ([string]$value -ceq "Pendiente") { $count++ } }
            } elseif ([string]$values -ceq "Pendiente") {
                $count++
            }
            Release-Com $range
            Release-Com $sheet
        }
    } finally { Release-Com $sheets }
    return $count
}

function Count-ErrorCells($Workbook) {
    $count = 0
    $sheets = $Workbook.Worksheets
    try {
        for ($i = 1; $i -le $sheets.Count; $i++) {
            $sheet = $sheets.Item($i)
            $used = $sheet.UsedRange
            foreach ($kind in -4123, 2) {
                try {
                    $errors = $used.SpecialCells($kind, 16)
                    $count += $errors.Count
                    Release-Com $errors
                } catch { }
            }
            Release-Com $used
            Release-Com $sheet
        }
    } finally { Release-Com $sheets }
    return $count
}

function Get-CellValue($Workbook, [string]$SheetName, [string]$Address) {
    $sheet = $Workbook.Worksheets.Item($SheetName)
    $range = $sheet.Range($Address)
    $value = $range.Value2
    Release-Com $range
    Release-Com $sheet
    return $value
}

function Test-Sheet($Workbook, [string]$Name) {
    try {
        $sheet = $Workbook.Worksheets.Item($Name)
        Release-Com $sheet
        return $true
    } catch { return $false }
}

function Wait-NoExcel([int]$Seconds = 15) {
    $deadline = (Get-Date).AddSeconds($Seconds)
    do {
        $processes = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
        if ($processes.Count -eq 0) { return 0 }
        Start-Sleep -Milliseconds 500
    } while ((Get-Date) -lt $deadline)
    return @(Get-Process EXCEL -ErrorAction SilentlyContinue).Count
}

function Ensure-NoExcel {
    $remaining = Wait-NoExcel 10
    $terminated = @()
    if ($remaining -gt 0) {
        foreach ($process in @(Get-Process EXCEL -ErrorAction SilentlyContinue)) {
            $terminated += $process.Id
            Stop-Process -Id $process.Id -Force
        }
        Start-Sleep -Seconds 2
    }
    return [pscustomobject]@{
        Remaining = @(Get-Process EXCEL -ErrorAction SilentlyContinue).Count
        FallbackTerminatedPids = $terminated
    }
}

$beforeProcesses = @(Get-Process EXCEL -ErrorAction SilentlyContinue).Count
if ($beforeProcesses -ne 0) { throw "Hay procesos EXCEL.EXE antes de la validación: $beforeProcesses" }

$excel = $null
$workbook = $null
$firstOpen = $null
$reopen = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    $workbook = $excel.Workbooks.Open($Path, 0, $false)
    $excel.Calculation = -4105
    $excel.CalculateFullRebuild()

    $links = $workbook.LinkSources(1)
    $linkCount = if ($null -eq $links) { 0 } else { @($links).Count }
    $worksheets = $workbook.Worksheets
    $sheetCount = $worksheets.Count
    Release-Com $worksheets
    $profileValues = @(Get-CellValue $workbook "09_Resultados_Custom" "F5:F8")
    $firstOpen = [ordered]@{
        Calculation = $excel.Calculation
        Sheets = $sheetCount
        Charts = Count-Charts $workbook
        Pending = Count-Pending $workbook
        ErrorCells = Count-ErrorCells $workbook
        ExternalLinks = $linkCount
        ResultsSheetPresent = Test-Sheet $workbook "09_Resultados_Custom"
        BenchmarkSheetPresent = Test-Sheet $workbook "09_Benchmark_Plan"
        CustomTotal = Get-CellValue $workbook "09_Resultados_Custom" "F9"
        ProfileSum = ($profileValues | Measure-Object -Sum).Sum
        WazuhBase = Get-CellValue $workbook "01_Dashboard" "T5"
        WazuhCustom = Get-CellValue $workbook "01_Dashboard" "T6"
        PublicCH = Get-CellValue $workbook "01_Dashboard" "T4"
    }
    if ($firstOpen.Sheets -ne 30 -or $firstOpen.Charts -ne 11 -or $firstOpen.Pending -ne 0 -or $firstOpen.ErrorCells -ne 0 -or $firstOpen.ExternalLinks -ne 0) {
        throw "Fallo en las comprobaciones COM previas al guardado"
    }
    if ($firstOpen.CustomTotal -ne 379 -or $firstOpen.ProfileSum -ne 379) { throw "Total custom distinto de 379" }
    if ($firstOpen.WazuhBase -ne 3 -or $firstOpen.WazuhCustom -ne 4 -or $firstOpen.PublicCH -ne 3) { throw "Cobertura comparativa incorrecta" }

    $workbook.Save()
    $workbook.Close($true)
    Release-Com $workbook
    $workbook = $null
    $excel.Quit()
    Release-Com $excel
    $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    $afterSaveCleanup = Ensure-NoExcel

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    $workbook = $excel.Workbooks.Open($Path, 0, $true)
    $worksheets = $workbook.Worksheets
    $reopenSheetCount = $worksheets.Count
    Release-Com $worksheets
    $reopen = [ordered]@{
        ReadOnly = $workbook.ReadOnly
        Sheets = $reopenSheetCount
        Charts = Count-Charts $workbook
        Pending = Count-Pending $workbook
        ErrorCells = Count-ErrorCells $workbook
        RepairMode = $false
        CorruptLoad = "xlNormalLoad"
        OpenedWithoutRepairPrompt = $true
        ResultsSheetPresent = Test-Sheet $workbook "09_Resultados_Custom"
        BenchmarkSheetPresent = Test-Sheet $workbook "09_Benchmark_Plan"
        CustomTotal = Get-CellValue $workbook "09_Resultados_Custom" "F9"
    }
    if (-not $reopen.ReadOnly -or $reopen.Sheets -ne 30 -or $reopen.Charts -ne 11 -or $reopen.Pending -ne 0 -or $reopen.ErrorCells -ne 0 -or $reopen.CustomTotal -ne 379) {
        throw "Fallo en la reapertura COM de solo lectura"
    }
    $workbook.Close($false)
    Release-Com $workbook
    $workbook = $null
    $excel.Quit()
    Release-Com $excel
    $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    $afterReopenCleanup = Ensure-NoExcel

    $item = Get-Item -LiteralPath $Path
    $report = [ordered]@{
        GeneratedAt = (Get-Date).ToString("o")
        Path = $Path
        Bytes = $item.Length
        SHA256 = (Get-FileHash -Algorithm SHA256 -LiteralPath $Path).Hash
        BeforeExcelProcesses = $beforeProcesses
        FirstOpen = $firstOpen
        OrphansAfterSave = $afterSaveCleanup.Remaining
        FallbackTerminatedAfterSave = $afterSaveCleanup.FallbackTerminatedPids
        ReopenReadOnly = $reopen
        OrphansAfterReopen = $afterReopenCleanup.Remaining
        FallbackTerminatedAfterReopen = $afterReopenCleanup.FallbackTerminatedPids
        Status = if ($afterSaveCleanup.Remaining -eq 0 -and $afterReopenCleanup.Remaining -eq 0) { "PASS" } else { "FAIL" }
    }
    $json = $report | ConvertTo-Json -Depth 6
    Set-Content -LiteralPath $Output -Value $json -Encoding UTF8
    $json
    if ($report.Status -ne "PASS") { exit 1 }
}
finally {
    if ($workbook -ne $null) { try { $workbook.Close($false) } catch { }; Release-Com $workbook }
    if ($excel -ne $null) { try { $excel.Quit() } catch { }; Release-Com $excel }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

