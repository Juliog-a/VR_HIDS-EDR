[CmdletBinding()]
param(
    [string]$WorkbookPath = "C:\Users\julio\Desktop\TFM\08_MEMORIA\ENTREGABLE\TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712.xlsx",
    [string]$OutputPath = "C:\Users\julio\Desktop\TFM\08_MEMORIA\AUDITORIA_EXCELES_FINAL_20260712\EXCEL_COM_VALIDACION_AJUSTE_FINAL_VISIBILIDAD_20260713.json",
    [switch]$SkipCalculateForDiagnostic
)

$ErrorActionPreference = "Stop"
$xlCalculationAutomatic = -4105
$xlCellTypeConstants = 2
$xlCellTypeFormulas = -4123
$xlErrors = 16
$xlLinkTypeExcelLinks = 1

if (-not ("TFMVisibilityValidationNative" -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class TFMVisibilityValidationNative {
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

function Get-ExactPendingCount {
    param($Workbook)
    $count = 0
    foreach ($worksheet in @($Workbook.Worksheets)) {
        $used = $null
        try {
            $used = $worksheet.UsedRange
            $values = $used.Value2
            if ($values -is [System.Array]) {
                foreach ($value in $values) {
                    if ($null -ne $value -and ([string]$value) -ceq "Pendiente") { $count++ }
                }
            } elseif ($null -ne $values -and ([string]$values) -ceq "Pendiente") {
                $count++
            }
        } finally {
            Release-ComObject $used
            Release-ComObject $worksheet
        }
    }
    return $count
}

function Add-Assertion {
    param(
        [System.Collections.Generic.List[object]]$List,
        [string]$Name,
        [bool]$Pass,
        $Actual
    )
    $List.Add([pscustomobject]@{ Name = $Name; Pass = $Pass; Actual = $Actual })
}

function Get-WorkbookInspection {
    param($Workbook)

    $sheetDetails = [System.Collections.Generic.List[object]]::new()
    $assertions = [System.Collections.Generic.List[object]]::new()
    $totalCharts = 0
    $emptyCharts = 0
    $brokenSeries = 0
    $formulaErrors = 0
    $storedErrors = 0

    foreach ($worksheet in @($Workbook.Worksheets)) {
        $used = $null
        $chartObjects = $null
        try {
            $used = $worksheet.UsedRange
            $sheetFormulaErrors = Get-SpecialCellCount -Range $used -CellType $xlCellTypeFormulas -ValueType $xlErrors
            $sheetStoredErrors = Get-SpecialCellCount -Range $used -CellType $xlCellTypeConstants -ValueType $xlErrors
            $formulaErrors += $sheetFormulaErrors
            $storedErrors += $sheetStoredErrors
            $chartObjects = $worksheet.ChartObjects()
            $chartCount = [int]$chartObjects.Count
            $totalCharts += $chartCount
            $charts = [System.Collections.Generic.List[object]]::new()

            for ($i = 1; $i -le $chartCount; $i++) {
                $chartObject = $null
                $chart = $null
                $seriesCollection = $null
                try {
                    $chartObject = $chartObjects.Item($i)
                    $chart = $chartObject.Chart
                    $seriesCollection = $chart.SeriesCollection()
                    $seriesCount = [int]$seriesCollection.Count
                    if ($seriesCount -eq 0) { $emptyCharts++ }
                    $seriesFormulas = [System.Collections.Generic.List[string]]::new()
                    for ($s = 1; $s -le $seriesCount; $s++) {
                        $series = $null
                        try {
                            $series = $seriesCollection.Item($s)
                            $formula = [string]$series.Formula
                            $seriesFormulas.Add($formula)
                            if ($formula -match "#REF!") { $brokenSeries++ }
                        } finally {
                            Release-ComObject $series
                        }
                    }
                    $charts.Add([pscustomobject]@{
                        Name = [string]$chartObject.Name
                        Title = if ($chart.HasTitle) { [string]$chart.ChartTitle.Text } else { "" }
                        SeriesCount = $seriesCount
                        SeriesFormulas = @($seriesFormulas)
                    })
                } finally {
                    Release-ComObject $seriesCollection
                    Release-ComObject $chart
                    Release-ComObject $chartObject
                }
            }

            $sheetDetails.Add([pscustomobject]@{
                Name = [string]$worksheet.Name
                FormulaErrors = $sheetFormulaErrors
                StoredErrors = $sheetStoredErrors
                ChartCount = $chartCount
                Charts = @($charts)
            })
        } finally {
            Release-ComObject $chartObjects
            Release-ComObject $used
            Release-ComObject $worksheet
        }
    }

    $pendingCount = Get-ExactPendingCount -Workbook $Workbook
    $externalLinks = 0
    try {
        $links = $Workbook.LinkSources($xlLinkTypeExcelLinks)
        if ($null -ne $links) { $externalLinks = @($links).Count }
    } catch {
        $externalLinks = 0
    }

    $repairMode = $false
    $repairModeReadable = $false
    try {
        $repairMode = [bool]$Workbook.RepairMode
        $repairModeReadable = $true
    } catch {}

    Add-Assertion $assertions "31 hojas" ([int]$Workbook.Worksheets.Count -eq 31) ([int]$Workbook.Worksheets.Count)
    Add-Assertion $assertions "11 graficos" ($totalCharts -eq 11) $totalCharts
    Add-Assertion $assertions "0 graficos vacios" ($emptyCharts -eq 0) $emptyCharts
    Add-Assertion $assertions "0 series rotas" ($brokenSeries -eq 0) $brokenSeries
    Add-Assertion $assertions "0 errores de formula" ($formulaErrors -eq 0) $formulaErrors
    Add-Assertion $assertions "0 errores almacenados" ($storedErrors -eq 0) $storedErrors
    Add-Assertion $assertions "0 enlaces externos" ($externalLinks -eq 0) $externalLinks
    Add-Assertion $assertions "0 celdas Pendiente" ($pendingCount -eq 0) $pendingCount
    Add-Assertion $assertions "Sin modo de reparacion" (-not $repairMode) $repairMode

    $dashboard = $null
    $dashboardCharts = $null
    try {
        $dashboard = $Workbook.Worksheets.Item("01_Dashboard")
        $dashboardCharts = $dashboard.ChartObjects()
        $layerChart = $null
        $specificChart = $null
        for ($i = 1; $i -le [int]$dashboardCharts.Count; $i++) {
            $co = $null
            $chart = $null
            try {
                $co = $dashboardCharts.Item($i)
                $chart = $co.Chart
                $title = if ($chart.HasTitle) { [string]$chart.ChartTitle.Text } else { "" }
                if ($title -like "*CLIENT_EVENT*custom frente a publicos*" -or $title -like "*CLIENT_EVENT*custom frente a p*blicos*") { $layerChart = $co.Name }
                if ($title -like "*deteccion especifica acreditada*" -or $title -like "*detecci*n espec*fica acreditada*") { $specificChart = $co.Name }
            } finally {
                Release-ComObject $chart
                Release-ComObject $co
            }
        }
        Add-Assertion $assertions "Grafico de capas localizado" ($null -ne $layerChart) $layerChart
        Add-Assertion $assertions "Grafico de deteccion especifica localizado" ($null -ne $specificChart) $specificChart
        Add-Assertion $assertions "Tacticas 5-6-7 sobre 14" ((@($dashboard.Range("I5:I7").Value2) -join ",") -eq "5,6,7") (@($dashboard.Range("I5:I7").Value2) -join ",")
    } finally {
        Release-ComObject $dashboardCharts
        Release-ComObject $dashboard
    }

    $graphs = $null
    $graphCharts = $null
    try {
        $graphs = $Workbook.Worksheets.Item("GRAFICAS")
        Add-Assertion $assertions "Wazuh especifico = 4" ([double]$graphs.Range("G63").Value2 -eq 4) $graphs.Range("G63").Value2
        $layerCategories = @($graphs.Range("J60:L60").Value2)
        $layerCustom = @($graphs.Range("J61:L61").Value2)
        $layerPublic = @($graphs.Range("J62:L62").Value2)
        Add-Assertion $assertions "Capas: tres categorias" ($layerCategories.Count -eq 3) ($layerCategories -join " | ")
        Add-Assertion $assertions "Capas custom 9-9-9" (($layerCustom -join ",") -eq "9,9,9") ($layerCustom -join ",")
        Add-Assertion $assertions "Capas publicos 9-3-3" (($layerPublic -join ",") -eq "9,3,3") ($layerPublic -join ",")
        $removedOutput = @($graphs.Range("M60:M62").Value2 | Where-Object { $null -ne $_ -and [string]$_ -ne "" }).Count
        Add-Assertion $assertions "Salida externa fuera de la fuente de capas" ($removedOutput -eq 0) $removedOutput
        $expectedStates = @("POSITIVO", "POSITIVO", "POSITIVO", "POSITIVO", "NO CONCLUYENTE", "POSITIVO")
        $actualStates = @($graphs.Range("G74:G79").Value2)
        Add-Assertion $assertions "Estados publicos 5 positivos y 1 no concluyente" (($actualStates -join "|") -eq ($expectedStates -join "|")) ($actualStates -join " | ")
        $actualTotals = @($graphs.Range("J74:J78").Value2)
        Add-Assertion $assertions "Totales publicos 5-0-1-0-0" (($actualTotals -join ",") -eq "5,0,1,0,0") ($actualTotals -join ",")
        $heatmapLabels = @($graphs.Range("M51:M55").Value2)
        Add-Assertion $assertions "Heatmap con cinco capacidades HIDS" ($heatmapLabels.Count -eq 5 -and (($heatmapLabels -join "|") -notlike "*Salida externa*")) ($heatmapLabels -join " | ")
        $graphCharts = $graphs.ChartObjects()
        $fpTitle = ""
        for ($i = 1; $i -le [int]$graphCharts.Count; $i++) {
            $co = $null
            $chart = $null
            try {
                $co = $graphCharts.Item($i)
                $chart = $co.Chart
                if ($chart.HasTitle -and [string]$chart.ChartTitle.Text -like "Resultado del runner FP custom:*") { $fpTitle = [string]$chart.ChartTitle.Text }
            } finally {
                Release-ComObject $chart
                Release-ComObject $co
            }
        }
        Add-Assertion $assertions "Grafico FP custom diferenciado" ($fpTitle -eq "Resultado del runner FP custom: 10/10 OK y 0 hits") $fpTitle
    } finally {
        Release-ComObject $graphCharts
        Release-ComObject $graphs
    }

    $publics = $null
    try {
        $publics = $Workbook.Worksheets.Item("15_Publicos_Definitivo")
        Add-Assertion $assertions "Columna Salida externa HTTP" ([string]$publics.Range("H4").Value2 -eq "Salida externa HTTP") $publics.Range("H4").Value2
        Add-Assertion $assertions "Discord TrackNetwork = No" ([string]$publics.Range("G10").Value2 -eq "No") $publics.Range("G10").Value2
        $httpTrack = [string]$publics.Range("H10").Value2
        Add-Assertion $assertions "HTTP atribuido a runner independiente" ($httpTrack -like "Si, acreditada por runner independiente*" -or $httpTrack -like "S*, acreditada por runner independiente*") $httpTrack
        $campaignResults = @($publics.Range("S5:S10").Value2)
        Add-Assertion $assertions "Seis CLIENT_EVENT: 5 positivos y 1 no concluyente" ((@($campaignResults | Where-Object { $_ -eq "POSITIVO" }).Count -eq 5) -and (@($campaignResults | Where-Object { $_ -eq "NO CONCLUYENTE" }).Count -eq 1)) ($campaignResults -join " | ")
        Add-Assertion $assertions "Componentes separados de las seis campanas" ([string]$publics.Range("A12").Value2 -like "Componentes hist*ricos, de transporte y salida externa*") $publics.Range("A12").Value2
        Add-Assertion $assertions "Area de impresion A1:T15" ([string]$publics.PageSetup.PrintArea -like "*A*1:*T*15") $publics.PageSetup.PrintArea
        Add-Assertion $assertions "Orientacion horizontal" ([int]$publics.PageSetup.Orientation -eq 2) $publics.PageSetup.Orientation
    } finally {
        Release-ComObject $publics
    }

    $failedAssertions = @($assertions | Where-Object { -not $_.Pass }).Count
    return [pscustomobject]@{
        SheetCount = [int]$Workbook.Worksheets.Count
        ChartCount = $totalCharts
        EmptyCharts = $emptyCharts
        BrokenSeries = $brokenSeries
        FormulaErrors = $formulaErrors
        StoredErrors = $storedErrors
        ExternalLinks = $externalLinks
        PendingExact = $pendingCount
        RepairModeReadable = $repairModeReadable
        RepairMode = $repairMode
        ReadOnly = [bool]$Workbook.ReadOnly
        Assertions = @($assertions)
        FailedAssertions = $failedAssertions
        Sheets = @($sheetDetails)
    }
}

$resolved = (Resolve-Path -LiteralPath $WorkbookPath).Path
$excel = $null
$writeWorkbook = $null
$readWorkbook = $null
$excelPid = 0
$saved = $false
$reopened = $false
$writeInspection = $null
$reopenInspection = $null
$exception = $null
$forcedProcessClose = $false

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    try { $excel.AutomationSecurity = 3 } catch {}
    [uint32]$pidValue = 0
    [void][TFMVisibilityValidationNative]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$pidValue)
    $excelPid = [int]$pidValue

    $writeWorkbook = $excel.Workbooks.Open($resolved, 0, $false)
    $excel.Calculation = $xlCalculationAutomatic
    $excel.CalculateBeforeSave = $true
    try { $writeWorkbook.ForceFullCalculation = $true } catch {}
    try { $writeWorkbook.FullCalculationOnLoad = $true } catch {}
    if (-not $SkipCalculateForDiagnostic) { $excel.CalculateFullRebuild() }
    $writeInspection = Get-WorkbookInspection -Workbook $writeWorkbook
    if ($writeInspection.FailedAssertions -ne 0) {
        $writeInspection.Assertions | Where-Object { -not $_.Pass } | Format-Table Name,Actual -AutoSize | Out-Host
        throw "Fallo de aserciones antes de guardar: $($writeInspection.FailedAssertions)"
    }
    $writeWorkbook.Save()
    $saved = $true
    $writeWorkbook.Close($false)
    Release-ComObject $writeWorkbook
    $writeWorkbook = $null

    $readWorkbook = $excel.Workbooks.Open($resolved, 0, $true)
    $reopened = $true
    $reopenInspection = Get-WorkbookInspection -Workbook $readWorkbook
    if (-not $readWorkbook.ReadOnly) { throw "La reapertura no es de solo lectura." }
    if ($reopenInspection.FailedAssertions -ne 0) {
        $reopenInspection.Assertions | Where-Object { -not $_.Pass } | Format-Table Name,Actual -AutoSize | Out-Host
        throw "Fallo de aserciones tras reapertura: $($reopenInspection.FailedAssertions)"
    }
    $readWorkbook.Close($false)
    Release-ComObject $readWorkbook
    $readWorkbook = $null
} catch {
    $exception = $_.Exception.Message
    throw
} finally {
    if ($readWorkbook) { try { $readWorkbook.Close($false) } catch {}; Release-ComObject $readWorkbook }
    if ($writeWorkbook) { try { $writeWorkbook.Close($false) } catch {}; Release-ComObject $writeWorkbook }
    if ($excel) { try { $excel.Quit() } catch {}; Release-ComObject $excel }
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 2
    if ($excelPid -gt 0 -and (Get-Process -Id $excelPid -ErrorAction SilentlyContinue)) {
        Stop-Process -Id $excelPid -Force
        $forcedProcessClose = $true
        Start-Sleep -Seconds 1
    }
}

$remainingExcelProcesses = @(Get-Process EXCEL -ErrorAction SilentlyContinue).Count
$valid = ($saved -and $reopened -and $null -eq $exception -and $writeInspection.FailedAssertions -eq 0 -and $reopenInspection.FailedAssertions -eq 0 -and $remainingExcelProcesses -eq 0)
$payload = [pscustomobject]@{
    GeneratedAt = (Get-Date).ToString("o")
    Path = $resolved
    ExcelPid = $excelPid
    CalculationModeRequested = "Automatic"
    CalculateFullRebuildExecuted = $true
    Saved = $saved
    ReopenedReadOnly = $reopened
    ForcedProcessClose = $forcedProcessClose
    RemainingExcelProcesses = $remainingExcelProcesses
    Exception = $exception
    WriteInspection = $writeInspection
    ReopenInspection = $reopenInspection
    Valid = $valid
}
$payload | ConvertTo-Json -Depth 14 | Set-Content -LiteralPath $OutputPath -Encoding utf8
$payload | Select-Object Path,Saved,ReopenedReadOnly,Valid,ForcedProcessClose,RemainingExcelProcesses,@{N="Sheets";E={$_.ReopenInspection.SheetCount}},@{N="Charts";E={$_.ReopenInspection.ChartCount}},@{N="FormulaErrors";E={$_.ReopenInspection.FormulaErrors}},@{N="StoredErrors";E={$_.ReopenInspection.StoredErrors}},@{N="PendingExact";E={$_.ReopenInspection.PendingExact}},@{N="FailedAssertions";E={$_.ReopenInspection.FailedAssertions}}

if (-not $valid) { exit 1 }
