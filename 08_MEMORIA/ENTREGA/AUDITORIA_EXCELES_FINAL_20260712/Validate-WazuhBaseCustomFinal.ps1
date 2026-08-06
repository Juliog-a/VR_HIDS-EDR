[CmdletBinding()]
param(
    [string]$WorkbookPath = 'C:\Users\julio\Desktop\TFM\08_MEMORIA\ENTREGABLE\TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712.xlsx',
    [string]$OutputPath = 'C:\Users\julio\Desktop\TFM\08_MEMORIA\AUDITORIA_EXCELES_FINAL_20260712\EXCEL_COM_VALIDACION_WAZUH_BASE_CUSTOM_20260713.json'
)

$ErrorActionPreference = 'Stop'
$xlCalculationAutomatic = -4105
$xlCellTypeConstants = 2
$xlCellTypeFormulas = -4123
$xlErrors = 16
$xlLinkTypeExcelLinks = 1

if (-not ('TFMWazuhValidationNative' -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class TFMWazuhValidationNative {
    [DllImport("user32.dll", SetLastError=true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
"@
}

function Release-ComObject {
    param([object]$Object)
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object) } catch {}
    }
}

function Add-Assertion {
    param(
        [System.Collections.Generic.List[object]]$List,
        [string]$Name,
        [bool]$Pass,
        $Actual
    )
    $List.Add([pscustomobject]@{ Name=$Name; Pass=$Pass; Actual=$Actual })
}

function Get-SpecialCellCount {
    param($Range, [int]$CellType, [int]$ValueType)
    $special = $null
    try {
        $special = $Range.SpecialCells($CellType, $ValueType)
        return [int64]$special.CountLarge
    }
    catch { return [int64]0 }
    finally { Release-ComObject $special }
}

function Get-RangeFlat {
    param($Sheet, [string]$Address)
    $range = $null
    $result = [System.Collections.Generic.List[object]]::new()
    try {
        $range = $Sheet.Range($Address)
        $values = $range.Value2
        if ($values -is [System.Array]) {
            foreach ($value in $values) { $result.Add($value) }
        }
        else { $result.Add($values) }
        return @($result)
    }
    finally { Release-ComObject $range }
}

function Get-CellText {
    param($Sheet, [string]$Address)
    $range = $null
    try {
        $range = $Sheet.Range($Address)
        if ($null -eq $range.Value2) { return '' }
        return [string]$range.Value2
    }
    finally { Release-ComObject $range }
}

function Get-ChartInfo {
    param($Sheet, [string]$ChartName)
    $chartObjects = $null
    $chartObject = $null
    $chart = $null
    $seriesCollection = $null
    $axis = $null
    try {
        $chartObjects = $Sheet.ChartObjects()
        $chartObject = $chartObjects.Item($ChartName)
        $chart = $chartObject.Chart
        $seriesCollection = $chart.SeriesCollection()
        $series = [System.Collections.Generic.List[object]]::new()
        for ($i = 1; $i -le [int]$seriesCollection.Count; $i++) {
            $item = $null
            try {
                $item = $seriesCollection.Item($i)
                $series.Add([pscustomobject]@{
                    Name = [string]$item.Name
                    Formula = [string]$item.Formula
                    HasDataLabels = [bool]$item.HasDataLabels
                })
            }
            finally { Release-ComObject $item }
        }
        $majorUnit = $null
        try { $axis = $chart.Axes(2); $majorUnit = [double]$axis.MajorUnit } catch {}
        return [pscustomobject]@{
            Name = [string]$chartObject.Name
            Title = if ($chart.HasTitle) { [string]$chart.ChartTitle.Text } else { '' }
            SeriesCount = [int]$seriesCollection.Count
            HasLegend = [bool]$chart.HasLegend
            MajorUnit = $majorUnit
            Series = @($series)
        }
    }
    finally {
        Release-ComObject $axis
        Release-ComObject $seriesCollection
        Release-ComObject $chart
        Release-ComObject $chartObject
        Release-ComObject $chartObjects
    }
}

function Get-TableRange {
    param($Sheet, [string]$TableName)
    $tables = $null
    $table = $null
    $range = $null
    try {
        $tables = $Sheet.ListObjects
        $table = $tables.Item($TableName)
        $range = $table.Range
        return [string]$range.Address()
    }
    finally {
        Release-ComObject $range
        Release-ComObject $table
        Release-ComObject $tables
    }
}

function Get-WorkbookInspection {
    param($Workbook)

    $assertions = [System.Collections.Generic.List[object]]::new()
    $sheetDetails = [System.Collections.Generic.List[object]]::new()
    $worksheets = $null
    $totalCharts = 0
    $emptyCharts = 0
    $brokenSeries = 0
    $formulaErrors = 0
    $storedErrors = 0
    $formulaRefErrors = 0
    $pendingCount = 0
    $ambiguousWazuhLabelCount = 0

    try {
        $worksheets = $Workbook.Worksheets
        for ($w = 1; $w -le [int]$worksheets.Count; $w++) {
            $worksheet = $null
            $used = $null
            $chartObjects = $null
            try {
                $worksheet = $worksheets.Item($w)
                $used = $worksheet.UsedRange
                $sheetFormulaErrors = Get-SpecialCellCount $used $xlCellTypeFormulas $xlErrors
                $sheetStoredErrors = Get-SpecialCellCount $used $xlCellTypeConstants $xlErrors
                $formulaErrors += $sheetFormulaErrors
                $storedErrors += $sheetStoredErrors

                $values = $used.Value2
                if ($values -isnot [System.Array]) { $values = @($values) }
                foreach ($value in $values) {
                    if ($null -ne $value) {
                        $text = [string]$value
                        if ($text -ceq 'Pendiente') { $pendingCount++ }
                        if ($text.IndexOf('Wazuh espec', [StringComparison]::OrdinalIgnoreCase) -ge 0) { $ambiguousWazuhLabelCount++ }
                    }
                }
                $formulas = $used.Formula
                if ($formulas -isnot [System.Array]) { $formulas = @($formulas) }
                foreach ($formula in $formulas) {
                    if ($null -ne $formula -and ([string]$formula).Contains('#REF!')) { $formulaRefErrors++ }
                }

                $chartObjects = $worksheet.ChartObjects()
                $chartCount = [int]$chartObjects.Count
                $totalCharts += $chartCount
                for ($c = 1; $c -le $chartCount; $c++) {
                    $chartObject = $null
                    $chart = $null
                    $seriesCollection = $null
                    try {
                        $chartObject = $chartObjects.Item($c)
                        $chart = $chartObject.Chart
                        $seriesCollection = $chart.SeriesCollection()
                        $seriesCount = [int]$seriesCollection.Count
                        if ($seriesCount -eq 0) { $emptyCharts++ }
                        for ($s = 1; $s -le $seriesCount; $s++) {
                            $series = $null
                            try {
                                $series = $seriesCollection.Item($s)
                                $formula = [string]$series.Formula
                                if ([string]::IsNullOrWhiteSpace($formula) -or $formula.Contains('#REF!')) { $brokenSeries++ }
                            }
                            finally { Release-ComObject $series }
                        }
                    }
                    finally {
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
                })
            }
            finally {
                Release-ComObject $chartObjects
                Release-ComObject $used
                Release-ComObject $worksheet
            }
        }
    }
    finally { Release-ComObject $worksheets }

    $externalLinks = 0
    try {
        $links = $Workbook.LinkSources($xlLinkTypeExcelLinks)
        if ($null -ne $links) { $externalLinks = @($links).Count }
    }
    catch { $externalLinks = 0 }

    $repairMode = $false
    $repairModeReadable = $false
    try { $repairMode = [bool]$Workbook.RepairMode; $repairModeReadable = $true } catch {}

    Add-Assertion $assertions '31 hojas' ([int]$Workbook.Worksheets.Count -eq 31) ([int]$Workbook.Worksheets.Count)
    Add-Assertion $assertions '11 graficos' ($totalCharts -eq 11) $totalCharts
    Add-Assertion $assertions '0 graficos vacios' ($emptyCharts -eq 0) $emptyCharts
    Add-Assertion $assertions '0 series rotas' ($brokenSeries -eq 0) $brokenSeries
    Add-Assertion $assertions '0 errores de formula' ($formulaErrors -eq 0) $formulaErrors
    Add-Assertion $assertions '0 errores almacenados' ($storedErrors -eq 0) $storedErrors
    Add-Assertion $assertions '0 referencias REF en formulas' ($formulaRefErrors -eq 0) $formulaRefErrors
    Add-Assertion $assertions '0 enlaces externos' ($externalLinks -eq 0) $externalLinks
    Add-Assertion $assertions '0 celdas Pendiente' ($pendingCount -eq 0) $pendingCount
    Add-Assertion $assertions '0 etiquetas Wazuh especifico ambiguas' ($ambiguousWazuhLabelCount -eq 0) $ambiguousWazuhLabelCount
    Add-Assertion $assertions 'RepairMode false' (-not $repairMode) $repairMode

    $dashboard = $null
    $graphs = $null
    $matrix = $null
    $comparison = $null
    $detail = $null
    $publics = $null
    try {
        $dashboard = $Workbook.Worksheets.Item('01_Dashboard')
        $graphs = $Workbook.Worksheets.Item('GRAFICAS')
        $matrix = $Workbook.Worksheets.Item('05_Matriz_Resultados')
        $comparison = $Workbook.Worksheets.Item('COMPARACION_VR_WAZUH')
        $detail = $Workbook.Worksheets.Item('WAZUH_DETALLE')
        $publics = $Workbook.Worksheets.Item('15_Publicos_Definitivo')

        $summaryChart = Get-ChartInfo $dashboard 'TFM_Final_Compare_Detection'
        $coverageChart = Get-ChartInfo $dashboard 'TFM_Final_Public_TEC'
        $layerChart = Get-ChartInfo $dashboard 'TFM_Final_Layers'
        $wazuhChart = Get-ChartInfo $graphs 'Chart 4'
        $fpChart = Get-ChartInfo $graphs 'Chart 3'

        $summaryFormula = [string]$summaryChart.Series[0].Formula
        Add-Assertion $assertions 'Dashboard: cuatro barras 9-3-3-4' (
            $summaryChart.SeriesCount -eq 1 -and
            $summaryFormula -like '*GRAFICAS!$F$61:$F$64*' -and
            $summaryFormula -like '*GRAFICAS!$G$61:$G$64*' -and
            $summaryChart.Series[0].HasDataLabels
        ) $summaryFormula
        Add-Assertion $assertions 'Dashboard: cobertura por TEC con cuatro series' ($coverageChart.SeriesCount -eq 4) $coverageChart.SeriesCount
        Add-Assertion $assertions 'Dashboard: capas sin salida externa' (
            $layerChart.SeriesCount -eq 2 -and
            (($layerChart.Series | ForEach-Object Formula) -join '|') -notlike '*$M$60*'
        ) (($layerChart.Series | ForEach-Object Formula) -join ' | ')

        $sourceSystems = Get-RangeFlat $graphs 'F61:F64'
        $sourceValues = Get-RangeFlat $graphs 'G61:G64'
        Add-Assertion $assertions 'Fuente resumen: cuatro sistemas separados' (($sourceSystems -join '|') -eq 'VR custom|Publicos CH|Wazuh base|Wazuh custom' -or ($sourceSystems -join '|') -like 'VR custom|P*blicos CH|Wazuh base|Wazuh custom') ($sourceSystems -join ' | ')
        Add-Assertion $assertions 'Fuente resumen: 9-3-3-4' (($sourceValues -join ',') -eq '9,3,3,4') ($sourceValues -join ',')
        Add-Assertion $assertions 'Wazuh base = 3 TEC' ([double](Get-CellText $graphs 'G63') -eq 3) (Get-CellText $graphs 'G63')
        Add-Assertion $assertions 'Wazuh custom = 4 TEC' ([double](Get-CellText $graphs 'G64') -eq 4) (Get-CellText $graphs 'G64')

        $baseVector = Get-RangeFlat $graphs 'D61:D69'
        $customVector = Get-RangeFlat $graphs 'E61:E69'
        Add-Assertion $assertions 'Vector Wazuh base TEC-002-004-006' (($baseVector -join ',') -eq '0,1,0,1,0,1,0,0,0') ($baseVector -join ',')
        Add-Assertion $assertions 'Vector Wazuh custom TEC-001-005-007-008' (($customVector -join ',') -eq '1,0,0,0,1,0,1,1,0') ($customVector -join ',')
        Add-Assertion $assertions 'Grafica Wazuh base/custom con dos series' ($wazuhChart.SeriesCount -eq 2 -and $wazuhChart.HasLegend) ([pscustomobject]@{Series=$wazuhChart.SeriesCount;Legend=$wazuhChart.HasLegend;Names=($wazuhChart.Series.Name -join '|')})
        Add-Assertion $assertions 'Escala binaria 0/1 legible' ([math]::Abs([double]$wazuhChart.MajorUnit - 1) -lt 0.0001 -and [math]::Abs([double]$coverageChart.MajorUnit - 1) -lt 0.0001) ([pscustomobject]@{Graficas=$wazuhChart.MajorUnit;Dashboard=$coverageChart.MajorUnit})

        $heatmapHeaders = Get-RangeFlat $graphs 'N50:Q50'
        $heatmapCapabilities = Get-RangeFlat $graphs 'M51:M55'
        Add-Assertion $assertions 'Heatmap: cuatro variantes separadas' (($heatmapHeaders -join '|') -like 'VR custom|P*blicos CH|Wazuh base|Wazuh custom') ($heatmapHeaders -join ' | ')
        Add-Assertion $assertions 'Heatmap: cinco capacidades sin salida externa' ($heatmapCapabilities.Count -eq 5 -and ($heatmapCapabilities -join '|') -notlike '*Salida externa*') ($heatmapCapabilities -join ' | ')
        Add-Assertion $assertions 'Runner FP custom diferenciado' ($fpChart.Title -eq 'Resultado del runner FP custom: 10/10 OK y 0 hits') $fpChart.Title

        $matrixTableRange = Get-TableRange $matrix 'MatrizResultadosTable'
        $matrixBase = Get-RangeFlat $matrix 'W5:W13'
        $matrixCustom = Get-RangeFlat $matrix 'Y5:Y13'
        Add-Assertion $assertions 'Matriz ampliada A4:Z13' ($matrixTableRange -like '*$A$4:$Z$13*') $matrixTableRange
        Add-Assertion $assertions 'Matriz: cabeceras Wazuh base/custom' ((Get-CellText $matrix 'W4') -eq 'Wazuh base' -and (Get-CellText $matrix 'Y4') -eq 'Wazuh custom') ([pscustomobject]@{Base=(Get-CellText $matrix 'W4');Custom=(Get-CellText $matrix 'Y4')})
        Add-Assertion $assertions 'Matriz: base 3 detectadas' (@($matrixBase | Where-Object { $_ -eq 'DETECTADA' }).Count -eq 3) ($matrixBase -join ' | ')
        Add-Assertion $assertions 'Matriz: custom 4 detectadas' (@($matrixCustom | Where-Object { $_ -eq 'DETECTADA' }).Count -eq 4) ($matrixCustom -join ' | ')

        $comparisonDetected = Get-RangeFlat $comparison 'C11:C14'
        $comparisonAlerts = Get-RangeFlat $comparison 'D11:D14'
        Add-Assertion $assertions 'Comparacion: deteccion 9/9-3/9-3/9-4/9' (($comparisonDetected -join '|') -eq '9/9|3/9|3/9|4/9') ($comparisonDetected -join ' | ')
        Add-Assertion $assertions 'Comparacion: alerta RT 9/9-3/9-3/9-4/9' (($comparisonAlerts -join '|') -eq '9/9|3/9|3/9|4/9') ($comparisonAlerts -join ' | ')
        Add-Assertion $assertions 'Union Wazuh solo complementaria 7/9' ((Get-CellText $comparison 'B20') -eq '7/9') (Get-CellText $comparison 'B20')
        Add-Assertion $assertions 'Dashboard KPI Wazuh base/custom' ((Get-CellText $dashboard 'B14') -eq '4/9' -and (Get-CellText $dashboard 'B16') -eq '3/9') ([pscustomobject]@{Custom=(Get-CellText $dashboard 'B14');Base=(Get-CellText $dashboard 'B16')})

        $ruleIds = Get-RangeFlat $detail 'A5:A16'
        $ruleIdText = @($ruleIds | ForEach-Object { [string]$_ })
        $missingBaseRules = @('92032','92302','91835','92077' | Where-Object { $ruleIdText -notcontains $_ }).Count
        $missingCustomRules = @('110301','110402','110203','110202' | Where-Object { $ruleIdText -notcontains $_ }).Count
        Add-Assertion $assertions 'Rules base trazadas' ($missingBaseRules -eq 0) ($ruleIdText -join ',')
        Add-Assertion $assertions 'Rules custom trazadas' ($missingCustomRules -eq 0) ($ruleIdText -join ',')

        $states = Get-RangeFlat $graphs 'G74:G79'
        $totals = Get-RangeFlat $graphs 'J74:J78'
        Add-Assertion $assertions 'Campanas publicas: cinco positivas y una no concluyente' (($states -join '|') -eq 'POSITIVO|POSITIVO|POSITIVO|POSITIVO|NO CONCLUYENTE|POSITIVO') ($states -join ' | ')
        Add-Assertion $assertions 'Totales publicos 5-0-1-0-0' (($totals -join ',') -eq '5,0,1,0,0') ($totals -join ',')
        Add-Assertion $assertions 'Discord TrackNetwork = No' ((Get-CellText $publics 'G10') -eq 'No') (Get-CellText $publics 'G10')
        Add-Assertion $assertions 'HTTP separado y runner independiente' ((Get-CellText $publics 'H10') -like '*runner independiente*') (Get-CellText $publics 'H10')
        Add-Assertion $assertions 'SERVER_EVENT y Router separados' ((Get-CellText $publics 'A12') -like 'Componentes hist*ricos, de transporte y salida externa*') (Get-CellText $publics 'A12')
    }
    finally {
        Release-ComObject $publics
        Release-ComObject $detail
        Release-ComObject $comparison
        Release-ComObject $matrix
        Release-ComObject $graphs
        Release-ComObject $dashboard
    }

    $failedAssertions = @($assertions | Where-Object { -not $_.Pass }).Count
    return [pscustomobject]@{
        SheetCount = [int]$Workbook.Worksheets.Count
        ChartCount = $totalCharts
        EmptyCharts = $emptyCharts
        BrokenSeries = $brokenSeries
        FormulaErrors = $formulaErrors
        StoredErrors = $storedErrors
        FormulaRefErrors = $formulaRefErrors
        ExternalLinks = $externalLinks
        PendingExact = $pendingCount
        AmbiguousWazuhLabelCount = $ambiguousWazuhLabelCount
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
$workbooks = $null
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
    [void][TFMWazuhValidationNative]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$pidValue)
    $excelPid = [int]$pidValue
    $workbooks = $excel.Workbooks

    $writeWorkbook = $workbooks.Open($resolved, 0, $false)
    $excel.Calculation = $xlCalculationAutomatic
    $excel.CalculateBeforeSave = $true
    try { $writeWorkbook.ForceFullCalculation = $true } catch {}
    try { $writeWorkbook.FullCalculationOnLoad = $true } catch {}
    $excel.CalculateFullRebuild()
    $writeInspection = Get-WorkbookInspection $writeWorkbook
    if ($writeInspection.FailedAssertions -ne 0) {
        $failed = $writeInspection.Assertions | Where-Object { -not $_.Pass }
        throw "Aserciones fallidas antes de guardar: $($failed.Name -join '; ')"
    }
    $writeWorkbook.Save()
    $saved = $true
    $writeWorkbook.Close($false)
    Release-ComObject $writeWorkbook
    $writeWorkbook = $null

    $readWorkbook = $workbooks.Open($resolved, 0, $true)
    $reopened = $true
    $reopenInspection = Get-WorkbookInspection $readWorkbook
    if (-not $readWorkbook.ReadOnly) { throw 'La reapertura no es de solo lectura.' }
    if ($reopenInspection.FailedAssertions -ne 0) {
        $failed = $reopenInspection.Assertions | Where-Object { -not $_.Pass }
        throw "Aserciones fallidas tras reapertura: $($failed.Name -join '; ')"
    }
    $readWorkbook.Close($false)
    Release-ComObject $readWorkbook
    $readWorkbook = $null
}
catch { $exception = $_.Exception.Message }
finally {
    if ($null -ne $readWorkbook) { try { $readWorkbook.Close($false) } catch {}; Release-ComObject $readWorkbook }
    if ($null -ne $writeWorkbook) { try { $writeWorkbook.Close($false) } catch {}; Release-ComObject $writeWorkbook }
    Release-ComObject $workbooks
    if ($null -ne $excel) { try { $excel.Quit() } catch {}; Release-ComObject $excel }
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
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
$writePass = $null -ne $writeInspection -and $writeInspection.FailedAssertions -eq 0
$reopenPass = $null -ne $reopenInspection -and $reopenInspection.FailedAssertions -eq 0
$valid = $saved -and $reopened -and $null -eq $exception -and $writePass -and $reopenPass -and $remainingExcelProcesses -eq 0
$item = Get-Item -LiteralPath $resolved
$hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $resolved).Hash
$payload = [pscustomobject]@{
    GeneratedAt = (Get-Date).ToString('o')
    Path = $resolved
    Bytes = $item.Length
    SHA256 = $hash
    ExcelPid = $excelPid
    CalculationModeRequested = 'Automatic'
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
$json = $payload | ConvertTo-Json -Depth 14
[System.IO.File]::WriteAllText($OutputPath, $json, [System.Text.UTF8Encoding]::new($false))
$payload | Select-Object Path,Bytes,SHA256,Saved,ReopenedReadOnly,Valid,ForcedProcessClose,RemainingExcelProcesses,@{N='Sheets';E={$_.ReopenInspection.SheetCount}},@{N='Charts';E={$_.ReopenInspection.ChartCount}},@{N='FormulaErrors';E={$_.ReopenInspection.FormulaErrors}},@{N='StoredErrors';E={$_.ReopenInspection.StoredErrors}},@{N='PendingExact';E={$_.ReopenInspection.PendingExact}},@{N='FailedAssertions';E={$_.ReopenInspection.FailedAssertions}},Exception | Format-List
if (-not $valid) { exit 1 }
