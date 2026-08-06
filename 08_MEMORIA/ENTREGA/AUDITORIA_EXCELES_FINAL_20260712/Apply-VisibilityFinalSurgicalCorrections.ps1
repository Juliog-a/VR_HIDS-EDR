[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [string]$LogPath
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$xlCalculationAutomatic = -4105
$xlCategory = 1
$xlValue = 2
$xlPrimary = 1
$xlLegendPositionBottom = -4107
$xlPasteFormats = -4122
$xlVAlignTop = -4160
$xlHAlignLeft = -4131
$xlLandscape = 2
$xlShiftToRight = -4161
$xlValues = -4163
$xlWhole = 1
$xlByRows = 1
$xlNext = 1
$xlContinuous = 1
$xlThin = 2

function U {
    param([Parameter(Mandatory = $true)][string]$Text)
    return [System.Text.RegularExpressions.Regex]::Unescape($Text)
}

function Rgb {
    param([int]$R, [int]$G, [int]$B)
    return $R + (256 * $G) + (65536 * $B)
}

function Release-ComObject {
    param([object]$Object)
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object) } catch {}
    }
}

function Set-CellValue {
    param($Worksheet, [string]$Address, $Value)
    $range = $null
    try {
        $range = $Worksheet.Range($Address)
        if ($null -eq $Value) {
            $range.ClearContents() | Out-Null
        }
        elseif ($Value -is [byte] -or $Value -is [int16] -or $Value -is [int32] -or $Value -is [int64] -or $Value -is [single] -or $Value -is [double] -or $Value -is [decimal]) {
            $range.Value2 = [double]$Value
        }
        else {
            $range.Value2 = [string]$Value
        }
    }
    finally { Release-ComObject $range }
}

function Set-RowValues {
    param($Worksheet, [int]$Row, [object[]]$Values, [int]$StartColumn = 1)
    for ($i = 0; $i -lt $Values.Count; $i++) {
        $cell = $null
        try {
            $cell = $Worksheet.Cells.Item($Row, $StartColumn + $i)
            $value = $Values[$i]
            if ($null -eq $value) {
                $cell.ClearContents() | Out-Null
            }
            elseif ($value -is [byte] -or $value -is [int16] -or $value -is [int32] -or $value -is [int64] -or $value -is [single] -or $value -is [double] -or $value -is [decimal]) {
                $cell.Value2 = [double]$value
            }
            else {
                $cell.Value2 = [string]$value
            }
        }
        finally { Release-ComObject $cell }
    }
}

function Set-RowStatus {
    param($Worksheet, [int]$Row, [string]$FpRisk, [string]$CustomNeeded, [string]$State, [string]$CampaignResult)
    Set-CellValue $Worksheet "S$Row" $FpRisk
    Set-CellValue $Worksheet "T$Row" $CustomNeeded
    Set-CellValue $Worksheet "U$Row" $State
    Set-CellValue $Worksheet "AE$Row" $CampaignResult
}

function Set-ChartPosition {
    param($Worksheet, $ChartObject, [string]$TopLeft, [string]$BottomRight)
    $from = $null
    $to = $null
    try {
        $from = $Worksheet.Range($TopLeft)
        $to = $Worksheet.Range($BottomRight)
        $ChartObject.Left = $from.Left
        $ChartObject.Top = $from.Top
        $ChartObject.Width = ($to.Left + $to.Width) - $from.Left
        $ChartObject.Height = ($to.Top + $to.Height) - $from.Top
    }
    finally {
        Release-ComObject $to
        Release-ComObject $from
    }
}

function Configure-Chart {
    param($Chart, [string]$Title, [string]$CategoryTitle, [string]$ValueTitle, [switch]$DataLabels)
    $Chart.HasTitle = $true
    $Chart.ChartTitle.Text = $Title
    try {
        $Chart.HasLegend = $true
        $Chart.Legend.Position = $xlLegendPositionBottom
    } catch {}
    $x = $null
    $y = $null
    try {
        $x = $Chart.Axes($xlCategory, $xlPrimary)
        $x.HasTitle = $true
        $x.AxisTitle.Text = $CategoryTitle
    } catch {}
    try {
        $y = $Chart.Axes($xlValue, $xlPrimary)
        $y.HasTitle = $true
        $y.AxisTitle.Text = $ValueTitle
        $y.MinimumScale = 0
        $y.TickLabels.NumberFormat = '0'
    } catch {}
    if ($DataLabels) {
        try {
            $Chart.ApplyDataLabels() | Out-Null
            $seriesCollection = $Chart.SeriesCollection()
            try {
                for ($i = 1; $i -le $seriesCollection.Count; $i++) {
                    $series = $null
                    try {
                        $series = $seriesCollection.Item($i)
                        if ($series.HasDataLabels) { $series.DataLabels().NumberFormat = '0' }
                    } finally { Release-ComObject $series }
                }
            } finally { Release-ComObject $seriesCollection }
        } catch {}
    }
    Release-ComObject $y
    Release-ComObject $x
}

function Style-NoteRange {
    param($Worksheet, [string]$Address, [string]$Text, [int]$FontSize = 9)
    $range = $null
    try {
        $range = $Worksheet.Range($Address)
        try { $range.UnMerge() } catch {}
        $range.Merge()
        $range.Value2 = $Text
        $range.WrapText = $true
        $range.HorizontalAlignment = $xlHAlignLeft
        $range.VerticalAlignment = $xlVAlignTop
        $range.Font.Size = $FontSize
        $range.Font.Italic = $true
        $range.Font.Color = Rgb 31 78 121
        $range.Interior.Color = Rgb 221 235 247
        $range.Borders.LineStyle = $xlContinuous
        $range.Borders.Weight = $xlThin
        $range.Borders.Color = Rgb 155 194 230
    }
    finally { Release-ComObject $range }
}

function Count-ExactMatches {
    param($Workbook, [string]$Needle)
    $count = 0
    $worksheets = $null
    try {
        $worksheets = $Workbook.Worksheets
        for ($i = 1; $i -le $worksheets.Count; $i++) {
            $sheet = $null
            $used = $null
            $after = $null
            $found = $null
            try {
                $sheet = $worksheets.Item($i)
                $used = $sheet.UsedRange
                $after = $used.Cells.Item($used.Cells.Count)
                $found = $used.Find($Needle, $after, $xlValues, $xlWhole, $xlByRows, $xlNext, $false)
                if ($null -eq $found) { continue }
                $firstAddress = $found.Address($false, $false)
                do {
                    $count++
                    $next = $used.Find($Needle, $found, $xlValues, $xlWhole, $xlByRows, $xlNext, $false)
                    Release-ComObject $found
                    $found = $next
                    if ($null -eq $found) { break }
                } while ($found.Address($false, $false) -ne $firstAddress)
            }
            finally {
                Release-ComObject $found
                Release-ComObject $after
                Release-ComObject $used
                Release-ComObject $sheet
            }
        }
    }
    finally { Release-ComObject $worksheets }
    return $count
}

function Count-Charts {
    param($Workbook)
    $count = 0
    $worksheets = $null
    try {
        $worksheets = $Workbook.Worksheets
        for ($i = 1; $i -le $worksheets.Count; $i++) {
            $sheet = $null
            $charts = $null
            try {
                $sheet = $worksheets.Item($i)
                $charts = $sheet.ChartObjects()
                $count += $charts.Count
            }
            finally {
                Release-ComObject $charts
                Release-ComObject $sheet
            }
        }
    }
    finally { Release-ComObject $worksheets }
    return $count
}

$excel = $null
$workbook = $null
$worksheets = $null
$beforePending = $null
$afterPending = $null
$beforeCharts = $null
$afterCharts = $null
$insertedHttpColumn = $false

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    try { $excel.AutomationSecurity = 3 } catch {}

    $workbook = $excel.Workbooks.Open($Path, 0, $false)
    if ($workbook.ReadOnly) { throw 'El libro se ha abierto en modo de solo lectura.' }
    $worksheets = $workbook.Worksheets
    $beforePending = Count-ExactMatches $workbook 'Pendiente'
    $beforeCharts = Count-Charts $workbook

    $dashboard = $null
    $graphs = $null
    $control = $null
    $matrix = $null
    $benign = $null
    $benchmarkPlan = $null
    $alertPlan = $null
    $lists = $null
    $public = $null
    $summary = $null
    $alertability = $null
    $hayabusa = $null
    $inconsistencies = $null
    $comparison = $null
    $readme = $null

    try {
        $dashboard = $worksheets.Item('01_Dashboard')
        $graphs = $worksheets.Item('GRAFICAS')
        $control = $worksheets.Item('04_Control_Publicos')
        $matrix = $worksheets.Item('05_Matriz_Resultados')
        $benign = $worksheets.Item('08_Benignas_FP')
        $benchmarkPlan = $worksheets.Item('09_Benchmark_Plan')
        $alertPlan = $worksheets.Item('12_Plan_Alertas')
        $lists = $worksheets.Item('99_Listas')
        $public = $worksheets.Item('15_Publicos_Definitivo')
        $summary = $worksheets.Item('RESUMEN_EJECUTIVO')
        $alertability = $worksheets.Item('11_Alertabilidad')
        $hayabusa = $worksheets.Item('13_Hayabusa_Resumen')
        $inconsistencies = $worksheets.Item('17_Incoherencias_Cerradas')
        $comparison = $worksheets.Item('COMPARACION_VR_WAZUH')
        $readme = $worksheets.Item('README')

        # Dashboard: source data and chart references only.
        Set-CellValue $dashboard 'I4' (U 'Cobertura t\u00E1ctica')
        Set-CellValue $dashboard 'J4' '%'
        $dashboard.Range('I5:I7').NumberFormat = '0"/14"'
        Set-CellValue $dashboard 'I5' 5
        Set-CellValue $dashboard 'I6' 6
        Set-CellValue $dashboard 'I7' 7
        Style-NoteRange $dashboard 'K4:Q7' (U '/14 corresponde a t\u00E1cticas MITRE ATT&CK Enterprise consideradas en la clasificaci\u00F3n; no corresponde a las 9 t\u00E9cnicas TEC ejecutadas en el laboratorio.') 10

        Set-CellValue $graphs 'F60' 'Sistema'
        Set-CellValue $graphs 'G60' (U 'TEC con detecci\u00F3n espec\u00EDfica')
        Set-CellValue $graphs 'F61' 'Custom CLIENT_EVENT'
        Set-CellValue $graphs 'F62' (U 'P\u00FAblicos CH')
        Set-CellValue $graphs 'F63' (U 'Wazuh espec\u00EDfico')
        $graphs.Range('G61').Formula = '=SUM(P61:P69)'
        $graphs.Range('G62').Formula = '=SUM(Q61:Q69)'
        $graphs.Range('G63').Formula = '=SUM(R61:R69)'

        Set-RowValues $graphs 60 @('Sistema', (U 'Visibilidad'), (U 'Detecci\u00F3n'), 'Alerta RT') 9
        Set-RowValues $graphs 61 @('Custom', 9, 9, 9) 9
        Set-RowValues $graphs 62 @((U 'P\u00FAblicos'), 9, 3, 3) 9
        $graphs.Range('M60:M62').ClearContents() | Out-Null

        $detectionChartObject = $null
        $detectionChart = $null
        $detectionSeriesCollection = $null
        $detectionSeries = $null
        try {
            $detectionChartObject = $dashboard.ChartObjects('TFM_Final_Compare_Detection')
            $detectionChart = $detectionChartObject.Chart
            $detectionSeriesCollection = $detectionChart.SeriesCollection()
            while ($detectionSeriesCollection.Count -gt 1) { $detectionSeriesCollection.Item($detectionSeriesCollection.Count).Delete() }
            if ($detectionSeriesCollection.Count -eq 0) { $detectionSeries = $detectionSeriesCollection.NewSeries() } else { $detectionSeries = $detectionSeriesCollection.Item(1) }
            $detectionSeries.Name = '=GRAFICAS!$G$60'
            $detectionSeries.XValues = '=GRAFICAS!$F$61:$F$63'
            $detectionSeries.Values = '=GRAFICAS!$G$61:$G$63'
            Configure-Chart $detectionChart (U 'T\u00E9cnicas con detecci\u00F3n espec\u00EDfica acreditada') 'Sistema' (U 'N.\u00BA de t\u00E9cnicas TEC') -DataLabels
        }
        finally {
            Release-ComObject $detectionSeries
            Release-ComObject $detectionSeriesCollection
            Release-ComObject $detectionChart
            Release-ComObject $detectionChartObject
        }

        $layersChartObject = $null
        $layersChart = $null
        $layersSeriesCollection = $null
        try {
            $layersChartObject = $dashboard.ChartObjects('TFM_Final_Layers')
            $layersChart = $layersChartObject.Chart
            $layersSeriesCollection = $layersChart.SeriesCollection()
            while ($layersSeriesCollection.Count -gt 2) { $layersSeriesCollection.Item($layersSeriesCollection.Count).Delete() }
            while ($layersSeriesCollection.Count -lt 2) { [void]$layersSeriesCollection.NewSeries() }
            $customSeries = $null
            $publicSeries = $null
            try {
                $customSeries = $layersSeriesCollection.Item(1)
                $customSeries.Name = '=GRAFICAS!$I$61'
                $customSeries.XValues = '=GRAFICAS!$J$60:$L$60'
                $customSeries.Values = '=GRAFICAS!$J$61:$L$61'
                $publicSeries = $layersSeriesCollection.Item(2)
                $publicSeries.Name = '=GRAFICAS!$I$62'
                $publicSeries.XValues = '=GRAFICAS!$J$60:$L$60'
                $publicSeries.Values = '=GRAFICAS!$J$62:$L$62'
            }
            finally {
                Release-ComObject $publicSeries
                Release-ComObject $customSeries
            }
            Configure-Chart $layersChart (U 'Capas acreditadas en CLIENT_EVENT: custom frente a p\u00FAblicos') 'Capa' (U 'N.\u00BA de t\u00E9cnicas TEC') -DataLabels
            Set-ChartPosition $dashboard $layersChartObject 'A36' 'H50'
        }
        finally {
            Release-ComObject $layersSeriesCollection
            Release-ComObject $layersChart
            Release-ComObject $layersChartObject
        }
        Style-NoteRange $dashboard 'A51:H53' (U 'Salida externa evaluada separadamente mediante SERVER_EVENT, Router, Discord, webhook o runner; no se considera una capacidad directa del CLIENT_EVENT.') 9

        # Public campaign state mapping and dependent chart.
        $graphs.Range('F73:J79').ClearContents() | Out-Null
        Set-RowValues $graphs 73 @('Artifact', 'Estado') 6
        $publicStates = @(
            @('Hayabusa Monitoring CH', 'POSITIVO'),
            @('ServiceCreation', 'POSITIVO'),
            @('ProcessCreation', 'POSITIVO'),
            @('SysmonLogForward', 'POSITIVO'),
            @('ETW Monitoring', 'NO CONCLUYENTE'),
            @('TrackNetworkConnections', 'POSITIVO')
        )
        for ($i = 0; $i -lt $publicStates.Count; $i++) { Set-RowValues $graphs (74 + $i) $publicStates[$i] 6 }
        Set-RowValues $graphs 73 @('Resultado', (U 'Campa\u00F1as')) 9
        Set-RowValues $graphs 74 @('POSITIVO', $null) 9
        Set-RowValues $graphs 75 @((U 'NEGATIVO V\u00C1LIDO'), $null) 9
        Set-RowValues $graphs 76 @('NO CONCLUYENTE', $null) 9
        Set-RowValues $graphs 77 @('NO EJECUTADO', $null) 9
        Set-RowValues $graphs 78 @('NO APLICABLE', $null) 9
        $graphs.Range('J74').Formula = '=COUNTIF(''15_Publicos_Definitivo''!$S$5:$S$10,"POSITIVO")'
        $graphs.Range('J75').Formula = U '=COUNTIF(''15_Publicos_Definitivo''!$S$5:$S$10,"NEGATIVO V\u00C1LIDO")'
        $graphs.Range('J76').Formula = '=COUNTIF(''15_Publicos_Definitivo''!$S$5:$S$10,"NO CONCLUYENTE")'
        $graphs.Range('J77').Formula = '=COUNTIF(''15_Publicos_Definitivo''!$S$5:$S$10,"NO EJECUTADO")'
        $graphs.Range('J78').Formula = '=COUNTIF(''15_Publicos_Definitivo''!$S$5:$S$10,"NO APLICABLE")'
        $graphs.Range('F73:G73').Copy() | Out-Null
        $graphs.Range('I73:J73').PasteSpecial($xlPasteFormats) | Out-Null
        try { $excel.CutCopyMode = $false } catch {}
        $outcomesObject = $null
        $outcomesChart = $null
        $outcomesSeriesCollection = $null
        $outcomesSeries = $null
        try {
            $outcomesObject = $dashboard.ChartObjects('TFM_Final_Outcomes')
            $outcomesChart = $outcomesObject.Chart
            $outcomesSeriesCollection = $outcomesChart.SeriesCollection()
            while ($outcomesSeriesCollection.Count -gt 1) { $outcomesSeriesCollection.Item($outcomesSeriesCollection.Count).Delete() }
            if ($outcomesSeriesCollection.Count -eq 0) { $outcomesSeries = $outcomesSeriesCollection.NewSeries() } else { $outcomesSeries = $outcomesSeriesCollection.Item(1) }
            $outcomesSeries.Name = '=GRAFICAS!$J$73'
            $outcomesSeries.XValues = '=GRAFICAS!$I$74:$I$78'
            $outcomesSeries.Values = '=GRAFICAS!$J$74:$J$78'
            Configure-Chart $outcomesChart (U 'Estado de las seis campa\u00F1as p\u00FAblicas') 'Resultado' (U 'Campa\u00F1as') -DataLabels
        }
        finally {
            Release-ComObject $outcomesSeries
            Release-ComObject $outcomesSeriesCollection
            Release-ComObject $outcomesChart
            Release-ComObject $outcomesObject
        }

        # Heatmap: preserve existing scores, only reclassify and remove external-output rows.
        $fixedHeatmapTitle = U 'Capacidades HIDS acreditadas por sistema'
        if ([string]$graphs.Range('M49').Value2 -ne $fixedHeatmapTitle) {
            $visibilityScores = @($graphs.Range('N51').Value2, $graphs.Range('O51').Value2, $graphs.Range('P51').Value2)
            $detectionScores = @($graphs.Range('N52').Value2, $graphs.Range('O52').Value2, $graphs.Range('P52').Value2)
            $forensicScores = @($graphs.Range('N53').Value2, $graphs.Range('O53').Value2, $graphs.Range('P53').Value2)
            $alertScores = @($graphs.Range('N54').Value2, $graphs.Range('O54').Value2, $graphs.Range('P54').Value2)
            $noiseScores = @($graphs.Range('N57').Value2, $graphs.Range('O57').Value2, $graphs.Range('P57').Value2)
            Set-RowValues $graphs 51 (@((U 'Visibilidad')) + $visibilityScores) 13
            Set-RowValues $graphs 52 (@((U 'Detecci\u00F3n')) + $detectionScores) 13
            Set-RowValues $graphs 53 (@('Alerta RT') + $alertScores) 13
            Set-RowValues $graphs 54 (@((U 'Evidencia forense')) + $forensicScores) 13
            Set-RowValues $graphs 55 (@('Ruido/FP') + $noiseScores) 13
        }
        Set-CellValue $graphs 'M49' $fixedHeatmapTitle
        Set-RowValues $graphs 50 @('Capacidad', 'VR custom', 'Wazuh base', 'Wazuh custom') 13
        $graphs.Range('M56:P57').ClearContents() | Out-Null
        $graphs.Range('M56:P57').Interior.Pattern = -4142
        $graphs.Range('M56:P57').Borders.LineStyle = -4142
        Style-NoteRange $graphs 'M58:P59' (U 'La salida externa se eval\u00FAa separadamente porque depende de componentes server-side, Router, webhook o runners auxiliares.') 9

        # FP runner chart and explicit separation from public Hayabusa FP.
        $fpChartObject = $null
        $fpChart = $null
        try {
            $fpChartObject = $graphs.ChartObjects('Chart 3')
            $fpChart = $fpChartObject.Chart
            $fpChart.HasTitle = $true
            $fpChart.ChartTitle.Text = 'Resultado del runner FP custom: 10/10 OK y 0 hits'
        }
        finally {
            Release-ComObject $fpChart
            Release-ComObject $fpChartObject
        }
        Style-NoteRange $graphs 'A25:E30' (U 'Este gr\u00E1fico representa exclusivamente el runner FP custom y no incluye las coincidencias producidas por Hayabusa Monitoring. Datos p\u00FAblicos separados: 4 coincidencias High, 3 eventos subyacentes y 1 caso benigno afectado (FP-003).') 9

        # 04_Control_Publicos: semantic closure of S/T/U/AE, never a generic replacement.
        $low = U 'Bajo'
        $high = U 'Alto'
        $indeterminate = U 'Indeterminado'
        $yes = U 'S\u00ED'
        $partial = U 'Parcial'
        $notEvaluable = U 'No evaluable'
        $outsideCampaign = U 'FUERA DE CAMPA\u00D1A'
        for ($r = 5; $r -le 19; $r++) { Set-RowStatus $control $r $low $yes 'CERRADO' 'POSITIVO' }
        for ($r = 20; $r -le 28; $r++) { Set-RowStatus $control $r $high $notEvaluable $outsideCampaign $outsideCampaign }
        for ($r = 29; $r -le 37; $r++) { Set-RowStatus $control $r $indeterminate $notEvaluable 'NO EJECUTADO' 'NO EJECUTADO' }
        for ($r = 38; $r -le 46; $r++) { Set-RowStatus $control $r $high $partial 'CERRADO' 'POSITIVO' }
        for ($r = 47; $r -le 55; $r++) { Set-RowStatus $control $r $indeterminate $notEvaluable 'NO EJECUTADO' 'NO EJECUTADO' }
        for ($r = 56; $r -le 61; $r++) { Set-RowStatus $control $r $indeterminate $notEvaluable $outsideCampaign $outsideCampaign }
        Set-RowStatus $control 62 $low $yes 'CERRADO' 'POSITIVO'
        Set-RowStatus $control 63 $low $yes 'CERRADO' 'POSITIVO'
        Set-RowStatus $control 64 $low $yes 'CERRADO' 'POSITIVO'
        Set-RowStatus $control 65 $high $partial 'CERRADO' 'POSITIVO'
        Set-RowStatus $control 66 $indeterminate $notEvaluable 'NO CONCLUYENTE' 'NO CONCLUYENTE'
        Set-RowStatus $control 67 $low $yes 'CERRADO' 'POSITIVO'
        $control.ListObjects.Item('ControlPublicosTable').Resize($control.Range('A4:AE67'))

        # 05_Matriz_Resultados: mark historical/non-executed fields explicitly.
        $matrix.Range('G5:G13').Value2 = U 'HIST\u00D3RICO; NO RT.'
        $matrix.Range('H5:H13').Value2 = 'NO EJECUTADA.'
        $matrix.Range('J5:J13').Value2 = 'NO EJECUTADA.'
        $matrix.Range('K7:K9').Value2 = U 'NO EJECUTADO EN CAMPA\u00D1A FINAL'

        # 08_Benignas_FP: historical OB rows were not individually executed.
        $benign.Range('G5:G12').Value2 = U 'SUSTITUIDO POR CAMPA\u00D1A FP FINAL'
        $benign.Range('H5:H12').Value2 = 'NO EVALUADO INDIVIDUALMENTE'
        $benign.Columns('H').ColumnWidth = 28

        # Additional exact Pending cells found in auxiliary sheets.
        $benchmarkPlan.Range('I5:I13').Value2 = U 'NO EJECUTADO EN CAMPA\u00D1A FINAL'
        Set-CellValue $alertPlan 'G11' 'EVALUADO SEPARADAMENTE; NO ES CAPACIDAD CLIENT_EVENT'
        Set-CellValue $lists 'A2' 'NO EVALUADO'
        Set-CellValue $lists 'B2' 'NO EJECUTADO'
        Set-CellValue $lists 'G5' 'No evaluable'

        # 15_Publicos_Definitivo: add HTTP beside Discord and separate non-detectors.
        $httpHeader = U 'Salida externa HTTP'
        if ([string]$public.Range('H4').Value2 -ne $httpHeader) {
            $public.Columns.Item(8).Insert($xlShiftToRight)
            $insertedHttpColumn = $true
        }
        $public.Range('A1:T15').UnMerge()
        $public.Range('A1:T1').Merge()
        $public.Range('A2:T2').Merge()
        Set-CellValue $public 'A1' (U 'Matriz definitiva de seis campa\u00F1as CLIENT_EVENT p\u00FAblicas')
        Set-CellValue $public 'A2' (U 'Criterio: CLIENT_EVENT aporta visibilidad, detecci\u00F3n o alerta RT seg\u00FAn el artifact; SERVER_EVENT, Router y runner se eval\u00FAan separadamente como transporte o salida externa.')
        Set-CellValue $public 'H4' $httpHeader
        for ($r = 5; $r -le 10; $r++) {
            Set-CellValue $public "G$r" 'No'
            Set-CellValue $public "H$r" 'No'
        }
        Set-CellValue $public 'H10' (U 'S\u00ED, acreditada por runner independiente; no atribuible al CLIENT_EVENT')
        Set-CellValue $public 'J10' (U 'Conexi\u00F3n visible; salida HTTP acreditada aparte')
        Set-CellValue $public 'K10' 'Diff added/removed; filtro por _ts'
        Set-CellValue $public 'T10' (U 'TrackNetwork acredit\u00F3 conexi\u00F3n hacia 192.168.1.129:8088 y eventos added/removed; PID=0 y ProcInfo vac\u00EDo, sin atribuci\u00F3n fiable a powershell.exe. Runner independiente: HTTPStatus=200; UploadSucceeded=True; ZIP=6245 bytes; SHA-256 coincidente; respuesta del receptor confirmada.')

        $public.Range('A11:T15').UnMerge()
        $totalFormatRow = if ($insertedHttpColumn) { 14 } else { 11 }
        $public.Range("A$totalFormatRow`:T$totalFormatRow").Copy() | Out-Null
        $public.Range('A11:T11').PasteSpecial($xlPasteFormats) | Out-Null
        $public.Range('A13:T13').Copy() | Out-Null
        $public.Range('A13:T15').PasteSpecial($xlPasteFormats) | Out-Null
        try { $excel.CutCopyMode = $false } catch {}
        $public.Range('A11:T15').ClearContents() | Out-Null

        Set-CellValue $public 'A11' (U 'TOTAL 6 CAMPA\u00D1AS')
        $public.Range('N11').Formula = '=SUM(N5:N10)'
        $public.Range('O11').Formula = '=SUM(O5:O10)'
        $public.Range('P11').Formula = '=SUM(P5:P10)'
        $public.Range('Q11').Formula = '=SUM(Q5:Q10)'
        Set-CellValue $public 'R11' (U 'Totales de filas por campa\u00F1a; no equivalen a detecciones entre artifacts heterog\u00E9neos.')
        $public.Range('S11').Formula = U '=COUNTIF(S5:S10,"POSITIVO")&" positivos; "&COUNTIF(S5:S10,"NEGATIVO V\u00C1LIDO")&" negativos v\u00E1lidos; "&COUNTIF(S5:S10,"NO CONCLUYENTE")&" no concluyentes"'
        Set-CellValue $public 'T11' (U 'Salida externa TEC-009 acreditada por runner independiente; no atribuible a TrackNetworkConnections.')

        $public.Range('A12:T12').Merge()
        Set-CellValue $public 'A12' (U 'Componentes hist\u00F3ricos, de transporte y salida externa \u2014 no incluidos en la cobertura de los seis CLIENT_EVENT evaluados')
        $public.Range('A12:T12').Interior.Color = Rgb 31 78 121
        $public.Range('A12:T12').Font.Color = Rgb 255 255 255
        $public.Range('A12:T12').Font.Bold = $true
        $public.Range('A12:T12').HorizontalAlignment = $xlHAlignLeft
        $public.Range('A12:T12').VerticalAlignment = $xlVAlignTop
        $public.Range('A12:T12').WrapText = $true

        Set-RowValues $public 13 @(
            'Windows.Hayabusa.Rules', (U 'CLIENT / hunt hist\u00F3rico'), 'No', (U 'S\u00ED, hist\u00F3rica'), (U 'Hist\u00F3rica; no campa\u00F1a final'), 'No', 'No', 'No', (U 'No incluida'), (U 'Hunt hist\u00F3rico; no alerta RT'), (U 'No ejecutado como campa\u00F1a final independiente'), $null, $null, $null, $null, $null, $null, $null, (U 'FUERA DE CAMPA\u00D1A'), (U 'No incluido en la cobertura de los seis CLIENT_EVENT evaluados.')
        )
        Set-RowValues $public 14 @(
            'Server.Alerts.TrackNetworkConnections', 'SERVER_EVENT', 'No aplica', 'No detector', 'No detector', 'Transporte', 'No', 'No', 'No aplica', (U 'Transporte server-side; no detector'), (U 'No incluido en cobertura'), $null, $null, $null, $null, $null, $null, $null, 'NO APLICABLE', (U 'SERVER_EVENT: transporta; no detecta y no cuenta en la cobertura.')
        )
        Set-RowValues $public 15 @(
            'Custom Router Discord', 'SERVER_EVENT', 'No aplica', 'No detector', 'No detector', (U 'Normalizaci\u00F3n y salida externa'), (U 'S\u00ED'), 'No', 'No aplica', (U 'Router Discord; no detector'), (U 'No incluido en cobertura'), $null, $null, $null, $null, $null, $null, $null, 'NO APLICABLE', (U 'SERVER_EVENT: normaliza y notifica; no sustituye la detecci\u00F3n CLIENT_EVENT.')
        )

        $public.Range('A4:T15').WrapText = $true
        $public.Range('A4:T15').VerticalAlignment = $xlVAlignTop
        $public.Range('N5:Q11').NumberFormat = '0'
        $columnWidths = @{
            'A'=38; 'B'=17; 'C'=12; 'D'=21; 'E'=29; 'F'=20; 'G'=14; 'H'=34; 'I'=24; 'J'=34;
            'K'=30; 'L'=27; 'M'=27; 'N'=12; 'O'=12; 'P'=14; 'Q'=12; 'R'=46; 'S'=20; 'T'=48
        }
        foreach ($key in $columnWidths.Keys) { $public.Columns($key).ColumnWidth = $columnWidths[$key] }
        $public.Rows('5:10').RowHeight = 78
        $public.Rows(11).RowHeight = 42
        $public.Rows(12).RowHeight = 34
        $public.Rows('13:15').RowHeight = 62
        try { $public.AutoFilterMode = $false } catch {}
        $public.Range('A4:T10').AutoFilter() | Out-Null
        $public.PageSetup.Orientation = $xlLandscape
        $public.PageSetup.PrintArea = '$A$1:$T$15'
        $public.PageSetup.PrintTitleRows = '$4:$4'
        $public.PageSetup.Zoom = $false
        $public.PageSetup.FitToPagesWide = 1
        $public.PageSetup.FitToPagesTall = $false
        $public.Activate()
        $excel.ActiveWindow.SplitRow = 4
        $excel.ActiveWindow.SplitColumn = 2
        $excel.ActiveWindow.FreezePanes = $true
        $excel.ActiveWindow.Zoom = 80
        $public.Range('C5').Select() | Out-Null

        # Reassert dependent totals after inserting the HTTP column so they point to final column S.
        $graphs.Range('J74').Formula = '=COUNTIF(''15_Publicos_Definitivo''!$S$5:$S$10,"POSITIVO")'
        $graphs.Range('J75').Formula = U '=COUNTIF(''15_Publicos_Definitivo''!$S$5:$S$10,"NEGATIVO V\u00C1LIDO")'
        $graphs.Range('J76').Formula = '=COUNTIF(''15_Publicos_Definitivo''!$S$5:$S$10,"NO CONCLUYENTE")'
        $graphs.Range('J77').Formula = '=COUNTIF(''15_Publicos_Definitivo''!$S$5:$S$10,"NO EJECUTADO")'
        $graphs.Range('J78').Formula = '=COUNTIF(''15_Publicos_Definitivo''!$S$5:$S$10,"NO APLICABLE")'

        # Directly affected narrative sheets only.
        Set-CellValue $summary 'C32' (U 'Conexi\u00F3n hacia 192.168.1.129:8088; eventos added/removed; PID=0 y ProcInfo vac\u00EDo; sin atribuci\u00F3n fiable a powershell.exe.')
        Set-CellValue $summary 'C33' (U 'Runner independiente: HTTPStatus=200, UploadSucceeded=True, ZIP=6245 bytes, SHA-256 coincidente y respuesta del receptor confirmada; no atribuible a TrackNetworkConnections.')

        Set-CellValue $alertability 'A2' (U 'Separar visibilidad, detecci\u00F3n, alerta en tiempo real, evidencia forense y transporte/salida externa; CLIENT_EVENT detecta y SERVER_EVENT/Router transportan o notifican.')
        Set-RowValues $alertability 8 @(
            (U 'Transporte / salida externa'),
            (U 'SERVER_EVENT o Router reenv\u00EDa una detecci\u00F3n ya emitida; un runner independiente puede acreditar HTTP.'),
            'Server.Alerts.Notification / Router Discord / runner HTTP',
            (U 'No: no es capacidad directa del CLIENT_EVENT'),
            (U 'Se eval\u00FAa separadamente de la cobertura detectora.')
        )
        Set-CellValue $alertability 'G17' (U 'Evaluar aparte como transporte; no incluir en cobertura CLIENT_EVENT')
        Set-CellValue $alertability 'B21' (U 'S\u00ED. ProcessCreation, ServiceCreation, SysmonLogForward y TrackNetwork aportan visibilidad o contexto; Hayabusa CH aporta detecci\u00F3n 3/9. SERVER_EVENT y Router se eval\u00FAan aparte como transporte/salida externa.')
        Set-CellValue $alertability 'B22' (U 'Cuando la fuente p\u00FAblica solo da visibilidad, cuando Hayabusa detecta gen\u00E9rico/ruidoso o cuando falta una alerta resumida por TEC con CU_ID, severidad y confianza. La salida externa no forma parte del detector CLIENT_EVENT.')

        Set-CellValue $hayabusa 'B17' (U 'HIST\u00D3RICO; NO RT. Es CLIENT/hunt y no CLIENT_EVENT.')
        Set-CellValue $hayabusa 'B29' 'NO EJECUTADA'
        Set-CellValue $hayabusa 'D29' (U 'No existe campa\u00F1a CHM independiente; no se incorpora a gr\u00E1ficos, totales ni cobertura final.')
        Set-CellValue $hayabusa 'D32' (U 'Detecci\u00F3n Sigma CH 3/9 en CLIENT_EVENT; salida externa no evaluada como capacidad del artifact.')
        Set-CellValue $hayabusa 'D33' (U 'Discord/JSONL/SERVER_EVENT son transporte o salida; no sustituyen la detecci\u00F3n.')

        Set-CellValue $inconsistencies 'C19' (U 'CLIENT_EVENT aporta visibilidad/detecci\u00F3n/alerta seg\u00FAn el artifact; SERVER_EVENT y Router transportan o notifican. El runner independiente acredita HTTP 200; TrackNetwork solo acredita conexi\u00F3n visible.')
        Set-CellValue $inconsistencies 'D19' (U 'Excluir salida externa de cobertura detectora y no atribuir el HTTP 200 a TrackNetworkConnections.')

        Set-CellValue $comparison 'J8' (U 'TrackNetwork: conexi\u00F3n visible sin atribuci\u00F3n fiable; HTTPStatus=200 solo por runner independiente.')
        Set-CellValue $comparison 'M8' (U 'Comparaci\u00F3n por capacidades HIDS; SERVER_EVENT, Router, Discord y runner HTTP quedan fuera de la cobertura de los seis CLIENT_EVENT.')
        Set-CellValue $comparison 'M5' (U 'Separar visibilidad, detecci\u00F3n y alerta CLIENT_EVENT de transporte SERVER_EVENT/Router y salida externa.')

        Set-CellValue $readme 'B13' (U 'ETW no concluyente; TrackNetwork acredita conexi\u00F3n sin proceso atribuible; SERVER_EVENT/Router/Discord no son detecci\u00F3n; HTTP 200 pertenece al runner independiente.')
        Set-CellValue $readme 'B15' (U 'Evento crudo=visibilidad; coincidencia Sigma CH=detecci\u00F3n 3/9; custom=9/9; Wazuh espec\u00EDfico=4/9. Transporte y salida externa se eval\u00FAan aparte.')

        $dashboard.Activate()
        $dashboard.Range('A1').Select() | Out-Null

        $excel.Calculation = $xlCalculationAutomatic
        try { $workbook.ForceFullCalculation = $true } catch {}
        try { $workbook.FullCalculationOnLoad = $true } catch {}
        try { $workbook.Application.CalculateFullRebuild() } catch { $workbook.Application.CalculateFull() }

        $afterPending = Count-ExactMatches $workbook 'Pendiente'
        if ($afterPending -ne 0) { throw "Quedan $afterPending celdas con valor exacto Pendiente." }
        $afterCharts = Count-Charts $workbook
        if ($workbook.Worksheets.Count -ne 31) { throw "El libro ya no conserva 31 hojas." }
        if ($afterCharts -ne 11) { throw "El libro ya no conserva 11 graficos." }
        $workbook.Save()
    }
    finally {
        foreach ($object in @($readme, $comparison, $inconsistencies, $hayabusa, $alertability, $summary, $public, $lists, $alertPlan, $benchmarkPlan, $benign, $matrix, $control, $graphs, $dashboard)) {
            Release-ComObject $object
        }
    }

    $result = [ordered]@{
        Path = $Path
        PendingBefore = $beforePending
        PendingAfter = $afterPending
        Worksheets = $workbook.Worksheets.Count
        ChartsBefore = $beforeCharts
        ChartsAfter = $afterCharts
        InsertedHttpColumn = $insertedHttpColumn
        Calculation = 'Automatic + CalculateFullRebuild'
        Saved = $true
        ChangedSheets = @(
            '01_Dashboard', 'GRAFICAS', '04_Control_Publicos', '05_Matriz_Resultados', '08_Benignas_FP',
            '09_Benchmark_Plan', '12_Plan_Alertas', '99_Listas', '15_Publicos_Definitivo',
            'RESUMEN_EJECUTIVO', '11_Alertabilidad', '13_Hayabusa_Resumen',
            '17_Incoherencias_Cerradas', 'COMPARACION_VR_WAZUH', 'README'
        )
    }
    $json = [pscustomobject]$result | ConvertTo-Json -Depth 6
    if (-not [string]::IsNullOrWhiteSpace($LogPath)) {
        [System.IO.File]::WriteAllText($LogPath, $json, [System.Text.UTF8Encoding]::new($false))
    }
    $json
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
    }
    Release-ComObject $worksheets
    Release-ComObject $workbook
    Release-ComObject $excel
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
}
