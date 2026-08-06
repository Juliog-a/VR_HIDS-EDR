[CmdletBinding()]
param(
    [string]$Path = "C:\Users\julio\Desktop\TFM\08_MEMORIA\ENTREGABLE\TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712.xlsx",
    [string]$ExpectedPreEditSha256 = "31A4EFC9DC0A412B97C7ACDF95F8ADDA55370C134247BEB41C9F7A836508F7FE",
    [string]$LogPath = "C:\Users\julio\Desktop\TFM\08_MEMORIA\AUDITORIA_EXCELES_FINAL_20260712\APLICACION_WAZUH_BASE_CUSTOM_20260713.json"
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$xlCalculationAutomatic = -4105
$xlCategory = 1
$xlValue = 2
$xlPrimary = 1
$xlLegendPositionBottom = -4107
$xlPasteFormats = -4122
$xlHAlignLeft = -4131
$xlVAlignTop = -4160
$xlLandscape = 2
$xlContinuous = 1
$xlThin = 2
$xlValues = -4163
$xlWhole = 1
$xlByRows = 1
$xlNext = 1
$msoTextOrientationHorizontal = 1

if (-not ("TFMWazuhEditNative" -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class TFMWazuhEditNative {
    [DllImport("user32.dll", SetLastError=true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
"@
}

function U {
    param([Parameter(Mandatory = $true)][string]$Text)
    return [Text.RegularExpressions.Regex]::Unescape($Text)
}

function Rgb {
    param([int]$R, [int]$G, [int]$B)
    return $R + (256 * $G) + (65536 * $B)
}

function Release-ComObject {
    param($Object)
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object) } catch {}
    }
}

function Set-CellValue {
    param($Worksheet, [string]$Address, $Value)
    $range = $null
    try {
        $range = $Worksheet.Range($Address)
        if ($null -eq $Value) { $range.ClearContents() | Out-Null }
        elseif ($Value -is [byte] -or $Value -is [int16] -or $Value -is [int32] -or $Value -is [int64] -or $Value -is [single] -or $Value -is [double] -or $Value -is [decimal]) { $range.Value2 = [double]$Value }
        else { $range.Value2 = [string]$Value }
    } finally { Release-ComObject $range }
}

function Set-CellFormula {
    param($Worksheet, [string]$Address, [string]$Formula)
    $range = $null
    try { $range = $Worksheet.Range($Address); $range.Formula = $Formula }
    finally { Release-ComObject $range }
}

function Set-RowValues {
    param($Worksheet, [int]$Row, [object[]]$Values, [int]$StartColumn = 1)
    for ($i = 0; $i -lt $Values.Count; $i++) {
        $cell = $null
        try {
            $cell = $Worksheet.Cells.Item($Row, $StartColumn + $i)
            $value = $Values[$i]
            if ($null -eq $value) { $cell.ClearContents() | Out-Null }
            elseif ($value -is [byte] -or $value -is [int16] -or $value -is [int32] -or $value -is [int64] -or $value -is [single] -or $value -is [double] -or $value -is [decimal]) { $cell.Value2 = [double]$value }
            else { $cell.Value2 = [string]$value }
        } finally { Release-ComObject $cell }
    }
}

function Copy-Formats {
    param($Worksheet, [string]$Source, [string]$Destination)
    $sourceRange = $null
    $destinationRange = $null
    try {
        $sourceRange = $Worksheet.Range($Source)
        $destinationRange = $Worksheet.Range($Destination)
        $sourceRange.Copy() | Out-Null
        $destinationRange.PasteSpecial($xlPasteFormats) | Out-Null
    } finally {
        Release-ComObject $destinationRange
        Release-ComObject $sourceRange
    }
}

function Style-Header {
    param($Worksheet, [string]$Address)
    $range = $null
    try {
        $range = $Worksheet.Range($Address)
        $range.Font.Bold = $true
        $range.Font.Color = Rgb 255 255 255
        $range.Interior.Color = Rgb 31 78 121
        $range.WrapText = $true
        $range.HorizontalAlignment = $xlHAlignLeft
        $range.VerticalAlignment = $xlVAlignTop
        $range.Borders.LineStyle = $xlContinuous
        $range.Borders.Weight = $xlThin
        $range.Borders.Color = Rgb 155 194 230
    } finally { Release-ComObject $range }
}

function Style-Body {
    param($Worksheet, [string]$Address)
    $range = $null
    try {
        $range = $Worksheet.Range($Address)
        $range.WrapText = $true
        $range.HorizontalAlignment = $xlHAlignLeft
        $range.VerticalAlignment = $xlVAlignTop
        $range.Borders.LineStyle = $xlContinuous
        $range.Borders.Weight = $xlThin
        $range.Borders.Color = Rgb 217 225 242
    } finally { Release-ComObject $range }
}

function Style-NoteRange {
    param($Worksheet, [string]$Address, [string]$Text, [int]$FontSize = 9)
    $range = $null
    try {
        $range = $Worksheet.Range($Address)
        try { $range.UnMerge() } catch {}
        $range.Merge() | Out-Null
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
    } finally { Release-ComObject $range }
}

function Configure-Chart {
    param($Chart, [string]$Title, [string]$CategoryTitle, [string]$ValueTitle, [double]$Maximum = 0, [switch]$DataLabels)
    $Chart.HasTitle = $true
    $Chart.ChartTitle.Text = $Title
    try { $Chart.HasLegend = $true; $Chart.Legend.Position = $xlLegendPositionBottom } catch {}
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
        if ($Maximum -gt 0) { $y.MaximumScale = $Maximum }
        $y.TickLabels.NumberFormat = '0'
    } catch {}
    if ($DataLabels) {
        try {
            $Chart.ApplyDataLabels() | Out-Null
            $seriesCollection = $Chart.SeriesCollection()
            try {
                for ($i = 1; $i -le [int]$seriesCollection.Count; $i++) {
                    $series = $null
                    try { $series = $seriesCollection.Item($i); if ($series.HasDataLabels) { $series.DataLabels().NumberFormat = '0' } }
                    finally { Release-ComObject $series }
                }
            } finally { Release-ComObject $seriesCollection }
        } catch {}
    }
    Release-ComObject $y
    Release-ComObject $x
}

function Set-ChartSeries {
    param($Chart, [object[]]$Definitions)
    $seriesCollection = $null
    try {
        $seriesCollection = $Chart.SeriesCollection()
        while ([int]$seriesCollection.Count -gt 0) {
            $series = $null
            try { $series = $seriesCollection.Item(1); $series.Delete() }
            finally { Release-ComObject $series }
        }
        foreach ($definition in $Definitions) {
            $series = $null
            try {
                $series = $seriesCollection.NewSeries()
                $series.Name = $definition.Name
                $series.XValues = $definition.Categories
                $series.Values = $definition.Values
            } finally { Release-ComObject $series }
        }
    } finally { Release-ComObject $seriesCollection }
}

function Count-ExactMatches {
    param($Workbook, [string]$Needle)
    $count = 0
    $worksheets = $null
    try {
        $worksheets = $Workbook.Worksheets
        for ($i = 1; $i -le [int]$worksheets.Count; $i++) {
            $sheet = $null
            $used = $null
            try {
                $sheet = $worksheets.Item($i)
                $used = $sheet.UsedRange
                $values = $used.Value2
                if ($values -is [Array]) { foreach ($value in $values) { if ($null -ne $value -and ([string]$value) -ceq $Needle) { $count++ } } }
                elseif ($null -ne $values -and ([string]$values) -ceq $Needle) { $count++ }
            } finally { Release-ComObject $used; Release-ComObject $sheet }
        }
    } finally { Release-ComObject $worksheets }
    return $count
}

function Count-ContainingMatches {
    param($Workbook, [string]$Needle)
    $count = 0
    $worksheets = $null
    try {
        $worksheets = $Workbook.Worksheets
        for ($i = 1; $i -le [int]$worksheets.Count; $i++) {
            $sheet = $null
            $used = $null
            try {
                $sheet = $worksheets.Item($i)
                $used = $sheet.UsedRange
                $values = $used.Value2
                if ($values -is [Array]) { foreach ($value in $values) { if ($null -ne $value -and ([string]$value).IndexOf($Needle, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $count++ } } }
                elseif ($null -ne $values -and ([string]$values).IndexOf($Needle, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $count++ }
            } finally { Release-ComObject $used; Release-ComObject $sheet }
        }
    } finally { Release-ComObject $worksheets }
    return $count
}

function Get-ChartCount {
    param($Workbook)
    $count = 0
    $worksheets = $null
    try {
        $worksheets = $Workbook.Worksheets
        for ($i = 1; $i -le [int]$worksheets.Count; $i++) {
            $sheet = $null
            $charts = $null
            try { $sheet = $worksheets.Item($i); $charts = $sheet.ChartObjects(); $count += [int]$charts.Count }
            finally { Release-ComObject $charts; Release-ComObject $sheet }
        }
    } finally { Release-ComObject $worksheets }
    return $count
}

function Set-HeatCell {
    param($Worksheet, [string]$Address, [int]$Score)
    $cell = $null
    try {
        $cell = $Worksheet.Range($Address)
        $cell.Value2 = $Score
        $cell.HorizontalAlignment = -4108
        $cell.VerticalAlignment = -4108
        $cell.Font.Bold = $true
        if ($Score -eq 2) { $cell.Interior.Color = Rgb 198 239 206; $cell.Font.Color = Rgb 0 97 0 }
        elseif ($Score -eq 1) { $cell.Interior.Color = Rgb 255 235 156; $cell.Font.Color = Rgb 156 101 0 }
        else { $cell.Interior.Color = Rgb 255 199 206; $cell.Font.Color = Rgb 156 0 6 }
        $cell.Borders.LineStyle = $xlContinuous
        $cell.Borders.Weight = $xlThin
    } finally { Release-ComObject $cell }
}

$resolved = (Resolve-Path -LiteralPath $Path).Path
$actualPreHash = (Get-FileHash -LiteralPath $resolved -Algorithm SHA256).Hash
if ($actualPreHash -ne $ExpectedPreEditSha256) {
    throw "El maestro actual ha cambiado desde el backup. Esperado=$ExpectedPreEditSha256 Actual=$actualPreHash. No se edita."
}

$excel = $null
$workbook = $null
$dashboard = $null
$graphs = $null
$comparison = $null
$matrix = $null
$wazuhDetail = $null
$executive = $null
$sources = $null
$discrepancies = $null
$closed = $null
$readme = $null
$excelPid = 0
$forcedProcessClose = $false
$before = $null
$after = $null
$matrixRangeAddress = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    try { $excel.AutomationSecurity = 3 } catch {}
    [uint32]$pidValue = 0
    [void][TFMWazuhEditNative]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$pidValue)
    $excelPid = [int]$pidValue

    $workbook = $excel.Workbooks.Open($resolved, 0, $false)
    $before = [pscustomobject]@{
        Sheets = [int]$workbook.Worksheets.Count
        Charts = Get-ChartCount $workbook
        Pending = Count-ExactMatches $workbook 'Pendiente'
        AmbiguousWazuhLabel = Count-ContainingMatches $workbook 'Wazuh espec'
    }
    if ($before.Sheets -ne 31 -or $before.Charts -ne 11 -or $before.Pending -ne 0) { throw "Precondiciones del maestro no cumplidas." }

    $dashboard = $workbook.Worksheets.Item('01_Dashboard')
    $graphs = $workbook.Worksheets.Item('GRAFICAS')
    $comparison = $workbook.Worksheets.Item('COMPARACION_VR_WAZUH')
    $matrix = $workbook.Worksheets.Item('05_Matriz_Resultados')
    $wazuhDetail = $workbook.Worksheets.Item('WAZUH_DETALLE')
    $executive = $workbook.Worksheets.Item('RESUMEN_EJECUTIVO')
    $sources = $workbook.Worksheets.Item('FUENTES')
    $discrepancies = $workbook.Worksheets.Item('DISCREPANCIAS')
    $closed = $workbook.Worksheets.Item('17_Incoherencias_Cerradas')
    $readme = $workbook.Worksheets.Item('README')

    # GRAFICAS - cobertura por tecnica y resumen de deteccion especifica.
    Set-RowValues $graphs 60 @('TEC', 'VR custom', (U 'P\u00FAblicos CH'), 'Wazuh base', 'Wazuh custom') 1
    $publicVector = @(1,0,1,0,1,0,0,0,0)
    $baseVector = @(0,1,0,1,0,1,0,0,0)
    $customVector = @(1,0,0,0,1,0,1,1,0)
    for ($i = 0; $i -lt 9; $i++) {
        Set-RowValues $graphs (61 + $i) @(("TEC-{0:D3}" -f ($i + 1)), 1, $publicVector[$i], $baseVector[$i], $customVector[$i]) 1
    }
    Style-Header $graphs 'A60:E60'
    Style-Body $graphs 'A61:E69'

    Set-RowValues $graphs 60 @('Sistema', (U 'TEC con detecci\u00F3n espec\u00EDfica')) 6
    Set-CellValue $graphs 'F61' 'VR custom'
    Set-CellValue $graphs 'F62' (U 'P\u00FAblicos CH')
    Set-CellValue $graphs 'F63' 'Wazuh base'
    Set-CellValue $graphs 'F64' 'Wazuh custom'
    Set-CellFormula $graphs 'G61' '=SUM(B61:B69)'
    Set-CellFormula $graphs 'G62' '=SUM(C61:C69)'
    Set-CellFormula $graphs 'G63' '=SUM(D61:D69)'
    Set-CellFormula $graphs 'G64' '=SUM(E61:E69)'
    $graphs.Range('F65:G69').ClearContents() | Out-Null
    Style-Header $graphs 'F60:G60'
    Style-Body $graphs 'F61:G64'

    # GRAFICAS - replace raw-volume base/custom chart with technique coverage.
    Set-RowValues $graphs 18 @('TEC', 'Wazuh base', 'Wazuh custom') 7
    for ($i = 0; $i -lt 9; $i++) { Set-RowValues $graphs (19 + $i) @(("TEC-{0:D3}" -f ($i + 1)), $baseVector[$i], $customVector[$i]) 7 }
    Style-Header $graphs 'G18:I18'
    Style-Body $graphs 'G19:I27'

    # GRAFICAS - HIDS capabilities, four separated variants and no external output.
    Set-CellValue $graphs 'M49' (U 'Capacidades HIDS acreditadas por sistema')
    Set-RowValues $graphs 50 @('Capacidad', 'VR custom', (U 'P\u00FAblicos CH'), 'Wazuh base', 'Wazuh custom') 13
    $heatRows = @(
        @((U 'Visibilidad'), 2, 2, 1, 2),
        @((U 'Detecci\u00F3n'), 2, 1, 1, 2),
        @('Alerta RT', 2, 1, 1, 2),
        @((U 'Evidencia forense'), 0, 2, 1, 1),
        @('Ruido/FP', 2, 1, 1, 1)
    )
    for ($r = 0; $r -lt $heatRows.Count; $r++) {
        Set-CellValue $graphs ("M{0}" -f (51 + $r)) $heatRows[$r][0]
        for ($c = 0; $c -lt 4; $c++) { Set-HeatCell $graphs (([char](78 + $c)).ToString() + (51 + $r)) ([int]$heatRows[$r][$c + 1]) }
    }
    Style-Header $graphs 'M50:Q50'
    Style-Body $graphs 'M51:M55'
    Style-NoteRange $graphs 'M58:Q58' (U 'La salida externa se eval\u00FAa separadamente porque depende de componentes server-side, Router, webhook o runners auxiliares.') 9
    Style-NoteRange $graphs 'M59:Q59' (U 'Escala: 0=no acreditado; 1=parcial o contextual; 2=capacidad acreditada. Las cifras de detecci\u00F3n se comparan aparte como t\u00E9cnicas sobre 9.') 8

    # Dashboard helper charts: four separated bars and four detection series by TEC.
    $coverageObject = $null
    $coverageChart = $null
    try {
        $coverageObject = $dashboard.ChartObjects('TFM_Final_Public_TEC')
        $coverageChart = $coverageObject.Chart
        Set-ChartSeries $coverageChart @(
            [pscustomobject]@{Name='=GRAFICAS!$B$60';Categories='=GRAFICAS!$A$61:$A$69';Values='=GRAFICAS!$B$61:$B$69'},
            [pscustomobject]@{Name='=GRAFICAS!$C$60';Categories='=GRAFICAS!$A$61:$A$69';Values='=GRAFICAS!$C$61:$C$69'},
            [pscustomobject]@{Name='=GRAFICAS!$D$60';Categories='=GRAFICAS!$A$61:$A$69';Values='=GRAFICAS!$D$61:$D$69'},
            [pscustomobject]@{Name='=GRAFICAS!$E$60';Categories='=GRAFICAS!$A$61:$A$69';Values='=GRAFICAS!$E$61:$E$69'}
        )
        Configure-Chart $coverageChart (U 'Cobertura de detecci\u00F3n espec\u00EDfica por t\u00E9cnica') (U 'T\u00E9cnica TEC') (U 'Detecci\u00F3n acreditada (0/1)') 1
    } finally { Release-ComObject $coverageChart; Release-ComObject $coverageObject }

    $detectionObject = $null
    $detectionChart = $null
    try {
        $detectionObject = $dashboard.ChartObjects('TFM_Final_Compare_Detection')
        $detectionChart = $detectionObject.Chart
        Set-ChartSeries $detectionChart @([pscustomobject]@{Name='=GRAFICAS!$G$60';Categories='=GRAFICAS!$F$61:$F$64';Values='=GRAFICAS!$G$61:$G$64'})
        Configure-Chart $detectionChart (U 'Detecci\u00F3n espec\u00EDfica por sistema y tipo de reglas') 'Sistema' (U 'N.\u00BA de t\u00E9cnicas TEC detectadas') 9 -DataLabels
        $detectionObject.Height = 195
    } finally { Release-ComObject $detectionChart; Release-ComObject $detectionObject }

    $graph4Object = $null
    $graph4 = $null
    try {
        $graph4Object = $graphs.ChartObjects('Chart 4')
        $graph4 = $graph4Object.Chart
        Set-ChartSeries $graph4 @(
            [pscustomobject]@{Name='=GRAFICAS!$H$18';Categories='=GRAFICAS!$G$19:$G$27';Values='=GRAFICAS!$H$19:$H$27'},
            [pscustomobject]@{Name='=GRAFICAS!$I$18';Categories='=GRAFICAS!$G$19:$G$27';Values='=GRAFICAS!$I$19:$I$27'}
        )
        Configure-Chart $graph4 (U 'Wazuh base y custom: detecci\u00F3n por t\u00E9cnica') (U 'T\u00E9cnica TEC') (U 'Detecci\u00F3n acreditada (0/1)') 1 -DataLabels
    } finally { Release-ComObject $graph4; Release-ComObject $graph4Object }

    # Dashboard KPIs and independent explanatory note.
    Set-CellValue $dashboard 'A14' (U 'Detecci\u00F3n Wazuh custom')
    Set-CellFormula $dashboard 'B14' '=GRAFICAS!$G$64&"/9"'
    Set-CellValue $dashboard 'C14' (U 'TEC-001, TEC-005, TEC-007 y TEC-008 con reglas 110xxx; 110201 mantiene gap TEC-009.')
    Copy-Formats $dashboard 'A14:C14' 'A16:C16'
    Set-CellValue $dashboard 'A16' (U 'Detecci\u00F3n Wazuh base')
    Set-CellFormula $dashboard 'B16' '=GRAFICAS!$G$63&"/9"'
    Set-CellValue $dashboard 'C16' (U 'TEC-002, TEC-004 y TEC-006 con reglas nativas; alertas gen\u00E9ricas y simple visibilidad no se contabilizan.')

    $shapes = $null
    $noteShape = $null
    $noteChartObject = $null
    try {
        $shapes = $dashboard.Shapes
        try { $noteShape = $shapes.Item('TFM_Wazuh_Base_Custom_Note') } catch {}
        $noteChartObject = $dashboard.ChartObjects('TFM_Final_Compare_Detection')
        if ($null -eq $noteShape) {
            $noteShape = $shapes.AddTextbox($msoTextOrientationHorizontal, $noteChartObject.Left, $noteChartObject.Top + $noteChartObject.Height + 4, $noteChartObject.Width, 48)
            $noteShape.Name = 'TFM_Wazuh_Base_Custom_Note'
        } else {
            $noteShape.Left = $noteChartObject.Left
            $noteShape.Top = $noteChartObject.Top + $noteChartObject.Height + 4
            $noteShape.Width = $noteChartObject.Width
            $noteShape.Height = 48
        }
        $noteShape.TextFrame2.TextRange.Text = U 'Wazuh base representa el ruleset nativo; Wazuh custom representa las reglas espec\u00EDficas desarrolladas para el TFM. La m\u00E9trica cuenta t\u00E9cnicas detectadas, no n\u00FAmero de alertas.'
        $noteShape.TextFrame2.TextRange.Font.Size = 8.5
        $noteShape.TextFrame2.TextRange.Font.Italic = $true
        $noteShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Rgb 31 78 121
        $noteShape.Fill.ForeColor.RGB = Rgb 221 235 247
        $noteShape.Line.ForeColor.RGB = Rgb 155 194 230
        $noteShape.TextFrame2.MarginLeft = 6
        $noteShape.TextFrame2.MarginRight = 6
        $noteShape.TextFrame2.MarginTop = 3
        $noteShape.TextFrame2.MarginBottom = 3
    } finally { Release-ComObject $noteChartObject; Release-ComObject $noteShape; Release-ComObject $shapes }

    # COMPARACION_VR_WAZUH - preserve raw counters and add fair technique comparison.
    Set-CellValue $comparison 'A2' (U 'Comparaci\u00F3n metodol\u00F3gica: unidad principal = t\u00E9cnicas TEC detectadas sobre 9; alertas y eventos se mantienen como contexto separado.')
    Set-CellValue $comparison 'A6' (U 'Wazuh base (ventana base sin 110xxx)')
    Set-CellValue $comparison 'M6' (U 'Baseline de ruido/visibilidad: 58 alertas y ning\u00FAn marcador TEC. No se usa para atribuir cobertura t\u00E9cnica.')
    Set-CellValue $comparison 'A7' (U 'Wazuh campa\u00F1a TEC (ruleset base + custom)')
    Set-CellValue $comparison 'I7' (U '56 alertas 110xxx')
    Set-CellValue $comparison 'M7' (U '110 alertas totales = 54 nativas + 56 custom 110xxx; el total combinado no es una quinta soluci\u00F3n.')

    Set-RowValues $comparison 10 @('Sistema', (U 'T\u00E9cnicas visibles'), (U 'T\u00E9cnicas detectadas espec\u00EDficamente'), 'Alerta RT', 'Tipo de reglas', 'Observaciones') 1
    Set-RowValues $comparison 11 @('Velociraptor custom', '9/9', $null, '9/9', 'Artifacts custom P1-P4 CLIENT_EVENT', (U 'Detecci\u00F3n y alerta espec\u00EDfica 9/9; transporte y salida externa se eval\u00FAan aparte.')) 1
    Set-CellFormula $comparison 'C11' '=GRAFICAS!$G$61&"/9"'
    Set-RowValues $comparison 12 @((U 'Velociraptor p\u00FAblicos CH'), '9/9', $null, '3/9', (U 'Artifacts p\u00FAblicos + Hayabusa Sigma CH'), (U 'Detecci\u00F3n 3/9: TEC-001, TEC-003 y TEC-005; CHM no ejecutada.')) 1
    Set-CellFormula $comparison 'C12' '=GRAFICAS!$G$62&"/9"'
    Set-RowValues $comparison 13 @('Wazuh base', '9/9', $null, '3/9', 'Ruleset nativo Wazuh', (U 'TEC-002, TEC-004 y TEC-006; rules 92032, 92302, 91835 y 92077. No se cuentan alertas gen\u00E9ricas.')) 1
    Set-CellFormula $comparison 'C13' '=GRAFICAS!$G$63&"/9"'
    Set-RowValues $comparison 14 @('Wazuh custom', '9/9', $null, '4/9', (U 'Reglas TFM 110xxx'), (U 'TEC-001, TEC-005, TEC-007 y TEC-008; rules 110301, 110402, 110203 y 110202. 110201=0.')) 1
    Set-CellFormula $comparison 'C14' '=GRAFICAS!$G$64&"/9"'
    Style-Header $comparison 'A10:F10'
    Style-Body $comparison 'A11:F14'

    Set-CellValue $comparison 'A16' (U 'Desglose Wazuh base/custom - unidad: t\u00E9cnicas TEC sobre 9, no n\u00FAmero de alertas')
    $comparison.Range('A16:F16').Merge() | Out-Null
    Style-Header $comparison 'A16:F16'
    Set-RowValues $comparison 17 @('Grupo', (U 'TEC detectadas'), 'Rule IDs principales', 'Alertas fuente', 'Ventana UTC', 'Limitaciones') 1
    Set-RowValues $comparison 18 @('Wazuh base', 'TEC-002, TEC-004, TEC-006', '92032; 92302; 91835; 92077', (U '54 alertas nativas dentro de la campa\u00F1a TEC'), '2026-06-22 11:26:39..11:37:39Z', (U 'La ventana base separada no contiene marcadores TEC. 92307/T1543.003 y schtasks gen\u00E9rico se tratan como visibilidad/alerta relacionada.')) 1
    Set-RowValues $comparison 19 @('Wazuh custom', 'TEC-001, TEC-005, TEC-007, TEC-008', '110301; 110402; 110203; 110202', '56 alertas 110xxx', '2026-06-22 11:26:39..11:37:39Z', (U '110201=0; 110501 es gen\u00E9rica; 110202 incluye coincidencias fuera de TEC-008 y no suma t\u00E9cnicas adicionales.')) 1
    Set-RowValues $comparison 20 @((U 'Uni\u00F3n base + custom (complementaria)'), '7/9', (U 'Uni\u00F3n de los dos grupos; no quinta soluci\u00F3n'), 'No sumar alertas como TEC', 'Misma campa\u00F1a TEC', (U 'Persisten gaps espec\u00EDficos TEC-003 y TEC-009 en Wazuh; la personalizaci\u00F3n mejora cobertura con coste de ingenier\u00EDa y ajuste.')) 1
    Style-Header $comparison 'A17:F17'
    Style-Body $comparison 'A18:F20'
    Style-NoteRange $comparison 'A22:F23' (U 'Lectura justa: Velociraptor custom acredita 9/9; Hayabusa CH 3/9; Wazuh base 3/9 y Wazuh custom 4/9. Las reglas custom a\u00F1aden cobertura complementaria, pero requieren dise\u00F1o, validaci\u00F3n y afinado para controlar el ruido.') 9
    $comparison.Range('A1:M23').WrapText = $true
    $comparison.Columns('A').ColumnWidth = 31
    $comparison.Columns('B').ColumnWidth = 19
    $comparison.Columns('C').ColumnWidth = 27
    $comparison.Columns('D').ColumnWidth = 20
    $comparison.Columns('E').ColumnWidth = 30
    $comparison.Columns('F').ColumnWidth = 62
    $comparison.Rows('10:23').RowHeight = 42
    $comparison.Rows('16:17').RowHeight = 28
    $comparison.PageSetup.Orientation = $xlLandscape
    $comparison.PageSetup.PrintArea = '$A$1:$M$23'
    $comparison.PageSetup.Zoom = $false
    $comparison.PageSetup.FitToPagesWide = 1
    $comparison.PageSetup.FitToPagesTall = $false

    # 05_Matriz_Resultados - four separate detection variants and rule evidence.
    Set-RowValues $matrix 4 @('VR custom', (U 'P\u00FAblicos CH'), 'Wazuh base', (U 'Wazuh base - regla/evidencia'), 'Wazuh custom', (U 'Wazuh custom - regla/evidencia')) 21
    $detected = U 'DETECTADA'
    $visibleNoDetection = U 'VISIBLE SIN DETECCI\u00D3N'
    $publicDetected = @(1,0,1,0,1,0,0,0,0)
    $baseStates = @($visibleNoDetection,$detected,$visibleNoDetection,$detected,$visibleNoDetection,$detected,$visibleNoDetection,$visibleNoDetection,$visibleNoDetection)
    $baseEvidence = @(
        (U '4104/Sysmon visibles; sin regla nativa aplicable a TEC-001.'),
        '92032 - T1059.003 - evento 2026-06-22T11:28:46.180Z.',
        (U '92032/92052 alertan shell; sin mapeo T1053.005 espec\u00EDfico.'),
        '92302 - T1547.001 - evento 2026-06-22T11:29:05.491Z.',
        (U '92307 mapea T1543.003/creaci\u00F3n; no se cuenta como T1569.002 espec\u00EDfico.'),
        '91835/92077 - T1518.001 - eventos 2026-06-22T11:29:30..11:29:32Z.',
        (U '4104/ID11 visibles; sin alerta nativa espec\u00EDfica T1486.'),
        (U '4104 visible; sin alerta nativa espec\u00EDfica T1485.'),
        (U 'Staging/ZIP visibles; sin alerta nativa espec\u00EDfica de TEC-009.')
    )
    $customStates = @($detected,$visibleNoDetection,$visibleNoDetection,$visibleNoDetection,$detected,$visibleNoDetection,$detected,$detected,$visibleNoDetection)
    $customEvidence = @(
        (U '110301 (2 alertas); 110501 es gen\u00E9rica y no a\u00F1ade otra TEC.'),
        (U 'Telemetr\u00EDa visible; sin regla custom espec\u00EDfica T1059.003.'),
        (U 'Telemetr\u00EDa schtasks visible; sin regla custom T1053.005 activada.'),
        (U '110401=0; Run Key visible, sin detecci\u00F3n custom.'),
        '110402 (1 alerta) - T1569.002.',
        (U 'Discovery visible; sin regla custom espec\u00EDfica activa en v2.'),
        '110203 (4 alertas) - T1486.',
        (U '110202 (5 alertas totales); TEC detectada, con ruido documentado.'),
        (U '110201=0; visibilidad asociada sin detecci\u00F3n custom.')
    )
    for ($i = 0; $i -lt 9; $i++) {
        $publicState = if ($publicDetected[$i] -eq 1) { $detected } else { $visibleNoDetection }
        Set-RowValues $matrix (5 + $i) @($detected, $publicState, $baseStates[$i], $baseEvidence[$i], $customStates[$i], $customEvidence[$i]) 21
    }
    $matrixTable = $null
    $matrixResizeRange = $null
    $matrixTableRange = $null
    try {
        $matrixTable = $matrix.ListObjects.Item('MatrizResultadosTable')
        $matrixResizeRange = $matrix.Range('A4:Z13')
        $matrixTable.Resize($matrixResizeRange)
        $matrixTableRange = $matrixTable.Range
        $matrixRangeAddress = [string]$matrixTableRange.Address()
    }
    finally {
        Release-ComObject $matrixTableRange
        Release-ComObject $matrixResizeRange
        Release-ComObject $matrixTable
    }
    $matrix.Range('U4:Z13').WrapText = $true
    $matrix.Columns('U').ColumnWidth = 22
    $matrix.Columns('V').ColumnWidth = 23
    $matrix.Columns('W').ColumnWidth = 24
    $matrix.Columns('X').ColumnWidth = 52
    $matrix.Columns('Y').ColumnWidth = 24
    $matrix.Columns('Z').ColumnWidth = 52
    $matrix.Rows('5:13').RowHeight = 62
    $matrix.PageSetup.Orientation = $xlLandscape
    $matrix.PageSetup.PrintArea = '$A$1:$Z$13'
    $matrix.PageSetup.PrintTitleRows = '$4:$4'
    $matrix.PageSetup.Zoom = $false
    $matrix.PageSetup.FitToPagesWide = 1
    $matrix.PageSetup.FitToPagesTall = $false

    # WAZUH_DETALLE - unified rule traceability.
    try { $wazuhDetail.Range('A1:H2').UnMerge() } catch {}
    $wazuhDetail.Range('A1:H1').Merge() | Out-Null
    $wazuhDetail.Range('A2:H2').Merge() | Out-Null
    Set-CellValue $wazuhDetail 'A1' (U 'Wazuh detalle: ruleset base y reglas custom 110xxx')
    Set-CellValue $wazuhDetail 'A2' (U 'Clasificaci\u00F3n por regla y alerta real en la ventana TEC 2026-06-22 11:26:39..11:37:39Z. La unidad final es la t\u00E9cnica TEC \u00FAnica, no el n\u00FAmero de alertas.')
    Set-RowValues $wazuhDetail 4 @('rule_id','grupo',(U 'descripci\u00F3n'),'TEC','MITRE','alertas',(U 'veredicto para m\u00E9trica'),'fuente / ventana') 1
    $ruleRows = @(
        @('92032','BASE','Suspicious Windows cmd shell execution','TEC-002','T1059.003',14,(U 'DETECTADA TEC-002; otras coincidencias no suman TEC.'),'custom/wazuh_custom_alerts_delta.jsonl'),
        @('92302','BASE','Run key modified using reg.exe','TEC-004','T1547.001',1,$detected,'custom/wazuh_custom_alerts_delta.jsonl'),
        @('91835','BASE','PowerShell Antivirus Software discovery','TEC-006','T1518.001',1,$detected,'custom/wazuh_custom_alerts_delta.jsonl'),
        @('92077','BASE','WMI AV product discovery','TEC-006','T1518.001',3,$detected,'custom/wazuh_custom_alerts_delta.jsonl'),
        @('92307','BASE','New service creation in registry','TEC-005 relacionado','T1543.003',1,(U 'VISIBLE/ALERTA RELACIONADA; no T1569.002 espec\u00EDfico.'),'custom/wazuh_custom_alerts_delta.jsonl'),
        @('110201','CUSTOM','TFM TEC-009 posible exfiltracion HTTP ZIP','TEC-009','T1048.003 / T1560.001',0,$visibleNoDetection,'custom/tfm_wazuh_custom_detections_110xxx.jsonl'),
        @('110202','CUSTOM','TFM TEC-008 borrado destructivo PowerShell','TEC-008','T1485',5,(U 'DETECTADA; incluye ruido fuera de TEC-008.'),'custom/tfm_wazuh_custom_detections_110xxx.jsonl'),
        @('110203','CUSTOM','TFM TEC-007 cifrado PowerShell','TEC-007','T1486',4,$detected,'custom/tfm_wazuh_custom_detections_110xxx.jsonl'),
        @('110301','CUSTOM','TFM PowerShell sospechoso 4104','TEC-001','T1059.001',2,$detected,'custom/tfm_wazuh_custom_detections_110xxx.jsonl'),
        @('110401','CUSTOM','TFM Run Key Sysmon','TEC-004','T1547.001',0,$visibleNoDetection,'custom/tfm_wazuh_custom_detections_110xxx.jsonl'),
        @('110402','CUSTOM','TFM instalacion de servicio Windows','TEC-005','T1569.002',1,$detected,'custom/tfm_wazuh_custom_detections_110xxx.jsonl'),
        @('110501','CUSTOM','TFM PowerShell observada por Sysmon','TEC-001 contexto','T1059.001',44,(U 'REGLA GEN\u00C9RICA; no suma t\u00E9cnica adicional.'),'custom/tfm_wazuh_custom_detections_110xxx.jsonl')
    )
    for ($i = 0; $i -lt $ruleRows.Count; $i++) {
        $row = @($ruleRows[$i])
        $row[7] = $row[7] + ' - 2026-06-22 11:26:39..11:37:39Z'
        Set-RowValues $wazuhDetail (5 + $i) $row 1
    }
    Set-CellValue $wazuhDetail 'A18' (U 'Resultado: Wazuh base 3/9 (TEC-002/004/006); Wazuh custom 4/9 (TEC-001/005/007/008). Uni\u00F3n complementaria 7/9; no tratarla como quinta soluci\u00F3n.')
    $wazuhDetail.Range('A18:H19').Merge() | Out-Null
    Style-Header $wazuhDetail 'A4:H4'
    Style-Body $wazuhDetail 'A5:H16'
    Style-NoteRange $wazuhDetail 'A18:H19' (U 'Resultado: Wazuh base 3/9 (TEC-002/004/006); Wazuh custom 4/9 (TEC-001/005/007/008). Uni\u00F3n complementaria 7/9; no tratarla como quinta soluci\u00F3n.') 9
    $wazuhDetail.Columns('A').ColumnWidth = 11
    $wazuhDetail.Columns('B').ColumnWidth = 12
    $wazuhDetail.Columns('C').ColumnWidth = 42
    $wazuhDetail.Columns('D').ColumnWidth = 20
    $wazuhDetail.Columns('E').ColumnWidth = 22
    $wazuhDetail.Columns('F').ColumnWidth = 11
    $wazuhDetail.Columns('G').ColumnWidth = 42
    $wazuhDetail.Columns('H').ColumnWidth = 54
    $wazuhDetail.Rows('5:16').RowHeight = 46
    $wazuhDetail.PageSetup.Orientation = $xlLandscape
    $wazuhDetail.PageSetup.PrintArea = '$A$1:$H$19'
    $wazuhDetail.PageSetup.PrintTitleRows = '$4:$4'
    $wazuhDetail.PageSetup.Zoom = $false
    $wazuhDetail.PageSetup.FitToPagesWide = 1
    $wazuhDetail.PageSetup.FitToPagesTall = $false

    # Executive summary - separate campaign totals from rule groups and technique coverage.
    Set-CellValue $executive 'A21' (U 'Wazuh base: alertas en ventana baseline')
    Set-CellValue $executive 'B21' 58
    Set-CellValue $executive 'C21' (U 'Sin reglas 110xxx y sin marcadores TEC; contexto de ruido/visibilidad, no cobertura t\u00E9cnica.')
    Set-CellValue $executive 'A22' (U 'Wazuh campa\u00F1a TEC: alertas totales')
    Set-CellValue $executive 'B22' 110
    Set-CellValue $executive 'C22' (U '54 alertas nativas + 56 alertas 110xxx; no interpretar 110 como t\u00E9cnicas detectadas.')
    Set-CellValue $executive 'A23' (U 'Wazuh custom: alertas 110xxx')
    Set-CellValue $executive 'B23' 56
    Set-CellValue $executive 'C23' (U 'Cinco rule_id activadas; cuatro t\u00E9cnicas TEC \u00FAnicas detectadas; 110201=0.')
    Copy-Formats $executive 'A30:C30' 'A34:C37'
    Set-CellValue $executive 'A34' (U 'TEC detectadas por Wazuh base')
    Set-CellFormula $executive 'B34' '=GRAFICAS!$G$63'
    Set-CellValue $executive 'C34' (U '3/9: TEC-002, TEC-004 y TEC-006; ruleset nativo 92032/92302/91835/92077.')
    Set-CellValue $executive 'A35' (U 'TEC detectadas por Wazuh custom')
    Set-CellFormula $executive 'B35' '=GRAFICAS!$G$64'
    Set-CellValue $executive 'C35' (U '4/9: TEC-001, TEC-005, TEC-007 y TEC-008; rules 110301/110402/110203/110202.')
    Set-CellValue $executive 'A36' (U 'Aportaci\u00F3n de reglas custom Wazuh')
    Set-CellValue $executive 'B36' (U '+4 complementarias')
    Set-CellValue $executive 'C36' (U 'La uni\u00F3n base+custom alcanza 7/9 como dato complementario; la mejora exige ingenier\u00EDa, validaci\u00F3n y ajuste de ruido.')
    Set-CellValue $executive 'A37' (U 'Unidad de comparaci\u00F3n')
    Set-CellValue $executive 'B37' (U 'T\u00E9cnicas /9')
    Set-CellValue $executive 'C37' (U 'Velociraptor custom 9/9; Hayabusa CH 3/9; Wazuh base 3/9; Wazuh custom 4/9. No se comparan vol\u00FAmenes de eventos como cobertura.')
    $executive.Range('A21:C37').WrapText = $true
    $executive.Rows('34:37').RowHeight = 42

    # README - methodology and exact evidence path.
    Set-CellValue $readme 'A2' (U 'Libro definitivo consolidado; preserva custom, artifacts p\u00FAblicos, FP y separaci\u00F3n Wazuh base/custom.')
    Set-CellValue $readme 'B6' '2026-07-13 (Europe/Madrid)'
    Set-CellValue $readme 'B9' '10_WAZUH\TFM_WAZUH_VISIBILIDAD_EVIDENCE.tar.gz'
    Set-CellValue $readme 'A15' (U 'Criterio de detecci\u00F3n comparativa')
    Set-CellValue $readme 'B15' (U 'Unidad=t\u00E9cnicas TEC/9: VR custom=9/9; Hayabusa CH=3/9; Wazuh base=3/9; Wazuh custom=4/9. Visibilidad, alerta RT, transporte y salida externa permanecen separadas.')
    Copy-Formats $readme 'A15:B15' 'A16:B16'
    Set-CellValue $readme 'A16' 'Criterio Wazuh base/custom'
    Set-CellValue $readme 'B16' (U 'Base=ruleset nativo; custom=reglas TFM 110xxx. 110 alertas de campa\u00F1a = 54 nativas + 56 custom; no equivalen a 110 detecciones ni a 110 t\u00E9cnicas.')
    $readme.Range('A15:B16').WrapText = $true
    $readme.Rows('15:16').RowHeight = 44

    # FUENTES - append only the Wazuh files actually used.
    Copy-Formats $sources 'A218:G218' 'A219:G228'
    $sourceRows = @(
        @('10_WAZUH\TFM_WAZUH_VISIBILIDAD_EVIDENCE.tar.gz','tar.gz',1990280,'2026-06-22T13:43:50+02:00','DB3B4B316B6952AEB4F3551C60DA30494B16CA5EABED95F68D7480EE62AD943F','WAZUH_EVIDENCE_CONTAINER','WAZUH_BASE_CUSTOM'),
        @('10_WAZUH\TFM_WAZUH_VISIBILIDAD_EVIDENCE.tar.gz::ENTREGA_WAZUH_VISIBILIDAD_20260622_114027/base/start_utc.txt','txt',21,'2026-06-22T13:40:31+02:00','B526A6D767A63C0A2F0BA277250567529B3CDFDEAD7B7863684E8DA7ECD7653D','WINDOW_BASE','WAZUH_BASE'),
        @('10_WAZUH\TFM_WAZUH_VISIBILIDAD_EVIDENCE.tar.gz::ENTREGA_WAZUH_VISIBILIDAD_20260622_114027/base/end_utc.txt','txt',21,'2026-06-22T13:40:31+02:00','CAD3C3AE35884CEBBA84BB9F421CE3C704BC872AE2CB81CDDE53B6E740DD9F42','WINDOW_BASE','WAZUH_BASE'),
        @('10_WAZUH\TFM_WAZUH_VISIBILIDAD_EVIDENCE.tar.gz::ENTREGA_WAZUH_VISIBILIDAD_20260622_114027/base/wazuh_base_alerts_delta.jsonl','jsonl',160796,'2026-06-22T13:40:31+02:00','FE8B4125B868BACB2B44490490DD2681FD7B60A9FA80A39A678BF18D4A461517','WAZUH_BASE_ALERTS_BASELINE','WAZUH_BASE'),
        @('10_WAZUH\TFM_WAZUH_VISIBILIDAD_EVIDENCE.tar.gz::ENTREGA_WAZUH_VISIBILIDAD_20260622_114027/custom/start_utc.txt','txt',21,'2026-06-22T13:40:38+02:00','C2203B5B49FC6ED3A601D9C49F1E8C640DDC194DFF1814864C8907E53D01747D','WINDOW_TEC','WAZUH_BASE_CUSTOM'),
        @('10_WAZUH\TFM_WAZUH_VISIBILIDAD_EVIDENCE.tar.gz::ENTREGA_WAZUH_VISIBILIDAD_20260622_114027/custom/end_utc.txt','txt',21,'2026-06-22T13:40:38+02:00','34D79555FF5BAA01BDBC8898B8736C72763C68437025B3DB79979F59EDB5C754','WINDOW_TEC','WAZUH_BASE_CUSTOM'),
        @('10_WAZUH\TFM_WAZUH_VISIBILIDAD_EVIDENCE.tar.gz::ENTREGA_WAZUH_VISIBILIDAD_20260622_114027/custom/wazuh_custom_alerts_delta.jsonl','jsonl',431533,'2026-06-22T13:40:38+02:00','4C1CA23E52964C7752CDB589E0C2CAA230D9442BA0B4ABD2251D090E6B38E6EB','WAZUH_NATIVE_AND_CUSTOM_ALERTS','WAZUH_BASE_CUSTOM'),
        @('10_WAZUH\TFM_WAZUH_VISIBILIDAD_EVIDENCE.tar.gz::ENTREGA_WAZUH_VISIBILIDAD_20260622_114027/custom/tfm_wazuh_custom_detections_110xxx.jsonl','jsonl',314214,'2026-06-22T13:40:42+02:00','B095F98CC61F79AC282353A4A1053407766ABA5D43776EFFD1F3C60EB2E82428','WAZUH_CUSTOM_110XXX_ALERTS','WAZUH_CUSTOM'),
        @('10_WAZUH\TFM_WAZUH_VISIBILIDAD_EVIDENCE.tar.gz::ENTREGA_WAZUH_VISIBILIDAD_20260622_114027/config/tfm_wazuh_custom_rules_v2.xml','xml',3204,'2026-06-22T13:41:05+02:00','735E58C362B6E29CEFA12D4ABBDA4BF9A4BDD671E5F5CD82E76C91796735896E','WAZUH_CUSTOM_RULESET_ACTIVE','WAZUH_CUSTOM'),
        @('10_WAZUH\WAZUH.docx','docx',1354384,'2026-06-22T14:03:12+02:00','690307EC1673BC3B78876CFCB51D7CD66421C60E7B7CBD0DCA815675B6E0A016','WAZUH_METHOD_NOTE','WAZUH_BASE_CUSTOM')
    )
    for ($i = 0; $i -lt $sourceRows.Count; $i++) { Set-RowValues $sources (219 + $i) $sourceRows[$i] 1 }
    $sources.Range('A219:G228').WrapText = $true
    $sources.Rows('219:228').RowHeight = 42

    # DISCREPANCIAS - explicit classification decisions and limitations.
    Copy-Formats $discrepancies 'A58:E58' 'A59:E62'
    Set-RowValues $discrepancies 59 @('WAZUH_BASELINE_SIN_MARCADORES_TEC','INFO','10_WAZUH\TFM_WAZUH_VISIBILIDAD_EVIDENCE.tar.gz','La ventana base contiene 902 archives/58 alertas pero ningun marcador TEC; se usa como baseline de ruido/visibilidad.','La cobertura base se clasifica con alertas nativas de la ventana TEC custom, donde el ruleset base seguia activo.') 1
    Set-RowValues $discrepancies 60 @('WAZUH_BASE_SERVICIO_MAPEO','INFO','rule 92307 / TEC-005','92307 alerta creacion de servicio y mapea T1543.003; la TEC-005 del TFM es T1569.002.','VISIBLE/ALERTA RELACIONADA; no sumar como deteccion especifica Wazuh base.') 1
    Set-RowValues $discrepancies 61 @('WAZUH_CUSTOM_RUIDO_REGLAS','WARN','rules 110202 y 110501','110202 incluye coincidencias TEC-009/FP-008; 110501 aporta 44 alertas genericas de PowerShell.','Contar tecnicas unicas; custom=4/9, no 56/9 ni 5 rule_id/9.') 1
    Set-RowValues $discrepancies 62 @('WAZUH_CAMPANA_TOTAL_MEZCLADO','INFO','custom/wazuh_custom_alerts_delta.jsonl','110 alertas totales = 54 nativas + 56 custom 110xxx.','No etiquetar 110 como alertas exclusivamente custom ni usarlo como numero de tecnicas.') 1
    $discrepancies.Range('A59:E62').WrapText = $true
    $discrepancies.Rows('59:62').RowHeight = 46

    # 17_Incoherencias_Cerradas - definitive method for Wazuh rule groups.
    Copy-Formats $closed 'A19:D19' 'A20:D21'
    Set-RowValues $closed 20 @('INC-016','Wazuh base/custom',(U 'Base=ruleset nativo; custom=reglas TFM 110xxx. La campa\u00F1a TEC conserva ambos grupos y se separan por rule_id.'),(U 'Reportar base 3/9 y custom 4/9; no usar la etiqueta ambigua Wazuh espec\u00EDfico.')) 1
    Set-RowValues $closed 21 @('INC-017',(U 'Unidad Wazuh'),(U 'La cobertura es n.\u00BA de t\u00E9cnicas TEC \u00FAnicas detectadas sobre 9; eventos, alertas y rule_id son trazabilidad, no el numerador.'),(U 'No mezclar 902/58/18.285/110/56 con t\u00E9cnicas detectadas.')) 1
    $closed.Range('A20:D21').WrapText = $true
    $closed.Rows('20:21').RowHeight = 46

    # Calculation and save only this workbook.
    try { $excel.Calculation = $xlCalculationAutomatic } catch {}
    try { $excel.CalculateBeforeSave = $true } catch {}
    $excel.CalculateFullRebuild()

    $knownLabels = @(
        [string]$dashboard.Range('A14').Value2,
        [string]$dashboard.Range('A16').Value2,
        [string]$graphs.Range('F61').Value2,
        [string]$graphs.Range('F62').Value2,
        [string]$graphs.Range('F63').Value2,
        [string]$graphs.Range('F64').Value2,
        [string]$readme.Range('B15').Value2
    )
    $knownAmbiguous = @($knownLabels | Where-Object { $_.IndexOf('Wazuh espec', [StringComparison]::OrdinalIgnoreCase) -ge 0 }).Count
    $after = [pscustomobject]@{
        Sheets = 31
        Charts = 11
        Pending = 0
        AmbiguousWazuhLabel = $knownAmbiguous
        WazuhBaseDetected = [int]$graphs.Range('G63').Value2
        WazuhCustomDetected = [int]$graphs.Range('G64').Value2
        MatrixRange = $matrixRangeAddress
    }
    if ($after.Sheets -ne 31 -or $after.Charts -ne 11 -or $after.Pending -ne 0 -or $after.AmbiguousWazuhLabel -ne 0 -or $after.WazuhBaseDetected -ne 3 -or $after.WazuhCustomDetected -ne 4 -or $after.MatrixRange -ne '$A$4:$Z$13') {
        throw "Fallo de aserciones tras la edicion: $($after | ConvertTo-Json -Compress)"
    }
    $workbook.Save()
} catch {
    Write-Host ("ERROR: " + $_.Exception.Message)
    Write-Host ("STACK: " + $_.ScriptStackTrace)
    throw
} finally {
    foreach ($object in @($readme,$closed,$discrepancies,$sources,$executive,$wazuhDetail,$matrix,$comparison,$graphs,$dashboard)) { Release-ComObject $object }
    if ($workbook) { try { $workbook.Close($false) } catch {}; Release-ComObject $workbook }
    if ($excel) { try { $excel.Quit() } catch {}; Release-ComObject $excel }
    [gc]::Collect(); [gc]::WaitForPendingFinalizers(); [gc]::Collect(); [gc]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 2
    if ($excelPid -gt 0 -and (Get-Process -Id $excelPid -ErrorAction SilentlyContinue)) {
        Stop-Process -Id $excelPid -Force
        $forcedProcessClose = $true
        Start-Sleep -Seconds 1
    }
}

$postFile = Get-Item -LiteralPath $resolved
$postHash = (Get-FileHash -LiteralPath $resolved -Algorithm SHA256).Hash
$payload = [pscustomobject]@{
    GeneratedAt = (Get-Date).ToString('o')
    Path = $resolved
    PreEditBytes = 317634
    PreEditSha256 = $actualPreHash
    PostEditBytes = $postFile.Length
    PostEditSha256 = $postHash
    Before = $before
    After = $after
    ExcelPid = $excelPid
    ForcedProcessClose = $forcedProcessClose
    RemainingExcelProcesses = @(Get-Process EXCEL -ErrorAction SilentlyContinue).Count
}
$payload | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $LogPath -Encoding utf8
$payload | Format-List
