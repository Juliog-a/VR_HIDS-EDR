param(
    [Parameter(Mandatory = $false)]
    [string]$Path = "C:\Users\julio\Desktop\TFM\08_MEMORIA\AUDITORIA_EXCELES_FINAL_20260712\TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712_CUSTOM_WORKING.xlsx"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$Root = "C:\Users\julio\Desktop\TFM"
$Normalized = Join-Path $Root "04_EVIDENCE\Excel_visibilidad_26062026\normalized"
$Validated = Join-Path $Root "01_ARTIFACTS\validated"
$AlertsPath = Join-Path $Normalized "alertas_vr.csv"
$VisibilityPath = Join-Path $Normalized "visibilidad_sistema.csv"
$TecPath = Join-Path $Normalized "tecnicas_real_vr.csv"
$FpPath = Join-Path $Normalized "fp_runner.csv"

if (-not (Test-Path -LiteralPath $Path)) { throw "No existe el libro temporal: $Path" }

function XlColor([int]$r, [int]$g, [int]$b) { return $r + (256 * $g) + (65536 * $b) }

$ColorNavy = XlColor 31 78 121
$ColorBlue = XlColor 68 114 196
$ColorTeal = XlColor 0 112 192
$ColorLightBlue = XlColor 221 235 247
$ColorPaleBlue = XlColor 242 247 252
$ColorGreen = XlColor 226 239 218
$ColorAmber = XlColor 255 242 204
$ColorRed = XlColor 244 204 204
$ColorGray = XlColor 242 242 242
$ColorDarkGray = XlColor 89 89 89
$ColorWhite = XlColor 255 255 255

function Write-Block {
    param($Worksheet, [int]$StartRow, [int]$StartColumn, [object[]]$Rows)
    if ($Rows.Count -eq 0) { return $null }
    $normalizedRows = @()
    foreach ($inputRow in $Rows) {
        $row = $inputRow
        while (($row -is [System.Array]) -and ($row.Count -eq 1) -and ($row[0] -is [System.Array])) {
            $row = $row[0]
        }
        if ($row -isnot [System.Array]) { throw "Write-Block recibió una fila escalar: $row" }
        $normalizedRows += ,$row
    }
    $maxCols = 0
    foreach ($row in $normalizedRows) {
        if ($row.Count -gt $maxCols) { $maxCols = $row.Count }
    }
    $range = $Worksheet.Range(
        $Worksheet.Cells.Item($StartRow, $StartColumn),
        $Worksheet.Cells.Item($StartRow + $normalizedRows.Count - 1, $StartColumn + $maxCols - 1)
    )
    for ($r = 0; $r -lt $normalizedRows.Count; $r++) {
        for ($c = 0; $c -lt $normalizedRows[$r].Count; $c++) {
            $value = $normalizedRows[$r][$c]
            while (($value -is [System.Array]) -and ($value.Count -eq 1)) { $value = $value[0] }
            if ($value -is [System.Array]) { $value = $value -join "; " }
            if ($null -ne $value) {
                if ($value -is [byte] -or $value -is [int16] -or $value -is [int32] -or $value -is [int64] -or $value -is [single] -or $value -is [double] -or $value -is [decimal]) {
                    $Worksheet.Cells.Item($StartRow + $r, $StartColumn + $c).Value2 = [double]$value
                } else {
                    $Worksheet.Cells.Item($StartRow + $r, $StartColumn + $c).NumberFormat = "@"
                    $Worksheet.Cells.Item($StartRow + $r, $StartColumn + $c).Value2 = [string]$value
                }
            }
        }
    }
    return $range
}

function Style-Title {
    param($Worksheet, [string]$Address, [string]$Text, [int]$FontSize = 18)
    $range = $Worksheet.Range($Address)
    $range.Merge()
    $range.Value2 = $Text
    $range.Interior.Color = $ColorNavy
    $range.Font.Color = $ColorWhite
    $range.Font.Bold = $true
    $range.Font.Size = $FontSize
    $range.HorizontalAlignment = -4131
    $range.VerticalAlignment = -4108
    $range.WrapText = $true
}

function Style-Subtitle {
    param($Worksheet, [string]$Address, [string]$Text)
    $range = $Worksheet.Range($Address)
    $range.Merge()
    $range.Value2 = $Text
    $range.Interior.Color = $ColorLightBlue
    $range.Font.Color = $ColorDarkGray
    $range.Font.Italic = $true
    $range.WrapText = $true
    $range.VerticalAlignment = -4108
}

function Style-Section {
    param($Worksheet, [string]$Address, [string]$Text)
    $range = $Worksheet.Range($Address)
    $range.Merge()
    $range.Value2 = $Text
    $range.Interior.Color = $ColorBlue
    $range.Font.Color = $ColorWhite
    $range.Font.Bold = $true
    $range.WrapText = $true
}

function Style-Header {
    param($Range)
    $Range.Interior.Color = $ColorNavy
    $Range.Font.Color = $ColorWhite
    $Range.Font.Bold = $true
    $Range.WrapText = $true
    $Range.HorizontalAlignment = -4108
    $Range.VerticalAlignment = -4108
}

function Delete-AllTables {
    param($Worksheet)
    while ($Worksheet.ListObjects.Count -gt 0) { $Worksheet.ListObjects.Item(1).Delete() }
}

function Add-Table {
    param($Worksheet, [string]$Address, [string]$Name, [string]$Style = "TableStyleMedium2")
    $table = $Worksheet.ListObjects.Add(1, $Worksheet.Range($Address), $null, 1)
    $table.Name = $Name
    $table.TableStyle = $Style
    return $table
}

function Configure-Sheet {
    param($Worksheet, [int]$SplitRow = 3, [int]$SplitColumn = 1, [int]$Zoom = 85)
    $Worksheet.Activate()
    $window = $Worksheet.Application.ActiveWindow
    $window.FreezePanes = $false
    $window.SplitRow = $SplitRow
    $window.SplitColumn = $SplitColumn
    $window.FreezePanes = $true
    $window.Zoom = $Zoom
    $Worksheet.Cells.Font.Name = "Aptos"
    $Worksheet.Cells.VerticalAlignment = -4160
}

function Configure-Print {
    param($Worksheet, [string]$Area, [int]$RepeatHeaderRow = 0)
    $Worksheet.PageSetup.Orientation = 2
    $Worksheet.PageSetup.Zoom = $false
    $Worksheet.PageSetup.FitToPagesWide = 1
    $Worksheet.PageSetup.FitToPagesTall = 1
    $Worksheet.PageSetup.PrintArea = $Area
    if ($RepeatHeaderRow -gt 0) { $Worksheet.PageSetup.PrintTitleRows = "`$$RepeatHeaderRow`:`$$RepeatHeaderRow" }
    $Worksheet.PageSetup.CenterHorizontally = $true
    $Worksheet.PageSetup.LeftMargin = $Worksheet.Application.InchesToPoints(0.25)
    $Worksheet.PageSetup.RightMargin = $Worksheet.Application.InchesToPoints(0.25)
    $Worksheet.PageSetup.TopMargin = $Worksheet.Application.InchesToPoints(0.4)
    $Worksheet.PageSetup.BottomMargin = $Worksheet.Application.InchesToPoints(0.4)
}

function Reset-ChartSeries {
    param($Chart)
    $series = $Chart.SeriesCollection()
    while ($series.Count -gt 0) { $series.Item(1).Delete() }
}

function Add-ChartSeries {
    param($Chart, [string]$Name, $Categories, $Values)
    $series = $Chart.SeriesCollection().NewSeries()
    $series.Name = $Name
    $series.XValues = $Categories
    $series.Values = $Values
    return $series
}

function Set-ChartBasics {
    param($ChartObject, [string]$Title, [int]$ChartType, [string]$XAxisTitle = "", [string]$YAxisTitle = "")
    $chart = $ChartObject.Chart
    $chart.ChartType = $ChartType
    $chart.HasTitle = $true
    $chart.ChartTitle.Text = $Title
    $chart.ChartTitle.Font.Size = 12
    $chart.ChartTitle.Font.Bold = $true
    $chart.HasLegend = $true
    $chart.Legend.Position = -4107
    $chart.PlotVisibleOnly = $false
    try {
        if ($XAxisTitle) {
            $axis = $chart.Axes(1)
            $axis.HasTitle = $true
            $axis.AxisTitle.Text = $XAxisTitle
        }
        if ($YAxisTitle) {
            $axis = $chart.Axes(2)
            $axis.HasTitle = $true
            $axis.AxisTitle.Text = $YAxisTitle
        }
    } catch { }
    return $chart
}

function Position-Chart {
    param($ChartObject, $Worksheet, [string]$TopLeft, [string]$BottomRight)
    $a = $Worksheet.Range($TopLeft)
    $b = $Worksheet.Range($BottomRight)
    $ChartObject.Left = $a.Left
    $ChartObject.Top = $a.Top
    $ChartObject.Width = ($b.Left + $b.Width) - $a.Left
    $ChartObject.Height = ($b.Top + $b.Height) - $a.Top
    $ChartObject.Placement = 1
}

function Principal-Artifact([string]$Profile) {
    switch ($Profile) {
        "P1" { return "Custom.TFM.HIDS.P1.Critical.Priority.Event_v1" }
        "P2" { return "Custom.TFM.HIDS.P2.High.Forensic.Event_v1" }
        "P3" { return "Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1" }
        "P4" { return "Custom.TFM.HIDS.P4.Low.Basic_v2" }
        default { return "" }
    }
}

$alerts = @(Import-Csv -LiteralPath $AlertsPath)
$visibility = @(Import-Csv -LiteralPath $VisibilityPath)
$tecs = @(Import-Csv -LiteralPath $TecPath)
$fps = @(Import-Csv -LiteralPath $FpPath)

$profileCounts = @{}
foreach ($p in "P1", "P2", "P3", "P4") { $profileCounts[$p] = @($alerts | Where-Object profile -eq $p).Count }

$excel = $null
$workbook = $null
$initialAlerts = $alerts.Count

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    $workbook = $excel.Workbooks.Open($Path, 0, $false)
    try { $excel.Calculation = -4135 } catch { }

    # 00_Guia
    $ws = $workbook.Worksheets.Item("00_Guia")
    try { $ws.Range("A1:C2").UnMerge() } catch { }
    Style-Title $ws "A1:C1" "Guía de uso · Visibilidad, detección, alertas y falsos positivos"
    Style-Subtitle $ws "A2:C2" "Libro maestro de eficacia HIDS/DFIR. Las medidas de rendimiento del sistema pertenecen a un entregable independiente."
    $ws.Range("B10").Value2 = "Resumen calculado de cobertura, detección y gaps"
    $ws.Range("C10").Value2 = "Usar para memoria"
    $ws.Range("B11").Value2 = "Catálogo completo: artifacts custom P1-P4, públicos y auxiliares"
    $ws.Range("C11").Value2 = "Trazar artifact principal, source interna, evidencia y rol metodológico"
    $ws.Range("A12").Value2 = "08_Benignas_FP"
    $ws.Range("B12").Value2 = "Operaciones benignas y resultados de falsos positivos"
    $ws.Range("C12").Value2 = "Separar FP custom de coincidencias públicas"
    $ws.Range("A13").Value2 = "09_Resultados_Custom"
    $ws.Range("B13").Value2 = "Resultados reales custom por perfil, técnica y source interna"
    $ws.Range("C13").Value2 = "Analizar eficacia CLIENT_EVENT, cobertura TEC y FP sin métricas de rendimiento"
    $ws.Columns("A").ColumnWidth = 28
    $ws.Columns("B").ColumnWidth = 62
    $ws.Columns("C").ColumnWidth = 50
    $ws.Rows("1:14").AutoFit()
    Configure-Sheet $ws 3 0 90
    Configure-Print $ws '$A$1:$C$14' 3

    # 03_Evaluacion_VR: artifact names exactos y estado de detección real.
    $ws = $workbook.Worksheets.Item("03_Evaluacion_VR")
    $artifactByTec = @{
        "TEC-001" = "Custom.TFM.HIDS.P2.High.Forensic.Event_v1; Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1; Custom.TFM.HIDS.P4.Low.Basic_v2"
        "TEC-002" = "Custom.TFM.HIDS.P2.High.Forensic.Event_v1; Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1; Custom.TFM.HIDS.P4.Low.Basic_v2"
        "TEC-003" = "Custom.TFM.HIDS.P2.High.Forensic.Event_v1; Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1; Custom.TFM.HIDS.P4.Low.Basic_v2"
        "TEC-004" = "Custom.TFM.HIDS.P1.Critical.Priority.Event_v1; Custom.TFM.HIDS.P2.High.Forensic.Event_v1; Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1; Custom.TFM.HIDS.P4.Low.Basic_v2"
        "TEC-005" = "Custom.TFM.HIDS.P1.Critical.Priority.Event_v1; Custom.TFM.HIDS.P2.High.Forensic.Event_v1; Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1; Custom.TFM.HIDS.P4.Low.Basic_v2"
        "TEC-006" = "Custom.TFM.HIDS.P2.High.Forensic.Event_v1; Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1"
        "TEC-007" = "Custom.TFM.HIDS.P1.Critical.Priority.Event_v1; Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1"
        "TEC-008" = "Custom.TFM.HIDS.P1.Critical.Priority.Event_v1; Custom.TFM.HIDS.P2.High.Forensic.Event_v1; Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1"
        "TEC-009" = "Custom.TFM.HIDS.P1.Critical.Priority.Event_v1; Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1"
    }
    for ($r = 5; $r -le 13; $r++) {
        $tec = [string]$ws.Cells.Item($r, 2).Value2
        $profiles = @($alerts | Where-Object TEC -eq $tec | Select-Object -ExpandProperty profile -Unique | Sort-Object)
        $principals = @($profiles | ForEach-Object { Principal-Artifact $_ })
        $ws.Cells.Item($r, 6).Value2 = $principals -join "; "
        $ws.Cells.Item($r, 7).Value2 = "Custom validado"
        $ws.Cells.Item($r, 8).Value2 = "CLIENT_EVENT"
        $ws.Cells.Item($r, 9).Value2 = "DETECTADA"
        $sources = @($alerts | Where-Object TEC -eq $tec | Select-Object -ExpandProperty Artifact -Unique)
        $ws.Cells.Item($r, 11).Value2 = "ALERTAS_VR / alertas_vr.csv; sources: " + ($sources -join "; ")
    }
    $ws.Columns("F").ColumnWidth = 58
    $ws.Columns("K").ColumnWidth = 62
    $ws.Range("A3:O12").WrapText = $true
    $ws.Rows("3:12").AutoFit()
    Configure-Sheet $ws 3 2 80
    Configure-Print $ws '$A$1:$O$12' 3

    # 06_Resumen: reconstrucción del resumen sin contenido de rendimiento.
    $ws = $workbook.Worksheets.Item("06_Resumen")
    Delete-AllTables $ws
    $ws.Cells.Clear()
    Style-Title $ws "A1:C1" "Resumen definitivo · eficacia de detección y trazabilidad"
    Style-Subtitle $ws "A2:C2" "Visibilidad, detección, alerta CLIENT_EVENT, evidencia forense, falsos positivos y comparación Wazuh se mantienen como capas separadas."
    $summaryRows = @(
        @("Indicador", "Valor", "Interpretación / evidencia"),
        @("Artifacts custom validados", 5, "Cuatro detectores CLIENT_EVENT P1-P4 y un Router SERVER_EVENT de transporte."),
        @("Detectores custom CLIENT_EVENT", 4, "P1, P2, P3 y P4; nombres exactos y SHA-256 en 07_Catalogo_Artifacts."),
        @("Sources detectoras declaradas", 27, "P1=4; P2=5; P3=13; P4=5."),
        @("Sources con alertas en campaña final", 25, "P2 Security ScheduledTask Optional y P3 ZIP FileCreate no generaron filas en ventana."),
        @("Técnicas custom detectadas", "9/9", "379 alertas CLIENT_EVENT consolidadas; volumen de alertas no equivale a cobertura."),
        @("Distribución de alertas", "P1=19; P2=118; P3=109; P4=133", "Suma=379; recuentos obtenidos de ALERTAS_VR."),
        @("Campaña FP custom", "10/10 OK; 0 hits", "No se mezcla con las cuatro coincidencias High de Hayabusa Monitoring."),
        @("Campañas públicas", "5 positivas / 1 no concluyente", "Hayabusa CH detecta 3/9; ETW permanece NO CONCLUYENTE."),
        @("Wazuh base", "3/9", "TEC-002/004/006 mediante ruleset nativo."),
        @("Wazuh custom", "4/9", "TEC-001/005/007/008 mediante reglas TFM 110xxx."),
        @("TEC-009", "HTTP 200 por runner", "TrackNetwork aporta visibilidad de conexión; el resultado HTTP no se atribuye al CLIENT_EVENT."),
        @("Cobertura táctica", "5/14; 6/14; 7/14", "/14 son tácticas ATT&CK Enterprise consideradas; no las 9 técnicas TEC."),
        @("Arquitectura metodológica", "CLIENT_EVENT ≠ SERVER_EVENT", "El detector alerta; Router/Discord/webhook transportan o notifican."),
        @("Conclusión", "Custom necesario", "Añade semántica, severidad, trazabilidad y control de ruido a la telemetría disponible.")
    )
    Write-Block $ws 3 1 $summaryRows | Out-Null
    Add-Table $ws "A3:C17" "ResumenEficaciaTable" | Out-Null
    $ws.Columns("A").ColumnWidth = 34
    $ws.Columns("B").ColumnWidth = 27
    $ws.Columns("C").ColumnWidth = 86
    $ws.Range("A3:C17").WrapText = $true
    $ws.Rows("1:17").AutoFit()
    Configure-Sheet $ws 3 0 90
    Configure-Print $ws '$A$1:$C$17' 3

    # 07_Catalogo_Artifacts: dos bloques y catálogo custom completo.
    $ws = $workbook.Worksheets.Item("07_Catalogo_Artifacts")
    Delete-AllTables $ws
    $ws.Cells.Clear()
    Style-Title $ws "A1:U1" "Catálogo completo de artifacts custom, públicos y auxiliares"
    Style-Subtitle $ws "A2:U2" "Artifact principal y source interna son niveles distintos. CLIENT_EVENT detecta; SERVER_EVENT, Router, Discord, webhook y receiver transportan o acreditan salida externa."
    Style-Section $ws "A3:U3" "Bloque A · Artifacts custom TFM validados"
    $catalogHeader = @("Artifact", "Categoría", "Tipo", "Modo", "Perfil", "Sources internas", "Técnicas TEC", "MITRE ATT&CK", "Fuente de telemetría", "Event IDs", "Función metodológica", "Uso en campaña final", "Técnicas detectadas", "Alertas CLIENT_EVENT", "FP hits", "Evidencia principal", "Fichero validado", "SHA-256", "Estado", "Decisión", "Limitaciones")
    $customRows = @(
        $catalogHeader,
        @("Custom.TFM.HIDS.P1.Critical.Priority.Event_v1", "Custom TFM", "CLIENT_EVENT", "Client Monitoring", "P1", "4: P1_EVENT_PowerShell4104_Basic; P1_EVENT_Sysmon_ID1_Process_Basic; P1_EVENT_System7045_ServiceCreation; P1_EVENT_TEC009_Sysmon_Context", "TEC-004/005/007/008/009", "T1547.001; T1569.002; T1486; T1485; T1048.003/T1560.001", "PowerShell, Sysmon, System", "4104; 1; 3; 11; 7045", "Correlación/prioridad crítica", "Sí", "5/9", 19, 0, "normalized/alertas_vr.csv; ALERTAS_VR", "Custom.TFM.HIDS.P1.Critical.Priority.Event_v1.yaml", "8A1FF2654405AF9E4412D32EF0C2BE08A383924654AA764A04AB264BAA42EDD7", "VALIDADO", "Mantener", "Señales de alta prioridad; el volumen no equivale a cobertura."),
        @("Custom.TFM.HIDS.P2.High.Forensic.Event_v1", "Custom TFM", "CLIENT_EVENT", "Client Monitoring", "P2", "5 declaradas; 4 con alertas. Incluye PowerShell, Process, 7045, Registry/File/Network y ScheduledTask opcional", "TEC-001 a TEC-009 declaradas", "T1059.001; T1059.003; T1053.005; T1547.001; T1569.002; T1518.001; T1486; T1485; T1048.003", "PowerShell, Sysmon, System, Security opcional", "4104; 1; 3; 11; 12/13/14; 23/26; 7045; 4698/4699", "Evidencia forense y contexto", "Sí", "7/9", 118, 0, "normalized/alertas_vr.csv; ALERTAS_VR", "Custom.TFM.HIDS.P2.High.Forensic.Event_v1.yaml", "255C5F7CA39E362FE45106C662C4C35730EFF57A4063A1E9973FD11A6375FAA2", "VALIDADO", "Mantener", "Source Security ScheduledTask Optional sin filas en la ventana final."),
        @("Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1", "Custom TFM", "CLIENT_EVENT", "Client Monitoring", "P3", "13 declaradas; 12 con alertas. Sources específicas TEC001-TEC009", "TEC-001 a TEC-009", "T1059.001; T1059.003; T1053.005; T1547.001; T1569.002; T1518.001; T1486; T1485; T1048.003", "PowerShell, Sysmon, System", "4104; 1; 3; 7045", "Detección de comportamiento", "Sí", "9/9", 109, 0, "normalized/alertas_vr.csv; ALERTAS_VR", "Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1.yaml", "C591CB849753194F6F87834C1E0F34EDE6CDFF6B31BC63650A1D281ADA1D410A", "VALIDADO", "Mantener", "TEC009_ZIP_FileCreate_Sysmon_Context no generó filas en ventana; otras sources cubren TEC-009."),
        @("Custom.TFM.HIDS.P4.Low.Basic_v2", "Custom TFM", "CLIENT_EVENT", "Client Monitoring", "P4", "5: CU001 PowerShell; CU002 CMD; CU003 ScheduledTask; CU004 RunKeys; CU005 ServiceCreation", "TEC-001/002/003/004/005", "T1059.001; T1059.003; T1053.005; T1547.001; T1569.002", "Sysmon y System", "1; 7045", "Señales básicas de baja complejidad", "Sí", "5/9", 133, 0, "normalized/alertas_vr.csv; ALERTAS_VR", "Custom.TFM.HIDS.P4.Low.Basic_v2.yaml", "1EBB29296A14475E38BCF34CA8DD9CC16572F4A8CA00A1AB0C3608E0B3097BC1", "VALIDADO", "Mantener", "Telemetría básica; requiere perfiles superiores para semántica completa.")
    )
    Write-Block $ws 4 1 $customRows | Out-Null
    Add-Table $ws "A4:U8" "CatalogoCustomTable" | Out-Null
    Style-Section $ws "A10:U10" "Bloque B · Artifacts públicos, históricos, transporte y evidencia externa"
    $publicRows = @(
        $catalogHeader,
        @("Windows.Events.ProcessCreation", "Público", "CLIENT_EVENT", "Client Monitoring", "Visibilidad", "N/A", "TEC-001/002/003/004/006; contexto 007/008/009", "Varios", "Windows/Sysmon proceso", "1", "Telemetría de proceso", "Sí; campaña positiva", "0/9 específicas", "Telemetría", "No aplicable", "Campaña pública normalizada", "Built-in", "No aplica", "POSITIVO", "Mantener", "Visibilidad no equivale a detección semántica."),
        @("Windows.Events.ServiceCreation", "Público", "CLIENT_EVENT", "Client Monitoring", "Visibilidad", "N/A", "TEC-005", "T1569.002 contextual", "System", "7045", "Visibilidad de creación de servicio", "Sí; campaña positiva", "0/9 específicas", "1 evento 7045", "No aplicable", "Campaña pública normalizada", "Built-in", "No aplica", "POSITIVO", "Mantener", "El evento 7045 requiere interpretación; no se suma a Hayabusa CH."),
        @("Windows.Sysinternals.SysmonLogForward", "Público", "CLIENT_EVENT", "Client Monitoring", "Visibilidad/forense", "N/A", "TEC-004/007/008/009", "Varios", "Sysmon", "1/3/11/12/13/14/26", "Reenvío de telemetría", "Sí; campaña positiva", "0/9 específicas", "Telemetría", "No aplicable", "Campaña pública normalizada", "Built-in", "No aplica", "POSITIVO", "Mantener", "ID26 es evidencia forense, no alerta individual."),
        @("Windows.Hayabusa.Monitoring", "Público", "CLIENT_EVENT", "Client Monitoring CH", "P3", "Reglas Sigma/Hayabusa", "TEC-001/003/005", "T1059.001; T1053.005; T1569.002", "EVTX/Sigma", "Varios", "Detección específica pública", "Sí; campaña positiva", "3/9", "7 High en TEC", "4 High / 1 caso", "Campaña Hayabusa CH", "Artifact público", "No aplica", "POSITIVO", "Mantener CH", "CHM no ejecutada; FP público separado del runner custom."),
        @("Windows.ETW.Monitoring", "Público", "CLIENT_EVENT", "Client Monitoring", "Experimental", "N/A", "No evaluable", "No evaluable", "ETW", "No determinado", "Visibilidad potencial", "Campaña incompleta", "NO EVALUABLE", 0, "No aplicable", "Evidencia pública incompleta", "Artifact público", "No aplica", "NO CONCLUYENTE", "Documentar", "Faltan runner/log/summary TEC y parámetros."),
        @("Generic.Events.TrackNetworkConnections", "Público", "CLIENT_EVENT", "Client Monitoring", "Contexto", "N/A", "TEC-009 contextual", "T1048.003 contextual", "netstat/diff", "N/A", "Visibilidad de conexión", "Sí; campaña positiva", "0/9 específicas", "3 filas destino", "No aplicable", "Conexión 192.168.1.129:8088", "Artifact público", "No aplica", "POSITIVO", "Mantener con filtro", "PID=0 y ProcInfo vacío; sin atribución fiable a powershell.exe."),
        @("Windows.Hayabusa.Rules", "Histórico", "HUNT", "Hunt histórico", "No RT", "Reglas Sigma/Hayabusa", "Histórico", "Varios", "EVTX", "Varios", "Análisis histórico", "No ejecutado como campaña final independiente", "No incluido", "No RT", "No aplicable", "Resultados históricos", "Artifact público", "No aplica", "FUERA DE CAMPAÑA", "Conservar como contraste", "No es alerta CLIENT_EVENT en tiempo real."),
        @("Server.Alerts.TrackNetworkConnections", "Transporte", "SERVER_EVENT", "Server Monitoring", "Transporte", "N/A", "No aplica", "No aplica", "Alertas servidor", "N/A", "Enrutamiento", "No incluido", "No aplica", "No detector", "No aplicable", "Arquitectura pública", "Companion", "No aplica", "TRANSPORTE", "No contar", "No detector CLIENT_EVENT."),
        @("Custom.TFM.HIDS.Router.JSONL.Discord.SOC_v3", "Custom auxiliar", "SERVER_EVENT", "Server Monitoring", "Transporte", "27 routes", "Todas las alertables", "Hereda del detector", "watch_monitoring", "N/A", "Normalización JSONL/Discord", "No acreditado como detector", "No aplica", "No detector", "No aplicable", "YAML validado", "Custom.TFM.HIDS.Router.JSONL.Discord.SOC_v3.yaml", "574E3E8BC03E7C4D655400D5BB5527771E438CA67C413853A0AAAC7A10C5E33F", "VALIDADO / TRANSPORTE", "Mantener separado", "Transporta; no detecta ni suma cobertura."),
        @("Receiver HTTP TEC-009", "Evidencia externa", "AUXILIAR", "Runner independiente", "Salida externa", "N/A", "TEC-009", "T1048.003 contextual", "Receptor HTTP", "N/A", "Acreditar upload externo", "Sí, por runner", "No aplica", "HTTP 200", "No aplicable", "HTTPStatus=200; ZIP=6245 B; SHA-256 coincidente", "Summary de runner", "Trazado en FUENTES", "EVIDENCIA", "Conservar separado", "No es artifact detector ni capacidad de TrackNetworkConnections.")
    )
    Write-Block $ws 11 1 $publicRows | Out-Null
    Add-Table $ws "A11:U21" "CatalogoPublicoAuxTable" "TableStyleMedium9" | Out-Null
    $ws.Range("A4:U21").WrapText = $true
    $ws.Columns("A").ColumnWidth = 42
    $ws.Columns("B:E").ColumnWidth = 18
    $ws.Columns("F").ColumnWidth = 48
    $ws.Columns("G:K").ColumnWidth = 27
    $ws.Columns("L:P").ColumnWidth = 25
    $ws.Columns("Q").ColumnWidth = 43
    $ws.Columns("R").ColumnWidth = 67
    $ws.Columns("S:U").ColumnWidth = 25
    $ws.Rows("1:21").AutoFit()
    Configure-Sheet $ws 4 2 75
    Configure-Print $ws '$A$1:$U$21' 4

    # 09_Resultados_Custom: sustitución completa de la hoja de rendimiento.
    $ws = $workbook.Worksheets.Item("09_Benchmark_Plan")
    Delete-AllTables $ws
    while ($ws.ChartObjects().Count -gt 0) { $ws.ChartObjects().Item(1).Delete() }
    $ws.Cells.Clear()
    $ws.Name = "09_Resultados_Custom"
    Style-Title $ws "A1:K1" "Resultados reales de eficacia · artifacts custom P1-P4"
    Style-Subtitle $ws "A2:K2" "Recuentos derivados de ALERTAS_VR y de las campañas TEC/FP. No contiene medidas de CPU, memoria, E/S ni consumo por proceso."
    Style-Section $ws "A3:K3" "Tabla 1 · Resumen por perfil"
    $profileRows = @(
        @("Perfil", "Artifact principal", "Sources internas ejecutadas", "Técnicas cubiertas declaradas", "Técnicas detectadas", "Alertas CLIENT_EVENT", "Porcentaje sobre total", "FP hits", "Función principal", "Evidencia", "Limitaciones"),
        @("P1", (Principal-Artifact "P1"), 4, 5, 5, $profileCounts["P1"], ([double]$profileCounts["P1"] / [double]$initialAlerts), 0, "Correlación y prioridad crítica", "ALERTAS_VR / alertas_vr.csv", "Volumen bajo y selectivo; no equivale a menor cobertura global."),
        @("P2", (Principal-Artifact "P2"), 4, 9, 7, $profileCounts["P2"], ([double]$profileCounts["P2"] / [double]$initialAlerts), 0, "Contexto forense", "ALERTAS_VR / alertas_vr.csv", "Source Security ScheduledTask opcional sin alertas en ventana."),
        @("P3", (Principal-Artifact "P3"), 12, 9, 9, $profileCounts["P3"], ([double]$profileCounts["P3"] / [double]$initialAlerts), 0, "Detección de comportamiento", "ALERTAS_VR / alertas_vr.csv", "ZIP FileCreate sin filas; TEC-009 detectada por otras sources."),
        @("P4", (Principal-Artifact "P4"), 5, 5, 5, $profileCounts["P4"], ([double]$profileCounts["P4"] / [double]$initialAlerts), 0, "Señal básica", "ALERTAS_VR / alertas_vr.csv", "Alta telemetría de proceso; requiere perfiles superiores para semántica."),
        @("TOTAL", "4 detectores CLIENT_EVENT", 25, 9, 9, $initialAlerts, 1.0, 0, "Cobertura custom completa", "dataset_summary.json", "La suma de alertas es 379; una alerta no equivale a una técnica.")
    )
    Write-Block $ws 4 1 $profileRows | Out-Null
    $ws.Range("G5:G9").NumberFormat = "0.0%"
    Add-Table $ws "A4:K9" "ResultadosCustomPerfilTable" | Out-Null

    Style-Section $ws "A11:J11" "Tabla 2 · Matriz TEC × perfil"
    $matrixHeader = @("TEC", "P1", "P2", "P3", "P4", "Total alertas", "Perfil máximo", "Sources principales", "Señal principal", "Estado de detección")
    $matrixRows = New-Object System.Collections.Generic.List[object]
    $matrixRows.Add($matrixHeader)
    $signalByTec = @{
        "TEC-001" = "PowerShell 4104 / ProcessCreation"
        "TEC-002" = "Sysmon ID 1 / CMD"
        "TEC-003" = "Sysmon ID 1 / Scheduled Task"
        "TEC-004" = "Sysmon ID 1 y 12/13"
        "TEC-005" = "System 7045 / Sysmon ID 1"
        "TEC-006" = "PowerShell 4104 / Sysmon ID 1"
        "TEC-007" = "PowerShell 4104"
        "TEC-008" = "PowerShell 4104 / contexto Sysmon"
        "TEC-009" = "PowerShell 4104 / Sysmon ID 1 y 3"
    }
    for ($i = 1; $i -le 9; $i++) {
        $tec = "TEC-{0:D3}" -f $i
        $rowAlerts = @($alerts | Where-Object TEC -eq $tec)
        $counts = @()
        foreach ($p in "P1", "P2", "P3", "P4") { $counts += @($rowAlerts | Where-Object profile -eq $p).Count }
        $max = ($counts | Measure-Object -Maximum).Maximum
        $maxProfile = ("P1", "P2", "P3", "P4")[[Array]::IndexOf($counts, $max)]
        $sources = @($rowAlerts | Select-Object -ExpandProperty Artifact -Unique)
        $matrixRows.Add(@($tec, $counts[0], $counts[1], $counts[2], $counts[3], $rowAlerts.Count, $maxProfile, ($sources -join "; "), $signalByTec[$tec], "DETECTADA"))
    }
    Write-Block $ws 12 1 $matrixRows.ToArray() | Out-Null
    Add-Table $ws "A12:J21" "ResultadosCustomTecPerfilTable" "TableStyleMedium4" | Out-Null

    Style-Section $ws "A23:L23" "Tabla 3 · Resultados por source interna y técnica"
    $sourceHeader = @("Artifact principal", "Source interna", "Perfil", "TEC", "MITRE", "EventID", "SignalType", "Alertas", "Campaña", "Ventana UTC", "Source file", "Validación")
    $sourceRows = New-Object System.Collections.Generic.List[object]
    $sourceRows.Add($sourceHeader)
    $groups = $alerts | Group-Object profile, Artifact, TEC, MITRE, EventID, SignalType | Sort-Object { $_.Group[0].profile }, { $_.Group[0].Artifact }, { $_.Group[0].TEC }, { $_.Group[0].EventID }
    foreach ($group in $groups) {
        $g = @($group.Group)
        $first = $g[0]
        $times = @($g | Select-Object -ExpandProperty timestamp | Sort-Object)
        $sourceFiles = @($g | Select-Object -ExpandProperty source_file -Unique)
        $validation = @($g | Select-Object -ExpandProperty validacion -Unique) -join "; "
        $sourceRows.Add(@(
            (Principal-Artifact $first.profile), $first.Artifact, $first.profile, $first.TEC, $first.MITRE,
            $first.EventID, $first.SignalType, $g.Count, $first.campana,
            ($times[0] + " .. " + $times[$times.Count - 1]), ($sourceFiles -join "; "), $validation
        ))
    }
    Write-Block $ws 24 1 $sourceRows.ToArray() | Out-Null
    $sourceEnd = 24 + $sourceRows.Count - 1
    Add-Table $ws ("A24:L{0}" -f $sourceEnd) "ResultadosCustomSourceTable" "TableStyleMedium5" | Out-Null

    $testTitleRow = $sourceEnd + 2
    Style-Section $ws ("A{0}:I{0}" -f $testTitleRow) "Tabla 4 · Pruebas reales y falsos positivos custom"
    $testRows = @(
        @("Campaña", "Planificadas", "Ejecutadas", "OK", "WARN", "FAIL", "Hits", "Evidencia", "Lectura"),
        @("TEC real", $tecs.Count, $tecs.Count, @($tecs | Where-Object estado_ejecucion -eq "OK").Count, @($tecs | Where-Object estado_ejecucion -eq "WARN").Count, @($tecs | Where-Object estado_ejecucion -eq "FAIL").Count, "No aplica", "tecnicas_real_vr.csv / dataset_summary.json", "9/9 detectadas por CLIENT_EVENT custom."),
        @("FP custom", $fps.Count, $fps.Count, @($fps | Where-Object estado -eq "OK").Count, @($fps | Where-Object estado -eq "WARN").Count, @($fps | Where-Object estado -eq "FAIL").Count, @($fps | Where-Object genero_alerta_vr -match "^SI").Count, "fp_runner.csv / fp_hits.csv / summary.json", "10/10 OK; 0 hits. Hayabusa FP queda separado.")
    )
    Write-Block $ws ($testTitleRow + 1) 1 $testRows | Out-Null
    $testEnd = $testTitleRow + 3
    Add-Table $ws ("A{0}:I{1}" -f ($testTitleRow + 1), $testEnd) "ResultadosCustomPruebasTable" "TableStyleMedium6" | Out-Null
    $ws.Range("A1:L$testEnd").WrapText = $true
    $ws.Columns("A:B").ColumnWidth = 45
    $ws.Columns("C:G").ColumnWidth = 16
    $ws.Columns("H:I").ColumnWidth = 48
    $ws.Columns("J:L").ColumnWidth = 55
    $ws.Rows("1:$testEnd").AutoFit()
    Configure-Sheet $ws 4 2 75
    Configure-Print $ws ("`$A`$1:`$L`$$testEnd") 4

    # 10_Gaps_Custom_P1P4: sustituir lenguaje de rendimiento por diseño operativo.
    $ws = $workbook.Worksheets.Item("10_Gaps_Custom_P1P4")
    $ws.Range("A2:I2").Merge()
    $ws.Range("A2").Value2 = "Justificación técnica: los públicos aportan visibilidad/alerta parcial; los custom aportan semántica, CU_ID, priorización, mantenibilidad y reducción de ruido."
    $ws.Range("C8").Value2 = "Un artifact monolítico dificulta granularidad, mantenibilidad, priorización y ajuste de ruido"
    $ws.Range("H8").Value2 = "Granularidad por riesgo, trazabilidad y control operativo"
    $ws.Range("A1:I10").WrapText = $true
    $ws.Rows("1:10").AutoFit()

    # 14_Arquitectura_Custom: nombres, ficheros y hashes reales; sin P0 inventado.
    $ws = $workbook.Worksheets.Item("14_Arquitectura_Custom")
    $ws.Cells.Clear()
    Style-Title $ws "A1:I1" "Arquitectura custom validada P4-P1"
    Style-Subtitle $ws "A2:I2" "Cuatro artifacts CLIENT_EVENT principales con sources internas por técnica. La complejidad es cualitativa y no representa una medida de rendimiento. No se localizó P0 en validated."
    $archRows = @(
        @("Perfil", "Nombre exacto declarado", "Fichero validado", "SHA-256", "Tipo", "Sources internas", "Técnicas declaradas", "Uso campaña final", "Limitaciones"),
        @("P1", (Principal-Artifact "P1"), "Custom.TFM.HIDS.P1.Critical.Priority.Event_v1.yaml", "8A1FF2654405AF9E4412D32EF0C2BE08A383924654AA764A04AB264BAA42EDD7", "CLIENT_EVENT", "4", "TEC-004/005/007/008/009", "Sí; 19 alertas", "Complejidad operativa cualitativa alta; señales selectivas."),
        @("P2", (Principal-Artifact "P2"), "Custom.TFM.HIDS.P2.High.Forensic.Event_v1.yaml", "255C5F7CA39E362FE45106C662C4C35730EFF57A4063A1E9973FD11A6375FAA2", "CLIENT_EVENT", "5 declaradas / 4 ejecutadas", "TEC-001 a TEC-009", "Sí; 118 alertas", "ScheduledTask Security opcional sin filas en ventana."),
        @("P3", (Principal-Artifact "P3"), "Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1.yaml", "C591CB849753194F6F87834C1E0F34EDE6CDFF6B31BC63650A1D281ADA1D410A", "CLIENT_EVENT", "13 declaradas / 12 ejecutadas", "TEC-001 a TEC-009", "Sí; 109 alertas", "ZIP FileCreate sin filas; otras sources cubren TEC-009."),
        @("P4", (Principal-Artifact "P4"), "Custom.TFM.HIDS.P4.Low.Basic_v2.yaml", "1EBB29296A14475E38BCF34CA8DD9CC16572F4A8CA00A1AB0C3608E0B3097BC1", "CLIENT_EVENT", "5", "TEC-001/002/003/004/005", "Sí; 133 alertas", "Complejidad operativa cualitativa baja; telemetría básica."),
        @("P0", "No incorporado", "No existe fichero P0 en 01_ARTIFACTS/validated", "No aplica", "LABORATORIO/DEBUG", "0", "No aplica", "No utilizado", "No se inventa artifact ni ejecución.")
    )
    Write-Block $ws 3 1 $archRows | Out-Null
    Style-Header $ws.Range("A3:I3")
    Style-Section $ws "A10:I10" "Relación entre artifact principal y sources internas ejecutadas"
    $sourceSummaryRows = @(
        @("Perfil", "Sources ejecutadas", "Sources declaradas sin filas", "Técnicas con alertas", "Señales"),
        @("P1", "P1_EVENT_PowerShell4104_Basic; P1_EVENT_Sysmon_ID1_Process_Basic; P1_EVENT_System7045_ServiceCreation; P1_EVENT_TEC009_Sysmon_Context", "Ninguna", "TEC-004/005/007/008/009", "4104; Sysmon 1/3; System 7045"),
        @("P2", "PowerShell4104 Strong; Sysmon ID1; System 7045; Registry/File/Network Context", "Forensic_EVENT_Security_ScheduledTask_Optional", "TEC-001/002/003/004/005/006/008", "4104; Sysmon 1/12/13; System 7045"),
        @("P3", "12 sources TEC001-TEC009 con filas en ALERTAS_VR", "TEC009_ZIP_FileCreate_Sysmon_Context", "TEC-001 a TEC-009", "4104; Sysmon 1/3; System 7045"),
        @("P4", "CU001; CU002; CU003; CU004; CU005", "Ninguna", "TEC-001/002/003/004/005", "ProcessCreation; System 7045")
    )
    Write-Block $ws 11 1 $sourceSummaryRows | Out-Null
    Style-Header $ws.Range("A11:E11")
    $ws.Range("A1:I15").WrapText = $true
    $ws.Columns("A").ColumnWidth = 12
    $ws.Columns("B:D").ColumnWidth = 46
    $ws.Columns("E:I").ColumnWidth = 29
    $ws.Rows("1:15").AutoFit()
    Configure-Sheet $ws 3 2 80
    Configure-Print $ws '$A$1:$I$15' 3

    # VISIBILIDAD_SISTEMA: matriz validada, totales y unidad explícita.
    $ws = $workbook.Worksheets.Item("VISIBILIDAD_SISTEMA")
    $ws.Cells.Clear()
    Style-Title $ws "A1:N1" "Visibilidad del sistema por fuente y técnica"
    Style-Subtitle $ws "A2:N2" "Los recuentos son filas de telemetría/evidencia en ventana, no técnicas detectadas. Receiver HTTP se mantiene separado de la telemetría host."
    $visHeader = @("Fuente", "TEC-001", "TEC-002", "TEC-003", "TEC-004", "TEC-005", "TEC-006", "TEC-007", "TEC-008", "TEC-009", "Total fuente", "Origen", "Unidad", "Observaciones metodológicas")
    $visRows = New-Object System.Collections.Generic.List[object]
    $visRows.Add($visHeader)
    foreach ($v in $visibility) {
        $values = @()
        foreach ($i in 1..9) { $values += [int]$v.("TEC-{0:D3}" -f $i) }
        $origin = if ($v.fuente -eq "receiver HTTP") { "summary/runner TEC-009" } else { "normalized/visibilidad_sistema.csv" }
        $unit = if ($v.fuente -eq "receiver HTTP") { "Confirmación externa" } else { "Filas de evento en ventana" }
        $note = switch ($v.fuente) {
            "Sysmon ID 26" { "Evidencia forense; no alerta individual." }
            "receiver HTTP" { "HTTP 200 del runner independiente; no detector host." }
            default { "Telemetría visible; no equivale automáticamente a detección." }
        }
        $visRows.Add(@($v.fuente) + $values + @(($values | Measure-Object -Sum).Sum, $origin, $unit, $note))
    }
    Write-Block $ws 3 1 $visRows.ToArray() | Out-Null
    $hostTotalRow = 12
    $generalTotalRow = 13
    $ws.Cells.Item($hostTotalRow, 1).Value2 = "TOTAL HOST (sin receiver)"
    $ws.Cells.Item($generalTotalRow, 1).Value2 = "TOTAL GENERAL"
    for ($c = 2; $c -le 11; $c++) {
        $col = [char](64 + $c)
        $ws.Cells.Item($hostTotalRow, $c).Formula = "=SUM(${col}4:${col}10)"
        $ws.Cells.Item($generalTotalRow, $c).Formula = "=SUM(${col}4:${col}11)"
    }
    $ws.Cells.Item($hostTotalRow, 12).Value2 = "Suma de telemetría host"
    $ws.Cells.Item($hostTotalRow, 13).Value2 = "Filas; fuentes pueden solaparse"
    $ws.Cells.Item($hostTotalRow, 14).Value2 = "Total por técnica y fuente; no combinar con conteos de técnicas detectadas."
    $ws.Cells.Item($generalTotalRow, 12).Value2 = "Host + receiver"
    $ws.Cells.Item($generalTotalRow, 13).Value2 = "Niveles distintos"
    $ws.Cells.Item($generalTotalRow, 14).Value2 = "Se conserva solo para trazabilidad; receiver no se representa como fuente host."
    Style-Header $ws.Range("A3:N3")
    $ws.Range("A12:N13").Interior.Color = $ColorLightBlue
    $ws.Range("A12:N13").Font.Bold = $true
    $ws.Range("A1:N13").WrapText = $true
    $ws.Columns("A").ColumnWidth = 27
    $ws.Columns("B:K").ColumnWidth = 12
    $ws.Columns("L:N").ColumnWidth = 42
    $ws.Rows("1:13").AutoFit()
    $ws.Range("A3:N11").AutoFilter()
    Configure-Sheet $ws 3 1 80
    Configure-Print $ws '$A$1:$N$13' 3

    # GRAFICAS: datos fuente reales y cinco gráficos existentes reutilizados.
    $ws = $workbook.Worksheets.Item("GRAFICAS")
    $chartObjects = @{}
    foreach ($co in $ws.ChartObjects()) { $chartObjects[$co.Name] = $co }
    $ws.Cells.Clear()
    Style-Title $ws "A1:U1" "Gráficas de eficacia, cobertura y trazabilidad"
    Style-Subtitle $ws "A2:U2" "Todos los gráficos usan tablas fuente visibles. Volumen de alertas ≠ cobertura; salida externa queda fuera de las capacidades del detector."

    $gMatrix = @()
    $gMatrix += ,@("TEC", "P1", "P2", "P3", "P4", "Total alertas", "Estado")
    for ($i = 1; $i -le 9; $i++) {
        $tec = "TEC-{0:D3}" -f $i
        $rowAlerts = @($alerts | Where-Object TEC -eq $tec)
        $gMatrix += ,@($tec,
            @($rowAlerts | Where-Object profile -eq "P1").Count,
            @($rowAlerts | Where-Object profile -eq "P2").Count,
            @($rowAlerts | Where-Object profile -eq "P3").Count,
            @($rowAlerts | Where-Object profile -eq "P4").Count,
            $rowAlerts.Count, "DETECTADA")
    }
    Write-Block $ws 3 1 $gMatrix | Out-Null
    Style-Header $ws.Range("A3:G3")
    $ws.Range("B4:E12").FormatConditions.Delete()
    $scale = $ws.Range("B4:E12").FormatConditions.AddColorScale(3)
    $scale.ColorScaleCriteria.Item(1).FormatColor.Color = $ColorWhite
    $scale.ColorScaleCriteria.Item(2).FormatColor.Color = $ColorAmber
    $scale.ColorScaleCriteria.Item(3).FormatColor.Color = $ColorBlue

    Write-Block $ws 3 9 @(
        @("Perfil", "Alertas"),
        @("P1", $profileCounts["P1"]),
        @("P2", $profileCounts["P2"]),
        @("P3", $profileCounts["P3"]),
        @("P4", $profileCounts["P4"])
    ) | Out-Null
    Style-Header $ws.Range("I3:J3")

    Write-Block $ws 3 12 @(
        @("Estado", "Campaña TEC", "Campaña FP"),
        @("OK", 9, 10),
        @("WARN", 0, 0),
        @("FAIL", 0, 0),
        @("HITS", 0, 0)
    ) | Out-Null
    Style-Header $ws.Range("L3:N3")
    $ws.Range("L8:N8").Merge()
    $ws.Range("L8").Value2 = "Denominadores separados: TEC=9; FP=10. HITS solo aplica a FP custom."
    $ws.Range("L8:N8").Interior.Color = $ColorAmber
    $ws.Range("L8:N8").WrapText = $true

    Write-Block $ws 3 16 @(
        @("Sistema", "TEC detectadas"),
        @("VR custom", 9),
        @("Públicos CH", 3),
        @("Wazuh base", 3),
        @("Wazuh custom", 4)
    ) | Out-Null
    Style-Header $ws.Range("P3:Q3")

    $hostVisibility = @($visibility | Where-Object fuente -ne "receiver HTTP")
    $visChartRows = New-Object System.Collections.Generic.List[object]
    $visChartRows.Add(@("Fuente", "TEC-001", "TEC-002", "TEC-003", "TEC-004", "TEC-005", "TEC-006", "TEC-007", "TEC-008", "TEC-009", "Total"))
    foreach ($v in $hostVisibility) {
        $values = @()
        foreach ($i in 1..9) { $values += [int]$v.("TEC-{0:D3}" -f $i) }
        $visChartRows.Add(@($v.fuente) + $values + @(($values | Measure-Object -Sum).Sum))
    }
    Write-Block $ws 15 1 $visChartRows.ToArray() | Out-Null
    Style-Header $ws.Range("A15:K15")

    $wazuhRows = @(
        @("TEC", "Wazuh base", "Wazuh custom"),
        @("TEC-001", 0, 1), ,@("TEC-002", 1, 0), ,@("TEC-003", 0, 0),
        @("TEC-004", 1, 0), ,@("TEC-005", 0, 1), ,@("TEC-006", 1, 0),
        @("TEC-007", 0, 1), ,@("TEC-008", 0, 1), ,@("TEC-009", 0, 0)
    )
    Write-Block $ws 15 13 $wazuhRows | Out-Null
    Style-Header $ws.Range("M15:O15")

    $capRows = @(
        @("Capacidad HIDS", "VR custom", "Públicos CH", "Wazuh base", "Wazuh custom"),
        @("Visibilidad", 2, 2, 1, 2),
        @("Detección", 2, 1, 1, 2),
        @("Alerta RT", 2, 1, 1, 2),
        @("Evidencia forense", 0, 2, 1, 1),
        @("Ruido/FP", 2, 1, 1, 1)
    )
    Write-Block $ws 15 17 $capRows | Out-Null
    Style-Header $ws.Range("Q15:U15")
    $ws.Range("R16:U20").FormatConditions.Delete()
    $capScale = $ws.Range("R16:U20").FormatConditions.AddColorScale(3)
    $capScale.ColorScaleCriteria.Item(1).FormatColor.Color = $ColorRed
    $capScale.ColorScaleCriteria.Item(2).FormatColor.Color = $ColorAmber
    $capScale.ColorScaleCriteria.Item(3).FormatColor.Color = $ColorGreen
    $ws.Range("Q21:U21").Merge()
    $ws.Range("Q21").Value2 = "Escala: 0=no acreditado; 1=parcial/contextual; 2=acreditado. La salida externa se evalúa aparte."
    $ws.Range("Q21:U21").Interior.Color = $ColorLightBlue
    $ws.Range("Q21:U21").WrapText = $true

    # Chart 1: visibilidad por técnica, siete fuentes host.
    $co = $chartObjects["Chart 1"]
    $chart = Set-ChartBasics $co "Fuentes de visibilidad host por técnica" 52 "Técnica TEC" "N.º de filas de telemetría"
    Reset-ChartSeries $chart
    for ($r = 16; $r -le 22; $r++) {
        $valueAddress = "B{0}:J{0}" -f $r
        Add-ChartSeries $chart ([string]$ws.Cells.Item($r, 1).Value2) $ws.Range("B15:J15") $ws.Range($valueAddress) | Out-Null
    }
    Position-Chart $co $ws "A27" "J42"

    # Chart 2: alertas por perfil.
    $co = $chartObjects["Chart 2"]
    $chart = Set-ChartBasics $co "Distribución de alertas CLIENT_EVENT por perfil custom" 51 "Perfil" "N.º de alertas"
    Reset-ChartSeries $chart
    $s = Add-ChartSeries $chart "Alertas CLIENT_EVENT" $ws.Range("I4:I7") $ws.Range("J4:J7")
    $s.ApplyDataLabels()
    $chart.HasLegend = $false
    Position-Chart $co $ws "L27" "U42"

    # Chart 3: campañas TEC y FP, denominadores separados.
    $co = $chartObjects["Chart 3"]
    $chart = Set-ChartBasics $co "Pruebas reales custom: TEC 9/9 OK; FP 10/10 OK y 0 hits" 51 "Resultado" "N.º de pruebas"
    Reset-ChartSeries $chart
    Add-ChartSeries $chart "Campaña TEC (n=9)" $ws.Range("L4:L7") $ws.Range("M4:M7") | Out-Null
    Add-ChartSeries $chart "Campaña FP (n=10)" $ws.Range("L4:L7") $ws.Range("N4:N7") | Out-Null
    Position-Chart $co $ws "A44" "J59"

    # Chart 4: Wazuh base/custom por técnica.
    $co = $chartObjects["Chart 4"]
    $chart = Set-ChartBasics $co "Wazuh base y custom: detección específica por técnica" 51 "Técnica TEC" "Detectada (1/0)"
    Reset-ChartSeries $chart
    Add-ChartSeries $chart "Wazuh base" $ws.Range("M16:M24") $ws.Range("N16:N24") | Out-Null
    Add-ChartSeries $chart "Wazuh custom 110xxx" $ws.Range("M16:M24") $ws.Range("O16:O24") | Out-Null
    Position-Chart $co $ws "L44" "U59"

    # Chart 5: alertas custom por técnica.
    $co = $chartObjects["Chart 5"]
    $chart = Set-ChartBasics $co "Alertas custom observadas por técnica" 51 "Técnica TEC" "N.º de alertas CLIENT_EVENT"
    Reset-ChartSeries $chart
    $s = Add-ChartSeries $chart "Alertas" $ws.Range("A4:A12") $ws.Range("F4:F12")
    $s.ApplyDataLabels()
    $chart.HasLegend = $false
    Position-Chart $co $ws "A61" "J76"
    $ws.Range("L61:U64").Merge()
    $ws.Range("L61").Value2 = "Lectura metodológica: el volumen de alertas muestra intensidad de señal, no mayor cobertura. La matriz TEC × perfil y el heatmap de capacidades conservan visibilidad, detección, alerta, evidencia y FP como dimensiones distintas."
    $ws.Range("L61:U64").Interior.Color = $ColorLightBlue
    $ws.Range("L61:U64").WrapText = $true
    $ws.Range("A1:U76").WrapText = $true
    $ws.Columns("A:U").ColumnWidth = 12
    $ws.Columns("A").ColumnWidth = 25
    $ws.Columns("Q").ColumnWidth = 24
    $ws.Rows("1:24").AutoFit()
    Configure-Sheet $ws 2 0 80
    Configure-Print $ws '$A$1:$U$76' 0

    # Dashboard: KPIs custom, navegación y fuentes de los seis gráficos existentes.
    $ws = $workbook.Worksheets.Item("01_Dashboard")
    $ws.Range("A1:N2").UnMerge()
    Style-Title $ws "A1:N1" "Dashboard definitivo · eficacia HIDS/DFIR"
    Style-Subtitle $ws "A2:N2" "Custom P1-P4, artifacts públicos, falsos positivos y Wazuh; visibilidad, detección, alerta, transporte y salida externa permanecen separados."
    try { $ws.Range("A3:N80").UnMerge() } catch { }
    $ws.Range("A3:N80").ClearContents()
    $dashboardRows = @(
        @("KPI", "Valor", "Evidencia / interpretación"),
        @("Artifacts custom validados", 5, "4 detectores CLIENT_EVENT + 1 Router SERVER_EVENT de transporte."),
        @("Técnicas custom detectadas", "9/9", "379 alertas CLIENT_EVENT; P1=19, P2=118, P3=109, P4=133."),
        @("Sources detectoras con alertas", "25/27", "Dos sources declaradas no generaron filas en la ventana final."),
        @("FP custom", "10/10 OK; 0 hits", "Runner FP custom; no incluye Hayabusa Monitoring."),
        @("Artifacts públicos", "5 POSITIVO; 1 NO CONCLUYENTE", "Hayabusa CH 3/9; ETW no concluyente."),
        @("Wazuh base", "3/9", "TEC-002/004/006; ruleset nativo."),
        @("Wazuh custom", "4/9", "TEC-001/005/007/008; reglas TFM 110xxx."),
        @("Visibilidad Sysmon", "ID 1/3/11/12-14/26", "Telemetría y evidencia; ID26 no es alerta individual."),
        @("TEC-009", "HTTP 200 por runner", "TrackNetwork observa conexión sin proceso atribuible; HTTP no se atribuye al artifact."),
        @("Separación de capas", "CLIENT_EVENT ≠ SERVER_EVENT", "Router/Discord/webhook transportan o notifican; no detectan.")
    )
    Write-Block $ws 3 1 $dashboardRows | Out-Null
    Style-Header $ws.Range("A3:C3")
    $ws.Range("A3:C14").Borders.LineStyle = 1
    $ws.Range("E3:H3").Merge()
    $ws.Range("E3").Value2 = "Cobertura táctica MITRE ATT&CK Enterprise"
    $ws.Range("E3:H3").Interior.Color = $ColorNavy
    $ws.Range("E3:H3").Font.Color = $ColorWhite
    $ws.Range("E3:H3").Font.Bold = $true
    Write-Block $ws 4 5 @(
        @("Principal", "5/14", ([double]5 / [double]14), "Execution, Persistence, Discovery, Collection, Impact"),
        @("Ampliada", "6/14", ([double]6 / [double]14), "Principal + Privilege Escalation contextual"),
        @("Contextual", "7/14", ([double]7 / [double]14), "Añade Exfiltration contextual para TEC-009")
    ) | Out-Null
    $ws.Range("G4:G6").NumberFormat = "0.0%"
    $ws.Range("E8:H10").Merge()
    $ws.Range("E8").Value2 = "/14 corresponde a tácticas MITRE ATT&CK Enterprise consideradas en la clasificación; no corresponde a las 9 técnicas TEC ejecutadas en el laboratorio."
    $ws.Range("E8:H10").Interior.Color = $ColorAmber
    $ws.Range("E8:H10").WrapText = $true
    $ws.Range("E12:H14").Merge()
    $ws.Range("E12").Value2 = "Salida externa evaluada separadamente mediante SERVER_EVENT, Router, Discord, webhook o runner; no se considera una capacidad directa del CLIENT_EVENT."
    $ws.Range("E12:H14").Interior.Color = $ColorLightBlue
    $ws.Range("E12:H14").WrapText = $true

    $ws.Range("J3:N3").Merge()
    $ws.Range("J3").Value2 = "Navegación interna"
    $ws.Range("J3:N3").Interior.Color = $ColorNavy
    $ws.Range("J3:N3").Font.Color = $ColorWhite
    $ws.Range("J3:N3").Font.Bold = $true
    $links = @(
        @("07_Catalogo_Artifacts", "Catálogo de artifacts"),
        @("09_Resultados_Custom", "Resultados custom"),
        @("VISIBILIDAD_SISTEMA", "Visibilidad del sistema"),
        @("ALERTAS_VR", "Alertas VR"),
        @("FP_RUNNER", "Runner FP"),
        @("FP_HITS", "Hits FP"),
        @("COMPARACION_VR_WAZUH", "Comparación VR-Wazuh"),
        @("WAZUH_DETALLE", "Detalle Wazuh"),
        @("GRAFICAS", "Gráficas"),
        @("FUENTES", "Fuentes")
    )
    try { $ws.Range("J4:N14").UnMerge() } catch { }
    $ws.Range("J4:N14").ClearContents()
    for ($i = 0; $i -lt $links.Count; $i++) {
        $linkRange = $ws.Range("J" + (4 + $i) + ":N" + (4 + $i))
        $linkRange.Merge()
        $cell = $ws.Cells.Item(4 + $i, 10)
        $ws.Hyperlinks.Add($cell, "", "'" + $links[$i][0] + "'!A1", "", $links[$i][1]) | Out-Null
        $linkRange.Interior.Color = if (($i % 2) -eq 0) { $ColorPaleBlue } else { $ColorWhite }
    }

    # Datos auxiliares de gráficos en S:W.
    $ws.Range("S1:W60").ClearContents()
    Write-Block $ws 2 19 @(
        @("Sistema", "Técnicas detectadas"),
        @("VR custom", 9), ,@("Públicos CH", 3), ,@("Wazuh base", 3), ,@("Wazuh custom", 4)
    ) | Out-Null
    $coverageRows = @(
        @("TEC", "VR custom", "Públicos CH", "Wazuh base", "Wazuh custom"),
        @("TEC-001",1,1,0,1), ,@("TEC-002",1,0,1,0), ,@("TEC-003",1,1,0,0),
        @("TEC-004",1,0,1,0), ,@("TEC-005",1,1,0,1), ,@("TEC-006",1,0,1,0),
        @("TEC-007",1,0,0,1), ,@("TEC-008",1,0,0,1), ,@("TEC-009",1,0,0,0)
    )
    Write-Block $ws 9 19 $coverageRows | Out-Null
    Write-Block $ws 21 19 @(
        @("Capa", "Custom", "Públicos"),
        @("Visibilidad", 9, 9), ,@("Detección", 9, 3), ,@("Alerta RT", 9, 3)
    ) | Out-Null
    Write-Block $ws 28 19 @(
        @("Cobertura /14", "Tácticas"), ,@("Principal",5), ,@("Ampliada",6), ,@("Contextual",7)
    ) | Out-Null
    Write-Block $ws 35 19 @(
        @("Artifact", "Filas FP", "FP CH High/Critical"),
        @("Hayabusa Monitoring CH",73,4), ,@("ServiceCreation",0,0), ,@("ProcessCreation",18,0),
        @("SysmonLogForward",24,0), ,@("ETW Monitoring",0,0), ,@("TrackNetworkConnections",12,0)
    ) | Out-Null
    Write-Block $ws 44 19 @(
        @("Estado", "Campañas"), ,@("POSITIVO",5), ,@("NEGATIVO VÁLIDO",0),
        @("NO CONCLUYENTE",1), ,@("NO EJECUTADO",0), ,@("NO APLICABLE",0)
    ) | Out-Null

    $dashboardCharts = @{}
    foreach ($co in $ws.ChartObjects()) { $dashboardCharts[$co.Name] = $co }

    $co = $dashboardCharts["Chart"]
    $chart = Set-ChartBasics $co "Cobertura táctica MITRE /14" 51 "Clasificación" "N.º de tácticas"
    Reset-ChartSeries $chart
    $s = Add-ChartSeries $chart "Cobertura /14" $ws.Range("S29:S31") $ws.Range("T29:T31")
    $s.ApplyDataLabels()
    $chart.HasLegend = $false
    Position-Chart $co $ws "A17" "G32"

    $co = $dashboardCharts["TFM_Final_Compare_Detection"]
    $chart = Set-ChartBasics $co "Detección específica por sistema y tipo de reglas" 51 "Sistema" "N.º de técnicas TEC detectadas"
    Reset-ChartSeries $chart
    $s = Add-ChartSeries $chart "Técnicas detectadas" $ws.Range("S3:S6") $ws.Range("T3:T6")
    $s.ApplyDataLabels()
    $chart.HasLegend = $false
    Position-Chart $co $ws "H17" "N32"

    $co = $dashboardCharts["TFM_Final_Public_TEC"]
    $chart = Set-ChartBasics $co "Cobertura de detección específica por técnica" 51 "Técnica TEC" "Detectada (1/0)"
    Reset-ChartSeries $chart
    Add-ChartSeries $chart "VR custom" $ws.Range("S10:S18") $ws.Range("T10:T18") | Out-Null
    Add-ChartSeries $chart "Públicos CH" $ws.Range("S10:S18") $ws.Range("U10:U18") | Out-Null
    Add-ChartSeries $chart "Wazuh base" $ws.Range("S10:S18") $ws.Range("V10:V18") | Out-Null
    Add-ChartSeries $chart "Wazuh custom" $ws.Range("S10:S18") $ws.Range("W10:W18") | Out-Null
    Position-Chart $co $ws "A34" "G49"

    $co = $dashboardCharts["TFM_Final_Layers"]
    $chart = Set-ChartBasics $co "Capas acreditadas en CLIENT_EVENT: custom frente a públicos" 51 "Capa" "N.º de técnicas TEC"
    Reset-ChartSeries $chart
    Add-ChartSeries $chart "Custom" $ws.Range("S22:S24") $ws.Range("T22:T24") | Out-Null
    Add-ChartSeries $chart "Públicos" $ws.Range("S22:S24") $ws.Range("U22:U24") | Out-Null
    Position-Chart $co $ws "H34" "N49"

    $co = $dashboardCharts["TFM_Final_FP"]
    $chart = Set-ChartBasics $co "Ventana FP por artifact público (capas separadas)" 51 "Artifact" "N.º de filas/coincidencias"
    Reset-ChartSeries $chart
    Add-ChartSeries $chart "Filas en ventana FP" $ws.Range("S36:S41") $ws.Range("T36:T41") | Out-Null
    Add-ChartSeries $chart "FP CH High/Critical" $ws.Range("S36:S41") $ws.Range("U36:U41") | Out-Null
    Position-Chart $co $ws "A51" "G66"

    $co = $dashboardCharts["TFM_Final_Outcomes"]
    $chart = Set-ChartBasics $co "Estado de las seis campañas públicas" 51 "Estado" "N.º de campañas"
    Reset-ChartSeries $chart
    $s = Add-ChartSeries $chart "Campañas" $ws.Range("S45:S49") $ws.Range("T45:T49")
    $s.ApplyDataLabels()
    $chart.HasLegend = $false
    Position-Chart $co $ws "H51" "N66"

    $ws.Columns("A").ColumnWidth = 29
    $ws.Columns("B").ColumnWidth = 22
    $ws.Columns("C").ColumnWidth = 60
    $ws.Columns("E:H").ColumnWidth = 18
    $ws.Columns("J:N").ColumnWidth = 14
    $ws.Columns("S:W").Hidden = $true
    $ws.Range("A1:N14").WrapText = $true
    $ws.Rows("1:14").AutoFit()
    Configure-Sheet $ws 2 0 80
    Configure-Print $ws '$A$1:$N$66' 0

    # README: alcance definitivo y exclusión explícita del rendimiento.
    $ws = $workbook.Worksheets.Item("README")
    $ws.Range("B4").Value2 = "Consolidar visibilidad, detección CLIENT_EVENT, artifacts custom/públicos, alertabilidad, FP y comparación Wazuh para la memoria del TFM."
    $ws.Range("B6").Value2 = "Maestro actual de ENTREGABLE; evidencia normalizada y YAML validados tienen prioridad."
    $ws.Range("A16").Value2 = "Artifacts custom validados"
    $ws.Range("B16").Value2 = "P1-P4 CLIENT_EVENT + Router SERVER_EVENT; 27 sources detectoras declaradas, 25 con alertas en campaña final."
    $ws.Range("A17").Value2 = "Alcance excluido"
    $ws.Range("B17").Value2 = "CPU, memoria, E/S, consumo por proceso y escenarios de rendimiento pertenecen a otro libro y no se analizan aquí."
    $ws.Range("A18").Value2 = "Excel antiguos"
    $ws.Range("B18").Value2 = "Los dos ficheros históricos comparados son idénticos; no se recuperan métricas obsoletas ni estados Pendiente."
    $ws.Range("A1:B18").WrapText = $true
    $ws.Columns("A").ColumnWidth = 34
    $ws.Columns("B").ColumnWidth = 105
    $ws.Rows("1:18").AutoFit()
    Configure-Sheet $ws 3 0 90

    # RESUMEN_EJECUTIVO: añade trazabilidad custom sin alterar Wazuh/públicos ya cerrados.
    $ws = $workbook.Worksheets.Item("RESUMEN_EJECUTIVO")
    $extraRows = @(
        @("Artifacts custom detectores", 4, "P1-P4 CLIENT_EVENT; nombres exactos y hashes en catálogo."),
        @("Artifact custom de transporte", 1, "Router SERVER_EVENT validado; no suma detección."),
        @("Sources detectoras declaradas", 27, "P1=4; P2=5; P3=13; P4=5."),
        @("Sources con alertas finales", 25, "Dos sources declaradas no generaron filas en la ventana final."),
        @("P0 Debug", "NO LOCALIZADO", "No existe YAML P0 en validated; no se inventa uso ni resultado.")
    )
    Write-Block $ws 37 1 $extraRows | Out-Null
    Style-Header $ws.Range("A3:C3")
    $ws.Range("A37:C41").Interior.Color = $ColorLightBlue
    $ws.Range("A1:C41").WrapText = $true
    $ws.Columns("A").ColumnWidth = 43
    $ws.Columns("B").ColumnWidth = 24
    $ws.Columns("C").ColumnWidth = 97
    $ws.Rows("1:41").AutoFit()

    # COMPARACION_VR_WAZUH: explicita custom principal/source y volumen/cobertura.
    $ws = $workbook.Worksheets.Item("COMPARACION_VR_WAZUH")
    $ws.Range("M4").Value2 = "Cuatro artifacts principales P1-P4, 25 sources con alertas y 9/9 técnicas. El volumen 379 no es la unidad de cobertura."
    $ws.Range("A19:M19").Merge()
    $ws.Range("A19").Value2 = "Aportación custom Velociraptor: P4 aporta señales básicas; P3 cubre 9/9 con semántica; P2 añade contexto forense; P1 prioriza señales críticas. Router SERVER_EVENT queda fuera de la detección."
    $ws.Range("A19:M19").Interior.Color = $ColorLightBlue
    $ws.Range("A19:M19").WrapText = $true
    $ws.Range("A1:M19").WrapText = $true
    $ws.Rows("1:19").AutoFit()

    # FUENTES: conserva el inventario actual, normaliza a diez columnas y añade artifacts/normalizados/contraste.
    $ws = $workbook.Worksheets.Item("FUENTES")
    $lastRow = $ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1
    $existing = New-Object System.Collections.Generic.List[object]
    for ($r = 4; $r -le $lastRow; $r++) {
        $pathValue = [string]$ws.Cells.Item($r, 1).Value2
        if ([string]::IsNullOrWhiteSpace($pathValue)) { continue }
        if ($pathValue -eq "ruta") { continue }
        $existing.Add(@(
            [string]$ws.Cells.Item($r, 2).Value2,
            $pathValue,
            [string]$ws.Cells.Item($r, 4).Value2,
            [string]$ws.Cells.Item($r, 3).Value2,
            [string]$ws.Cells.Item($r, 5).Value2,
            [string]$ws.Cells.Item($r, 6).Value2,
            [string]$ws.Cells.Item($r, 7).Value2,
            "USADA"
        ))
    }
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $existing) { [void]$seen.Add([string]$row[1]) }

    $additionalPaths = @(
        "01_ARTIFACTS\validated\Custom.TFM.HIDS.P1.Critical.Priority.Event_v1.yaml",
        "01_ARTIFACTS\validated\Custom.TFM.HIDS.P2.High.Forensic.Event_v1.yaml",
        "01_ARTIFACTS\validated\Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1.yaml",
        "01_ARTIFACTS\validated\Custom.TFM.HIDS.P4.Low.Basic_v2.yaml",
        "01_ARTIFACTS\validated\Custom.TFM.HIDS.Router.JSONL.Discord.SOC_v3.yaml",
        "01_ARTIFACTS\validated\SHA256SUMS_P1_P2_RECALIBRATION.txt",
        "04_EVIDENCE\Excel_visibilidad_26062026\normalized\alertas_vr.csv",
        "04_EVIDENCE\Excel_visibilidad_26062026\normalized\visibilidad_sistema.csv",
        "04_EVIDENCE\Excel_visibilidad_26062026\normalized\tecnicas_real_vr.csv",
        "04_EVIDENCE\Excel_visibilidad_26062026\normalized\fp_runner.csv",
        "04_EVIDENCE\Excel_visibilidad_26062026\normalized\fp_hits.csv",
        "04_EVIDENCE\Excel_visibilidad_26062026\normalized\dataset_summary.json",
        "04_EVIDENCE\Excel_visibilidad_26062026\TFM_VISIBILIDAD_DETECCION_FP_WAZUH_FINAL_26062026.xlsx",
        "04_EVIDENCE\Excel_visibilidad_26062026\Analisis_Tecnicas_TFM_V.5.xlsx"
    )
    foreach ($relative in $additionalPaths) {
        if ($seen.Contains($relative)) { continue }
        $absolute = Join-Path $Root $relative
        if (-not (Test-Path -LiteralPath $absolute)) { continue }
        $item = Get-Item -LiteralPath $absolute
        $hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $absolute).Hash
        $use = if ($relative -like "01_ARTIFACTS\validated\*") { "ARTIFACT_VALIDADO" } elseif ($relative -like "*normalized*") { "EVIDENCIA_NORMALIZADA" } else { "CONTRASTE_HISTORICO" }
        $sheetMetric = if ($use -eq "ARTIFACT_VALIDADO") { "07_Catalogo_Artifacts; 14_Arquitectura_Custom" } elseif ($use -eq "EVIDENCIA_NORMALIZADA") { "09_Resultados_Custom; VISIBILIDAD_SISTEMA; GRAFICAS" } else { "Auditoría comparativa; no trasladado" }
        $status = if ($use -eq "CONTRASTE_HISTORICO") { "SUPERADO / SOLO CONTRASTE" } else { "USADA" }
        $existing.Add(@($item.Extension.TrimStart('.').ToLowerInvariant(), $relative, $item.LastWriteTimeUtc.ToString("o"), $item.Length, $hash, $use, $sheetMetric, $status))
        [void]$seen.Add($relative)
    }

    Delete-AllTables $ws
    try { if ($ws.AutoFilterMode) { $ws.AutoFilterMode = $false } } catch { }
    $ws.Cells.Clear()
    Style-Title $ws "A1:J1" "Fuentes analizadas y trazabilidad SHA-256"
    Style-Subtitle $ws "A2:J2" "Prioridad: evidencia normalizada, campañas reales, artifacts validados y maestro actual. Los Excel históricos se usan solo como contraste."
    $sourceInventory = New-Object System.Collections.Generic.List[object]
    $sourceInventory.Add(@("ID", "Tipo", "Fichero", "Ruta relativa", "Fecha", "Tamaño", "SHA-256", "Uso", "Hoja o métrica derivada", "Estado"))
    $idx = 0
    foreach ($row in $existing) {
        $idx++
        $pathValue = [string]$row[1]
        $fileName = [System.IO.Path]::GetFileName($pathValue)
        $sourceInventory.Add(@(("F-{0:D4}" -f $idx), $row[0], $fileName, $pathValue, $row[2], $row[3], $row[4], $row[5], $row[6], $row[7]))
    }
    Write-Block $ws 3 1 $sourceInventory.ToArray() | Out-Null
    $sourcesEnd = 3 + $sourceInventory.Count - 1
    Add-Table $ws ("A3:J{0}" -f $sourcesEnd) "FuentesTrazabilidadTable" "TableStyleMedium2" | Out-Null
    $ws.Range("A1:J$sourcesEnd").WrapText = $true
    $ws.Columns("A:B").ColumnWidth = 13
    $ws.Columns("C:D").ColumnWidth = 58
    $ws.Columns("E:F").ColumnWidth = 24
    $ws.Columns("G").ColumnWidth = 68
    $ws.Columns("H:J").ColumnWidth = 34
    $ws.Rows("1:$sourcesEnd").AutoFit()
    Configure-Sheet $ws 3 2 75
    Configure-Print $ws ("`$A`$1:`$J`$$sourcesEnd") 3

    # DISCREPANCIAS: añade cierres trazables, conservando registros existentes.
    $ws = $workbook.Worksheets.Item("DISCREPANCIAS")
    $last = $ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1
    $append = @(
        @("SOURCE_DECLARADA_SIN_FILAS", "INFO", "01_ARTIFACTS/validated/Custom.TFM.HIDS.P2.High.Forensic.Event_v1.yaml", "Forensic_EVENT_Security_ScheduledTask_Optional no generó filas en ALERTAS_VR dentro de la campaña final.", "No se atribuyen alertas; el artifact principal sigue validado."),
        @("SOURCE_DECLARADA_SIN_FILAS", "INFO", "01_ARTIFACTS/validated/Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1.yaml", "TEC009_ZIP_FileCreate_Sysmon_Context no generó filas en ALERTAS_VR dentro de la campaña final.", "TEC-009 se mantiene detectada por otras sources P1/P3; no se inventa ejecución de esta source."),
        @("P0_NO_LOCALIZADO", "INFO", "01_ARTIFACTS/validated", "No existe artifact P0 en la carpeta validated.", "P0 se documenta como no incorporado y no utilizado."),
        @("EXCEL_HISTORICOS_IDENTICOS", "INFO", "04_EVIDENCE/Excel_visibilidad_26062026", "TFM_VISIBILIDAD...FINAL_26062026.xlsx y Analisis_Tecnicas_TFM_V.5.xlsx son binariamente idénticos (SHA F9CD06E7...).", "Solo contraste; no se trasladan estados Pendiente ni contenido superado."),
        @("HOJA_AUXILIAR_ANTIGUA", "INFO", "99_Listas", "Hoja auxiliar vacía presente solo en los Excel históricos.", "No se recupera: no aporta métrica, fuente ni nota metodológica nueva.")
    )
    Write-Block $ws ($last + 1) 1 $append | Out-Null
    $ws.Range("A1:E" + ($last + $append.Count)).WrapText = $true

    # Hipervínculo de la guía a la hoja renombrada.
    $wsGuide = $workbook.Worksheets.Item("00_Guia")
    try { $wsGuide.Range("A12").Hyperlinks.Delete() } catch { }
    try { $wsGuide.Hyperlinks.Add($wsGuide.Range("A13"), "", "'09_Resultados_Custom'!A1", "", "09_Resultados_Custom") | Out-Null } catch { }

    # Validaciones internas antes de guardar.
    if ($workbook.Worksheets.Count -ne 30) { throw "Número de hojas inesperado: $($workbook.Worksheets.Count)" }
    try { $null = $workbook.Worksheets.Item("09_Resultados_Custom") } catch { throw "No existe 09_Resultados_Custom" }
    foreach ($sheet in $workbook.Worksheets) { if ($sheet.Name -eq "09_Benchmark_Plan") { throw "Sigue existiendo 09_Benchmark_Plan" } }
    $chartCount = 0
    foreach ($sheet in $workbook.Worksheets) { $chartCount += $sheet.ChartObjects().Count }
    if ($chartCount -ne 11) { throw "Número de gráficos inesperado: $chartCount" }
    if ($profileCounts["P1"] + $profileCounts["P2"] + $profileCounts["P3"] + $profileCounts["P4"] -ne 379) { throw "Suma de perfiles distinta de 379" }

    $excel.Calculation = -4105
    $excel.CalculateFullRebuild()
    $workbook.Save()
    $workbook.Close($true)
    $workbook = $null
    $excel.Quit()
    $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    $item = Get-Item -LiteralPath $Path
    $hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $Path).Hash
    [pscustomobject]@{
        Status = "OK"
        Path = $Path
        Bytes = $item.Length
        SHA256 = $hash
        Sheets = 30
        Charts = 11
        Alerts = $initialAlerts
        Profiles = $profileCounts
        SourcesGrouped = $groups.Count
        SourcesInventory = $existing.Count
    } | ConvertTo-Json -Depth 5
}
catch {
    $line = $_.InvocationInfo.ScriptLineNumber
    $text = $_.InvocationInfo.Line
    Write-Error ("FALLO COM en línea {0}: {1}`n{2}`nSTACK:`n{3}" -f $line, $_.Exception.Message, $text, $_.ScriptStackTrace)
    throw
}
finally {
    if ($workbook -ne $null) { try { $workbook.Close($false) } catch { } }
    if ($excel -ne $null) { try { $excel.Quit() } catch { } }
    if ($workbook -ne $null) { try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) } catch { } }
    if ($excel -ne $null) { try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) } catch { } }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
