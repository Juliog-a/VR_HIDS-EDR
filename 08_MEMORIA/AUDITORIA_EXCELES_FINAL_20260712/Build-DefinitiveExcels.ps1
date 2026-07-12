[CmdletBinding()]
param(
    [string]$VisibilityPath = "C:\Users\julio\Desktop\TFM\08_MEMORIA\ENTREGABLE\TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712.xlsx",
    [string]$BenchmarkPath = "C:\Users\julio\Desktop\TFM\08_MEMORIA\ENTREGABLE\TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_DEFINITIVO_20260712.xlsx",
    [string]$EvidenceRoot = "C:\Users\julio\Desktop\TFM\04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS",
    [string]$AuditDir = "C:\Users\julio\Desktop\TFM\08_MEMORIA\AUDITORIA_EXCELES_FINAL_20260712"
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

# Excel constants.
$xlCalculationAutomatic = -4105
$xlColumnClustered = 51
$xlCategory = 1
$xlValue = 2
$xlPrimary = 1
$xlLegendPositionBottom = -4107
$xlPasteFormats = -4122
$xlVAlignTop = -4160
$xlHAlignLeft = -4131

function Set-RowValues {
    param(
        [Parameter(Mandatory)]$Worksheet,
        [Parameter(Mandatory)][int]$Row,
        [Parameter(Mandatory)][object[]]$Values,
        [int]$StartColumn = 1
    )
    for ($i = 0; $i -lt $Values.Count; $i++) {
        $cell = $Worksheet.Cells.Item($Row, $StartColumn + $i)
        $value = $Values[$i]
        if ($null -eq $value) {
            $cell.ClearContents() | Out-Null
        } elseif ($value -is [byte] -or $value -is [int16] -or $value -is [int32] -or $value -is [int64] -or $value -is [single] -or $value -is [double] -or $value -is [decimal]) {
            $cell.Value2 = [double]$value
        } elseif ($value -is [bool]) {
            $cell.Value2 = [bool]$value
        } else {
            $cell.Value2 = [string]$value
        }
    }
}

function Copy-RowFormat {
    param($Excel, $Worksheet, [int]$SourceRow, [int]$FirstTargetRow, [int]$LastTargetRow, [string]$LastColumn)
    $Worksheet.Range("A$SourceRow`:$LastColumn$SourceRow").Copy() | Out-Null
    $Worksheet.Range("A$FirstTargetRow`:$LastColumn$LastTargetRow").PasteSpecial($xlPasteFormats) | Out-Null
    try { $Excel.CutCopyMode = $false } catch {}
    $Worksheet.Rows("$FirstTargetRow`:$LastTargetRow").RowHeight = $Worksheet.Rows($SourceRow).RowHeight
}

function Set-ChartPosition {
    param($Worksheet, $ChartObject, [string]$TopLeft, [string]$BottomRight)
    $from = $Worksheet.Range($TopLeft)
    $to = $Worksheet.Range($BottomRight)
    $ChartObject.Left = $from.Left
    $ChartObject.Top = $from.Top
    $ChartObject.Width = ($to.Left + $to.Width) - $from.Left
    $ChartObject.Height = ($to.Top + $to.Height) - $from.Top
}

function Configure-Chart {
    param(
        $Chart,
        [string]$Title,
        [string]$CategoryTitle,
        [string]$ValueTitle,
        [string]$ValueNumberFormat = "0",
        [switch]$DataLabels
    )
    $Chart.HasTitle = $true
    $Chart.ChartTitle.Text = $Title
    $Chart.HasLegend = $true
    $Chart.Legend.Position = $xlLegendPositionBottom
    try { $Chart.ChartStyle = 10 } catch {}
    try {
        $x = $Chart.Axes($xlCategory, $xlPrimary)
        $x.HasTitle = $true
        $x.AxisTitle.Text = $CategoryTitle
    } catch {}
    try {
        $y = $Chart.Axes($xlValue, $xlPrimary)
        $y.HasTitle = $true
        $y.AxisTitle.Text = $ValueTitle
        $y.TickLabels.NumberFormat = $ValueNumberFormat
        try { $y.TickLabels.NumberFormatLocal = ($ValueNumberFormat -replace '\.', ',') } catch {}
        $y.MinimumScale = 0
    } catch {}
    if ($DataLabels) {
        try { $Chart.ApplyDataLabels() } catch {}
    }
}

function Add-DashboardChart {
    param(
        $Dashboard,
        $SourceSheet,
        [string]$Name,
        [string]$SourceRange,
        [string]$Title,
        [string]$CategoryTitle,
        [string]$ValueTitle,
        [string]$TopLeft,
        [string]$BottomRight
    )
    $co = $Dashboard.ChartObjects().Add(0, 0, 420, 245)
    $co.Name = $Name
    $chart = $co.Chart
    $chart.ChartType = $xlColumnClustered
    $chart.SetSourceData($SourceSheet.Range($SourceRange))
    Configure-Chart -Chart $chart -Title $Title -CategoryTitle $CategoryTitle -ValueTitle $ValueTitle -DataLabels
    Set-ChartPosition -Worksheet $Dashboard -ChartObject $co -TopLeft $TopLeft -BottomRight $BottomRight
    return $co
}

function Get-EvidenceRole {
    param([System.IO.FileInfo]$File)
    if ($File.Name -match "_REAL\.(csv|json)$") { return "PUBLIC_EXPORT_TEC" }
    if ($File.Name -match "_FP\.(csv|json)$") { return "PUBLIC_EXPORT_FP" }
    if ($File.Name -match "TEC_(start|end)_utc") { return "WINDOW_TEC" }
    if ($File.Name -match "FP_(start|end)_utc") { return "WINDOW_FP" }
    if ($File.Name -match "TFM_TEC.*summary") { return "RUNNER_TEC_SUMMARY" }
    if ($File.Name -match "TFM_FP.*summary") { return "RUNNER_FP_SUMMARY" }
    if ($File.Name -match "vr_hits") { return "FP_HITS_SUPPORT" }
    if ($File.Name -match "SHA256") { return "SOURCE_HASH_MANIFEST" }
    return "PUBLIC_SUPPORT"
}

function Update-VisibilityWorkbook {
    param($Excel, $Workbook, [string]$EvidenceRootPath)

    $dashboard = $Workbook.Worksheets.Item("01_Dashboard")
    $dashboard.Range("A1").Value2 = "Dashboard definitivo · Velociraptor HIDS/DFIR"
    $dashboard.Range("A2").Value2 = "Cierre 12/07/2026: custom P1-P4, artifacts públicos, falsos positivos y Wazuh; capas metodológicas separadas."
    Set-RowValues $dashboard 4 @("Campañas públicas ejecutadas", 6, "Seis artifacts públicos con ventanas TEC/FP trazables.")
    Set-RowValues $dashboard 5 @("Campañas positivas", 5, "ProcessCreation, ServiceCreation, SysmonLogForward, Hayabusa.Monitoring CH y TrackNetworkConnections.")
    Set-RowValues $dashboard 6 @("Negativos válidos", 0, "Ninguna campaña cumple todos los requisitos para este estado.")
    Set-RowValues $dashboard 7 @("No concluyentes", 1, "ETW: faltan runner/log/summary TEC y manifiesto de parámetros.")
    Set-RowValues $dashboard 8 @("Eventos públicos TEC canónicos", 516, "Deduplicados y filtrados por ventana; 521 filas raw en ventana. Capas heterogéneas, no sumar como detecciones.")
    Set-RowValues $dashboard 9 @("Filas públicas FP en ventana", 127, "Telemetría/coincidencias observadas; no equivalen todas a falsos positivos.")
    $dashboard.Range("B4:B9").NumberFormat = "0"
    $dashboard.Range("B10").NumberFormat = "@"
    Set-RowValues $dashboard 10 @("FP Hayabusa CH", "4 matches / 1 caso", "Cuatro high, cero critical; tres eventos subyacentes, todos ligados a FP-003 benigno.")
    $dashboard.Range("B11:B14").NumberFormat = "@"
    Set-RowValues $dashboard 11 @("Visibilidad pública por TEC", "9/9", "ProcessCreation y fuentes forenses aportan telemetría; no implica detección.")
    Set-RowValues $dashboard 12 @("Detección pública CH específica", "3/9", "TEC-001, TEC-003 y TEC-005 con coincidencias high; no existe campaña CHM separada.")
    Set-RowValues $dashboard 13 @("Detección custom CLIENT_EVENT", "9/9", "P1/P2/P3/P4; 379 filas consolidadas.")
    Set-RowValues $dashboard 14 @("Detección Wazuh específica", "4/9", "TEC-001, TEC-005, TEC-007 y TEC-008 con reglas 110xxx > 0; TEC-009 mantiene gap.")
    Set-RowValues $dashboard 15 @("TEC-009", "HTTP 200 confirmado", "Runner: UploadSucceeded=True, ZIP 6245 B y SHA-256; TrackNetwork observa destino sin PID/proceso atribuible.")
    $dashboard.Range("A4:C15").WrapText = $true
    $dashboard.Range("A4:C15").VerticalAlignment = $xlVAlignTop

    $guide = $Workbook.Worksheets.Item("00_Guia")
    $guide.Range("C8").Value2 = "Campaña final pública cerrada el 12/07/2026; detalle y fuentes integrados."
    $guide.Range("C12").Value2 = "Resultados FP custom y públicos; separar telemetría de detección injustificada."

    $readme = $Workbook.Worksheets.Item("README")
    $readme.Range("A2").Value2 = "Libro definitivo consolidado con campañas públicas de 12/07/2026 y preservación íntegra de custom, FP y Wazuh."
    $readme.Range("A6").Value2 = "Fecha de revisión final"
    $readme.Range("B6").Value2 = "2026-07-12 (Europe/Madrid)"
    $readme.Range("A13").Value2 = "Limitaciones"
    $readme.Range("B13").Value2 = "ETW no concluyente por falta de evidencia TEC de configuración/runner; TrackNetwork no permite atribución de proceso; JSONL/Discord no son detección."
    Set-RowValues $readme 14 @("Artifacts públicos", "04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS; 6/6 campañas, 112 ficheros, ventanas TEC/FP.")
    Set-RowValues $readme 15 @("Criterio público", "Se filtra por ventana; evento crudo=visibilidad; coincidencia Sigma=detec.; salida externa se acredita aparte.")
    $readme.Range("A14:B15").WrapText = $true
    try { $readme.AutoFilterMode = $false; $readme.Range("A4:B15").AutoFilter() | Out-Null } catch {}

    # Matriz definitiva por artifact público, ampliada sin crear una hoja redundante.
    $public = $Workbook.Worksheets.Item("15_Publicos_Definitivo")
    $newHeaders = @("Configuración / criterio", "Ventana TEC UTC", "Ventana FP UTC", "Filas TEC", "Filas FP", "FP CH high/critical", "Casos FP", "Fuentes exactas", "Resultado campaña", "Limitación / anomalía")
    Set-RowValues $public 4 $newHeaders 10
    $publicRows = @(
        @("Windows.Hayabusa.Monitoring", "CLIENT_EVENT", "Sí", "Sí", "Sí (Sigma CH)", "Sí", "No confirmada", "TEC-001/003/005 en CH high", "Positivo con FP; no extrapolar niveles inferiores", "Control planificado Critical/High/Stable; configuración efectiva no preservada", "2026-07-12T13:25:03.318Z–13:31:18.641Z", "2026-07-12T13:41:00.378Z–13:41:29.197Z", 185, 73, 4, 1, "04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\04_HayabusaMonitoring_CH\Windows.Hayabusa.Monitoring_REAL.json | ..._FP.json | ventanas", "POSITIVO", "185 matches canónicos/187 filas raw TEC: high=7, critical=0. FP CH: 4 high/3 eventos/FP-003. No existe campaña CHM."),
        @("Windows.Events.ServiceCreation", "CLIENT_EVENT", "Sí", "Sí: 7045", "No: 7045 aporta visibilidad, no atribuye malicia", "No", "No", "TEC-005", "Creación de servicio observada", "Event ID 7045", "2026-07-12T12:26:12.998Z–12:33:05.540Z", "2026-07-12T12:37:57.100Z–12:38:39.571Z", 1, 0, 0, 0, "04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\02_ServiceCreation\Windows.Events.ServiceCreation_REAL.json | ..._FP.json | summary TEC", "POSITIVO", "StartExit=1053 esperado para wrapper cmd.exe; no invalida la creación observada ni el 7045."),
        @("Windows.Events.ProcessCreation", "CLIENT_EVENT", "Sí", "Sí", "No por sí solo", "No por sí solo", "No", "TEC-001 a TEC-009", "Telemetría de procesos", "Monitoring; CommandLine/parent/PID", "2026-07-12T11:50:43.356Z–11:58:00.217Z", "2026-07-12T12:01:47.092Z–12:02:15.729Z", 89, 18, 0, 0, "04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\01_ProcessCreation\Windows.Events.ProcessCreation_REAL.json | ..._FP.json | ventanas", "POSITIVO", "Proceso observado no equivale a técnica detectada; no usar rutas/nombres TEC como IOC."),
        @("Windows.Sysinternals.SysmonLogForward", "CLIENT_EVENT", "Sí", "Sí", "No por sí solo", "No por sí solo", "No", "Contexto TEC-001 a TEC-009", "Reenvío/forense", "Sysmon IDs observados 1,3,11,12,13", "2026-07-12T12:46:16.670Z–12:52:34.859Z", "2026-07-12T12:55:41.452Z–12:56:21.064Z", 113, 24, 0, 0, "04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\03_SysmonLogForward\Windows.Sysinternals.SysmonLogForward_REAL.json | ..._FP.json | ventanas", "POSITIVO", "Sin ID 14/26. Dos ID3 al destino se atribuyen a powershell.exe en esta campaña; siguen siendo visibilidad/forense. ID 26 nunca sería alerta individual."),
        @("Windows.ETW.Monitoring", "CLIENT_EVENT", "Sí", "Sin coincidencias", "No acreditada", "No", "No", "No atribuible", "0 filas, evidencia incompleta", "Parámetros TEC no preservados", "2026-07-12T14:22:25.782Z–14:30:04.869Z", "2026-07-12T14:32:20.307Z–14:32:57.519Z", 0, 0, 0, 0, "04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\05_ETWMonitoring\Windows.ETW.Monitoring_REAL.json | ..._FP.json | ventanas", "NO CONCLUYENTE", "Faltan runner/log/summary TEC y manifiesto de configuración; no escribir ‘ETW no funciona’ ni ‘negativo válido’."),
        @("Generic.Events.TrackNetworkConnections", "CLIENT_EVENT", "Sí", "Sí: destino observado", "No por sí solo", "No por sí solo", "No por artifact; sí runner", "TEC-009 contextual", "Conexión visible; upload acreditado aparte", "Diff added/removed; filtro por _ts", "2026-07-12T14:43:12.205Z–14:50:12.514Z", "2026-07-12T14:52:39.396Z–14:53:04.206Z", 128, 12, 0, 0, "04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\06_TrackNetworkConnections\Generic.Events.TrackNetworkConnections_REAL.json | summary TEC", "POSITIVO", "128 eventos canónicos/131 filas raw. 3 filas objetivo en ventana; Timestamp=1601, PID=0, ProcInfo vacío. No atribuir a PowerShell. Runner: HTTP 200/UploadSucceeded=True/6245 B." )
    )
    for ($i = 0; $i -lt $publicRows.Count; $i++) { Set-RowValues $public (5 + $i) $publicRows[$i] }
    $public.Range("J4:S14").WrapText = $true
    $public.Range("A4:S14").VerticalAlignment = $xlVAlignTop
    Set-RowValues $public 14 @("TOTAL 6 CAMPAÑAS", "", "", "", "", "", "", "", "", "", "", "")
    $public.Range("M14").Formula = "=SUM(M5:M10)"
    $public.Range("N14").Formula = "=SUM(N5:N10)"
    $public.Range("O14").Formula = "=SUM(O5:O10)"
    $public.Range("P14").Formula = "=SUM(P5:P10)"
    $public.Range("Q14").Value2 = "Totales de filas por campaña; no equivalen a detecciones entre artifacts heterogéneos."
    $public.Range("R14").Formula = '=COUNTIF(R5:R10,"POSITIVO")&" positivos; "&COUNTIF(R5:R10,"NEGATIVO VÁLIDO")&" negativos válidos; "&COUNTIF(R5:R10,"NO CONCLUYENTE")&" no concluyentes"'
    $public.Range("S14").Value2 = "Salida externa TEC-009 acreditada por runner, no por TrackNetwork."
    $public.Range("M5:P14").NumberFormat = "0"
    foreach ($c in "J", "K", "L", "Q", "S") { $public.Columns($c).ColumnWidth = if ($c -in @("Q", "S")) { 45 } else { 26 } }
    $public.Columns("M:P").ColumnWidth = 13
    $public.Columns("R").ColumnWidth = 19
    try { $public.AutoFilterMode = $false; $public.Range("A4:S14").AutoFilter() | Out-Null } catch {}
    $public.PageSetup.PrintArea = "`$A`$1:`$S`$14"

    # Cerrar filas pendientes en el control operativo y preservar las planificadas como histórico/N/A.
    $control = $Workbook.Worksheets.Item("04_Control_Publicos")
    Set-RowValues $control 4 @("Visibilidad", "Detección_Regla", "Alerta_RT", "Artifact_Alertante", "Salida_Externa", "Veredicto_Alertabilidad", "Resultado_Campaña") 25
    for ($r = 5; $r -le 13; $r++) {
        $control.Cells.Item($r, 14).Value2 = "Visible; 89 filas TEC / 18 FP en ventana"
        $control.Cells.Item($r, 21).Value2 = "Cerrado 12/07/2026"
        $control.Cells.Item($r, 22).Value2 = "2026-07-12"
        $control.Cells.Item($r, 23).Value2 = "04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\01_ProcessCreation\Windows.Events.ProcessCreation_REAL.json"
        $control.Cells.Item($r, 24).Value2 = "Telemetría; no detección por nombre de proceso."
        Set-RowValues $control $r @("Sí", "No", "No", "N/A", "No", "VISIBILIDAD", "POSITIVO") 25
    }
    Set-RowValues $control 14 @("1 evento 7045 TEC; 0 FP", "Event ID 7045", "N/A", "T1569.002", "ServiceName/ImagePath; StartExit=1053 esperado", "Bajo", "No", "Cerrado 12/07/2026", "2026-07-12", "04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\02_ServiceCreation\Windows.Events.ServiceCreation_REAL.json", "Creación observada; visibilidad, no detección de actividad maliciosa.", "Sí", "No", "No", "N/A", "No", "VISIBILIDAD", "POSITIVO") 14
    for ($r = 15; $r -le 19; $r++) {
        $control.Cells.Item($r, 14).Value2 = "113 filas TEC / 24 FP; IDs 1,3,11,12,13"
        $control.Cells.Item($r, 21).Value2 = "Cerrado 12/07/2026"
        $control.Cells.Item($r, 22).Value2 = "2026-07-12"
        $control.Cells.Item($r, 23).Value2 = "04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\03_SysmonLogForward\Windows.Sysinternals.SysmonLogForward_REAL.json"
        $control.Cells.Item($r, 24).Value2 = "Reenvío/forense; sin lógica adicional de detección."
        Set-RowValues $control $r @("Sí", "No", "No", "N/A", "No", "VISIBILIDAD / FORENSE", "POSITIVO") 25
    }
    for ($r = 20; $r -le 28; $r++) {
        $control.Cells.Item($r, 14).Value2 = "Histórico conservado; no forma parte de la campaña RT 12/07"
        $control.Cells.Item($r, 21).Value2 = "HISTÓRICO"
        Set-RowValues $control $r @("Sí histórico", "Parcial", "No", "N/A", "No", "HUNT HISTÓRICO", "N/A") 25
    }
    for ($r = 29; $r -le 37; $r++) {
        $control.Cells.Item($r, 14).Value2 = "NO EJECUTADA; no se añaden resultados Medium"
        $control.Cells.Item($r, 21).Value2 = "NO EJECUTADA"
        Set-RowValues $control $r @("N/A", "N/A", "N/A", "N/A", "N/A", "SIN CAMPAÑA CHM", "N/A") 25
    }
    $chHigh = @{ "TEC-001" = "2 high; detectado CH"; "TEC-003" = "1 high; detectado CH"; "TEC-005" = "4 high en ventana; 3 atribuibles semánticamente al servicio" }
    for ($r = 38; $r -le 46; $r++) {
        $tec = [string]$control.Cells.Item($r, 8).Value2
        $result = if ($chHigh.ContainsKey($tec)) { $chHigh[$tec] } else { "0 high/critical específicos; niveles inferiores observados no forman campaña separada" }
        $det = if ($chHigh.ContainsKey($tec)) { "Sí CH" } else { "No CH" }
        $control.Cells.Item($r, 14).Value2 = $result
        $control.Cells.Item($r, 16).Value2 = if ($chHigh.ContainsKey($tec)) { "high" } else { "N/A CH" }
        $control.Cells.Item($r, 21).Value2 = "Cerrado 12/07/2026"
        $control.Cells.Item($r, 22).Value2 = "2026-07-12"
        $control.Cells.Item($r, 23).Value2 = "04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\04_HayabusaMonitoring_CH\Windows.Hayabusa.Monitoring_REAL.json"
        $control.Cells.Item($r, 24).Value2 = "TEC: 185 matches canónicos/187 filas raw; high=7, critical=0. FP: 4 high, 1 caso."
        Set-RowValues $control $r @("Sí", $det, $det, "Windows.Hayabusa.Monitoring", "No confirmada", "DETECCIÓN SIGMA CH", "POSITIVO") 25
    }
    for ($r = 47; $r -le 55; $r++) {
        $control.Cells.Item($r, 14).Value2 = "NO EJECUTADA; no existe campaña CHM separada"
        $control.Cells.Item($r, 21).Value2 = "NO EJECUTADA"
        Set-RowValues $control $r @("N/A", "N/A", "N/A", "N/A", "N/A", "SIN CAMPAÑA CHM", "N/A") 25
    }
    for ($r = 56; $r -le 61; $r++) {
        $control.Cells.Item($r, 14).Value2 = "No ejecutado en campaña final 12/07; conservado como plan/histórico"
        $control.Cells.Item($r, 21).Value2 = "NO APLICABLE"
        Set-RowValues $control $r @("N/A", "N/A", "N/A", "N/A", "N/A", "FUERA DE CAMPAÑA", "N/A") 25
    }
    Copy-RowFormat -Excel $Excel -Worksheet $control -SourceRow 61 -FirstTargetRow 62 -LastTargetRow 67 -LastColumn "AE"
    $summaryRows = @(
        @("PUB-FINAL-01", "Cierre", "P4", "Final", "Windows.Events.ProcessCreation", "CLIENT_EVENT", "Process creation", "TEC-001..009", "Varios", "Varias", "Validar visibilidad", "Monitoring", "Filtrado por ventana", "89 TEC / 18 FP", "Evento crudo", "N/A", "N/A", "CommandLine/parent/PID", "No FP de detección", "Sí para semántica", "Cerrado", "2026-07-12", "...01_ProcessCreation...", "Telemetría", "Sí", "No", "No", "N/A", "No", "VISIBILIDAD", "POSITIVO"),
        @("PUB-FINAL-02", "Cierre", "P4/P2", "Final", "Windows.Events.ServiceCreation", "CLIENT_EVENT", "System 7045", "TEC-005", "T1569.002", "Execution", "Validar creación", "Monitoring", "Filtrado por ventana", "1 TEC / 0 FP", "7045", "N/A", "T1569.002", "ServiceName/ImagePath", "Bajo", "Parcial", "Cerrado", "2026-07-12", "...02_ServiceCreation...", "1053 esperado", "Sí", "No", "No", "N/A", "No", "VISIBILIDAD", "POSITIVO"),
        @("PUB-FINAL-03", "Cierre", "P4/P2", "Final", "Windows.Sysinternals.SysmonLogForward", "CLIENT_EVENT", "Sysmon", "TEC-001..009", "Varios", "Varias", "Validar reenvío", "IDs 1,3,11,12,13", "Filtrado por ventana", "113 TEC / 24 FP", "IDs Sysmon", "N/A", "N/A", "Forense", "No FP de detección", "Sí para correlación", "Cerrado", "2026-07-12", "...03_SysmonLogForward...", "Sin ID14/26", "Sí", "No", "No", "N/A", "No", "FORENSE", "POSITIVO"),
        @("PUB-FINAL-04", "Cierre", "P3", "Final", "Windows.Hayabusa.Monitoring", "CLIENT_EVENT", "Sigma", "TEC-001..009", "Varios", "Varias", "Validar CH", "Plan CH; config efectiva no acreditada", "Filtrado/deduplicado", "185 TEC / 73 FP", "RuleTitle", "high", "Varios", "7 high TEC; 4 high FP", "Alto", "Sí anti-FP", "Cerrado", "2026-07-12", "...04_HayabusaMonitoring_CH...", "187 raw; no CHM", "Sí", "Sí CH 3/9", "Sí", "Hayabusa.Monitoring", "No", "DETECCIÓN SIGMA", "POSITIVO"),
        @("PUB-FINAL-05", "Cierre", "P3", "Final", "Windows.ETW.Monitoring", "CLIENT_EVENT", "ETW/Sigma", "TEC-001..009", "Varios", "Varias", "Validar ETW", "No preservada", "Ventanas y exports", "0 TEC / 0 FP", "Sin coincidencias", "N/A", "N/A", "Exports vacíos", "Indeterminado", "Indeterminado", "Cerrado", "2026-07-12", "...05_ETWMonitoring...", "Falta runner/config TEC", "No acreditada", "No", "No", "N/A", "No", "NO CONCLUYENTE", "NO CONCLUYENTE"),
        @("PUB-FINAL-06", "Cierre", "P4/P2", "Final", "Generic.Events.TrackNetworkConnections", "CLIENT_EVENT", "Conexiones", "TEC-009", "T1048.003 contextual", "Exfiltration", "Validar destino", "Diff + _ts", "Filtrado por _ts", "128 TEC / 12 FP", "3 filas destino", "N/A", "Contextual", "192.168.1.129:8088", "Bajo", "Sí para atribución", "Cerrado", "2026-07-12", "...06_TrackNetworkConnections...", "131 raw; PID0/ProcInfo vacío", "Sí", "No", "No", "N/A", "No por artifact", "VISIBILIDAD RED", "POSITIVO")
    )
    for ($i = 0; $i -lt $summaryRows.Count; $i++) { Set-RowValues $control (62 + $i) $summaryRows[$i] }
    $control.Range("A62:AE67").WrapText = $true
    $control.Range("A4:AE67").VerticalAlignment = $xlVAlignTop
    try { $control.ListObjects.Item("ControlPublicosTable").Resize($control.Range("A4:AE67")) } catch {}

    # Matriz por técnica: no se atribuyen detecciones a proceso crudo ni se inventa CHM.
    $matrix = $Workbook.Worksheets.Item("05_Matriz_Resultados")
    $matrixData = @(
        @("TEC-001", "ID1", "2 high CH", "Sí CH", "TEC-001: PowerShell Base64", "Detección CH específica"),
        @("TEC-002", "ID1", "0 high/critical", "No CH", "Reglas low; no CH", "Visibilidad, sin detección CH"),
        @("TEC-003", "ID1", "1 high CH", "Sí CH", "Scheduled Task", "Detección CH específica"),
        @("TEC-004", "ID12/13", "0 high/critical", "No CH", "Run Key en medium; no CH", "Visibilidad, sin detección CH"),
        @("TEC-005", "ID1/13 + 7045", "4 high en ventana; 3 semánticos", "Sí CH", "Servicio/7045", "Detección CH específica"),
        @("TEC-006", "ID1", "0 high/critical", "No CH", "WMIC low/medium; no CH", "Visibilidad, sin detección CH"),
        @("TEC-007", "ID11/proceso", "0 high/critical", "No CH", "Reglas genéricas; no cifrado específico", "Visibilidad, sin detección CH específica"),
        @("TEC-008", "Sin ID26 en ventana", "0 high/critical", "No CH", "ID26 no observado; sería forense", "Visibilidad, sin detección CH específica"),
        @("TEC-009", "ID1/3/11 + red", "0 high/critical", "No CH", "3 filas destino; HTTP 200 runner", "Visibilidad red; sin detección CH")
    )
    for ($i = 0; $i -lt $matrixData.Count; $i++) {
        $r = 5 + $i
        $d = $matrixData[$i]
        $matrix.Cells.Item($r, 4).Value2 = "Sí: proceso (visibilidad)"
        $matrix.Cells.Item($r, 5).Value2 = if ($d[0] -eq "TEC-005") { "Sí: 7045" } else { "N/A" }
        $matrix.Cells.Item($r, 6).Value2 = $d[1]
        $matrix.Cells.Item($r, 7).Value2 = "Histórico; no RT"
        $matrix.Cells.Item($r, 8).Value2 = "NO EJECUTADA"
        $matrix.Cells.Item($r, 9).Value2 = $d[2]
        $matrix.Cells.Item($r, 10).Value2 = "NO EJECUTADA"
        if ($d[0] -eq "TEC-009") { $matrix.Cells.Item($r, 11).Value2 = "TrackNetwork: 3 filas destino por _ts; PID=0" }
        $matrix.Cells.Item($r, 12).Value2 = $d[5]
        $matrix.Cells.Item($r, 14).Value2 = $d[4]
        $matrix.Cells.Item($r, 15).Value2 = "Custom P1-P4 se conserva como cobertura final"
        $matrix.Cells.Item($r, 16).Value2 = "Sí"
        $matrix.Cells.Item($r, 17).Value2 = $d[3]
        $matrix.Cells.Item($r, 18).Value2 = $d[3]
        $matrix.Cells.Item($r, 19).Value2 = if ($d[3] -eq "Sí CH") { "Windows.Hayabusa.Monitoring" } else { "N/A" }
        $matrix.Cells.Item($r, 20).Value2 = "Cubierto por custom validado"
    }
    try { $matrix.ListObjects.Item("MatrizResultadosTable").Resize($matrix.Range("A4:T13")) } catch {}
    $matrix.Range("A4:T13").WrapText = $true

    # Evaluación por técnica: enlazar la nueva evidencia pública sin degradar P1-P4.
    $eval = $Workbook.Worksheets.Item("03_Evaluacion_VR")
    $evalEvidence = @(
        "ProcessCreation visible + Hayabusa CH: 2 high",
        "ProcessCreation/Sysmon visibles; 0 high/critical CH",
        "ProcessCreation visible + Hayabusa CH: 1 high",
        "Sysmon 12/13 visible; reglas medium fuera de CH",
        "ServiceCreation 7045 + Hayabusa: 4 high (3 semánticos)",
        "WMIC/proceso visible; 0 high/critical CH",
        "Proceso/ID11 y reglas genéricas; sin regla CH específica de cifrado",
        "Proceso visible; sin ID26 y sin regla CH específica de destrucción",
        "TrackNetwork: 3 filas destino sin PID; runner HTTP 200/UploadSucceeded=True"
    )
    for ($i = 0; $i -lt 9; $i++) { $eval.Cells.Item(5 + $i, 12).Value2 = $evalEvidence[$i] }

    $summary = $Workbook.Worksheets.Item("06_Resumen")
    $summary.Range("B8").Value2 = 1
    $summary.Range("C8").Value2 = "Hayabusa.Monitoring CH es el único artifact público con detección por reglas; alcanza 3/9 TEC. ServiceCreation aporta visibilidad 7045."
    $summary.Range("B9").Value2 = 3
    $summary.Range("C9").Value2 = "ProcessCreation, SysmonLogForward y TrackNetwork aportan visibilidad/forense; no detección por sí solos."
    $summary.Range("B10").Value2 = "NO CONCLUYENTE"
    $summary.Range("C10").Value2 = "0 coincidencias, pero faltan runner/log/summary TEC y manifiesto de parámetros; no se clasifica como negativo válido."
    $summary.Range("A17").Value2 = "TEC-009"
    $summary.Range("B17").Value2 = "HTTP 200 confirmado por runner"
    $summary.Range("C17").Value2 = "UploadSucceeded=True; ZIP 6245 B; SHA-256 y ReceiverResponse coinciden. TrackNetwork observa destino sin PID/proceso."
    $summary.Range("B20").Value2 = "5 positivos / 1 no concluyente"
    $summary.Range("C20").Value2 = "Públicos: visibilidad 9/9; detección CH específica 3/9; 4 FP high en 1 caso benigno."

    $exec = $Workbook.Worksheets.Item("RESUMEN_EJECUTIVO")
    Copy-RowFormat -Excel $Excel -Worksheet $exec -SourceRow 24 -FirstTargetRow 25 -LastTargetRow 33 -LastColumn "C"
    $execRows = @(
        @("Campañas públicas ejecutadas", 6, "Seis carpetas esperadas presentes."),
        @("Campañas públicas positivas", 5, "ETW queda no concluyente."),
        @("Campañas públicas no concluyentes", 1, "ETW: evidencia TEC incompleta."),
        @("Eventos públicos TEC canónicos", 516, "521 filas raw; no equivalen todas a detecciones."),
        @("Filas públicas FP en ventana", 127, "Telemetría/coincidencias; capas separadas."),
        @("TEC detectadas por Hayabusa CH", 3, "TEC-001/003/005; high=7, critical=0."),
        @("FP Hayabusa CH", 4, "4 high, 3 eventos subyacentes, 1 caso FP-003."),
        @("TrackNetwork TEC-009", 3, "Filas destino 192.168.1.129:8088 por _ts; PID=0."),
        @("Salida externa TEC-009", "HTTP 200", "Runner independiente: UploadSucceeded=True, 6245 B, SHA-256 coincidente.")
    )
    for ($i = 0; $i -lt $execRows.Count; $i++) { Set-RowValues $exec (25 + $i) $execRows[$i] }
    $exec.Range("A25:C33").WrapText = $true
    $exec.Columns("A").ColumnWidth = 37
    $exec.Columns("B").ColumnWidth = 18
    $exec.Columns("C").ColumnWidth = 66
    $exec.Rows("25:33").AutoFit()
    $exec.Rows(1).RowHeight = 32
    try { $exec.AutoFilterMode = $false; $exec.Range("A4:C33").AutoFilter() | Out-Null } catch {}

    # FP públicos: añadir resumen por artifact sin alterar los 10 FP custom reconciliados.
    $benign = $Workbook.Worksheets.Item("08_Benignas_FP")
    Copy-RowFormat -Excel $Excel -Worksheet $benign -SourceRow 12 -FirstTargetRow 13 -LastTargetRow 18 -LastColumn "J"
    $fpRows = @(
        @("PUB-FP-01", "ProcessCreation", "Ventana FP", "Telemetría", "No detección", "Windows.Events.ProcessCreation", "18 eventos", "NO FP", "Bajo", "Proceso observado no es detección."),
        @("PUB-FP-02", "ServiceCreation", "Ventana FP", "TEC-005", "0", "Windows.Events.ServiceCreation", "0 eventos", "NO FP", "Bajo", "El export raw repetía 7045 TEC fuera de ventana; excluido."),
        @("PUB-FP-03", "SysmonLogForward", "Ventana FP", "Forense", "No detección", "SysmonLogForward", "24 eventos", "NO FP", "Bajo", "Telemetría IDs 1/3/11; no regla."),
        @("PUB-FP-04", "Hayabusa CH", "FP-003 benigno", "Scheduled Task", "0 high/critical deseado", "Hayabusa.Monitoring", "4 high / 3 eventos", "FP CONFIRMADO", "Alto", "Un caso benigno afectado; 0 critical. 73 filas totales en ventana."),
        @("PUB-FP-05", "ETW", "Ventana FP", "Sin coincidencias", "No evaluable", "ETW.Monitoring", "0 eventos", "NO CONCLUYENTE", "Indeterminado", "Falta evidencia TEC/configuración para adjudicar campaña."),
        @("PUB-FP-06", "TrackNetwork", "Ventana FP", "Telemetría red", "No detección", "TrackNetworkConnections", "12 eventos", "NO FP", "Bajo", "No confundir conexión con alerta injustificada.")
    )
    for ($i = 0; $i -lt $fpRows.Count; $i++) { Set-RowValues $benign (13 + $i) $fpRows[$i] }
    $benign.Range("A13:J18").WrapText = $true
    try { $benign.ListObjects.Item("BenignasFPTable").Resize($benign.Range("A4:J18")) } catch {}

    # Conservar Hayabusa.Rules histórico y añadir la campaña Monitoring CH real.
    $hay = $Workbook.Worksheets.Item("13_Hayabusa_Resumen")
    Copy-RowFormat -Excel $Excel -Worksheet $hay -SourceRow 19 -FirstTargetRow 21 -LastTargetRow 35 -LastColumn "D"
    Set-RowValues $hay 21 @("Campaña Windows.Hayabusa.Monitoring CH · 12/07/2026", "", "", "")
    $hay.Range("A21:D21").Merge()
    $hay.Range("A21").Font.Bold = $true
    $hayRows = @(
        @("Ventana TEC UTC", "13:25:03.318–13:31:18.641", "", ""),
        @("Matches TEC canónicos", 185, "187 filas raw", "Se filtra por ventana y se deduplica con clave semántica."),
        @("Niveles únicos TEC", "informational=92; low=38; medium=48; high=7; critical=0", "", "Anomalía: salida contiene niveles inferiores pese a control CH."),
        @("Detección CH por TEC", "3/9", "TEC-001, TEC-003, TEC-005", "Solo high/critical para el veredicto CH."),
        @("Ventana FP UTC", "13:41:00.378–13:41:29.197", "", ""),
        @("Filas FP en ventana", 73, "", "Incluye niveles inferiores observados."),
        @("FP CH", "4 high; 0 critical", "3 eventos subyacentes", "Un caso benigno: FP-003 Scheduled Task."),
        @("Campaña CHM", "NO EXISTE", "", "No se añaden resultados Medium como campaña separada."),
        @("Fuente REAL", "...04_HayabusaMonitoring_CH\Windows.Hayabusa.Monitoring_REAL.json", "", ""),
        @("Fuente FP", "...04_HayabusaMonitoring_CH\Windows.Hayabusa.Monitoring_FP.json", "", ""),
        @("Veredicto", "POSITIVO CON FP", "", "Detección Sigma en tiempo de ejecución; salida externa no confirmada."),
        @("Regla metodológica", "Coincidencia Sigma=detec.; CLIENT_EVENT=unidad primaria", "", "Discord/JSONL no sustituyen la detección."),
        @("Limitación", "Configuración efectiva no preservada fuera del control del workbook", "", "Se reportan niveles realmente observados."),
        @("Estado", "CERRADO", "", "Auditoría final 12/07/2026.")
    )
    $hay.Range("B25").NumberFormat = "@"
    for ($i = 0; $i -lt $hayRows.Count; $i++) { Set-RowValues $hay (22 + $i) $hayRows[$i] }
    $hay.Range("A21:D35").WrapText = $true
    $hay.Columns("B").ColumnWidth = 38
    $hay.Columns("C").ColumnWidth = 25
    $hay.Columns("D").ColumnWidth = 58
    $hay.Rows("21:35").AutoFit()

    # Cerrar contradicciones vigentes.
    $inc = $Workbook.Worksheets.Item("17_Incoherencias_Cerradas")
    $inc.Range("C6").Value2 = "La ejecución final del 12/07 confirma HTTP 200, UploadSucceeded=True, ZIP 6245 B, SHA-256 y bytes del ReceiverResponse."
    $inc.Range("D6").Value2 = "Afirmar transferencia HTTP controlada del runner; separar de la visibilidad TrackNetwork."
    Copy-RowFormat -Excel $Excel -Worksheet $inc -SourceRow 14 -FirstTargetRow 15 -LastTargetRow 19 -LastColumn "D"
    $incRows = @(
        @("INC-011", "Hayabusa CH", "La salida contiene niveles informational/low/medium pese al control CH; el veredicto CH usa solo high/critical.", "No inventar campaña CHM."),
        @("INC-012", "ETW vacío", "0 filas no es negativo válido porque faltan runner/log/summary TEC y parámetros aplicados.", "Clasificar NO CONCLUYENTE."),
        @("INC-013", "TrackNetwork TEC-009", "Tres filas destino se sitúan en ventana por _ts, pero Timestamp=1601, PID=0 y ProcInfo vacío.", "No atribuir la conexión a PowerShell."),
        @("INC-014", "Filtrado temporal", "Exports contienen filas fuera de ventana y duplicados; se usan conteos filtrados/deduplicados documentados.", "No agregar conteos raw."),
        @("INC-015", "Capas", "El HTTP 200 del runner acredita salida externa; TrackNetwork acredita conexión visible, no el POST ni su éxito.", "Mantener capas separadas.")
    )
    for ($i = 0; $i -lt $incRows.Count; $i++) { Set-RowValues $inc (15 + $i) $incRows[$i] }
    $inc.Range("A15:D19").WrapText = $true

    $disc = $Workbook.Worksheets.Item("DISCREPANCIAS")
    Copy-RowFormat -Excel $Excel -Worksheet $disc -SourceRow 52 -FirstTargetRow 53 -LastTargetRow 58 -LastColumn "E"
    $discRows = @(
        @("PUBLICOS_FUERA_DE_VENTANA", "INFO", "04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS", "Conteos raw superan los conteos canónicos en las seis campañas.", "Filtrar por ventanas TEC/FP; no sumar raw."),
        @("HAYABUSA_DUPLICADOS", "INFO", "04_HayabusaMonitoring_CH", "187 filas TEC; 185 matches únicos.", "Conservar ambos conteos y usar únicos para distribución por nivel."),
        @("HAYABUSA_CONFIG_SALIDA", "WARN", "04_HayabusaMonitoring_CH", "Control CH, pero salida contiene informational/low/medium.", "Veredicto CH solo high/critical; no campaña CHM."),
        @("ETW_EVIDENCIA_TEC_INCOMPLETA", "WARN", "05_ETWMonitoring", "Faltan runner/log/summary TEC y manifiesto de parámetros.", "NO CONCLUYENTE."),
        @("TRACKNETWORK_TIMESTAMP_PROCESO", "WARN", "06_TrackNetworkConnections", "Filas objetivo con Timestamp=1601, PID=0 y ProcInfo vacío; _ts sí cae en ventana.", "Conexión visible sin atribución a proceso."),
        @("TEC009_CAPAS", "INFO", "06_TrackNetworkConnections + summary TEC", "TrackNetwork ve destino; runner confirma POST HTTP 200/6245 B/hash.", "No atribuir el éxito HTTP al artifact de red.")
    )
    for ($i = 0; $i -lt $discRows.Count; $i++) { Set-RowValues $disc (53 + $i) $discRows[$i] }
    $disc.Range("A53:E58").WrapText = $true
    try { $disc.AutoFilterMode = $false; $disc.Range("A4:E58").AutoFilter() | Out-Null } catch {}

    # Comparación metodológica: nueva fila pública, manteniendo Wazuh dentro del mismo libro.
    $compare = $Workbook.Worksheets.Item("COMPARACION_VR_WAZUH")
    Copy-RowFormat -Excel $Excel -Worksheet $compare -SourceRow 7 -FirstTargetRow 8 -LastTargetRow 8 -LastColumn "M"
    Set-RowValues $compare 8 @(
        "Artifacts públicos 12/07",
        "521 filas heterogéneas en ventana; no comparables con archives/CLIENT_EVENT custom",
        "CH high=7; 3/9 TEC específicas; Service 7045=1",
        "Incluido en Process/Hayabusa",
        "Sysmon ID1 incluido en reenvío",
        "Sysmon ID11 incluido en reenvío",
        "Service 7045=1",
        "No aplica: artifacts públicos",
        "No aplica",
        "Conexión destino visible; HTTP 200 solo runner",
        "Hayabusa FP CH=4 matches/1 caso",
        "ETW no concluyente; TrackNetwork sin PID/proceso",
        "Comparación por capacidades/capas, no por volumen bruto"
    )
    $compare.Range("A8:M8").WrapText = $true
    try { $compare.AutoFilterMode = $false; $compare.Range("A4:M8").AutoFilter() | Out-Null } catch {}

    # Trazabilidad exhaustiva de los 112 ficheros de evidencia pública.
    $sources = $Workbook.Worksheets.Item("FUENTES")
    $files = @(Get-ChildItem -LiteralPath $EvidenceRootPath -Recurse -File -Force | Sort-Object FullName)
    $startRow = 107
    Copy-RowFormat -Excel $Excel -Worksheet $sources -SourceRow 106 -FirstTargetRow $startRow -LastTargetRow ($startRow + $files.Count - 1) -LastColumn "G"
    for ($i = 0; $i -lt $files.Count; $i++) {
        $f = $files[$i]
        $relative = $f.FullName.Substring($EvidenceRootPath.TrimEnd('\').Length + 1)
        $campaign = $relative.Split([char]'\')[0]
        Set-RowValues $sources ($startRow + $i) @(
            "04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\$relative",
            $f.Extension.TrimStart('.').ToLowerInvariant(),
            [double]$f.Length,
            $f.LastWriteTimeUtc.ToString("yyyy-MM-ddTHH:mm:ss.fff'Z'"),
            (Get-FileHash -Algorithm SHA256 -LiteralPath $f.FullName).Hash,
            (Get-EvidenceRole $f),
            $campaign
        )
    }
    $sources.Range("A$startRow:G$($startRow + $files.Count - 1)").WrapText = $true
    $sources.Range("C$startRow:C$($startRow + $files.Count - 1)").NumberFormat = "0"
    try { $sources.AutoFilterMode = $false; $sources.Range("A4:G$($startRow + $files.Count - 1)").AutoFilter() | Out-Null } catch {}

    # Helper auditable y cinco gráficos nuevos en el dashboard.
    $graphs = $Workbook.Worksheets.Item("GRAFICAS")
    $graphs.Range("A60:R80").Clear()
    Set-RowValues $graphs 60 @("TEC", "Visibilidad pública", "Detección CH high/critical")
    $detectedPublic = @("TEC-001", "TEC-003", "TEC-005")
    for ($i = 1; $i -le 9; $i++) {
        $tec = "TEC-{0:D3}" -f $i
        Set-RowValues $graphs (60 + $i) @($tec, 1, $(if ($detectedPublic -contains $tec) { 1 } else { 0 }))
    }
    Set-RowValues $graphs 60 @("Sistema", "TEC con detección específica") 6
    Set-RowValues $graphs 61 @("Custom CLIENT_EVENT", "") 6
    Set-RowValues $graphs 62 @("Públicos CH", "") 6
    Set-RowValues $graphs 63 @("Wazuh 110xxx", "") 6
    Set-RowValues $graphs 60 @("Sistema", "Visibilidad", "Detección", "Alerta RT", "Salida externa") 9
    Set-RowValues $graphs 61 @("Custom", 9, 9, 9, 0) 9
    Set-RowValues $graphs 62 @("Públicos", 9, 3, 3, 0) 9
    Set-RowValues $graphs 60 @("TEC", "Custom", "Públicos CH", "Wazuh específico") 15
    for ($i = 1; $i -le 9; $i++) {
        $r = 60 + $i
        $tec = "TEC-{0:D3}" -f $i
        $graphs.Cells.Item($r, 15).Value2 = $tec
        $graphs.Cells.Item($r, 16).Formula = '=IF(TECNICAS_REAL_VR!O' + (4 + $i) + '="SI_CLIENT_EVENT",1,0)'
        $publicDetection = if ($detectedPublic -contains $tec) { [double]1 } else { [double]0 }
        Set-RowValues $graphs $r ([object[]]@($publicDetection)) 17
        $graphs.Cells.Item($r, 18).Formula = '=--(SUMIF(WAZUH_DETALLE!$D$5:$D$11,"*' + $tec + '*",WAZUH_DETALLE!$F$5:$F$11)>0)'
    }
    $graphs.Range("G61").Formula = "=SUM(P61:P69)"
    $graphs.Range("G62").Formula = "=SUM(Q61:Q69)"
    $graphs.Range("G63").Formula = "=SUM(R61:R69)"
    Set-RowValues $graphs 73 @("Artifact", "Filas en ventana FP", "FP CH high/critical")
    $fpArtifactRows = @(
        @("Hayabusa Monitoring CH", 5), @("ServiceCreation", 6), @("ProcessCreation", 7),
        @("SysmonLogForward", 8), @("ETW Monitoring", 9), @("TrackNetworkConnections", 10)
    )
    for ($i = 0; $i -lt $fpArtifactRows.Count; $i++) {
        $row = 74 + $i
        $sourceRow = $fpArtifactRows[$i][1]
        $graphs.Cells.Item($row, 1).Value2 = $fpArtifactRows[$i][0]
        $graphs.Cells.Item($row, 2).Formula = "='15_Publicos_Definitivo'!N$sourceRow"
        $graphs.Cells.Item($row, 3).Formula = "='15_Publicos_Definitivo'!O$sourceRow"
    }
    Set-RowValues $graphs 73 @("Resultado", "Campañas") 6
    Set-RowValues $graphs 74 @("POSITIVO", "") 6
    Set-RowValues $graphs 75 @("NEGATIVO VÁLIDO", "") 6
    Set-RowValues $graphs 76 @("NO CONCLUYENTE", "") 6
    Set-RowValues $graphs 77 @("NO EJECUTADO", "") 6
    Set-RowValues $graphs 78 @("NO APLICABLE", "") 6
    $graphs.Range("G74").Formula = '=COUNTIF(''15_Publicos_Definitivo''!$R$5:$R$10,"POSITIVO")'
    $graphs.Range("G75").Formula = '=COUNTIF(''15_Publicos_Definitivo''!$R$5:$R$10,"NEGATIVO VÁLIDO")'
    $graphs.Range("G76").Formula = '=COUNTIF(''15_Publicos_Definitivo''!$R$5:$R$10,"NO CONCLUYENTE")'
    $graphs.Range("G77").Formula = '=--(COUNTIF(''04_Control_Publicos''!$U$29:$U$55,"NO EJECUTADA")>0)'
    $graphs.Range("G78").Formula = '=COUNTIF(''15_Publicos_Definitivo''!$R$5:$R$10,"NO APLICABLE")'
    try { $graphs.Range("A60:R80").NumberFormat = "General" } catch {}
    $graphs.Range("A60:R80").WrapText = $true
    foreach ($headerRange in @("A60:C60", "F60:G60", "I60:M60", "O60:R60", "A73:C73", "F73:G73")) {
        try { $graphs.Range($headerRange).Font.Bold = $true } catch {}
    }

    # Reparar gráfico Wazuh sin categorías.
    try {
        $wazuhChart = $graphs.ChartObjects("Chart 5").Chart
        while ($wazuhChart.SeriesCollection().Count -gt 0) { $wazuhChart.SeriesCollection(1).Delete() }
        $series = $wazuhChart.SeriesCollection().NewSeries()
        $series.Name = "=GRAFICAS!`$B`$34"
        $series.XValues = "=GRAFICAS!`$A`$35:`$A`$41"
        $series.Values = "=GRAFICAS!`$B`$35:`$B`$41"
        Configure-Chart -Chart $wazuhChart -Title "Alertas por regla Wazuh 110xxx" -CategoryTitle "Regla" -ValueTitle "Alertas" -DataLabels
    } catch {}

    # Conservar y reparar el gráfico previo del dashboard; eliminar solo gráficos finales reejecutables.
    foreach ($co in @($dashboard.ChartObjects())) {
        if ($co.Name -like "TFM_Final_*") { $co.Delete() }
    }
    try {
        $existing = $dashboard.ChartObjects(1)
        Configure-Chart -Chart $existing.Chart -Title "Cobertura táctica MITRE /14 (criterio tutoría)" -CategoryTitle "Cobertura" -ValueTitle "Proporción" -ValueNumberFormat "0%" -DataLabels
        Set-ChartPosition -Worksheet $dashboard -ChartObject $existing -TopLeft "J54" -BottomRight "Q70"
    } catch {}
    Add-DashboardChart $dashboard $graphs "TFM_Final_Public_TEC" "A60:C69" "Visibilidad y detección CH por técnica" "Técnica" "TEC (0/1)" "A18" "H34" | Out-Null
    Add-DashboardChart $dashboard $graphs "TFM_Final_Compare_Detection" "F60:G63" "Detección específica demostrada" "Sistema" "Nº de TEC" "J18" "Q34" | Out-Null
    Add-DashboardChart $dashboard $graphs "TFM_Final_Layers" "I60:M62" "Capas demostradas: custom frente a públicos" "Sistema" "Nº de TEC" "A36" "H52" | Out-Null
    Add-DashboardChart $dashboard $graphs "TFM_Final_FP" "A73:C79" "Ventana FP por artifact (capas separadas)" "Artifact" "Filas / matches" "J36" "Q52" | Out-Null
    Add-DashboardChart $dashboard $graphs "TFM_Final_Outcomes" "F73:G78" "Campañas públicas y control CHM" "Resultado" "Campañas / controles" "A54" "H70" | Out-Null
    $dashboard.PageSetup.PrintArea = "`$A`$1:`$Q`$70"
    $dashboard.Activate()
    $dashboard.Range("A1").Select()

    # Metadatos.
    try { $Workbook.BuiltinDocumentProperties("Title").Value = "TFM - Visibilidad, detección, FP y Wazuh - Definitivo 20260712" } catch {}
    try { $Workbook.BuiltinDocumentProperties("Subject").Value = "Cierre técnico trazable de Velociraptor y Wazuh" } catch {}
    try { $Workbook.BuiltinDocumentProperties("Keywords").Value = "TFM;Velociraptor;Wazuh;HIDS;DFIR;MITRE;20260712" } catch {}
}

function Update-BenchmarkWorkbook {
    param($Excel, $Workbook)
    $readme = $Workbook.Worksheets.Item("README")
    $readme.Range("A3").Value2 = "Fecha de revisión final"
    $readme.Range("B3").Value2 = "2026-07-12 (Europe/Madrid)"
    Set-RowValues $readme 12 @("Estado final", "APTO: 9/9 runs válidos; SERVER_GUI excluido; gráficos reparados y unidades separadas.")
    $readme.Range("A12:B12").WrapText = $true
    try { $readme.AutoFilterMode = $false; $readme.Range("A1:B12").AutoFilter() | Out-Null } catch {}

    $exec = $Workbook.Worksheets.Item("RESUMEN_EJECUTIVO")
    $exec.Range("C3").NumberFormat = "@"
    $exec.Range("C3").Value2 = "9/9"
    $exec.Range("B5").NumberFormat = "@"
    $exec.Range("B5").Value2 = "3/3"
    $exec.Range("B2:B4").NumberFormat = "0"

    $quality = $Workbook.Worksheets.Item("CALIDAD_DATOS")
    $quality.Range("C6").NumberFormat = "@"
    $quality.Range("C6").Value2 = "9/9"
    $quality.Range("C18").Value2 = "Fórmulas helper transparentes añadidas solo para el gráfico de repeticiones; métricas fuente no modificadas."
    $quality.Range("C19").Value2 = "Seis gráficos con series no vacías; comparación de repeticiones usa solo duración (s)."
    Copy-RowFormat -Excel $Excel -Worksheet $quality -SourceRow 20 -FirstTargetRow 21 -LastTargetRow 22 -LastColumn "E"
    Set-RowValues $quality 21 @("Unidades de gráficos separadas", "OK", "CPU %, RAM MB y duración s no comparten eje.", "CRITICA", "Gráfico 6 reparado.")
    Set-RowValues $quality 22 @("Comparación tres repeticiones", "OK", "REP_01/02/03 por los tres escenarios; duración en segundos.", "CRITICA", "Fuente RUNS_VALIDOS F2:F10.")

    $disc = $Workbook.Worksheets.Item("DISCREPANCIAS")
    Copy-RowFormat -Excel $Excel -Worksheet $disc -SourceRow 4 -FirstTargetRow 5 -LastTargetRow 6 -LastColumn "E"
    Set-RowValues $disc 5 @("FORMATO_CORREGIDO", "RESUMEN_EJECUTIVO", "C3 y B5 se interpretaban como fechas por formato/valor serial.", "Fijados como texto 9/9 y 3/3.", "Auditoría Excel COM/openpyxl 12/07/2026")
    Set-RowValues $disc 6 @("GRAFICO_REPARADO", "GRAFICAS Chart 6", "Mezclaba CPU %, RAM MB y duración s en un eje y no comparaba repeticiones.", "Sustituido por duración por repetición y escenario (s).", "RUNS_VALIDOS F2:F10")
    $disc.Range("A5:E6").WrapText = $true

    $graphs = $Workbook.Worksheets.Item("GRAFICAS")
    Set-RowValues $graphs 24 @("Repetición", "BASELINE_NO_VR (s)", "VR_IDLE (s)", "VR_TEC_RUNNER (s)")
    for ($i = 1; $i -le 3; $i++) {
        $r = 24 + $i
        $graphs.Cells.Item($r, 1).Value2 = "REP_{0:D2}" -f $i
        $graphs.Cells.Item($r, 2).Formula = "=RUNS_VALIDOS!F$($i + 1)"
        $graphs.Cells.Item($r, 3).Formula = "=RUNS_VALIDOS!F$($i + 4)"
        $graphs.Cells.Item($r, 4).Formula = "=RUNS_VALIDOS!F$($i + 7)"
    }
    $graphs.Range("A24:D24").Font.Bold = $true
    $graphs.Range("B25:D27").NumberFormat = "0.00"

    $charts = @($graphs.ChartObjects())
    $chartPositions = @(
        @("A57", "G74"), @("H57", "N74"),
        @("A76", "G93"), @("H76", "N93"),
        @("A95", "G112"), @("H95", "N112")
    )
    for ($i = 1; $i -le $charts.Count; $i++) {
        $chart = $charts[$i - 1].Chart
        switch ($i) {
            1 { Configure-Chart $chart "CPU media del cliente por escenario (%)" "Escenario" "CPU media (%)" "0.00" -DataLabels }
            2 { Configure-Chart $chart "CPU máxima del cliente por escenario (%)" "Escenario" "CPU máxima (%)" "0.00" -DataLabels }
            3 { Configure-Chart $chart "RAM media del cliente por escenario (MB)" "Escenario" "RAM media (MB)" "0.00" -DataLabels }
            4 { Configure-Chart $chart "RAM máxima del cliente por escenario (MB)" "Escenario" "RAM máxima (MB)" "0.00" -DataLabels }
            5 { Configure-Chart $chart "Resumen: duración media por escenario (s)" "Escenario" "Duración media (s)" "0.00" -DataLabels }
            6 {
                $chart.ChartType = $xlColumnClustered
                $chart.SetSourceData($graphs.Range("A24:D27"))
                Configure-Chart $chart "Duración por repetición y escenario (s)" "Escenario" "Duración (s)" "0.00" -DataLabels
            }
        }
        if ($i -le $chartPositions.Count) {
            Set-ChartPosition -Worksheet $graphs -ChartObject $charts[$i - 1] -TopLeft $chartPositions[$i - 1][0] -BottomRight $chartPositions[$i - 1][1]
        }
        try { $chart.Refresh() } catch {}
    }
    $graphs.PageSetup.PrintArea = "`$A`$1:`$N`$112"
    try {
        $graphs.PageSetup.Orientation = 2
        $graphs.PageSetup.Zoom = $false
        $graphs.PageSetup.FitToPagesWide = 1
        $graphs.PageSetup.FitToPagesTall = 0
    } catch {}
    $graphs.Activate()
    $Excel.Goto($graphs.Range("A55"), $true)
    try { $Workbook.BuiltinDocumentProperties("Title").Value = "TFM - Benchmark rendimiento Velociraptor - Definitivo 20260712" } catch {}
    try { $Workbook.BuiltinDocumentProperties("Subject").Value = "Benchmark 3 escenarios x 3 repeticiones" } catch {}
    try { $Workbook.BuiltinDocumentProperties("Keywords").Value = "TFM;Velociraptor;benchmark;CPU;RAM;20260712" } catch {}
}

if (-not (Test-Path -LiteralPath $VisibilityPath -PathType Leaf)) { throw "No existe: $VisibilityPath" }
if (-not (Test-Path -LiteralPath $BenchmarkPath -PathType Leaf)) { throw "No existe: $BenchmarkPath" }
if (-not (Test-Path -LiteralPath $EvidenceRoot -PathType Container)) { throw "No existe: $EvidenceRoot" }
New-Item -ItemType Directory -Path $AuditDir -Force | Out-Null

if (-not ("TFMExcelNative" -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class TFMExcelNative {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
"@
}

$excel = $null
$excelPid = 0
$log = [System.Collections.Generic.List[object]]::new()
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    try { $excel.AutomationSecurity = 3 } catch {}
    [uint32]$pidValue = 0
    [void][TFMExcelNative]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$pidValue)
    $excelPid = [int]$pidValue
    try { $excel.Calculation = $xlCalculationAutomatic } catch {}
    try { $excel.CalculateBeforeSave = $true } catch {}

    foreach ($item in @(
        [pscustomobject]@{ Path = $VisibilityPath; Kind = "VISIBILIDAD" },
        [pscustomobject]@{ Path = $BenchmarkPath; Kind = "BENCHMARK" }
    )) {
        $wb = $null
        try {
            $wb = $excel.Workbooks.Open($item.Path, 0, $false)
            try { $excel.Calculation = $xlCalculationAutomatic } catch {}
            try { $excel.CalculateBeforeSave = $true } catch {}
            if ($item.Kind -eq "VISIBILIDAD") {
                Update-VisibilityWorkbook -Excel $excel -Workbook $wb -EvidenceRootPath $EvidenceRoot
            } else {
                Update-BenchmarkWorkbook -Excel $excel -Workbook $wb
            }
            try { $wb.ForceFullCalculation = $true } catch {}
            try { $wb.FullCalculationOnLoad = $true } catch {}
            foreach ($ws in @($wb.Worksheets)) {
                foreach ($co in @($ws.ChartObjects())) { try { $co.Chart.Refresh() } catch {} }
            }
            $excel.CalculateFullRebuild()
            $wb.Save()
            $log.Add([pscustomobject]@{
                Kind = $item.Kind
                Path = $item.Path
                Sheets = $wb.Worksheets.Count
                Connections = $wb.Connections.Count
                Saved = $true
                Time = (Get-Date).ToString("o")
            })
        } finally {
            if ($wb) { $wb.Close($true) }
            if ($wb) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) }
        }
    }
} catch {
    Write-Host ("BUILD_ERROR_POSITION: " + $_.InvocationInfo.PositionMessage)
    Write-Host ("BUILD_ERROR_STACK: " + $_.ScriptStackTrace)
    throw
} finally {
    if ($excel) { $excel.Quit() }
    if ($excel) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 2
    if ($excelPid -gt 0 -and (Get-Process -Id $excelPid -ErrorAction SilentlyContinue)) {
        Stop-Process -Id $excelPid -Force
    }
}

$logPath = Join-Path $AuditDir "EXCEL_COM_BUILD_20260712.json"
$log | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $logPath -Encoding UTF8
$log | Format-Table -AutoSize









