param(
    [string]$Root = "${PSScriptRoot}\..\..\.."
)

$ErrorActionPreference = "Stop"

$TraceDir = Join-Path $Root "04_EVIDENCE\Excel_visibilidad_26062026"
$NormDir = Join-Path $TraceDir "normalized"
$FinalDir = Join-Path $Root "08_MEMORIA\Excel_visibilidad_FP"
$BaseXlsx = Join-Path $Root "04_EVIDENCE\Analisis_Tecnicas_TFM_V.4.xlsx"
$FinalName = "TFM_VISIBILIDAD_DETECCION_FP_WAZUH_FINAL_26062026"
$FinalXlsx = Join-Path $FinalDir "$FinalName.xlsx"
$V5Xlsx = Join-Path $FinalDir "Analisis_Tecnicas_TFM_V.5.xlsx"
$TraceFinalXlsx = Join-Path $TraceDir "$FinalName.xlsx"
$TraceV5Xlsx = Join-Path $TraceDir "Analisis_Tecnicas_TFM_V.5.xlsx"
$AuditFinal = Join-Path $FinalDir "$FinalName`_AUDIT.md"
$AuditTrace = Join-Path $TraceDir "$FinalName`_AUDIT.md"
$DqTrace = Join-Path $TraceDir "$FinalName`_DATA_QUALITY.json"
$DqFinal = Join-Path $FinalDir "$FinalName`_DATA_QUALITY.json"
$LogPath = Join-Path $TraceDir ("generation_log_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".txt")

New-Item -ItemType Directory -Force -Path $TraceDir, $NormDir, $FinalDir | Out-Null

function Read-CsvRows {
    param([string]$Name)
    $p = Join-Path $NormDir $Name
    if (-not (Test-Path -LiteralPath $p)) { throw "Missing normalized file: $p" }
    return @(Import-Csv -LiteralPath $p)
}

function Get-PropNames {
    param([object[]]$Rows)
    if ($Rows.Count -eq 0) { return @() }
    return @($Rows[0].PSObject.Properties.Name)
}

function Add-DataSheet {
    param(
        [object]$Workbook,
        [string]$Name,
        [string]$Title,
        [object[]]$Rows,
        [string[]]$Columns,
        [string]$Note = ""
    )

    foreach ($ws0 in @($Workbook.Worksheets)) {
        if ($ws0.Name -eq $Name) {
            $ws0.Delete()
            break
        }
    }

    $ws = $Workbook.Worksheets.Add([System.Type]::Missing, $Workbook.Worksheets.Item($Workbook.Worksheets.Count))
    $ws.Name = $Name
    $ws.Cells.Font.Name = "Calibri"
    $ws.Cells.Font.Size = 10
    $ws.Range("A1").Value2 = $Title
    $ws.Range("A1").Font.Bold = $true
    $ws.Range("A1").Font.Size = 15
    $ws.Range("A1").Interior.Color = 0x703000
    $ws.Range("A1").Font.Color = 0xFFFFFF
    $ws.Range("A2").Value2 = $Note
    $ws.Range("A2").Font.Italic = $true
    $ws.Range("A2").WrapText = $true

    if ($Columns.Count -eq 0) {
        $ws.Range("A4").Value2 = "Sin datos"
        Write-Output -NoEnumerate $ws
        return
    }

    $rowCount = [Math]::Max(1, $Rows.Count) + 1
    $colCount = $Columns.Count
    $data = New-Object "object[,]" $rowCount, $colCount
    for ($c = 0; $c -lt $colCount; $c++) {
        $data[0, $c] = $Columns[$c]
    }
    if ($Rows.Count -eq 0) {
        for ($c = 0; $c -lt $colCount; $c++) { $data[1, $c] = "" }
    } else {
        for ($r = 0; $r -lt $Rows.Count; $r++) {
            for ($c = 0; $c -lt $colCount; $c++) {
                $v = $Rows[$r].PSObject.Properties[$Columns[$c]].Value
                if ($null -eq $v) { $v = "" }
                $data[($r + 1), $c] = [string]$v
            }
        }
    }

    $startRow = 4
    $startCol = 1
    $endRow = $startRow + $rowCount - 1
    $endCol = $startCol + $colCount - 1
    $range = $ws.Range($ws.Cells.Item($startRow, $startCol), $ws.Cells.Item($endRow, $endCol))
    $range.Value2 = $data

    $header = $ws.Range($ws.Cells.Item($startRow, $startCol), $ws.Cells.Item($startRow, $endCol))
    $header.Font.Bold = $true
    $header.Font.Color = 0xFFFFFF
    $header.Interior.Color = 0x703000
    $header.WrapText = $true

    $range.Borders.LineStyle = 1
    $range.Borders.Color = 0xD9D9D9
    $range.VerticalAlignment = -4160

    $range.AutoFilter() | Out-Null

    $ws.Activate()
    $ws.Application.ActiveWindow.DisplayGridlines = $false
    $ws.Application.ActiveWindow.SplitRow = $startRow
    $ws.Application.ActiveWindow.FreezePanes = $true
    for ($c = 1; $c -le $colCount; $c++) {
        $name = [string]$Columns[$c - 1]
        if ($name -match "evidencia|observaciones|limitaciones|detalle|decision|raw|ruta|Artifact|lectura|metodologica|descripcion") {
            $ws.Columns.Item($c).ColumnWidth = 42
        } elseif ($name -match "timestamp|inicio|fin|fecha") {
            $ws.Columns.Item($c).ColumnWidth = 22
        } elseif ($name -match "valor|alertas|events|seg|bytes|OK|WARN|FAIL|SKIPPED") {
            $ws.Columns.Item($c).ColumnWidth = 12
        } else {
            $ws.Columns.Item($c).ColumnWidth = 18
        }
    }
    $ws.Rows.Item(1).RowHeight = 28
    $ws.Rows.Item(2).RowHeight = 36
    $range.WrapText = $false
    $header.WrapText = $true
    Write-Output -NoEnumerate $ws
    return
}

function Add-Chart {
    param(
        [object]$Worksheet,
        [object]$SourceRange,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [string]$Title
    )
    $co = $Worksheet.ChartObjects().Add($Left, $Top, $Width, $Height)
    $chart = $co.Chart
    $chart.ChartType = 51
    $chart.SetSourceData($SourceRange)
    $chart.HasTitle = $true
    $chart.ChartTitle.Text = $Title
    $chart.HasLegend = $true
    Write-Output -NoEnumerate $co
    return
}

function Write-Grid {
    param(
        [object]$Worksheet,
        [int]$StartRow,
        [int]$StartCol,
        [object[]]$Rows,
        [string[]]$Columns
    )
    $rowCount = $Rows.Count + 1
    $colCount = $Columns.Count
    $data = New-Object "object[,]" $rowCount, $colCount
    for ($c = 0; $c -lt $colCount; $c++) { $data[0, $c] = $Columns[$c] }
    for ($r = 0; $r -lt $Rows.Count; $r++) {
        for ($c = 0; $c -lt $colCount; $c++) {
            $data[($r + 1), $c] = $Rows[$r].PSObject.Properties[$Columns[$c]].Value
        }
    }
    $range = $Worksheet.Range($Worksheet.Cells.Item($StartRow, $StartCol), $Worksheet.Cells.Item($StartRow + $rowCount - 1, $StartCol + $colCount - 1))
    $range.Value2 = $data
    $range.Borders.LineStyle = 1
    $range.Borders.Color = 0xD9D9D9
    $hdr = $Worksheet.Range($Worksheet.Cells.Item($StartRow, $StartCol), $Worksheet.Cells.Item($StartRow, $StartCol + $colCount - 1))
    $hdr.Font.Bold = $true
    $hdr.Interior.Color = 0x703000
    $hdr.Font.Color = 0xFFFFFF
    Write-Output -NoEnumerate $range
    return
}

function Apply-Heatmap {
    param(
        [object]$Worksheet,
        [int]$StartRow,
        [int]$StartCol,
        [int]$Rows,
        [int]$Cols
    )
    for ($r = 1; $r -le $Rows; $r++) {
        for ($c = 1; $c -le $Cols; $c++) {
            $cell = $Worksheet.Cells.Item($StartRow + $r, $StartCol + $c)
            $v = 0
            [void][double]::TryParse([string]$cell.Value2, [ref]$v)
            if ($v -le 0) {
                $cell.Interior.Color = 0xE7E6E6
            } elseif ($v -eq 1) {
                $cell.Interior.Color = 0xB7DEE8
            } else {
                $cell.Interior.Color = 0x92D050
            }
        }
    }
}

$kpis = Read-CsvRows "kpis.csv"
$tecnicas = Read-CsvRows "tecnicas_real_vr.csv"
$visibilidad = Read-CsvRows "visibilidad_sistema.csv"
$alertas = Read-CsvRows "alertas_vr.csv"
$fpRunner = Read-CsvRows "fp_runner.csv"
$fpHits = Read-CsvRows "fp_hits.csv"
$comparacion = Read-CsvRows "comparacion_vr_wazuh.csv"
$wazuhDetalle = Read-CsvRows "wazuh_detalle.csv"
$discrepancias = Read-CsvRows "discrepancias.csv"
$fuentes = Read-CsvRows "fuentes.csv"
$chartEvents = Read-CsvRows "chart_events_by_tec.csv"
$chartProfiles = Read-CsvRows "chart_alerts_by_profile.csv"
$chartFp = Read-CsvRows "chart_fp.csv"
$chartWazuh = Read-CsvRows "chart_wazuh.csv"
$chartRules = Read-CsvRows "chart_wazuh_rules.csv"
$heatVisibility = Read-CsvRows "heatmap_visibility.csv"
$heatCapacity = Read-CsvRows "heatmap_capacity.csv"
$dq = Get-Content -Raw -LiteralPath $DqTrace | ConvertFrom-Json

Copy-Item -LiteralPath $BaseXlsx -Destination $FinalXlsx -Force

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Open($FinalXlsx)

    $readmeRows = @(
        [pscustomobject]@{ campo = "Objetivo"; valor = "Consolidar visibilidad, deteccion, alertas/perfiles, FP y comparacion Wazuh para memoria TFM." },
        [pscustomobject]@{ campo = "Fecha de generacion"; valor = (Get-Date -Format "yyyy-MM-dd HH:mm:ss K") },
        [pscustomobject]@{ campo = "Excel base"; valor = "04_EVIDENCE\Analisis_Tecnicas_TFM_V.4.xlsx" },
        [pscustomobject]@{ campo = "Rutas VR"; valor = "05_LOGS\Validas\Prueba_24_06_2026 y 05_LOGS\Validas\Prueba_24_06_2026_FP" },
        [pscustomobject]@{ campo = "Rutas Wazuh"; valor = "10_WAZUH\TFM_WAZUH_VISIBILIDAD_EVIDENCE" },
        [pscustomobject]@{ campo = "Criterio REAL"; valor = "Solo filas CLIENT_EVENT de carpeta real dentro de $($dq.windows.real_start_utc)..$($dq.windows.real_end_utc)." },
        [pscustomobject]@{ campo = "Criterio FP"; valor = "FP por summary/vr_hits del RunId TFM_FP_20260626_101948 dentro de $($dq.windows.fp_start_utc)..$($dq.windows.fp_end_utc)." },
        [pscustomobject]@{ campo = "Advertencia"; valor = "CSV con fechas mezcladas se registran en DISCREPANCIAS y no se agregan como definitivos fuera de ventana." },
        [pscustomobject]@{ campo = "Limitaciones"; valor = "No se localiza receiver_log.jsonl/ZIP recibido de TEC-009; JSONL/Discord no se tratan como deteccion." }
    )

    Add-DataSheet $wb "README" "README - alcance y metodologia" $readmeRows @("campo", "valor") "Libro generado desde V.4 con normalizacion trazable." | Out-Null
    Add-DataSheet $wb "RESUMEN_EJECUTIVO" "Resumen ejecutivo" $kpis (Get-PropNames $kpis) "KPIs principales y conteos reconciliados." | Out-Null
    Add-DataSheet $wb "TECNICAS_REAL_VR" "Tecnicas reales Velociraptor" $tecnicas (Get-PropNames $tecnicas) "TEC-001 a TEC-009; ejecucion real 24/06; filas fuera de ventana excluidas." | Out-Null
    Add-DataSheet $wb "VISIBILIDAD_SISTEMA" "Visibilidad del sistema por fuente" $visibilidad (Get-PropNames $visibilidad) "Cruce fuente vs tecnica; Sysmon ID 26 se mantiene como evidencia forense." | Out-Null
    Add-DataSheet $wb "ALERTAS_VR" "Alertas / detecciones VR normalizadas" $alertas (Get-PropNames $alertas) "Filas CLIENT_EVENT reales dentro de ventana canonica. JSONL/Discord no equivalen a deteccion." | Out-Null
    Add-DataSheet $wb "FP_RUNNER" "Runner de falsos positivos" $fpRunner (Get-PropNames $fpRunner) "FP-001 a FP-010; RunId TFM_FP_20260626_101948." | Out-Null
    Add-DataSheet $wb "FP_HITS" "Hits FP normalizados" $fpHits (Get-PropNames $fpHits) "TotalHits=0; no se inventan filas de deteccion." | Out-Null
    Add-DataSheet $wb "COMPARACION_VR_WAZUH" "Comparacion VR / Wazuh" $comparacion (Get-PropNames $comparacion) "Comparacion metodologica base/custom y VR custom." | Out-Null
    Add-DataSheet $wb "WAZUH_DETALLE" "Wazuh detalle reglas 110xxx" $wazuhDetalle (Get-PropNames $wazuhDetalle) "110201=0 se conserva como gap TEC-009." | Out-Null
    Add-DataSheet $wb "DISCREPANCIAS" "Discrepancias y decisiones automaticas" $discrepancias (Get-PropNames $discrepancias) "Eventos ambiguos, mezclas de fechas y exclusiones." | Out-Null
    Add-DataSheet $wb "FUENTES" "Fuentes analizadas y hashes" $fuentes (Get-PropNames $fuentes) "Lista de fuentes usadas con SHA-256." | Out-Null

    foreach ($ws0 in @($wb.Worksheets)) {
        if ($ws0.Name -eq "GRAFICAS") { $ws0.Delete(); break }
    }
    $g = $wb.Worksheets.Add([System.Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
    $g.Name = "GRAFICAS"
    $g.Cells.Font.Name = "Calibri"
    $g.Cells.Font.Size = 10
    $g.Range("A1").Value2 = "GRAFICAS - memoria"
    $g.Range("A1").Font.Bold = $true
    $g.Range("A1").Font.Size = 15
    $g.Range("A1").Interior.Color = 0x703000
    $g.Range("A1").Font.Color = 0xFFFFFF
    $g.Application.ActiveWindow.DisplayGridlines = $false

    $r1 = Write-Grid $g 3 1 $chartEvents (Get-PropNames $chartEvents)
    Add-Chart $g $r1 430 40 520 250 "Eventos Sysmon/4104 por TEC" | Out-Null
    $r2 = Write-Grid $g 3 7 $chartProfiles (Get-PropNames $chartProfiles)
    Add-Chart $g $r2 980 40 360 250 "Alertas VR por perfil" | Out-Null
    $r3 = Write-Grid $g 18 1 $chartFp (Get-PropNames $chartFp)
    Add-Chart $g $r3 430 330 520 250 "FP OK/SKIPPED/HITS" | Out-Null
    $r4 = Write-Grid $g 18 7 $chartWazuh (Get-PropNames $chartWazuh)
    Add-Chart $g $r4 980 330 520 250 "Wazuh base vs custom" | Out-Null
    $r5 = Write-Grid $g 34 1 $chartRules (Get-PropNames $chartRules)
    Add-Chart $g $r5 430 620 520 250 "Reglas Wazuh 110xxx" | Out-Null
    $r6 = Write-Grid $g 50 1 $heatVisibility (Get-PropNames $heatVisibility)
    Apply-Heatmap $g 50 1 $heatVisibility.Count 9
    $r7 = Write-Grid $g 50 13 $heatCapacity (Get-PropNames $heatCapacity)
    Apply-Heatmap $g 50 13 $heatCapacity.Count 3
    $g.Range("A49").Value2 = "Heatmap visibilidad por TEC y fuente"
    $g.Range("M49").Value2 = "Heatmap VR vs Wazuh por capacidad"
    $g.Range("A49:M49").Font.Bold = $true
    $g.Columns.AutoFit() | Out-Null
    for ($c = 1; $c -le 18; $c++) { if ($g.Columns.Item($c).ColumnWidth -gt 24) { $g.Columns.Item($c).ColumnWidth = 24 } }

    $wb.Worksheets.Item("README").Activate()
    $wb.Save()
    $wb.Close($true)
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
}

Copy-Item -LiteralPath $FinalXlsx -Destination $V5Xlsx -Force
Copy-Item -LiteralPath $FinalXlsx -Destination $TraceFinalXlsx -Force
Copy-Item -LiteralPath $V5Xlsx -Destination $TraceV5Xlsx -Force
Copy-Item -LiteralPath $DqTrace -Destination $DqFinal -Force

$audit = @"
# Auditoria $FinalName

Fecha generacion: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss K")

## Entradas canonicas

- Excel base: 04_EVIDENCE\Analisis_Tecnicas_TFM_V.4.xlsx
- VR REAL: 05_LOGS\Validas\Prueba_24_06_2026
- VR FP: 05_LOGS\Validas\Prueba_24_06_2026_FP
- Wazuh: 10_WAZUH\TFM_WAZUH_VISIBILIDAD_EVIDENCE

## Criterio de clasificacion

- REAL: filas CLIENT_EVENT de carpeta real dentro de $($dq.windows.real_start_utc)..$($dq.windows.real_end_utc).
- FP: summary/vr_hits del RunId TFM_FP_20260626_101948 dentro de $($dq.windows.fp_start_utc)..$($dq.windows.fp_end_utc).
- Filas fuera de ventana o duplicadas en carpeta FP quedan en DISCREPANCIAS.
- JSONL/Discord no se consideran deteccion.

## Resultados globales

- TEC reales: 9/9 presentes; OK=$($dq.metrics.tecnicas).
- Alertas VR reales incluidas: $($dq.metrics.alertas_vr_real_incluidas).
- Alertas VR por perfil: P1=$($dq.metrics.alertas_vr_por_perfil.P1); P2=$($dq.metrics.alertas_vr_por_perfil.P2); P3=$($dq.metrics.alertas_vr_por_perfil.P3); P4=$($dq.metrics.alertas_vr_por_perfil.P4).
- FP: 10/10 presentes; hits FP=$($dq.metrics.fp_hits).
- Wazuh base: archives=$($dq.metrics.wazuh_base.archives), alerts=$($dq.metrics.wazuh_base.alerts).
- Wazuh custom: archives=$($dq.metrics.wazuh_custom.archives), alerts=$($dq.metrics.wazuh_custom.alerts), reglas 110xxx=56.
- 110201=0: gap TEC-009 conservado.

## Advertencias

- Discrepancias registradas: $($dq.metrics.discrepancias).
- Filas CLIENT_EVENT en ventana FP no reconciliadas con vr_hits: $($dq.metrics.fp_client_event_rows_revisar).
- TEC-009 tiene HTTP 200, UploadSucceeded=True, SHA-256 y bytes en ReceiverResponse, pero no se localiza receiver_log.jsonl ni ZIP recibido en las rutas revisadas.
- Se genero el XLSX con Excel COM local porque @oai/artifact-tool no estaba disponible en el entorno.

## Salidas

- $FinalXlsx
- $V5Xlsx
- $DqFinal
- $AuditFinal
"@

Set-Content -LiteralPath $AuditFinal -Value $audit -Encoding UTF8
Copy-Item -LiteralPath $AuditFinal -Destination $AuditTrace -Force

"Workbook generated: $FinalXlsx" | Tee-Object -FilePath $LogPath
"V5 copy: $V5Xlsx" | Tee-Object -FilePath $LogPath -Append
"Audit: $AuditFinal" | Tee-Object -FilePath $LogPath -Append
"Data quality: $DqFinal" | Tee-Object -FilePath $LogPath -Append
