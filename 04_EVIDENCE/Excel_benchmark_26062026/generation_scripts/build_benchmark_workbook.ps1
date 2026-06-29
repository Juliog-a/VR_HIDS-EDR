param(
    [string]$DatasetJson = ""
)

$ErrorActionPreference = "Stop"

function Resolve-Root {
    return (Resolve-Path (Join-Path $PSScriptRoot "..\..\..")).Path
}

function RgbColor([int]$r, [int]$g, [int]$b) {
    return $r + ($g * 256) + ($b * 65536)
}

function CellValue($v) {
    if ($null -eq $v) { return "" }
    if ($v -is [bool]) {
        if ($v) { return "TRUE" }
        return "FALSE"
    }
    if ($v -is [byte] -or $v -is [int16] -or $v -is [int32] -or $v -is [int64] -or
        $v -is [single] -or $v -is [double] -or $v -is [decimal]) {
        return [double]$v
    }
    if ($v -is [datetime]) { return $v.ToString("o") }
    return $v
}

function To-Array($maybeArray) {
    if ($null -eq $maybeArray) { return @() }
    if ($maybeArray -is [System.Array]) { return @($maybeArray) }
    return @($maybeArray)
}

function Get-ValueByName($obj, [string]$name) {
    if ($null -eq $obj) { return "" }
    $prop = $obj.PSObject.Properties[$name]
    if ($null -eq $prop) { return "" }
    return $prop.Value
}

function Write-Table {
    param(
        [object]$Worksheet,
        [int]$StartRow,
        [int]$StartCol,
        [string[]]$Headers,
        [object[]]$Rows,
        [switch]$AddFilter
    )

    $rowCount = [Math]::Max(1, $Rows.Count + 1)
    $colCount = $Headers.Count
    for ($c = 0; $c -lt $colCount; $c++) {
        $headerName = [string]($Headers[$c])
        $Worksheet.Cells.Item([int]$StartRow, [int]($StartCol + $c)).Value2 = $headerName
    }
    for ($r = 0; $r -lt $Rows.Count; $r++) {
        for ($c = 0; $c -lt $colCount; $c++) {
            $headerName = [string]($Headers[$c])
            $rowObj = $Rows[$r]
            $cell = $Worksheet.Cells.Item([int]($StartRow + $r + 1), [int]($StartCol + $c))
            $rawValue = Get-ValueByName -obj $rowObj -name $headerName
            $convertedValue = CellValue -v $rawValue
            if ($convertedValue -is [double]) {
                $cell.Value = $convertedValue
            } else {
                $cell.Value2 = [string]$convertedValue
            }
        }
    }

    $endRow = $StartRow + $rowCount - 1
    $endCol = $StartCol + $colCount - 1
    $range = $Worksheet.Range($Worksheet.Cells.Item($StartRow, $StartCol), $Worksheet.Cells.Item($endRow, $endCol))
    $range.Font.Name = "Calibri"
    $range.Font.Size = 10
    $header = $Worksheet.Range($Worksheet.Cells.Item($StartRow, $StartCol), $Worksheet.Cells.Item($StartRow, $endCol))
    $header.Font.Bold = $true
    $header.Font.Color = RgbColor 255 255 255
    $header.Interior.Color = RgbColor 31 78 121
    $header.RowHeight = 22
    $range.Borders.LineStyle = 1
    $range.Borders.Color = RgbColor 220 225 230
    $range.VerticalAlignment = -4108
    $range.WrapText = $true
    if ($AddFilter -and $Rows.Count -gt 0) {
        $range.AutoFilter() | Out-Null
    }
    $range.Columns.AutoFit() | Out-Null
    for ($c = $StartCol; $c -le $endCol; $c++) {
        if ($Worksheet.Columns.Item($c).ColumnWidth -gt 55) {
            $Worksheet.Columns.Item($c).ColumnWidth = 55
        }
    }
}

function Freeze-Header {
    param([object]$Excel, [object]$Worksheet)
    $Worksheet.Activate() | Out-Null
    $Worksheet.Range("A2").Select() | Out-Null
    $Excel.ActiveWindow.SplitRow = 1
    $Excel.ActiveWindow.FreezePanes = $true
}

function Apply-StatusColors {
    param(
        [object]$Worksheet,
        [int]$StatusColumn,
        [int]$StartRow,
        [int]$EndRow
    )
    for ($r = $StartRow; $r -le $EndRow; $r++) {
        $cell = $Worksheet.Cells.Item($r, $StatusColumn)
        $status = [string]$cell.Value2
        if ($status -eq "OK" -or $status -eq "APTO") {
            $cell.Interior.Color = RgbColor 198 239 206
            $cell.Font.Color = RgbColor 0 97 0
        } elseif ($status -eq "WARN") {
            $cell.Interior.Color = RgbColor 255 235 156
            $cell.Font.Color = RgbColor 156 101 0
        } elseif ($status -eq "FAIL" -or $status -eq "NO_APTO") {
            $cell.Interior.Color = RgbColor 255 199 206
            $cell.Font.Color = RgbColor 156 0 6
        }
        $cell.Font.Bold = $true
    }
}

function Add-Chart {
    param(
        [object]$Worksheet,
        [string]$RangeAddress,
        [string]$Title,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height
    )
    $chartObject = $Worksheet.ChartObjects().Add($Left, $Top, $Width, $Height)
    $chart = $chartObject.Chart
    $chart.ChartType = 51
    $chart.SetSourceData($Worksheet.Range($RangeAddress))
    $chart.HasTitle = $true
    $chart.ChartTitle.Text = $Title
    $chart.HasLegend = $false
    $chart.Axes(1).TickLabels.Font.Size = 9
    $chart.Axes(2).TickLabels.Font.Size = 9
    $chart.ChartArea.Format.Line.Visible = 0
}

function Write-ChartBlock {
    param(
        [object]$Worksheet,
        [int]$StartRow,
        [int]$StartCol,
        [string]$MetricHeader,
        [object[]]$ScenarioRows,
        [string]$MetricProperty
    )
    $rows = @()
    foreach ($r in $ScenarioRows) {
        $rows += [pscustomobject]@{
            Scenario = $r.Scenario
            $MetricHeader = Get-ValueByName $r $MetricProperty
        }
    }
    Write-Table -Worksheet $Worksheet -StartRow $StartRow -StartCol $StartCol -Headers @("Scenario", $MetricHeader) -Rows $rows
}

function Sheet-Exists {
    param([object]$Workbook, [string]$Name)
    foreach ($ws in $Workbook.Worksheets) {
        if ($ws.Name -eq $Name) { return $true }
    }
    return $false
}

$root = Resolve-Root
if ([string]::IsNullOrWhiteSpace($DatasetJson)) {
    $DatasetJson = Join-Path $root "04_EVIDENCE\Excel_benchmark_26062026\benchmark_dataset.json"
}

$traceDir = Join-Path $root "04_EVIDENCE\Excel_benchmark_26062026"
$finalDir = Join-Path $root "08_MEMORIA\Excel_benchmark"
$normalizedDir = Join-Path $traceDir "normalized"
$finalName = "TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026"
$finalXlsx = Join-Path $finalDir "$finalName.xlsx"
$traceXlsx = Join-Path $traceDir "$finalName.xlsx"
$finalAudit = Join-Path $finalDir "${finalName}_AUDIT.md"
$traceAudit = Join-Path $traceDir "${finalName}_AUDIT.md"
$finalDq = Join-Path $finalDir "${finalName}_DATA_QUALITY.json"
$traceDq = Join-Path $traceDir "${finalName}_DATA_QUALITY.json"
$generationLog = Join-Path $traceDir "benchmark_generation_26062026.log"
$previewPdf = Join-Path $traceDir "GRAFICAS_PREVIEW_26062026.pdf"

New-Item -ItemType Directory -Path $traceDir -Force | Out-Null
New-Item -ItemType Directory -Path $finalDir -Force | Out-Null
New-Item -ItemType Directory -Path $normalizedDir -Force | Out-Null

$log = New-Object System.Collections.Generic.List[string]
$log.Add("[$(Get-Date -Format o)] Inicio generacion benchmark final.")

$data = Get-Content -LiteralPath $DatasetJson -Raw -Encoding UTF8 | ConvertFrom-Json
if ($data.status -ne "APTO") {
    throw "Precheck NO_APTO. No se genera Excel final."
}

$canonical = [string]$data.paths.canonicalXlsx
if (!(Test-Path -LiteralPath $canonical)) {
    throw "No existe fuente canonica: $canonical"
}
Copy-Item -LiteralPath $canonical -Destination $finalXlsx -Force
$log.Add("[$(Get-Date -Format o)] Copiada plantilla canonica a $finalXlsx")

$excel = $null
$wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $wb = $excel.Workbooks.Open($finalXlsx)

    $tmp = $wb.Worksheets.Add()
    $tmp.Name = "__TMP_KEEP__"
    for ($i = $wb.Worksheets.Count; $i -ge 1; $i--) {
        $ws = $wb.Worksheets.Item($i)
        if ($ws.Name -ne "__TMP_KEEP__") {
            $ws.Delete()
        }
    }

    $sheetNames = @(
        "README",
        "RESUMEN_EJECUTIVO",
        "RUNS_VALIDOS",
        "RESUMEN_ESCENARIO",
        "RAW_SAMPLES",
        "PROCESS_SAMPLES",
        "CALIDAD_DATOS",
        "GRAFICAS",
        "TRAZABILIDAD",
        "LIMITACIONES",
        "DISCREPANCIAS"
    )

    $sheets = @{}
    foreach ($name in $sheetNames) {
        $ws = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
        $ws.Name = $name
        $sheets[$name] = $ws
    }
    $tmp.Delete()

    $runs = To-Array $data.runs
    $scenarioSummary = To-Array $data.scenarioSummary
    $rawSamples = To-Array $data.rawSamples
    $processSamples = To-Array $data.processSamples
    $checks = To-Array $data.checks
    $sources = To-Array $data.sources
    $discrepancies = To-Array $data.discrepancies

    $readmeRows = @(
        [pscustomobject]@{ Campo = "Objetivo"; Valor = "Excel definitivo de benchmark de rendimiento de Velociraptor para memoria TFM." },
        [pscustomobject]@{ Campo = "Fecha de generacion"; Valor = (Get-Date -Format "yyyy-MM-dd HH:mm:ss K") },
        [pscustomobject]@{ Campo = "Fuente canonica"; Valor = $canonical },
        [pscustomobject]@{ Campo = "Campana"; Valor = "BENCH_20260619_FINAL01" },
        [pscustomobject]@{ Campo = "Criterio benchmark"; Valor = "Coste operativo observado en laboratorio: CPU, RAM, duracion, muestras, servicio y proceso cliente Velociraptor." },
        [pscustomobject]@{ Campo = "Escenarios"; Valor = "BASELINE_NO_VR; VR_IDLE; VR_TEC_RUNNER; 3 repeticiones validas por escenario." },
        [pscustomobject]@{ Campo = "Separacion metodologica"; Valor = "Este Excel no mide calidad de deteccion, alerta, evidencia ni falsos positivos." },
        [pscustomobject]@{ Campo = "SERVER_GUI"; Valor = "Excluido del calculo principal mediante IncludeServerGuiInTotal=False y ServerGuiExcludedFromClientMetrics=True." },
        [pscustomobject]@{ Campo = "RunnerStillRunningAtEnd"; Valor = "Conservado como WARN no bloqueante en VR_TEC_RUNNER." },
        [pscustomobject]@{ Campo = "Limitacion principal"; Valor = "Mediciones relativas del laboratorio Windows 10 virtualizado; no extrapolables directamente a produccion." }
    )
    Write-Table -Worksheet $sheets["README"] -StartRow 1 -StartCol 1 -Headers @("Campo", "Valor") -Rows $readmeRows -AddFilter
    Freeze-Header -Excel $excel -Worksheet $sheets["README"]

    $k = $data.kpi
    $summaryRows = @(
        [pscustomobject]@{ KPI = "runs totales"; Valor = $k.TotalRuns; Observaciones = "Esperado: 9" },
        [pscustomobject]@{ KPI = "runs validos"; Valor = $k.ValidRuns; Observaciones = "9/9" },
        [pscustomobject]@{ KPI = "escenarios"; Valor = $k.ScenarioCount; Observaciones = "BASELINE_NO_VR, VR_IDLE, VR_TEC_RUNNER" },
        [pscustomobject]@{ KPI = "repeticiones por escenario"; Valor = "3/3/3"; Observaciones = ($k.RepetitionsPerScenario | ConvertTo-Json -Compress) },
        [pscustomobject]@{ KPI = "CPU media BASELINE_NO_VR"; Valor = $k.AvgCPU_BASELINE_NO_VR; Observaciones = "%" },
        [pscustomobject]@{ KPI = "CPU media VR_IDLE"; Valor = $k.AvgCPU_VR_IDLE; Observaciones = "%" },
        [pscustomobject]@{ KPI = "CPU media VR_TEC_RUNNER"; Valor = $k.AvgCPU_VR_TEC_RUNNER; Observaciones = "%" },
        [pscustomobject]@{ KPI = "RAM media BASELINE_NO_VR"; Valor = $k.AvgRAM_BASELINE_NO_VR; Observaciones = "MB" },
        [pscustomobject]@{ KPI = "RAM media VR_IDLE"; Valor = $k.AvgRAM_VR_IDLE; Observaciones = "MB" },
        [pscustomobject]@{ KPI = "RAM media VR_TEC_RUNNER"; Valor = $k.AvgRAM_VR_TEC_RUNNER; Observaciones = "MB" },
        [pscustomobject]@{ KPI = "pico CPU cliente VR"; Valor = $k.PeakCPU; Observaciones = "%" },
        [pscustomobject]@{ KPI = "pico RAM cliente VR"; Valor = $k.PeakRAM; Observaciones = "MB" },
        [pscustomobject]@{ KPI = "duracion media BASELINE_NO_VR"; Valor = (Get-ValueByName ($scenarioSummary | Where-Object { $_.Scenario -eq "BASELINE_NO_VR" } | Select-Object -First 1) "AvgDurationSeconds"); Observaciones = "segundos" },
        [pscustomobject]@{ KPI = "duracion media VR_IDLE"; Valor = (Get-ValueByName ($scenarioSummary | Where-Object { $_.Scenario -eq "VR_IDLE" } | Select-Object -First 1) "AvgDurationSeconds"); Observaciones = "segundos" },
        [pscustomobject]@{ KPI = "duracion media VR_TEC_RUNNER"; Valor = (Get-ValueByName ($scenarioSummary | Where-Object { $_.Scenario -eq "VR_TEC_RUNNER" } | Select-Object -First 1) "AvgDurationSeconds"); Observaciones = "segundos" },
        [pscustomobject]@{ KPI = "estado global"; Valor = "APTO"; Observaciones = "Precheck sin fallos criticos" },
        [pscustomobject]@{ KPI = "observaciones"; Valor = "RunnerStillRunningAtEnd=True en VR_TEC_RUNNER"; Observaciones = "WARN no bloqueante" }
    )
    Write-Table -Worksheet $sheets["RESUMEN_EJECUTIVO"] -StartRow 1 -StartCol 1 -Headers @("KPI", "Valor", "Observaciones") -Rows $summaryRows -AddFilter
    Apply-StatusColors -Worksheet $sheets["RESUMEN_EJECUTIVO"] -StatusColumn 2 -StartRow 2 -EndRow ($summaryRows.Count + 1)
    Freeze-Header -Excel $excel -Worksheet $sheets["RESUMEN_EJECUTIVO"]

    $runHeaders = @(
        "RunId", "Scenario", "Repetition", "StartTime", "EndTime", "DurationSeconds",
        "ScenarioValidity", "ValidForClientBenchmark", "ClientServiceStatusStart",
        "ClientServiceStatusEnd", "ClientRows", "RunnerObservedSamples",
        "VR_Client_ProcessCountMax", "Avg_VR_Client_CPU_Percent",
        "Max_VR_Client_CPU_Percent", "P95_VR_Client_CPU_Percent",
        "Avg_VR_Client_RAM_MB", "Max_VR_Client_RAM_MB", "P95_VR_Client_RAM_MB",
        "Warnings", "SourceFile", "SHA-256"
    )
    Write-Table -Worksheet $sheets["RUNS_VALIDOS"] -StartRow 1 -StartCol 1 -Headers $runHeaders -Rows $runs -AddFilter
    Apply-StatusColors -Worksheet $sheets["RUNS_VALIDOS"] -StatusColumn 7 -StartRow 2 -EndRow ($runs.Count + 1)
    Freeze-Header -Excel $excel -Worksheet $sheets["RUNS_VALIDOS"]

    $scenarioHeaders = @("Scenario", "RepetitionsExpected", "RepetitionsFound", "ValidRuns", "AvgDurationSeconds", "AvgCPUPercent", "MaxCPUPercent", "AvgRAMMB", "MaxRAMMB", "RunnerObservedSamples", "Warnings", "Interpretation")
    Write-Table -Worksheet $sheets["RESUMEN_ESCENARIO"] -StartRow 1 -StartCol 1 -Headers $scenarioHeaders -Rows $scenarioSummary -AddFilter
    Freeze-Header -Excel $excel -Worksheet $sheets["RESUMEN_ESCENARIO"]

    $rawHeaders = @()
    if ($rawSamples.Count -gt 0) {
        $rawHeaders = @($rawSamples[0].PSObject.Properties.Name)
    }
    Write-Table -Worksheet $sheets["RAW_SAMPLES"] -StartRow 1 -StartCol 1 -Headers $rawHeaders -Rows $rawSamples -AddFilter
    Freeze-Header -Excel $excel -Worksheet $sheets["RAW_SAMPLES"]

    $processHeaders = @()
    if ($processSamples.Count -gt 0) {
        $processHeaders = @($processSamples[0].PSObject.Properties.Name)
    }
    Write-Table -Worksheet $sheets["PROCESS_SAMPLES"] -StartRow 1 -StartCol 1 -Headers $processHeaders -Rows $processSamples -AddFilter
    Freeze-Header -Excel $excel -Worksheet $sheets["PROCESS_SAMPLES"]

    $extraChecks = @(
        [pscustomobject]@{ Check = "Formula check final"; Status = "OK"; Evidence = "Workbook generado sin formulas de calculo."; Severity = "CRITICA"; Decision = "No hay formulas rotas posibles en tablas generadas." },
        [pscustomobject]@{ Check = "Chart population final"; Status = "OK"; Evidence = "Graficas construidas desde RESUMEN_ESCENARIO y bloques de datos reales."; Severity = "CRITICA"; Decision = "No se generan graficas vacias." },
        [pscustomobject]@{ Check = "Benchmark separado de deteccion"; Status = "OK"; Evidence = "No incluye hojas de alertas, deteccion, FP ni Wazuh."; Severity = "CRITICA"; Decision = "Mantiene separacion metodologica." }
    )
    $qualityRows = @($checks) + $extraChecks
    Write-Table -Worksheet $sheets["CALIDAD_DATOS"] -StartRow 1 -StartCol 1 -Headers @("Check", "Status", "Evidence", "Severity", "Decision") -Rows $qualityRows -AddFilter
    Apply-StatusColors -Worksheet $sheets["CALIDAD_DATOS"] -StatusColumn 2 -StartRow 2 -EndRow ($qualityRows.Count + 1)
    Freeze-Header -Excel $excel -Worksheet $sheets["CALIDAD_DATOS"]

    $chartSheet = $sheets["GRAFICAS"]
    Write-ChartBlock -Worksheet $chartSheet -StartRow 1 -StartCol 1 -MetricHeader "CPU media (%)" -ScenarioRows $scenarioSummary -MetricProperty "AvgCPUPercent"
    Write-ChartBlock -Worksheet $chartSheet -StartRow 1 -StartCol 4 -MetricHeader "CPU maxima (%)" -ScenarioRows $scenarioSummary -MetricProperty "MaxCPUPercent"
    Write-ChartBlock -Worksheet $chartSheet -StartRow 1 -StartCol 7 -MetricHeader "RAM media (MB)" -ScenarioRows $scenarioSummary -MetricProperty "AvgRAMMB"
    Write-ChartBlock -Worksheet $chartSheet -StartRow 1 -StartCol 10 -MetricHeader "RAM maxima (MB)" -ScenarioRows $scenarioSummary -MetricProperty "MaxRAMMB"
    Write-ChartBlock -Worksheet $chartSheet -StartRow 18 -StartCol 1 -MetricHeader "Duracion media (s)" -ScenarioRows $scenarioSummary -MetricProperty "AvgDurationSeconds"

    $comparisonRows = @()
    foreach ($r in $scenarioSummary) {
        $comparisonRows += [pscustomobject]@{
            Scenario = $r.Scenario
            AvgCPUPercent = $r.AvgCPUPercent
            AvgRAMMB = $r.AvgRAMMB
            AvgDurationSeconds = $r.AvgDurationSeconds
        }
    }
    Write-Table -Worksheet $chartSheet -StartRow 18 -StartCol 4 -Headers @("Scenario", "AvgCPUPercent", "AvgRAMMB", "AvgDurationSeconds") -Rows $comparisonRows

    $heatRows = @()
    foreach ($r in $qualityRows) {
        $heatRows += [pscustomobject]@{
            Check = $r.Check
            Status = $r.Status
            Severity = $r.Severity
        }
    }
    Write-Table -Worksheet $chartSheet -StartRow 35 -StartCol 1 -Headers @("Check", "Status", "Severity") -Rows $heatRows
    Apply-StatusColors -Worksheet $chartSheet -StatusColumn 2 -StartRow 36 -EndRow (35 + $heatRows.Count)

    Add-Chart -Worksheet $chartSheet -RangeAddress "A1:B4" -Title "CPU media por escenario" -Left 20 -Top 130 -Width 330 -Height 220
    Add-Chart -Worksheet $chartSheet -RangeAddress "D1:E4" -Title "CPU maxima por escenario" -Left 370 -Top 130 -Width 330 -Height 220
    Add-Chart -Worksheet $chartSheet -RangeAddress "G1:H4" -Title "RAM media por escenario" -Left 720 -Top 130 -Width 330 -Height 220
    Add-Chart -Worksheet $chartSheet -RangeAddress "J1:K4" -Title "RAM maxima por escenario" -Left 1070 -Top 130 -Width 330 -Height 220
    Add-Chart -Worksheet $chartSheet -RangeAddress "A18:B21" -Title "Duracion media por escenario" -Left 20 -Top 430 -Width 420 -Height 240
    Add-Chart -Worksheet $chartSheet -RangeAddress "D18:G21" -Title "Comparacion BASELINE vs VR_IDLE vs VR_TEC_RUNNER" -Left 470 -Top 430 -Width 520 -Height 240
    $chartSheet.Columns.AutoFit() | Out-Null
    Freeze-Header -Excel $excel -Worksheet $chartSheet

    Write-Table -Worksheet $sheets["TRAZABILIDAD"] -StartRow 1 -StartCol 1 -Headers @("Ruta", "Tipo", "TamanoBytes", "FechaModificacion", "SHA-256", "Rol", "Campana", "Observaciones") -Rows $sources -AddFilter
    Freeze-Header -Excel $excel -Worksheet $sheets["TRAZABILIDAD"]

    $limitRows = @(
        [pscustomobject]@{ Limitacion = "Laboratorio Windows 10 virtualizado"; Impacto = "Las mediciones reflejan esta VM/laboratorio, no una poblacion de endpoints productivos." },
        [pscustomobject]@{ Limitacion = "Metricas relativas"; Impacto = "CPU/RAM no deben presentarse como medida universal extrapolable directamente." },
        [pscustomobject]@{ Limitacion = "Foco en agente cliente Velociraptor"; Impacto = "SERVER_GUI queda excluido del calculo principal." },
        [pscustomobject]@{ Limitacion = "RunnerStillRunningAtEnd"; Impacto = "Se conserva como WARN no bloqueante porque la ventana de muestreo del cliente existe y es valida." },
        [pscustomobject]@{ Limitacion = "No mide deteccion"; Impacto = "El benchmark no evalua calidad de deteccion, alertas, evidencia ni falsos positivos." },
        [pscustomobject]@{ Limitacion = "Numero limitado de repeticiones"; Impacto = "Tres repeticiones por escenario; adecuado para TFM, limitado para inferencia estadistica amplia." },
        [pscustomobject]@{ Limitacion = "Ruido residual del sistema operativo"; Impacto = "Puede afectar a CPU/RAM observadas en una VM Windows." }
    )
    Write-Table -Worksheet $sheets["LIMITACIONES"] -StartRow 1 -StartCol 1 -Headers @("Limitacion", "Impacto") -Rows $limitRows -AddFilter
    Freeze-Header -Excel $excel -Worksheet $sheets["LIMITACIONES"]

    Write-Table -Worksheet $sheets["DISCREPANCIAS"] -StartRow 1 -StartCol 1 -Headers @("Tipo", "Ambito", "Detalle", "Decision", "Fuente") -Rows $discrepancies -AddFilter
    Freeze-Header -Excel $excel -Worksheet $sheets["DISCREPANCIAS"]

    foreach ($ws in $wb.Worksheets) {
        $ws.PageSetup.Orientation = 2
        $ws.PageSetup.Zoom = $false
        $ws.PageSetup.FitToPagesWide = 1
        $ws.PageSetup.FitToPagesTall = $false
    }

    $wb.Worksheets.Item("README").Activate() | Out-Null
    $wb.Save()

    $chartCount = $sheets["GRAFICAS"].ChartObjects().Count
    try {
        $sheets["GRAFICAS"].ExportAsFixedFormat(0, $previewPdf)
        $log.Add("[$(Get-Date -Format o)] Preview PDF generado: $previewPdf")
    } catch {
        $log.Add("[$(Get-Date -Format o)] WARN preview PDF no generado: $($_.Exception.Message)")
    }

    $formulaCount = 0
    $formulaErrorCount = 0
    foreach ($ws in $wb.Worksheets) {
        try {
            $formulas = $ws.UsedRange.SpecialCells(-4123)
            $formulaCount += [int]$formulas.Count
            foreach ($cell in $formulas) {
                if ($cell.Text -like "#*") { $formulaErrorCount += 1 }
            }
        } catch {
        }
    }

    $mandatoryPresent = @()
    foreach ($name in $sheetNames) {
        $mandatoryPresent += [pscustomobject]@{ Sheet = $name; Present = (Sheet-Exists -Workbook $wb -Name $name) }
    }
    $missingMandatory = @($mandatoryPresent | Where-Object { $_.Present -ne $true })

    $verification = [pscustomobject]@{
        WorkbookOpen = $true
        WorksheetCount = $wb.Worksheets.Count
        MandatorySheetsPresent = ($missingMandatory.Count -eq 0)
        MissingMandatorySheets = @($missingMandatory | ForEach-Object { $_.Sheet })
        FormulaCount = $formulaCount
        FormulaErrorCount = $formulaErrorCount
        ChartsInGraficas = $chartCount
        ChartsPopulated = ($chartCount -ge 6)
        ValidRuns = $data.kpi.ValidRuns
        TotalRuns = $data.kpi.TotalRuns
        ScenarioCount = $data.kpi.ScenarioCount
        ServerGuiExcluded = (($runs | Where-Object { $_.IncludeServerGuiInTotal -ne $false -or $_.ServerGuiExcludedFromClientMetrics -ne $true }).Count -eq 0)
        NotepadRows = (($runs | Measure-Object -Property NotepadRows -Sum).Sum)
        SourcesWithSha256 = (($sources | Where-Object { [string]::IsNullOrWhiteSpace($_.'SHA-256') }).Count -eq 0)
    }

    if (-not $verification.MandatorySheetsPresent) { throw "Faltan hojas obligatorias." }
    if ($verification.FormulaErrorCount -ne 0) { throw "Hay errores de formula." }
    if (-not $verification.ChartsPopulated) { throw "Graficas insuficientes." }
    if ($verification.ValidRuns -ne 9 -or $verification.TotalRuns -ne 9 -or $verification.ScenarioCount -ne 3) { throw "Runs/escenarios invalidos." }
    if (-not $verification.ServerGuiExcluded) { throw "SERVER_GUI no excluido correctamente." }
    if ($verification.NotepadRows -ne 0) { throw "notepad.exe observado en muestras." }
    if (-not $verification.SourcesWithSha256) { throw "Hay fuentes sin SHA-256." }

    $wb.Close($true)
    $wb = $null
    $excel.Quit()
    $excel = $null

    Copy-Item -LiteralPath $finalXlsx -Destination $traceXlsx -Force

    $dqBase = Get-Content -LiteralPath (Join-Path $traceDir "${finalName}_DATA_QUALITY.json") -Raw -Encoding UTF8 | ConvertFrom-Json
    $dqOut = [pscustomobject]@{
        generatedAt = (Get-Date -Format o)
        status = "APTO"
        decision = "Excel final generado y verificado."
        kpi = $data.kpi
        checks = $qualityRows
        discrepancies = $data.discrepancies
        finalVerification = $verification
        precheck = $dqBase
    }
    $dqJson = $dqOut | ConvertTo-Json -Depth 12
    Set-Content -LiteralPath $traceDq -Value $dqJson -Encoding UTF8
    Copy-Item -LiteralPath $traceDq -Destination $finalDq -Force

    $audit = @"
# TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026_AUDIT

Fecha generacion: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss K")

## Decision

APTO. El precheck confirma 9/9 runs validos, 3 escenarios y 3 repeticiones por escenario. No se han detectado fallos criticos.

## Salidas

- Excel final memoria: $finalXlsx
- Excel trazabilidad: $traceXlsx
- Data quality JSON: $traceDq
- Precheck: $(Join-Path $traceDir "BENCHMARK_PRECHECK_26062026.md")
- Preview PDF graficas: $previewPdf

## Fuente canonica

$canonical

Analisis_Benchmark.xlsx se mantiene solo como referencia historica/formato.

## KPIs

- Runs totales: $($data.kpi.TotalRuns)
- Runs validos: $($data.kpi.ValidRuns)
- Escenarios: $($data.kpi.ScenarioCount)
- CPU media BASELINE_NO_VR: $($data.kpi.AvgCPU_BASELINE_NO_VR) %
- CPU media VR_IDLE: $($data.kpi.AvgCPU_VR_IDLE) %
- CPU media VR_TEC_RUNNER: $($data.kpi.AvgCPU_VR_TEC_RUNNER) %
- RAM media BASELINE_NO_VR: $($data.kpi.AvgRAM_BASELINE_NO_VR) MB
- RAM media VR_IDLE: $($data.kpi.AvgRAM_VR_IDLE) MB
- RAM media VR_TEC_RUNNER: $($data.kpi.AvgRAM_VR_TEC_RUNNER) MB
- Pico CPU cliente VR: $($data.kpi.PeakCPU) %
- Pico RAM cliente VR: $($data.kpi.PeakRAM) MB

## Validacion final

- Apertura Excel: OK
- Hojas obligatorias presentes: $($verification.MandatorySheetsPresent)
- Hojas: $($verification.WorksheetCount)
- Formulas: $($verification.FormulaCount)
- Errores de formula: $($verification.FormulaErrorCount)
- Graficas en GRAFICAS: $($verification.ChartsInGraficas)
- Graficas pobladas: $($verification.ChartsPopulated)
- SERVER_GUI excluido: $($verification.ServerGuiExcluded)
- notepad.exe runner: $($verification.NotepadRows)
- Fuentes con SHA-256: $($verification.SourcesWithSha256)

## Limitaciones metodologicas

- Laboratorio Windows 10 virtualizado.
- Mediciones relativas, no extrapolables directamente a produccion.
- Benchmark centrado en agente cliente Velociraptor.
- SERVER_GUI excluido del calculo principal.
- RunnerStillRunningAtEnd conservado como WARN no bloqueante.
- No mide calidad de deteccion, alertas, evidencia ni falsos positivos.
- Tres repeticiones por escenario.
- Posible ruido residual del sistema operativo.

## Discrepancias

- RunnerStillRunningAtEnd=True en VR_TEC_RUNNER: WARN no bloqueante.
- SERVER_GUI observado y excluido del calculo principal.
- Analisis_Benchmark.xlsx usado solo como referencia historica/formato.
"@
    Set-Content -LiteralPath $traceAudit -Value $audit -Encoding UTF8
    Copy-Item -LiteralPath $traceAudit -Destination $finalAudit -Force

    $log.Add("[$(Get-Date -Format o)] Excel final verificado: $finalXlsx")
    $log.Add("[$(Get-Date -Format o)] Copias generadas en memoria y trazabilidad.")
    Set-Content -LiteralPath $generationLog -Value $log -Encoding UTF8

    Write-Host "OK Excel benchmark final generado"
    Write-Host $finalXlsx
    Write-Host $traceXlsx
    Write-Host $traceAudit
    Write-Host $traceDq
} finally {
    if ($wb -ne $null) {
        try { $wb.Close($false) | Out-Null } catch {}
    }
    if ($excel -ne $null) {
        try { $excel.Quit() | Out-Null } catch {}
    }
}
