param(
    [Parameter(Mandatory = $true)]
    [string]$DetectionPath,
    [Parameter(Mandatory = $true)]
    [string]$BenchmarkPath,
    [Parameter(Mandatory = $true)]
    [string]$LogPath,
    [Parameter(Mandatory = $true)]
    [string]$ProgressPath
)

$ErrorActionPreference = 'Stop'
$xlCalculationAutomatic = -4105
$xlCellTypeLastCell = 11
$script:corrections = 0
$script:changes = [System.Collections.Generic.List[object]]::new()

function Set-Progress {
    param([string]$Text)
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`t$Text" |
        Set-Content -LiteralPath $ProgressPath -Encoding UTF8
}

function Release-ComObject {
    param([object]$Object)
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Add-Change {
    param([string]$Workbook, [string]$Location, [string]$Description)
    $script:corrections++
    [void]$script:changes.Add([pscustomobject]@{
        Workbook = $Workbook
        Location = $Location
        Description = $Description
    })
}

function Get-Sheet {
    param([object]$Workbook, [string]$Name)
    return $Workbook.Worksheets.Item($Name)
}

function Set-Cell {
    param(
        [object]$Workbook,
        [string]$SheetName,
        [string]$Address,
        [object]$Value,
        [string]$Description
    )
    $sheet = Get-Sheet $Workbook $SheetName
    $cell = $null
    try {
        $cell = $sheet.Range($Address)
        $old = [string]$cell.Value2
        if ($old -ne [string]$Value) {
            $cell.Value2 = $Value
            $cell.WrapText = $true
            Add-Change ([System.IO.Path]::GetFileName($Workbook.FullName)) "$SheetName!$Address" $Description
        }
    }
    finally {
        Release-ComObject $cell
        Release-ComObject $sheet
    }
}

function Verify-Formula {
    param(
        [object]$Workbook,
        [string]$SheetName,
        [string]$Address,
        [string]$ExpectedFormula,
        [string]$ExpectedText
    )
    $sheet = Get-Sheet $Workbook $SheetName
    $cell = $null
    try {
        $cell = $sheet.Range($Address)
        if ([string]$cell.Formula -ne $ExpectedFormula) {
            throw "Fórmula inesperada en ${SheetName}!${Address}: $($cell.Formula)"
        }
        if ([string]$cell.Text -ne $ExpectedText) {
            throw "Resultado inesperado en ${SheetName}!${Address}: $($cell.Text)"
        }
        return [pscustomobject]@{
            Sheet = $SheetName
            Address = $Address
            Formula = [string]$cell.Formula
            Value = [string]$cell.Text
        }
    }
    finally {
        Release-ComObject $cell
        Release-ComObject $sheet
    }
}

function Audit-Workbook {
    param([object]$Workbook)
    $errors = [System.Collections.Generic.List[object]]::new()
    $charts = [System.Collections.Generic.List[object]]::new()
    $sheets = [System.Collections.Generic.List[object]]::new()

    for ($i = 1; $i -le $Workbook.Worksheets.Count; $i++) {
        $sheet = $Workbook.Worksheets.Item($i)
        $used = $null
        try {
            $used = $sheet.UsedRange
            $rows = [int]$used.Rows.Count
            $cols = [int]$used.Columns.Count
            $values = $used.Value2
            if ($rows -eq 1 -and $cols -eq 1) {
                $values = @($values)
            }
            for ($r = 1; $r -le $rows; $r++) {
                for ($c = 1; $c -le $cols; $c++) {
                    try {
                        $value = if ($rows -eq 1 -and $cols -eq 1) { $values[0] } else { $values[$r, $c] }
                        if ([string]$value -match '^(#REF!|#N/A|#VALUE!|#DIV/0!|#NAME\?)$|^/9$') {
                            [void]$errors.Add([pscustomobject]@{
                                Sheet = $sheet.Name
                                Row = $r
                                Column = $c
                                Value = [string]$value
                            })
                        }
                    }
                    catch {}
                }
            }

            $chartCount = [int]$sheet.ChartObjects().Count
            for ($ci = 1; $ci -le $chartCount; $ci++) {
                $chartObject = $sheet.ChartObjects($ci)
                $chart = $null
                try {
                    $chart = $chartObject.Chart
                    $seriesCount = 0
                    try { $seriesCount = [int]$chart.SeriesCollection().Count } catch {}
                    [void]$charts.Add([pscustomobject]@{
                        Sheet = $sheet.Name
                        Name = $chartObject.Name
                        SeriesCount = $seriesCount
                        HasData = ($seriesCount -gt 0)
                    })
                }
                finally {
                    Release-ComObject $chart
                    Release-ComObject $chartObject
                }
            }
            [void]$sheets.Add([pscustomobject]@{
                Name = $sheet.Name
                Rows = $rows
                Columns = $cols
                PrintArea = [string]$sheet.PageSetup.PrintArea
                AutoFilter = [bool]$sheet.AutoFilterMode
                Charts = $chartCount
            })
        }
        finally {
            Release-ComObject $used
            Release-ComObject $sheet
        }
    }

    return [pscustomobject]@{
        Path = $Workbook.FullName
        Sheets = @($sheets)
        FormulaErrors = @($errors)
        Charts = @($charts)
        EmptyCharts = @($charts | Where-Object { -not $_.HasData })
    }
}

$excel = $null
$detection = $null
$benchmark = $null
try {
    Set-Progress 'Inicio de Excel COM'
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    Set-Progress 'Corrección del Excel de detección'
    $detection = $excel.Workbooks.Open($DetectionPath, 0, $false)
    $excel.Calculation = $xlCalculationAutomatic

    $note82 = '82 filas no reconciliadas fueron excluidas y no se computaron como falsos positivos definitivos'
    Set-Cell $detection '01_Dashboard' 'C5' '379 filas CLIENT_EVENT emitidas/exportadas, incluidas emisiones repetidas; P1=19, P2=118, P3=109 y P4=133; no equivalen a alertas únicas.' 'Unidad correcta de las 379 filas'
    Set-Cell $detection '01_Dashboard' 'C7' "Runner FP custom; no incluye Hayabusa Monitoring. $note82." 'Advertencia visible de las 82 filas junto al KPI'
    Set-Cell $detection '01_Dashboard' 'E3' 'Tácticas ATT&CK representadas por el subconjunto evaluado' 'Denominación acotada de cobertura táctica'
    Set-Cell $detection '01_Dashboard' 'E8' 'Los cocientes /14 describen tácticas ATT&CK representadas en la clasificación del subconjunto evaluado; no equivalen a las 9 técnicas TEC ejecutadas ni a cobertura general de ATT&CK Enterprise.' 'Alcance de la cobertura táctica'

    Set-Cell $detection '06_Resumen' 'C8' '379 filas CLIENT_EVENT emitidas/exportadas, incluidas emisiones repetidas; el volumen no equivale a alertas únicas ni a cobertura.' 'Unidad correcta de las 379 filas'
    Set-Cell $detection '06_Resumen' 'A9' 'Distribución de filas CLIENT_EVENT' 'Denominación del recuento por perfil'
    Set-Cell $detection '06_Resumen' 'C9' 'Suma=379 filas emitidas/exportadas, incluidas emisiones repetidas; recuentos obtenidos de ALERTAS_VR.' 'Aclaración sobre repeticiones'
    Set-Cell $detection '06_Resumen' 'C10' "No se mezcla con las cuatro coincidencias High de Hayabusa Monitoring. $note82." 'Advertencia visible de las 82 filas junto al KPI'
    Set-Cell $detection '06_Resumen' 'A15' 'Tácticas ATT&CK representadas por el subconjunto evaluado' 'Denominación acotada de cobertura táctica'
    Set-Cell $detection '06_Resumen' 'C15' 'Los cocientes /14 corresponden a la clasificación táctica del subconjunto evaluado; no representan cobertura general de ATT&CK Enterprise.' 'Alcance de la cobertura táctica'

    Set-Cell $detection 'RESUMEN_EJECUTIVO' 'A15' 'Filas CLIENT_EVENT emitidas/exportadas incluidas' 'Unidad correcta del KPI de volumen'
    Set-Cell $detection 'RESUMEN_EJECUTIVO' 'C15' '379 filas CLIENT_EVENT emitidas/exportadas, incluidas emisiones repetidas; no son 379 alertas únicas.' 'Aclaración del KPI de volumen'
    foreach ($row in 16..19) {
        $sheet = Get-Sheet $detection 'RESUMEN_EJECUTIVO'
        $cell = $null
        try {
            $cell = $sheet.Range("A$row")
            $new = ([string]$cell.Value2) -replace '(?i)alertas', 'filas'
            if ([string]$cell.Value2 -ne $new) {
                $cell.Value2 = $new
                Add-Change ([System.IO.Path]::GetFileName($detection.FullName)) "RESUMEN_EJECUTIVO!A$row" 'Perfil expresado en filas CLIENT_EVENT'
            }
        }
        finally {
            Release-ComObject $cell
            Release-ComObject $sheet
        }
    }
    Set-Cell $detection 'RESUMEN_EJECUTIVO' 'C20' "summary.json y vr_hits.csv reconciliados. $note82." 'Advertencia visible de las 82 filas junto al KPI'

    Set-Cell $detection 'FP_HITS' 'A2' "TotalHits=0. $note82." 'Advertencia visible de exclusión en la fuente FP'

    Set-Cell $detection '09_Resultados_Custom' 'A2' 'Resultados finales por perfil. El recuento corresponde a filas CLIENT_EVENT emitidas/exportadas; incluye emisiones repetidas y no equivale a alertas únicas. P2 contiene detección, contexto y evidencia forense y debe interpretarse por source.' 'Nota metodológica de volumen y EvidenceRole'
    Set-Cell $detection '09_Resultados_Custom' 'F4' 'Filas CLIENT_EVENT' 'Unidad de recuento'
    Set-Cell $detection '09_Resultados_Custom' 'F12' 'Filas CLIENT_EVENT' 'Unidad de recuento'
    Set-Cell $detection '09_Resultados_Custom' 'G12' 'Perfil de mayor prioridad observado' 'Corrección de Perfil máximo'
    $priorities = @('P2','P2','P2','P1','P1','P2','P1','P1','P1')
    for ($i = 0; $i -lt $priorities.Count; $i++) {
        Set-Cell $detection '09_Resultados_Custom' ("G" + (13 + $i)) $priorities[$i] 'Prioridad máxima observada por técnica'
    }
    Set-Cell $detection '09_Resultados_Custom' 'I6' 'Incluye señales detectoras, contexto y evidencia forense. No interpretar todas las filas P2 como alertas homogéneas; revisar la source.' 'Separación de EvidenceRole para P2'

    Set-Cell $detection 'ALERTAS_VR' 'A1' 'Filas CLIENT_EVENT emitidas/exportadas · campaña final' 'Denominación correcta del dataset'
    Set-Cell $detection 'ALERTAS_VR' 'A2' '379 filas CLIENT_EVENT emitidas/exportadas, incluidas emisiones repetidas. TEC_detector/source describe la clasificación inferida por detector/source; TEC_escenario pertenece al ground truth de ejecución y no se reconstruye fila a fila sin evidencia.' 'Diferencia TEC_escenario y TEC_detector/source'
    Set-Cell $detection 'ALERTAS_VR' 'C4' 'TEC_detector/source' 'Nombre explícito del campo TEC derivado'

    Set-Cell $detection '12_Plan_Alertas' 'A1' 'HISTÓRICO — Plan de alertas; no representa el estado final' 'Marcado de hoja histórica'
    Set-Cell $detection '12_Plan_Alertas' 'A2' 'Hoja conservada exclusivamente para trazabilidad histórica. El estado final se documenta en 01_Dashboard, 06_Resumen, RESUMEN_EJECUTIVO, 09_Resultados_Custom, ALERTAS_VR y COMPARACION_VR_WAZUH.' 'Alcance histórico del plan'
    Set-Cell $detection '12_Plan_Alertas' 'G8' 'Registro histórico; el estado SOC final debe consultarse en las hojas de resultados.' 'Neutralización del estado provisional'

    Set-Cell $detection 'COMPARACION_VR_WAZUH' 'C4' 'Filas/alertas según sistema' 'Unidades no equivalentes en la comparación'
    Set-Cell $detection 'COMPARACION_VR_WAZUH' 'C5' '379 filas CLIENT_EVENT emitidas/exportadas, incluidas emisiones repetidas' 'Unidad correcta de Velociraptor'

    $discrepancies = @(
        @('CLIENT_EVENT_REPETICIONES_EXACTAS','ALERTAS_VR','182','La auditoría documental identificó 182 filas adicionales en grupos de repetición exacta, 177 grupos repetidos y 197 registros distintos al comparar los 15 campos.','No afirmar 379 alertas únicas; conservar las filas y documentar las repeticiones.'),
        @('TEC_ESCENARIO_VS_DETECTOR_SOURCE','ALERTAS_VR','206/379','La clasificación TEC derivada del detector/source no coincide necesariamente con el escenario que originó la actividad.','Conservar TEC_detector/source y no fabricar TEC_escenario fila a fila.'),
        @('EVIDENCE_ROLE_NO_HOMOGENEO','09_Resultados_Custom','P2','P2 contiene detección, contexto y evidencia forense; las filas no comparten un único rol semántico.','Interpretar por source y separar detección, contexto y evidencia en la memoria.'),
        @('POSIBLE_TELEMETRIA_VALIDACION','ALERTAS_VR!I383','1','La fila contiene código de validación con -match y puede corresponder a telemetría generada por el propio proceso de validación.','Conservarla, señalar la limitación y no usarla como único soporte de cobertura.'),
        @('PLAN_ALERTAS_HISTORICO','12_Plan_Alertas','HISTÓRICO','La hoja conserva un plan previo y no representa por sí sola el estado final.','Usar las hojas finales de resultados para métricas y conclusiones.')
    )
    for ($i = 0; $i -lt $discrepancies.Count; $i++) {
        $row = 67 + $i
        for ($c = 0; $c -lt 5; $c++) {
            $address = ([char](65 + $c)).ToString() + $row
            Set-Cell $detection 'DISCREPANCIAS' $address $discrepancies[$i][$c] 'Registro de discrepancia metodológica'
        }
    }

    $detection.ForceFullCalculation = $true
    $excel.CalculateFullRebuild()
    $detection.Save()

    $formulaChecks = @(
        (Verify-Formula $detection 'COMPARACION_VR_WAZUH' 'C11' '=GRAFICAS!$Q$4&"/9"' '9/9'),
        (Verify-Formula $detection 'COMPARACION_VR_WAZUH' 'C12' '=GRAFICAS!$Q$5&"/9"' '3/9'),
        (Verify-Formula $detection 'COMPARACION_VR_WAZUH' 'C13' '=GRAFICAS!$Q$6&"/9"' '3/9'),
        (Verify-Formula $detection 'COMPARACION_VR_WAZUH' 'C14' '=GRAFICAS!$Q$7&"/9"' '4/9'),
        (Verify-Formula $detection 'RESUMEN_EJECUTIVO' 'B34' '=GRAFICAS!$Q$6' '3'),
        (Verify-Formula $detection 'RESUMEN_EJECUTIVO' 'B35' '=GRAFICAS!$Q$7' '4')
    )
    $detectionAudit = Audit-Workbook $detection
    $detection.Close($true)
    Release-ComObject $detection
    $detection = $null

    Set-Progress 'Corrección y validación del Excel de rendimiento'
    $benchmark = $excel.Workbooks.Open($BenchmarkPath, 0, $false)
    $method = 'El benchmark principal mide exclusivamente el impacto del cliente Velociraptor sobre el endpoint. SERVER_GUI se registra separadamente y se excluye del cálculo principal porque pertenece al servidor y no al cliente evaluado. IncludeServerGuiInTotal=False. ServerGuiExcludedFromClientMetrics=True.'
    Set-Cell $benchmark 'README' 'B9' $method 'Explicación metodológica canónica de SERVER_GUI'

    foreach ($item in @(
        @('CALIDAD_DATOS', '$A$1:$E$22'),
        @('DISCREPANCIAS', '$A$1:$E$6')
    )) {
        $sheet = Get-Sheet $benchmark $item[0]
        $range = $null
        try {
            $range = $sheet.Range($item[1])
            if (-not $sheet.AutoFilterMode) {
                [void]$range.AutoFilter()
                Add-Change ([System.IO.Path]::GetFileName($benchmark.FullName)) "$($item[0])!$($item[1])" 'Filtro habilitado'
            }
        }
        finally {
            Release-ComObject $range
            Release-ComObject $sheet
        }
    }
    $graphSheet = Get-Sheet $benchmark 'GRAFICAS'
    try {
        if ([string]$graphSheet.PageSetup.PrintArea -ne '$A$1:$N$114') {
            $graphSheet.PageSetup.PrintArea = '$A$1:$N$114'
            Add-Change ([System.IO.Path]::GetFileName($benchmark.FullName)) 'GRAFICAS!PrintArea' 'Área de impresión ajustada'
        }
    }
    finally { Release-ComObject $graphSheet }

    $benchmark.ForceFullCalculation = $true
    $excel.CalculateFullRebuild()
    $benchmark.Save()
    $benchmarkAudit = Audit-Workbook $benchmark
    $benchmark.Close($true)
    Release-ComObject $benchmark
    $benchmark = $null

    $excel.Quit()
    Release-ComObject $excel
    $excel = $null

    if ($detectionAudit.FormulaErrors.Count -gt 0 -or $benchmarkAudit.FormulaErrors.Count -gt 0) {
        throw 'Se detectaron errores de cálculo en los libros finales.'
    }
    if ($detectionAudit.EmptyCharts.Count -gt 0 -or $benchmarkAudit.EmptyCharts.Count -gt 0) {
        throw 'Se detectaron gráficos sin series de datos.'
    }

    $result = [pscustomobject]@{
        Corrections = $script:corrections
        FormulaChecks = $formulaChecks
        DetectionAudit = $detectionAudit
        BenchmarkAudit = $benchmarkAudit
        DetectionSHA256 = (Get-FileHash -LiteralPath $DetectionPath -Algorithm SHA256).Hash
        BenchmarkSHA256 = (Get-FileHash -LiteralPath $BenchmarkPath -Algorithm SHA256).Hash
        Changes = @($script:changes)
    }
    $result | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $LogPath -Encoding UTF8
    Set-Progress 'COMPLETADO'
    $result | ConvertTo-Json -Depth 5
}
finally {
    if ($null -ne $detection) {
        try { $detection.Close($false) } catch {}
        Release-ComObject $detection
    }
    if ($null -ne $benchmark) {
        try { $benchmark.Close($false) } catch {}
        Release-ComObject $benchmark
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
    }
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
}
