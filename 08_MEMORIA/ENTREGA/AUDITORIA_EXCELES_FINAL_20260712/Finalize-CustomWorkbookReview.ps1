param(
    [string]$Path = "C:\Users\julio\Desktop\TFM\08_MEMORIA\AUDITORIA_EXCELES_FINAL_20260712\TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712_CUSTOM_WORKING.xlsx"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function XlColor([int]$r, [int]$g, [int]$b) { return $r + (256 * $g) + (65536 * $b) }

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    $workbook = $excel.Workbooks.Open($Path, 0, $false)

    $dashboard = $workbook.Worksheets.Item("01_Dashboard")
    $deletedShapes = @()
    for ($i = $dashboard.Shapes.Count; $i -ge 1; $i--) {
        $shape = $dashboard.Shapes.Item($i)
        if ($shape.Type -ne 3) {
            $deletedShapes += $shape.Name
            $shape.Delete()
        }
    }
    $dashboard.Range("H15:N16").UnMerge()
    $dashboard.Range("H15:N16").ClearContents()
    $dashboard.Range("H15:N16").Merge()
    $dashboard.Range("H15").NumberFormat = "@"
    $dashboard.Range("H15").Value2 = "Wazuh base usa el ruleset nativo; Wazuh custom usa reglas TFM 110xxx. La métrica cuenta técnicas detectadas, no alertas."
    $dashboard.Range("H15:N16").Interior.Color = XlColor 221 235 247
    $dashboard.Range("H15:N16").Font.Italic = $true
    $dashboard.Range("H15:N16").WrapText = $true
    $dashboard.Range("H4").Value2 = "5 tácticas principales"
    $dashboard.Range("H5").Value2 = "6 con contexto ampliado"
    $dashboard.Range("H6").Value2 = "7 con red contextual"
    $dashboard.Rows("4:6").RowHeight = 24
    $dashboard.Rows("15:16").RowHeight = 22

    $tacticalChart = $dashboard.ChartObjects().Item("Chart").Chart
    $axis = $tacticalChart.Axes(2)
    $axis.TickLabels.NumberFormat = "0"
    $axis.MinimumScale = 0
    $axis.MaximumScale = 8
    $axis.MajorUnit = 1

    $graphs = $workbook.Worksheets.Item("GRAFICAS")
    $graphs.Range("M7").ClearContents()
    $graphs.PageSetup.Zoom = $false
    $graphs.PageSetup.FitToPagesWide = 1
    $graphs.PageSetup.FitToPagesTall = 1

    $catalog = $workbook.Worksheets.Item("07_Catalogo_Artifacts")
    $catalog.PageSetup.Zoom = $false
    $catalog.PageSetup.FitToPagesWide = 2
    $catalog.PageSetup.FitToPagesTall = 1

    $results = $workbook.Worksheets.Item("09_Resultados_Custom")
    $results.PageSetup.Zoom = $false
    $results.PageSetup.FitToPagesWide = 1
    $results.PageSetup.FitToPagesTall = 2

    $visibility = $workbook.Worksheets.Item("VISIBILIDAD_SISTEMA")
    $visibility.PageSetup.Zoom = $false
    $visibility.PageSetup.FitToPagesWide = 1
    $visibility.PageSetup.FitToPagesTall = 1

    $gaps = $workbook.Worksheets.Item("10_Gaps_Custom_P1P4")
    $gaps.Range("C9").Value2 = "Un artifact monolítico dificulta granularidad, mantenibilidad, priorización y ajuste de ruido"
    $gaps.Range("H9").Value2 = "Granularidad por riesgo, trazabilidad y control operativo"

    $dashboard.PageSetup.Zoom = $false
    $dashboard.PageSetup.FitToPagesWide = 1
    $dashboard.PageSetup.FitToPagesTall = 1

    $excel.Calculation = -4105
    $excel.CalculateFullRebuild()
    $workbook.Save()
    $workbook.Close($true)
    $workbook = $null
    $excel.Quit()
    $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    $item = Get-Item -LiteralPath $Path
    [pscustomobject]@{
        Status = "OK"
        Bytes = $item.Length
        SHA256 = (Get-FileHash -Algorithm SHA256 -LiteralPath $Path).Hash
        DeletedNonChartShapes = $deletedShapes
        TEC_HitsCell = "blank"
        CatalogPagesWide = 2
    } | ConvertTo-Json -Depth 4
}
finally {
    if ($workbook -ne $null) { try { $workbook.Close($false) } catch { } }
    if ($excel -ne $null) { try { $excel.Quit() } catch { } }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

