param(
    [string]$Path = "C:\Users\julio\Desktop\TFM\08_MEMORIA\ENTREGABLE\TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712.xlsx",
    [string]$OutputDirectory = "C:\Users\julio\Desktop\TFM\08_MEMORIA\AUDITORIA_EXCELES_FINAL_20260712\CUSTOM_PREVIEW_PDF"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $workbook = $excel.Workbooks.Open($Path, 0, $true)
    $targets = @(
        @{ Name = "01_Dashboard"; File = "01_Dashboard.pdf" },
        @{ Name = "07_Catalogo_Artifacts"; File = "02_07_Catalogo_Artifacts.pdf" },
        @{ Name = "09_Resultados_Custom"; File = "03_09_Resultados_Custom.pdf" },
        @{ Name = "VISIBILIDAD_SISTEMA"; File = "04_VISIBILIDAD_SISTEMA.pdf" },
        @{ Name = "GRAFICAS"; File = "05_GRAFICAS.pdf" }
    )
    $results = @()
    foreach ($target in $targets) {
        $sheet = $workbook.Worksheets.Item($target.Name)
        $output = Join-Path $OutputDirectory $target.File
        if (Test-Path -LiteralPath $output) { Remove-Item -LiteralPath $output -Force }
        $sheet.ExportAsFixedFormat(0, $output, 0, $true, $false)
        $item = Get-Item -LiteralPath $output
        $results += [pscustomobject]@{ Sheet = $target.Name; Path = $output; Bytes = $item.Length }
    }
    $workbook.Close($false)
    $workbook = $null
    $excel.Quit()
    $excel = $null
    $results | ConvertTo-Json -Depth 3
}
finally {
    if ($workbook -ne $null) { try { $workbook.Close($false) } catch { } }
    if ($excel -ne $null) { try { $excel.Quit() } catch { } }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
