[CmdletBinding()]
param(
    [string]$OutputDir = "C:\Users\julio\Desktop\TFM\tmp\pdfs\excel_visual_qa"
)

$ErrorActionPreference = "Stop"
$xlTypePDF = 0
$xlQualityStandard = 0
$xlLandscape = 2
$xlPortrait = 1
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

if (-not ("TFMExcelPdfNative" -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class TFMExcelPdfNative {
    [DllImport("user32.dll", SetLastError=true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
"@
}

function Release-ComObject {
    param($Object)
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object) } catch {}
    }
}

$targets = @(
    [pscustomobject]@{ Book="C:\Users\julio\Desktop\TFM\08_MEMORIA\ENTREGABLE\TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712.xlsx"; Prefix="VIS"; Sheet="01_Dashboard"; Area="A1:Q70"; Orientation=$xlLandscape; Wide=1; Tall=3 },
    [pscustomobject]@{ Book="C:\Users\julio\Desktop\TFM\08_MEMORIA\ENTREGABLE\TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712.xlsx"; Prefix="VIS"; Sheet="RESUMEN_EJECUTIVO"; Area="A1:C33"; Orientation=$xlPortrait; Wide=1; Tall=2 },
    [pscustomobject]@{ Book="C:\Users\julio\Desktop\TFM\08_MEMORIA\ENTREGABLE\TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712.xlsx"; Prefix="VIS"; Sheet="15_Publicos_Definitivo"; Area="A1:S14"; Orientation=$xlLandscape; Wide=1; Tall=2 },
    [pscustomobject]@{ Book="C:\Users\julio\Desktop\TFM\08_MEMORIA\ENTREGABLE\TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712.xlsx"; Prefix="VIS"; Sheet="COMPARACION_VR_WAZUH"; Area="A1:M8"; Orientation=$xlLandscape; Wide=1; Tall=1 },
    [pscustomobject]@{ Book="C:\Users\julio\Desktop\TFM\08_MEMORIA\ENTREGABLE\TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712.xlsx"; Prefix="VIS"; Sheet="13_Hayabusa_Resumen"; Area="A1:D35"; Orientation=$xlPortrait; Wide=1; Tall=2 },
    [pscustomobject]@{ Book="C:\Users\julio\Desktop\TFM\08_MEMORIA\ENTREGABLE\TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712.xlsx"; Prefix="VIS"; Sheet="GRAFICAS"; Area="A1:R80"; Orientation=$xlLandscape; Wide=1; Tall=4 },
    [pscustomobject]@{ Book="C:\Users\julio\Desktop\TFM\08_MEMORIA\ENTREGABLE\TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_DEFINITIVO_20260712.xlsx"; Prefix="BEN"; Sheet="RESUMEN_EJECUTIVO"; Area="A1:C18"; Orientation=$xlPortrait; Wide=1; Tall=1 },
    [pscustomobject]@{ Book="C:\Users\julio\Desktop\TFM\08_MEMORIA\ENTREGABLE\TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_DEFINITIVO_20260712.xlsx"; Prefix="BEN"; Sheet="RUNS_VALIDOS"; Area="A1:V10"; Orientation=$xlLandscape; Wide=1; Tall=1 },
    [pscustomobject]@{ Book="C:\Users\julio\Desktop\TFM\08_MEMORIA\ENTREGABLE\TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_DEFINITIVO_20260712.xlsx"; Prefix="BEN"; Sheet="GRAFICAS"; Area="A1:N112"; Orientation=$xlLandscape; Wide=1; Tall=0 }
)

$excel = $null
$excelPid = 0
$log = [System.Collections.Generic.List[object]]::new()
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    [uint32]$pidValue = 0
    [void][TFMExcelPdfNative]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$pidValue)
    $excelPid = [int]$pidValue

    foreach ($group in ($targets | Group-Object Book)) {
        $workbook = $null
        try {
            $workbook = $excel.Workbooks.Open($group.Name, 0, $true)
            foreach ($target in $group.Group) {
                $worksheet = $null
                $pageSetup = $null
                try {
                    $worksheet = $workbook.Worksheets.Item($target.Sheet)
                    $pageSetup = $worksheet.PageSetup
                    $pageSetup.PrintArea = $worksheet.Range($target.Area).Address()
                    $pageSetup.Orientation = $target.Orientation
                    $pageSetup.Zoom = $false
                    $pageSetup.FitToPagesWide = $target.Wide
                    if ($target.Tall -eq 0) { $pageSetup.FitToPagesTall = $false } else { $pageSetup.FitToPagesTall = $target.Tall }
                    $pageSetup.CenterHorizontally = $true
                    $outputPath = Join-Path $OutputDir ("{0}_{1}.pdf" -f $target.Prefix,($target.Sheet -replace '[^A-Za-z0-9_-]','_'))
                    if (Test-Path -LiteralPath $outputPath) { Remove-Item -LiteralPath $outputPath -Force }
                    $worksheet.ExportAsFixedFormat($xlTypePDF, $outputPath, $xlQualityStandard, $true, $false)
                    $item = Get-Item -LiteralPath $outputPath
                    $log.Add([pscustomobject]@{ Workbook=$group.Name; Sheet=$target.Sheet; Area=$target.Area; Path=$outputPath; SizeBytes=$item.Length; Exported=($item.Length -gt 0) })
                } finally {
                    Release-ComObject $pageSetup
                    Release-ComObject $worksheet
                }
            }
            $workbook.Close($false)
        } finally {
            if ($workbook) { try { $workbook.Close($false) } catch {}; Release-ComObject $workbook }
        }
    }
} finally {
    if ($excel) { try { $excel.Quit() } catch {}; Release-ComObject $excel }
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 2
    if ($excelPid -gt 0 -and (Get-Process -Id $excelPid -ErrorAction SilentlyContinue)) { Stop-Process -Id $excelPid -Force }
}

$log | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDir "EXPORT_LOG.json") -Encoding utf8
$log | Select-Object Sheet,Area,SizeBytes,Exported,Path
