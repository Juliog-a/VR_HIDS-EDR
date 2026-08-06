[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [Parameter(Mandatory = $true)]
    [string]$OutputDir
)

$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

if (-not ('TFMWazuhPdfQANative' -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class TFMWazuhPdfQANative {
    [DllImport("user32.dll", SetLastError=true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
"@
}

function Release-ComObject {
    param([object]$Object)
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object) } catch {}
    }
}

$xlTypePDF = 0
$xlQualityStandard = 0
$xlLandscape = 2
$xlPortrait = 1

$targets = @(
    [pscustomobject]@{ Sheet='01_Dashboard'; Area='$A$1:$Q$70'; Name='01_dashboard'; Orientation=$xlLandscape; Tall=2 },
    [pscustomobject]@{ Sheet='GRAFICAS'; Area='$G$17:$W$40'; Name='02_graficas_wazuh_chart'; Orientation=$xlLandscape; Tall=1 },
    [pscustomobject]@{ Sheet='GRAFICAS'; Area='$M$48:$Q$60'; Name='03_graficas_heatmap'; Orientation=$xlPortrait; Tall=1 },
    [pscustomobject]@{ Sheet='COMPARACION_VR_WAZUH'; Area='$A$1:$M$23'; Name='04_comparacion_vr_wazuh'; Orientation=$xlLandscape; Tall=2 },
    [pscustomobject]@{ Sheet='05_Matriz_Resultados'; Area='$A$1:$Z$13'; Name='05_matriz_completa'; Orientation=$xlLandscape; Tall=1 },
    [pscustomobject]@{ Sheet='05_Matriz_Resultados'; Area='$T$1:$Z$13'; Name='06_matriz_wazuh_foco'; Orientation=$xlLandscape; Tall=1 },
    [pscustomobject]@{ Sheet='RESUMEN_EJECUTIVO'; Area='$A$1:$C$37'; Name='07_resumen_ejecutivo'; Orientation=$xlPortrait; Tall=2 },
    [pscustomobject]@{ Sheet='WAZUH_DETALLE'; Area='$A$1:$H$19'; Name='08_wazuh_detalle'; Orientation=$xlLandscape; Tall=1 }
)

$excel = $null
$workbook = $null
$worksheets = $null
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
    [void][TFMWazuhPdfQANative]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$pidValue)
    $excelPid = [int]$pidValue
    $workbook = $excel.Workbooks.Open($Path, 0, $true)
    $worksheets = $workbook.Worksheets

    foreach ($target in $targets) {
        $sheet = $null
        try {
            $sheet = $worksheets.Item($target.Sheet)
            $sheet.Activate()
            try { $excel.PrintCommunication = $false } catch {}
            $sheet.PageSetup.PrintArea = $target.Area
            $sheet.PageSetup.Orientation = $target.Orientation
            $sheet.PageSetup.Zoom = $false
            $sheet.PageSetup.FitToPagesWide = 1
            $sheet.PageSetup.FitToPagesTall = $target.Tall
            $sheet.PageSetup.LeftMargin = $excel.InchesToPoints(0.2)
            $sheet.PageSetup.RightMargin = $excel.InchesToPoints(0.2)
            $sheet.PageSetup.TopMargin = $excel.InchesToPoints(0.25)
            $sheet.PageSetup.BottomMargin = $excel.InchesToPoints(0.25)
            $sheet.PageSetup.HeaderMargin = $excel.InchesToPoints(0.1)
            $sheet.PageSetup.FooterMargin = $excel.InchesToPoints(0.1)
            try { $excel.PrintCommunication = $true } catch {}
            $outPath = Join-Path $OutputDir ($target.Name + '.pdf')
            $sheet.ExportAsFixedFormat($xlTypePDF, $outPath, $xlQualityStandard, $true, $false)
            if (-not (Test-Path -LiteralPath $outPath) -or (Get-Item -LiteralPath $outPath).Length -eq 0) {
                throw "No se pudo exportar $($target.Sheet) $($target.Area)"
            }
            $log.Add([pscustomobject]@{
                Sheet = $target.Sheet
                Area = $target.Area
                Pdf = $outPath
                Bytes = (Get-Item -LiteralPath $outPath).Length
            })
        }
        finally { Release-ComObject $sheet }
    }
}
finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $worksheets
    Release-ComObject $workbook
    Release-ComObject $excel
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 2
    if ($excelPid -gt 0 -and (Get-Process -Id $excelPid -ErrorAction SilentlyContinue)) {
        Stop-Process -Id $excelPid -Force
    }
}

$jsonPath = Join-Path $OutputDir 'PDF_QA_EXPORT_LOG.json'
$json = $log | ConvertTo-Json -Depth 5
[System.IO.File]::WriteAllText($jsonPath, $json, [System.Text.UTF8Encoding]::new($false))
$log | Format-Table Sheet,Area,Bytes,Pdf -AutoSize
