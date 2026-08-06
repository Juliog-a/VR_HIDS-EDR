[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [Parameter(Mandatory = $true)]
    [string]$OutputDir
)

$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

if (-not ('TFMVisibilityPdfQANative' -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class TFMVisibilityPdfQANative {
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
    [pscustomobject]@{ Sheet='GRAFICAS'; Area='$A$17:$P$33'; Name='02_graficas_runner_fp'; Orientation=$xlLandscape; Tall=1 },
    [pscustomobject]@{ Sheet='GRAFICAS'; Area='$M$48:$P$59'; Name='03_graficas_heatmap'; Orientation=$xlPortrait; Tall=1 },
    [pscustomobject]@{ Sheet='GRAFICAS'; Area='$F$72:$J$79'; Name='04_graficas_estados'; Orientation=$xlPortrait; Tall=1 },
    [pscustomobject]@{ Sheet='04_Control_Publicos'; Area='$A$4:$U$67'; Name='05_control_publicos_a_u'; Orientation=$xlLandscape; Tall=2 },
    [pscustomobject]@{ Sheet='04_Control_Publicos'; Area='$S$4:$AE$67'; Name='06_control_publicos_s_ae'; Orientation=$xlLandscape; Tall=2 },
    [pscustomobject]@{ Sheet='05_Matriz_Resultados'; Area='$A$1:$T$13'; Name='07_matriz_resultados'; Orientation=$xlLandscape; Tall=1 },
    [pscustomobject]@{ Sheet='08_Benignas_FP'; Area='$A$1:$J$18'; Name='08_benignas_fp'; Orientation=$xlLandscape; Tall=1 },
    [pscustomobject]@{ Sheet='15_Publicos_Definitivo'; Area='$A$1:$T$15'; Name='09_publicos_definitivo'; Orientation=$xlLandscape; Tall=1 },
    [pscustomobject]@{ Sheet='RESUMEN_EJECUTIVO'; Area='$A$1:$C$33'; Name='10_resumen_ejecutivo'; Orientation=$xlPortrait; Tall=1 }
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
    [void][TFMVisibilityPdfQANative]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$pidValue)
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
