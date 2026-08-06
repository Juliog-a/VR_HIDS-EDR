[CmdletBinding()]
param(
    [string]$OutputDir = "C:\Users\julio\Desktop\TFM\08_MEMORIA\AUDITORIA_EXCELES_FINAL_20260712\VISUAL_QA"
)

$ErrorActionPreference = "Stop"
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

if (-not ("TFMExcelVisualNative" -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class TFMExcelVisualNative {
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

function Export-RangePicture {
    param($Worksheet, [string]$Address, [string]$OutputPath)
    $range = $null
    $chartObject = $null
    $chart = $null
    try {
        $Worksheet.Activate()
        $range = $Worksheet.Range($Address)
        $range.CopyPicture(1, 2)
        $width = [math]::Min([math]::Max([double]$range.Width, 700), 2800)
        $height = [math]::Min([math]::Max([double]$range.Height, 320), 2000)
        $chartObject = $Worksheet.ChartObjects().Add(0, 0, $width, $height)
        $chart = $chartObject.Chart
        $chart.Paste()
        $ok = $chart.Export($OutputPath, "PNG")
        if (-not $ok -or -not (Test-Path -LiteralPath $OutputPath)) { throw "No se pudo exportar $Address" }
    } finally {
        if ($chartObject) { try { $chartObject.Delete() } catch {} }
        Release-ComObject $chart
        Release-ComObject $chartObject
        Release-ComObject $range
    }
}

function Export-AllCharts {
    param($Worksheet, [string]$Prefix, [string]$OutputDirectory, $Log)
    $chartObjects = $null
    try {
        $chartObjects = $Worksheet.ChartObjects()
        for ($i = 1; $i -le $chartObjects.Count; $i++) {
            $chartObject = $null
            $chart = $null
            try {
                $chartObject = $chartObjects.Item($i)
                $chart = $chartObject.Chart
                $safeName = ([string]$chartObject.Name -replace '[^A-Za-z0-9_-]', '_')
                $path = Join-Path $OutputDirectory ("{0}_chart_{1:D2}_{2}.png" -f $Prefix,$i,$safeName)
                $ok = $chart.Export($path, "PNG")
                $Log.Add([pscustomobject]@{ Type="CHART"; Sheet=$Worksheet.Name; Source=$chartObject.Name; Path=$path; Exported=[bool]$ok })
            } finally {
                Release-ComObject $chart
                Release-ComObject $chartObject
            }
        }
    } finally {
        Release-ComObject $chartObjects
    }
}

$targets = @(
    [pscustomobject]@{
        Path="C:\Users\julio\Desktop\TFM\08_MEMORIA\ENTREGABLE\TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712.xlsx";
        Prefix="VIS";
        Ranges=@(
            @("01_Dashboard","A1:Q16","dashboard_kpi"),
            @("RESUMEN_EJECUTIVO","A1:C34","resumen_ejecutivo"),
            @("15_Publicos_Definitivo","A1:S14","publicos_definitivo"),
            @("COMPARACION_VR_WAZUH","A1:M8","comparacion_vr_wazuh"),
            @("13_Hayabusa_Resumen","A1:D35","hayabusa_resumen"),
            @("GRAFICAS","A60:R80","graficas_fuentes_finales")
        );
        ChartSheets=@("01_Dashboard","GRAFICAS")
    },
    [pscustomobject]@{
        Path="C:\Users\julio\Desktop\TFM\08_MEMORIA\ENTREGABLE\TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_DEFINITIVO_20260712.xlsx";
        Prefix="BEN";
        Ranges=@(
            @("RESUMEN_EJECUTIVO","A1:C12","resumen_ejecutivo"),
            @("RUNS_VALIDOS","A1:K10","runs_validos"),
            @("GRAFICAS","A1:K27","graficas_fuentes")
        );
        ChartSheets=@("GRAFICAS")
    }
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
    [void][TFMExcelVisualNative]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$pidValue)
    $excelPid = [int]$pidValue
    foreach ($target in $targets) {
        $workbook = $null
        try {
            $workbook = $excel.Workbooks.Open($target.Path, 0, $true)
            foreach ($rangeTarget in $target.Ranges) {
                $worksheet = $null
                try {
                    $worksheet = $workbook.Worksheets.Item($rangeTarget[0])
                    $outPath = Join-Path $OutputDir ("{0}_{1}.png" -f $target.Prefix,$rangeTarget[2])
                    Export-RangePicture -Worksheet $worksheet -Address $rangeTarget[1] -OutputPath $outPath
                    $log.Add([pscustomobject]@{ Type="RANGE"; Sheet=$rangeTarget[0]; Source=$rangeTarget[1]; Path=$outPath; Exported=$true })
                } catch {
                    $log.Add([pscustomobject]@{ Type="RANGE"; Sheet=$rangeTarget[0]; Source=$rangeTarget[1]; Path=$outPath; Exported=$false; Error=$_.Exception.Message })
                } finally {
                    Release-ComObject $worksheet
                }
            }
            foreach ($sheetName in $target.ChartSheets) {
                $worksheet = $null
                try {
                    $worksheet = $workbook.Worksheets.Item($sheetName)
                    Export-AllCharts -Worksheet $worksheet -Prefix ("{0}_{1}" -f $target.Prefix,$sheetName) -OutputDirectory $OutputDir -Log $log
                } finally {
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

$log | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDir "VISUAL_QA_EXPORT_LOG.json") -Encoding utf8
$log | Select-Object Type,Sheet,Source,Exported,Path,Error
