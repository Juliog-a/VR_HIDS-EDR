[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [Parameter(Mandatory = $true)]
    [string]$OutputDir
)

$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

if (-not ('TFMVisibilityQANative' -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class TFMVisibilityQANative {
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

function Export-RangePicture {
    param($Worksheet, [string]$Address, [string]$OutputPath)
    $range = $null
    $chartObjects = $null
    $chartObject = $null
    $chart = $null
    try {
        $Worksheet.Activate()
        $range = $Worksheet.Range($Address)
        $range.CopyPicture(1, 2)
        $width = [math]::Min([math]::Max([double]$range.Width, 900), 4200)
        $height = [math]::Min([math]::Max([double]$range.Height, 420), 3200)
        $chartObjects = $Worksheet.ChartObjects()
        $chartObject = $chartObjects.Add(0, 0, $width, $height)
        $chart = $chartObject.Chart
        $chart.Paste()
        $ok = $chart.Export($OutputPath, 'PNG')
        if (-not $ok -or -not (Test-Path -LiteralPath $OutputPath)) { throw "No se pudo exportar $Address" }
    }
    finally {
        if ($null -ne $chartObject) { try { $chartObject.Delete() } catch {} }
        Release-ComObject $chart
        Release-ComObject $chartObject
        Release-ComObject $chartObjects
        Release-ComObject $range
    }
}

function Export-SheetCharts {
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
                $outPath = Join-Path $OutputDirectory ("{0}_chart_{1:D2}_{2}.png" -f $Prefix, $i, $safeName)
                $ok = $chart.Export($outPath, 'PNG')
                $Log.Add([pscustomobject]@{ Type='CHART'; Sheet=$Worksheet.Name; Source=$chartObject.Name; Output=$outPath; Exported=[bool]$ok })
            }
            finally {
                Release-ComObject $chart
                Release-ComObject $chartObject
            }
        }
    }
    finally { Release-ComObject $chartObjects }
}

$ranges = @(
    @('01_Dashboard', 'A1:Q16', 'dashboard_kpi'),
    @('GRAFICAS', 'A18:E30', 'graficas_runner_fp_note'),
    @('GRAFICAS', 'M49:P59', 'graficas_heatmap_capacidades'),
    @('GRAFICAS', 'F73:J79', 'graficas_estados_publicos'),
    @('04_Control_Publicos', 'A4:U34', 'control_publicos_1'),
    @('04_Control_Publicos', 'A35:U67', 'control_publicos_2'),
    @('05_Matriz_Resultados', 'A1:T13', 'matriz_resultados'),
    @('08_Benignas_FP', 'A1:J18', 'benignas_fp'),
    @('15_Publicos_Definitivo', 'A1:T15', 'publicos_definitivo'),
    @('RESUMEN_EJECUTIVO', 'A1:C33', 'resumen_ejecutivo')
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
    [void][TFMVisibilityQANative]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$pidValue)
    $excelPid = [int]$pidValue
    $workbook = $excel.Workbooks.Open($Path, 0, $true)
    $worksheets = $workbook.Worksheets

    foreach ($item in $ranges) {
        $sheet = $null
        try {
            $sheet = $worksheets.Item($item[0])
            $outPath = Join-Path $OutputDir ("VIS_{0}.png" -f $item[2])
            Export-RangePicture $sheet $item[1] $outPath
            $log.Add([pscustomobject]@{ Type='RANGE'; Sheet=$item[0]; Source=$item[1]; Output=$outPath; Exported=$true })
        }
        finally { Release-ComObject $sheet }
    }

    foreach ($sheetName in @('01_Dashboard', 'GRAFICAS')) {
        $sheet = $null
        try {
            $sheet = $worksheets.Item($sheetName)
            Export-SheetCharts $sheet ("VIS_{0}" -f $sheetName) $OutputDir $log
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

$jsonPath = Join-Path $OutputDir 'VISUAL_QA_AJUSTE_FINAL_LOG.json'
$json = $log | ConvertTo-Json -Depth 5
[System.IO.File]::WriteAllText($jsonPath, $json, [System.Text.UTF8Encoding]::new($false))
$log | Select-Object Type,Sheet,Source,Exported,Output
