[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Path
)

$ErrorActionPreference = 'Stop'

if (-not ('TFMWazuhDisplayNative' -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class TFMWazuhDisplayNative {
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

$excel = $null
$workbook = $null
$dashboard = $null
$comparison = $null
$graficas = $null
$closed = $null
$excelPid = 0
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    try { $excel.AutomationSecurity = 3 } catch {}
    [uint32]$pidValue = 0
    [void][TFMWazuhDisplayNative]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$pidValue)
    $excelPid = [int]$pidValue

    $workbook = $excel.Workbooks.Open($Path, 0, $false)
    $dashboard = $workbook.Worksheets.Item('01_Dashboard')
    $comparison = $workbook.Worksheets.Item('COMPARACION_VR_WAZUH')
    $graficas = $workbook.Worksheets.Item('GRAFICAS')
    $closed = $workbook.Worksheets.Item('17_Incoherencias_Cerradas')

    $graficas.Range('R60').Value2 = 'Wazuh custom'
    $closed.Range('D20').Value2 = 'Reportar base 3/9 y custom 4/9; no usar una etiqueta agregada ambigua.'

    foreach ($address in @('B14', 'B16')) {
        $cell = $null
        try {
            $cell = $dashboard.Range($address)
            $cell.ClearContents()
            if ($address -eq 'B14') {
                $cell.Value2 = "'4/9"
            }
            else {
                $cell.Value2 = "'3/9"
            }
        }
        finally { Release-ComObject $cell }
    }

    $displayValues = [ordered]@{
        'B11' = '9/9'; 'D11' = '9/9'
        'B12' = '9/9'; 'D12' = '3/9'
        'B13' = '9/9'; 'D13' = '3/9'
        'B14' = '9/9'; 'D14' = '4/9'
        'B20' = '7/9'
    }
    foreach ($entry in $displayValues.GetEnumerator()) {
        $cell = $null
        try {
            $cell = $comparison.Range($entry.Key)
            $cell.ClearContents()
            $cell.Value2 = "'" + [string]$entry.Value
        }
        finally { Release-ComObject $cell }
    }

    foreach ($chartSpec in @(
        [pscustomobject]@{ Sheet=$dashboard; Name='TFM_Final_Public_TEC' },
        [pscustomobject]@{ Sheet=$graficas; Name='Chart 4' }
    )) {
        $chartObjects = $null
        $chartObject = $null
        $chart = $null
        $axis = $null
        try {
            $chartObjects = $chartSpec.Sheet.ChartObjects()
            $chartObject = $chartObjects.Item($chartSpec.Name)
            $chart = $chartObject.Chart
            $axis = $chart.Axes(2)
            $axis.MinimumScale = 0
            $axis.MaximumScale = 1.1
            $axis.MajorUnit = 1
            $axis.TickLabels.NumberFormat = '0'
        }
        finally {
            Release-ComObject $axis
            Release-ComObject $chart
            Release-ComObject $chartObject
            Release-ComObject $chartObjects
        }
    }

    $excel.Calculation = -4105
    $excel.CalculateFull()
    $workbook.Save()
}
finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $comparison
    Release-ComObject $graficas
    Release-ComObject $closed
    Release-ComObject $dashboard
    Release-ComObject $workbook
    Release-ComObject $excel
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 2
    if ($excelPid -gt 0 -and (Get-Process -Id $excelPid -ErrorAction SilentlyContinue)) {
        Stop-Process -Id $excelPid -Force
    }
}

$item = Get-Item -LiteralPath $Path
$hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $Path).Hash
[pscustomobject]@{ Path=$item.FullName; Bytes=$item.Length; SHA256=$hash } | Format-List
