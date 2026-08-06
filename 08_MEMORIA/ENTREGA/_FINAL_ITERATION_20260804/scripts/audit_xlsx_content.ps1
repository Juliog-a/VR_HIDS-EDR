param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,
    [Parameter(Mandatory = $true)]
    [string]$OutPath
)

$ErrorActionPreference = 'Stop'

function Release-ComObject {
    param([object]$Object)
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Get-RangeSnapshot {
    param(
        [object]$Workbook,
        [string]$SheetName,
        [string]$Address
    )
    $sheet = $null
    $range = $null
    try {
        $sheet = $Workbook.Worksheets.Item($SheetName)
        $range = $sheet.Range($Address)
        $values = $range.Value2
        $formulas = $range.Formula
        $rows = [System.Collections.Generic.List[object]]::new()
        if ($range.Cells.Count -eq 1) {
            [void]$rows.Add([pscustomobject]@{
                Address = [string]$range.Address()
                Value = $values
                Formula = [string]$formulas
            })
        }
        else {
            for ($r = 1; $r -le $range.Rows.Count; $r++) {
                for ($c = 1; $c -le $range.Columns.Count; $c++) {
                    $cell = $range.Cells.Item($r, $c)
                    try {
                        $value = $values[$r, $c]
                        $formula = $formulas[$r, $c]
                        if ($null -ne $value -or (-not [string]::IsNullOrWhiteSpace([string]$formula))) {
                            [void]$rows.Add([pscustomobject]@{
                                Address = [string]$cell.Address()
                                Value = $value
                                Formula = [string]$formula
                            })
                        }
                    }
                    finally {
                        Release-ComObject $cell
                    }
                }
            }
        }
        return [pscustomobject]@{
            Sheet = $SheetName
            Range = $Address
            Cells = @($rows)
        }
    }
    finally {
        Release-ComObject $range
        Release-ComObject $sheet
    }
}

$excel = $null
$wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    $wb = $excel.Workbooks.Open($WorkbookPath, 0, $true, 5, '', '', $true)
    $excel.CalculateFullRebuild()

    $regex = '(?i)(379|alertas?|CLIENT_EVENT|82\s+filas|no\s+reconciliad|10/10|0\s*hits|Perfil\s+m[aá]ximo|cobertura\s+t[aá]ctica|MITRE\s+ATT&CK\s+Enterprise|HIST[OÓ]RICO|alerta\s+SOC\s+pendiente|TEC_escenario|TEC_detector|EvidenceRole|I383|validaci[oó]n|tácticas\s+representadas)'
    $cellMatches = [System.Collections.Generic.List[object]]::new()
    foreach ($sheet in @($wb.Worksheets)) {
        $used = $null
        try {
            $used = $sheet.UsedRange
            $values = $used.Value2
            if ($used.Cells.Count -eq 1) {
                if ([string]$values -match $regex) {
                    [void]$cellMatches.Add([pscustomobject]@{
                        Sheet = [string]$sheet.Name
                        Address = [string]$used.Address()
                        Value = [string]$values
                    })
                }
            }
            else {
                $startRow = [int]$used.Row
                $startCol = [int]$used.Column
                for ($r = 1; $r -le $used.Rows.Count; $r++) {
                    for ($c = 1; $c -le $used.Columns.Count; $c++) {
                        $value = $values[$r, $c]
                        if ($null -ne $value -and [string]$value -match $regex) {
                            $cell = $sheet.Cells.Item($startRow + $r - 1, $startCol + $c - 1)
                            try {
                                [void]$cellMatches.Add([pscustomobject]@{
                                    Sheet = [string]$sheet.Name
                                    Address = [string]$cell.Address()
                                    Value = [string]$value
                                })
                            }
                            finally {
                                Release-ComObject $cell
                            }
                        }
                    }
                }
            }
        }
        finally {
            Release-ComObject $used
            Release-ComObject $sheet
        }
    }

    $snapshots = [System.Collections.Generic.List[object]]::new()
    foreach ($spec in @(
        @('01_Dashboard','A1:H12'),
        @('06_Resumen','A1:H18'),
        @('RESUMEN_EJECUTIVO','A1:H40'),
        @('FP_HITS','A1:O8'),
        @('09_Resultados_Custom','A1:I24'),
        @('09_Resultados_Custom','A48:I58'),
        @('12_Plan_Alertas','A1:H10'),
        @('ALERTAS_VR','A1:O8'),
        @('GRAFICAS','O1:R10'),
        @('COMPARACION_VR_WAZUH','A1:H18')
    )) {
        [void]$snapshots.Add((Get-RangeSnapshot -Workbook $wb -SheetName $spec[0] -Address $spec[1]))
    }

    $result = [pscustomobject]@{
        Path = $WorkbookPath
        SHA256 = (Get-FileHash -LiteralPath $WorkbookPath -Algorithm SHA256).Hash
        Matches = @($cellMatches)
        Snapshots = @($snapshots)
    }
    $result | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $OutPath -Encoding UTF8
    [pscustomobject]@{
        MatchCount = $cellMatches.Count
        Matches = @($cellMatches)
        Snapshots = @($snapshots)
    } | ConvertTo-Json -Depth 8
}
finally {
    if ($null -ne $wb) {
        try { $wb.Close($false) } catch {}
        Release-ComObject $wb
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
    }
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
}
