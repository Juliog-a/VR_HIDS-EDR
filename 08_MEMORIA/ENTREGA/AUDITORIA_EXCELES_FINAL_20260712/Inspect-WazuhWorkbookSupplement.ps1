param(
    [string]$Path = "C:\Users\julio\Desktop\TFM\08_MEMORIA\ENTREGABLE\TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712.xlsx",
    [string]$OutputPath = "C:\Users\julio\Desktop\TFM\08_MEMORIA\AUDITORIA_EXCELES_FINAL_20260712\_INSPECCION_WAZUH_SUPLEMENTARIA_ANTES.json"
)

$ErrorActionPreference = 'Stop'

function Release-ComObject {
    param($Object)
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object) } catch {}
    }
}

function Get-RangeRows {
    param($Worksheet, [string]$Address)
    $range = $null
    try {
        $range = $Worksheet.Range($Address)
        $values = $range.Value2
        $rows = [System.Collections.Generic.List[object]]::new()
        $firstRow = [int]$range.Row
        $firstColumn = [int]$range.Column
        $rowCount = [int]$range.Rows.Count
        $columnCount = [int]$range.Columns.Count
        for ($r = 1; $r -le $rowCount; $r++) {
            $item = [ordered]@{ Row = $firstRow + $r - 1 }
            for ($c = 1; $c -le $columnCount; $c++) {
                $columnNumber = $firstColumn + $c - 1
                $columnName = ''
                $n = $columnNumber
                while ($n -gt 0) {
                    $n--
                    $columnName = [char](65 + ($n % 26)) + $columnName
                    $n = [math]::Floor($n / 26)
                }
                if ($values -is [System.Array]) { $item[$columnName] = $values[$r,$c] }
                else { $item[$columnName] = $values }
            }
            $rows.Add([pscustomobject]$item)
        }
        return @($rows)
    } finally {
        Release-ComObject $range
    }
}

function Get-SheetDetails {
    param($Workbook, [string]$Name, [string]$Address)
    $sheet = $null
    $used = $null
    $tables = $null
    try {
        $sheet = $Workbook.Worksheets.Item($Name)
        $used = $sheet.UsedRange
        $tables = $sheet.ListObjects
        $tableRows = [System.Collections.Generic.List[object]]::new()
        for ($i = 1; $i -le [int]$tables.Count; $i++) {
            $table = $null
            try {
                $table = $tables.Item($i)
                $tableRows.Add([pscustomobject]@{ Name = [string]$table.Name; Range = [string]$table.Range.Address() })
            } finally { Release-ComObject $table }
        }
        return [pscustomobject]@{
            Name = $Name
            UsedRange = [string]$used.Address()
            Tables = @($tableRows)
            Rows = Get-RangeRows -Worksheet $sheet -Address $Address
        }
    } finally {
        Release-ComObject $tables
        Release-ComObject $used
        Release-ComObject $sheet
    }
}

$excel = $null
$workbook = $null
$result = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $workbook = $excel.Workbooks.Open($Path, 0, $true)
    $result = [pscustomobject]@{
        Path = $Path
        ReadOnly = [bool]$workbook.ReadOnly
        WazuhDetalle = Get-SheetDetails $workbook 'WAZUH_DETALLE' 'A1:G20'
        Fuentes = Get-SheetDetails $workbook 'FUENTES' 'A205:G230'
        Discrepancias = Get-SheetDetails $workbook 'DISCREPANCIAS' 'A45:E70'
        Comparacion = Get-SheetDetails $workbook 'COMPARACION_VR_WAZUH' 'A1:M30'
        Matriz = Get-SheetDetails $workbook '05_Matriz_Resultados' 'A4:Z13'
    }
} finally {
    if ($workbook) { try { $workbook.Close($false) } catch {}; Release-ComObject $workbook }
    if ($excel) { try { $excel.Quit() } catch {}; Release-ComObject $excel }
    [gc]::Collect(); [gc]::WaitForPendingFinalizers()
}

$result | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $OutputPath -Encoding utf8
[pscustomobject]@{ OutputPath = $OutputPath; Bytes = (Get-Item -LiteralPath $OutputPath).Length }
