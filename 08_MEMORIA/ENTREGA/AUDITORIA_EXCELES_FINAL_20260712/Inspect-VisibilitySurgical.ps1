param(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

function Release-ComObject {
    param([object]$Object)
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Get-CellValue {
    param($Worksheet, [string]$Address)
    $range = $null
    try {
        $range = $Worksheet.Range($Address)
        return $range.Value2
    }
    finally {
        Release-ComObject $range
    }
}

function Get-RangeMatrix {
    param($Worksheet, [string]$Address)
    $range = $null
    try {
        $range = $Worksheet.Range($Address)
        $values = $range.Value2
        $rows = @()
        if ($values -is [System.Array] -and $values.Rank -eq 2) {
            for ($r = 1; $r -le $values.GetLength(0); $r++) {
                $row = [ordered]@{ Row = $range.Row + $r - 1 }
                for ($c = 1; $c -le $values.GetLength(1); $c++) {
                    $columnNumber = $range.Column + $c - 1
                    $columnName = ''
                    $n = $columnNumber
                    while ($n -gt 0) {
                        $n--
                        $columnName = [char]([int](65 + ($n % 26))) + $columnName
                        $n = [int][math]::Floor($n / 26)
                    }
                    $row[$columnName] = $values[$r, $c]
                }
                $rows += [pscustomobject]$row
            }
        }
        else {
            $rows += [pscustomobject]@{ Row = $range.Row; Value = $values }
        }
        return $rows
    }
    finally {
        Release-ComObject $range
    }
}

function Get-ExactMatches {
    param($Worksheet, [string]$Needle)
    $used = $null
    $after = $null
    $found = $null
    $matches = @()
    try {
        $used = $Worksheet.UsedRange
        $after = $used.Cells.Item($used.Cells.Count)
        $found = $used.Find($Needle, $after, -4163, 1, 1, 1, $false)
        if ($null -eq $found) { return @() }
        $firstAddress = $found.Address($false, $false)
        do {
            $address = $found.Address($false, $false)
            $matches += [pscustomobject]@{
                Sheet = $Worksheet.Name
                Address = $address
                Row = $found.Row
                Column = $found.Column
                Value = $found.Value2
            }
            $next = $used.Find($Needle, $found, -4163, 1, 1, 1, $false)
            Release-ComObject $found
            $found = $next
            if ($null -eq $found) { break }
        } while ($found.Address($false, $false) -ne $firstAddress)
        return $matches
    }
    finally {
        Release-ComObject $found
        Release-ComObject $after
        Release-ComObject $used
    }
}

function Get-Tables {
    param($Worksheet)
    $tables = @()
    $listObjects = $null
    try {
        $listObjects = $Worksheet.ListObjects
        for ($i = 1; $i -le $listObjects.Count; $i++) {
            $table = $null
            $range = $null
            try {
                $table = $listObjects.Item($i)
                $range = $table.Range
                $tables += [pscustomobject]@{
                    Name = $table.Name
                    Range = $range.Address($false, $false)
                }
            }
            finally {
                Release-ComObject $range
                Release-ComObject $table
            }
        }
        return $tables
    }
    finally {
        Release-ComObject $listObjects
    }
}

function Get-Charts {
    param($Worksheet)
    $result = @()
    $chartObjects = $null
    try {
        $chartObjects = $Worksheet.ChartObjects()
        for ($i = 1; $i -le $chartObjects.Count; $i++) {
            $chartObject = $null
            $chart = $null
            $seriesCollection = $null
            try {
                $chartObject = $chartObjects.Item($i)
                $chart = $chartObject.Chart
                $title = ''
                if ($chart.HasTitle) { $title = $chart.ChartTitle.Text }
                $series = @()
                $seriesCollection = $chart.SeriesCollection()
                for ($s = 1; $s -le $seriesCollection.Count; $s++) {
                    $oneSeries = $null
                    try {
                        $oneSeries = $seriesCollection.Item($s)
                        $series += [pscustomobject]@{
                            Index = $s
                            Name = $oneSeries.Name
                            Formula = $oneSeries.Formula
                            HasDataLabels = $oneSeries.HasDataLabels
                        }
                    }
                    finally {
                        Release-ComObject $oneSeries
                    }
                }
                $result += [pscustomobject]@{
                    Name = $chartObject.Name
                    Title = $title
                    ChartType = $chart.ChartType
                    Left = [math]::Round($chartObject.Left, 1)
                    Top = [math]::Round($chartObject.Top, 1)
                    Width = [math]::Round($chartObject.Width, 1)
                    Height = [math]::Round($chartObject.Height, 1)
                    SeriesCount = $seriesCollection.Count
                    Series = $series
                }
            }
            finally {
                Release-ComObject $seriesCollection
                Release-ComObject $chart
                Release-ComObject $chartObject
            }
        }
        return $result
    }
    finally {
        Release-ComObject $chartObjects
    }
}

function Get-Shapes {
    param($Worksheet)
    $result = @()
    $shapes = $null
    try {
        $shapes = $Worksheet.Shapes
        for ($i = 1; $i -le $shapes.Count; $i++) {
            $shape = $null
            try {
                $shape = $shapes.Item($i)
                $text = ''
                try {
                    if ($shape.TextFrame2.HasText) { $text = $shape.TextFrame2.TextRange.Text }
                }
                catch {}
                $result += [pscustomobject]@{
                    Name = $shape.Name
                    Type = $shape.Type
                    Left = [math]::Round($shape.Left, 1)
                    Top = [math]::Round($shape.Top, 1)
                    Width = [math]::Round($shape.Width, 1)
                    Height = [math]::Round($shape.Height, 1)
                    Text = $text
                }
            }
            finally {
                Release-ComObject $shape
            }
        }
        return $result
    }
    finally {
        Release-ComObject $shapes
    }
}

$excel = $null
$workbook = $null
$worksheets = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    try { $excel.AutomationSecurity = 3 } catch {}

    $workbook = $excel.Workbooks.Open($Path, 0, $true)
    $worksheets = $workbook.Worksheets
    $allPendiente = @()
    $sheetSummary = @()
    $chartCount = 0

    for ($i = 1; $i -le $worksheets.Count; $i++) {
        $sheet = $null
        $used = $null
        $chartObjects = $null
        try {
            $sheet = $worksheets.Item($i)
            $used = $sheet.UsedRange
            $chartObjects = $sheet.ChartObjects()
            $chartCount += $chartObjects.Count
            $sheetSummary += [pscustomobject]@{
                Index = $i
                Name = $sheet.Name
                UsedRange = $used.Address($false, $false)
                Charts = $chartObjects.Count
                Tables = (Get-Tables $sheet)
            }
            $allPendiente += Get-ExactMatches $sheet 'Pendiente'
        }
        finally {
            Release-ComObject $chartObjects
            Release-ComObject $used
            Release-ComObject $sheet
        }
    }

    $target = [ordered]@{}
    foreach ($sheetName in @('01_Dashboard', 'GRAFICAS', '04_Control_Publicos', '05_Matriz_Resultados', '08_Benignas_FP', '09_Benchmark_Plan', '12_Plan_Alertas', '15_Publicos_Definitivo', 'RESUMEN_EJECUTIVO', '11_Alertabilidad', '13_Hayabusa_Resumen', '17_Incoherencias_Cerradas', 'COMPARACION_VR_WAZUH', 'README', '99_Listas')) {
        $sheet = $null
        try {
            $sheet = $worksheets.Item($sheetName)
            $entry = [ordered]@{
                Charts = Get-Charts $sheet
                Tables = Get-Tables $sheet
            }
            switch ($sheetName) {
                '01_Dashboard' {
                    $entry['Cells'] = Get-RangeMatrix $sheet 'A1:Q70'
                    $entry['Shapes'] = Get-Shapes $sheet
                }
                'GRAFICAS' {
                    $entry['Cells_1_59'] = Get-RangeMatrix $sheet 'A1:R59'
                    $entry['Cells_60_80'] = Get-RangeMatrix $sheet 'A60:R80'
                }
                '04_Control_Publicos' { $entry['Cells'] = Get-RangeMatrix $sheet 'A4:AE67' }
                '05_Matriz_Resultados' { $entry['Cells'] = Get-RangeMatrix $sheet 'A4:T13' }
                '08_Benignas_FP' { $entry['Cells'] = Get-RangeMatrix $sheet 'A4:J18' }
                '15_Publicos_Definitivo' {
                    $entry['Cells'] = Get-RangeMatrix $sheet 'A1:S20'
                    $entry['Shapes'] = Get-Shapes $sheet
                }
                default {
                    $used = $null
                    try {
                        $used = $sheet.UsedRange
                        $entry['Cells'] = Get-RangeMatrix $sheet $used.Address($false, $false)
                    }
                    finally { Release-ComObject $used }
                }
            }
            $target[$sheetName] = [pscustomobject]$entry
        }
        finally {
            Release-ComObject $sheet
        }
    }

    $json = [pscustomobject]@{
        Path = $Path
        ReadOnly = $workbook.ReadOnly
        Worksheets = $worksheets.Count
        Charts = $chartCount
        PendienteCount = $allPendiente.Count
        Pendiente = $allPendiente
        SheetSummary = $sheetSummary
        Target = [pscustomobject]$target
    } | ConvertTo-Json -Depth 12
    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $json
    }
    else {
        [System.IO.File]::WriteAllText($OutputPath, $json, [System.Text.UTF8Encoding]::new($false))
        [pscustomobject]@{ OutputPath = $OutputPath; Bytes = (Get-Item -LiteralPath $OutputPath).Length } | ConvertTo-Json
    }
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
    }
    Release-ComObject $worksheets
    Release-ComObject $workbook
    Release-ComObject $excel
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
}
