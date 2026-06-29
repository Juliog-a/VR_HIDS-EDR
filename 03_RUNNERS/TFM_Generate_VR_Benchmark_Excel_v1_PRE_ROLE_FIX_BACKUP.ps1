<#
TFM Velociraptor resource benchmark Excel generator v1

Reads benchmark outputs produced by TFM_Benchmark_VR_Resource_Usage_v1.ps1 and
generates the workbook/reporting package for the performance-impact section.
#>

[CmdletBinding()]
param(
    [string]$HostBenchmarkDir = "C:\Users\julio\Desktop\TFM\05_LOGS\BENCHMARKS",
    [string]$VmBenchmarkDir = "C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas\Logs_Pruebas_TFM\BENCHMARKS",
    [string]$OutputXlsx = "C:\Users\julio\Desktop\TFM\VR_RESOURCE_BENCHMARK.xlsx",
    [string]$SummaryTxt = "C:\Users\julio\Desktop\TFM\VR_RESOURCE_BENCHMARK_RESUMEN.txt",
    [string]$ConclusionsMd = "C:\Users\julio\Desktop\TFM\VR_RESOURCE_BENCHMARK_CONCLUSIONES.md",
    [string]$ContextMd = "C:\Users\julio\Desktop\TFM\00_CONTEXT\VR_RESOURCE_BENCHMARK_CONTEXT.md",
    [string]$ReadmeMd = "C:\Users\julio\Desktop\TFM\05_LOGS\BENCHMARKS\README_BENCHMARKS.md",
    [string]$HashFile = "C:\Users\julio\Desktop\TFM\SHA256SUMS_VR_RESOURCE_BENCHMARK.txt"
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$script:Warnings = New-Object System.Collections.ArrayList
$script:ChartCount = 0

function Add-Warn {
    param([string]$Message)
    [void]$script:Warnings.Add($Message)
    Write-Host "[WARN] $Message"
}

function Write-Info {
    param([string]$Message)
    Write-Host "[INFO] $Message"
}

function Write-Ok {
    param([string]$Message)
    Write-Host "[OK] $Message"
}

function Escape-Xml {
    param([object]$Value)
    if ($null -eq $Value) { return "" }
    return [System.Security.SecurityElement]::Escape([string]$Value)
}

function ConvertTo-ColumnName {
    param([int]$Index)
    $name = ""
    $n = $Index
    while ($n -gt 0) {
        $rem = [int](($n - 1) % 26)
        $name = ([string][char]([int]65 + $rem)) + $name
        $n = [int][math]::Floor(($n - 1) / 26)
    }
    return $name
}

function Get-Number {
    param(
        [object]$Value,
        [double]$Default = 0
    )
    try {
        if ($null -eq $Value) { return $Default }
        if ($Value -is [double] -or $Value -is [int] -or $Value -is [long] -or $Value -is [decimal]) {
            return [double]$Value
        }
        $text = [string]$Value
        if ([string]::IsNullOrWhiteSpace($text)) { return $Default }
        return [double]::Parse($text, [System.Globalization.CultureInfo]::InvariantCulture)
    } catch {
        return $Default
    }
}

function Get-Prop {
    param(
        [object]$Object,
        [string[]]$Names,
        [object]$Default = ""
    )
    if ($null -eq $Object) { return $Default }
    foreach ($name in $Names) {
        $prop = $Object.PSObject.Properties[$name]
        if ($null -ne $prop) {
            return $prop.Value
        }
    }
    return $Default
}

function New-CellXml {
    param(
        [string]$Ref,
        [object]$Value,
        [bool]$Header = $false
    )

    if ($null -eq $Value) {
        return "<c r=""$Ref""/>"
    }

    if ($Value -is [bool]) {
        $v = if ($Value) { "1" } else { "0" }
        return "<c r=""$Ref"" t=""b""><v>$v</v></c>"
    }

    if ((-not $Header) -and ($Value -is [double] -or $Value -is [int] -or $Value -is [long] -or $Value -is [decimal])) {
        $v = ([double]$Value).ToString("0.############", [System.Globalization.CultureInfo]::InvariantCulture)
        return "<c r=""$Ref""><v>$v</v></c>"
    }

    $style = if ($Header) { ' s="1"' } else { "" }
    $text = Escape-Xml -Value $Value
    return "<c r=""$Ref"" t=""inlineStr""$style><is><t>$text</t></is></c>"
}

function New-WorksheetXml {
    param(
        [object[]]$Rows,
        [int[]]$ColumnWidths = @()
    )

    $maxCols = 1
    foreach ($row in $Rows) {
        $cells = @($row)
        if ($cells.Count -gt $maxCols) { $maxCols = $cells.Count }
    }
    $maxRows = [math]::Max(1, $Rows.Count)
    $dimension = "A1:{0}{1}" -f (ConvertTo-ColumnName -Index $maxCols), $maxRows

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.AppendLine('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    [void]$sb.AppendLine("<dimension ref=""$dimension""/>")
    [void]$sb.AppendLine('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')

    if ($ColumnWidths.Count -gt 0) {
        [void]$sb.AppendLine('<cols>')
        for ($i = 0; $i -lt $ColumnWidths.Count; $i++) {
            $col = $i + 1
            $width = [math]::Max(8, [int]$ColumnWidths[$i])
            [void]$sb.AppendLine("<col min=""$col"" max=""$col"" width=""$width"" customWidth=""1""/>")
        }
        [void]$sb.AppendLine('</cols>')
    }

    [void]$sb.AppendLine('<sheetData>')
    for ($r = 0; $r -lt $Rows.Count; $r++) {
        $rowNum = $r + 1
        [void]$sb.Append("<row r=""$rowNum"">")
        $cells = @($Rows[$r])
        for ($c = 0; $c -lt $cells.Count; $c++) {
            $ref = "{0}{1}" -f (ConvertTo-ColumnName -Index ($c + 1)), $rowNum
            [void]$sb.Append((New-CellXml -Ref $ref -Value $cells[$c] -Header:($r -eq 0)))
        }
        [void]$sb.AppendLine('</row>')
    }
    [void]$sb.AppendLine('</sheetData>')
    [void]$sb.AppendLine('</worksheet>')
    return $sb.ToString()
}

function Add-ZipText {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [string]$EntryName,
        [string]$Text
    )

    $entry = $Zip.CreateEntry($EntryName)
    $stream = $entry.Open()
    $encoding = New-Object System.Text.UTF8Encoding $false
    $writer = New-Object System.IO.StreamWriter($stream, $encoding)
    try {
        $writer.Write($Text)
    } finally {
        $writer.Dispose()
        $stream.Dispose()
    }
}

function New-XlsxOpenXml {
    param(
        [string]$Path,
        [object[]]$Sheets
    )

    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Force
    }
    $parent = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }

    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::CreateNew)
    $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $contentTypes = New-Object System.Text.StringBuilder
        [void]$contentTypes.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        [void]$contentTypes.AppendLine('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">')
        [void]$contentTypes.AppendLine('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>')
        [void]$contentTypes.AppendLine('<Default Extension="xml" ContentType="application/xml"/>')
        [void]$contentTypes.AppendLine('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>')
        [void]$contentTypes.AppendLine('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')

        for ($i = 0; $i -lt $Sheets.Count; $i++) {
            $sheetIndex = $i + 1
            [void]$contentTypes.AppendLine("<Override PartName=""/xl/worksheets/sheet$sheetIndex.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>")
        }
        [void]$contentTypes.AppendLine('</Types>')
        Add-ZipText -Zip $zip -EntryName "[Content_Types].xml" -Text $contentTypes.ToString()

        Add-ZipText -Zip $zip -EntryName "_rels/.rels" -Text @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
'@

        $workbook = New-Object System.Text.StringBuilder
        [void]$workbook.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        [void]$workbook.AppendLine('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
        [void]$workbook.AppendLine('<sheets>')
        for ($i = 0; $i -lt $Sheets.Count; $i++) {
            $sheetId = $i + 1
            $name = Escape-Xml -Value $Sheets[$i].Name
            [void]$workbook.AppendLine("<sheet name=""$name"" sheetId=""$sheetId"" r:id=""rId$sheetId""/>")
        }
        [void]$workbook.AppendLine('</sheets>')
        [void]$workbook.AppendLine('</workbook>')
        Add-ZipText -Zip $zip -EntryName "xl/workbook.xml" -Text $workbook.ToString()

        $rels = New-Object System.Text.StringBuilder
        [void]$rels.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        [void]$rels.AppendLine('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
        for ($i = 0; $i -lt $Sheets.Count; $i++) {
            $sheetId = $i + 1
            [void]$rels.AppendLine("<Relationship Id=""rId$sheetId"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet$sheetId.xml""/>")
        }
        $styleRel = $Sheets.Count + 1
        [void]$rels.AppendLine("<Relationship Id=""rId$styleRel"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/>")
        [void]$rels.AppendLine('</Relationships>')
        Add-ZipText -Zip $zip -EntryName "xl/_rels/workbook.xml.rels" -Text $rels.ToString()

        Add-ZipText -Zip $zip -EntryName "xl/styles.xml" -Text @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
'@

        for ($i = 0; $i -lt $Sheets.Count; $i++) {
            $sheetIndex = $i + 1
            $rows = @($Sheets[$i].Rows)
            $widths = @($Sheets[$i].Widths)
            $xml = New-WorksheetXml -Rows $rows -ColumnWidths $widths
            Add-ZipText -Zip $zip -EntryName "xl/worksheets/sheet$sheetIndex.xml" -Text $xml
        }
    } finally {
        $zip.Dispose()
        $fs.Dispose()
    }
}

function Test-XlsxPackage {
    param([string]$Path)

    $result = [pscustomobject]@{
        Opens      = $false
        SheetCount = 0
        ChartCount = 0
        Error      = ""
    }

    try {
        $fs = [System.IO.File]::OpenRead($Path)
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            $wbEntry = $zip.GetEntry("xl/workbook.xml")
            if ($null -eq $wbEntry) { throw "xl/workbook.xml no existe" }
            $reader = New-Object System.IO.StreamReader($wbEntry.Open())
            try {
                [xml]$wbXml = $reader.ReadToEnd()
            } finally {
                $reader.Dispose()
            }
            $result.SheetCount = @($wbXml.workbook.sheets.sheet).Count
            $result.ChartCount = @($zip.Entries | Where-Object { $_.FullName -like "xl/charts/chart*.xml" }).Count
            $result.Opens = $true
        } finally {
            $zip.Dispose()
            $fs.Dispose()
        }
    } catch {
        $result.Error = $_.Exception.Message
    }

    return $result
}

function Try-AddExcelComCharts {
    param(
        [string]$Path,
        [bool]$HasData,
        [int]$ScenarioCount
    )

    if (-not $HasData -or $ScenarioCount -lt 1) {
        return 0
    }

    $excel = $null
    $workbook = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($Path)
        $ws = $workbook.Worksheets.Item("GRAFICAS")
        $last = 3 + $ScenarioCount

        $xlColumnClustered = 51
        $xlLine = 4
        $created = 0

        function Add-ComChart {
            param(
                [object]$Worksheet,
                [string]$RangeAddress,
                [string]$Title,
                [int]$ChartType,
                [double]$Left,
                [double]$Top
            )
            $chartObj = $Worksheet.ChartObjects().Add($Left, $Top, 420, 250)
            $chart = $chartObj.Chart
            $chart.ChartType = $ChartType
            $chart.SetSourceData($Worksheet.Range($RangeAddress))
            $chart.HasTitle = $true
            $chart.ChartTitle.Text = $Title
            return $chartObj
        }

        [void](Add-ComChart -Worksheet $ws -RangeAddress ("A3:D$last") -Title "VR CPU media/max/P95 por escenario" -ChartType $xlColumnClustered -Left 20 -Top 260)
        $created++
        [void](Add-ComChart -Worksheet $ws -RangeAddress ("F3:I$last") -Title "VR RAM media/max/P95 por escenario" -ChartType $xlColumnClustered -Left 470 -Top 260)
        $created++
        [void](Add-ComChart -Worksheet $ws -RangeAddress ("K3:M$last") -Title "VR IO media/max por escenario" -ChartType $xlColumnClustered -Left 920 -Top 260)
        $created++
        [void](Add-ComChart -Worksheet $ws -RangeAddress ("O3:Q$last") -Title "CPU sistema media/max por escenario" -ChartType $xlColumnClustered -Left 20 -Top 540)
        $created++
        [void](Add-ComChart -Worksheet $ws -RangeAddress ("S3:U$last") -Title "RAM sistema media/max por escenario" -ChartType $xlColumnClustered -Left 470 -Top 540)
        $created++

        try {
            $compWs = $workbook.Worksheets.Item("COMPARATIVA")
            $compRows = $compWs.UsedRange.Rows.Count
            if ($compRows -gt 1) {
                $compChart = $ws.ChartObjects().Add(920, 820, 420, 250)
                $compChart.Chart.ChartType = $xlColumnClustered
                $compChart.Chart.SetSourceData($compWs.Range("A1:C$compRows"))
                $compChart.Chart.HasTitle = $true
                $compChart.Chart.ChartTitle.Text = "Comparativa incremental CPU/RAM/IO"
                $created++
            }
        } catch {
            Add-Warn "No se pudo crear grafica de comparativa incremental: $($_.Exception.Message)"
        }

        try {
            $used = $workbook.Worksheets.Item("RAW_SAMPLES").UsedRange
            if ($used.Rows.Count -gt 1) {
                $rawWs = $workbook.Worksheets.Item("RAW_SAMPLES")
                $rawRangeLast = [math]::Min($used.Rows.Count, 250)

                $cpuChart = $ws.ChartObjects().Add(20, 820, 420, 250)
                $cpuChart.Chart.ChartType = $xlLine
                $series = $cpuChart.Chart.SeriesCollection().NewSeries()
                $series.Name = "VR_CPU_Percent"
                $series.XValues = $rawWs.Range("D2:D$rawRangeLast")
                $series.Values = $rawWs.Range("H2:H$rawRangeLast")
                $cpuChart.Chart.HasTitle = $true
                $cpuChart.Chart.ChartTitle.Text = "VR CPU % en el tiempo"
                $created++

                $ramChart = $ws.ChartObjects().Add(470, 820, 420, 250)
                $ramChart.Chart.ChartType = $xlLine
                $ramSeries = $ramChart.Chart.SeriesCollection().NewSeries()
                $ramSeries.Name = "VR_RAM_MB"
                $ramSeries.XValues = $rawWs.Range("D2:D$rawRangeLast")
                $ramSeries.Values = $rawWs.Range("I2:I$rawRangeLast")
                $ramChart.Chart.HasTitle = $true
                $ramChart.Chart.ChartTitle.Text = "VR RAM MB en el tiempo"
                $created++
            }
        } catch {
            Add-Warn "No se pudo crear grafica lineal desde RAW_SAMPLES: $($_.Exception.Message)"
        }

        $workbook.Save()
        $workbook.Close($true)
        $excel.Quit()

        return $created
    } catch {
        Add-Warn "Excel COM no disponible o fallo al crear graficas reales: $($_.Exception.Message)"
        try { if ($workbook) { $workbook.Close($false) } } catch {}
        try { if ($excel) { $excel.Quit() } } catch {}
        return 0
    } finally {
        try {
            if ($workbook) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) }
            if ($excel) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
        } catch {
        }
    }
}

function Normalize-Summary {
    param(
        [object]$Summary,
        [string]$SourceFile
    )

    $scenario = [string](Get-Prop -Object $Summary -Names @("Scenario", "ScenarioName") -Default "UNKNOWN")
    $samples = [int](Get-Number -Value (Get-Prop -Object $Summary -Names @("Samples") -Default 0))
    $vrMaxCount = [int](Get-Number -Value (Get-Prop -Object $Summary -Names @("VRProcessCountMax", "VR_ProcessCountMax") -Default 0))
    $runnerStarted = [string](Get-Prop -Object $Summary -Names @("RunnerStarted") -Default "")
    $quality = "OK"

    if ($samples -lt 1) {
        $quality = "INVALID_NO_SAMPLES"
    } elseif ($scenario -eq "BASELINE_NO_VR" -and $vrMaxCount -gt 0) {
        $quality = "INVALID_BASELINE_VR_PRESENT"
    } elseif ($scenario -eq "VR_IDLE" -and $vrMaxCount -eq 0) {
        $quality = "INVALID_IDLE_NO_VR"
    } elseif ($scenario -eq "VR_TEC_RUNNER" -and ($vrMaxCount -eq 0 -or $runnerStarted -ne "True")) {
        $quality = "INVALID_RUNNER_OR_NO_VR"
    }

    return [pscustomobject]@{
        Scenario                   = $scenario
        RunId                      = [string](Get-Prop -Object $Summary -Names @("RunId") -Default "")
        DataQuality                = $quality
        Samples                    = $samples
        DurationSec                = Get-Number -Value (Get-Prop -Object $Summary -Names @("DurationSec") -Default 0)
        ActualDurationSec          = Get-Number -Value (Get-Prop -Object $Summary -Names @("ActualDurationSec") -Default 0)
        ServiceStatusStart         = [string](Get-Prop -Object $Summary -Names @("ServiceStatusStart") -Default "")
        ServiceStatusEnd           = [string](Get-Prop -Object $Summary -Names @("ServiceStatusEnd") -Default "")
        RunnerStarted              = $runnerStarted
        VRProcessCountAvg          = Get-Number -Value (Get-Prop -Object $Summary -Names @("VRProcessCountAvg") -Default 0)
        VRProcessCountMax          = $vrMaxCount
        Avg_VR_CPU_Percent         = Get-Number -Value (Get-Prop -Object $Summary -Names @("Avg_VR_CPU_Percent") -Default 0)
        Max_VR_CPU_Percent         = Get-Number -Value (Get-Prop -Object $Summary -Names @("Max_VR_CPU_Percent") -Default 0)
        P95_VR_CPU_Percent         = Get-Number -Value (Get-Prop -Object $Summary -Names @("P95_VR_CPU_Percent") -Default 0)
        Avg_VR_RAM_MB              = Get-Number -Value (Get-Prop -Object $Summary -Names @("Avg_VR_RAM_MB") -Default 0)
        Max_VR_RAM_MB              = Get-Number -Value (Get-Prop -Object $Summary -Names @("Max_VR_RAM_MB") -Default 0)
        P95_VR_RAM_MB              = Get-Number -Value (Get-Prop -Object $Summary -Names @("P95_VR_RAM_MB") -Default 0)
        Avg_VR_IO_BytesSec         = Get-Number -Value (Get-Prop -Object $Summary -Names @("Avg_VR_IO_BytesSec") -Default 0)
        Max_VR_IO_BytesSec         = Get-Number -Value (Get-Prop -Object $Summary -Names @("Max_VR_IO_BytesSec") -Default 0)
        Avg_System_CPU_Percent     = Get-Number -Value (Get-Prop -Object $Summary -Names @("Avg_System_CPU_Percent") -Default 0)
        Max_System_CPU_Percent     = Get-Number -Value (Get-Prop -Object $Summary -Names @("Max_System_CPU_Percent") -Default 0)
        Avg_System_RAM_Used_MB     = Get-Number -Value (Get-Prop -Object $Summary -Names @("Avg_System_RAM_Used_MB") -Default 0)
        Max_System_RAM_Used_MB     = Get-Number -Value (Get-Prop -Object $Summary -Names @("Max_System_RAM_Used_MB") -Default 0)
        Avg_Disk_BytesSec          = Get-Number -Value (Get-Prop -Object $Summary -Names @("Avg_Disk_BytesSec") -Default 0)
        Max_Disk_BytesSec          = Get-Number -Value (Get-Prop -Object $Summary -Names @("Max_Disk_BytesSec") -Default 0)
        Avg_Net_BytesSec           = Get-Number -Value (Get-Prop -Object $Summary -Names @("Avg_Net_BytesSec") -Default 0)
        Max_Net_BytesSec           = Get-Number -Value (Get-Prop -Object $Summary -Names @("Max_Net_BytesSec") -Default 0)
        StartTime                  = [string](Get-Prop -Object $Summary -Names @("StartTime") -Default "")
        EndTime                    = [string](Get-Prop -Object $Summary -Names @("EndTime") -Default "")
        SourceFile                 = $SourceFile
    }
}

function Get-LatestByScenario {
    param([object[]]$Rows)
    $selected = New-Object System.Collections.ArrayList
    foreach ($scenario in @("BASELINE_NO_VR", "VR_IDLE", "VR_TEC_RUNNER")) {
        $row = @($Rows | Where-Object { $_.Scenario -eq $scenario } | Sort-Object -Property StartTime -Descending | Select-Object -First 1)
        if ($row.Count -gt 0) {
            [void]$selected.Add($row[0])
        }
    }
    return @($selected)
}

function New-DeltaRow {
    param(
        [string]$Metric,
        [object]$A,
        [object]$B,
        [string]$Comparison
    )

    if ($null -eq $A -or $null -eq $B) {
        return @($Comparison, $Metric, "PENDIENTE", "", "", "")
    }
    $aVal = Get-Number -Value $A.$Metric
    $bVal = Get-Number -Value $B.$Metric
    $delta = [math]::Round(($bVal - $aVal), 2)
    $ratio = ""
    if ([math]::Abs($aVal) -gt 0.000001) {
        $ratio = [math]::Round(($bVal / $aVal), 2)
    }
    return @($Comparison, $Metric, $delta, $aVal, $bVal, $ratio)
}

$projectRoot = Split-Path -Parent $PSScriptRoot
$benchmarkScript = Join-Path $PSScriptRoot "TFM_Benchmark_VR_Resource_Usage_v1.ps1"

foreach ($dir in @((Split-Path -Parent $OutputXlsx), (Split-Path -Parent $SummaryTxt), (Split-Path -Parent $ConclusionsMd), (Split-Path -Parent $ContextMd), (Split-Path -Parent $ReadmeMd), (Split-Path -Parent $HashFile), $HostBenchmarkDir)) {
    if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

$sourceDirs = New-Object System.Collections.ArrayList
foreach ($dir in @($HostBenchmarkDir, $VmBenchmarkDir)) {
    if (Test-Path -LiteralPath $dir) {
        if (-not ($sourceDirs -contains $dir)) {
            [void]$sourceDirs.Add($dir)
        }
    } else {
        Add-Warn "Ruta de benchmark no disponible: $dir"
    }
}

$summaryRows = New-Object System.Collections.ArrayList
$rawRows = New-Object System.Collections.ArrayList
$processRows = New-Object System.Collections.ArrayList
$usedFiles = New-Object System.Collections.ArrayList

foreach ($dir in $sourceDirs) {
    $summaryFiles = @(Get-ChildItem -LiteralPath $dir -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*summary*.json" })
    foreach ($file in $summaryFiles) {
        try {
            $json = Get-Content -LiteralPath $file.FullName -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            $norm = Normalize-Summary -Summary $json -SourceFile $file.FullName
            [void]$summaryRows.Add($norm)
            [void]$usedFiles.Add($file.FullName)

            $runDir = Split-Path -Parent $file.FullName
            $sampleFiles = @(Get-ChildItem -LiteralPath $runDir -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq "samples.csv" -or $_.Name -like "*_samples.csv" })
            foreach ($sampleFile in $sampleFiles) {
                try {
                    $items = @(Import-Csv -LiteralPath $sampleFile.FullName -ErrorAction Stop)
                    foreach ($item in $items) {
                        $item | Add-Member -NotePropertyName SourceFile -NotePropertyValue $sampleFile.FullName -Force
                        [void]$rawRows.Add($item)
                    }
                    [void]$usedFiles.Add($sampleFile.FullName)
                } catch {
                    Add-Warn "No se pudo importar samples CSV '$($sampleFile.FullName)': $($_.Exception.Message)"
                }
            }

            $procFiles = @(Get-ChildItem -LiteralPath $runDir -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq "process_samples.csv" -or $_.Name -like "*process_samples.csv" })
            foreach ($procFile in $procFiles) {
                try {
                    $items = @(Import-Csv -LiteralPath $procFile.FullName -ErrorAction Stop)
                    foreach ($item in $items) {
                        $item | Add-Member -NotePropertyName SourceFile -NotePropertyValue $procFile.FullName -Force
                        [void]$processRows.Add($item)
                    }
                    [void]$usedFiles.Add($procFile.FullName)
                } catch {
                    Add-Warn "No se pudo importar process_samples CSV '$($procFile.FullName)': $($_.Exception.Message)"
                }
            }
        } catch {
            Add-Warn "No se pudo leer summary JSON '$($file.FullName)': $($_.Exception.Message)"
        }
    }
}

$hasData = ($summaryRows.Count -gt 0)
$latestRows = @(Get-LatestByScenario -Rows @($summaryRows))

$baseline = @($latestRows | Where-Object { $_.Scenario -eq "BASELINE_NO_VR" } | Select-Object -First 1)
$idle = @($latestRows | Where-Object { $_.Scenario -eq "VR_IDLE" } | Select-Object -First 1)
$runner = @($latestRows | Where-Object { $_.Scenario -eq "VR_TEC_RUNNER" } | Select-Object -First 1)
$baselineObj = if ($baseline.Count -gt 0) { $baseline[0] } else { $null }
$idleObj = if ($idle.Count -gt 0) { $idle[0] } else { $null }
$runnerObj = if ($runner.Count -gt 0) { $runner[0] } else { $null }

$executiveRows = New-Object System.Collections.ArrayList
[void]$executiveRows.Add(@("Metrica", "Valor", "Origen", "Nota"))
if (-not $hasData) {
    [void]$executiveRows.Add(@("Estado", "PENDIENTE_DE_EJECUCION", "LOCAL", "No hay summary.json de benchmark en rutas revisadas."))
    [void]$executiveRows.Add(@("Escenarios requeridos", 3, "REQUISITO", "BASELINE_NO_VR, VR_IDLE, VR_TEC_RUNNER"))
    [void]$executiveRows.Add(@("Excel", "PLANTILLA", "GENERADO", "Sin metricas reales ni graficas reales hasta ejecutar benchmark."))
} else {
    [void]$executiveRows.Add(@("Escenarios con datos", $latestRows.Count, "LOG", "Ultimo RunId por escenario."))
    [void]$executiveRows.Add(@("Samples totales importadas", $rawRows.Count, "LOG", "RAW_SAMPLES normalizado."))
    [void]$executiveRows.Add(@("Process samples importadas", $processRows.Count, "LOG", "PROCESS_DETAIL normalizado."))
    if ($idleObj) {
        [void]$executiveRows.Add(@("CPU VR media idle", $idleObj.Avg_VR_CPU_Percent, "LOG", "Criterio bajo si < 2 %."))
        [void]$executiveRows.Add(@("RAM VR media idle MB", $idleObj.Avg_VR_RAM_MB, "LOG", "Consumo medio en reposo."))
    }
    if ($runnerObj) {
        [void]$executiveRows.Add(@("CPU VR maxima runner", $runnerObj.Max_VR_CPU_Percent, "LOG", "Pico maximo durante runner TEC."))
        [void]$executiveRows.Add(@("RAM VR maxima runner MB", $runnerObj.Max_VR_RAM_MB, "LOG", "Pico maximo durante runner TEC."))
    }
}

$summaryHeader = @(
    "Scenario", "RunId", "DataQuality", "Samples", "DurationSec", "ActualDurationSec",
    "ServiceStatusStart", "ServiceStatusEnd", "RunnerStarted", "VRProcessCountAvg",
    "VRProcessCountMax", "Avg_VR_CPU_Percent", "Max_VR_CPU_Percent",
    "P95_VR_CPU_Percent", "Avg_VR_RAM_MB", "Max_VR_RAM_MB", "P95_VR_RAM_MB",
    "Avg_VR_IO_BytesSec", "Max_VR_IO_BytesSec", "Avg_System_CPU_Percent",
    "Max_System_CPU_Percent", "Avg_System_RAM_Used_MB", "Max_System_RAM_Used_MB",
    "Avg_Disk_BytesSec", "Max_Disk_BytesSec", "Avg_Net_BytesSec",
    "Max_Net_BytesSec", "SourceFile"
)

$scenarioSummaryRows = New-Object System.Collections.ArrayList
[void]$scenarioSummaryRows.Add($summaryHeader)
if ($summaryRows.Count -eq 0) {
    [void]$scenarioSummaryRows.Add(@("PENDIENTE", "", "PENDIENTE_EJECUCION", 0, 0, 0, "", "", "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, ""))
} else {
    foreach ($row in @($summaryRows | Sort-Object Scenario, StartTime)) {
        [void]$scenarioSummaryRows.Add(@(
            $row.Scenario, $row.RunId, $row.DataQuality, $row.Samples, $row.DurationSec,
            $row.ActualDurationSec, $row.ServiceStatusStart, $row.ServiceStatusEnd,
            $row.RunnerStarted, $row.VRProcessCountAvg, $row.VRProcessCountMax,
            $row.Avg_VR_CPU_Percent, $row.Max_VR_CPU_Percent, $row.P95_VR_CPU_Percent,
            $row.Avg_VR_RAM_MB, $row.Max_VR_RAM_MB, $row.P95_VR_RAM_MB,
            $row.Avg_VR_IO_BytesSec, $row.Max_VR_IO_BytesSec,
            $row.Avg_System_CPU_Percent, $row.Max_System_CPU_Percent,
            $row.Avg_System_RAM_Used_MB, $row.Max_System_RAM_Used_MB,
            $row.Avg_Disk_BytesSec, $row.Max_Disk_BytesSec, $row.Avg_Net_BytesSec,
            $row.Max_Net_BytesSec, $row.SourceFile
        ))
    }
}

$rawHeader = @(
    "RunId", "Scenario", "Timestamp", "ElapsedSec", "ServiceName", "ServiceStatus",
    "VR_ProcessCount", "VR_CPU_Percent", "VR_RAM_MB", "VR_WorkingSet_MB",
    "VR_PrivateMemory_MB", "VR_IO_Total_BytesSec", "System_CPU_Percent",
    "System_RAM_Used_MB", "System_RAM_Free_MB", "System_Disk_BytesSec",
    "System_Net_BytesSec", "RunnerActive", "SourceFile"
)
$rawSheetRows = New-Object System.Collections.ArrayList
[void]$rawSheetRows.Add($rawHeader)
if ($rawRows.Count -eq 0) {
    [void]$rawSheetRows.Add(@("PENDIENTE", "", "", 0, "", "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", ""))
} else {
    foreach ($row in @($rawRows | Select-Object -First 5000)) {
        [void]$rawSheetRows.Add(@(
            (Get-Prop $row @("RunId") ""),
            (Get-Prop $row @("Scenario") ""),
            (Get-Prop $row @("Timestamp", "Time") ""),
            (Get-Number (Get-Prop $row @("ElapsedSec") 0)),
            (Get-Prop $row @("ServiceName") ""),
            (Get-Prop $row @("ServiceStatus") ""),
            (Get-Number (Get-Prop $row @("VR_ProcessCount") 0)),
            (Get-Number (Get-Prop $row @("VR_CPU_Percent") 0)),
            (Get-Number (Get-Prop $row @("VR_RAM_MB") 0)),
            (Get-Number (Get-Prop $row @("VR_WorkingSet_MB") 0)),
            (Get-Number (Get-Prop $row @("VR_PrivateMemory_MB") 0)),
            (Get-Number (Get-Prop $row @("VR_IO_Total_BytesSec", "VR_IO_BytesSec") 0)),
            (Get-Number (Get-Prop $row @("System_CPU_Percent") 0)),
            (Get-Number (Get-Prop $row @("System_RAM_Used_MB") 0)),
            (Get-Number (Get-Prop $row @("System_RAM_Free_MB") 0)),
            (Get-Number (Get-Prop $row @("System_Disk_BytesSec") 0)),
            (Get-Number (Get-Prop $row @("System_Net_BytesSec") 0)),
            (Get-Prop $row @("RunnerActive", "RunnerRunning") ""),
            (Get-Prop $row @("SourceFile") "")
        ))
    }
}

$processHeader = @(
    "RunId", "Scenario", "Timestamp", "ElapsedSec", "ServiceStatus", "PID",
    "ProcessName", "ExecutablePath", "CommandLine", "CPU_Percent", "RAM_MB",
    "WorkingSet_MB", "PrivateMemory_MB", "Handles", "Threads", "IO_Read_BytesSec",
    "IO_Write_BytesSec", "IO_Total_BytesSec", "RunnerActive", "SourceFile"
)
$processSheetRows = New-Object System.Collections.ArrayList
[void]$processSheetRows.Add($processHeader)
if ($processRows.Count -eq 0) {
    [void]$processSheetRows.Add(@("PENDIENTE", "", "", 0, "", "", "", "", "", 0, 0, 0, 0, 0, 0, 0, 0, 0, "", ""))
} else {
    foreach ($row in @($processRows | Select-Object -First 5000)) {
        [void]$processSheetRows.Add(@(
            (Get-Prop $row @("RunId") ""),
            (Get-Prop $row @("Scenario") ""),
            (Get-Prop $row @("Timestamp", "Time") ""),
            (Get-Number (Get-Prop $row @("ElapsedSec") 0)),
            (Get-Prop $row @("ServiceStatus") ""),
            (Get-Prop $row @("PID", "ProcessId") ""),
            (Get-Prop $row @("ProcessName", "Name") ""),
            (Get-Prop $row @("ExecutablePath") ""),
            (Get-Prop $row @("CommandLine") ""),
            (Get-Number (Get-Prop $row @("CPU_Percent") 0)),
            (Get-Number (Get-Prop $row @("RAM_MB") 0)),
            (Get-Number (Get-Prop $row @("WorkingSet_MB") 0)),
            (Get-Number (Get-Prop $row @("PrivateMemory_MB") 0)),
            (Get-Number (Get-Prop $row @("Handles") 0)),
            (Get-Number (Get-Prop $row @("Threads") 0)),
            (Get-Number (Get-Prop $row @("IO_Read_BytesSec") 0)),
            (Get-Number (Get-Prop $row @("IO_Write_BytesSec") 0)),
            (Get-Number (Get-Prop $row @("IO_Total_BytesSec") 0)),
            (Get-Prop $row @("RunnerActive") ""),
            (Get-Prop $row @("SourceFile") "")
        ))
    }
}

$comparisonRows = New-Object System.Collections.ArrayList
[void]$comparisonRows.Add(@("Comparacion", "Metrica", "Delta", "Valor_A", "Valor_B", "Ratio_B_sobre_A"))
$metricsForDelta = @(
    "Avg_VR_CPU_Percent", "Avg_VR_RAM_MB", "Avg_VR_IO_BytesSec",
    "Avg_System_CPU_Percent", "Avg_System_RAM_Used_MB"
)
foreach ($metric in $metricsForDelta) {
    [void]$comparisonRows.Add((New-DeltaRow -Comparison "VR_IDLE_vs_BASELINE_NO_VR" -Metric $metric -A $baselineObj -B $idleObj))
}
foreach ($metric in $metricsForDelta) {
    [void]$comparisonRows.Add((New-DeltaRow -Comparison "VR_TEC_RUNNER_vs_VR_IDLE" -Metric $metric -A $idleObj -B $runnerObj))
}
foreach ($metric in $metricsForDelta) {
    [void]$comparisonRows.Add((New-DeltaRow -Comparison "VR_TEC_RUNNER_vs_BASELINE_NO_VR" -Metric $metric -A $baselineObj -B $runnerObj))
}

$metricsRows = New-Object System.Collections.ArrayList
[void]$metricsRows.Add(@("Metrica", "Valor", "Origen", "Interpretacion"))
if (-not $hasData) {
    [void]$metricsRows.Add(@("Consumo incremental CPU por activar Velociraptor", "PENDIENTE", "SIN_LOG", "Requiere BASELINE_NO_VR y VR_IDLE."))
    [void]$metricsRows.Add(@("Consumo incremental RAM por activar Velociraptor", "PENDIENTE", "SIN_LOG", "Requiere BASELINE_NO_VR y VR_IDLE."))
    [void]$metricsRows.Add(@("Consumo adicional CPU durante runner", "PENDIENTE", "SIN_LOG", "Requiere VR_IDLE y VR_TEC_RUNNER."))
    [void]$metricsRows.Add(@("Consumo adicional RAM durante runner", "PENDIENTE", "SIN_LOG", "Requiere VR_IDLE y VR_TEC_RUNNER."))
} else {
    $idleCpuDelta = if ($baselineObj -and $idleObj) { [math]::Round($idleObj.Avg_VR_CPU_Percent - $baselineObj.Avg_VR_CPU_Percent, 2) } else { "PENDIENTE" }
    $idleRamDelta = if ($baselineObj -and $idleObj) { [math]::Round($idleObj.Avg_VR_RAM_MB - $baselineObj.Avg_VR_RAM_MB, 2) } else { "PENDIENTE" }
    $runCpuDelta = if ($idleObj -and $runnerObj) { [math]::Round($runnerObj.Avg_VR_CPU_Percent - $idleObj.Avg_VR_CPU_Percent, 2) } else { "PENDIENTE" }
    $runRamDelta = if ($idleObj -and $runnerObj) { [math]::Round($runnerObj.Avg_VR_RAM_MB - $idleObj.Avg_VR_RAM_MB, 2) } else { "PENDIENTE" }
    [void]$metricsRows.Add(@("Consumo incremental CPU por activar Velociraptor", $idleCpuDelta, "LOG", "Avg_VR_CPU_IDLE - Avg_VR_CPU_BASELINE"))
    [void]$metricsRows.Add(@("Consumo incremental RAM por activar Velociraptor", $idleRamDelta, "LOG", "Avg_VR_RAM_IDLE - Avg_VR_RAM_BASELINE"))
    [void]$metricsRows.Add(@("Consumo adicional CPU durante runner", $runCpuDelta, "LOG", "Avg_VR_CPU_RUNNER - Avg_VR_CPU_IDLE"))
    [void]$metricsRows.Add(@("Consumo adicional RAM durante runner", $runRamDelta, "LOG", "Avg_VR_RAM_RUNNER - Avg_VR_RAM_IDLE"))
    if ($runnerObj) {
        [void]$metricsRows.Add(@("Pico maximo CPU VR durante runner", $runnerObj.Max_VR_CPU_Percent, "LOG", "Max_VR_CPU_Percent en VR_TEC_RUNNER"))
        [void]$metricsRows.Add(@("Pico maximo RAM VR durante runner", $runnerObj.Max_VR_RAM_MB, "LOG", "Max_VR_RAM_MB en VR_TEC_RUNNER"))
    }
}

$graphRows = New-Object System.Collections.ArrayList
[void]$graphRows.Add(@("GRAFICAS", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""))
[void]$graphRows.Add(@("Estado", $(if ($hasData) { "Datos disponibles; Excel COM intentara crear graficas reales." } else { "PENDIENTE_EJECUCION: no hay datos reales para graficar." }), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""))
[void]$graphRows.Add(@("Scenario", "Avg_CPU", "Max_CPU", "P95_CPU", "", "Scenario", "Avg_RAM", "Max_RAM", "P95_RAM", "", "Scenario", "Avg_IO", "Max_IO", "", "Scenario", "Avg_SYS_CPU", "Max_SYS_CPU", "", "Scenario", "Avg_SYS_RAM", "Max_SYS_RAM"))
if ($latestRows.Count -eq 0) {
    [void]$graphRows.Add(@("PENDIENTE", 0, 0, 0, "", "PENDIENTE", 0, 0, 0, "", "PENDIENTE", 0, 0, "", "PENDIENTE", 0, 0, "", "PENDIENTE", 0, 0))
} else {
    foreach ($row in $latestRows) {
        [void]$graphRows.Add(@(
            $row.Scenario, $row.Avg_VR_CPU_Percent, $row.Max_VR_CPU_Percent, $row.P95_VR_CPU_Percent,
            "", $row.Scenario, $row.Avg_VR_RAM_MB, $row.Max_VR_RAM_MB, $row.P95_VR_RAM_MB,
            "", $row.Scenario, $row.Avg_VR_IO_BytesSec, $row.Max_VR_IO_BytesSec,
            "", $row.Scenario, $row.Avg_System_CPU_Percent, $row.Max_System_CPU_Percent,
            "", $row.Scenario, $row.Avg_System_RAM_Used_MB, $row.Max_System_RAM_Used_MB
        ))
    }
}

$conclusionRows = New-Object System.Collections.ArrayList
[void]$conclusionRows.Add(@("Apartado", "Conclusion", "Base"))
if (-not $hasData) {
    [void]$conclusionRows.Add(@("Estado", "PENDIENTE DE EJECUCION. No hay datos reales de benchmark.", "SIN_LOG"))
    [void]$conclusionRows.Add(@("Uso en memoria", "No usar metricas de rendimiento hasta ejecutar los tres escenarios y copiar resultados.", "CRITERIO"))
} else {
    $idleImpact = "PENDIENTE"
    if ($idleObj) {
        $idleImpact = if ($idleObj.Avg_VR_CPU_Percent -lt 2) { "Impacto CPU idle bajo segun criterio < 2 %." } else { "Impacto CPU idle no bajo segun criterio < 2 %; revisar." }
    }
    [void]$conclusionRows.Add(@("Impacto en reposo", $idleImpact, "LOG"))
    if ($runnerObj) {
        $runnerImpact = if ($runnerObj.Avg_VR_CPU_Percent -lt 10) { "Impacto CPU durante runner bajo/moderado segun criterio < 10 %." } else { "Impacto CPU durante runner alto segun criterio >= 10 %; revisar." }
        [void]$conclusionRows.Add(@("Impacto durante runner", $runnerImpact, "LOG"))
    }
    [void]$conclusionRows.Add(@("Limitacion", "Laboratorio controlado; interpretar como benchmark reproducible, no como estadistica de produccion.", "CRITERIO"))
}

$traceRows = New-Object System.Collections.ArrayList
[void]$traceRows.Add(@("Tipo", "Ruta", "SHA256", "Nota"))
foreach ($dir in @($HostBenchmarkDir, $VmBenchmarkDir)) {
    [void]$traceRows.Add(@("Ruta revisada", $dir, "", $(if (Test-Path -LiteralPath $dir) { "Disponible" } else { "No disponible" })))
}
foreach ($file in @($usedFiles | Select-Object -Unique)) {
    $hash = ""
    try { $hash = (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash } catch { $hash = "" }
    [void]$traceRows.Add(@("Archivo procesado", $file, $hash, "LOG"))
}
if ($usedFiles.Count -eq 0) {
    [void]$traceRows.Add(@("Archivo procesado", "Ninguno", "", "PENDIENTE_EJECUCION"))
}
foreach ($warning in $script:Warnings) {
    [void]$traceRows.Add(@("WARN", "", "", $warning))
}

$sheets = @(
    [pscustomobject]@{ Name = "RESUMEN_EJECUTIVO"; Rows = @($executiveRows); Widths = @(34, 30, 18, 80) },
    [pscustomobject]@{ Name = "RAW_SAMPLES"; Rows = @($rawSheetRows); Widths = @(26, 20, 28, 12, 20, 18, 16, 16, 16, 18, 20, 20, 18, 20, 20, 22, 22, 14, 80) },
    [pscustomobject]@{ Name = "SCENARIO_SUMMARY"; Rows = @($scenarioSummaryRows); Widths = @(20, 32, 28, 12, 12, 16, 18, 18, 14, 18, 18, 18, 18, 18, 16, 16, 16, 18, 18, 20, 20, 22, 22, 18, 18, 18, 18, 80) },
    [pscustomobject]@{ Name = "PROCESS_DETAIL"; Rows = @($processSheetRows); Widths = @(26, 20, 28, 12, 18, 12, 20, 60, 80, 14, 12, 14, 18, 12, 12, 18, 18, 18, 14, 80) },
    [pscustomobject]@{ Name = "COMPARATIVA"; Rows = @($comparisonRows); Widths = @(34, 28, 14, 14, 14, 18) },
    [pscustomobject]@{ Name = "GRAFICAS"; Rows = @($graphRows); Widths = @(22, 14, 14, 14, 4, 22, 14, 14, 14, 4, 22, 14, 14, 4, 22, 16, 16, 4, 22, 16, 16) },
    [pscustomobject]@{ Name = "CONCLUSIONES"; Rows = @($conclusionRows); Widths = @(28, 90, 18) },
    [pscustomobject]@{ Name = "TRAZABILIDAD"; Rows = @($traceRows); Widths = @(24, 90, 68, 60) }
)

New-XlsxOpenXml -Path $OutputXlsx -Sheets $sheets
$script:ChartCount = Try-AddExcelComCharts -Path $OutputXlsx -HasData:$hasData -ScenarioCount $latestRows.Count
$xlsxCheck = Test-XlsxPackage -Path $OutputXlsx

if (-not $xlsxCheck.Opens) {
    throw "El XLSX generado no es valido: $($xlsxCheck.Error)"
}

$statusText = if ($hasData) { "DATOS_IMPORTADOS" } else { "PENDIENTE_DE_EJECUCION" }
$chartNote = if ($script:ChartCount -gt 0) { "Graficas reales creadas mediante Excel COM: $script:ChartCount." } elseif ($hasData) { "Graficas reales no creadas; workbook contiene tablas normalizadas y helper de graficas." } else { "Sin graficas reales porque no hay datos de benchmark." }

$summaryLines = @(
    "VR_RESOURCE_BENCHMARK",
    "",
    "Estado: $statusText",
    "Fecha: $((Get-Date).ToString('o'))",
    "Excel: $OutputXlsx",
    "Hojas detectadas: $($xlsxCheck.SheetCount)",
    "Graficas detectadas en paquete: $($xlsxCheck.ChartCount)",
    "Nota graficas: $chartNote",
    "",
    "Rutas revisadas:",
    "- $HostBenchmarkDir",
    "- $VmBenchmarkDir",
    "",
    "Summary JSON importados: $($summaryRows.Count)",
    "Samples importadas: $($rawRows.Count)",
    "Process samples importadas: $($processRows.Count)",
    "",
    "Escenarios esperados:",
    "- BASELINE_NO_VR",
    "- VR_IDLE",
    "- VR_TEC_RUNNER",
    "",
    "Advertencias:"
)
if ($script:Warnings.Count -eq 0) {
    $summaryLines += "- none"
} else {
    foreach ($warning in $script:Warnings) { $summaryLines += "- $warning" }
}
Set-Content -LiteralPath $SummaryTxt -Value $summaryLines -Encoding UTF8

$conclusionMdLines = @(
    "# VR_RESOURCE_BENCHMARK_CONCLUSIONES",
    "",
    "## Estado",
    "",
    "- Estado del informe: $statusText.",
    "- Excel generado: $OutputXlsx.",
    "- Hojas: $($xlsxCheck.SheetCount).",
    "- Graficas reales: $($xlsxCheck.ChartCount).",
    "",
    "## Conclusiones tecnicas"
)
if (-not $hasData) {
    $conclusionMdLines += @(
        "",
        "- No hay datos reales de benchmark en las rutas revisadas.",
        "- No se calculan consumo incremental, picos ni coste bajo/moderado/alto.",
        "- El resultado actual es una plantilla trazable pendiente de ejecucion experimental.",
        "- Para la memoria no deben usarse porcentajes de rendimiento hasta ejecutar los tres escenarios."
    )
} else {
    $conclusionMdLines += @(
        "",
        '- Los resultados proceden de `summary.json` y `samples.csv` generados por el script de benchmark.',
        '- La calidad de cada escenario queda indicada en `SCENARIO_SUMMARY.DataQuality`.',
        '- El coste en reposo se interpreta con el umbral `Avg_VR_CPU_IDLE < 2 %`.',
        '- El coste durante runner se interpreta con el umbral `Avg_VR_CPU_RUNNER < 10 %`.',
        '- La estabilidad de RAM debe revisarse con `RAW_SAMPLES` y `Max/P95_VR_RAM_MB`.'
    )
}
$conclusionMdLines += @(
    "",
    "## Limitaciones",
    "",
    "- Laboratorio controlado.",
    "- Ventanas de muestreo cortas por diseno reproducible.",
    '- Si falta un escenario, las comparativas incrementales quedan marcadas como `PENDIENTE`.',
    "- Si Excel COM no esta disponible, las graficas reales se omiten y quedan tablas normalizadas para graficar."
)
Set-Content -LiteralPath $ConclusionsMd -Value $conclusionMdLines -Encoding UTF8

$contextLines = @(
    "# VR_RESOURCE_BENCHMARK_CONTEXT",
    "",
    "## Scripts creados",
    "",
    '- `03_RUNNERS/TFM_Benchmark_VR_Resource_Usage_v1.ps1`',
    '- `03_RUNNERS/TFM_Generate_VR_Benchmark_Excel_v1.ps1`',
    "",
    "## Como ejecutar en VM",
    "",
    '```powershell',
    "cd C:\Users\seguridad\Desktop\TFM\03_RUNNERS",
    ".\TFM_Benchmark_VR_Resource_Usage_v1.ps1 -ScenarioName `"BASELINE_NO_VR`" -DurationSec 180 -IntervalSec 2 -StopVRBeforeRun `$true",
    ".\TFM_Benchmark_VR_Resource_Usage_v1.ps1 -ScenarioName `"VR_IDLE`" -DurationSec 180 -IntervalSec 2 -StartVRAfterStop `$true -WarmupSec 60",
    ".\TFM_Benchmark_VR_Resource_Usage_v1.ps1 -ScenarioName `"VR_TEC_RUNNER`" -DurationSec 300 -IntervalSec 2 -RunnerPath `"C:\Users\seguridad\Desktop\TFM\03_RUNNERS\TFM_Run_All_TEC_Tests_v6.ps1`"",
    '```',
    "",
    "## Carpetas generadas",
    "",
    '- VM: `C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas\Logs_Pruebas_TFM\BENCHMARKS`.',
    '- Host/revision: `C:\Users\julio\Desktop\TFM\05_LOGS\BENCHMARKS`.',
    '- Cada ejecucion crea una carpeta `TFM_BENCH_<ScenarioName>_<yyyyMMdd_HHmmss>`.',
    "",
    "## Excel generado",
    "",
    "- $OutputXlsx.",
    "- Estado actual: $statusText.",
    '- Hojas: `RESUMEN_EJECUTIVO`, `RAW_SAMPLES`, `SCENARIO_SUMMARY`, `PROCESS_DETAIL`, `COMPARATIVA`, `GRAFICAS`, `CONCLUSIONES`, `TRAZABILIDAD`.',
    "",
    "## Metricas contenidas",
    "",
    "- CPU/RAM/IO de proceso Velociraptor: media, maximo y P95 donde aplica.",
    "- CPU/RAM/disco/red del sistema.",
    "- Conteo medio/maximo de procesos Velociraptor.",
    "- Comparativas incrementales entre baseline, idle y runner.",
    "",
    "## Cautelas de calidad de datos",
    "",
    '- `BASELINE_NO_VR` es invalido si aparece `VRProcessCountMax > 0`.',
    '- `VR_IDLE` es invalido si `VRProcessCountMax = 0`.',
    '- `VR_TEC_RUNNER` es invalido si no arranca el runner o no hay proceso VR.',
    '- Si no existen datos, el Excel queda como plantilla `PENDIENTE_DE_EJECUCION`.',
    "",
    "## Interpretacion",
    "",
    "- CPU media VR idle < 2 %: impacto bajo.",
    "- CPU media VR runner < 10 %: impacto moderado/bajo.",
    "- RAM estable sin crecimiento progresivo: aceptable.",
    "- Picos breves no sostenidos: aceptables si se documentan.",
    "",
    "## Pendiente",
    "",
    "- Ejecutar los tres escenarios en VM.",
    '- Copiar las carpetas `TFM_BENCH_*` al host en `05_LOGS\BENCHMARKS`.',
    '- Reejecutar `TFM_Generate_VR_Benchmark_Excel_v1.ps1` para generar metricas y graficas finales.'
)
Set-Content -LiteralPath $ContextMd -Value $contextLines -Encoding UTF8

$readmeLines = @(
    "# README_BENCHMARKS",
    "",
    "Directorio para resultados de benchmark de consumo de Velociraptor.",
    "",
    "## Flujo",
    "",
    '1. Copiar `03_RUNNERS/TFM_Benchmark_VR_Resource_Usage_v1.ps1` a la VM si no esta ya copiado.',
    '2. Ejecutar los escenarios `BASELINE_NO_VR`, `VR_IDLE` y `VR_TEC_RUNNER`.',
    '3. Copiar las carpetas `TFM_BENCH_*` desde la VM a este directorio.',
    '4. Ejecutar `03_RUNNERS/TFM_Generate_VR_Benchmark_Excel_v1.ps1` en host.',
    "",
    "## Archivos esperados por RunId",
    "",
    '- `samples.csv`',
    '- `process_samples.csv`',
    '- `summary.json`',
    '- `summary.txt`',
    '- `process_list_start.txt`',
    '- `process_list_end.txt`',
    '- `environment.txt`',
    "",
    "## Estado actual",
    "",
    "- $statusText."
)
Set-Content -LiteralPath $ReadmeMd -Value $readmeLines -Encoding UTF8

$hashTargets = @(
    $benchmarkScript,
    $PSCommandPath,
    $OutputXlsx,
    $SummaryTxt,
    $ConclusionsMd,
    $ContextMd,
    $ReadmeMd
)
$hashLines = @("SHA256SUMS_VR_RESOURCE_BENCHMARK", "Generated: $((Get-Date).ToString('o'))", "")
foreach ($target in $hashTargets) {
    if (Test-Path -LiteralPath $target) {
        try {
            $hash = Get-FileHash -LiteralPath $target -Algorithm SHA256
            $hashLines += ("{0}  {1}" -f $hash.Hash, $target)
        } catch {
            $hashLines += ("ERROR  {0}  {1}" -f $target, $_.Exception.Message)
        }
    } else {
        $hashLines += ("MISSING  {0}" -f $target)
    }
}
Set-Content -LiteralPath $HashFile -Value $hashLines -Encoding UTF8

Write-Ok "Excel generado: $OutputXlsx"
Write-Ok "Resumen: $SummaryTxt"
Write-Ok "Conclusiones: $ConclusionsMd"
Write-Ok "Contexto: $ContextMd"
Write-Ok "Hashes: $HashFile"
Write-Ok "Estado: $statusText"
