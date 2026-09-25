<#
TFM Velociraptor resource benchmark Excel generator v1.

Reads benchmark outputs and generates a technical workbook. New benchmark
schema 1.1 separates CLIENT_SERVICE from SERVER_GUI. Legacy outputs are kept
for traceability, but are not used for final client-agent conclusions.
#>

[CmdletBinding()]
param(
    [string]$HostBenchmarkDir = "${PSScriptRoot}\..\05_LOGS\BENCHMARKS",
    [string]$VmBenchmarkDir = "C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas\Logs_Pruebas_TFM\BENCHMARKS",
    [string]$OutputXlsx = "${PSScriptRoot}\..\VR_RESOURCE_BENCHMARK.xlsx",
    [string]$SummaryTxt = "${PSScriptRoot}\..\VR_RESOURCE_BENCHMARK_RESUMEN.txt",
    [string]$ConclusionsMd = "${PSScriptRoot}\..\VR_RESOURCE_BENCHMARK_CONCLUSIONES.md",
    [string]$ContextMd = "${PSScriptRoot}\..\00_CONTEXT\VR_RESOURCE_BENCHMARK_CONTEXT.md",
    [string]$ReadmeMd = "${PSScriptRoot}\..\05_LOGS\BENCHMARKS\README_BENCHMARKS.md",
    [string]$HashFile = "${PSScriptRoot}\..\SHA256SUMS_VR_RESOURCE_BENCHMARK.txt"
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$script:Warnings = New-Object System.Collections.ArrayList

function Add-Warn {
    param([string]$Message)
    [void]$script:Warnings.Add($Message)
    Write-Host "[WARN] $Message"
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

function Get-Prop {
    param(
        [object]$Object,
        [string[]]$Names,
        [object]$Default = ""
    )
    if ($null -eq $Object) { return $Default }
    foreach ($name in $Names) {
        $prop = $Object.PSObject.Properties[$name]
        if ($null -ne $prop) { return $prop.Value }
    }
    return $Default
}

function Get-NumberOrNull {
    param([object]$Value)
    try {
        if ($null -eq $Value) { return $null }
        if ($Value -is [double] -or $Value -is [int] -or $Value -is [long] -or $Value -is [decimal]) {
            return [double]$Value
        }
        $text = [string]$Value
        if ([string]::IsNullOrWhiteSpace($text) -or $text -eq "NA" -or $text -eq "PENDIENTE") { return $null }
        return [double]::Parse($text, [System.Globalization.CultureInfo]::InvariantCulture)
    } catch {
        return $null
    }
}

function Get-BoolText {
    param([object]$Value)
    if ($null -eq $Value) { return "" }
    if ($Value -is [bool]) {
        if ($Value) { return "true" }
        return "false"
    }
    return [string]$Value
}

function Get-NestedMetricAvailability {
    param(
        [object]$Summary,
        [string]$Name
    )
    $ma = Get-Prop -Object $Summary -Names @("MetricsAvailability") -Default $null
    if ($null -eq $ma) { return "" }
    return Get-BoolText -Value (Get-Prop -Object $ma -Names @($Name) -Default "")
}

function New-CellXml {
    param(
        [string]$Ref,
        [object]$Value,
        [bool]$Header = $false
    )

    if ($null -eq $Value) {
        return "<c r=""$Ref"" t=""inlineStr""><is><t>NA</t></is></c>"
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

    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $parent = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }

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
            $xml = New-WorksheetXml -Rows @($Sheets[$i].Rows) -ColumnWidths @($Sheets[$i].Widths)
            Add-ZipText -Zip $zip -EntryName "xl/worksheets/sheet$sheetIndex.xml" -Text $xml
        }
    } finally {
        $zip.Dispose()
        $fs.Dispose()
    }
}

function Test-XlsxPackage {
    param([string]$Path)
    $result = [pscustomobject]@{ Opens = $false; SheetCount = 0; ChartCount = 0; Error = "" }
    try {
        $fs = [System.IO.File]::OpenRead($Path)
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            $wbEntry = $zip.GetEntry("xl/workbook.xml")
            if ($null -eq $wbEntry) { throw "xl/workbook.xml no existe" }
            $reader = New-Object System.IO.StreamReader($wbEntry.Open())
            try { [xml]$wbXml = $reader.ReadToEnd() } finally { $reader.Dispose() }
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

function Get-ProcessRole {
    param(
        [string]$ProcessName,
        [string]$ExecutablePath,
        [string]$CommandLine
    )
    $name = if ($null -eq $ProcessName) { "" } else { $ProcessName }
    $exe = if ($null -eq $ExecutablePath) { "" } else { $ExecutablePath }
    $cmd = if ($null -eq $CommandLine) { "" } else { $CommandLine }
    $combined = "$name $exe $cmd"

    if ($name -ieq "notepad.exe") { return "EXCLUDED_NOTEPAD" }
    if ($combined -notmatch '(?i)velociraptor') { return "NOT_VELOCIRAPTOR" }
    if (($exe -match '(?i)^C:\\Program Files\\Velociraptor\\Velociraptor\.exe$') -or (($cmd -match '(?i)client\.config\.yaml') -and ($cmd -match '(?i)service\s+run'))) { return "CLIENT_SERVICE" }
    if (($exe -match '(?i)^C:\\Users\\seguridad\\Desktop\\TFM\\Velociraptor\\') -or (($cmd -match '(?i)server\.config\.yaml') -and ($cmd -match '(?i)\bgui\b'))) { return "SERVER_GUI" }
    return "OTHER_VELOCIRAPTOR"
}

function Normalize-Summary {
    param(
        [object]$Summary,
        [string]$SourceFile
    )

    $hasClientFields = ($null -ne $Summary.PSObject.Properties["Avg_VR_Client_CPU_Percent"])
    $schema = [string](Get-Prop -Object $Summary -Names @("SchemaVersion") -Default "")
    $scenario = [string](Get-Prop -Object $Summary -Names @("Scenario", "ScenarioName") -Default "UNKNOWN")
    $legacy = (-not $hasClientFields)
    $quality = if ($legacy) { "LEGACY_PARTIAL_ROLE_UNSEPARATED" } else { [string](Get-Prop -Object $Summary -Names @("ScenarioValidity", "DataQuality") -Default "UNKNOWN") }
    $supportsClient = (-not $legacy)

    return [pscustomobject]@{
        Scenario                          = $scenario
        RunId                             = [string](Get-Prop -Object $Summary -Names @("RunId") -Default "")
        SchemaVersion                     = $schema
        DataQuality                       = $quality
        SupportsClientOnlyMetrics         = $supportsClient
        Samples                           = Get-NumberOrNull (Get-Prop -Object $Summary -Names @("Samples") -Default $null)
        DurationSec                       = Get-NumberOrNull (Get-Prop -Object $Summary -Names @("DurationSec") -Default $null)
        ActualDurationSec                 = Get-NumberOrNull (Get-Prop -Object $Summary -Names @("ActualDurationSec") -Default $null)
        ServiceControlSkipped             = Get-BoolText (Get-Prop -Object $Summary -Names @("ServiceControlSkipped") -Default "")
        ServiceControlAvailable           = Get-BoolText (Get-Prop -Object $Summary -Names @("ServiceControlAvailable") -Default "")
        ClientServiceName                 = [string](Get-Prop -Object $Summary -Names @("ClientServiceName", "ServiceName") -Default "")
        ClientServiceStatusStart          = [string](Get-Prop -Object $Summary -Names @("ClientServiceStatusStart", "ServiceStatusStart") -Default "")
        ClientServiceStatusEnd            = [string](Get-Prop -Object $Summary -Names @("ClientServiceStatusEnd", "ServiceStatusEnd") -Default "")
        ClientServicePath                 = [string](Get-Prop -Object $Summary -Names @("ClientServicePath") -Default "")
        VR_Client_ProcessCountStart       = Get-NumberOrNull (Get-Prop -Object $Summary -Names @("VR_Client_ProcessCountStart") -Default $null)
        VR_Client_ProcessCountEnd         = Get-NumberOrNull (Get-Prop -Object $Summary -Names @("VR_Client_ProcessCountEnd") -Default $null)
        VR_Client_ProcessCountMax         = Get-NumberOrNull (Get-Prop -Object $Summary -Names @("VR_Client_ProcessCountMax", "VRProcessCountMax") -Default $null)
        VR_ServerGUI_ProcessCountMax      = Get-NumberOrNull (Get-Prop -Object $Summary -Names @("VR_ServerGUI_ProcessCountMax") -Default $null)
        ServerGuiDetected                 = Get-BoolText (Get-Prop -Object $Summary -Names @("ServerGuiDetected") -Default "")
        ServerGuiExcludedFromClientMetrics = Get-BoolText (Get-Prop -Object $Summary -Names @("ServerGuiExcludedFromClientMetrics") -Default "")
        RunnerStarted                     = Get-BoolText (Get-Prop -Object $Summary -Names @("RunnerStarted") -Default "")
        RunnerObservedSamples             = Get-NumberOrNull (Get-Prop -Object $Summary -Names @("RunnerObservedSamples") -Default $null)
        Avg_VR_Client_CPU_Percent         = if ($legacy) { $null } else { Get-NumberOrNull (Get-Prop -Object $Summary -Names @("Avg_VR_Client_CPU_Percent") -Default $null) }
        Max_VR_Client_CPU_Percent         = if ($legacy) { $null } else { Get-NumberOrNull (Get-Prop -Object $Summary -Names @("Max_VR_Client_CPU_Percent") -Default $null) }
        P95_VR_Client_CPU_Percent         = if ($legacy) { $null } else { Get-NumberOrNull (Get-Prop -Object $Summary -Names @("P95_VR_Client_CPU_Percent") -Default $null) }
        Avg_VR_Client_RAM_MB              = if ($legacy) { $null } else { Get-NumberOrNull (Get-Prop -Object $Summary -Names @("Avg_VR_Client_RAM_MB") -Default $null) }
        Max_VR_Client_RAM_MB              = if ($legacy) { $null } else { Get-NumberOrNull (Get-Prop -Object $Summary -Names @("Max_VR_Client_RAM_MB") -Default $null) }
        P95_VR_Client_RAM_MB              = if ($legacy) { $null } else { Get-NumberOrNull (Get-Prop -Object $Summary -Names @("P95_VR_Client_RAM_MB") -Default $null) }
        Avg_VR_Client_IO_BytesSec         = if ($legacy) { $null } else { Get-NumberOrNull (Get-Prop -Object $Summary -Names @("Avg_VR_Client_IO_BytesSec") -Default $null) }
        Max_VR_Client_IO_BytesSec         = if ($legacy) { $null } else { Get-NumberOrNull (Get-Prop -Object $Summary -Names @("Max_VR_Client_IO_BytesSec") -Default $null) }
        Avg_System_CPU_Percent            = Get-NumberOrNull (Get-Prop -Object $Summary -Names @("Avg_System_CPU_Percent") -Default $null)
        Max_System_CPU_Percent            = Get-NumberOrNull (Get-Prop -Object $Summary -Names @("Max_System_CPU_Percent") -Default $null)
        Avg_System_RAM_Used_MB            = Get-NumberOrNull (Get-Prop -Object $Summary -Names @("Avg_System_RAM_Used_MB") -Default $null)
        Max_System_RAM_Used_MB            = Get-NumberOrNull (Get-Prop -Object $Summary -Names @("Max_System_RAM_Used_MB") -Default $null)
        Avg_Disk_BytesSec                 = Get-NumberOrNull (Get-Prop -Object $Summary -Names @("Avg_Disk_BytesSec") -Default $null)
        Avg_Net_BytesSec                  = Get-NumberOrNull (Get-Prop -Object $Summary -Names @("Avg_Net_BytesSec") -Default $null)
        MetricsSystemCPUAvailable         = Get-NestedMetricAvailability -Summary $Summary -Name "SystemCPUAvailable"
        MetricsSystemDiskAvailable        = Get-NestedMetricAvailability -Summary $Summary -Name "SystemDiskAvailable"
        MetricsSystemNetAvailable         = Get-NestedMetricAvailability -Summary $Summary -Name "SystemNetAvailable"
        MetricsProcessPerfAvailable       = Get-NestedMetricAvailability -Summary $Summary -Name "ProcessPerfAvailable"
        SourceFile                        = $SourceFile
        StartTime                         = [string](Get-Prop -Object $Summary -Names @("StartTime") -Default "")
    }
}

function Latest-ValidScenario {
    param(
        [object[]]$Rows,
        [string]$Scenario
    )
    $items = @($Rows | Where-Object {
        $_.Scenario -eq $Scenario -and $_.SupportsClientOnlyMetrics -eq $true -and $_.DataQuality -eq "VALID"
    } | Sort-Object StartTime -Descending | Select-Object -First 1)
    if ($items.Count -gt 0) { return $items[0] }
    return $null
}

function New-DeltaRow {
    param(
        [string]$Comparison,
        [string]$Metric,
        [object]$A,
        [object]$B
    )
    if ($null -eq $A -or $null -eq $B) { return @($Comparison, $Metric, "PENDIENTE", "", "", "") }
    $aVal = Get-NumberOrNull $A.$Metric
    $bVal = Get-NumberOrNull $B.$Metric
    if ($null -eq $aVal -or $null -eq $bVal) { return @($Comparison, $Metric, "PENDIENTE", $aVal, $bVal, "") }
    $delta = [math]::Round(($bVal - $aVal), 2)
    $ratio = ""
    if ([math]::Abs($aVal) -gt 0.000001) { $ratio = [math]::Round(($bVal / $aVal), 2) }
    return @($Comparison, $Metric, $delta, $aVal, $bVal, $ratio)
}

foreach ($dir in @((Split-Path -Parent $OutputXlsx), (Split-Path -Parent $SummaryTxt), (Split-Path -Parent $ConclusionsMd), (Split-Path -Parent $ContextMd), (Split-Path -Parent $ReadmeMd), (Split-Path -Parent $HashFile), $HostBenchmarkDir)) {
    if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

$sourceDirs = New-Object System.Collections.ArrayList
if (Test-Path -LiteralPath $HostBenchmarkDir) { [void]$sourceDirs.Add($HostBenchmarkDir) }
if (Test-Path -LiteralPath $VmBenchmarkDir) {
    if (-not ($sourceDirs -contains $VmBenchmarkDir)) { [void]$sourceDirs.Add($VmBenchmarkDir) }
} else {
    Add-Warn "Ruta VM no disponible desde host; no se consulta: $VmBenchmarkDir"
}

$summaryRows = New-Object System.Collections.ArrayList
$rawRows = New-Object System.Collections.ArrayList
$processRows = New-Object System.Collections.ArrayList
$usedFiles = New-Object System.Collections.ArrayList

foreach ($dir in $sourceDirs) {
    $summaryFiles = @(Get-ChildItem -LiteralPath $dir -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq "summary.json" -or $_.Name -like "*summary*.json" })
    foreach ($file in $summaryFiles) {
        try {
            $json = Get-Content -LiteralPath $file.FullName -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            $norm = Normalize-Summary -Summary $json -SourceFile $file.FullName
            [void]$summaryRows.Add($norm)
            [void]$usedFiles.Add($file.FullName)
            $runDir = Split-Path -Parent $file.FullName

            foreach ($sampleFile in @(Get-ChildItem -LiteralPath $runDir -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq "samples.csv" -or $_.Name -like "*_samples.csv" })) {
                try {
                    foreach ($item in @(Import-Csv -LiteralPath $sampleFile.FullName -ErrorAction Stop)) {
                        $item | Add-Member -NotePropertyName SourceFile -NotePropertyValue $sampleFile.FullName -Force
                        $item | Add-Member -NotePropertyName DataQuality -NotePropertyValue $norm.DataQuality -Force
                        [void]$rawRows.Add($item)
                    }
                    [void]$usedFiles.Add($sampleFile.FullName)
                } catch {
                    Add-Warn "No se pudo importar samples CSV '$($sampleFile.FullName)': $($_.Exception.Message)"
                }
            }

            foreach ($procFile in @(Get-ChildItem -LiteralPath $runDir -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq "process_samples.csv" -or $_.Name -like "*process_samples.csv" })) {
                try {
                    foreach ($item in @(Import-Csv -LiteralPath $procFile.FullName -ErrorAction Stop)) {
                        $item | Add-Member -NotePropertyName SourceFile -NotePropertyValue $procFile.FullName -Force
                        $item | Add-Member -NotePropertyName DataQuality -NotePropertyValue $norm.DataQuality -Force
                        if ($null -eq $item.PSObject.Properties["Role"] -or [string]::IsNullOrWhiteSpace([string]$item.Role)) {
                            $role = Get-ProcessRole -ProcessName ([string](Get-Prop $item @("ProcessName", "Name") "")) -ExecutablePath ([string](Get-Prop $item @("ExecutablePath") "")) -CommandLine ([string](Get-Prop $item @("CommandLine") ""))
                            $item | Add-Member -NotePropertyName Role -NotePropertyValue $role -Force
                        }
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

$baselineObj = Latest-ValidScenario -Rows @($summaryRows) -Scenario "BASELINE_NO_VR"
$idleObj = Latest-ValidScenario -Rows @($summaryRows) -Scenario "VR_IDLE"
$runnerObj = Latest-ValidScenario -Rows @($summaryRows) -Scenario "VR_TEC_RUNNER"
$hasDefinitiveData = ($null -ne $baselineObj -and $null -ne $idleObj -and $null -ne $runnerObj)
$statusText = if ($hasDefinitiveData) { "DATOS_DEFINITIVOS_CLIENTE" } elseif ($summaryRows.Count -gt 0) { "DATOS_PARCIALES_O_LEGACY" } else { "PENDIENTE_DE_EJECUCION" }

$executiveRows = New-Object System.Collections.ArrayList
[void]$executiveRows.Add(@("Metrica", "Valor", "Origen", "Nota"))
[void]$executiveRows.Add(@("Estado", $statusText, "GENERADO", "Solo se concluye sobre CLIENT_SERVICE con schema 1.1 valido."))
[void]$executiveRows.Add(@("Summary JSON importados", $summaryRows.Count, "LOG", "Incluye legacy si existe."))
[void]$executiveRows.Add(@("Samples importadas", $rawRows.Count, "LOG", "RAW_SAMPLES normalizado."))
[void]$executiveRows.Add(@("Process samples importadas", $processRows.Count, "LOG", "PROCESS_DETAIL normalizado."))
if ($hasDefinitiveData) {
    [void]$executiveRows.Add(@("CPU media cliente idle", $idleObj.Avg_VR_Client_CPU_Percent, "LOG", "Impacto bajo si < 2 %."))
    [void]$executiveRows.Add(@("RAM media cliente idle MB", $idleObj.Avg_VR_Client_RAM_MB, "LOG", "CLIENT_SERVICE solo."))
    [void]$executiveRows.Add(@("CPU max cliente runner", $runnerObj.Max_VR_Client_CPU_Percent, "LOG", "Pico durante runner TEC."))
    [void]$executiveRows.Add(@("RAM max cliente runner MB", $runnerObj.Max_VR_Client_RAM_MB, "LOG", "Pico durante runner TEC."))
} else {
    [void]$executiveRows.Add(@("Metricas finales", "PENDIENTE", "CALIDAD_DATOS", "Faltan tres escenarios validos con schema 1.1."))
}

$summaryHeader = @(
    "Scenario", "RunId", "SchemaVersion", "DataQuality", "SupportsClientOnlyMetrics",
    "Samples", "DurationSec", "ActualDurationSec", "ServiceControlSkipped",
    "ServiceControlAvailable", "ClientServiceName", "ClientServiceStatusStart",
    "ClientServiceStatusEnd", "ClientServicePath", "VR_Client_ProcessCountStart",
    "VR_Client_ProcessCountEnd", "VR_Client_ProcessCountMax",
    "VR_ServerGUI_ProcessCountMax", "ServerGuiDetected",
    "ServerGuiExcludedFromClientMetrics", "RunnerStarted", "RunnerObservedSamples",
    "Avg_VR_Client_CPU_Percent", "Max_VR_Client_CPU_Percent",
    "P95_VR_Client_CPU_Percent", "Avg_VR_Client_RAM_MB", "Max_VR_Client_RAM_MB",
    "P95_VR_Client_RAM_MB", "Avg_VR_Client_IO_BytesSec",
    "Max_VR_Client_IO_BytesSec", "Avg_System_CPU_Percent",
    "Avg_System_RAM_Used_MB", "Avg_Disk_BytesSec", "Avg_Net_BytesSec", "SourceFile"
)
$scenarioSummaryRows = New-Object System.Collections.ArrayList
[void]$scenarioSummaryRows.Add($summaryHeader)
if ($summaryRows.Count -eq 0) {
    [void]$scenarioSummaryRows.Add(@("PENDIENTE", "", "", "PENDIENTE_EJECUCION", $false, 0, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""))
} else {
    foreach ($row in @($summaryRows | Sort-Object Scenario, StartTime)) {
        [void]$scenarioSummaryRows.Add(@(
            $row.Scenario, $row.RunId, $row.SchemaVersion, $row.DataQuality,
            $row.SupportsClientOnlyMetrics, $row.Samples, $row.DurationSec,
            $row.ActualDurationSec, $row.ServiceControlSkipped,
            $row.ServiceControlAvailable, $row.ClientServiceName,
            $row.ClientServiceStatusStart, $row.ClientServiceStatusEnd,
            $row.ClientServicePath, $row.VR_Client_ProcessCountStart,
            $row.VR_Client_ProcessCountEnd, $row.VR_Client_ProcessCountMax,
            $row.VR_ServerGUI_ProcessCountMax, $row.ServerGuiDetected,
            $row.ServerGuiExcludedFromClientMetrics, $row.RunnerStarted,
            $row.RunnerObservedSamples, $row.Avg_VR_Client_CPU_Percent,
            $row.Max_VR_Client_CPU_Percent, $row.P95_VR_Client_CPU_Percent,
            $row.Avg_VR_Client_RAM_MB, $row.Max_VR_Client_RAM_MB,
            $row.P95_VR_Client_RAM_MB, $row.Avg_VR_Client_IO_BytesSec,
            $row.Max_VR_Client_IO_BytesSec, $row.Avg_System_CPU_Percent,
            $row.Avg_System_RAM_Used_MB, $row.Avg_Disk_BytesSec,
            $row.Avg_Net_BytesSec, $row.SourceFile
        ))
    }
}

$rawHeader = @(
    "RunId", "Scenario", "Timestamp", "ElapsedSec", "ServiceStatus",
    "VR_Client_ProcessCount", "VR_Client_CPU_Percent", "VR_Client_RAM_MB",
    "VR_Client_IO_Total_BytesSec", "VR_ServerGUI_ProcessCount",
    "VR_ServerGUI_CPU_Percent", "VR_ServerGUI_RAM_MB", "System_CPU_Percent",
    "System_RAM_Used_MB", "System_Disk_BytesSec", "System_Net_BytesSec",
    "RunnerActive", "DataQuality", "SourceFile"
)
$rawSheetRows = New-Object System.Collections.ArrayList
[void]$rawSheetRows.Add($rawHeader)
if ($rawRows.Count -eq 0) {
    [void]$rawSheetRows.Add(@("PENDIENTE", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "PENDIENTE_EJECUCION", ""))
} else {
    foreach ($row in @($rawRows | Select-Object -First 10000)) {
        [void]$rawSheetRows.Add(@(
            (Get-Prop $row @("RunId") ""),
            (Get-Prop $row @("Scenario") ""),
            (Get-Prop $row @("Timestamp", "Time") ""),
            (Get-Prop $row @("ElapsedSec") ""),
            (Get-Prop $row @("ServiceStatus") ""),
            (Get-Prop $row @("VR_Client_ProcessCount") ""),
            (Get-Prop $row @("VR_Client_CPU_Percent") ""),
            (Get-Prop $row @("VR_Client_RAM_MB") ""),
            (Get-Prop $row @("VR_Client_IO_Total_BytesSec") ""),
            (Get-Prop $row @("VR_ServerGUI_ProcessCount") ""),
            (Get-Prop $row @("VR_ServerGUI_CPU_Percent") ""),
            (Get-Prop $row @("VR_ServerGUI_RAM_MB") ""),
            (Get-Prop $row @("System_CPU_Percent") ""),
            (Get-Prop $row @("System_RAM_Used_MB") ""),
            (Get-Prop $row @("System_Disk_BytesSec") ""),
            (Get-Prop $row @("System_Net_BytesSec") ""),
            (Get-Prop $row @("RunnerActive", "RunnerRunning") ""),
            (Get-Prop $row @("DataQuality") ""),
            (Get-Prop $row @("SourceFile") "")
        ))
    }
}

$processHeader = @(
    "RunId", "Scenario", "Timestamp", "ElapsedSec", "ServiceStatus", "PID", "Role",
    "ProcessName", "ExecutablePath", "CommandLine", "CPU_Percent", "RAM_MB",
    "WorkingSet_MB", "PrivateMemory_MB", "Handles", "Threads", "IO_Total_BytesSec",
    "ProcessPerfAvailable", "RunnerActive", "DataQuality", "SourceFile"
)
$processSheetRows = New-Object System.Collections.ArrayList
[void]$processSheetRows.Add($processHeader)
if ($processRows.Count -eq 0) {
    [void]$processSheetRows.Add(@("PENDIENTE", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "PENDIENTE_EJECUCION", ""))
} else {
    foreach ($row in @($processRows | Select-Object -First 10000)) {
        [void]$processSheetRows.Add(@(
            (Get-Prop $row @("RunId") ""),
            (Get-Prop $row @("Scenario") ""),
            (Get-Prop $row @("Timestamp", "Time") ""),
            (Get-Prop $row @("ElapsedSec") ""),
            (Get-Prop $row @("ServiceStatus") ""),
            (Get-Prop $row @("PID", "ProcessId") ""),
            (Get-Prop $row @("Role") ""),
            (Get-Prop $row @("ProcessName", "Name") ""),
            (Get-Prop $row @("ExecutablePath") ""),
            (Get-Prop $row @("CommandLine") ""),
            (Get-Prop $row @("CPU_Percent") ""),
            (Get-Prop $row @("RAM_MB") ""),
            (Get-Prop $row @("WorkingSet_MB") ""),
            (Get-Prop $row @("PrivateMemory_MB") ""),
            (Get-Prop $row @("Handles") ""),
            (Get-Prop $row @("Threads") ""),
            (Get-Prop $row @("IO_Total_BytesSec", "VR_IO_BytesSec") ""),
            (Get-Prop $row @("ProcessPerfAvailable") ""),
            (Get-Prop $row @("RunnerActive") ""),
            (Get-Prop $row @("DataQuality") ""),
            (Get-Prop $row @("SourceFile") "")
        ))
    }
}

$comparisonRows = New-Object System.Collections.ArrayList
[void]$comparisonRows.Add(@("Comparacion", "Metrica", "Delta", "Valor_A", "Valor_B", "Ratio_B_sobre_A"))
foreach ($metric in @("Avg_VR_Client_CPU_Percent", "Avg_VR_Client_RAM_MB", "Avg_VR_Client_IO_BytesSec")) {
    [void]$comparisonRows.Add((New-DeltaRow -Comparison "VR_IDLE_vs_BASELINE_NO_VR" -Metric $metric -A $baselineObj -B $idleObj))
    [void]$comparisonRows.Add((New-DeltaRow -Comparison "VR_TEC_RUNNER_vs_VR_IDLE" -Metric $metric -A $idleObj -B $runnerObj))
    [void]$comparisonRows.Add((New-DeltaRow -Comparison "VR_TEC_RUNNER_vs_BASELINE_NO_VR" -Metric $metric -A $baselineObj -B $runnerObj))
}

$graphRows = New-Object System.Collections.ArrayList
[void]$graphRows.Add(@("Scenario", "Avg_Client_CPU", "Max_Client_CPU", "P95_Client_CPU", "Avg_Client_RAM_MB", "Max_Client_RAM_MB", "P95_Client_RAM_MB", "Avg_Client_IO_BytesSec", "Max_Client_IO_BytesSec"))
foreach ($row in @($baselineObj, $idleObj, $runnerObj)) {
    if ($null -ne $row) {
        [void]$graphRows.Add(@($row.Scenario, $row.Avg_VR_Client_CPU_Percent, $row.Max_VR_Client_CPU_Percent, $row.P95_VR_Client_CPU_Percent, $row.Avg_VR_Client_RAM_MB, $row.Max_VR_Client_RAM_MB, $row.P95_VR_Client_RAM_MB, $row.Avg_VR_Client_IO_BytesSec, $row.Max_VR_Client_IO_BytesSec))
    }
}
if ($graphRows.Count -eq 1) {
    [void]$graphRows.Add(@("PENDIENTE", "", "", "", "", "", "", "", ""))
}

$conclusionRows = New-Object System.Collections.ArrayList
[void]$conclusionRows.Add(@("Apartado", "Conclusion", "Base"))
if (-not $hasDefinitiveData) {
    [void]$conclusionRows.Add(@("Estado", "PENDIENTE/NO DEFINITIVO. No hay tres escenarios validos con schema 1.1 cliente/server.", "CALIDAD_DATOS"))
    [void]$conclusionRows.Add(@("Benchmarks anteriores", "Parciales/no definitivos si proceden de schema legacy o mezclan cliente y server GUI.", "TRAZABILIDAD"))
} else {
    $idleImpact = if ($idleObj.Avg_VR_Client_CPU_Percent -lt 2) { "Impacto CPU cliente en reposo bajo segun criterio < 2 %." } else { "Impacto CPU cliente en reposo no bajo; revisar." }
    $runnerImpact = if ($runnerObj.Avg_VR_Client_CPU_Percent -lt 10) { "Impacto CPU cliente durante runner bajo/moderado segun criterio < 10 %." } else { "Impacto CPU cliente durante runner alto; revisar." }
    [void]$conclusionRows.Add(@("Impacto idle", $idleImpact, "LOG_SCHEMA_1_1"))
    [void]$conclusionRows.Add(@("Impacto runner", $runnerImpact, "LOG_SCHEMA_1_1"))
    [void]$conclusionRows.Add(@("Server GUI", "No se usa para conclusiones del agente cliente.", "DISEÑO"))
}

$qualityRows = New-Object System.Collections.ArrayList
[void]$qualityRows.Add(@("Control", "Resultado", "Detalle"))
[void]$qualityRows.Add(@("Estado global", $statusText, "Solo schema 1.1 con DataQuality VALID permite conclusiones."))
[void]$qualityRows.Add(@("Summaries legacy", @($summaryRows | Where-Object { $_.SupportsClientOnlyMetrics -eq $false }).Count, "No usar para porcentajes finales del agente cliente."))
[void]$qualityRows.Add(@("Summaries schema 1.1 validos", @($summaryRows | Where-Object { $_.SupportsClientOnlyMetrics -eq $true -and $_.DataQuality -eq "VALID" }).Count, "Base potencial para conclusiones."))
[void]$qualityRows.Add(@("BASELINE_NO_VR valido", ($null -ne $baselineObj), "Requiere VR_Client_ProcessCountMax = 0."))
[void]$qualityRows.Add(@("VR_IDLE valido", ($null -ne $idleObj), "Requiere cliente observado."))
[void]$qualityRows.Add(@("VR_TEC_RUNNER valido", ($null -ne $runnerObj), "Requiere cliente y runner observado."))
foreach ($warning in $script:Warnings) { [void]$qualityRows.Add(@("WARN", "WARN", $warning)) }

$traceRows = New-Object System.Collections.ArrayList
[void]$traceRows.Add(@("Tipo", "Ruta", "SHA256", "Nota"))
foreach ($dir in @($HostBenchmarkDir, $VmBenchmarkDir)) {
    [void]$traceRows.Add(@("Ruta revisada", $dir, "", $(if (Test-Path -LiteralPath $dir) { "Disponible" } else { "No disponible desde host" })))
}
foreach ($file in @($usedFiles | Select-Object -Unique)) {
    $hash = ""
    try { $hash = (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash } catch { $hash = "" }
    [void]$traceRows.Add(@("Archivo procesado", $file, $hash, "LOG"))
}
if ($usedFiles.Count -eq 0) { [void]$traceRows.Add(@("Archivo procesado", "Ninguno", "", "PENDIENTE_EJECUCION")) }

$sheets = @(
    [pscustomobject]@{ Name = "RESUMEN_EJECUTIVO"; Rows = @($executiveRows); Widths = @(34, 32, 18, 90) },
    [pscustomobject]@{ Name = "RAW_SAMPLES"; Rows = @($rawSheetRows); Widths = @(26, 20, 28, 12, 18, 18, 18, 18, 22, 20, 20, 18, 18, 20, 20, 20, 14, 32, 90) },
    [pscustomobject]@{ Name = "SCENARIO_SUMMARY"; Rows = @($scenarioSummaryRows); Widths = @(20, 32, 14, 30, 18, 10, 12, 16, 18, 20, 20, 18, 18, 70, 18, 18, 18, 20, 16, 22, 14, 18, 18, 18, 18, 18, 18, 18, 22, 22, 18, 20, 18, 18, 90) },
    [pscustomobject]@{ Name = "PROCESS_DETAIL"; Rows = @($processSheetRows); Widths = @(26, 20, 28, 12, 18, 12, 22, 24, 70, 90, 16, 14, 16, 18, 12, 12, 20, 20, 14, 32, 90) },
    [pscustomobject]@{ Name = "COMPARATIVA"; Rows = @($comparisonRows); Widths = @(34, 32, 14, 14, 14, 18) },
    [pscustomobject]@{ Name = "GRAFICAS"; Rows = @($graphRows); Widths = @(22, 18, 18, 18, 20, 20, 20, 24, 24) },
    [pscustomobject]@{ Name = "CONCLUSIONES"; Rows = @($conclusionRows); Widths = @(28, 100, 24) },
    [pscustomobject]@{ Name = "TRAZABILIDAD"; Rows = @($traceRows); Widths = @(24, 100, 68, 44) },
    [pscustomobject]@{ Name = "CALIDAD_DATOS"; Rows = @($qualityRows); Widths = @(32, 24, 100) }
)

New-XlsxOpenXml -Path $OutputXlsx -Sheets $sheets
$xlsxCheck = Test-XlsxPackage -Path $OutputXlsx
if (-not $xlsxCheck.Opens) { throw "El XLSX generado no es valido: $($xlsxCheck.Error)" }

$summaryLines = @(
    "VR_RESOURCE_BENCHMARK",
    "",
    "Estado: $statusText",
    "Fecha: $((Get-Date).ToString('o'))",
    "Excel: $OutputXlsx",
    "Hojas detectadas: $($xlsxCheck.SheetCount)",
    "Graficas detectadas en paquete: $($xlsxCheck.ChartCount)",
    "",
    "Summary JSON importados: $($summaryRows.Count)",
    "Samples importadas: $($rawRows.Count)",
    "Process samples importadas: $($processRows.Count)",
    "",
    "Nota: las ejecuciones legacy/parciales no se usan para conclusiones del agente cliente.",
    "",
    "Advertencias:"
)
if ($script:Warnings.Count -eq 0) { $summaryLines += "- none" } else { foreach ($warning in $script:Warnings) { $summaryLines += "- $warning" } }
Set-Content -LiteralPath $SummaryTxt -Value $summaryLines -Encoding UTF8

$conclusionMdLines = @(
    "# VR_RESOURCE_BENCHMARK_CONCLUSIONES",
    "",
    "## Estado",
    "",
    "- Estado del informe: $statusText.",
    "- Excel generado: $OutputXlsx.",
    "- Hojas: $($xlsxCheck.SheetCount).",
    "",
    "## Criterio tecnico",
    "",
    "- Las conclusiones de impacto se basan solo en `CLIENT_SERVICE`.",
    "- `SERVER_GUI` queda excluido de las metricas del agente cliente.",
    "- Los benchmarks legacy/parciales quedan como trazabilidad, no como resultado final.",
    "",
    "## Pendiente",
    "",
    "- Ejecutar de nuevo los tres escenarios con `-SkipServiceControl `$true` y el script corregido.",
    "- Regenerar el Excel tras copiar los resultados definitivos a `05_LOGS\\BENCHMARKS`."
)
Set-Content -LiteralPath $ConclusionsMd -Value $conclusionMdLines -Encoding UTF8

$contextLines = @(
    "# VR_RESOURCE_BENCHMARK_CONTEXT",
    "",
    "## Correccion vigente",
    "",
    "- `TFM_Benchmark_VR_Resource_Usage_v1.ps1` incluye `SkipServiceControl` en `param()`.",
    "- El benchmark separa `CLIENT_SERVICE`, `SERVER_GUI` y `OTHER_VELOCIRAPTOR`.",
    "- El Excel tiene hoja `CALIDAD_DATOS`.",
    "- No usar benchmarks anteriores como definitivos si no proceden del script corregido.",
    "",
    "## Estado Excel",
    "",
    "- Estado: $statusText.",
    "- Hojas: 9.",
    "",
    "## Interpretacion",
    "",
    "- CPU media cliente idle < 2 %: impacto bajo.",
    "- CPU media cliente runner < 10 %: impacto bajo/moderado.",
    "- RAM estable sin crecimiento progresivo: aceptable.",
    "- Server/GUI local no forma parte del coste del agente HIDS."
)
Set-Content -LiteralPath $ContextMd -Value $contextLines -Encoding UTF8

$readmeLines = @(
    "# README_BENCHMARKS",
    "",
    "Directorio para resultados de benchmark de consumo de Velociraptor.",
    "",
    "## Estado",
    "",
    "- Estado actual: $statusText.",
    "- Los resultados previos al fix de roles/SkipServiceControl son parciales/no definitivos.",
    "",
    "## Flujo",
    "",
    "1. Ejecutar `INSTRUCCIONES_REPETIR_BENCHMARK_v3.md` en VM.",
    "2. Copiar las carpetas `TFM_BENCH_*` a este directorio.",
    "3. Ejecutar `03_RUNNERS/TFM_Generate_VR_Benchmark_Excel_v1.ps1` en host."
)
Set-Content -LiteralPath $ReadmeMd -Value $readmeLines -Encoding UTF8

$projectRoot = Split-Path -Parent $PSScriptRoot
$hashTargets = @(
    (Join-Path $PSScriptRoot "TFM_Benchmark_VR_Resource_Usage_v1.ps1"),
    (Join-Path $PSScriptRoot "TFM_Benchmark_VR_Resource_Usage_v1_PRE_SKIP_PARAM_BROKEN_BACKUP.ps1"),
    $PSCommandPath,
    (Join-Path $PSScriptRoot "TFM_Generate_VR_Benchmark_Excel_v1_PRE_ROLE_FIX_BACKUP.ps1"),
    (Join-Path $projectRoot "05_LOGS\BENCHMARKS\INSTRUCCIONES_REPETIR_BENCHMARK_v3.md"),
    (Join-Path $projectRoot "05_LOGS\BENCHMARKS\PARAMETER_VALIDATION_BENCHMARK_v1.txt"),
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
Write-Ok "Hojas: $($xlsxCheck.SheetCount)"
Write-Ok "Estado: $statusText"
Write-Ok "Hashes: $HashFile"
