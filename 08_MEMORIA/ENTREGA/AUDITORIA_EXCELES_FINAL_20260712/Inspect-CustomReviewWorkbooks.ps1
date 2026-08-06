[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$CurrentPath,
    [Parameter(Mandatory = $true)][string]$OldPath,
    [Parameter(Mandatory = $true)][string]$AnalysisPath,
    [Parameter(Mandatory = $true)][string]$OutputPath
)

$ErrorActionPreference = 'Stop'

if (-not ('TFMCustomReviewInspectNative' -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class TFMCustomReviewInspectNative {
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

function Get-ColumnName {
    param([int]$Column)
    $name = ''
    while ($Column -gt 0) {
        $Column--
        $name = [char](65 + ($Column % 26)) + $name
        $Column = [math]::Floor($Column / 26)
    }
    return $name
}

function Get-FlatValues {
    param($Range)
    $result = [System.Collections.Generic.List[object]]::new()
    $values = $Range.Value2
    if ($values -is [System.Array]) {
        foreach ($value in $values) { $result.Add($value) }
    }
    else { $result.Add($values) }
    return @($result)
}

function Get-WorkbookReport {
    param($Excel, [string]$Path, [string]$Role)

    $resolved = (Resolve-Path -LiteralPath $Path).Path
    $item = Get-Item -LiteralPath $resolved
    $workbooks = $null
    $workbook = $null
    $worksheets = $null
    $names = $null
    $sheetReports = [System.Collections.Generic.List[object]]::new()
    $keywordHits = [System.Collections.Generic.List[object]]::new()
    $formulaHits = [System.Collections.Generic.List[object]]::new()
    $hyperlinks = [System.Collections.Generic.List[object]]::new()
    $definedNames = [System.Collections.Generic.List[object]]::new()
    $targetSheets = @(
        '00_Guia','01_Dashboard','03_Evaluacion_VR','05_Matriz_Resultados','06_Resumen',
        '07_Catalogo_Artifacts','08_Benignas_FP','09_Benchmark_Plan','09_Resultados_Custom',
        '10_Gaps_Custom_P1P4','14_Arquitectura_Custom','README','RESUMEN_EJECUTIVO',
        'TECNICAS_REAL_VR','VISIBILIDAD_SISTEMA','ALERTAS_VR','FP_RUNNER','FP_HITS',
        'COMPARACION_VR_WAZUH','WAZUH_DETALLE','FUENTES','GRAFICAS'
    )
    $keywordPattern = '(?i)(\bbenchmark\b|CPU|RAM|Working\s*Set|Private\s*Bytes|\bI/O\b|consumo(?:\s+esperado)?|\bB[0-8]\b)'

    try {
        $workbooks = $Excel.Workbooks
        $workbook = $workbooks.Open($resolved, 0, $true)
        $worksheets = $workbook.Worksheets

        for ($w = 1; $w -le [int]$worksheets.Count; $w++) {
            $worksheet = $null
            $used = $null
            $rows = $null
            $columns = $null
            $tables = $null
            $chartObjects = $null
            try {
                $worksheet = $worksheets.Item($w)
                $sheetName = [string]$worksheet.Name
                $used = $worksheet.UsedRange
                $rows = $used.Rows
                $columns = $used.Columns
                $startRow = [int]$used.Row
                $startCol = [int]$used.Column
                $rowCount = [int]$rows.Count
                $colCount = [int]$columns.Count
                $values = $used.Value2
                $formulas = $used.Formula

                $nonEmpty = 0
                if ($values -is [System.Array]) {
                    for ($r = 1; $r -le $rowCount; $r++) {
                        for ($c = 1; $c -le $colCount; $c++) {
                            $value = $values[$r,$c]
                            if ($null -ne $value -and [string]$value -ne '') {
                                $nonEmpty++
                                $text = [string]$value
                                if ($text -match $keywordPattern -and $keywordHits.Count -lt 1000) {
                                    $keywordHits.Add([pscustomobject]@{Sheet=$sheetName;Cell=(Get-ColumnName ($startCol+$c-1))+($startRow+$r-1);Value=$text.Substring(0,[math]::Min(500,$text.Length))})
                                }
                            }
                            $formula = $formulas[$r,$c]
                            if ($null -ne $formula -and [string]$formula -like '=*') {
                                $formulaText = [string]$formula
                                if ($formulaText -match $keywordPattern -or $formulaText -match '(?i)09_Benchmark_Plan') {
                                    $formulaHits.Add([pscustomobject]@{Sheet=$sheetName;Cell=(Get-ColumnName ($startCol+$c-1))+($startRow+$r-1);Formula=$formulaText})
                                }
                            }
                        }
                    }
                }
                elseif ($null -ne $values -and [string]$values -ne '') {
                    $nonEmpty = 1
                    $text = [string]$values
                    if ($text -match $keywordPattern) {
                        $keywordHits.Add([pscustomobject]@{Sheet=$sheetName;Cell=(Get-ColumnName $startCol)+$startRow;Value=$text.Substring(0,[math]::Min(500,$text.Length))})
                    }
                }

                $tableList = [System.Collections.Generic.List[object]]::new()
                $tables = $worksheet.ListObjects
                for ($t = 1; $t -le [int]$tables.Count; $t++) {
                    $table = $null
                    $tableRange = $null
                    try {
                        $table = $tables.Item($t)
                        $tableRange = $table.Range
                        $tableList.Add([pscustomobject]@{Name=[string]$table.Name;Range=[string]$tableRange.Address()})
                    }
                    finally { Release-ComObject $tableRange; Release-ComObject $table }
                }

                $chartList = [System.Collections.Generic.List[object]]::new()
                $chartObjects = $worksheet.ChartObjects()
                for ($c = 1; $c -le [int]$chartObjects.Count; $c++) {
                    $chartObject = $null
                    $chart = $null
                    $seriesCollection = $null
                    try {
                        $chartObject = $chartObjects.Item($c)
                        $chart = $chartObject.Chart
                        $seriesCollection = $chart.SeriesCollection()
                        $seriesList = [System.Collections.Generic.List[object]]::new()
                        for ($s = 1; $s -le [int]$seriesCollection.Count; $s++) {
                            $series = $null
                            try {
                                $series = $seriesCollection.Item($s)
                                $seriesList.Add([pscustomobject]@{Name=[string]$series.Name;Formula=[string]$series.Formula})
                            }
                            finally { Release-ComObject $series }
                        }
                        $chartList.Add([pscustomobject]@{
                            Name=[string]$chartObject.Name
                            Title=if($chart.HasTitle){[string]$chart.ChartTitle.Text}else{''}
                            Type=[int]$chart.ChartType
                            Left=[double]$chartObject.Left
                            Top=[double]$chartObject.Top
                            Width=[double]$chartObject.Width
                            Height=[double]$chartObject.Height
                            Series=@($seriesList)
                        })
                    }
                    finally { Release-ComObject $seriesCollection; Release-ComObject $chart; Release-ComObject $chartObject }
                }

                $preview = @()
                if ($targetSheets -contains $sheetName) {
                    $previewRows = if($sheetName -eq 'ALERTAS_VR'){12}elseif($sheetName -eq 'FUENTES'){30}else{[math]::Min(110,$rowCount)}
                    $previewCols = [math]::Min(32,$colCount)
                    $previewRange = $null
                    try {
                        $previewRange = $worksheet.Range((Get-ColumnName $startCol)+$startRow, (Get-ColumnName ($startCol+$previewCols-1))+($startRow+$previewRows-1))
                        $previewValues = $previewRange.Value2
                        if ($previewValues -is [System.Array]) {
                            for ($r = 1; $r -le $previewRows; $r++) {
                                $rowValues = [System.Collections.Generic.List[object]]::new()
                                $hasValue = $false
                                for ($c = 1; $c -le $previewCols; $c++) {
                                    $value = $previewValues[$r,$c]
                                    if ($null -ne $value -and [string]$value -ne '') { $hasValue = $true }
                                    $rowValues.Add($value)
                                }
                                if ($hasValue) { $preview += ,@($rowValues) }
                            }
                        }
                        elseif ($null -ne $previewValues) { $preview += ,@($previewValues) }
                    }
                    finally { Release-ComObject $previewRange }
                }

                $sheetHyperlinks = $null
                try {
                    $sheetHyperlinks = $worksheet.Hyperlinks
                    for ($h = 1; $h -le [int]$sheetHyperlinks.Count; $h++) {
                        $link = $null
                        $linkRange = $null
                        try {
                            $link = $sheetHyperlinks.Item($h)
                            $linkRange = $link.Range
                            $hyperlinks.Add([pscustomobject]@{Sheet=$sheetName;Cell=[string]$linkRange.Address();Address=[string]$link.Address;SubAddress=[string]$link.SubAddress;Text=[string]$link.TextToDisplay})
                        }
                        finally { Release-ComObject $linkRange; Release-ComObject $link }
                    }
                }
                finally { Release-ComObject $sheetHyperlinks }

                $sheetReports.Add([pscustomobject]@{
                    Index=$w;Name=$sheetName;UsedRange=[string]$used.Address();Rows=$rowCount;Columns=$colCount;
                    NonEmpty=$nonEmpty;Tables=@($tableList);Charts=@($chartList);Preview=$preview
                })
            }
            finally {
                Release-ComObject $chartObjects
                Release-ComObject $tables
                Release-ComObject $columns
                Release-ComObject $rows
                Release-ComObject $used
                Release-ComObject $worksheet
            }
        }

        $names = $workbook.Names
        for ($n = 1; $n -le [int]$names.Count; $n++) {
            $nameItem = $null
            try {
                $nameItem = $names.Item($n)
                $definedNames.Add([pscustomobject]@{Name=[string]$nameItem.Name;RefersTo=[string]$nameItem.RefersTo;Visible=[bool]$nameItem.Visible})
            }
            finally { Release-ComObject $nameItem }
        }

        $externalLinks = 0
        try { $links=$workbook.LinkSources(1); if($null -ne $links){$externalLinks=@($links).Count} } catch {}

        return [pscustomobject]@{
            Role=$Role;Path=$resolved;Bytes=$item.Length;LastWriteTime=$item.LastWriteTime.ToString('o');
            SHA256=(Get-FileHash -Algorithm SHA256 -LiteralPath $resolved).Hash;ReadOnly=[bool]$workbook.ReadOnly;
            SheetCount=[int]$worksheets.Count;ExternalLinks=$externalLinks;Sheets=@($sheetReports);
            KeywordHits=@($keywordHits);FormulaKeywordHits=@($formulaHits);DefinedNames=@($definedNames);Hyperlinks=@($hyperlinks)
        }
    }
    finally {
        if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
        Release-ComObject $names
        Release-ComObject $worksheets
        Release-ComObject $workbook
        Release-ComObject $workbooks
    }
}

$excel = $null
$excelPid = 0
$reports = [System.Collections.Generic.List[object]]::new()
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    try { $excel.AutomationSecurity = 3 } catch {}
    [uint32]$pidValue = 0
    [void][TFMCustomReviewInspectNative]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$pidValue)
    $excelPid = [int]$pidValue
    $reports.Add((Get-WorkbookReport $excel $CurrentPath 'CURRENT'))
    $reports.Add((Get-WorkbookReport $excel $OldPath 'OLD_26062026'))
    $reports.Add((Get-WorkbookReport $excel $AnalysisPath 'ANALISIS_V5'))
}
finally {
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $excel
    [gc]::Collect(); [gc]::WaitForPendingFinalizers(); [gc]::Collect(); [gc]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 2
    if($excelPid -gt 0 -and (Get-Process -Id $excelPid -ErrorAction SilentlyContinue)){Stop-Process -Id $excelPid -Force}
}

$payload=[pscustomobject]@{GeneratedAt=(Get-Date).ToString('o');ExcelPid=$excelPid;Reports=@($reports)}
$json=$payload | ConvertTo-Json -Depth 14
[System.IO.File]::WriteAllText($OutputPath,$json,[System.Text.UTF8Encoding]::new($false))
$reports | Select-Object Role,Path,Bytes,SHA256,SheetCount,ExternalLinks,@{N='Charts';E={($_.Sheets.Charts | Measure-Object).Count}},@{N='KeywordHits';E={$_.KeywordHits.Count}},@{N='DefinedNames';E={$_.DefinedNames.Count}},@{N='Hyperlinks';E={$_.Hyperlinks.Count}} | Format-Table -Wrap -AutoSize
