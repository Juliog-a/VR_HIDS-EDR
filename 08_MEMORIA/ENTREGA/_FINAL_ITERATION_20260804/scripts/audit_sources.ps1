param(
    [Parameter(Mandatory = $true)]
    [string]$WorkRoot
)

$ErrorActionPreference = 'Stop'
$qaDir = Join-Path $WorkRoot 'qa'
New-Item -ItemType Directory -Force -Path $qaDir | Out-Null

function Release-ComObject {
    param([object]$Object)
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Get-CellAudit {
    param(
        [object]$Workbook,
        [string]$SheetName,
        [string]$Address
    )
    $sheet = $null
    $cell = $null
    try {
        $sheet = $Workbook.Worksheets.Item($SheetName)
        $cell = $sheet.Range($Address)
        [pscustomobject]@{
            Sheet   = $SheetName
            Address = $Address
            Formula = [string]$cell.Formula
            FormulaLocal = [string]$cell.FormulaLocal
            Value2  = $cell.Value2
            Text    = [string]$cell.Text
        }
    }
    finally {
        Release-ComObject $cell
        Release-ComObject $sheet
    }
}

function Find-WorkbookMatches {
    param(
        [object]$Workbook,
        [string]$Needle,
        [int]$MaxResults = 200
    )
    $matches = [System.Collections.Generic.List[object]]::new()
    foreach ($sheet in @($Workbook.Worksheets)) {
        $used = $null
        $found = $null
        try {
            $used = $sheet.UsedRange
            $found = $used.Find($Needle, $null, -4123, 2, 1, 1, $false, $false, $false)
            if ($null -eq $found) { continue }
            $firstAddress = $found.Address()
            do {
                $matches.Add([pscustomobject]@{
                    Sheet  = [string]$sheet.Name
                    Address = [string]$found.Address()
                    Text    = [string]$found.Text
                    Formula = [string]$found.Formula
                })
                $next = $used.FindNext($found)
                Release-ComObject $found
                $found = $next
                if ($null -eq $found -or $matches.Count -ge $MaxResults) { break }
            } while ($found.Address() -ne $firstAddress)
        }
        catch {
            $matches.Add([pscustomobject]@{
                Sheet = [string]$sheet.Name
                Address = ''
                Text = "AUDIT_ERROR: $($_.Exception.Message)"
                Formula = ''
            })
        }
        finally {
            Release-ComObject $found
            Release-ComObject $used
            Release-ComObject $sheet
        }
        if ($matches.Count -ge $MaxResults) { break }
    }
    return @($matches)
}

function Audit-Workbook {
    param(
        [object]$Excel,
        [string]$Path,
        [string]$Kind
    )
    $wb = $null
    try {
        $wb = $Excel.Workbooks.Open($Path, 0, $true, 5, '', '', $true)
        $Excel.CalculateFullRebuild()

        $sheetAudits = [System.Collections.Generic.List[object]]::new()
        $totalCharts = 0
        $totalTables = 0
        foreach ($sheet in @($wb.Worksheets)) {
            $used = $null
            $listObjects = $null
            $chartObjects = $null
            try {
                $used = $sheet.UsedRange
                $listObjects = $sheet.ListObjects
                $chartObjects = $sheet.ChartObjects()
                $chartCount = [int]$chartObjects.Count
                $tableCount = [int]$listObjects.Count
                $totalCharts += $chartCount
                $totalTables += $tableCount
                $sheetAudits.Add([pscustomobject]@{
                    Name       = [string]$sheet.Name
                    Visible    = [int]$sheet.Visible
                    UsedRows   = [int]$used.Rows.Count
                    UsedCols   = [int]$used.Columns.Count
                    Charts     = $chartCount
                    Tables     = $tableCount
                    PrintArea  = [string]$sheet.PageSetup.PrintArea
                    AutoFilter = [bool]$sheet.AutoFilterMode
                })
            }
            finally {
                Release-ComObject $chartObjects
                Release-ComObject $listObjects
                Release-ComObject $used
                Release-ComObject $sheet
            }
        }

        $targets = [System.Collections.Generic.List[object]]::new()
        if ($Kind -eq 'detection') {
            foreach ($spec in @(
                @('COMPARACION_VR_WAZUH','C11'),
                @('COMPARACION_VR_WAZUH','C12'),
                @('COMPARACION_VR_WAZUH','C13'),
                @('COMPARACION_VR_WAZUH','C14'),
                @('RESUMEN_EJECUTIVO','B34'),
                @('RESUMEN_EJECUTIVO','B35'),
                @('01_Dashboard','A7'),
                @('06_Resumen','A10'),
                @('RESUMEN_EJECUTIVO','A20'),
                @('FP_HITS','A2'),
                @('ALERTAS_VR','I383'),
                @('09_Resultados_Custom','H54'),
                @('12_Plan_Alertas','A1'),
                @('12_Plan_Alertas','A2'),
                @('12_Plan_Alertas','G8')
            )) {
                $targets.Add((Get-CellAudit -Workbook $wb -SheetName $spec[0] -Address $spec[1]))
            }
            foreach ($row in 13..21) {
                $targets.Add((Get-CellAudit -Workbook $wb -SheetName '09_Resultados_Custom' -Address "G$row"))
            }
        }
        else {
            foreach ($spec in @(
                @('README','B3'),
                @('README','B9'),
                @('RUNS_VALIDOS','A1'),
                @('RESUMEN_EJECUTIVO','A1'),
                @('RESUMEN_ESCENARIO','A1'),
                @('CALIDAD_DATOS','A14'),
                @('DISCREPANCIAS','A2')
            )) {
                $targets.Add((Get-CellAudit -Workbook $wb -SheetName $spec[0] -Address $spec[1]))
            }
        }

        $searchTerms = @(
            '#REF!', '#N/A', '#VALUE!', '#DIV/0!', '#NAME?', '/9',
            'IncludeServerGuiInTotal', 'ServerGuiExcludedFromClientMetrics',
            'RunnerStillRunningAtEnd', 'SERVER_GUI', 'webhook',
            '82 filas', 'no reconciliadas', 'Perfil máximo', 'HISTÓRICO'
        )
        $search = [ordered]@{}
        foreach ($term in $searchTerms) {
            $search[$term] = @(Find-WorkbookMatches -Workbook $wb -Needle $term -MaxResults 100)
        }

        $result = [pscustomobject]@{
            Kind          = $Kind
            Path          = $Path
            FileLength    = (Get-Item -LiteralPath $Path).Length
            SHA256        = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
            ExcelVersion  = [string]$Excel.Version
            ReadOnly      = [bool]$wb.ReadOnly
            Worksheets    = [int]$wb.Worksheets.Count
            TotalCharts   = $totalCharts
            TotalTables   = $totalTables
            CalculationVersion = [int]$wb.CalculationVersion
            SheetAudit    = @($sheetAudits)
            Targets       = @($targets)
            Search        = $search
        }
        $out = Join-Path $qaDir ("audit_{0}.json" -f $Kind)
        $result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $out -Encoding UTF8
        return $result
    }
    finally {
        if ($null -ne $wb) {
            $wb.Close($false)
        }
        Release-ComObject $wb
    }
}

function Audit-Word {
    param(
        [object]$Word,
        [string]$Path
    )
    $doc = $null
    try {
        $doc = $Word.Documents.Open($Path, $false, $true, $false)
        $fullText = [string]$doc.Content.Text
        $fullText | Set-Content -LiteralPath (Join-Path $qaDir 'word_fulltext.txt') -Encoding UTF8

        $patterns = [ordered]@{
            FuenteGenerica = '(?im)^\s*Fuente:\s*elaboraci[oó]n'
            FuenteTotal = '(?im)^\s*Fuente:'
            PersonalJulio = 'C:\\Users\\julio\\'
            PersonalSeguridad = 'C:\\Users\\seguridad\\'
            HomeAdmin = '/home/admin/'
            DUMB_LAB = 'DUMB_LAB'
            Candidate = '(?i)\bcandidate\b'
            Working = '(?i)\bworking\b'
            Backup = '(?i)\bbackup\b'
            Commit = '(?i)\bcommit\b'
            Push = '(?i)\bpush\b'
            GitAdd = '(?i)git add \.'
            EquationResidual = 'Equation Chapter 1 Section 1'
            ErrorReference = 'Error[.!]?\s*(No se encuentra el origen de la referencia|Reference source not found)'
            Alertas379 = '(?i)379\s+alertas'
            Filas379 = '(?i)379\s+filas\s+CLIENT_EVENT'
            Rows82 = '(?i)82\s+filas'
            RulesV1 = 'tfm_wazuh_custom_rules_v1\.xml'
            RulesV2 = 'tfm_wazuh_custom_rules_v2\.xml'
            FlagsInclude = 'IncludeServerGuiInTotal=False'
            FlagsExclude = 'ServerGuiExcludedFromClientMetrics=True'
            RunnerWarn = 'RunnerStillRunningAtEnd=True'
            RQ = '(?i)\bRQ[1-9]\b'
            AnnexM = '(?i)Anexo M'
            MIAutorBad = '(?i)\ba mi como autor\b|\ba mi, como autor\b'
            MIAutorGood = '(?i)\ba mí, como autor\b'
            OseaBad = '(?i)\bósea\b'
            Almacenas = '(?i)\balmacenas evidencias\b'
            ScriptBlockBad = '(?i)script block login'
            HIDBad = '(?i)evaluación HID\b'
            QuedaPero = 'queda,pero'
            PasarPasar = 'pasar,pasar'
            CoverageEnterprise = '(?i)cobertura táctica MITRE ATT&CK Enterprise'
            IocField = '(?i)\bIOC\b'
            Discord = '(?i)\bDiscord\b'
            Http200 = '(?i)HTTP(?:Status)?\s*[=:]?\s*200'
        }
        $patternCounts = [ordered]@{}
        foreach ($entry in $patterns.GetEnumerator()) {
            $patternCounts[$entry.Key] = [regex]::Matches($fullText, $entry.Value).Count
        }

        $paragraphFindings = [System.Collections.Generic.List[object]]::new()
        $captionParagraphs = [System.Collections.Generic.List[object]]::new()
        $headingParagraphs = [System.Collections.Generic.List[object]]::new()
        $sourceParagraphs = [System.Collections.Generic.List[object]]::new()
        $rqParagraphs = [System.Collections.Generic.List[object]]::new()
        for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
            $p = $doc.Paragraphs.Item($i)
            $text = ([string]$p.Range.Text).Replace("`r", '').Replace("`a", '').Trim()
            if ([string]::IsNullOrWhiteSpace($text)) {
                Release-ComObject $p
                continue
            }
            $styleName = ''
            try { $styleName = [string]$p.Range.Style.NameLocal } catch { $styleName = [string]$p.Range.Style }
            $page = 0
            try { $page = [int]$p.Range.Information(3) } catch {}
            if ($text -match '^(Tabla|Figura|Código)\s*') {
                $captionParagraphs.Add([pscustomobject]@{Index=$i;Page=$page;Style=$styleName;Text=$text})
            }
            if ($styleName -match 'Título|Heading') {
                $headingParagraphs.Add([pscustomobject]@{Index=$i;Page=$page;Style=$styleName;Text=$text})
            }
            if ($text -match '^(?i)Fuente:') {
                $sourceParagraphs.Add([pscustomobject]@{Index=$i;Page=$page;Style=$styleName;Text=$text})
            }
            if ($text -match '(?i)\bRQ[1-9]\b|pregunta de investigación') {
                $rqParagraphs.Add([pscustomobject]@{Index=$i;Page=$page;Style=$styleName;Text=$text})
            }
            if ($text -match '(?i)(Tabla 63|Tabla 65|Código 5|Figura 15|Tabla 32|Tabla 45|Equation Chapter|Fuente:|379 alertas|IncludeServerGui|ServerGuiExcluded|tfm_wazuh_custom_rules_v1|a mi como autor|ósea|almacenas evidencias|script block login|evaluación HID|queda,pero|pasar,pasar)') {
                $paragraphFindings.Add([pscustomobject]@{Index=$i;Page=$page;Style=$styleName;Text=$text})
            }
            Release-ComObject $p
        }

        $tableAudits = [System.Collections.Generic.List[object]]::new()
        $focus = @(23, 45, 47, 52, 53, 54, 56, 65, 68)
        foreach ($index in $focus) {
            if ($index -gt $doc.Tables.Count) { continue }
            $table = $doc.Tables.Item($index)
            try {
                $cols = [int]$table.Columns.Count
                $rows = [int]$table.Rows.Count
                $width = 0.0
                foreach ($col in @($table.Columns)) {
                    try { $width += [double]$col.Width } catch {}
                    Release-ComObject $col
                }
                $page = 0
                try { $page = [int]$table.Range.Information(3) } catch {}
                $sample = ([string]$table.Range.Text).Replace("`r", ' | ').Replace("`a", '').Trim()
                if ($sample.Length -gt 500) { $sample = $sample.Substring(0, 500) }
                $tableAudits.Add([pscustomobject]@{
                    Index = $index
                    Page = $page
                    Rows = $rows
                    Columns = $cols
                    WidthPoints = [math]::Round($width, 2)
                    AllowAutoFit = [bool]$table.AllowAutoFit
                    Sample = $sample
                })
            }
            finally {
                Release-ComObject $table
            }
        }

        $fieldAudits = [System.Collections.Generic.List[object]]::new()
        foreach ($field in @($doc.Fields)) {
            $code = ([string]$field.Code.Text).Trim()
            $resultText = ([string]$field.Result.Text).Replace("`r", '').Replace("`a", '').Trim()
            if ($code -match '^(SEQ|TOC|TOA|REF|PAGEREF|PAGE|NUMPAGES)\b') {
                $fieldAudits.Add([pscustomobject]@{
                    Type = [int]$field.Type
                    Code = $code
                    Result = $resultText
                })
            }
            Release-ComObject $field
        }

        $hyperlinks = [System.Collections.Generic.List[object]]::new()
        foreach ($link in @($doc.Hyperlinks)) {
            $hyperlinks.Add([pscustomobject]@{
                Text = ([string]$link.TextToDisplay).Trim()
                Address = [string]$link.Address
                SubAddress = [string]$link.SubAddress
            })
            Release-ComObject $link
        }

        $sections = [System.Collections.Generic.List[object]]::new()
        foreach ($section in @($doc.Sections)) {
            $setup = $section.PageSetup
            $sections.Add([pscustomobject]@{
                Index = [int]$section.Index
                Orientation = [int]$setup.Orientation
                PageWidth = [double]$setup.PageWidth
                PageHeight = [double]$setup.PageHeight
                LeftMargin = [double]$setup.LeftMargin
                RightMargin = [double]$setup.RightMargin
                TopMargin = [double]$setup.TopMargin
                BottomMargin = [double]$setup.BottomMargin
                SectionStart = [int]$setup.SectionStart
            })
            Release-ComObject $setup
            Release-ComObject $section
        }

        $result = [pscustomobject]@{
            Path = $Path
            FileLength = (Get-Item -LiteralPath $Path).Length
            SHA256 = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
            WordVersion = [string]$Word.Version
            Pages = [int]$doc.ComputeStatistics(2)
            Words = [int]$doc.ComputeStatistics(0)
            Characters = [int]$doc.ComputeStatistics(3)
            Paragraphs = [int]$doc.Paragraphs.Count
            Tables = [int]$doc.Tables.Count
            InlineShapes = [int]$doc.InlineShapes.Count
            FloatingShapes = [int]$doc.Shapes.Count
            Sections = @($sections)
            TOCs = [int]$doc.TablesOfContents.Count
            FiguresLists = [int]$doc.TablesOfFigures.Count
            Fields = [int]$doc.Fields.Count
            PatternCounts = $patternCounts
            ParagraphFindings = @($paragraphFindings)
            Captions = @($captionParagraphs)
            Headings = @($headingParagraphs)
            SourceParagraphs = @($sourceParagraphs)
            RQParagraphs = @($rqParagraphs)
            FocusTables = @($tableAudits)
            FieldAudit = @($fieldAudits)
            Hyperlinks = @($hyperlinks)
        }
        $result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $qaDir 'audit_word.json') -Encoding UTF8
        return $result
    }
    finally {
        if ($null -ne $doc) {
            $doc.Close($false)
        }
        Release-ComObject $doc
    }
}

$excel = $null
$word = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3

    $detectionPath = Join-Path $WorkRoot 'TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712_WORK.xlsx'
    $benchmarkPath = Join-Path $WorkRoot 'TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_DEFINITIVO_20260712_WORK.xlsx'
    $detectionAudit = Audit-Workbook -Excel $excel -Path $detectionPath -Kind 'detection'
    $benchmarkAudit = Audit-Workbook -Excel $excel -Path $benchmarkPath -Kind 'benchmark'

    $excel.Quit()
    Release-ComObject $excel
    $excel = $null

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    $docPath = Join-Path $WorkRoot 'TFM_ENTREGA_FINAL_WORK.docx'
    $wordAudit = Audit-Word -Word $word -Path $docPath

    $word.Quit()
    Release-ComObject $word
    $word = $null

    $summary = [pscustomobject]@{
        Detection = [pscustomobject]@{
            Worksheets = $detectionAudit.Worksheets
            Charts = $detectionAudit.TotalCharts
            Tables = $detectionAudit.TotalTables
            TargetCells = $detectionAudit.Targets
        }
        Benchmark = [pscustomobject]@{
            Worksheets = $benchmarkAudit.Worksheets
            Charts = $benchmarkAudit.TotalCharts
            Tables = $benchmarkAudit.TotalTables
            TargetCells = $benchmarkAudit.Targets
        }
        Word = [pscustomobject]@{
            Pages = $wordAudit.Pages
            Words = $wordAudit.Words
            Tables = $wordAudit.Tables
            SourceParagraphs = $wordAudit.SourceParagraphs.Count
            Captions = $wordAudit.Captions.Count
            Headings = $wordAudit.Headings.Count
            PatternCounts = $wordAudit.PatternCounts
            FocusTables = $wordAudit.FocusTables
        }
    }
    $summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $qaDir 'audit_summary.json') -Encoding UTF8
    $summary | ConvertTo-Json -Depth 6
}
finally {
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
    }
    if ($null -ne $word) {
        try { $word.Quit() } catch {}
        Release-ComObject $word
    }
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
}
