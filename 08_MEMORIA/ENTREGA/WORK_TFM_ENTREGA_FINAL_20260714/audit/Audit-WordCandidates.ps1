[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ManifestPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Release-ComObject {
    param([object]$Object)
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Clean-WordText {
    param([AllowNull()][string]$Text)
    if ($null -eq $Text) { return '' }
    return (($Text -replace "`r", ' ' -replace "`a", ' ' -replace "`v", ' ' -replace "`t", ' ') -replace '\s+', ' ').Trim()
}

function Get-StyleName {
    param([object]$Range)
    $style = $null
    try {
        $style = $Range.Style
        if ($style -is [string]) { return [string]$style }
        try { return [string]$style.NameLocal } catch { return [string]$style }
    }
    catch {
        return ''
    }
    finally {
        if ($null -ne $style -and $style -isnot [string]) { Release-ComObject $style }
    }
}

function Get-PageNumber {
    param([object]$Range)
    try { return [int]$Range.Information(3) } catch { return $null }
}

$manifestFull = (Resolve-Path -LiteralPath $ManifestPath).Path
$manifest = Get-Content -Raw -LiteralPath $manifestFull | ConvertFrom-Json
if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}
$outFull = (Resolve-Path -LiteralPath $OutputDir).Path
$progressPath = Join-Path $outFull 'word_candidates_audit_progress.log'
[IO.File]::WriteAllText($progressPath, "INICIO $(Get-Date -Format o)`r`n", $utf8NoBom)

function Write-ProgressLog {
    param([string]$Message)
    [IO.File]::AppendAllText($progressPath, "$(Get-Date -Format o) | $Message`r`n", $utf8NoBom)
}

$baselinePids = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$word = $null
$documents = $null
$report = [ordered]@{
    generated_at = (Get-Date).ToString('o')
    word_version = $null
    word_build = $null
    baseline_winword_pids = $baselinePids
    documents = @()
}

try {
    Write-ProgressLog 'Creando Word.Application'
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    try { $word.AutomationSecurity = 3 } catch {}
    try { $word.Options.UpdateLinksAtOpen = $false } catch {}
    try { $word.Options.SaveNormalPrompt = $false } catch {}
    $report.word_version = [string]$word.Version
    $report.word_build = [string]$word.Build
    $documents = $word.Documents
    Write-ProgressLog "Word creado. Version=$($report.word_version) Build=$($report.word_build)"

    foreach ($inputPath in @($manifest.documents)) {
        if (-not (Test-Path -LiteralPath $inputPath)) {
            throw "No existe el Word candidato: $inputPath"
        }

        $fullPath = (Resolve-Path -LiteralPath $inputPath).Path
        $item = Get-Item -LiteralPath $fullPath
        $hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $fullPath).Hash
        $doc = $null
        $content = $null
        $paragraphs = $null
        $tables = $null
        $inlineShapes = $null
        $shapes = $null
        $sections = $null
        $fields = $null
        $headings = [System.Collections.Generic.List[object]]::new()
        $captions = [System.Collections.Generic.List[object]]::new()
        $paragraphRecords = [System.Collections.Generic.List[object]]::new()
        $tableRecords = [System.Collections.Generic.List[object]]::new()
        $inlineRecords = [System.Collections.Generic.List[object]]::new()
        $shapeRecords = [System.Collections.Generic.List[object]]::new()
        $sectionRecords = [System.Collections.Generic.List[object]]::new()
        $fieldRecords = [System.Collections.Generic.List[object]]::new()

        try {
            Write-ProgressLog "Abriendo solo lectura y sin reparacion: $fullPath"
            $doc = $documents.OpenNoRepairDialog($fullPath, $false, $true, $false)
            if ($null -eq $doc) { throw "Word devolvio un documento nulo al abrir: $fullPath" }
            Write-ProgressLog "Abierto: $fullPath"

            $content = $doc.Content
            if ($null -eq $content) { throw "Word devolvio Content nulo: $fullPath" }
            $fullText = [string]$content.Text
            Write-ProgressLog "Texto extraido. Longitud=$($fullText.Length)"
            $stem = [IO.Path]::GetFileNameWithoutExtension($fullPath)
            $textOut = Join-Path $outFull ($stem + '_' + $hash.Substring(0, 12) + '_fulltext.txt')
            [IO.File]::WriteAllText($textOut, $fullText, $utf8NoBom)
            Write-ProgressLog "Texto guardado: $textOut"

            $paragraphs = $doc.Paragraphs
            for ($i = 1; $i -le $paragraphs.Count; $i++) {
                $p = $null
                $range = $null
                try {
                    $p = $paragraphs.Item($i)
                    $range = $p.Range
                    $text = Clean-WordText ([string]$range.Text)
                    $styleName = Get-StyleName $range
                    $outline = $null
                    try { $outline = [int]$p.OutlineLevel } catch {}
                    if ($text.Length -gt 0) {
                        $paragraphRecords.Add([ordered]@{
                            index = $i
                            style = $styleName
                            outline_level = $outline
                            text = $text
                        })
                    }
                    $isHeading = (($null -ne $outline -and $outline -ge 1 -and $outline -le 9) -or $styleName -match '^(Heading|T[ií]tulo)\s*[1-9]')
                    if ($isHeading -and $text.Length -gt 0) {
                        $headings.Add([ordered]@{
                            index = $i
                            page = Get-PageNumber $range
                            style = $styleName
                            outline_level = $outline
                            text = $text
                        })
                    }
                    $isCaption = ($styleName -match '(Caption|Ep[ií]grafe|Descripci[oó]n)' -or $text -match '^(Figura|Tabla|Figure|Table|Cuadro|Ilustraci[oó]n)\s*\d+')
                    if ($isCaption -and $text.Length -gt 0) {
                        $captions.Add([ordered]@{
                            index = $i
                            page = Get-PageNumber $range
                            style = $styleName
                            text = $text
                        })
                    }
                }
                finally {
                    Release-ComObject $range
                    Release-ComObject $p
                }
            }

            $tables = $doc.Tables
            for ($i = 1; $i -le $tables.Count; $i++) {
                $table = $null
                $range = $null
                try {
                    $table = $tables.Item($i)
                    $range = $table.Range
                    $rowCount = $null
                    $colCount = $null
                    try { $rowCount = [int]$table.Rows.Count } catch {}
                    try { $colCount = [int]$table.Columns.Count } catch {}
                    $sample = Clean-WordText ([string]$range.Text)
                    if ($sample.Length -gt 700) { $sample = $sample.Substring(0, 700) }
                    $tableRecords.Add([ordered]@{
                        index = $i
                        page = Get-PageNumber $range
                        rows = $rowCount
                        columns = $colCount
                        allow_autofit = $(try { [bool]$table.AllowAutoFit } catch { $null })
                        sample = $sample
                    })
                }
                finally {
                    Release-ComObject $range
                    Release-ComObject $table
                }
            }

            $inlineShapes = $doc.InlineShapes
            for ($i = 1; $i -le $inlineShapes.Count; $i++) {
                $shape = $null
                $range = $null
                try {
                    $shape = $inlineShapes.Item($i)
                    $range = $shape.Range
                    $inlineRecords.Add([ordered]@{
                        index = $i
                        page = Get-PageNumber $range
                        type = $(try { [int]$shape.Type } catch { $null })
                        width_points = $(try { [math]::Round([double]$shape.Width, 2) } catch { $null })
                        height_points = $(try { [math]::Round([double]$shape.Height, 2) } catch { $null })
                        title = $(try { [string]$shape.Title } catch { '' })
                        alt_text = $(try { [string]$shape.AlternativeText } catch { '' })
                    })
                }
                finally {
                    Release-ComObject $range
                    Release-ComObject $shape
                }
            }

            $shapes = $doc.Shapes
            for ($i = 1; $i -le $shapes.Count; $i++) {
                $shape = $null
                $anchor = $null
                try {
                    $shape = $shapes.Item($i)
                    try { $anchor = $shape.Anchor } catch {}
                    $shapeRecords.Add([ordered]@{
                        index = $i
                        page = $(if ($null -ne $anchor) { Get-PageNumber $anchor } else { $null })
                        name = $(try { [string]$shape.Name } catch { '' })
                        type = $(try { [int]$shape.Type } catch { $null })
                        width_points = $(try { [math]::Round([double]$shape.Width, 2) } catch { $null })
                        height_points = $(try { [math]::Round([double]$shape.Height, 2) } catch { $null })
                        wrap_type = $(try { [int]$shape.WrapFormat.Type } catch { $null })
                        title = $(try { [string]$shape.Title } catch { '' })
                        alt_text = $(try { [string]$shape.AlternativeText } catch { '' })
                    })
                }
                finally {
                    Release-ComObject $anchor
                    Release-ComObject $shape
                }
            }

            $sections = $doc.Sections
            for ($i = 1; $i -le $sections.Count; $i++) {
                $section = $null
                $setup = $null
                try {
                    $section = $sections.Item($i)
                    $setup = $section.PageSetup
                    $sectionRecords.Add([ordered]@{
                        index = $i
                        orientation = $(try { [int]$setup.Orientation } catch { $null })
                        width_points = $(try { [math]::Round([double]$setup.PageWidth, 2) } catch { $null })
                        height_points = $(try { [math]::Round([double]$setup.PageHeight, 2) } catch { $null })
                        margin_top = $(try { [math]::Round([double]$setup.TopMargin, 2) } catch { $null })
                        margin_bottom = $(try { [math]::Round([double]$setup.BottomMargin, 2) } catch { $null })
                        margin_left = $(try { [math]::Round([double]$setup.LeftMargin, 2) } catch { $null })
                        margin_right = $(try { [math]::Round([double]$setup.RightMargin, 2) } catch { $null })
                        different_first_page = $(try { [bool]$setup.DifferentFirstPageHeaderFooter } catch { $null })
                        odd_even = $(try { [bool]$setup.OddAndEvenPagesHeaderFooter } catch { $null })
                    })
                }
                finally {
                    Release-ComObject $setup
                    Release-ComObject $section
                }
            }

            $fields = $doc.Fields
            for ($i = 1; $i -le $fields.Count; $i++) {
                $field = $null
                $codeRange = $null
                $resultRange = $null
                try {
                    $field = $fields.Item($i)
                    $codeRange = $field.Code
                    $resultRange = $field.Result
                    $fieldRecords.Add([ordered]@{
                        index = $i
                        type = $(try { [int]$field.Type } catch { $null })
                        locked = $(try { [bool]$field.Locked } catch { $null })
                        code = Clean-WordText ([string]$codeRange.Text)
                        result = Clean-WordText ([string]$resultRange.Text)
                    })
                }
                finally {
                    Release-ComObject $resultRange
                    Release-ComObject $codeRange
                    Release-ComObject $field
                }
            }

            $docRecord = [ordered]@{
                name = $item.Name
                path = $fullPath
                size_bytes = [int64]$item.Length
                modified = $item.LastWriteTime.ToString('o')
                sha256 = $hash
                read_only_open = [bool]$doc.ReadOnly
                compatibility_mode = $(try { [int]$doc.CompatibilityMode } catch { $null })
                pages = [int]$doc.ComputeStatistics(2, $false)
                words = [int]$doc.ComputeStatistics(0, $false)
                characters = [int]$doc.ComputeStatistics(3, $false)
                paragraphs = [int]$paragraphs.Count
                tables = [int]$tables.Count
                inline_shapes = [int]$inlineShapes.Count
                floating_shapes = [int]$shapes.Count
                sections = [int]$sections.Count
                fields = [int]$fields.Count
                tables_of_contents = $(try { [int]$doc.TablesOfContents.Count } catch { $null })
                tables_of_figures = $(try { [int]$doc.TablesOfFigures.Count } catch { $null })
                hyperlinks = $(try { [int]$doc.Hyperlinks.Count } catch { $null })
                bookmarks = $(try { [int]$doc.Bookmarks.Count } catch { $null })
                footnotes = $(try { [int]$doc.Footnotes.Count } catch { $null })
                endnotes = $(try { [int]$doc.Endnotes.Count } catch { $null })
                comments = $(try { [int]$doc.Comments.Count } catch { $null })
                revisions = $(try { [int]$doc.Revisions.Count } catch { $null })
                headings = @($headings)
                captions = @($captions)
                table_details = @($tableRecords)
                inline_shape_details = @($inlineRecords)
                floating_shape_details = @($shapeRecords)
                section_details = @($sectionRecords)
                field_details = @($fieldRecords)
                paragraph_details = @($paragraphRecords)
                suspicious_text = [ordered]@{
                    error_reference = [regex]::Matches($fullText, '(?i)(Error!|¡Error!|referencia no encontrada|origen de la referencia no encontrado)').Count
                    todo_pending = [regex]::Matches($fullText, '(?i)\b(TODO|TBD|PENDIENTE|POR COMPLETAR)\b').Count
                    client_event = [regex]::Matches($fullText, '(?i)CLIENT_EVENT').Count
                    server_event = [regex]::Matches($fullText, '(?i)SERVER_EVENT').Count
                    discord = [regex]::Matches($fullText, '(?i)Discord').Count
                    jsonl = [regex]::Matches($fullText, '(?i)JSONL').Count
                    etw = [regex]::Matches($fullText, '(?i)\bETW\b').Count
                    tracknetwork = [regex]::Matches($fullText, '(?i)TrackNetwork').Count
                    wazuh = [regex]::Matches($fullText, '(?i)Wazuh').Count
                }
            }
            $report.documents += $docRecord
            Write-ProgressLog "Auditoria completada: $fullPath"
        }
        finally {
            if ($null -ne $doc) {
                try { $doc.Close(0) } catch {}
            }
            Release-ComObject $fields
            Release-ComObject $sections
            Release-ComObject $shapes
            Release-ComObject $inlineShapes
            Release-ComObject $tables
            Release-ComObject $paragraphs
            Release-ComObject $content
            Release-ComObject $doc
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }
}
catch {
    $errorRecord = [ordered]@{
        generated_at = (Get-Date).ToString('o')
        message = $_.Exception.Message
        exception_type = $_.Exception.GetType().FullName
        position = $_.InvocationInfo.PositionMessage
        script_stack = $_.ScriptStackTrace
    }
    $errorPath = Join-Path $outFull 'word_candidates_audit_error.json'
    [IO.File]::WriteAllText($errorPath, ($errorRecord | ConvertTo-Json -Depth 6), $utf8NoBom)
    Write-ProgressLog "ERROR: $($_.Exception.Message) | $($_.InvocationInfo.PositionMessage)"
    throw
}
finally {
    Release-ComObject $documents
    if ($null -ne $word) {
        try { $word.Quit(0) } catch {}
    }
    Release-ComObject $word
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 2
}

$afterPids = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$orphanPids = @($afterPids | Where-Object { $baselinePids -notcontains $_ })
$report.after_winword_pids = $afterPids
$report.new_winword_pids_after_quit = $orphanPids

$reportPath = Join-Path $outFull 'word_candidates_audit.json'
[IO.File]::WriteAllText($reportPath, ($report | ConvertTo-Json -Depth 12), $utf8NoBom)

[pscustomobject]@{
    Report = $reportPath
    WordVersion = $report.word_version
    WordBuild = $report.word_build
    Documents = $report.documents.Count
    NewWinwordPidsAfterQuit = ($orphanPids -join ',')
} | Format-List

if ($orphanPids.Count -gt 0) {
    throw "Quedaron procesos WINWORD nuevos tras la auditoría: $($orphanPids -join ',')"
}
