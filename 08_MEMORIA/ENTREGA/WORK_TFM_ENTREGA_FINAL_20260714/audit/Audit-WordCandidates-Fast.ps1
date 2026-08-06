[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ManifestPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$wUri = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$wpUri = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
$rUri = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$bUri = 'http://schemas.openxmlformats.org/officeDocument/2006/bibliography'

function Release-ComObject {
    param([object]$Object)
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Clean-Text {
    param([AllowNull()][string]$Text)
    if ($null -eq $Text) { return '' }
    return (($Text -replace "`r", ' ' -replace "`a", ' ' -replace "`v", ' ' -replace "`t", ' ') -replace '\s+', ' ').Trim()
}

function Read-ZipXml {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [string]$EntryName
    )
    $entry = $Zip.GetEntry($EntryName)
    if ($null -eq $entry) { return $null }
    $stream = $null
    try {
        $stream = $entry.Open()
        $xml = New-Object System.Xml.XmlDocument
        $xml.PreserveWhitespace = $true
        $xml.Load($stream)
        return ,$xml
    }
    finally {
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function New-WordNs {
    param([System.Xml.XmlDocument]$Xml)
    $ns = New-Object System.Xml.XmlNamespaceManager($Xml.NameTable)
    $ns.AddNamespace('w', $wUri)
    $ns.AddNamespace('wp', $wpUri)
    $ns.AddNamespace('r', $rUri)
    $ns.AddNamespace('b', $bUri)
    return ,$ns
}

function Get-AttributeNs {
    param(
        [System.Xml.XmlNode]$Node,
        [string]$LocalName,
        [string]$NamespaceUri
    )
    if ($null -eq $Node) { return $null }
    $attr = $Node.Attributes.GetNamedItem($LocalName, $NamespaceUri)
    if ($null -eq $attr) { return $null }
    return [string]$attr.Value
}

function Get-ParagraphText {
    param(
        [System.Xml.XmlNode]$Paragraph,
        [System.Xml.XmlNamespaceManager]$Ns
    )
    $builder = New-Object System.Text.StringBuilder
    foreach ($node in @($Paragraph.SelectNodes('.//w:t|.//w:tab|.//w:br', $Ns))) {
        switch ($node.LocalName) {
            't' { [void]$builder.Append($node.InnerText) }
            'tab' { [void]$builder.Append("`t") }
            'br' { [void]$builder.Append("`n") }
        }
    }
    return Clean-Text $builder.ToString()
}

function Resolve-OutlineLevel {
    param(
        [string]$StyleId,
        [hashtable]$StyleMap
    )
    $seen = @{}
    $current = $StyleId
    while ($current -and $StyleMap.ContainsKey($current) -and -not $seen.ContainsKey($current)) {
        $seen[$current] = $true
        $entry = $StyleMap[$current]
        if ($null -ne $entry.outline_level) { return $entry.outline_level }
        $current = [string]$entry.based_on
    }
    return $null
}

function Get-OpenXmlAudit {
    param(
        [string]$Path,
        [string]$TextOutputPath
    )
    $zip = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
        $documentXml = Read-ZipXml $zip 'word/document.xml'
        if ($null -eq $documentXml) { throw "Falta word/document.xml en $Path" }
        $stylesXml = Read-ZipXml $zip 'word/styles.xml'
        $ns = New-WordNs $documentXml
        $styleMap = @{}

        if ($null -ne $stylesXml) {
            $stylesNs = New-WordNs $stylesXml
            foreach ($style in @($stylesXml.SelectNodes('//w:style[@w:type="paragraph"]', $stylesNs))) {
                $id = Get-AttributeNs $style 'styleId' $wUri
                if (-not $id) { continue }
                $nameNode = $style.SelectSingleNode('./w:name', $stylesNs)
                $basedNode = $style.SelectSingleNode('./w:basedOn', $stylesNs)
                $outlineNode = $style.SelectSingleNode('./w:pPr/w:outlineLvl', $stylesNs)
                $outline = $null
                if ($null -ne $outlineNode) {
                    $raw = Get-AttributeNs $outlineNode 'val' $wUri
                    if ($raw -match '^\d+$') { $outline = [int]$raw }
                }
                $styleMap[$id] = [ordered]@{
                    id = $id
                    name = $(if ($null -ne $nameNode) { Get-AttributeNs $nameNode 'val' $wUri } else { '' })
                    based_on = $(if ($null -ne $basedNode) { Get-AttributeNs $basedNode 'val' $wUri } else { '' })
                    outline_level = $outline
                }
            }
        }

        $paragraphRecords = [System.Collections.Generic.List[object]]::new()
        $headings = [System.Collections.Generic.List[object]]::new()
        $captions = [System.Collections.Generic.List[object]]::new()
        $textBuilder = New-Object System.Text.StringBuilder
        $paragraphNodes = @($documentXml.SelectNodes('//w:body//w:p', $ns))
        $paragraphIndex = 0
        foreach ($p in $paragraphNodes) {
            $paragraphIndex++
            $text = Get-ParagraphText $p $ns
            $pStyleNode = $p.SelectSingleNode('./w:pPr/w:pStyle', $ns)
            $styleId = $(if ($null -ne $pStyleNode) { Get-AttributeNs $pStyleNode 'val' $wUri } else { '' })
            $styleName = $(if ($styleId -and $styleMap.ContainsKey($styleId)) { [string]$styleMap[$styleId].name } else { $styleId })
            $outlineNode = $p.SelectSingleNode('./w:pPr/w:outlineLvl', $ns)
            $outline = $null
            if ($null -ne $outlineNode) {
                $rawOutline = Get-AttributeNs $outlineNode 'val' $wUri
                if ($rawOutline -match '^\d+$') { $outline = [int]$rawOutline }
            }
            if ($null -eq $outline -and $styleId) { $outline = Resolve-OutlineLevel $styleId $styleMap }
            if ($text.Length -gt 0) {
                [void]$textBuilder.AppendLine($text)
                $paragraphRecords.Add([ordered]@{
                    index = $paragraphIndex
                    style_id = $styleId
                    style_name = $styleName
                    outline_level_zero_based = $outline
                    text = $text
                })
            }
            $headingStyle = (($styleId -match '(?i)^Heading[1-9]$') -or ($styleName -match '(?i)^(Heading|T[ií]tulo)\s*[1-9]$'))
            if ($text.Length -gt 0 -and ($headingStyle -or ($null -ne $outline -and $outline -ge 0 -and $outline -le 8))) {
                $headings.Add([ordered]@{
                    index = $paragraphIndex
                    style_id = $styleId
                    style_name = $styleName
                    level = $(if ($null -ne $outline) { $outline + 1 } else { $null })
                    text = $text
                })
            }
            if ($text.Length -gt 0 -and (($styleId -match '(?i)Caption') -or ($styleName -match '(?i)(Caption|Ep[ií]grafe|Descripci[oó]n)') -or ($text -match '^(Figura|Tabla|Figure|Table|Cuadro|Ilustraci[oó]n)\s*\d+'))) {
                $captions.Add([ordered]@{
                    index = $paragraphIndex
                    style_id = $styleId
                    style_name = $styleName
                    text = $text
                })
            }
        }
        [IO.File]::WriteAllText($TextOutputPath, $textBuilder.ToString(), $utf8NoBom)

        $tableRecords = [System.Collections.Generic.List[object]]::new()
        $tableIndex = 0
        foreach ($table in @($documentXml.SelectNodes('//w:tbl', $ns))) {
            $tableIndex++
            $rows = @($table.SelectNodes('./w:tr', $ns))
            $gridCols = @($table.SelectNodes('./w:tblGrid/w:gridCol', $ns))
            $tableTextBuilder = New-Object System.Text.StringBuilder
            foreach ($t in @($table.SelectNodes('.//w:t', $ns))) {
                if ($tableTextBuilder.Length -gt 0) { [void]$tableTextBuilder.Append(' | ') }
                [void]$tableTextBuilder.Append($t.InnerText)
                if ($tableTextBuilder.Length -gt 900) { break }
            }
            $sample = Clean-Text $tableTextBuilder.ToString()
            if ($sample.Length -gt 700) { $sample = $sample.Substring(0, 700) }
            $tableRecords.Add([ordered]@{
                index = $tableIndex
                rows = $rows.Count
                grid_columns = $gridCols.Count
                sample = $sample
            })
        }

        $sectionRecords = [System.Collections.Generic.List[object]]::new()
        $sectionIndex = 0
        foreach ($sect in @($documentXml.SelectNodes('//w:sectPr', $ns))) {
            $sectionIndex++
            $pgSz = $sect.SelectSingleNode('./w:pgSz', $ns)
            $pgMar = $sect.SelectSingleNode('./w:pgMar', $ns)
            $sectionRecords.Add([ordered]@{
                index = $sectionIndex
                width_twips = $(if ($null -ne $pgSz) { Get-AttributeNs $pgSz 'w' $wUri } else { $null })
                height_twips = $(if ($null -ne $pgSz) { Get-AttributeNs $pgSz 'h' $wUri } else { $null })
                orientation = $(if ($null -ne $pgSz) { Get-AttributeNs $pgSz 'orient' $wUri } else { $null })
                margin_top_twips = $(if ($null -ne $pgMar) { Get-AttributeNs $pgMar 'top' $wUri } else { $null })
                margin_bottom_twips = $(if ($null -ne $pgMar) { Get-AttributeNs $pgMar 'bottom' $wUri } else { $null })
                margin_left_twips = $(if ($null -ne $pgMar) { Get-AttributeNs $pgMar 'left' $wUri } else { $null })
                margin_right_twips = $(if ($null -ne $pgMar) { Get-AttributeNs $pgMar 'right' $wUri } else { $null })
            })
        }

        $fieldCodes = [System.Collections.Generic.List[string]]::new()
        foreach ($p in $paragraphNodes) {
            $parts = @($p.SelectNodes('.//w:instrText', $ns) | ForEach-Object { $_.InnerText })
            if ($parts.Count -gt 0) {
                $code = Clean-Text ($parts -join ' ')
                if ($code) { $fieldCodes.Add($code) }
            }
            foreach ($simple in @($p.SelectNodes('.//w:fldSimple', $ns))) {
                $code = Clean-Text (Get-AttributeNs $simple 'instr' $wUri)
                if ($code) { $fieldCodes.Add($code) }
            }
        }

        $mediaEntries = @($zip.Entries | Where-Object { $_.FullName -like 'word/media/*' })
        $customSourceCount = 0
        $customSourceTitles = [System.Collections.Generic.List[string]]::new()
        foreach ($entry in @($zip.Entries | Where-Object { $_.FullName -like 'customXml/item*.xml' -and $_.FullName -notlike '*itemProps*' })) {
            $xml = Read-ZipXml $zip $entry.FullName
            if ($null -eq $xml) { continue }
            $sourceNodes = @($xml.SelectNodes("//*[local-name()='Source' and namespace-uri()='$bUri']"))
            $customSourceCount += $sourceNodes.Count
            foreach ($source in $sourceNodes) {
                $titleNode = $source.SelectSingleNode("./*[local-name()='Title' and namespace-uri()='$bUri']")
                if ($null -ne $titleNode -and $titleNode.InnerText) { $customSourceTitles.Add((Clean-Text $titleNode.InnerText)) }
            }
        }

        $fullText = $textBuilder.ToString()
        return [ordered]@{
            paragraphs_nonempty = $paragraphRecords.Count
            headings = @($headings)
            captions = @($captions)
            tables = @($tableRecords)
            inline_drawings = @($documentXml.SelectNodes('//wp:inline', $ns)).Count
            floating_drawings = @($documentXml.SelectNodes('//wp:anchor', $ns)).Count
            legacy_pictures = @($documentXml.SelectNodes("//*[local-name()='pict']", $ns)).Count
            media_parts = $mediaEntries.Count
            section_details = @($sectionRecords)
            field_codes = @($fieldCodes)
            toc_field_count = @($fieldCodes | Where-Object { $_ -match '(?i)^\s*TOC\b' }).Count
            ref_field_count = @($fieldCodes | Where-Object { $_ -match '(?i)^\s*(REF|PAGEREF)\b' }).Count
            seq_table_field_count = @($fieldCodes | Where-Object { $_ -match '(?i)^\s*SEQ\s+(Tabla|Table)\b' }).Count
            seq_figure_field_count = @($fieldCodes | Where-Object { $_ -match '(?i)^\s*SEQ\s+(Figura|Figure)\b' }).Count
            page_breaks = @($documentXml.SelectNodes('//w:br[@w:type="page"]', $ns)).Count
            bookmarks = @($documentXml.SelectNodes('//w:bookmarkStart', $ns)).Count
            hyperlinks = @($documentXml.SelectNodes('//w:hyperlink', $ns)).Count
            bibliography_source_count = $customSourceCount
            bibliography_titles = @($customSourceTitles)
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
    }
    finally {
        if ($null -ne $zip) { $zip.Dispose() }
    }
}

$manifestFull = (Resolve-Path -LiteralPath $ManifestPath).Path
$manifest = Get-Content -Raw -LiteralPath $manifestFull | ConvertFrom-Json
if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir | Out-Null }
$outFull = (Resolve-Path -LiteralPath $OutputDir).Path
$progressPath = Join-Path $outFull 'word_candidates_fast_progress.log'
[IO.File]::WriteAllText($progressPath, "INICIO $(Get-Date -Format o)`r`n", $utf8NoBom)
function Log([string]$Message) { [IO.File]::AppendAllText($progressPath, "$(Get-Date -Format o) | $Message`r`n", $utf8NoBom) }

$baselinePids = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$fatalError = $null
$word = $null
$documents = $null
$report = [ordered]@{
    generated_at = (Get-Date).ToString('o')
    word_version = $null
    word_build = $null
    baseline_winword_pids = $baselinePids
    created_winword_pids = @()
    documents = @()
    after_winword_pids = @()
    new_winword_pids_after_quit = @()
    forced_cleanup_pids = @()
    final_winword_pids = @()
}

try {
    Log 'Creando Word.Application'
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    try { $word.AutomationSecurity = 3 } catch {}
    try { $word.Options.UpdateLinksAtOpen = $false } catch {}
    try { $word.Options.SaveNormalPrompt = $false } catch {}
    $report.word_version = [string]$word.Version
    $report.word_build = [string]$word.Build
    $documents = $word.Documents
    $report.created_winword_pids = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id | Where-Object { $baselinePids -notcontains $_ })
    Log "Word listo: $($report.word_version) build $($report.word_build)"

    foreach ($inputPath in @($manifest.documents)) {
        if (-not (Test-Path -LiteralPath $inputPath)) { throw "No existe: $inputPath" }
        $fullPath = (Resolve-Path -LiteralPath $inputPath).Path
        $item = Get-Item -LiteralPath $fullPath
        $hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $fullPath).Hash
        $stem = [IO.Path]::GetFileNameWithoutExtension($fullPath)
        $textOut = Join-Path $outFull ($stem + '_' + $hash.Substring(0, 12) + '_fulltext.txt')
        Log "OOXML: $fullPath"
        $openXml = Get-OpenXmlAudit -Path $fullPath -TextOutputPath $textOut

        $doc = $null
        try {
            Log "Word abre solo lectura, sin reparacion: $fullPath"
            $doc = $documents.OpenNoRepairDialog($fullPath, $false, $true, $false)
            if ($null -eq $doc) { throw "Word devolvio documento nulo: $fullPath" }
            $record = [ordered]@{
                name = $item.Name
                path = $fullPath
                size_bytes = [int64]$item.Length
                modified = $item.LastWriteTime.ToString('o')
                sha256 = $hash
                read_only_open = [bool]$doc.ReadOnly
                compatibility_mode = $(try { [int]$doc.CompatibilityMode } catch { $null })
                pages_word = $(try { [int]$doc.ComputeStatistics(2, $false) } catch { $null })
                words_word = $(try { [int]$doc.ComputeStatistics(0, $false) } catch { $null })
                characters_word = $(try { [int]$doc.ComputeStatistics(3, $false) } catch { $null })
                paragraphs_word = $(try { [int]$doc.Paragraphs.Count } catch { $null })
                tables_word = $(try { [int]$doc.Tables.Count } catch { $null })
                inline_shapes_word = $(try { [int]$doc.InlineShapes.Count } catch { $null })
                floating_shapes_word = $(try { [int]$doc.Shapes.Count } catch { $null })
                sections_word = $(try { [int]$doc.Sections.Count } catch { $null })
                fields_word = $(try { [int]$doc.Fields.Count } catch { $null })
                tables_of_contents_word = $(try { [int]$doc.TablesOfContents.Count } catch { $null })
                tables_of_figures_word = $(try { [int]$doc.TablesOfFigures.Count } catch { $null })
                hyperlinks_word = $(try { [int]$doc.Hyperlinks.Count } catch { $null })
                bookmarks_word = $(try { [int]$doc.Bookmarks.Count } catch { $null })
                footnotes_word = $(try { [int]$doc.Footnotes.Count } catch { $null })
                endnotes_word = $(try { [int]$doc.Endnotes.Count } catch { $null })
                comments_word = $(try { [int]$doc.Comments.Count } catch { $null })
                revisions_word = $(try { [int]$doc.Revisions.Count } catch { $null })
                openxml = $openXml
            }
            $report.documents += $record
            Log "Completado: $fullPath"
        }
        finally {
            if ($null -ne $doc) { try { $doc.Close(0) } catch {} }
            Release-ComObject $doc
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }
}
catch {
    $fatalError = $_
    $errorRecord = [ordered]@{
        generated_at = (Get-Date).ToString('o')
        message = $_.Exception.Message
        exception_type = $_.Exception.GetType().FullName
        position = $_.InvocationInfo.PositionMessage
        script_stack = $_.ScriptStackTrace
    }
    [IO.File]::WriteAllText((Join-Path $outFull 'word_candidates_fast_error.json'), ($errorRecord | ConvertTo-Json -Depth 6), $utf8NoBom)
    Log "ERROR: $($_.Exception.Message)"
}
finally {
    Release-ComObject $documents
    if ($null -ne $word) { try { $word.Quit(0) } catch {} }
    Release-ComObject $word
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 2
}

$afterPids = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$orphanPids = @($afterPids | Where-Object { $report.created_winword_pids -contains $_ })
$report.after_winword_pids = $afterPids
$report.new_winword_pids_after_quit = $orphanPids
if ($orphanPids.Count -gt 0) {
    foreach ($pidToStop in $orphanPids) {
        try { Stop-Process -Id $pidToStop -Force -ErrorAction Stop } catch {}
    }
    Start-Sleep -Seconds 2
    $report.forced_cleanup_pids = $orphanPids
}
$report.final_winword_pids = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$reportPath = Join-Path $outFull 'word_candidates_audit.json'
[IO.File]::WriteAllText($reportPath, ($report | ConvertTo-Json -Depth 14), $utf8NoBom)

[pscustomobject]@{
    Report = $reportPath
    WordVersion = $report.word_version
    WordBuild = $report.word_build
    Documents = $report.documents.Count
    NewWinwordPidsAfterQuit = ($orphanPids -join ',')
    ForcedCleanupPids = ($report.forced_cleanup_pids -join ',')
    FinalWinwordPids = ($report.final_winword_pids -join ',')
} | Format-List

if (@($report.final_winword_pids | Where-Object { $report.created_winword_pids -contains $_ }).Count -gt 0) {
    throw "Persisten WINWORD creados por la auditoria: $($report.final_winword_pids -join ',')"
}
if ($null -ne $fatalError) {
    throw "Auditoria fallida tras limpieza COM: $($fatalError.Exception.Message)"
}
