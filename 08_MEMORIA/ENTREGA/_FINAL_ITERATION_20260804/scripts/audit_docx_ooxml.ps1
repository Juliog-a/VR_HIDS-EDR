param(
    [Parameter(Mandatory = $true)]
    [string]$DocxPath,
    [Parameter(Mandatory = $true)]
    [string]$OutPath
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-ZipText {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [string]$EntryName
    )
    $entry = $Zip.GetEntry($EntryName)
    if ($null -eq $entry) { return $null }
    $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8, $true)
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Get-NodeText {
    param(
        [System.Xml.XmlNode]$Node,
        [System.Xml.XmlNamespaceManager]$Ns
    )
    $parts = [System.Collections.Generic.List[string]]::new()
    foreach ($child in @($Node.SelectNodes('.//w:t|.//w:tab|.//w:br|.//w:cr', $Ns))) {
        switch ($child.LocalName) {
            't' { $parts.Add([string]$child.InnerText) }
            'tab' { $parts.Add("`t") }
            default { $parts.Add("`n") }
        }
    }
    return (($parts -join '') -replace "`a", '').Trim()
}

function Get-StyleId {
    param(
        [System.Xml.XmlNode]$Paragraph,
        [System.Xml.XmlNamespaceManager]$Ns
    )
    $node = $Paragraph.SelectSingleNode('./w:pPr/w:pStyle', $Ns)
    if ($null -eq $node) { return '' }
    return [string]$node.GetAttribute('val', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
}

$zip = [System.IO.Compression.ZipFile]::OpenRead($DocxPath)
try {
    [xml]$docXml = Read-ZipText -Zip $zip -EntryName 'word/document.xml'
    [xml]$stylesXml = Read-ZipText -Zip $zip -EntryName 'word/styles.xml'
    [xml]$settingsXml = Read-ZipText -Zip $zip -EntryName 'word/settings.xml'
    [xml]$relsXml = Read-ZipText -Zip $zip -EntryName 'word/_rels/document.xml.rels'

    $ns = New-Object System.Xml.XmlNamespaceManager($docXml.NameTable)
    $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    $ns.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')

    $stylesNs = New-Object System.Xml.XmlNamespaceManager($stylesXml.NameTable)
    $stylesNs.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    $styleMap = @{}
    foreach ($style in @($stylesXml.SelectNodes('//w:style', $stylesNs))) {
        $id = [string]$style.GetAttribute('styleId', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        $nameNode = $style.SelectSingleNode('./w:name', $stylesNs)
        $name = if ($null -ne $nameNode) {
            [string]$nameNode.GetAttribute('val', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        } else { $id }
        $styleMap[$id] = $name
    }

    $body = $docXml.SelectSingleNode('//w:body', $ns)
    $paragraphs = [System.Collections.Generic.List[object]]::new()
    $tables = [System.Collections.Generic.List[object]]::new()
    $captionCandidates = [System.Collections.Generic.List[object]]::new()
    $headings = [System.Collections.Generic.List[object]]::new()
    $sourceParagraphs = [System.Collections.Generic.List[object]]::new()
    $fieldCodes = [System.Collections.Generic.List[object]]::new()
    $blockIndex = 0
    $paragraphIndex = 0
    $tableIndex = 0
    $previousParagraphText = ''
    $previousParagraphStyle = ''

    foreach ($node in @($body.ChildNodes)) {
        $blockIndex++
        if ($node.LocalName -eq 'p') {
            $paragraphIndex++
            $text = Get-NodeText -Node $node -Ns $ns
            $styleId = Get-StyleId -Paragraph $node -Ns $ns
            $styleName = if ($styleMap.ContainsKey($styleId)) { [string]$styleMap[$styleId] } else { $styleId }
            $codes = @($node.SelectNodes('.//w:instrText', $ns) | ForEach-Object { ([string]$_.InnerText).Trim() })
            foreach ($code in $codes) {
                $fieldCodes.Add([pscustomobject]@{
                    ParagraphIndex = $paragraphIndex
                    StyleId = $styleId
                    StyleName = $styleName
                    Code = $code
                    Text = $text
                })
            }
            $record = [pscustomobject]@{
                ParagraphIndex = $paragraphIndex
                BlockIndex = $blockIndex
                StyleId = $styleId
                StyleName = $styleName
                Text = $text
                FieldCodes = $codes
            }
            $paragraphs.Add($record)
            if ($text -match '^(Tabla|Figura|Código)\s*\d*') {
                $captionCandidates.Add($record)
            }
            if ($styleName -match 'Heading|Título|title') {
                $headings.Add($record)
            }
            if ($text -match '^(?i)Fuente:') {
                $sourceParagraphs.Add($record)
            }
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                $previousParagraphText = $text
                $previousParagraphStyle = $styleName
            }
        }
        elseif ($node.LocalName -eq 'tbl') {
            $tableIndex++
            $gridWidths = @($node.SelectNodes('./w:tblGrid/w:gridCol', $ns) | ForEach-Object {
                $raw = $_.GetAttribute('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
                if ([string]::IsNullOrWhiteSpace($raw)) {
                    $raw = $_.GetAttribute('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w')
                }
                if ([string]::IsNullOrWhiteSpace($raw)) {
                    $raw = $_.GetAttribute('val', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
                }
                if ([string]::IsNullOrWhiteSpace($raw)) {
                    $raw = $_.Attributes | Where-Object { $_.LocalName -eq 'w' } | Select-Object -ExpandProperty Value -First 1
                }
                [int]$raw
            })
            $rows = @($node.SelectNodes('./w:tr', $ns))
            $maxCols = 0
            foreach ($row in $rows) {
                $count = @($row.SelectNodes('./w:tc', $ns)).Count
                if ($count -gt $maxCols) { $maxCols = $count }
            }
            $headerCells = @()
            if ($rows.Count -gt 0) {
                foreach ($cell in @($rows[0].SelectNodes('./w:tc', $ns))) {
                    $headerCells += (Get-NodeText -Node $cell -Ns $ns)
                }
            }
            $sample = Get-NodeText -Node $node -Ns $ns
            if ($sample.Length -gt 800) { $sample = $sample.Substring(0, 800) }
            $tables.Add([pscustomobject]@{
                TableIndex = $tableIndex
                BlockIndex = $blockIndex
                Rows = $rows.Count
                MaxColumns = $maxCols
                GridWidths = $gridWidths
                GridTotalTwips = ($gridWidths | Measure-Object -Sum).Sum
                Header = $headerCells
                PreviousParagraphText = $previousParagraphText
                PreviousParagraphStyle = $previousParagraphStyle
                Sample = $sample
                RepeatingHeader = ($null -ne $rows[0].SelectSingleNode('./w:trPr/w:tblHeader', $ns))
            })
        }
    }

    $seqCodes = @($fieldCodes | Where-Object { $_.Code -match '^SEQ\s+' })
    $seqSummary = @($seqCodes | ForEach-Object {
        if ($_.Code -match '^SEQ\s+([^\s\\]+)') { $matches[1] } else { 'UNKNOWN' }
    } | Group-Object | Sort-Object Name | ForEach-Object {
        [pscustomobject]@{Label=$_.Name;Count=$_.Count}
    })

    $bookmarkStarts = @($docXml.SelectNodes('//w:bookmarkStart', $ns))
    $bookmarkEnds = @($docXml.SelectNodes('//w:bookmarkEnd', $ns))

    $settingsNs = New-Object System.Xml.XmlNamespaceManager($settingsXml.NameTable)
    $settingsNs.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    $updateFields = $settingsXml.SelectSingleNode('//w:updateFields', $settingsNs)

    $relsNs = New-Object System.Xml.XmlNamespaceManager($relsXml.NameTable)
    $relsNs.AddNamespace('pr', 'http://schemas.openxmlformats.org/package/2006/relationships')
    $externalRelationships = @($relsXml.SelectNodes('//pr:Relationship[@TargetMode="External"]', $relsNs) | ForEach-Object {
        [pscustomobject]@{
            Id = [string]$_.Id
            Type = [string]$_.Type
            Target = [string]$_.Target
        }
    })

    $focusTables = @($tables | Where-Object { $_.TableIndex -in @(23,45,47,52,53,54,56,65,68) })
    $specificCaptions = @($captionCandidates | Where-Object {
        $_.Text -match '^(Tabla\s+(32|45|63|65)|Figura\s+15|Código\s+5)'
    })

    $result = [pscustomobject]@{
        Path = $DocxPath
        FileLength = (Get-Item -LiteralPath $DocxPath).Length
        SHA256 = (Get-FileHash -LiteralPath $DocxPath -Algorithm SHA256).Hash
        Paragraphs = $paragraphs.Count
        Tables = $tables.Count
        Captions = $captionCandidates.Count
        Headings = $headings.Count
        SourceParagraphs = $sourceParagraphs.Count
        SourceGeneric = @($sourceParagraphs | Where-Object { $_.Text -match '^(?i)Fuente:\s*(elaboración|elaboraci[oó]n|elaboración o captura)' }).Count
        FieldCodes = $fieldCodes.Count
        SeqSummary = $seqSummary
        BookmarkStarts = $bookmarkStarts.Count
        BookmarkEnds = $bookmarkEnds.Count
        UpdateFieldsOnOpen = if ($null -eq $updateFields) { $false } else {
            $val = [string]$updateFields.GetAttribute('val', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
            ($val -ne 'false' -and $val -ne '0')
        }
        ExternalRelationships = $externalRelationships
        SpecificCaptions = $specificCaptions
        FocusTables = $focusTables
        SourceParagraphDetails = $sourceParagraphs
        ParagraphDetails = $paragraphs
        CaptionsDetails = $captionCandidates
        HeadingsDetails = $headings
        FieldCodeDetails = $fieldCodes
        TablesDetails = $tables
    }
    $result | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $OutPath -Encoding UTF8
    [pscustomobject]@{
        Paragraphs = $result.Paragraphs
        Tables = $result.Tables
        Captions = $result.Captions
        Headings = $result.Headings
        SourceParagraphs = $result.SourceParagraphs
        SourceGeneric = $result.SourceGeneric
        SeqSummary = $result.SeqSummary
        UpdateFieldsOnOpen = $result.UpdateFieldsOnOpen
        ExternalRelationships = $result.ExternalRelationships
        SpecificCaptions = $result.SpecificCaptions
        FocusTables = $result.FocusTables
    } | ConvertTo-Json -Depth 8
}
finally {
    $zip.Dispose()
}
