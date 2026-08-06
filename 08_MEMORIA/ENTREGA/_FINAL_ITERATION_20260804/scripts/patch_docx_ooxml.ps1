param(
    [Parameter(Mandatory = $true)]
    [string]$DocumentPath,
    [Parameter(Mandatory = $true)]
    [string]$LogPath
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-EntryText {
    param([System.IO.Compression.ZipArchiveEntry]$Entry)
    $reader = New-Object System.IO.StreamReader($Entry.Open(), [System.Text.Encoding]::UTF8, $true)
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Write-XmlEntry {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [string]$Name,
        [xml]$Xml
    )
    $old = $Zip.GetEntry($Name)
    if ($null -ne $old) { $old.Delete() }
    $entry = $Zip.CreateEntry($Name, [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
    $settings.Indent = $false
    $settings.OmitXmlDeclaration = $false
    $writer = [System.Xml.XmlWriter]::Create($stream, $settings)
    try { $Xml.Save($writer) } finally { $writer.Dispose(); $stream.Dispose() }
}

$zip = [System.IO.Compression.ZipFile]::Open($DocumentPath, [System.IO.Compression.ZipArchiveMode]::Update)
try {
    $documentEntry = $zip.GetEntry('word/document.xml')
    $settingsEntry = $zip.GetEntry('word/settings.xml')
    if ($null -eq $documentEntry -or $null -eq $settingsEntry) {
        throw 'El DOCX no contiene document.xml o settings.xml.'
    }

    [xml]$documentXml = Read-EntryText $documentEntry
    [xml]$settingsXml = Read-EntryText $settingsEntry
    $w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    $docNs = New-Object System.Xml.XmlNamespaceManager($documentXml.NameTable)
    $docNs.AddNamespace('w', $w)
    $blankCaptionsRemoved = 0
    $paragraphs = @($documentXml.SelectNodes('//w:p[w:pPr/w:pStyle[@w:val="Descripcin"]]', $docNs))
    foreach ($paragraph in $paragraphs) {
        $text = ''
        foreach ($node in @($paragraph.SelectNodes('.//w:t', $docNs))) {
            $text += [string]$node.InnerText
        }
        $protected = $paragraph.SelectSingleNode('.//w:drawing|.//w:pict|.//w:fldSimple|.//w:instrText|.//w:object', $docNs)
        if ([string]::IsNullOrWhiteSpace($text) -and $null -eq $protected) {
            [void]$paragraph.ParentNode.RemoveChild($paragraph)
            $blankCaptionsRemoved++
        }
    }

    $mathTypeResiduesRemoved = 0
    $equationNode = $documentXml.SelectSingleNode('//w:instrText[contains(., "Equation Chapter")]', $docNs)
    if ($null -ne $equationNode) {
        $equationRun = $equationNode.SelectSingleNode('ancestor::w:r[1]', $docNs)
        $macroRun = $equationRun.PreviousSibling
        while ($null -ne $macroRun -and $null -eq $macroRun.SelectSingleNode('./w:instrText[contains(., "MTEditEquationSection")]', $docNs)) {
            $macroRun = $macroRun.PreviousSibling
        }
        $outerBeginRun = if ($null -ne $macroRun) { $macroRun.PreviousSibling } else { $null }
        if ($null -ne $outerBeginRun -and $null -ne $outerBeginRun.SelectSingleNode('./w:fldChar[@w:fldCharType="begin"]', $docNs)) {
            $runsToRemove = [System.Collections.Generic.List[System.Xml.XmlNode]]::new()
            $depth = 0
            $cursor = $outerBeginRun
            while ($null -ne $cursor) {
                $next = $cursor.NextSibling
                if ($cursor.LocalName -eq 'r') {
                    [void]$runsToRemove.Add($cursor)
                    $depth += @($cursor.SelectNodes('./w:fldChar[@w:fldCharType="begin"]', $docNs)).Count
                    $depth -= @($cursor.SelectNodes('./w:fldChar[@w:fldCharType="end"]', $docNs)).Count
                    if ($depth -eq 0) { break }
                }
                $cursor = $next
            }
            foreach ($run in $runsToRemove) {
                [void]$run.ParentNode.RemoveChild($run)
                $mathTypeResiduesRemoved++
            }
        }
    }

    $settingsNs = New-Object System.Xml.XmlNamespaceManager($settingsXml.NameTable)
    $settingsNs.AddNamespace('w', $w)
    $updateFields = $settingsXml.SelectSingleNode('/w:settings/w:updateFields', $settingsNs)
    if ($null -eq $updateFields) {
        $updateFields = $settingsXml.CreateElement('w', 'updateFields', $w)
        [void]$settingsXml.DocumentElement.AppendChild($updateFields)
    }
    $updateFields.SetAttribute('val', $w, 'true')

    Write-XmlEntry -Zip $zip -Name 'word/document.xml' -Xml $documentXml
    Write-XmlEntry -Zip $zip -Name 'word/settings.xml' -Xml $settingsXml
}
finally {
    $zip.Dispose()
}

$result = [pscustomobject]@{
    BlankCaptionParagraphsRemoved = $blankCaptionsRemoved
    MathTypeResidueRunsRemoved = $mathTypeResiduesRemoved
    UpdateFieldsOnOpen = $true
    SHA256 = (Get-FileHash -LiteralPath $DocumentPath -Algorithm SHA256).Hash
}
$result | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $LogPath -Encoding UTF8
$result | ConvertTo-Json -Depth 4
