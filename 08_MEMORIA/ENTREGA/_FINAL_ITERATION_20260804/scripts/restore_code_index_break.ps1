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
    $reader = [System.IO.StreamReader]::new($Entry.Open(), [System.Text.Encoding]::UTF8, $true)
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
    $settings = [System.Xml.XmlWriterSettings]::new()
    $settings.Encoding = [System.Text.UTF8Encoding]::new($false)
    $settings.Indent = $false
    $writer = [System.Xml.XmlWriter]::Create($stream, $settings)
    try { $Xml.Save($writer) } finally { $writer.Dispose(); $stream.Dispose() }
}

$resolvedDocument = (Resolve-Path -LiteralPath $DocumentPath).Path
$zip = [System.IO.Compression.ZipFile]::Open(
    $resolvedDocument,
    [System.IO.Compression.ZipArchiveMode]::Update
)
try {
    $entry = $zip.GetEntry('word/document.xml')
    if ($null -eq $entry) { throw 'El DOCX no contiene word/document.xml.' }
    [xml]$xml = Read-EntryText $entry
    $w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    $ns = [System.Xml.XmlNamespaceManager]::new($xml.NameTable)
    $ns.AddNamespace('w', $w)

    $heading = $null
    foreach ($paragraph in @($xml.SelectNodes('//w:body//w:p', $ns))) {
        $text = (($paragraph.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }) -join '')
        if ($text -eq 'Índice de códigos') {
            $heading = $paragraph
            break
        }
    }
    if ($null -eq $heading) { throw 'No se encontró Índice de códigos.' }

    $previous = $heading.PreviousSibling
    if ($null -ne $previous -and
        $previous.LocalName -eq 'p' -and
        $null -ne $previous.SelectSingleNode('.//w:br[@w:type="page"]', $ns)) {
        throw 'El salto antes del índice de códigos ya existe.'
    }

    $breakParagraph = $xml.CreateElement('w', 'p', $w)
    $run = $xml.CreateElement('w', 'r', $w)
    $break = $xml.CreateElement('w', 'br', $w)
    $type = $xml.CreateAttribute('w', 'type', $w)
    $type.Value = 'page'
    [void]$break.Attributes.Append($type)
    [void]$run.AppendChild($break)
    [void]$breakParagraph.AppendChild($run)
    [void]$heading.ParentNode.InsertBefore($breakParagraph, $heading)

    Write-XmlEntry -Zip $zip -Name 'word/document.xml' -Xml $xml
}
finally {
    $zip.Dispose()
}

$report = [pscustomobject]@{
    CodeIndexPageBreakInserted=$true
    Purpose='Aprovechar el reverso anterior sin aumentar el número total de páginas'
    SHA256=(Get-FileHash -LiteralPath $resolvedDocument -Algorithm SHA256).Hash
}
$json = $report | ConvertTo-Json -Depth 4
[System.IO.File]::WriteAllText(
    [System.IO.Path]::GetFullPath($LogPath),
    $json,
    [System.Text.UTF8Encoding]::new($false)
)
$json
