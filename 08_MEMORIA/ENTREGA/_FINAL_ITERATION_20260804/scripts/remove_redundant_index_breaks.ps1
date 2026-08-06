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
$removed = 0
try {
    $entry = $zip.GetEntry('word/document.xml')
    if ($null -eq $entry) { throw 'El DOCX no contiene word/document.xml.' }
    [xml]$xml = Read-EntryText $entry
    $w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    $ns = [System.Xml.XmlNamespaceManager]::new($xml.NameTable)
    $ns.AddNamespace('w', $w)

    foreach ($headingText in @('Índice de figuras', 'Índice de códigos')) {
        $heading = $null
        foreach ($paragraph in @($xml.SelectNodes('//w:body//w:p', $ns))) {
            $text = (($paragraph.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }) -join '')
            if ($text -eq $headingText) {
                $heading = $paragraph
                break
            }
        }
        if ($null -eq $heading) { throw "No se encontró $headingText." }
        $previous = $heading.PreviousSibling
        if ($null -ne $previous -and
            $previous.LocalName -eq 'p' -and
            $null -ne $previous.SelectSingleNode('.//w:br[@w:type="page"]', $ns)) {
            [void]$previous.ParentNode.RemoveChild($previous)
            $removed++
        }
    }

    if ($removed -ne 2) {
        throw "Se esperaban dos saltos redundantes y se retiraron $removed."
    }
    Write-XmlEntry -Zip $zip -Name 'word/document.xml' -Xml $xml
}
finally {
    $zip.Dispose()
}

$report = [pscustomobject]@{
    RedundantPageBreakParagraphsRemoved=$removed
    SHA256=(Get-FileHash -LiteralPath $resolvedDocument -Algorithm SHA256).Hash
}
$json = $report | ConvertTo-Json -Depth 4
[System.IO.File]::WriteAllText(
    [System.IO.Path]::GetFullPath($LogPath),
    $json,
    [System.Text.UTF8Encoding]::new($false)
)
$json
