param(
    [Parameter(Mandatory = $true)]
    [string]$DocxPath,
    [Parameter(Mandatory = $true)]
    [string]$OutPath
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$resolved = (Resolve-Path -LiteralPath $DocxPath).Path
$directory = Split-Path -Parent $resolved
$backup = Join-Path $directory 'qa\TFM_ENTREGA_FINAL_PRE_PLACEHOLDER_ROW_FIX.docx'
$temporary = Join-Path $directory 'TFM_ENTREGA_FINAL_WORK.placeholder-fix.tmp.docx'

Copy-Item -LiteralPath $resolved -Destination $backup -Force
Copy-Item -LiteralPath $resolved -Destination $temporary -Force

$removed = 0
$archive = [System.IO.Compression.ZipFile]::Open($temporary, [System.IO.Compression.ZipArchiveMode]::Update)
try {
    $entry = $archive.GetEntry('word/document.xml')
    if (-not $entry) {
        throw 'No se ha encontrado word/document.xml.'
    }

    $reader = [System.IO.StreamReader]::new($entry.Open())
    try {
        [xml]$xml = $reader.ReadToEnd()
    }
    finally {
        $reader.Dispose()
    }

    $namespace = [System.Xml.XmlNamespaceManager]::new($xml.NameTable)
    $namespace.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')

    foreach ($table in @($xml.SelectNodes('//w:body/w:tbl', $namespace))) {
        foreach ($row in @($table.SelectNodes('./w:tr', $namespace))) {
            $cellTexts = @(
                foreach ($cell in $row.SelectNodes('./w:tc', $namespace)) {
                    (($cell.SelectNodes('.//w:t', $namespace) | ForEach-Object { $_.InnerText }) -join '').Trim()
                }
            )

            if (
                $cellTexts.Count -eq 4 -and
                @($cellTexts | Where-Object { $_ -eq 'T1105' }).Count -eq 4
            ) {
                [void]$table.RemoveChild($row)
                $removed++
            }
        }
    }

    if ($removed -ne 1) {
        throw "Se esperaba eliminar una fila placeholder y se localizaron $removed."
    }

    $entry.Delete()
    $replacement = $archive.CreateEntry('word/document.xml', [System.IO.Compression.CompressionLevel]::Optimal)
    $settings = [System.Xml.XmlWriterSettings]::new()
    $settings.Encoding = [System.Text.UTF8Encoding]::new($false)
    $settings.Indent = $false
    $settings.OmitXmlDeclaration = $false
    $writer = [System.Xml.XmlWriter]::Create($replacement.Open(), $settings)
    try {
        $xml.Save($writer)
    }
    finally {
        $writer.Dispose()
    }
}
finally {
    $archive.Dispose()
}

Copy-Item -LiteralPath $temporary -Destination $resolved -Force
Remove-Item -LiteralPath $temporary -Force

$result = [ordered]@{
    DocumentPath = $resolved
    PlaceholderRowsRemoved = $removed
    RemovedContent = 'Fila íntegramente compuesta por T1105 en sus cuatro celdas'
    Rationale = 'Residuo incompleto sin contenido académico verificable; no se inventan los tres campos ausentes'
    SHA256 = (Get-FileHash -LiteralPath $resolved -Algorithm SHA256).Hash
}

$json = $result | ConvertTo-Json -Depth 4
Set-Content -LiteralPath $OutPath -Value $json -Encoding UTF8
$json
