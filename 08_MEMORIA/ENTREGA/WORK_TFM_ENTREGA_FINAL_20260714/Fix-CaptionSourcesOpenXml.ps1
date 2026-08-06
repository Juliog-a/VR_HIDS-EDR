[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$InputDoc,
    [Parameter(Mandatory = $true)][string]$OutputDoc,
    [Parameter(Mandatory = $true)][string]$ReportPath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$oAcute = [char]0x00F3
$inputFull = (Resolve-Path -LiteralPath $InputDoc).Path
$outputFull = [IO.Path]::GetFullPath((Join-Path (Get-Location) $OutputDoc))
$reportFull = [IO.Path]::GetFullPath((Join-Path (Get-Location) $ReportPath))

if (Test-Path -LiteralPath $outputFull) {
    $archive = Join-Path (Split-Path -Parent $outputFull) ("previous_sources_fixed_" + (Get-Date -Format 'yyyyMMdd_HHmmss_fff') + '.docx')
    Move-Item -LiteralPath $outputFull -Destination $archive
}
Copy-Item -LiteralPath $inputFull -Destination $outputFull

$zip = [IO.Compression.ZipFile]::Open($outputFull, [IO.Compression.ZipArchiveMode]::Update)
try {
    $entry = $zip.GetEntry('word/document.xml')
    if ($null -eq $entry) { throw 'word/document.xml not found.' }
    $stream = $entry.Open()
    $reader = New-Object IO.StreamReader($stream, [Text.Encoding]::UTF8, $true)
    try { $xmlText = $reader.ReadToEnd() } finally { $reader.Dispose(); $stream.Dispose() }

    $xml = New-Object Xml.XmlDocument
    $xml.PreserveWhitespace = $true
    $xml.LoadXml($xmlText)
    $ns = New-Object Xml.XmlNamespaceManager($xml.NameTable)
    $wNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    $ns.AddNamespace('w', $wNs)

    $removed = 0
    foreach ($p in @($xml.SelectNodes('//w:body//w:p', $ns))) {
        $text = (($p.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }) -join '').Trim()
        if ($text.StartsWith('Fuente:')) {
            [void]$p.ParentNode.RemoveChild($p)
            $removed++
        }
    }

    $captions = [System.Collections.Generic.List[object]]::new()
    foreach ($p in @($xml.SelectNodes('//w:body//w:p', $ns))) {
        $codes = @()
        $codes += @($p.SelectNodes('.//w:instrText', $ns) | ForEach-Object { $_.InnerText })
        $codes += @($p.SelectNodes('.//w:fldSimple', $ns) | ForEach-Object { $_.GetAttribute('instr', $wNs) })
        $codeText = ($codes -join ' ')
        if ($codeText -notmatch '\bSEQ\s+(Tabla|Figura)\b') { continue }
        $category = $matches[1]
        $captionText = (($p.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }) -join '').Trim()
        $captions.Add([pscustomobject]@{ node = $p; category = $category; text = $captionText })
    }

    $inserted = 0
    $figure38Fixed = 0
    foreach ($caption in $captions) {
        if ($caption.text -like '*WazuhSvc*' -and $caption.text -notmatch '^Figura\s+\d+:') {
            foreach ($t in @($caption.node.SelectNodes('.//w:t', $ns))) {
                if ($t.InnerText -match '^\s+.*WazuhSvc') {
                    $t.InnerText = ':' + $t.InnerText
                    $figure38Fixed++
                    break
                }
            }
        }

        if ($caption.text -match '(?i)benchmark|rendimiento|CPU|RAM|escenario') {
            $source = "Fuente: elaboraci${oAcute}n propia a partir de TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_DEFINITIVO_20260712.xlsx."
        }
        elseif ($caption.text -match '(?i)Wazuh|CLIENT_EVENT|detecci[oó]n|falsos positivos|artifacts p[uú]blicos|runner|TEC-00|Sysmon|PowerShell|alerta|cobertura') {
            $source = "Fuente: elaboraci${oAcute}n propia a partir de TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712.xlsx."
        }
        elseif ($caption.category -eq 'Tabla') {
            $source = "Fuente: elaboraci${oAcute}n propia a partir de las evidencias y referencias citadas en el texto."
        }
        else {
            $source = "Fuente: elaboraci${oAcute}n o captura propia del laboratorio."
        }

        $p = $xml.CreateElement('w', 'p', $wNs)
        $pPr = $xml.CreateElement('w', 'pPr', $wNs)
        $style = $xml.CreateElement('w', 'pStyle', $wNs)
        $style.SetAttribute('val', $wNs, 'Normal')
        [void]$pPr.AppendChild($style)
        $keep = $xml.CreateElement('w', 'keepLines', $wNs)
        [void]$pPr.AppendChild($keep)
        $spacing = $xml.CreateElement('w', 'spacing', $wNs)
        $spacing.SetAttribute('before', $wNs, '0')
        $spacing.SetAttribute('after', $wNs, '120')
        [void]$pPr.AppendChild($spacing)
        [void]$p.AppendChild($pPr)

        $r = $xml.CreateElement('w', 'r', $wNs)
        $rPr = $xml.CreateElement('w', 'rPr', $wNs)
        $fonts = $xml.CreateElement('w', 'rFonts', $wNs)
        $fonts.SetAttribute('ascii', $wNs, 'Arial')
        $fonts.SetAttribute('hAnsi', $wNs, 'Arial')
        [void]$rPr.AppendChild($fonts)
        [void]$rPr.AppendChild($xml.CreateElement('w', 'i', $wNs))
        $sz = $xml.CreateElement('w', 'sz', $wNs)
        $sz.SetAttribute('val', $wNs, '16')
        [void]$rPr.AppendChild($sz)
        $szCs = $xml.CreateElement('w', 'szCs', $wNs)
        $szCs.SetAttribute('val', $wNs, '16')
        [void]$rPr.AppendChild($szCs)
        [void]$r.AppendChild($rPr)
        $tNode = $xml.CreateElement('w', 't', $wNs)
        $tNode.InnerText = $source
        [void]$r.AppendChild($tNode)
        [void]$p.AppendChild($r)
        [void]$caption.node.ParentNode.InsertAfter($p, $caption.node)
        $inserted++
    }

    $entry.Delete()
    $newEntry = $zip.CreateEntry('word/document.xml', [IO.Compression.CompressionLevel]::Optimal)
    $outStream = $newEntry.Open()
    $settings = New-Object Xml.XmlWriterSettings
    $settings.Encoding = $utf8NoBom
    $settings.Indent = $false
    $settings.CloseOutput = $false
    $writer = [Xml.XmlWriter]::Create($outStream, $settings)
    try { $xml.Save($writer) } finally { $writer.Dispose(); $outStream.Dispose() }
}
finally { $zip.Dispose() }

$report = [ordered]@{
    generated_at = (Get-Date).ToString('o')
    input = $inputFull
    output = $outputFull
    prior_source_paragraphs_removed = $removed
    caption_fields = $captions.Count
    source_paragraphs_inserted = $inserted
    figure38_colon_fixed = $figure38Fixed
    sha256 = (Get-FileHash -Algorithm SHA256 -LiteralPath $outputFull).Hash
}
[IO.File]::WriteAllText($reportFull, ($report | ConvertTo-Json -Depth 5), $utf8NoBom)
$report | ConvertTo-Json -Depth 5
