param(
    [Parameter(Mandatory = $true)]
    [string]$DocumentPath,
    [Parameter(Mandatory = $true)]
    [string]$LogPath,
    [Parameter(Mandatory = $true)]
    [string]$ProgressPath
)

$ErrorActionPreference = 'Stop'
$wdFindStop = 0
$wdFindContinue = 1
$wdReplaceAll = 2
$wdStyleNormal = -1
$wdStyleCaption = -35

function Set-Progress {
    param([string]$Text)
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`t$Text" |
        Set-Content -LiteralPath $ProgressPath -Encoding UTF8
}

function Release-ComObject {
    param([object]$Object)
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Delete-SourceParagraphs {
    param([object]$Document, [string]$SourceText)
    $deleted = 0
    while ($true) {
        $range = $Document.Content.Duplicate
        $find = $range.Find
        try {
            $find.ClearFormatting()
            $find.Text = $SourceText
            $find.Forward = $true
            $find.Wrap = $wdFindStop
            $find.MatchCase = $true
            if (-not $find.Execute()) { break }
            $paragraph = $range.Paragraphs.Item(1)
            try {
                $text = ([string]$paragraph.Range.Text -replace "[`r`a]", '').Trim()
                if ($text -eq $SourceText) {
                    [void]$paragraph.Range.Delete()
                    $deleted++
                }
                else {
                    break
                }
            }
            finally { Release-ComObject $paragraph }
        }
        finally {
            Release-ComObject $find
            Release-ComObject $range
        }
    }
    return $deleted
}

function Set-NormalStyle {
    param([object]$Document, [string]$Needle)
    $range = $Document.Content.Duplicate
    $find = $range.Find
    try {
        $find.ClearFormatting()
        $find.Text = $Needle
        $find.Wrap = $wdFindStop
        if ($find.Execute()) {
            $paragraph = $range.Paragraphs.Item(1)
            try {
                $paragraph.Range.Style = $Document.Styles.Item($wdStyleNormal)
                return $true
            }
            finally { Release-ComObject $paragraph }
        }
        return $false
    }
    finally {
        Release-ComObject $find
        Release-ComObject $range
    }
}

function Delete-BlankCaptionParagraphs {
    param([object]$Document)
    $deleted = 0
    $searchStart = 0
    while ($searchStart -lt $Document.Content.End) {
        $range = $Document.Range($searchStart, $Document.Content.End)
        $find = $range.Find
        try {
            $find.ClearFormatting()
            $find.Style = $Document.Styles.Item($wdStyleCaption)
            $find.Text = '^p'
            $find.Forward = $true
            $find.Wrap = $wdFindStop
            $find.Format = $true
            if (-not $find.Execute()) { break }
            $paragraph = $range.Paragraphs.Item(1)
            try {
                $paragraphText = ([string]$paragraph.Range.Text -replace "[`r`a]", '').Trim()
                if ([string]::IsNullOrWhiteSpace($paragraphText)) {
                    $nextStart = [int]$paragraph.Range.Start
                    [void]$paragraph.Range.Delete()
                    $deleted++
                    $searchStart = $nextStart
                }
                else {
                    $searchStart = [int]$paragraph.Range.End
                }
            }
            finally { Release-ComObject $paragraph }
        }
        finally {
            Release-ComObject $find
            Release-ComObject $range
        }
    }
    return $deleted
}

$word = $null
$doc = $null
try {
    Set-Progress 'Inicio'
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    $doc = $word.Documents.Open($DocumentPath, $false, $false, $false)
    $doc.TrackRevisions = $false

    Set-Progress 'Eliminación residual de fuentes'
    $sourceLines = @(
        'Fuente: elaboración propia a partir de las evidencias y referencias citadas en el texto.',
        'Fuente: elaboración propia a partir de TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712.xlsx.',
        'Fuente: elaboración o captura propia del laboratorio.',
        'Fuente: elaboración propia a partir de TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_DEFINITIVO_20260712.xlsx.'
    )
    $sourceDeleted = 0
    foreach ($source in $sourceLines) {
        $sourceDeleted += Delete-SourceParagraphs -Document $doc -SourceText $source
    }

    Set-Progress 'Normalización de estilos residuales'
    $normalStyleChanges = 0
    foreach ($needle in @(
        'La fuente canónica del benchmark es TFM_BENCHMARK_',
        'La Figura 29 representa el volumen por perfil.',
        'La Figura 43 compara técnicas detectadas'
    )) {
        if (Set-NormalStyle -Document $doc -Needle $needle) {
            $normalStyleChanges++
        }
    }
    $blankCaptionsDeleted = Delete-BlankCaptionParagraphs -Document $doc

    Set-Progress 'Actualización final de campos e índices'
    try { [void]$doc.Fields.Update() } catch {}
    for ($i = 1; $i -le $doc.TablesOfContents.Count; $i++) {
        $toc = $doc.TablesOfContents.Item($i)
        try { $toc.Update() } finally { Release-ComObject $toc }
    }
    for ($i = 1; $i -le $doc.TablesOfFigures.Count; $i++) {
        $tof = $doc.TablesOfFigures.Item($i)
        try { $tof.Update() } finally { Release-ComObject $tof }
    }
    try { [void]$doc.Fields.Update() } catch {}

    Set-Progress 'Guardado y reapertura'
    $doc.Repaginate()
    $doc.Save()
    $doc.Close($false)
    Release-ComObject $doc
    $doc = $null

    $doc = $word.Documents.Open($DocumentPath, $false, $true, $false)
    $doc.Repaginate()
    $text = [string]$doc.Content.Text
    $pages = [int]$doc.ComputeStatistics(2)
    $remainingSources = [regex]::Matches($text, '(?im)^\s*Fuente:\s*(elaboración|elaboración o captura)').Count
    $brokenReferences = [regex]::Matches($text, '(?i)Error\.\s*No se encuentra el origen de la referencia').Count
    $doc.Close($false)
    Release-ComObject $doc
    $doc = $null
    $word.Quit()
    Release-ComObject $word
    $word = $null

    $result = [pscustomobject]@{
        ResidualSourceParagraphsDeleted = $sourceDeleted
        BlankCaptionParagraphsDeleted = $blankCaptionsDeleted
        NormalStyleChanges = $normalStyleChanges
        RemainingGenericSources = $remainingSources
        BrokenReferences = $brokenReferences
        Pages = $pages
        SHA256 = (Get-FileHash -LiteralPath $DocumentPath -Algorithm SHA256).Hash
    }
    $result | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $LogPath -Encoding UTF8
    Set-Progress 'COMPLETADO'
    $result | ConvertTo-Json -Depth 5
}
finally {
    if ($null -ne $doc) {
        try { $doc.Close($false) } catch {}
        Release-ComObject $doc
    }
    if ($null -ne $word) {
        try { $word.Quit() } catch {}
        Release-ComObject $word
    }
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
}
