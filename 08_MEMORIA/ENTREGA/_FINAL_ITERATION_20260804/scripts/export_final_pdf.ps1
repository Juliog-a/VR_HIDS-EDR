param(
    [Parameter(Mandatory = $true)]
    [string]$DocumentPath,
    [Parameter(Mandatory = $true)]
    [string]$PdfPath,
    [Parameter(Mandatory = $true)]
    [string]$LogPath,
    [Parameter(Mandatory = $true)]
    [string]$ProgressPath
)

$ErrorActionPreference = 'Stop'
$word = $null
$doc = $null

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

try {
    Set-Progress 'Apertura y actualización final'
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3

    $doc = $word.Documents.Open($DocumentPath, $false, $false, $false)
    $doc.TrackRevisions = $false
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
    $doc.Repaginate()
    $doc.Save()
    $pagesBeforeExport = [int]$doc.ComputeStatistics(2)

    Set-Progress 'Exportación PDF'
    $doc.ExportAsFixedFormat(
        $PdfPath,
        17,
        $false,
        0,
        0,
        1,
        9999,
        0,
        $true,
        $true,
        1,
        $true,
        $true,
        $false
    )
    $doc.Close($false)
    Release-ComObject $doc
    $doc = $null

    Set-Progress 'Reapertura de estabilidad'
    $doc = $word.Documents.Open($DocumentPath, $false, $true, $false)
    $doc.Repaginate()
    $pagesAfterReopen = [int]$doc.ComputeStatistics(2)
    $text = [string]$doc.Content.Text
    $genericSources = [regex]::Matches($text, '(?im)^\s*Fuente:\s*(elaboración|elaboración o captura)').Count
    $brokenReferences = [regex]::Matches($text, '(?i)Error\.\s*No se encuentra el origen de la referencia').Count
    $doc.Close($false)
    Release-ComObject $doc
    $doc = $null
    $word.Quit()
    Release-ComObject $word
    $word = $null

    if (-not (Test-Path -LiteralPath $PdfPath)) {
        throw 'Word no generó el PDF.'
    }
    $pdfInfo = Get-Item -LiteralPath $PdfPath
    if ($pdfInfo.Length -le 0) {
        throw 'El PDF generado está vacío.'
    }
    if ($pagesBeforeExport -ne $pagesAfterReopen) {
        throw "La paginación no es estable: $pagesBeforeExport frente a $pagesAfterReopen."
    }

    $result = [pscustomobject]@{
        Pages = $pagesAfterReopen
        StableReopen = $true
        GenericSources = $genericSources
        BrokenReferences = $brokenReferences
        DocxSHA256 = (Get-FileHash -LiteralPath $DocumentPath -Algorithm SHA256).Hash
        PdfSHA256 = (Get-FileHash -LiteralPath $PdfPath -Algorithm SHA256).Hash
        PdfBytes = $pdfInfo.Length
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
