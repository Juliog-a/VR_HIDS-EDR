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
$toc = $null
$pagesBefore = 0
$pagesAfter = 0
$reopenedPages = 0

function Release-ComObject {
    param([object]$Object)
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Write-Progress {
    param([string]$Message)
    [System.IO.File]::WriteAllText(
        [System.IO.Path]::GetFullPath($ProgressPath),
        "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`t$Message",
        [System.Text.UTF8Encoding]::new($false)
    )
}

function Find-ParagraphRange {
    param(
        [object]$Document,
        [string]$Text
    )
    $range = $Document.Content.Duplicate
    $find = $range.Find
    try {
        $find.ClearFormatting()
        $find.Text = $Text
        $find.Forward = $true
        $find.Wrap = 0
        $found = $find.Execute()
        if (-not $found) { throw "No se encontró el párrafo: $Text" }
        $paragraph = $range.Paragraphs.Item(1)
        try { return $paragraph.Range.Duplicate }
        finally { Release-ComObject $paragraph }
    }
    finally {
        Release-ComObject $find
        Release-ComObject $range
    }
}

$resolvedDocument = (Resolve-Path -LiteralPath $DocumentPath).Path
$resolvedPdf = [System.IO.Path]::GetFullPath($PdfPath)

try {
    Write-Progress '1/8 Apertura editable'
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.ScreenUpdating = $false
    $word.AutomationSecurity = 3
    $doc = $word.Documents.Open($resolvedDocument, $false, $false, $false)
    $pagesBefore = [int]$doc.ComputeStatistics(2)

    Write-Progress '2/8 Delimitación del índice general'
    $indexHeading = Find-ParagraphRange -Document $doc -Text 'Índice'
    $tablesHeading = Find-ParagraphRange -Document $doc -Text 'Índice de tablas'
    try {
        $deleteRange = $doc.Range([int]$indexHeading.End, [int]$tablesHeading.Start)
        try { $deleteRange.Delete() } finally { Release-ComObject $deleteRange }
    }
    finally {
        Release-ComObject $tablesHeading
    }

    Write-Progress '3/8 Inserción del TOC automático'
    $insertRange = $doc.Range([int]$indexHeading.End, [int]$indexHeading.End)
    try {
        $toc = $doc.TablesOfContents.Add(
            $insertRange,
            $true,
            1,
            3,
            $false,
            '',
            $true,
            $true,
            '',
            $true,
            $true,
            $true
        )
    }
    finally {
        Release-ComObject $insertRange
        Release-ComObject $indexHeading
    }

    Write-Progress '4/8 Salto de página antes del índice de tablas'
    $tablesHeading = Find-ParagraphRange -Document $doc -Text 'Índice de tablas'
    try {
        $breakRange = $doc.Range([int]$tablesHeading.Start, [int]$tablesHeading.Start)
        try { $breakRange.InsertBreak(7) } finally { Release-ComObject $breakRange }
    }
    finally { Release-ComObject $tablesHeading }

    Write-Progress '5/8 Actualización de índices y campos'
    $toc.Update()
    for ($i = 1; $i -le $doc.TablesOfFigures.Count; $i++) {
        $tof = $doc.TablesOfFigures.Item($i)
        try { $tof.Update() } finally { Release-ComObject $tof }
    }
    [void]$doc.Fields.Update()
    $doc.Repaginate()
    $toc.UpdatePageNumbers()
    $doc.Repaginate()
    $pagesAfter = [int]$doc.ComputeStatistics(2)

    Write-Progress '6/8 Guardado'
    $doc.Save()
    $doc.Close($false)
    Release-ComObject $doc
    $doc = $null
    Release-ComObject $toc
    $toc = $null

    Write-Progress '7/8 Reapertura de estabilidad'
    $doc = $word.Documents.Open($resolvedDocument, $false, $true, $false)
    $doc.Repaginate()
    $reopenedPages = [int]$doc.ComputeStatistics(2)
    if ($doc.TablesOfContents.Count -ne 1) {
        throw "Se esperaba un índice general automático y se obtuvieron $($doc.TablesOfContents.Count)."
    }
    $text = [string]$doc.Content.Text
    if ($text -match '(?i)(Error|¡Error!).*(marcador|bookmark|referencia)') {
        throw 'Persisten errores de marcador o referencia tras reconstruir el índice.'
    }

    Write-Progress '8/8 Exportación PDF'
    $doc.ExportAsFixedFormat(
        $resolvedPdf,
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
}
finally {
    if ($null -ne $toc) { Release-ComObject $toc }
    if ($null -ne $doc) {
        try { $doc.Close($false) } catch {}
        Release-ComObject $doc
    }
    if ($null -ne $word) {
        try { $word.Quit() } catch {}
        Release-ComObject $word
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$report = [pscustomobject]@{
    DocumentPath=$resolvedDocument
    PdfPath=$resolvedPdf
    PagesBefore=$pagesBefore
    PagesAfter=$pagesAfter
    ReopenedPages=$reopenedPages
    TablesOfContents=1
    TablesOfFigures=3
    BrokenReferenceErrors=0
    DocxSHA256=(Get-FileHash -LiteralPath $resolvedDocument -Algorithm SHA256).Hash
    PdfSHA256=(Get-FileHash -LiteralPath $resolvedPdf -Algorithm SHA256).Hash
    PdfBytes=(Get-Item -LiteralPath $resolvedPdf).Length
}
$json = $report | ConvertTo-Json -Depth 5
[System.IO.File]::WriteAllText(
    [System.IO.Path]::GetFullPath($LogPath),
    $json,
    [System.Text.UTF8Encoding]::new($false)
)
Write-Progress 'COMPLETADO'
$json
