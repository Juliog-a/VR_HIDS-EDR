param(
    [Parameter(Mandatory = $true)]
    [string]$DocumentPath,
    [Parameter(Mandatory = $true)]
    [string]$PdfPath,
    [Parameter(Mandatory = $true)]
    [string]$LogPath
)

$ErrorActionPreference = 'Stop'
$word = $null
$doc = $null

function Release-ComObject {
    param([object]$Object)
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Find-Paragraph {
    param([object]$Document, [string]$Text)
    $range = $Document.Content.Duplicate
    $find = $range.Find
    try {
        $find.ClearFormatting()
        $find.Text = $Text
        $find.Forward = $true
        $find.Wrap = 0
        if (-not $find.Execute()) { throw "No se encontró: $Text" }
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
$pagesBefore = 0
$pagesAfter = 0
$reopenedPages = 0

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.ScreenUpdating = $false
    $word.AutomationSecurity = 3
    $doc = $word.Documents.Open($resolvedDocument, $false, $false, $false)
    $pagesBefore = [int]$doc.ComputeStatistics(2)

    $tablesHeading = Find-Paragraph -Document $doc -Text 'Índice de tablas'
    try { $tablesHeading.ParagraphFormat.PageBreakBefore = -1 }
    finally { Release-ComObject $tablesHeading }

    $figuresHeading = Find-Paragraph -Document $doc -Text 'Índice de figuras'
    try { $figuresHeading.ParagraphFormat.PageBreakBefore = 0 }
    finally { Release-ComObject $figuresHeading }

    $codeHeading = Find-Paragraph -Document $doc -Text 'Índice de códigos'
    try {
        $codeHeading.ParagraphFormat.PageBreakBefore = 0
    }
    finally { Release-ComObject $codeHeading }

    if ($doc.TablesOfContents.Count -ne 1) {
        throw 'El índice general automático no está disponible.'
    }
    $toc = $doc.TablesOfContents.Item(1)
    try { $toc.UpdatePageNumbers() } finally { Release-ComObject $toc }
    for ($i = 1; $i -le $doc.TablesOfFigures.Count; $i++) {
        $tof = $doc.TablesOfFigures.Item($i)
        try { $tof.UpdatePageNumbers() } finally { Release-ComObject $tof }
    }
    $doc.Repaginate()
    $pagesAfter = [int]$doc.ComputeStatistics(2)
    $doc.Save()
    $doc.Close($false)
    Release-ComObject $doc
    $doc = $null

    $doc = $word.Documents.Open($resolvedDocument, $false, $true, $false)
    $doc.Repaginate()
    $reopenedPages = [int]$doc.ComputeStatistics(2)
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
    PagesBefore=$pagesBefore
    PagesAfter=$pagesAfter
    ReopenedPages=$reopenedPages
    SecondaryIndicesStartOnNewPage=$true
    CodeIndexHeading='Índice de códigos'
    DocxSHA256=(Get-FileHash -LiteralPath $resolvedDocument -Algorithm SHA256).Hash
    PdfSHA256=(Get-FileHash -LiteralPath $resolvedPdf -Algorithm SHA256).Hash
}
$json = $report | ConvertTo-Json -Depth 4
[System.IO.File]::WriteAllText(
    [System.IO.Path]::GetFullPath($LogPath),
    $json,
    [System.Text.UTF8Encoding]::new($false)
)
$json
