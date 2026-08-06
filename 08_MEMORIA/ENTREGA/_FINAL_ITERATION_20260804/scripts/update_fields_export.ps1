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
$fieldsUpdated = 0
$tocUpdated = 0
$tofUpdated = 0
$storyUpdates = 0
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
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`t$Message"
    [System.IO.File]::WriteAllText(
        [System.IO.Path]::GetFullPath($ProgressPath),
        $line,
        [System.Text.UTF8Encoding]::new($false)
    )
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
    $word.Options.UpdateFieldsAtPrint = $true
    $word.Options.UpdateLinksAtPrint = $false
    $doc = $word.Documents.Open($resolvedDocument, $false, $false, $false)
    $pagesBefore = [int]$doc.ComputeStatistics(2)

    Write-Progress '2/8 Campos del cuerpo principal'
    try {
        $fieldsUpdated = [int]$doc.Fields.Count
        [void]$doc.Fields.Update()
    }
    catch {}

    Write-Progress '3/8 Indice general'
    for ($i = 1; $i -le $doc.TablesOfContents.Count; $i++) {
        $toc = $doc.TablesOfContents.Item($i)
        try {
            $toc.Update()
            $tocUpdated++
        }
        finally { Release-ComObject $toc }
    }

    Write-Progress '4/8 Indices de tablas, figuras y codigos'
    for ($i = 1; $i -le $doc.TablesOfFigures.Count; $i++) {
        $tof = $doc.TablesOfFigures.Item($i)
        try {
            $tof.Update()
            $tofUpdated++
        }
        finally { Release-ComObject $tof }
    }

    Write-Progress '5/8 Campos de cabeceras y pies'
    foreach ($section in @($doc.Sections)) {
        try {
            foreach ($header in @($section.Headers)) {
                try {
                    [void]$header.Range.Fields.Update()
                    $storyUpdates++
                }
                finally { Release-ComObject $header }
            }
            foreach ($footer in @($section.Footers)) {
                try {
                    [void]$footer.Range.Fields.Update()
                    $storyUpdates++
                }
                finally { Release-ComObject $footer }
            }
        }
        finally { Release-ComObject $section }
    }

    Write-Progress '6/8 Repaginacion y guardado'
    $doc.Repaginate()
    $pagesAfter = [int]$doc.ComputeStatistics(2)
    $doc.Save()
    $doc.Close($false)
    Release-ComObject $doc
    $doc = $null

    Write-Progress '7/8 Reapertura de estabilidad'
    $doc = $word.Documents.Open($resolvedDocument, $false, $true, $false)
    $doc.Repaginate()
    $reopenedPages = [int]$doc.ComputeStatistics(2)

    Write-Progress '8/8 Exportacion PDF'
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
    DocumentPath=$resolvedDocument
    PdfPath=$resolvedPdf
    PagesBefore=$pagesBefore
    PagesAfter=$pagesAfter
    ReopenedPages=$reopenedPages
    MainStoryFields=$fieldsUpdated
    TablesOfContentsUpdated=$tocUpdated
    TablesOfFiguresUpdated=$tofUpdated
    HeaderFooterStoriesUpdated=$storyUpdates
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
