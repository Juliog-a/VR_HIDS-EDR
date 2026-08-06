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

function Release-ComObject {
    param([object]$Object)
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

try {
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tApertura de solo lectura" |
        Set-Content -LiteralPath $ProgressPath -Encoding UTF8
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    $doc = $word.Documents.Open($DocumentPath, $false, $true, $false)
    $doc.Repaginate()
    $text = [string]$doc.Content.Text
    if ($text -match 'Equation Chapter 1 Section 1') {
        throw 'El residuo Equation Chapter sigue presente al abrir el DOCX.'
    }
    if ($text -match '(?im)^\s*Fuente:\s*(elaboración|elaboración o captura)') {
        throw 'Persisten líneas Fuente genéricas al abrir el DOCX.'
    }
    $pages = [int]$doc.ComputeStatistics(2)

    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tExportación PDF desde DOCX final" |
        Set-Content -LiteralPath $ProgressPath -Encoding UTF8
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
    $word.Quit()
    Release-ComObject $word
    $word = $null

    $result = [pscustomobject]@{
        Pages = $pages
        ReadOnlyExport = $true
        DocxSHA256 = (Get-FileHash -LiteralPath $DocumentPath -Algorithm SHA256).Hash
        PdfSHA256 = (Get-FileHash -LiteralPath $PdfPath -Algorithm SHA256).Hash
        PdfBytes = (Get-Item -LiteralPath $PdfPath).Length
    }
    $result | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $LogPath -Encoding UTF8
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tCOMPLETADO" |
        Set-Content -LiteralPath $ProgressPath -Encoding UTF8
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
