[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$InputDoc,
    [Parameter(Mandatory = $true)][string]$ReportPath,
    [Parameter(Mandatory = $true)][string]$PdfPath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Release-ComObject {
    param([AllowNull()][object]$Object)
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object) } catch {}
    }
}

function Clean-Text {
    param([AllowNull()][string]$Text)
    if ($null -eq $Text) { return '' }
    return (($Text -replace "`r", ' ' -replace "`a", ' ' -replace "`v", ' ' -replace "`t", ' ') -replace '\s+', ' ').Trim()
}

$inputFull = (Resolve-Path -LiteralPath $InputDoc).Path
$reportFull = [IO.Path]::GetFullPath((Join-Path (Get-Location) $ReportPath))
$pdfFull = [IO.Path]::GetFullPath((Join-Path (Get-Location) $PdfPath))
$pdfDir = Split-Path -Parent $pdfFull
if (-not (Test-Path -LiteralPath $pdfDir)) { New-Item -ItemType Directory -Path $pdfDir | Out-Null }
if (Test-Path -LiteralPath $pdfFull) { Move-Item -LiteralPath $pdfFull -Destination (Join-Path $pdfDir ("previous_validation_" + (Get-Date -Format 'yyyyMMdd_HHmmss_fff') + '.pdf')) }

$baselinePids = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$word = $null
$documents = $null
$doc = $null
$report = [ordered]@{
    validated_at = (Get-Date).ToString('o')
    input = $inputFull
    sha256 = (Get-FileHash -Algorithm SHA256 -LiteralPath $inputFull).Hash
    opened_no_repair_dialog = $false
    read_only = $false
    word_version = $null
    word_build = $null
    compatibility_mode = $null
    pages = $null
    words = $null
    word_tables = $null
    inline_shapes = $null
    floating_shapes = $null
    sections = $null
    fields = $null
    seq_tables = 0
    seq_figures = 0
    seq_code = 0
    tables_of_contents = $null
    tables_of_figures = $null
    comments = $null
    revisions = $null
    key_controls = [ordered]@{}
    suspicious = @()
    pdf_exported = $false
    pdf_path = $pdfFull
    forced_process_cleanup = @()
}

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.ScreenUpdating = $false
    try { $word.AutomationSecurity = 3 } catch {}
    try { $word.Options.UpdateLinksAtOpen = $false } catch {}
    try { $word.Options.SaveNormalPrompt = $false } catch {}
    $report.word_version = [string]$word.Version
    $report.word_build = [string]$word.Build
    $documents = $word.Documents
    $doc = $documents.OpenNoRepairDialog($inputFull, $false, $true, $false)
    if ($null -eq $doc) { throw 'Word returned a null document.' }
    $report.opened_no_repair_dialog = $true
    $report.read_only = [bool]$doc.ReadOnly
    $report.compatibility_mode = [int]$doc.CompatibilityMode
    [void]$doc.Repaginate()

    $report.pages = [int]$doc.ComputeStatistics(2)
    $report.words = [int]$doc.ComputeStatistics(0)
    $report.word_tables = [int]$doc.Tables.Count
    $report.inline_shapes = [int]$doc.InlineShapes.Count
    $report.floating_shapes = [int]$doc.Shapes.Count
    $report.sections = [int]$doc.Sections.Count
    $report.fields = [int]$doc.Fields.Count
    $report.tables_of_contents = [int]$doc.TablesOfContents.Count
    $report.tables_of_figures = [int]$doc.TablesOfFigures.Count
    $report.comments = [int]$doc.Comments.Count
    $report.revisions = [int]$doc.Revisions.Count

    $fields = $null
    try {
        $fields = $doc.Fields
        for ($i = 1; $i -le $fields.Count; $i++) {
            $field = $null
            $code = $null
            try {
                $field = $fields.Item($i)
                $code = $field.Code
                $text = Clean-Text ([string]$code.Text)
                if ($text -match '^SEQ\s+Tabla\b') { $report.seq_tables++ }
                elseif ($text -match '^SEQ\s+Figura\b') { $report.seq_figures++ }
                elseif ($text -match '^SEQ\s+C') { $report.seq_code++ }
            }
            finally { Release-ComObject $code; Release-ComObject $field }
        }
    }
    finally { Release-ComObject $fields }

    $fullText = [string]$doc.Content.Text
    $controls = [ordered]@{
        custom_9_of_9 = ($fullText.IndexOf('9/9', [StringComparison]::Ordinal) -ge 0)
        client_event_379 = ($fullText.IndexOf('379', [StringComparison]::Ordinal) -ge 0)
        p1_19 = ($fullText.IndexOf('P1=19', [StringComparison]::OrdinalIgnoreCase) -ge 0)
        p2_118 = ($fullText.IndexOf('P2=118', [StringComparison]::OrdinalIgnoreCase) -ge 0)
        p3_109 = ($fullText.IndexOf('P3=109', [StringComparison]::OrdinalIgnoreCase) -ge 0)
        p4_133 = ($fullText.IndexOf('P4=133', [StringComparison]::OrdinalIgnoreCase) -ge 0)
        fp_10_of_10 = ($fullText.IndexOf('10/10', [StringComparison]::Ordinal) -ge 0)
        fp_zero = ($fullText.IndexOf('0 hits', [StringComparison]::OrdinalIgnoreCase) -ge 0)
        hayabusa_3_of_9 = ($fullText.IndexOf('Hayabusa CH', [StringComparison]::OrdinalIgnoreCase) -ge 0 -and $fullText.IndexOf('3/9', [StringComparison]::Ordinal) -ge 0)
        wazuh_base_3_of_9 = ($fullText.IndexOf('Wazuh base', [StringComparison]::OrdinalIgnoreCase) -ge 0)
        wazuh_custom_4_of_9 = ($fullText.IndexOf('Wazuh custom', [StringComparison]::OrdinalIgnoreCase) -ge 0 -and $fullText.IndexOf('4/9', [StringComparison]::Ordinal) -ge 0)
        etw_inconclusive = ($fullText.IndexOf('NO CONCLUYENTE', [StringComparison]::OrdinalIgnoreCase) -ge 0)
        tracknetwork_limit = ($fullText.IndexOf('sin atribución fiable', [StringComparison]::OrdinalIgnoreCase) -ge 0)
        http_runner = ($fullText.IndexOf('runner independiente', [StringComparison]::OrdinalIgnoreCase) -ge 0)
        definitive_visibility_excel = ($fullText.IndexOf('TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712.xlsx', [StringComparison]::Ordinal) -ge 0)
        definitive_benchmark_excel = ($fullText.IndexOf('TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_DEFINITIVO_20260712.xlsx', [StringComparison]::Ordinal) -ge 0)
    }
    $report.key_controls = $controls
    foreach ($name in $controls.Keys) { if (-not $controls[$name]) { $report.suspicious += "missing_control:$name" } }
    foreach ($bad in @('Error! Reference source not found', '¡Error! No se encuentra el origen de la referencia', 'Error! Bookmark not defined', '[[REF:', '[[SEQ_', '[[TABLE_', '[[IMAGE_', '09_Benchmark_Plan', 'v15.21')) {
        if ($fullText.IndexOf($bad, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $report.suspicious += $bad }
    }

    $doc.ExportAsFixedFormat($pdfFull, 17)
    $report.pdf_exported = (Test-Path -LiteralPath $pdfFull)
    $doc.Close(0)
    Release-ComObject $doc; $doc = $null
    $word.Quit()
    Release-ComObject $documents; $documents = $null
    Release-ComObject $word; $word = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
catch {
    $report.error = $_.Exception.ToString()
    throw
}
finally {
    if ($null -ne $doc) { try { $doc.Close(0) } catch {}; Release-ComObject $doc }
    if ($null -ne $word) { try { $word.Quit() } catch {}; Release-ComObject $word }
    Release-ComObject $documents
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 2
    $afterPids = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    foreach ($pidValue in @($afterPids | Where-Object { $_ -notin $baselinePids })) {
        try { Stop-Process -Id $pidValue -Force -ErrorAction Stop; $report.forced_process_cleanup += $pidValue } catch {}
    }
    [IO.File]::WriteAllText($reportFull, ($report | ConvertTo-Json -Depth 8), $utf8NoBom)
}

$report | ConvertTo-Json -Depth 8
