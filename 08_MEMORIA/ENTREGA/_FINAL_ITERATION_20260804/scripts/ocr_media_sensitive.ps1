param(
    [Parameter(Mandatory = $true)]
    [string]$DocxPath,
    [Parameter(Mandatory = $true)]
    [string]$ExtractDirectory,
    [Parameter(Mandatory = $true)]
    [string]$OutJson,
    [int]$MaxImages = 0
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Runtime.WindowsRuntime

$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Storage.FileAccessMode, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]

$asTaskGeneric = [System.WindowsRuntimeSystemExtensions].GetMethods() |
    Where-Object {
        $_.Name -eq 'AsTask' -and
        $_.IsGenericMethodDefinition -and
        $_.GetGenericArguments().Count -eq 1 -and
        $_.GetParameters().Count -eq 1 -and
        $_.ToString() -match 'IAsyncOperation`1'
    } |
    Select-Object -First 1

function Await-WinRt {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Operation,
        [Parameter(Mandatory = $true)]
        [type]$ResultType
    )
    try {
        $task = $asTaskGeneric.MakeGenericMethod($ResultType).Invoke($null, @($Operation))
        try {
            [void]$task.Wait()
            return $task.Result
        }
        finally {
            if ($task -is [System.IDisposable]) { $task.Dispose() }
        }
    }
    catch {
        while ([int]$Operation.Status -eq 0) {
            Start-Sleep -Milliseconds 10
        }
        if ([int]$Operation.Status -eq 1) {
            return $Operation.GetResults()
        }
        throw
    }
}

function Get-LineRectangle {
    param([object]$Line)
    $left = [double]::PositiveInfinity
    $top = [double]::PositiveInfinity
    $right = [double]::NegativeInfinity
    $bottom = [double]::NegativeInfinity
    foreach ($word in $Line.Words) {
        $rect = $word.BoundingRect
        $left = [Math]::Min($left, [double]$rect.X)
        $top = [Math]::Min($top, [double]$rect.Y)
        $right = [Math]::Max($right, [double]($rect.X + $rect.Width))
        $bottom = [Math]::Max($bottom, [double]($rect.Y + $rect.Height))
    }
    if ([double]::IsInfinity($left)) { return $null }
    return [pscustomobject]@{
        X = [Math]::Floor($left)
        Y = [Math]::Floor($top)
        Width = [Math]::Ceiling($right - $left)
        Height = [Math]::Ceiling($bottom - $top)
    }
}

$resolvedDocx = (Resolve-Path -LiteralPath $DocxPath).Path
$resolvedExtract = [System.IO.Path]::GetFullPath($ExtractDirectory)
[System.IO.Directory]::CreateDirectory($resolvedExtract) | Out-Null

$zip = [System.IO.Compression.ZipFile]::OpenRead($resolvedDocx)
try {
    foreach ($entry in $zip.Entries) {
        if ($entry.FullName -notmatch '^word/media/[^/]+\.(png|jpe?g|bmp|tiff?)$') { continue }
        $destination = Join-Path $resolvedExtract ([System.IO.Path]::GetFileName($entry.FullName))
        $sourceStream = $entry.Open()
        try {
            $targetStream = [System.IO.File]::Open(
                $destination,
                [System.IO.FileMode]::Create,
                [System.IO.FileAccess]::Write,
                [System.IO.FileShare]::None
            )
            try { $sourceStream.CopyTo($targetStream) }
            finally { $targetStream.Dispose() }
        }
        finally { $sourceStream.Dispose() }
    }
}
finally {
    $zip.Dispose()
}

$ocr = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
if ($null -eq $ocr) { throw 'No se pudo inicializar Windows OCR.' }

$sensitivePattern = '(?ix)' +
    '(?:[A-Z]\s*:\s*\\\s*Users\s*\\\s*(?:julio|seguridad))|' +
    '(?:\bUsers\b.{0,40}\b(?:julio|seguridad)\b)|' +
    '(?:\b(?:Users|Llsers)\b.{0,60}\bseguridad\b)|' +
    '(?:\b(?:julio|seguridad)\b.{0,80}\b(?:Desktop|TFM|Pruebas|Runners)\b)|' +
    '(?:\bLT2?9\s*\\+\s*seguridad\b)|' +
    '(?:\bHost\b.{0,60}\bUser\b.{0,40}\bseguridad\b)|' +
    '(?:\bsegur\s*\\\s*TEC)|' +
    '(?:/\s*home\s*/\s*admin)|' +
    '(?:\bhome\b.{0,40}\badmin\b)|' +
    '(?:\badmin\b.{0,80}\bTFM[_/\\])|' +
    '(?:\b(?:1)?192\s*[\._]\s*168\s*[\._]\s*\d{1,3}\s*[\._]\s*\d{1,3}\b)|' +
    '(?:\b10\s*\.\s*\d{1,3}\s*\.\s*\d{1,3}\s*\.\s*\d{1,3}\b)|' +
    '(?:\bcandidate\b)|' +
    '(?:\bDUMB[_ -]?LAB\b)|' +
    '(?:\bdump[_ -]?exfil\b)|' +
    '(?:\bdatos[_ -]?robados\b)'

$results = [System.Collections.Generic.List[object]]::new()
$allLines = [System.Collections.Generic.List[object]]::new()
$filesToScan = @(Get-ChildItem -LiteralPath $resolvedExtract -File | Sort-Object Name)
if ($MaxImages -gt 0) { $filesToScan = @($filesToScan | Select-Object -First $MaxImages) }
foreach ($file in $filesToScan) {
    Write-Output ("OCR " + $file.Name)
    $fileStream = [System.IO.File]::Open(
        $file.FullName,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::Read,
        [System.IO.FileShare]::Read
    )
    $stream = [System.IO.WindowsRuntimeStreamExtensions]::AsRandomAccessStream($fileStream)
    try {
        Write-Output '  stream'
        $decoder = Await-WinRt ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
        Write-Output '  decoder'
        $bitmap = Await-WinRt ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
        try {
            Write-Output '  bitmap'
            $ocrResult = Await-WinRt ($ocr.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
            Write-Output '  recognized'
            foreach ($line in $ocrResult.Lines) {
                $text = [string]$line.Text
                $rect = Get-LineRectangle -Line $line
                if ($null -eq $rect) { continue }
                $lineRecord = [pscustomobject]@{
                    Image = $file.Name
                    Text = $text
                    X = $rect.X
                    Y = $rect.Y
                    Width = $rect.Width
                    Height = $rect.Height
                }
                [void]$allLines.Add($lineRecord)
                if ($text -match $sensitivePattern) {
                    [void]$results.Add($lineRecord)
                }
            }
        }
        finally {
            if ($bitmap -is [System.IDisposable]) { $bitmap.Dispose() }
        }
    }
    finally {
        if ($stream -is [System.IDisposable]) { $stream.Dispose() }
        $fileStream.Dispose()
    }
}

$report = [pscustomobject]@{
    DocxPath = $resolvedDocx
    ExtractDirectory = $resolvedExtract
    ImagesScanned = $filesToScan.Count
    SensitiveLines = $results.Count
    Findings = @($results)
    AllLines = @($allLines)
}

$json = $report | ConvertTo-Json -Depth 6
[System.IO.File]::WriteAllText(
    [System.IO.Path]::GetFullPath($OutJson),
    $json,
    [System.Text.UTF8Encoding]::new($false)
)
$json
