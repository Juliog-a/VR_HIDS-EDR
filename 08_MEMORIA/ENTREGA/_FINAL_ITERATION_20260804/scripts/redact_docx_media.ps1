param(
    [Parameter(Mandatory = $true)]
    [string]$DocxPath,
    [Parameter(Mandatory = $true)]
    [string]$FindingsJson,
    [Parameter(Mandatory = $true)]
    [string]$OutJson
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.Drawing

function New-Redaction {
    param(
        [string]$Image,
        [double]$X,
        [double]$Y,
        [double]$Width,
        [double]$Height,
        [string]$Replacement,
        [string]$Reason
    )
    [pscustomobject]@{
        Image = $Image
        X = [int][Math]::Floor($X)
        Y = [int][Math]::Floor($Y)
        Width = [int][Math]::Ceiling($Width)
        Height = [int][Math]::Ceiling($Height)
        Replacement = $Replacement
        Reason = $Reason
    }
}

function Get-Replacement {
    param([string]$Text)
    if ($Text -match '(?i)(?:1)?192\s*[\._]\s*168|http.{0,20}192') {
        return '<HOST_LAB>'
    }
    if ($Text -match '(?i)LT2?9\s*\\|Host.{0,40}User') {
        return '<HOST_LAB>'
    }
    if ($Text -match '(?i)candidate|DUMB|Users|Llsers|segur|home|admin.{0,40}TFM') {
        return '<RUTA_EVIDENCIAS>'
    }
    return '<DATO_LAB>'
}

function Get-MedianColor {
    param(
        [System.Drawing.Bitmap]$Bitmap,
        [int]$X,
        [int]$Y,
        [int]$Width,
        [int]$Height
    )
    $rs = [System.Collections.Generic.List[int]]::new()
    $gs = [System.Collections.Generic.List[int]]::new()
    $bs = [System.Collections.Generic.List[int]]::new()
    for ($iy = 0; $iy -lt 5; $iy++) {
        for ($ix = 0; $ix -lt 9; $ix++) {
            $px = [Math]::Min(
                $Bitmap.Width - 1,
                [Math]::Max(0, $X + [int](($ix + 0.5) * [Math]::Max(1, $Width) / 9))
            )
            $py = [Math]::Min(
                $Bitmap.Height - 1,
                [Math]::Max(0, $Y + [int](($iy + 0.5) * [Math]::Max(1, $Height) / 5))
            )
            $c = $Bitmap.GetPixel($px, $py)
            [void]$rs.Add($c.R)
            [void]$gs.Add($c.G)
            [void]$bs.Add($c.B)
        }
    }
    $r = @($rs | Sort-Object)[[int]($rs.Count / 2)]
    $g = @($gs | Sort-Object)[[int]($gs.Count / 2)]
    $b = @($bs | Sort-Object)[[int]($bs.Count / 2)]
    [System.Drawing.Color]::FromArgb(255, $r, $g, $b)
}

$resolvedDocx = (Resolve-Path -LiteralPath $DocxPath).Path
$findings = Get-Content -LiteralPath $FindingsJson -Raw -Encoding UTF8 | ConvertFrom-Json
$redactions = [System.Collections.Generic.List[object]]::new()

foreach ($finding in $findings.Findings) {
    $paddingX = [Math]::Max(2, [int]($finding.Height * 0.25))
    $paddingY = [Math]::Max(1, [int]($finding.Height * 0.12))
    [void]$redactions.Add((New-Redaction `
        -Image $finding.Image `
        -X ($finding.X - $paddingX) `
        -Y ($finding.Y - $paddingY) `
        -Width ($finding.Width + 2 * $paddingX) `
        -Height ($finding.Height + 2 * $paddingY) `
        -Replacement (Get-Replacement $finding.Text) `
        -Reason $finding.Text
    ))
}

# Hallazgos manuales que Windows OCR no segmenta de forma fiable por el
# contraste o por el tamaño de la fuente. Las coordenadas se limitan a las
# cadenas privadas; no alteran el resto de la evidencia.
$manual = @(
    (New-Redaction 'image22.png' 0 0 440 19 '<RUTA_EVIDENCIAS>' 'Ruta personal en prompt'),
    (New-Redaction 'image23.png' 0 0 535 19 '<RUTA_EVIDENCIAS>' 'Ruta personal en prompt'),
    (New-Redaction 'image23.png' 0 42 120 18 '<HOST_LAB>' 'Usuario del laboratorio'),
    (New-Redaction 'image24.png' 0 0 828 45 '<RUTA_EVIDENCIAS>' 'Rutas personales en prompts'),
    (New-Redaction 'image24.png' 385 343 135 18 '<HOST_LAB>' 'Usuario autor de la tarea'),
    (New-Redaction 'image24.png' 385 444 135 18 '<HOST_LAB>' 'Usuario de ejecución'),
    (New-Redaction 'image25.png' 0 0 205 19 '<RUTA_EVIDENCIAS>' 'Ruta personal en prompt'),
    (New-Redaction 'image29.png' 0 0 711 222 '<RUTA_EVIDENCIAS>' 'Listado de rutas personales y etiqueta de laboratorio'),
    (New-Redaction 'image30.png' 0 0 715 29 '<RUTA_ARTIFACTS>' 'Ruta personal y directorio candidate en prompt'),
    (New-Redaction 'image31.png' 0 0 723 29 '<RUTA_EVIDENCIAS>' 'Ruta personal en prompt'),
    (New-Redaction 'image31.png' 0 44 723 39 '<HOST_LAB>' 'IP privada del receptor'),
    (New-Redaction 'image31.png' 0 123 723 84 '<RUTA_EVIDENCIAS>' 'Rutas de ZIP controlado'),
    (New-Redaction 'image31.png' 0 256 723 66 '<RUTA_EVIDENCIAS>' 'Ruta de resumen y prompt'),
    (New-Redaction 'image32.png' 0 24 831 119 '<RUTA_EVIDENCIAS>' 'Rutas de salida y denominación candidate'),
    (New-Redaction 'image34.png' 0 0 490 19 '<RUTA_EVIDENCIAS>' 'Ruta personal en prompt'),
    (New-Redaction 'image63.png' 0 10 450 21 '<RUTA_EVIDENCIAS>' 'Ruta personal en prompt'),
    (New-Redaction 'image65.png' 72 35 745 20 '<RUTA_EVIDENCIAS>' 'Ruta /home/admin de carpeta base'),
    (New-Redaction 'image65.png' 25 96 800 20 '<RUTA_EVIDENCIAS>' 'Ruta /home/admin de archives'),
    (New-Redaction 'image65.png' 22 151 805 20 '<RUTA_EVIDENCIAS>' 'Ruta /home/admin de alertas'),
    (New-Redaction 'image67.png' 70 18 750 20 '<RUTA_EVIDENCIAS>' 'Ruta /home/admin de carpeta custom'),
    (New-Redaction 'image67.png' 22 73 805 20 '<RUTA_EVIDENCIAS>' 'Ruta /home/admin de archives custom'),
    (New-Redaction 'image67.png' 22 126 805 20 '<RUTA_EVIDENCIAS>' 'Ruta /home/admin de alertas custom'),
    (New-Redaction 'image72.png' 450 282 470 55 '<RUTA_EVIDENCIAS>' 'Ruta personal expandida en detalle de evento'),
    (New-Redaction 'image72.png' 185 300 190 20 '<HOST_LAB>' 'Usuario del laboratorio en detalle de evento'),
    (New-Redaction 'image72.png' 322 712 165 21 '<HOST_LAB>' 'Usuario del laboratorio en campo user'),
    (New-Redaction 'image73.png' 205 319 790 46 '<RUTA_EVIDENCIAS>' 'Ruta personal en Script Block Logging'),
    (New-Redaction 'image73.png' 205 491 790 39 '<RUTA_EVIDENCIAS>' 'Ruta personal del script TEC-009'),
    (New-Redaction 'image73.png' 205 708 790 70 '<RUTA_EVIDENCIAS>' 'Ruta personal del objetivo TEC-008'),
    (New-Redaction 'image75.png' 902 317 90 23 '<RUTA_EVIDENCIAS>' 'Inicio de ruta personal en dashboard'),
    (New-Redaction 'image75.png' 205 338 790 23 '<RUTA_EVIDENCIAS>' 'Continuación de ruta personal en dashboard'),
    (New-Redaction 'image75.png' 902 421 90 23 '<RUTA_EVIDENCIAS>' 'Inicio de ruta personal en dashboard'),
    (New-Redaction 'image75.png' 205 442 790 72 '<RUTA_EVIDENCIAS>' 'Ruta y Script Block con ruta personal')
)
foreach ($item in $manual) { [void]$redactions.Add($item) }

foreach ($line in $findings.AllLines) {
    if ($line.Text -notmatch '(?i)(?:[A-Z]\s*:\s*\\\s*Users\b|C\s*:\s*\\\s*Users\b|C\s*:\s*\\\s*Llsers\b|segur\s*\\\s*TEC)') {
        continue
    }
    [void]$redactions.Add((New-Redaction `
        -Image $line.Image `
        -X ($line.X - 2) `
        -Y ($line.Y - 1) `
        -Width ($line.Width + 4) `
        -Height ($line.Height + 2) `
        -Replacement '<RUTA_EVIDENCIAS>' `
        -Reason $line.Text
    ))
}

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('tfm_media_redact_' + [Guid]::NewGuid().ToString('N'))
[System.IO.Directory]::CreateDirectory($tempRoot) | Out-Null

$archive = [System.IO.Compression.ZipFile]::Open(
    $resolvedDocx,
    [System.IO.Compression.ZipArchiveMode]::Update
)
try {
    $byImage = @($redactions | Group-Object Image)
    foreach ($group in $byImage) {
        $entry = $archive.GetEntry('word/media/' + $group.Name)
        if ($null -eq $entry) { throw "No se encontró word/media/$($group.Name)." }

        $sourcePath = Join-Path $tempRoot $group.Name
        $entryRead = $entry.Open()
        try {
            $sourceFile = [System.IO.File]::Open(
                $sourcePath,
                [System.IO.FileMode]::Create,
                [System.IO.FileAccess]::Write,
                [System.IO.FileShare]::None
            )
            try { $entryRead.CopyTo($sourceFile) }
            finally { $sourceFile.Dispose() }
        }
        finally { $entryRead.Dispose() }

        $bitmap = [System.Drawing.Bitmap]::FromFile($sourcePath)
        try {
            $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
            try {
                $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
                foreach ($redaction in $group.Group) {
                    $x = [Math]::Max(0, $redaction.X)
                    $y = [Math]::Max(0, $redaction.Y)
                    $width = [Math]::Min($bitmap.Width - $x, [Math]::Max(1, $redaction.Width))
                    $height = [Math]::Min($bitmap.Height - $y, [Math]::Max(1, $redaction.Height))
                    if ($width -le 0 -or $height -le 0) { continue }

                    $background = Get-MedianColor -Bitmap $bitmap -X $x -Y $y -Width $width -Height $height
                    $fill = [System.Drawing.SolidBrush]::new($background)
                    try { $graphics.FillRectangle($fill, $x, $y, $width, $height) }
                    finally { $fill.Dispose() }

                    $brightness = (0.299 * $background.R) + (0.587 * $background.G) + (0.114 * $background.B)
                    $foreground = if ($brightness -gt 145) {
                        [System.Drawing.Color]::FromArgb(255, 64, 64, 64)
                    }
                    else {
                        [System.Drawing.Color]::White
                    }
                    $fontSize = [Math]::Max(6.0, [Math]::Min(11.0, $height * 0.58))
                    $font = [System.Drawing.Font]::new('Consolas', [single]$fontSize, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Pixel)
                    $brush = [System.Drawing.SolidBrush]::new($foreground)
                    try {
                        $graphics.DrawString(
                            $redaction.Replacement,
                            $font,
                            $brush,
                            [single]($x + 1),
                            [single]($y + [Math]::Max(0, ($height - $fontSize) / 2 - 1))
                        )
                    }
                    finally {
                        $brush.Dispose()
                        $font.Dispose()
                    }
                }
            }
            finally { $graphics.Dispose() }

            $editedPath = Join-Path $tempRoot ('edited_' + $group.Name)
            $bitmap.Save($editedPath, [System.Drawing.Imaging.ImageFormat]::Png)
        }
        finally { $bitmap.Dispose() }

        $entryWrite = $entry.Open()
        try {
            $entryWrite.SetLength(0)
            $editedFile = [System.IO.File]::OpenRead($editedPath)
            try { $editedFile.CopyTo($entryWrite) }
            finally { $editedFile.Dispose() }
        }
        finally { $entryWrite.Dispose() }
    }
}
finally {
    $archive.Dispose()
}

$report = [pscustomobject]@{
    DocxPath = $resolvedDocx
    ImagesRedacted = @($redactions | Select-Object -ExpandProperty Image -Unique).Count
    RedactionRectangles = $redactions.Count
    Redactions = @($redactions)
    SHA256 = (Get-FileHash -LiteralPath $resolvedDocx -Algorithm SHA256).Hash
}
$json = $report | ConvertTo-Json -Depth 7
[System.IO.File]::WriteAllText(
    [System.IO.Path]::GetFullPath($OutJson),
    $json,
    [System.Text.UTF8Encoding]::new($false)
)
$json
