param(
    [Parameter(Mandatory = $true)]
    [string]$DocxPath,
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
        [int]$X,
        [int]$Y,
        [int]$Width,
        [int]$Height,
        [string]$Replacement,
        [string]$Reason
    )
    [pscustomobject]@{
        Image=$Image
        X=$X
        Y=$Y
        Width=$Width
        Height=$Height
        Replacement=$Replacement
        Reason=$Reason
    }
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

$redactions = [System.Collections.Generic.List[object]]::new()

# Captura TEC-007: rutas de scripts, dataset y evidencias.
foreach ($item in @(
    (New-Redaction 'image28.png' 145 12 570 14 '<RUTA_ARTIFACTS>' 'Ruta del script de enumeración'),
    (New-Redaction 'image28.png' 125 26 590 14 '<RUTA_ARTIFACTS>' 'Ruta del script de cifrado'),
    (New-Redaction 'image28.png' 76 40 639 14 '<RUTA_EVIDENCIAS>' 'Ruta del dataset'),
    (New-Redaction 'image28.png' 90 96 625 14 '<RUTA_EVIDENCIAS>' 'Ruta del TXT generado'),
    (New-Redaction 'image28.png' 90 110 625 14 '<RUTA_EVIDENCIAS>' 'Ruta del CSV generado'),
    (New-Redaction 'image28.png' 20 486 690 18 '<RUTA_EVIDENCIAS> > .\Scriptransom.ps1' 'Prompt con ruta personal'),
    (New-Redaction 'image28.png' 76 514 639 14 '<RUTA_EVIDENCIAS>' 'WorkDir personal'),
    (New-Redaction 'image28.png' 82 528 633 14 '<RUTA_EVIDENCIAS>' 'ListFile personal')
)) { [void]$redactions.Add($item) }
for ($i = 0; $i -lt 10; $i++) {
    [void]$redactions.Add((New-Redaction `
        'image28.png' 97 (557 + (14 * $i)) 490 13 `
        '<RUTA_EVIDENCIAS>' 'Prefijo de ruta personal en evidencia AES'))
}

# Receptor HTTP y alertas Velociraptor.
foreach ($item in @(
    (New-Redaction 'image31.png' 160 212 180 18 '<HOST_LAB>' 'IP privada del receptor'),
    (New-Redaction 'image46.png' 98 539 420 18 '<RAÍZ_TFM>\Pruebas' 'Ruta base personal'),
    (New-Redaction 'image47.png' 19 322 500 18 '<RUTA_EVIDENCIAS>' 'Ruta personal TEC-008 IOA'),
    (New-Redaction 'image47.png' 19 422 500 18 '<RUTA_EVIDENCIAS>' 'Ruta personal TEC-008 IOC'),
    (New-Redaction 'image48.png' 19 337 500 18 '<RUTA_EVIDENCIAS>' 'Ruta personal TEC-008 IOA'),
    (New-Redaction 'image48.png' 19 437 500 18 '<RUTA_EVIDENCIAS>' 'Ruta personal TEC-008 IOC'),
    (New-Redaction 'image49.png' 19 340 500 18 '<RUTA_EVIDENCIAS>' 'Ruta personal TEC-008 IOA'),
    (New-Redaction 'image49.png' 19 440 500 18 '<RUTA_EVIDENCIAS>' 'Ruta personal TEC-008 IOC'),
    (New-Redaction 'image51.png' 20 361 500 42 '<RUTA_ARTIFACTS>\exfiltracion_v3.ps1' 'Ruta personal TEC-009 IOA'),
    (New-Redaction 'image51.png' 20 461 500 42 '<RUTA_ARTIFACTS>\exfiltracion_v3.ps1' 'Ruta personal TEC-009 IOC')
)) { [void]$redactions.Add($item) }

# Runners, dashboard Wazuh y rutas del servidor.
foreach ($item in @(
    (New-Redaction 'image53.png' 20 70 1480 42 '<RAÍZ_TFM>\03_RUNNERS> .\TFM_Run_All_TEC_Tests_v6_PATCHED_4104.ps1' 'Ruta personal del runner TEC'),
    (New-Redaction 'image57.png' 20 70 1580 42 '<RAÍZ_TFM>\03_RUNNERS> .\TFM_Run_FP_Tests_v1.ps1' 'Ruta personal del runner FP'),
    (New-Redaction 'image61.png' 120 4 880 32 '<HOST_LAB>' 'IP privada en la barra del navegador'),
    (New-Redaction 'image64.png' 0 0 965 20 'sudo mv <RUTA_ARTIFACTS> /var/ossec/etc/rules/tfm_wazuh_custom_rules_v2.xml' 'Ruta /home/admin'),
    (New-Redaction 'image64.png' 0 20 250 18 '' 'Continuación residual de la ruta /home/admin'),
    (New-Redaction 'image70.png' 210 308 760 44 '<RUTA_EVIDENCIAS>' 'Ruta personal en la vista resumida de Script Block Logging'),
    (New-Redaction 'image70.png' 150 638 820 44 '<RUTA_EVIDENCIAS>' 'Ruta personal en Script Block Logging'),
    (New-Redaction 'image71.png' 335 563 125 20 '<HOST_LAB>' 'Usuario de laboratorio parentUser'),
    (New-Redaction 'image71.png' 335 704 125 20 '<HOST_LAB>' 'Usuario de laboratorio user'),
    (New-Redaction 'image72.png' 475 340 145 20 '<HOST_LAB>' 'Usuario de laboratorio en resumen de evento')
)) { [void]$redactions.Add($item) }

$resolvedDocx = (Resolve-Path -LiteralPath $DocxPath).Path
$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('tfm_media_remaining_' + [Guid]::NewGuid().ToString('N'))
[System.IO.Directory]::CreateDirectory($tempRoot) | Out-Null

$archive = [System.IO.Compression.ZipFile]::Open(
    $resolvedDocx,
    [System.IO.Compression.ZipArchiveMode]::Update
)
try {
    foreach ($group in @($redactions | Group-Object Image)) {
        $entry = $archive.GetEntry('word/media/' + $group.Name)
        if ($null -eq $entry) { throw "No se encontró word/media/$($group.Name)." }

        $sourcePath = Join-Path $tempRoot $group.Name
        $source = $entry.Open()
        try {
            $target = [System.IO.File]::Open(
                $sourcePath,
                [System.IO.FileMode]::Create,
                [System.IO.FileAccess]::Write,
                [System.IO.FileShare]::None
            )
            try { $source.CopyTo($target) } finally { $target.Dispose() }
        }
        finally { $source.Dispose() }

        $bitmap = [System.Drawing.Bitmap]::FromFile($sourcePath)
        try {
            $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
            try {
                $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
                foreach ($redaction in $group.Group) {
                    $x = [Math]::Max(0, $redaction.X)
                    $y = [Math]::Max(0, $redaction.Y)
                    $width = [Math]::Min($bitmap.Width - $x, $redaction.Width)
                    $height = [Math]::Min($bitmap.Height - $y, $redaction.Height)
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
                    $fontSize = [Math]::Max(6.0, [Math]::Min(18.0, $height * 0.58))
                    $font = [System.Drawing.Font]::new(
                        'Consolas',
                        [single]$fontSize,
                        [System.Drawing.FontStyle]::Regular,
                        [System.Drawing.GraphicsUnit]::Pixel
                    )
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
            try { $editedFile.CopyTo($entryWrite) } finally { $editedFile.Dispose() }
        }
        finally { $entryWrite.Dispose() }
    }
}
finally {
    $archive.Dispose()
}

$report = [pscustomobject]@{
    DocxPath=$resolvedDocx
    ImagesRedacted=@($redactions | Select-Object -ExpandProperty Image -Unique).Count
    RedactionRectangles=$redactions.Count
    Redactions=@($redactions)
    SHA256=(Get-FileHash -LiteralPath $resolvedDocx -Algorithm SHA256).Hash
}
$json = $report | ConvertTo-Json -Depth 6
[System.IO.File]::WriteAllText(
    [System.IO.Path]::GetFullPath($OutJson),
    $json,
    [System.Text.UTF8Encoding]::new($false)
)
$json
