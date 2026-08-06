param(
    [Parameter(Mandatory = $true)]
    [string[]]$WorkbookPaths,
    [Parameter(Mandatory = $true)]
    [string]$OutPath
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$results = @()

foreach ($workbookPath in $WorkbookPaths) {
    $resolved = (Resolve-Path -LiteralPath $workbookPath).Path
    $directory = Split-Path -Parent $resolved
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($resolved)
    $backup = Join-Path $directory "qa\$stem.PRE_SANITIZE.xlsx"
    $temporary = Join-Path $directory "$stem.sanitize.tmp.xlsx"

    Copy-Item -LiteralPath $resolved -Destination $backup -Force
    Copy-Item -LiteralPath $resolved -Destination $temporary -Force

    $counts = [ordered]@{
        ProjectRootJulio = 0
        ProjectRootSeguridad = 0
        RemainingWindowsUserRoots = 0
        LinuxAdminRoots = 0
        DiscordWebhookUrls = 0
        PrivateIPv4 = 0
    }
    $changedEntries = 0

    $archive = [System.IO.Compression.ZipFile]::Open(
        $temporary,
        [System.IO.Compression.ZipArchiveMode]::Update
    )
    try {
        $targets = @(
            $archive.Entries |
                Where-Object { $_.FullName -match '\.(xml|rels)$' } |
                ForEach-Object { $_.FullName }
        )

        foreach ($entryName in $targets) {
            $entry = $archive.GetEntry($entryName)
            $reader = [System.IO.StreamReader]::new($entry.Open())
            try {
                $text = $reader.ReadToEnd()
            }
            finally {
                $reader.Dispose()
            }

            $original = $text

            $pattern = 'C:\\Users\\julio\\Desktop\\TFM'
            $matches = [regex]::Matches($text, $pattern, 'IgnoreCase').Count
            $counts.ProjectRootJulio += $matches
            if ($matches) {
                $text = [regex]::Replace($text, $pattern, '&lt;RAÍZ_TFM&gt;', 'IgnoreCase')
            }

            $pattern = 'C:\\Users\\seguridad\\Desktop\\TFM'
            $matches = [regex]::Matches($text, $pattern, 'IgnoreCase').Count
            $counts.ProjectRootSeguridad += $matches
            if ($matches) {
                $text = [regex]::Replace($text, $pattern, '&lt;RAÍZ_TFM&gt;', 'IgnoreCase')
            }

            $pattern = 'C:\\Users\\(?:julio|seguridad)'
            $matches = [regex]::Matches($text, $pattern, 'IgnoreCase').Count
            $counts.RemainingWindowsUserRoots += $matches
            if ($matches) {
                $text = [regex]::Replace($text, $pattern, '&lt;HOST_LAB&gt;', 'IgnoreCase')
            }

            $pattern = '/home/admin'
            $matches = [regex]::Matches($text, $pattern, 'IgnoreCase').Count
            $counts.LinuxAdminRoots += $matches
            if ($matches) {
                $text = [regex]::Replace($text, $pattern, '&lt;HOST_LAB&gt;', 'IgnoreCase')
            }

            $pattern = 'https://discord\.com/api/webhooks/[A-Za-z0-9_./-]+'
            $matches = [regex]::Matches($text, $pattern, 'IgnoreCase').Count
            $counts.DiscordWebhookUrls += $matches
            if ($matches) {
                $text = [regex]::Replace(
                    $text,
                    $pattern,
                    '&lt;SALIDA_EXTERNA_REDACTADA&gt;',
                    'IgnoreCase'
                )
            }

            $pattern = '(?<![\d.])(?:10(?:\.\d{1,3}){3}|192\.168(?:\.\d{1,3}){2}|172\.(?:1[6-9]|2\d|3[01])(?:\.\d{1,3}){2})(?![\d.])'
            $matches = [regex]::Matches($text, $pattern).Count
            $counts.PrivateIPv4 += $matches
            if ($matches) {
                $text = [regex]::Replace($text, $pattern, '&lt;HOST_LAB&gt;')
            }

            if ($text -ne $original) {
                $entry.Delete()
                $replacement = $archive.CreateEntry(
                    $entryName,
                    [System.IO.Compression.CompressionLevel]::Optimal
                )
                $writer = [System.IO.StreamWriter]::new(
                    $replacement.Open(),
                    [System.Text.UTF8Encoding]::new($false)
                )
                try {
                    $writer.Write($text)
                }
                finally {
                    $writer.Dispose()
                }
                $changedEntries++
            }
        }
    }
    finally {
        $archive.Dispose()
    }

    Copy-Item -LiteralPath $temporary -Destination $resolved -Force
    Remove-Item -LiteralPath $temporary -Force

    $results += [ordered]@{
        WorkbookPath = $resolved
        ChangedXmlEntries = $changedEntries
        Replacements = $counts
        TotalReplacements = (
            $counts.ProjectRootJulio +
            $counts.ProjectRootSeguridad +
            $counts.RemainingWindowsUserRoots +
            $counts.LinuxAdminRoots +
            $counts.DiscordWebhookUrls +
            $counts.PrivateIPv4
        )
        SHA256 = (Get-FileHash -LiteralPath $resolved -Algorithm SHA256).Hash
    }
}

$json = $results | ConvertTo-Json -Depth 6
Set-Content -LiteralPath $OutPath -Value $json -Encoding UTF8
$json
