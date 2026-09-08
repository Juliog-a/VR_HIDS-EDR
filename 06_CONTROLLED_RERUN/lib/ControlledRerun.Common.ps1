Set-StrictMode -Version Latest

function Write-CRStatus {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet('INFO','OK','WARN','FAIL')][string]$Level = 'INFO'
    )
    $color = switch ($Level) {
        'OK' { 'Green' }
        'WARN' { 'Yellow' }
        'FAIL' { 'Red' }
        default { 'Cyan' }
    }
    Write-Host ("[{0}] {1}" -f $Level, $Message) -ForegroundColor $color
}

function Assert-CRAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw 'Ejecute PowerShell como administrador.'
    }
}

function New-CRDirectory {
    param([Parameter(Mandatory=$true)][string]$Path)
    [System.IO.Directory]::CreateDirectory($Path) | Out-Null
    return $Path
}

function Write-CRUtf8NoBom {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [AllowEmptyString()][string]$Text = ''
    )
    $encoding = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Text, $encoding)
}

function Write-CRJson {
    param(
        [Parameter(Mandatory=$true)]$Object,
        [Parameter(Mandatory=$true)][string]$Path,
        [int]$Depth = 12
    )
    $json = $Object | ConvertTo-Json -Depth $Depth
    Write-CRUtf8NoBom -Path $Path -Text $json
}

function Get-CRSha256 {
    param([Parameter(Mandatory=$true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return ''
    }
    return (Get-FileHash -Algorithm SHA256 -LiteralPath $Path).Hash.ToUpperInvariant()
}

function Copy-CRFile {
    param(
        [Parameter(Mandatory=$true)][string]$Source,
        [Parameter(Mandatory=$true)][string]$DestinationDirectory,
        [string]$DestinationName = ''
    )
    if (-not (Test-Path -LiteralPath $Source -PathType Leaf)) {
        throw "No existe el fichero requerido: $Source"
    }
    New-CRDirectory -Path $DestinationDirectory | Out-Null
    if ([string]::IsNullOrWhiteSpace($DestinationName)) {
        $DestinationName = [System.IO.Path]::GetFileName($Source)
    }
    $destination = Join-Path $DestinationDirectory $DestinationName
    Copy-Item -LiteralPath $Source -Destination $destination -Force
    return $destination
}

function Backup-AndClear-CRJsonl {
    param(
        [Parameter(Mandatory=$true)][string]$JsonlPath,
        [Parameter(Mandatory=$true)][string]$BackupDirectory,
        [Parameter(Mandatory=$true)][string]$Label
    )
    $parent = Split-Path -Parent $JsonlPath
    if ([string]::IsNullOrWhiteSpace($parent) -or -not (Test-Path -LiteralPath $parent)) {
        throw "No se puede acceder al directorio del JSONL: $parent"
    }
    New-CRDirectory -Path $BackupDirectory | Out-Null
    $backup = ''
    if (Test-Path -LiteralPath $JsonlPath -PathType Leaf) {
        $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $backup = Join-Path $BackupDirectory ("soc_alerts_BEFORE_{0}_{1}.jsonl" -f $Label, $stamp)
        Copy-Item -LiteralPath $JsonlPath -Destination $backup -Force
    }
    Write-CRUtf8NoBom -Path $JsonlPath -Text ''
    return $backup
}

function Read-CRJsonl {
    param([Parameter(Mandatory=$true)][string]$Path)
    $rows = New-Object System.Collections.Generic.List[object]
    $validLines = New-Object System.Collections.Generic.List[object]
    $badLines = New-Object System.Collections.Generic.List[object]
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return [pscustomobject]@{ Rows=@(); ValidLines=@(); BadLines=@([pscustomobject]@{Line=0;Error='File not found';Raw=''}) }
    }
    $lineNo = 0
    foreach ($line in [System.IO.File]::ReadLines($Path)) {
        $lineNo++
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        try {
            $row = $line | ConvertFrom-Json -ErrorAction Stop
            [void]$rows.Add($row)
            [void]$validLines.Add([pscustomobject]@{ Line=$lineNo; Raw=$line; Row=$row })
        } catch {
            [void]$badLines.Add([pscustomobject]@{ Line=$lineNo; Error=$_.Exception.Message; Raw=$line })
        }
    }
    return [pscustomobject]@{
        Rows = $rows.ToArray()
        ValidLines = $validLines.ToArray()
        BadLines = $badLines.ToArray()
    }
}

function Export-CRHashManifest {
    param(
        [Parameter(Mandatory=$true)][string]$Root,
        [Parameter(Mandatory=$true)][string]$OutputPath
    )
    $outputFull = [System.IO.Path]::GetFullPath($OutputPath)
    $records = New-Object System.Collections.Generic.List[object]
    foreach ($file in @(Get-ChildItem -LiteralPath $Root -Recurse -File | Sort-Object FullName)) {
        if ([System.IO.Path]::GetFullPath($file.FullName) -eq $outputFull) { continue }
        $relative = $file.FullName.Substring($Root.TrimEnd('\').Length).TrimStart('\')
        [void]$records.Add([pscustomobject]@{
            RelativePath = $relative
            Length = $file.Length
            SHA256 = Get-CRSha256 -Path $file.FullName
            LastWriteTimeUtc = $file.LastWriteTimeUtc.ToString('o')
        })
    }
    $records.ToArray() | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
}

function Get-CRProjectRoot {
    param(
        [string]$RequestedRoot,
        [Parameter(Mandatory=$true)][string]$PackageRoot
    )
    if (-not [string]::IsNullOrWhiteSpace($RequestedRoot)) {
        return [System.IO.Path]::GetFullPath($RequestedRoot)
    }
    return [System.IO.Path]::GetFullPath((Split-Path -Parent $PackageRoot))
}

function Get-CRConfigDefinitions {
    return @(
        [pscustomobject]@{Role='P1_CRITICAL';RelativePath='01_ARTIFACTS\validated\Custom.TFM.HIDS.P1.Critical.Priority.Event_v1.yaml';SHA256='8A1FF2654405AF9E4412D32EF0C2BE08A383924654AA764A04AB264BAA42EDD7'},
        [pscustomobject]@{Role='P2_EVENT';RelativePath='01_ARTIFACTS\validated\Custom.TFM.HIDS.P2.High.Forensic.Event_v1.yaml';SHA256='255C5F7CA39E362FE45106C662C4C35730EFF57A4063A1E9973FD11A6375FAA2'},
        [pscustomobject]@{Role='P3_EVENT';RelativePath='01_ARTIFACTS\validated\Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1.yaml';SHA256='C591CB849753194F6F87834C1E0F34EDE6CDFF6B31BC63650A1D281ADA1D410A'},
        [pscustomobject]@{Role='P4';RelativePath='01_ARTIFACTS\validated\Custom.TFM.HIDS.P4.Low.Basic_v2.yaml';SHA256='1EBB29296A14475E38BCF34CA8DD9CC16572F4A8CA00A1AB0C3608E0B3097BC1'},
        [pscustomobject]@{Role='SOC_ROUTER';RelativePath='01_ARTIFACTS\validated\Custom.TFM.HIDS.Router.JSONL.Discord.SOC_v3.yaml';SHA256='574E3E8BC03E7C4D655400D5BB5527771E438CA67C413853A0AAAC7A10C5E33F'},
        [pscustomobject]@{Role='TEC_RUNNER';RelativePath='03_RUNNERS\validate\TFM_Run_All_TEC_Tests_v6_PATCHED_4104.ps1';SHA256='C023F8F7DC4D90977C4F6C76C8147A90C3558D62E14D9E1A265358DCDFF68367'},
        [pscustomobject]@{Role='FP_RUNNER';RelativePath='03_RUNNERS\validate\TFM_Run_FP_Tests_v1.ps1';SHA256='557B9799DAA62CFD68A31EF79296997137DBB8E355175615F20F44F8AD032528'},
        [pscustomobject]@{Role='BENCHMARK_RUNNER';RelativePath='03_RUNNERS\validate\TFM_Benchmark_VR_Resource_Usage_v1.ps1';SHA256='1C86D7493485C81D80398416FA0671B06F4A1134E89E1F9B6FE29C6E0F6EAB4F'},
        [pscustomobject]@{Role='RECEIVER';RelativePath='02_SCRIPTS\validated\receiver_tfm_v4.py';SHA256='D61A86270EA822D2D4195E8C506738D9544D7964C23CB174D9F09EE52DD84CF6'}
    )
}

function Test-CRConfigHashes {
    param(
        [Parameter(Mandatory=$true)][string]$ProjectRoot,
        [switch]$ThrowOnMismatch
    )
    $results = New-Object System.Collections.Generic.List[object]
    foreach ($definition in @(Get-CRConfigDefinitions)) {
        $path = Join-Path $ProjectRoot $definition.RelativePath
        $exists = Test-Path -LiteralPath $path -PathType Leaf
        $actual = if ($exists) { Get-CRSha256 -Path $path } else { '' }
        $match = $exists -and ($actual -eq $definition.SHA256)
        [void]$results.Add([pscustomobject]@{
            Role = $definition.Role
            RelativePath = $definition.RelativePath
            ExpectedSHA256 = $definition.SHA256
            ActualSHA256 = $actual
            Exists = $exists
            Match = $match
        })
    }
    $bad = @($results | Where-Object { -not $_.Match })
    if ($ThrowOnMismatch -and $bad.Count -gt 0) {
        $details = ($bad | ForEach-Object { "{0}: {1}" -f $_.Role, $_.RelativePath }) -join '; '
        throw "Ficheros requeridos ausentes o con hash distinto: $details"
    }
    return $results.ToArray()
}

function Copy-CRConfigSnapshot {
    param(
        [Parameter(Mandatory=$true)][string]$ProjectRoot,
        [Parameter(Mandatory=$true)][string]$Destination,
        [string[]]$Roles = @('P1_CRITICAL','P2_EVENT','P3_EVENT','P4','SOC_ROUTER','TEC_RUNNER')
    )
    New-CRDirectory -Path $Destination | Out-Null
    $records = New-Object System.Collections.Generic.List[object]
    foreach ($definition in @(Get-CRConfigDefinitions | Where-Object { $_.Role -in $Roles })) {
        $source = Join-Path $ProjectRoot $definition.RelativePath
        $name = "{0}__{1}" -f $definition.Role, [System.IO.Path]::GetFileName($source)
        $copied = Copy-CRFile -Source $source -DestinationDirectory $Destination -DestinationName $name
        [void]$records.Add([pscustomobject]@{
            Role = $definition.Role
            SourceRelativePath = $definition.RelativePath
            SnapshotFile = [System.IO.Path]::GetFileName($copied)
            SHA256 = Get-CRSha256 -Path $copied
        })
    }
    $records.ToArray() | Export-Csv -LiteralPath (Join-Path $Destination 'config_snapshot_manifest.csv') -NoTypeInformation -Encoding UTF8
    return $records.ToArray()
}
