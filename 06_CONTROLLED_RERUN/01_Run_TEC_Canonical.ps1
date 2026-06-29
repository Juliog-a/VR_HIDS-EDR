[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][ValidatePattern('^TEC_[A-Za-z0-9_-]+$')][string]$CampaignId,
    [string]$ProjectRoot = '',
    [string]$JsonlPath = '\\VBOXSVR\Carpeta_Compartida_TFM\inbox\soc_alerts.jsonl',
    [string]$ReceiverUrl = 'http://192.168.1.129:8088/upload',
    [int]$PostRunWaitSeconds = 45,
    [switch]$IConfirmMonitoringConfigured
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$packageRoot = $PSScriptRoot
. (Join-Path $packageRoot 'lib\ControlledRerun.Common.ps1')
Assert-CRAdministrator
if (-not $IConfirmMonitoringConfigured) {
    throw 'Falta -IConfirmMonitoringConfigured. Confirme antes en GUI P1/P2/P3/P4, SOC_v3, JSONL on y Discord off.'
}

$root = Get-CRProjectRoot -RequestedRoot $ProjectRoot -PackageRoot $packageRoot
$basePath = Join-Path $root '01_ACTIVE_TESTS\Pruebas'
$runner = Join-Path $root '03_RUNNERS\TFM_Run_All_TEC_Tests_v6_PATCHED_4104.ps1'
$campaignRoot = Join-Path $packageRoot ("OUTPUT\TEC\{0}" -f $CampaignId)
if ((Test-Path -LiteralPath $campaignRoot) -and @(Get-ChildItem -LiteralPath $campaignRoot -Force).Count -gt 0) {
    throw "La campaña ya existe y no se sobrescribe: $campaignRoot"
}
New-CRDirectory -Path $campaignRoot | Out-Null
$rawDir = New-CRDirectory -Path (Join-Path $campaignRoot 'raw')
$runnerDir = New-CRDirectory -Path (Join-Path $rawDir 'runner')
$tec009Dir = New-CRDirectory -Path (Join-Path $rawDir 'tec009_source')
$byTecDir = New-CRDirectory -Path (Join-Path $campaignRoot 'by_technique')
$configDir = New-CRDirectory -Path (Join-Path $campaignRoot 'config_snapshot')
$backupDir = New-CRDirectory -Path (Join-Path $campaignRoot 'preexisting_jsonl_backup')
$exportDir = New-CRDirectory -Path (Join-Path $campaignRoot 'velociraptor_exports')

if (-not (Test-Path -LiteralPath $basePath -PathType Container)) {
    throw "No existe el BasePath activo: $basePath"
}
$requiredRoles = @('P1_CRITICAL','P2_EVENT','P3_EVENT','P4','SOC_ROUTER','TEC_RUNNER')
$hashChecks = @(Test-CRConfigHashes -ProjectRoot $root | Where-Object { $_.Role -in $requiredRoles })
$badHashes = @($hashChecks | Where-Object { -not $_.Match })
if ($badHashes.Count -gt 0) {
    $badHashes | Format-Table Role,RelativePath,ExpectedSHA256,ActualSHA256 -AutoSize
    throw 'Precondición fallida: artifacts/router/runner no coinciden con CONFIG_SOURCE_MANIFEST.csv.'
}
@($hashChecks) | Export-Csv -LiteralPath (Join-Path $campaignRoot 'preflight_hashes.csv') -NoTypeInformation -Encoding UTF8
Copy-CRConfigSnapshot -ProjectRoot $root -Destination $configDir | Out-Null
$activeScriptsDir = New-CRDirectory -Path (Join-Path $configDir 'active_test_scripts')
$activeScriptRecords = New-Object System.Collections.Generic.List[object]
foreach ($scriptFile in @(Get-ChildItem -LiteralPath $basePath -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in @('.ps1','.cmd','.bat') } | Sort-Object FullName)) {
    $relative = $scriptFile.FullName.Substring($basePath.TrimEnd('\').Length).TrimStart('\')
    $destination = Join-Path $activeScriptsDir $relative
    New-CRDirectory -Path (Split-Path -Parent $destination) | Out-Null
    Copy-Item -LiteralPath $scriptFile.FullName -Destination $destination -Force
    [void]$activeScriptRecords.Add([pscustomobject]@{
        RelativeToBasePath = $relative
        Length = $scriptFile.Length
        SHA256 = Get-CRSha256 -Path $scriptFile.FullName
    })
}
$activeScriptRecords.ToArray() | Export-Csv -LiteralPath (Join-Path $configDir 'active_test_scripts_manifest.csv') -NoTypeInformation -Encoding UTF8

try {
    $receiverUri = [Uri]$ReceiverUrl
    $receiverPort = if ($receiverUri.IsDefaultPort) { if ($receiverUri.Scheme -eq 'https') { 443 } else { 80 } } else { $receiverUri.Port }
    $reachable = Test-NetConnection -ComputerName $receiverUri.Host -Port $receiverPort -InformationLevel Quiet -WarningAction SilentlyContinue
    if (-not $reachable) {
        throw "Receiver no accesible en $($receiverUri.Host):$receiverPort. Inícielo antes de limpiar el JSONL."
    }
} catch {
    throw "Precheck receiver fallido: $($_.Exception.Message)"
}

$preexistingBackup = Backup-AndClear-CRJsonl -JsonlPath $JsonlPath -BackupDirectory $backupDir -Label $CampaignId
Start-Sleep -Seconds 2
$startLocal = Get-Date
$startUtc = $startLocal.ToUniversalTime()
Write-CRUtf8NoBom -Path (Join-Path $campaignRoot 'START_UTC.txt') -Text $startUtc.ToString('o')
Write-CRStatus -Message "Campaña TEC iniciada: $CampaignId" -Level INFO
Write-CRStatus -Message "JSONL exclusivo: $JsonlPath" -Level INFO

$consoleLog = Join-Path $runnerDir 'wrapper_console.log'
$runnerError = ''
try {
    $runnerParams = @{
        BasePath = $basePath
        ReceiverUrl = $ReceiverUrl
        ExfilTimeoutSec = 60
        KeepExfilArtifacts = $true
        EnableExfilUpload = $true
        PauseBetweenTestsSeconds = 10
        ShortPauseSeconds = 2
    }
    & $runner @runnerParams 2>&1 | Tee-Object -LiteralPath $consoleLog
} catch {
    $runnerError = $_.Exception.ToString()
    Write-CRStatus -Message "El runner lanzó una excepción: $runnerError" -Level FAIL
}

Write-CRStatus -Message "Esperando ${PostRunWaitSeconds}s para cerrar la ventana SOC." -Level INFO
Start-Sleep -Seconds $PostRunWaitSeconds
$endLocal = Get-Date
$endUtc = $endLocal.ToUniversalTime()
Write-CRUtf8NoBom -Path (Join-Path $campaignRoot 'END_UTC.txt') -Text $endUtc.ToString('o')

$rawJsonl = Join-Path $rawDir ("soc_alerts_{0}.jsonl" -f $CampaignId)
Copy-Item -LiteralPath $JsonlPath -Destination $rawJsonl -Force
$jsonlRead = Read-CRJsonl -Path $rawJsonl

$runnerSummarySource = @(Get-ChildItem -LiteralPath (Join-Path $basePath 'Logs_Pruebas_TFM') -Filter 'TFM_TEC_Run_CANDIDATE_v5_*_summary.json' -File -ErrorAction SilentlyContinue |
    Where-Object { $_.LastWriteTime -ge $startLocal.AddMinutes(-1) } | Sort-Object LastWriteTime -Descending | Select-Object -First 1)
$runnerSummary = $null
$runnerSummaryCopied = ''
$runnerSummaryCanonical = ''
if ($runnerSummarySource.Count -eq 1) {
    $runnerSummary = Get-Content -LiteralPath $runnerSummarySource[0].FullName -Raw -Encoding UTF8 | ConvertFrom-Json
    $runnerSummaryCopied = Copy-CRFile -Source $runnerSummarySource[0].FullName -DestinationDirectory $runnerDir
    $runnerSummaryCanonical = Join-Path $campaignRoot 'runner_summary.json'
    Copy-Item -LiteralPath $runnerSummarySource[0].FullName -Destination $runnerSummaryCanonical -Force
    foreach ($property in @('LogFile','SummaryTxt','SummaryCsv')) {
        $sourcePath = [string]$runnerSummary.Campaign.$property
        if (-not [string]::IsNullOrWhiteSpace($sourcePath) -and (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
            Copy-CRFile -Source $sourcePath -DestinationDirectory $runnerDir | Out-Null
        }
    }
    foreach ($sourcePath in @([string]$runnerSummary.TEC009.ZipFile,[string]$runnerSummary.TEC009.CopyFile,[string]$runnerSummary.TEC009.SummaryFile)) {
        if (-not [string]::IsNullOrWhiteSpace($sourcePath) -and (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
            Copy-CRFile -Source $sourcePath -DestinationDirectory $tec009Dir | Out-Null
        }
    }
} else {
    $runnerError = ($runnerError + ' No se localizó summary.json nuevo del runner.').Trim()
}

$indexRows = New-Object System.Collections.Generic.List[object]
$sha = [System.Security.Cryptography.SHA256]::Create()
try {
    foreach ($entry in @($jsonlRead.ValidLines)) {
        $row = $entry.Row
        $tec = [string]$row.TEC
        if ($tec -notmatch '^TEC-00[1-9]$') { continue }
        $tecDir = New-CRDirectory -Path (Join-Path $byTecDir $tec)
        $tecJsonl = Join-Path $tecDir 'alerts.jsonl'
        $lineWithNewLine = $entry.Raw + [Environment]::NewLine
        [System.IO.File]::AppendAllText($tecJsonl, $lineWithNewLine, (New-Object System.Text.UTF8Encoding($false)))
        $lineBytes = [System.Text.Encoding]::UTF8.GetBytes($entry.Raw)
        $lineHash = ([BitConverter]::ToString($sha.ComputeHash($lineBytes))).Replace('-','')
        [void]$indexRows.Add([pscustomobject]@{
            JsonlLine = $entry.Line
            LineSHA256 = $lineHash
            TEC = $tec
            Profile = [string]$row.Profile
            Artifact = [string]$row.Artifact
            Source = [string]$row.Source
            Timestamp = [string]$row.Timestamp
            DetectionTime = [string]$row.DetectionTime
            DetectionName = [string]$row.DetectionName
        })
    }
} finally {
    $sha.Dispose()
}
$indexRows.ToArray() | Export-Csv -LiteralPath (Join-Path $campaignRoot 'jsonl_line_index.csv') -NoTypeInformation -Encoding UTF8

$expectedTecs = 1..9 | ForEach-Object { 'TEC-{0:D3}' -f $_ }
foreach ($tec in $expectedTecs) {
    $tecDir = New-CRDirectory -Path (Join-Path $byTecDir $tec)
    $rows = @($jsonlRead.Rows | Where-Object { [string]$_.TEC -eq $tec })
    $runnerTechnique = if ($runnerSummary) { @($runnerSummary.Techniques | Where-Object { [string]$_.TEC_ID -eq $tec }) } else { @() }
    Write-CRJson -Object ([pscustomobject]@{
        TEC = $tec
        AlertCount = $rows.Count
        Profiles = @($rows | Group-Object Profile | ForEach-Object { [pscustomobject]@{Profile=$_.Name;Count=$_.Count} })
        RunnerTechnique = @($runnerTechnique)
    }) -Path (Join-Path $tecDir 'technique_summary.json')
    if (-not (Test-Path -LiteralPath (Join-Path $tecDir 'alerts.jsonl'))) {
        Write-CRUtf8NoBom -Path (Join-Path $tecDir 'alerts.jsonl') -Text ''
    }
}

$exportInstructions = @'
Exportar desde Velociraptor, limitado a START_UTC/END_UTC, y guardar aquí:
- P1_CRITICAL_CLIENT_EVENT.csv
- P2_EVENT_CLIENT_EVENT.csv
- P3_EVENT_CLIENT_EVENT.csv
- P4_CLIENT_EVENT.csv

No exportar el router como sustituto de CLIENT_EVENT. Conservar Artifact, Source,
timestamp, TEC/CU, Evidence y campos originales.
'@
Write-CRUtf8NoBom -Path (Join-Path $exportDir 'REQUIRED_EXPORTS.txt') -Text $exportInstructions

$campaignStatus = if ($runnerSummary) { [string]$runnerSummary.Campaign.GlobalStatus } else { 'MISSING_SUMMARY' }
$manifest = [pscustomobject]@{
    SchemaVersion = '1.0'
    Mode = 'TEC_CANONICAL'
    CampaignId = $CampaignId
    ComputerName = $env:COMPUTERNAME
    User = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    ProjectRoot = $root
    BasePath = $basePath
    JsonlPath = $JsonlPath
    ReceiverUrl = $ReceiverUrl
    StartLocal = $startLocal.ToString('o')
    StartUtc = $startUtc.ToString('o')
    EndLocal = $endLocal.ToString('o')
    EndUtc = $endUtc.ToString('o')
    PostRunWaitSeconds = $PostRunWaitSeconds
    PreexistingJsonlBackup = $preexistingBackup
    RawJsonl = $rawJsonl
    RawJsonlSHA256 = Get-CRSha256 -Path $rawJsonl
    JsonlValidRows = @($jsonlRead.Rows).Count
    JsonlBadRows = @($jsonlRead.BadLines).Count
    RunnerSummary = $runnerSummaryCopied
    RunnerSummaryCanonical = $runnerSummaryCanonical
    RunnerGlobalStatus = $campaignStatus
    RunnerError = $runnerError
    MonitoringConfirmation = $true
    ConfigHashesVerified = ($badHashes.Count -eq 0)
}
Write-CRJson -Object $manifest -Path (Join-Path $campaignRoot 'campaign_manifest.json')
Export-CRHashManifest -Root $campaignRoot -OutputPath (Join-Path $campaignRoot 'HASHES_SHA256.csv')

$coverage = @($indexRows.ToArray() | ForEach-Object { $_.TEC } | Sort-Object -Unique)
$missing = @($expectedTecs | Where-Object { $_ -notin $coverage })
if ($runnerError -or @($jsonlRead.BadLines).Count -gt 0 -or $missing.Count -gt 0 -or $campaignStatus -ne 'OK') {
    Write-CRStatus -Message ("Campaña capturada pero NO APTA. Runner={0}; JSONL_bad={1}; missing={2}. Salida: {3}" -f $campaignStatus,@($jsonlRead.BadLines).Count,($missing -join ','),$campaignRoot) -Level FAIL
    exit 2
}
Write-CRStatus -Message "Campaña TEC capturada con cobertura 9/9. Exporte ahora los CLIENT_EVENT a $exportDir" -Level OK
