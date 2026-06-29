[CmdletBinding()]
param(
    [string]$CampaignId = '',
    [string]$ProjectRoot = '',
    [string]$JsonlPath = '\\VBOXSVR\Carpeta_Compartida_TFM\inbox\soc_alerts.jsonl',
    [ValidateRange(1,10)][int]$Repetitions = 3,
    [int]$PostRunWaitSeconds = 45,
    [int]$InterRepetitionQuietSeconds = 20,
    [switch]$IConfirmMonitoringConfigured
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$packageRoot = $PSScriptRoot
. (Join-Path $packageRoot 'lib\ControlledRerun.Common.ps1')
Assert-CRAdministrator
if (-not $IConfirmMonitoringConfigured) {
    throw 'Falta -IConfirmMonitoringConfigured. Confirme en GUI P1/P2/P3/P4, SOC_v3, JSONL on y Discord off.'
}
if ([string]::IsNullOrWhiteSpace($CampaignId)) {
    $CampaignId = 'FP_{0}' -f (Get-Date -Format 'yyyyMMdd_HHmmss')
}
if ($CampaignId -notmatch '^FP_[A-Za-z0-9_-]+$') {
    throw 'CampaignId debe comenzar por FP_ y contener solo letras, números, guion o guion bajo.'
}

$root = Get-CRProjectRoot -RequestedRoot $ProjectRoot -PackageRoot $packageRoot
$runner = Join-Path $root '03_RUNNERS\TFM_Run_FP_Tests_v1.ps1'
$campaignRoot = Join-Path $packageRoot ("OUTPUT\FP\{0}" -f $CampaignId)
if ((Test-Path -LiteralPath $campaignRoot) -and @(Get-ChildItem -LiteralPath $campaignRoot -Force).Count -gt 0) {
    throw "La campaña FP ya existe y no se sobrescribe: $campaignRoot"
}
New-CRDirectory -Path $campaignRoot | Out-Null
$backupDir = New-CRDirectory -Path (Join-Path $campaignRoot 'preexisting_jsonl_backup')
$workspaceRoot = New-CRDirectory -Path (Join-Path $campaignRoot 'workspaces')
$listenerLog = Join-Path $campaignRoot 'fp009_benign_listener.log'

$requiredRoles = @('P1_CRITICAL','P2_EVENT','P3_EVENT','P4','SOC_ROUTER','FP_RUNNER')
$hashChecks = @(Test-CRConfigHashes -ProjectRoot $root | Where-Object { $_.Role -in $requiredRoles })
$badHashes = @($hashChecks | Where-Object { -not $_.Match })
if ($badHashes.Count -gt 0) {
    $badHashes | Format-Table Role,RelativePath,ExpectedSHA256,ActualSHA256 -AutoSize
    throw 'Precondición fallida: configuración o runner FP con hash distinto.'
}
@($hashChecks) | Export-Csv -LiteralPath (Join-Path $campaignRoot 'preflight_hashes.csv') -NoTypeInformation -Encoding UTF8
Copy-CRConfigSnapshot -ProjectRoot $root -Destination (Join-Path $campaignRoot 'config_snapshot') -Roles $requiredRoles | Out-Null

if (Test-NetConnection -ComputerName '127.0.0.1' -Port 80 -InformationLevel Quiet -WarningAction SilentlyContinue) {
    throw 'El puerto 80 ya está ocupado. No se usará un listener no controlado para FP-009.'
}

$listenerJob = $null
$overallFailure = $false
$repetitionRecords = New-Object System.Collections.Generic.List[object]
try {
    $listenerJob = Start-Job -Name ("TFM_FP009_{0}" -f $CampaignId) -ArgumentList 80,$listenerLog -ScriptBlock {
        param($Port,$LogPath)
        $encoding = New-Object System.Text.UTF8Encoding($false)
        $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback,$Port)
        $listener.Start()
        [System.IO.File]::AppendAllText($LogPath, "START $(Get-Date -Format o) 127.0.0.1:$Port`r`n", $encoding)
        try {
            while ($true) {
                $client = $listener.AcceptTcpClient()
                try {
                    $stream = $client.GetStream()
                    $stream.ReadTimeout = 5000
                    $buffer = New-Object byte[] 8192
                    $read = $stream.Read($buffer,0,$buffer.Length)
                    $request = if($read -gt 0){[System.Text.Encoding]::ASCII.GetString($buffer,0,$read)}else{''}
                    $firstLine = ($request -split "`r?`n")[0]
                    $body = 'TFM benign FP009 listener'
                    $response = "HTTP/1.1 200 OK`r`nContent-Type: text/plain`r`nContent-Length: $($body.Length)`r`nConnection: close`r`n`r`n$body"
                    $bytes = [System.Text.Encoding]::ASCII.GetBytes($response)
                    $stream.Write($bytes,0,$bytes.Length)
                    $stream.Flush()
                    [System.IO.File]::AppendAllText($LogPath, "REQUEST $(Get-Date -Format o) $firstLine STATUS=200`r`n", $encoding)
                } finally {
                    $client.Close()
                }
            }
        } finally {
            $listener.Stop()
            [System.IO.File]::AppendAllText($LogPath, "STOP $(Get-Date -Format o)`r`n", $encoding)
        }
    }
    Start-Sleep -Seconds 2
    if (-not (Test-NetConnection -ComputerName '127.0.0.1' -Port 80 -InformationLevel Quiet -WarningAction SilentlyContinue)) {
        $jobOutput = Receive-Job -Job $listenerJob -Keep -ErrorAction SilentlyContinue | Out-String
        throw "No se pudo iniciar el listener benigno FP-009. Job: $jobOutput"
    }
    Write-CRStatus -Message 'Listener benigno FP-009 activo en 127.0.0.1:80.' -Level OK

    $preexistingBackup = Backup-AndClear-CRJsonl -JsonlPath $JsonlPath -BackupDirectory $backupDir -Label $CampaignId

    for ($rep = 1; $rep -le $Repetitions; $rep++) {
        if ($rep -gt 1) {
            Write-CRStatus -Message "Ventana de silencio ${InterRepetitionQuietSeconds}s antes de REP_$('{0:D2}' -f $rep)." -Level INFO
            Start-Sleep -Seconds $InterRepetitionQuietSeconds
            Write-CRUtf8NoBom -Path $JsonlPath -Text ''
        }
        $repName = 'REP_{0:D2}' -f $rep
        $repRoot = New-CRDirectory -Path (Join-Path $campaignRoot $repName)
        $runnerOutput = New-CRDirectory -Path (Join-Path $repRoot 'runner_native')
        $exportDir = New-CRDirectory -Path (Join-Path $repRoot 'velociraptor_exports')
        $startLocal = Get-Date
        $startUtc = $startLocal.ToUniversalTime()
        Write-CRStatus -Message "Ejecutando $CampaignId $repName" -Level INFO

        $runnerError = ''
        try {
            & $runner `
                -ClearJsonlBeforeRun $false `
                -KeepWorkspace $true `
                -JsonlPath $JsonlPath `
                -OutputDir $runnerOutput `
                -WorkspaceRoot $workspaceRoot `
                -PostRunWaitSeconds $PostRunWaitSeconds 2>&1 |
                Tee-Object -LiteralPath (Join-Path $runnerOutput 'wrapper_console.log')
        } catch {
            $runnerError = $_.Exception.ToString()
            $overallFailure = $true
        }
        $endLocal = Get-Date
        $endUtc = $endLocal.ToUniversalTime()
        $rawJsonl = Join-Path $repRoot ("soc_alerts_{0}_{1}.jsonl" -f $CampaignId,$repName)
        Copy-Item -LiteralPath $JsonlPath -Destination $rawJsonl -Force
        $jsonlRead = Read-CRJsonl -Path $rawJsonl

        $summaryFiles = @(Get-ChildItem -LiteralPath $runnerOutput -Filter 'TFM_FP_*_summary.json' -File | Sort-Object LastWriteTime -Descending)
        $summary = $null
        $summaryPath = ''
        if ($summaryFiles.Count -gt 0) {
            $summaryPath = $summaryFiles[0].FullName
            $summary = Get-Content -LiteralPath $summaryPath -Raw -Encoding UTF8 | ConvertFrom-Json
        } else {
            $runnerError = ($runnerError + ' No existe summary.json FP.').Trim()
            $overallFailure = $true
        }

        $testIds = if($summary){@($summary.Tests | ForEach-Object {[string]$_.FP_ID} | Sort-Object -Unique)}else{@()}
        $expectedIds = 1..10 | ForEach-Object { 'FP-{0:D3}' -f $_ }
        $missingIds = @($expectedIds | Where-Object { $_ -notin $testIds })
        $skipped = if($summary){[int]$summary.TestCounts.SKIPPED}else{99}
        $failed = if($summary){[int]$summary.TestCounts.FAIL}else{99}
        $p1Hits = @($jsonlRead.Rows | Where-Object { [string]$_.Profile -eq 'P1_CRITICAL' }).Count
        if ($missingIds.Count -gt 0 -or $skipped -gt 0 -or $failed -gt 0 -or @($jsonlRead.BadLines).Count -gt 0 -or $p1Hits -gt 0 -or $runnerError) {
            $overallFailure = $true
        }

        $manifest = [pscustomobject]@{
            SchemaVersion = '1.0'
            Mode = 'FP_CONTROLLED'
            CampaignId = $CampaignId
            Repetition = $rep
            RepetitionName = $repName
            ComputerName = $env:COMPUTERNAME
            JsonlPath = $JsonlPath
            RawJsonl = $rawJsonl
            RawJsonlSHA256 = Get-CRSha256 -Path $rawJsonl
            StartLocal = $startLocal.ToString('o')
            StartUtc = $startUtc.ToString('o')
            EndLocal = $endLocal.ToString('o')
            EndUtc = $endUtc.ToString('o')
            RunnerSummary = $summaryPath
            RunnerRunId = if($summary){[string]$summary.RunId}else{''}
            TestsExpected = 10
            MissingTestIds = $missingIds
            Skipped = $skipped
            Failed = $failed
            JsonlValidRows = @($jsonlRead.Rows).Count
            JsonlBadRows = @($jsonlRead.BadLines).Count
            P1CriticalHits = $p1Hits
            RunnerError = $runnerError
        }
        Write-CRJson -Object $manifest -Path (Join-Path $repRoot 'repetition_manifest.json')
        Write-CRUtf8NoBom -Path (Join-Path $exportDir 'REQUIRED_EXPORTS.txt') -Text "Exportar P1/P2/P3/P4 CLIENT_EVENT para StartUtc=$($startUtc.ToString('o')) EndUtc=$($endUtc.ToString('o')). No usar el router como sustituto."
        Export-CRHashManifest -Root $repRoot -OutputPath (Join-Path $repRoot 'HASHES_SHA256.csv')
        [void]$repetitionRecords.Add($manifest)
        Write-CRStatus -Message ("{0}: missing={1}; skipped={2}; failed={3}; badJSON={4}; P1={5}" -f $repName,($missingIds -join ','),$skipped,$failed,@($jsonlRead.BadLines).Count,$p1Hits) -Level $(if($overallFailure){'WARN'}else{'OK'})
    }
} finally {
    if ($listenerJob) {
        Stop-Job -Job $listenerJob -ErrorAction SilentlyContinue
        Receive-Job -Job $listenerJob -ErrorAction SilentlyContinue | Out-Null
        Remove-Job -Job $listenerJob -Force -ErrorAction SilentlyContinue
    }
}

$campaignManifest = [pscustomobject]@{
    SchemaVersion = '1.0'
    Mode = 'FP_CONTROLLED_SET'
    CampaignId = $CampaignId
    RepetitionsRequested = $Repetitions
    RepetitionsCompleted = $repetitionRecords.Count
    Listener = '127.0.0.1:80 controlled TcpListener'
    ListenerLog = $listenerLog
    JsonlSeparation = 'TRUNCATE_BEFORE_EACH_REPETITION'
    OverallAutomatedStatus = if($overallFailure){'FAIL'}else{'PASS'}
    Repetitions = $repetitionRecords.ToArray()
}
Write-CRJson -Object $campaignManifest -Path (Join-Path $campaignRoot 'campaign_manifest.json')
Export-CRHashManifest -Root $campaignRoot -OutputPath (Join-Path $campaignRoot 'HASHES_SHA256.csv')
if ($overallFailure) {
    Write-CRStatus -Message "Campaña FP capturada con incidencias. No usarla para Excel. Salida: $campaignRoot" -Level FAIL
    exit 2
}
Write-CRStatus -Message "Campaña FP completada: $Repetitions repeticiones, FP-001..FP-010, sin SKIPPED ni P1." -Level OK
