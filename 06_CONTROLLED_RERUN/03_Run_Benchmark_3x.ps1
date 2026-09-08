[CmdletBinding()]
param(
    [string]$CampaignId = '',
    [string]$ProjectRoot = '',
    [string]$JsonlPath = '\\VBOXSVR\Carpeta_Compartida_TFM\inbox\soc_alerts.jsonl',
    [ValidateRange(1,10)][int]$Repetitions = 3,
    [int]$BaselineDurationSec = 180,
    [int]$IdleDurationSec = 180,
    [int]$RunnerDurationSec = 300,
    [int]$IntervalSec = 2,
    [int]$ServiceSettleSec = 30,
    [int]$IdleWarmupSec = 30,
    [string]$ServiceName = 'Velociraptor'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$packageRoot = $PSScriptRoot
. (Join-Path $packageRoot 'lib\ControlledRerun.Common.ps1')
Assert-CRAdministrator
if ([string]::IsNullOrWhiteSpace($CampaignId)) {
    $CampaignId = 'BENCH_{0}' -f (Get-Date -Format 'yyyyMMdd_HHmmss')
}
if ($CampaignId -notmatch '^BENCH_[A-Za-z0-9_-]+$') {
    throw 'CampaignId debe comenzar por BENCH_.'
}

$root = Get-CRProjectRoot -RequestedRoot $ProjectRoot -PackageRoot $packageRoot
$benchmark = Join-Path $root '03_RUNNERS\validate\TFM_Benchmark_VR_Resource_Usage_v1.ps1'
$loadRunner = Join-Path $packageRoot 'support\TFM_Benchmark_TEC_Load_v1.ps1'
$campaignRoot = Join-Path $packageRoot ("OUTPUT\BENCHMARK\{0}" -f $CampaignId)
if ((Test-Path -LiteralPath $campaignRoot) -and @(Get-ChildItem -LiteralPath $campaignRoot -Force).Count -gt 0) {
    throw "La campaña benchmark ya existe y no se sobrescribe: $campaignRoot"
}
New-CRDirectory -Path $campaignRoot | Out-Null

$benchmarkHash = @(Test-CRConfigHashes -ProjectRoot $root | Where-Object { $_.Role -eq 'BENCHMARK_RUNNER' })
if ($benchmarkHash.Count -ne 1 -or -not $benchmarkHash[0].Match) {
    throw 'El benchmark v1 no coincide con el hash controlado de SchemaVersion 1.1.'
}
if (-not (Test-Path -LiteralPath $loadRunner -PathType Leaf)) {
    throw "No existe el lanzador de carga: $loadRunner"
}
$service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if (-not $service) {
    throw "No existe el servicio $ServiceName."
}

$preexisting = Backup-AndClear-CRJsonl -JsonlPath $JsonlPath -BackupDirectory (Join-Path $campaignRoot 'preexisting_jsonl_backup') -Label $CampaignId
$campaignStart = Get-Date
$runRecords = New-Object System.Collections.Generic.List[object]
$failed = $false

function Set-ServiceState {
    param([ValidateSet('Running','Stopped')][string]$Desired)
    $svc = Get-Service -Name $ServiceName -ErrorAction Stop
    if ($Desired -eq 'Stopped' -and $svc.Status -ne 'Stopped') {
        Stop-Service -Name $ServiceName -Force -ErrorAction Stop
    } elseif ($Desired -eq 'Running' -and $svc.Status -ne 'Running') {
        Start-Service -Name $ServiceName -ErrorAction Stop
    }
    $desiredEnum = [System.ServiceProcess.ServiceControllerStatus]::Stopped
    if ($Desired -eq 'Running') { $desiredEnum = [System.ServiceProcess.ServiceControllerStatus]::Running }
    $svc.WaitForStatus($desiredEnum,[TimeSpan]::FromSeconds(30))
    Start-Sleep -Seconds $ServiceSettleSec
}

function Invoke-ControlledBenchmark {
    param(
        [int]$Repetition,
        [ValidateSet('BASELINE_NO_VR','VR_IDLE','VR_TEC_RUNNER')][string]$Scenario,
        [int]$DurationSec
    )
    $repName = 'REP_{0:D2}' -f $Repetition
    $scenarioRoot = New-CRDirectory -Path (Join-Path $campaignRoot (Join-Path $repName $Scenario))
    $before = @(Get-ChildItem -LiteralPath $scenarioRoot -Directory -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName)
    $start = Get-Date
    $params = @{
        ScenarioName = $Scenario
        DurationSec = $DurationSec
        IntervalSec = $IntervalSec
        OutputDir = $scenarioRoot
        SkipServiceControl = $true
        IncludeServerGuiInTotal = $false
        WarmupSec = 0
    }
    if ($Scenario -eq 'VR_IDLE') {
        $params.WarmupSec = $IdleWarmupSec
    }
    if ($Scenario -eq 'VR_TEC_RUNNER') {
        $params.RunnerPath = $loadRunner
    }
    Write-CRStatus -Message "$repName $Scenario iniciado." -Level INFO
    & $benchmark @params 2>&1 | Tee-Object -LiteralPath (Join-Path $scenarioRoot 'wrapper_console.log')
    $end = Get-Date
    $runDirs = @(Get-ChildItem -LiteralPath $scenarioRoot -Directory | Where-Object { $_.FullName -notin $before } | Sort-Object LastWriteTime -Descending)
    $summaryPath = ''
    $summary = $null
    if ($runDirs.Count -gt 0 -and (Test-Path -LiteralPath (Join-Path $runDirs[0].FullName 'summary.json'))) {
        $summaryPath = Join-Path $runDirs[0].FullName 'summary.json'
        $summary = Get-Content -LiteralPath $summaryPath -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    $valid = $false
    $reason = 'summary.json ausente'
    if ($summary) {
        $valid = ([string]$summary.SchemaVersion -eq '1.1') -and ([string]$summary.ScenarioValidity -eq 'VALID') -and ([bool]$summary.ServerGuiExcludedFromClientMetrics) -and (-not [bool]$summary.IncludeServerGuiInTotal)
        if ($Scenario -eq 'BASELINE_NO_VR') {
            $valid = $valid -and ([int]$summary.VR_Client_ProcessCountMax -eq 0)
        } else {
            $valid = $valid -and ([int]$summary.VR_Client_ProcessCountMax -gt 0)
        }
        if ($Scenario -eq 'VR_TEC_RUNNER') {
            $valid = $valid -and ([int]$summary.RunnerObservedSamples -gt 0)
        }
        $reason = (@($summary.ValidityReasons) -join '; ')
    }
    $record = [pscustomobject]@{
        Repetition = $Repetition
        RepetitionName = $repName
        Scenario = $Scenario
        WrapperStart = $start.ToString('o')
        WrapperEnd = $end.ToString('o')
        SummaryPath = $summaryPath
        SchemaVersion = if($summary){[string]$summary.SchemaVersion}else{''}
        ScenarioValidity = if($summary){[string]$summary.ScenarioValidity}else{'MISSING'}
        ClientProcessMax = if($summary){[int]$summary.VR_Client_ProcessCountMax}else{-1}
        RunnerObservedSamples = if($summary){[int]$summary.RunnerObservedSamples}else{-1}
        ServerGuiExcluded = if($summary){[bool]$summary.ServerGuiExcludedFromClientMetrics}else{$false}
        IncludeServerGuiInTotal = if($summary){[bool]$summary.IncludeServerGuiInTotal}else{$true}
        AutomatedPass = $valid
        Evidence = $reason
    }
    [void]$runRecords.Add($record)
    Write-CRJson -Object $record -Path (Join-Path $scenarioRoot 'wrapper_result.json')
    Export-CRHashManifest -Root $scenarioRoot -OutputPath (Join-Path $scenarioRoot 'HASHES_SHA256.csv')
    if (-not $valid) {
        $script:failed = $true
        throw "Benchmark inválido en $repName/$Scenario. $reason"
    }
    Write-CRStatus -Message "$repName $Scenario VALID." -Level OK
}

try {
    for ($rep = 1; $rep -le $Repetitions; $rep++) {
        Set-ServiceState -Desired 'Stopped'
        Invoke-ControlledBenchmark -Repetition $rep -Scenario 'BASELINE_NO_VR' -DurationSec $BaselineDurationSec

        Set-ServiceState -Desired 'Running'
        Invoke-ControlledBenchmark -Repetition $rep -Scenario 'VR_IDLE' -DurationSec $IdleDurationSec

        Set-ServiceState -Desired 'Running'
        Invoke-ControlledBenchmark -Repetition $rep -Scenario 'VR_TEC_RUNNER' -DurationSec $RunnerDurationSec
    }
} catch {
    $failed = $true
    Write-CRStatus -Message $_.Exception.Message -Level FAIL
} finally {
    try {
        $finalService = Get-Service -Name $ServiceName -ErrorAction Stop
        if ($finalService.Status -ne 'Running') {
            Start-Service -Name $ServiceName -ErrorAction Stop
            $finalService.WaitForStatus([System.ServiceProcess.ServiceControllerStatus]::Running,[TimeSpan]::FromSeconds(30))
        }
    } catch {
        $failed = $true
        Write-CRStatus -Message "No se pudo dejar el servicio Running: $($_.Exception.Message)" -Level FAIL
    }
}

$campaignEnd = Get-Date
$benchmarkJsonl = Join-Path $campaignRoot ("soc_alerts_{0}_PERFORMANCE_ONLY.jsonl" -f $CampaignId)
if (Test-Path -LiteralPath $JsonlPath -PathType Leaf) {
    Copy-Item -LiteralPath $JsonlPath -Destination $benchmarkJsonl -Force
}
$manifest = [pscustomobject]@{
    SchemaVersion = '1.0'
    Mode = 'BENCHMARK_CONTROLLED_SET'
    CampaignId = $CampaignId
    StartLocal = $campaignStart.ToString('o')
    EndLocal = $campaignEnd.ToString('o')
    RepetitionsRequested = $Repetitions
    ExpectedRuns = 3 * $Repetitions
    CompletedRuns = $runRecords.Count
    Scenarios = @('BASELINE_NO_VR','VR_IDLE','VR_TEC_RUNNER')
    BenchmarkScript = $benchmark
    BenchmarkScriptSHA256 = Get-CRSha256 -Path $benchmark
    LoadRunner = $loadRunner
    LoadRunnerSHA256 = Get-CRSha256 -Path $loadRunner
    JsonlClassification = 'PERFORMANCE_ONLY_NOT_DETECTION_EVIDENCE'
    PreexistingJsonlBackup = $preexisting
    AutomatedStatus = if($failed){'FAIL'}else{'PASS'}
    Runs = $runRecords.ToArray()
}
Write-CRJson -Object $manifest -Path (Join-Path $campaignRoot 'campaign_manifest.json')
$runRecords.ToArray() | Export-Csv -LiteralPath (Join-Path $campaignRoot 'benchmark_run_index.csv') -NoTypeInformation -Encoding UTF8
Export-CRHashManifest -Root $campaignRoot -OutputPath (Join-Path $campaignRoot 'HASHES_SHA256.csv')
if ($failed -or $runRecords.Count -ne (3 * $Repetitions)) {
    Write-CRStatus -Message "Benchmark incompleto o inválido. Salida conservada: $campaignRoot" -Level FAIL
    exit 2
}
Write-CRStatus -Message "Benchmark válido: $($runRecords.Count) ejecuciones separadas. Salida: $campaignRoot" -Level OK
