[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$TecCampaignId,
    [Parameter(Mandatory=$true)][string]$FpCampaignId,
    [Parameter(Mandatory=$true)][string]$BenchmarkCampaignId,
    [string]$OutputRoot = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$packageRoot = $PSScriptRoot
. (Join-Path $packageRoot 'lib\ControlledRerun.Common.ps1')
if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $OutputRoot = Join-Path $packageRoot 'OUTPUT'
}

$validationId = 'VALIDATION_{0}' -f (Get-Date -Format 'yyyyMMdd_HHmmss')
$validationRoot = New-CRDirectory -Path (Join-Path $OutputRoot ("VALIDATION\{0}" -f $validationId))
$results = New-Object System.Collections.Generic.List[object]

function Add-Validation {
    param(
        [ValidateSet('BLOCKER','MAJOR','MINOR','INFO')][string]$Severity,
        [string]$Scope,
        [string]$Check,
        [bool]$Pass,
        [string]$Evidence,
        [string]$Path = ''
    )
    [void]$results.Add([pscustomobject]@{
        Severity = $Severity
        Scope = $Scope
        Check = $Check
        Pass = $Pass
        Evidence = $Evidence
        Path = $Path
    })
    $level = if($Pass){'OK'}elseif($Severity -in @('BLOCKER','MAJOR')){'FAIL'}else{'WARN'}
    Write-CRStatus -Message ("{0} | {1} | {2}" -f $Scope,$Check,$Evidence) -Level $level
}

function Get-SafePropertyValue {
    param(
        $Object,
        [Parameter(Mandatory=$true)][string]$Name,
        $Default = $null
    )
    if ($null -eq $Object) { return $Default }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $Default }
    return $property.Value
}

function Get-SafeNameList {
    param([object[]]$Items)
    return @($Items | ForEach-Object {
        $value = Get-SafePropertyValue -Object $_ -Name 'Name' -Default ''
        if (-not [string]::IsNullOrWhiteSpace([string]$value)) { [string]$value }
    })
}

function Test-TimestampsInWindow {
    param(
        [object[]]$Rows,
        [datetimeoffset]$StartUtc,
        [datetimeoffset]$EndUtc,
        [int]$ToleranceSeconds = 60
    )
    $unparseable = 0
    $outside = 0
    $parsed = 0
    $min = $StartUtc.AddSeconds(-$ToleranceSeconds)
    $max = $EndUtc.AddSeconds($ToleranceSeconds)
    foreach ($row in @($Rows)) {
        # DetectionTime pertenece al CLIENT_EVENT original. Timestamp puede ser
        # la hora posterior de persistencia del router y no sirve para asignar
        # una detección retrasada a una repetición FP.
        $raw = [string](Get-SafePropertyValue -Object $row -Name 'EventTimeUtc' -Default '')
        if ([string]::IsNullOrWhiteSpace($raw)) { $raw = [string](Get-SafePropertyValue -Object $row -Name 'DetectionTime' -Default '') }
        if ([string]::IsNullOrWhiteSpace($raw)) { $raw = [string](Get-SafePropertyValue -Object $row -Name 'Timestamp' -Default '') }
        if ([string]::IsNullOrWhiteSpace($raw)) { $raw = [string](Get-SafePropertyValue -Object $row -Name '_ts' -Default '') }
        $dto = [datetimeoffset]::MinValue
        $number = 0.0
        $parsedOk = $false
        if([double]::TryParse($raw,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$number)){
            try{
                $milliseconds = if([Math]::Abs($number) -ge 100000000000){[long][Math]::Round($number)}else{[long][Math]::Round($number * 1000.0)}
                $dto = [DateTimeOffset]::FromUnixTimeMilliseconds($milliseconds)
                $parsedOk = $true
            }catch{}
        } elseif([datetimeoffset]::TryParse($raw,[ref]$dto)) {
            $parsedOk = $true
        }
        if (-not $parsedOk) {
            $unparseable++
            continue
        }
        $parsed++
        if ($dto.ToUniversalTime() -lt $min -or $dto.ToUniversalTime() -gt $max) { $outside++ }
    }
    return [pscustomobject]@{Parsed=$parsed;Unparseable=$unparseable;Outside=$outside;Total=@($Rows).Count}
}

function Get-ExpectedExports {
    param([string]$Directory)
    $names = @('P1_CRITICAL_CLIENT_EVENT.csv','P2_EVENT_CLIENT_EVENT.csv','P3_EVENT_CLIENT_EVENT.csv','P4_CLIENT_EVENT.csv')
    return @($names | ForEach-Object {
        $p = Join-Path $Directory $_
        [pscustomobject]@{Name=$_;Path=$p;Exists=(Test-Path -LiteralPath $p -PathType Leaf);Length=$(if(Test-Path -LiteralPath $p -PathType Leaf){(Get-Item -LiteralPath $p).Length}else{0})}
    })
}

function Read-ClientEventExportSet {
    param([string]$Directory)
    $rows = New-Object System.Collections.ArrayList
    $errors = New-Object System.Collections.ArrayList
    $schemaBad = New-Object System.Collections.ArrayList
    foreach($item in @(Get-ExpectedExports -Directory $Directory)){
        if(-not $item.Exists -or $item.Length -le 0){continue}
        try{
            $header = [string](Get-Content -LiteralPath $item.Path -TotalCount 1)
            if($header -notmatch '(^|,)"?Artifact"?(,|$)' -or $header -notmatch '(^|,)"?Source"?(,|$)'){
                [void]$schemaBad.Add($item.Name)
            }
            foreach($row in @(Import-Csv -LiteralPath $item.Path)){[void]$rows.Add($row)}
        }catch{[void]$errors.Add(("{0}: {1}" -f $item.Name,$_.Exception.Message))}
    }
    return [pscustomobject]@{Rows=$rows.ToArray();Errors=$errors.ToArray();SchemaBad=$schemaBad.ToArray()}
}

function Test-RequiredFiles {
    param([string]$RunDirectory,[string[]]$Names)
    return @($Names | ForEach-Object {
        $p = Join-Path $RunDirectory $_
        [pscustomobject]@{Name=$_;Exists=(Test-Path -LiteralPath $p -PathType Leaf);Path=$p}
    })
}

# ---------------------------------------------------------------------------
# TEC canónica
# ---------------------------------------------------------------------------
$tecRoot = Join-Path $OutputRoot ("TEC\{0}" -f $TecCampaignId)
Add-Validation -Severity BLOCKER -Scope TEC -Check 'Carpeta campaña' -Pass (Test-Path -LiteralPath $tecRoot -PathType Container) -Evidence $tecRoot -Path $tecRoot
if (Test-Path -LiteralPath $tecRoot -PathType Container) {
    $tecManifestPath = Join-Path $tecRoot 'campaign_manifest.json'
    $tecManifest = $null
    try { $tecManifest = Get-Content -LiteralPath $tecManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json } catch {}
    Add-Validation -Severity BLOCKER -Scope TEC -Check 'Manifest parseable' -Pass ($null -ne $tecManifest) -Evidence $tecManifestPath -Path $tecManifestPath

    $tecJsonlPath = Join-Path $tecRoot ("raw\soc_alerts_{0}.jsonl" -f $TecCampaignId)
    $tecJsonl = Read-CRJsonl -Path $tecJsonlPath
    Add-Validation -Severity BLOCKER -Scope TEC -Check 'JSONL válido' -Pass (@($tecJsonl.BadLines).Count -eq 0 -and (Test-Path -LiteralPath $tecJsonlPath)) -Evidence ("valid={0}; bad={1}" -f @($tecJsonl.Rows).Count,@($tecJsonl.BadLines).Count) -Path $tecJsonlPath

    $expectedTecs = 1..9 | ForEach-Object { 'TEC-{0:D3}' -f $_ }
    $actualTecs = @($tecJsonl.Rows | ForEach-Object { [string]$_.TEC } | Where-Object { $_ -match '^TEC-00[1-9]$' } | Sort-Object -Unique)
    $missingTecs = @($expectedTecs | Where-Object { $_ -notin $actualTecs })
    Add-Validation -Severity BLOCKER -Scope TEC -Check 'Cobertura TEC-001..TEC-009' -Pass ($missingTecs.Count -eq 0) -Evidence ("present={0}; missing={1}" -f ($actualTecs -join ','),($missingTecs -join ',')) -Path $tecJsonlPath

    if ($tecManifest) {
        $startUtc = [datetimeoffset]::Parse([string]$tecManifest.StartUtc)
        $endUtc = [datetimeoffset]::Parse([string]$tecManifest.EndUtc)
        $timeCheck = Test-TimestampsInWindow -Rows @($tecJsonl.Rows) -StartUtc $startUtc -EndUtc $endUtc
        Add-Validation -Severity BLOCKER -Scope TEC -Check 'Timestamps dentro de campaña' -Pass ($timeCheck.Unparseable -eq 0 -and $timeCheck.Outside -eq 0) -Evidence ("parsed={0}; unparseable={1}; outside={2}" -f $timeCheck.Parsed,$timeCheck.Unparseable,$timeCheck.Outside) -Path $tecJsonlPath
        Add-Validation -Severity MAJOR -Scope TEC -Check 'Hash JSONL coincide con manifest' -Pass ((Get-CRSha256 -Path $tecJsonlPath) -eq [string]$tecManifest.RawJsonlSHA256) -Evidence (Get-CRSha256 -Path $tecJsonlPath) -Path $tecJsonlPath
    }

    $summaryFiles = @(Get-ChildItem -LiteralPath (Join-Path $tecRoot 'raw\runner') -Filter '*_summary.json' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)
    $tecSummary = $null
    if ($summaryFiles.Count -gt 0) {
        try { $tecSummary = Get-Content -LiteralPath $summaryFiles[0].FullName -Raw -Encoding UTF8 | ConvertFrom-Json } catch {}
    }
    Add-Validation -Severity BLOCKER -Scope TEC -Check 'summary.json runner' -Pass ($null -ne $tecSummary) -Evidence $(if($summaryFiles.Count -gt 0){$summaryFiles[0].FullName}else{'ausente'})
    if ($tecSummary) {
        $techniques = @($tecSummary.Techniques)
        $badTechnique = @($techniques | Where-Object { [string]$_.Status -ne 'OK' })
        $summaryIds = @($techniques | ForEach-Object { [string]$_.TEC_ID } | Sort-Object -Unique)
        Add-Validation -Severity BLOCKER -Scope TEC -Check 'Runner 9 OK / 0 WARN / 0 FAIL' -Pass ($techniques.Count -eq 9 -and $badTechnique.Count -eq 0 -and [string]$tecSummary.Campaign.GlobalStatus -eq 'OK') -Evidence ("count={0}; global={1}; bad={2}; ids={3}" -f $techniques.Count,$tecSummary.Campaign.GlobalStatus,$badTechnique.Count,($summaryIds -join ',')) -Path $summaryFiles[0].FullName

        $localHash = ([string]$tecSummary.TEC009.LocalSHA256).ToUpperInvariant()
        $sourceFiles = @(Get-ChildItem -LiteralPath (Join-Path $tecRoot 'raw\tec009_source') -File -ErrorAction SilentlyContinue)
        $sourceMatches = @($sourceFiles | Where-Object { (Get-CRSha256 -Path $_.FullName) -eq $localHash })
        Add-Validation -Severity BLOCKER -Scope TEC009 -Check 'Hash ZIP origen' -Pass (-not [string]::IsNullOrWhiteSpace($localHash) -and $sourceMatches.Count -gt 0) -Evidence ("runnerHash={0}; matchingSourceFiles={1}" -f $localHash,$sourceMatches.Count) -Path (Join-Path $tecRoot 'raw\tec009_source')

        $httpStatus = 0
        $httpParsed = [int]::TryParse(([string]$tecSummary.TEC009.HttpStatus).Trim(),[ref]$httpStatus)
        $httpOk = $httpParsed -and $httpStatus -ge 200 -and $httpStatus -le 299 -and ([string]$tecSummary.TEC009.UploadSucceeded -match '^(?i:true)$')
        Add-Validation -Severity BLOCKER -Scope TEC009 -Check 'HTTP 2xx y UploadSucceeded' -Pass $httpOk -Evidence ("HTTP={0}; UploadSucceeded={1}" -f $tecSummary.TEC009.HttpStatus,$tecSummary.TEC009.UploadSucceeded) -Path $summaryFiles[0].FullName

        $receiverRoot = Join-Path $OutputRoot ("RECEIVER\{0}" -f $TecCampaignId)
        $receiverLog = Join-Path $receiverRoot 'receiver_log.jsonl'
        $receiverJsonl = Read-CRJsonl -Path $receiverLog
        $matchingPosts = @($receiverJsonl.Rows | Where-Object {
            $methodProperty = $_.PSObject.Properties['method']
            $hashProperty = $_.PSObject.Properties['sha256']
            $typeProperty = $_.PSObject.Properties['content_type']
            $methodProperty -and $hashProperty -and $typeProperty -and
            [string]$methodProperty.Value -eq 'POST' -and
            ([string]$hashProperty.Value).ToUpperInvariant() -eq $localHash -and
            [string]$typeProperty.Value -eq 'application/zip'
        })
        $receivedZips = @(Get-ChildItem -LiteralPath (Join-Path $receiverRoot 'received') -Filter '*.zip' -File -ErrorAction SilentlyContinue)
        $receivedMatches = @($receivedZips | Where-Object { (Get-CRSha256 -Path $_.FullName) -eq $localHash })
        Add-Validation -Severity BLOCKER -Scope TEC009 -Check 'Receiver POST y hash destino' -Pass (@($receiverJsonl.BadLines).Count -eq 0 -and $matchingPosts.Count -eq 1 -and $receivedMatches.Count -eq 1) -Evidence ("postsMatching={0}; zipsMatching={1}; receiverBadJson={2}" -f $matchingPosts.Count,$receivedMatches.Count,@($receiverJsonl.BadLines).Count) -Path $receiverRoot
    }

    $exports = Get-ExpectedExports -Directory (Join-Path $tecRoot 'velociraptor_exports')
    $missingExports = @($exports | Where-Object { -not $_.Exists -or $_.Length -le 0 })
    Add-Validation -Severity BLOCKER -Scope TEC -Check 'Exports CLIENT_EVENT P1/P2/P3/P4' -Pass ($missingExports.Count -eq 0) -Evidence ("missing={0}" -f ((Get-SafeNameList -Items $missingExports) -join ',')) -Path (Join-Path $tecRoot 'velociraptor_exports')
    if($missingExports.Count -eq 0 -and $tecManifest){
        $exportSet = Read-ClientEventExportSet -Directory (Join-Path $tecRoot 'velociraptor_exports')
        Add-Validation -Severity BLOCKER -Scope TEC -Check 'Exports CLIENT_EVENT parseables' -Pass (@($exportSet.Errors).Count -eq 0) -Evidence ("rows={0}; errors={1}" -f @($exportSet.Rows).Count,(@($exportSet.Errors) -join '; ')) -Path (Join-Path $tecRoot 'velociraptor_exports')
        Add-Validation -Severity MAJOR -Scope TEC -Check 'Exports conservan Artifact y Source' -Pass (@($exportSet.SchemaBad).Count -eq 0) -Evidence ("bad={0}" -f (@($exportSet.SchemaBad) -join ',')) -Path (Join-Path $tecRoot 'velociraptor_exports')
        $exportTime = Test-TimestampsInWindow -Rows @($exportSet.Rows) -StartUtc ([datetimeoffset]::Parse([string]$tecManifest.StartUtc)) -EndUtc ([datetimeoffset]::Parse([string]$tecManifest.EndUtc)) -ToleranceSeconds 0
        Add-Validation -Severity BLOCKER -Scope TEC -Check 'Cronología CLIENT_EVENT limitada a FINAL02' -Pass ($exportTime.Unparseable -eq 0 -and $exportTime.Outside -eq 0) -Evidence ("parsed={0}; unparseable={1}; outside={2}" -f $exportTime.Parsed,$exportTime.Unparseable,$exportTime.Outside) -Path (Join-Path $tecRoot 'velociraptor_exports')
        $exportTecs = @($exportSet.Rows | ForEach-Object {[string](Get-SafePropertyValue -Object $_ -Name 'ID_Tecnica_Interna' -Default '')} | Where-Object {$_ -match '^TEC-00[1-9]$'} | Sort-Object -Unique)
        $missingExportTecs = @($expectedTecs | Where-Object {$_ -notin $exportTecs})
        Add-Validation -Severity BLOCKER -Scope TEC -Check 'Cobertura CLIENT_EVENT TEC-001..TEC-009' -Pass ($missingExportTecs.Count -eq 0) -Evidence ("present={0}; missing={1}" -f ($exportTecs -join ','),($missingExportTecs -join ',')) -Path (Join-Path $tecRoot 'velociraptor_exports')
    }
}

# ---------------------------------------------------------------------------
# Falsos positivos
# ---------------------------------------------------------------------------
$fpRoot = Join-Path $OutputRoot ("FP\{0}" -f $FpCampaignId)
Add-Validation -Severity BLOCKER -Scope FP -Check 'Carpeta campaña' -Pass (Test-Path -LiteralPath $fpRoot -PathType Container) -Evidence $fpRoot -Path $fpRoot
if (Test-Path -LiteralPath $fpRoot -PathType Container) {
    $repDirs = @(Get-ChildItem -LiteralPath $fpRoot -Directory -Filter 'REP_*' | Sort-Object Name)
    Add-Validation -Severity BLOCKER -Scope FP -Check 'Tres repeticiones separadas' -Pass ($repDirs.Count -eq 3) -Evidence ("count={0}; names={1}" -f $repDirs.Count,((Get-SafeNameList -Items $repDirs) -join ',')) -Path $fpRoot
    foreach ($repDir in $repDirs) {
        $scope = "FP/$($repDir.Name)"
        $manifestPath = Join-Path $repDir.FullName 'repetition_manifest.json'
        $manifest = $null
        try { $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json } catch {}
        Add-Validation -Severity BLOCKER -Scope $scope -Check 'Manifest parseable' -Pass ($null -ne $manifest) -Evidence $manifestPath -Path $manifestPath
        $rawFiles = @(Get-ChildItem -LiteralPath $repDir.FullName -Filter 'soc_alerts_*.jsonl' -File)
        $rawPath = if($rawFiles.Count -eq 1){$rawFiles[0].FullName}else{''}
        $read = Read-CRJsonl -Path $rawPath
        Add-Validation -Severity BLOCKER -Scope $scope -Check 'JSONL exclusivo válido' -Pass ($rawFiles.Count -eq 1 -and @($read.BadLines).Count -eq 0) -Evidence ("files={0}; valid={1}; bad={2}" -f $rawFiles.Count,@($read.Rows).Count,@($read.BadLines).Count) -Path $rawPath

        $summaries = @(Get-ChildItem -LiteralPath (Join-Path $repDir.FullName 'runner_native') -Filter 'TFM_FP_*_summary.json' -File | Sort-Object LastWriteTime -Descending)
        $summary = $null
        if($summaries.Count -gt 0){try{$summary=Get-Content -LiteralPath $summaries[0].FullName -Raw -Encoding UTF8 | ConvertFrom-Json}catch{}}
        Add-Validation -Severity BLOCKER -Scope $scope -Check 'summary.json FP' -Pass ($null -ne $summary) -Evidence $(if($summaries.Count){$summaries[0].FullName}else{'ausente'})
        if($summary){
            $ids = @($summary.Tests | ForEach-Object {[string]$_.FP_ID} | Sort-Object -Unique)
            $missing = @(1..10 | ForEach-Object {'FP-{0:D3}' -f $_} | Where-Object {$_ -notin $ids})
            $passTests = @($summary.Tests).Count -eq 10 -and $missing.Count -eq 0 -and [int]$summary.TestCounts.SKIPPED -eq 0 -and [int]$summary.TestCounts.FAIL -eq 0
            Add-Validation -Severity BLOCKER -Scope $scope -Check 'FP-001..FP-010 sin SKIPPED/FAIL' -Pass $passTests -Evidence ("tests={0}; missing={1}; skipped={2}; fail={3}" -f @($summary.Tests).Count,($missing -join ','),$summary.TestCounts.SKIPPED,$summary.TestCounts.FAIL) -Path $summaries[0].FullName
            if($manifest){
                $time = Test-TimestampsInWindow -Rows @($read.Rows) -StartUtc ([datetimeoffset]::Parse([string]$manifest.StartUtc)) -EndUtc ([datetimeoffset]::Parse([string]$manifest.EndUtc))
                Add-Validation -Severity MAJOR -Scope $scope -Check 'Timestamps sin contaminación temporal' -Pass ($time.Unparseable -eq 0 -and $time.Outside -eq 0) -Evidence ("parsed={0}; unparseable={1}; outside={2}" -f $time.Parsed,$time.Unparseable,$time.Outside) -Path $rawPath
                Add-Validation -Severity MAJOR -Scope $scope -Check 'Hash JSONL coincide' -Pass ((Get-CRSha256 -Path $rawPath) -eq [string]$manifest.RawJsonlSHA256) -Evidence (Get-CRSha256 -Path $rawPath) -Path $rawPath
            }
            $rawText = if(Test-Path -LiteralPath $rawPath){Get-Content -LiteralPath $rawPath -Raw -Encoding UTF8}else{''}
            $runIdsInJsonl = @([regex]::Matches($rawText,'TFM_FP_\d{8}_\d{6}') | ForEach-Object {$_.Value} | Sort-Object -Unique)
            $foreignRunIds = @($runIdsInJsonl | Where-Object { $_ -ne [string]$summary.RunId })
            Add-Validation -Severity BLOCKER -Scope $scope -Check 'Sin RunId ajeno' -Pass ($foreignRunIds.Count -eq 0) -Evidence ("own={0}; foreign={1}" -f $summary.RunId,($foreignRunIds -join ',')) -Path $rawPath
        }
        $hitsFiles = @(Get-ChildItem -LiteralPath (Join-Path $repDir.FullName 'runner_native') -Filter 'TFM_FP_*_vr_hits.csv' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)
        $canonicalHits = @()
        if ($hitsFiles.Count -gt 0) { $canonicalHits = @(Import-Csv -LiteralPath $hitsFiles[0].FullName) }
        $p1 = @($canonicalHits | Where-Object {[string](Get-SafePropertyValue -Object $_ -Name 'Profile' -Default '') -eq 'P1_CRITICAL'}).Count
        $p2 = @($canonicalHits | Where-Object {[string](Get-SafePropertyValue -Object $_ -Name 'Profile' -Default '') -eq 'P2_EVENT'}).Count
        $unknown = @($canonicalHits | Where-Object {[string](Get-SafePropertyValue -Object $_ -Name 'FP_ID' -Default 'UNKNOWN') -eq 'UNKNOWN'}).Count
        Add-Validation -Severity BLOCKER -Scope $scope -Check 'P1 crítico igual a cero' -Pass ($p1 -eq 0) -Evidence "P1_CRITICAL=$p1; source=runner vr_hits window" -Path $(if($hitsFiles.Count){$hitsFiles[0].FullName}else{$rawPath})
        Add-Validation -Severity MINOR -Scope $scope -Check 'P2 forense observado/adjudicación requerida' -Pass ($p2 -eq 0) -Evidence "P2_EVENT=$p2; P2 dual-use/forensic does not invalidate execution, but must be classified before conclusions" -Path $(if($hitsFiles.Count){$hitsFiles[0].FullName}else{$rawPath})
        Add-Validation -Severity MAJOR -Scope $scope -Check 'Hits asociados a FP_ID' -Pass ($unknown -eq 0) -Evidence "UNKNOWN=$unknown; source=runner vr_hits window" -Path $(if($hitsFiles.Count){$hitsFiles[0].FullName}else{$rawPath})

        $exports = Get-ExpectedExports -Directory (Join-Path $repDir.FullName 'velociraptor_exports')
        $missingExports = @($exports | Where-Object {-not $_.Exists -or $_.Length -le 0})
        Add-Validation -Severity MAJOR -Scope $scope -Check 'Exports CLIENT_EVENT P1/P2/P3/P4' -Pass ($missingExports.Count -eq 0) -Evidence ("missing={0}" -f ((Get-SafeNameList -Items $missingExports) -join ',')) -Path (Join-Path $repDir.FullName 'velociraptor_exports')
        if($missingExports.Count -eq 0 -and $manifest){
            $exportSet = Read-ClientEventExportSet -Directory (Join-Path $repDir.FullName 'velociraptor_exports')
            Add-Validation -Severity MAJOR -Scope $scope -Check 'Exports CLIENT_EVENT parseables' -Pass (@($exportSet.Errors).Count -eq 0) -Evidence ("rows={0}; errors={1}" -f @($exportSet.Rows).Count,(@($exportSet.Errors) -join '; ')) -Path (Join-Path $repDir.FullName 'velociraptor_exports')
            Add-Validation -Severity MAJOR -Scope $scope -Check 'Exports conservan Artifact y Source' -Pass (@($exportSet.SchemaBad).Count -eq 0) -Evidence ("bad={0}" -f (@($exportSet.SchemaBad) -join ',')) -Path (Join-Path $repDir.FullName 'velociraptor_exports')
            $exportTime = Test-TimestampsInWindow -Rows @($exportSet.Rows) -StartUtc ([datetimeoffset]::Parse([string]$manifest.StartUtc)) -EndUtc ([datetimeoffset]::Parse([string]$manifest.EndUtc)) -ToleranceSeconds 0
            Add-Validation -Severity MAJOR -Scope $scope -Check 'Cronología CLIENT_EVENT limitada a repetición' -Pass ($exportTime.Unparseable -eq 0 -and $exportTime.Outside -eq 0) -Evidence ("parsed={0}; unparseable={1}; outside={2}" -f $exportTime.Parsed,$exportTime.Unparseable,$exportTime.Outside) -Path (Join-Path $repDir.FullName 'velociraptor_exports')
            $filteringReportCsv = Join-Path $repDir.FullName 'velociraptor_exports\CLIENT_EVENT_FILTERING_REPORT.csv'
            $filteringReportMd = Join-Path $repDir.FullName 'velociraptor_exports\CLIENT_EVENT_FILTERING_REPORT.md'
            $reportsExist = (Test-Path -LiteralPath $filteringReportCsv -PathType Leaf) -and (Test-Path -LiteralPath $filteringReportMd -PathType Leaf)
            Add-Validation -Severity MAJOR -Scope $scope -Check 'Informes de filtrado CLIENT_EVENT' -Pass $reportsExist -Evidence ("csv={0}; md={1}" -f (Test-Path -LiteralPath $filteringReportCsv),(Test-Path -LiteralPath $filteringReportMd)) -Path (Join-Path $repDir.FullName 'velociraptor_exports')
            if($reportsExist){
                $filterRows=@(Import-Csv -LiteralPath $filteringReportCsv)
                $reportPass=$filterRows.Count -eq 4 -and @($filterRows|Where-Object {[string]$_.Repetition -ne $repDir.Name -or [string]$_.StartUtc -ne ([datetimeoffset]::Parse([string]$manifest.StartUtc).UtcDateTime.ToString('o')) -or [string]$_.EndUtc -ne ([datetimeoffset]::Parse([string]$manifest.EndUtc).UtcDateTime.ToString('o')) -or [string]$_.ImportCsvVerified -ne 'True'}).Count -eq 0
                Add-Validation -Severity MAJOR -Scope $scope -Check 'Informe usa repetition_manifest e Import-Csv verificado' -Pass $reportPass -Evidence ("rows={0}; repetition={1}" -f $filterRows.Count,$repDir.Name) -Path $filteringReportCsv
            }
        }
    }
    $listenerLog = Join-Path $fpRoot 'fp009_benign_listener.log'
    $listenerText = if(Test-Path -LiteralPath $listenerLog){Get-Content -LiteralPath $listenerLog -Raw -Encoding UTF8}else{''}
    $listener200 = @([regex]::Matches($listenerText,'(?m)REQUEST .* GET / HTTP/1\.1 STATUS=200')).Count
    Add-Validation -Severity BLOCKER -Scope FP -Check 'Listener benigno FP-009 respondió a tres GET' -Pass ($listener200 -eq 3) -Evidence "HTTP_GET_200_count=$listener200" -Path $listenerLog
}

# ---------------------------------------------------------------------------
# Benchmark
# ---------------------------------------------------------------------------
$benchRoot = Join-Path $OutputRoot ("BENCHMARK\{0}" -f $BenchmarkCampaignId)
Add-Validation -Severity BLOCKER -Scope BENCHMARK -Check 'Carpeta campaña' -Pass (Test-Path -LiteralPath $benchRoot -PathType Container) -Evidence $benchRoot -Path $benchRoot
if (Test-Path -LiteralPath $benchRoot -PathType Container) {
    $summaryFiles = @(Get-ChildItem -LiteralPath $benchRoot -Recurse -Filter 'summary.json' -File | Sort-Object FullName)
    Add-Validation -Severity BLOCKER -Scope BENCHMARK -Check 'Nueve summary.json' -Pass ($summaryFiles.Count -eq 9) -Evidence "count=$($summaryFiles.Count)" -Path $benchRoot
    $scenarioCounts = @{}
    foreach($summaryFile in $summaryFiles){
        $relativeSummary = $summaryFile.FullName.Substring($benchRoot.TrimEnd('\').Length).TrimStart('\')
        $pathParts = @($relativeSummary -split '\\')
        $repLabel = if($pathParts.Count -ge 1){[string]$pathParts[0]}else{'REP_UNKNOWN'}
        $scenarioLabel = if($pathParts.Count -ge 2){[string]$pathParts[1]}else{'SCENARIO_UNKNOWN'}
        $scope = "BENCHMARK/$repLabel/$scenarioLabel"
        try {
            $summary = $null
            try{$summary=Get-Content -LiteralPath $summaryFile.FullName -Raw -Encoding UTF8 | ConvertFrom-Json}catch{}
            Add-Validation -Severity BLOCKER -Scope $scope -Check 'summary.json parseable' -Pass ($null -ne $summary) -Evidence $summaryFile.FullName -Path $summaryFile.FullName
            if(-not $summary){continue}

            $scenario = [string](Get-SafePropertyValue -Object $summary -Name 'Scenario' -Default $scenarioLabel)
            if(-not $scenarioCounts.ContainsKey($scenario)){$scenarioCounts[$scenario]=0}
            $scenarioCounts[$scenario]++
            $schemaVersion = [string](Get-SafePropertyValue -Object $summary -Name 'SchemaVersion' -Default '')
            $scenarioValidity = [string](Get-SafePropertyValue -Object $summary -Name 'ScenarioValidity' -Default 'MISSING')
            $validityReasons = @(Get-SafePropertyValue -Object $summary -Name 'ValidityReasons' -Default @())
            Add-Validation -Severity BLOCKER -Scope $scope -Check 'SchemaVersion 1.1' -Pass ($schemaVersion -eq '1.1') -Evidence $schemaVersion -Path $summaryFile.FullName
            Add-Validation -Severity BLOCKER -Scope $scope -Check 'ScenarioValidity VALID' -Pass ($scenarioValidity -eq 'VALID') -Evidence ($validityReasons -join '; ') -Path $summaryFile.FullName

            $clientMax = [int](Get-SafePropertyValue -Object $summary -Name 'VR_Client_ProcessCountMax' -Default -1)
            $serviceStart = [string](Get-SafePropertyValue -Object $summary -Name 'ClientServiceStatusStart' -Default 'MISSING')
            $servicePath = [string](Get-SafePropertyValue -Object $summary -Name 'ClientServicePath' -Default '')
            $clientPass = $false
            if($scenario -eq 'BASELINE_NO_VR'){$clientPass=($clientMax -eq 0 -and $serviceStart -eq 'Stopped')}
            if($scenario -in @('VR_IDLE','VR_TEC_RUNNER')){$clientPass=($clientMax -gt 0 -and $serviceStart -eq 'Running')}
            Add-Validation -Severity BLOCKER -Scope $scope -Check 'CLIENT_SERVICE coherente' -Pass $clientPass -Evidence ("scenario={0}; max={1}; serviceStart={2}; path={3}" -f $scenario,$clientMax,$serviceStart,$servicePath) -Path $summaryFile.FullName
            $serviceShape = ($servicePath -match '(?i)client\.config\.yaml') -and ($servicePath -match '(?i)service\s+run')
            Add-Validation -Severity MAJOR -Scope $scope -Check 'Ruta identifica cliente Velociraptor' -Pass $serviceShape -Evidence $servicePath -Path $summaryFile.FullName

            $runnerObserved = [int](Get-SafePropertyValue -Object $summary -Name 'RunnerObservedSamples' -Default 0)
            if($scenario -eq 'VR_TEC_RUNNER'){
                Add-Validation -Severity BLOCKER -Scope $scope -Check 'Runner observado' -Pass ($runnerObserved -gt 0) -Evidence "RunnerObservedSamples=$runnerObserved" -Path $summaryFile.FullName
                $runnerStillRunning = [bool](Get-SafePropertyValue -Object $summary -Name 'RunnerStillRunningAtEnd' -Default $true)
                Add-Validation -Severity MINOR -Scope $scope -Check 'Estado del runner al cerrar ventana' -Pass (-not $runnerStillRunning) -Evidence "RunnerStillRunningAtEnd=$runnerStillRunning; no invalida la muestra si RunnerObservedSamples>0 y ScenarioValidity=VALID" -Path $summaryFile.FullName
            }
            $serverExcludedValue = [bool](Get-SafePropertyValue -Object $summary -Name 'ServerGuiExcludedFromClientMetrics' -Default $false)
            $includeServerValue = [bool](Get-SafePropertyValue -Object $summary -Name 'IncludeServerGuiInTotal' -Default $true)
            $serverExcluded = $serverExcludedValue -and (-not $includeServerValue)
            Add-Validation -Severity BLOCKER -Scope $scope -Check 'SERVER_GUI excluido' -Pass $serverExcluded -Evidence ("Excluded={0}; IncludeInTotal={1}" -f $serverExcludedValue,$includeServerValue) -Path $summaryFile.FullName

            $required = Test-RequiredFiles -RunDirectory $summaryFile.Directory.FullName -Names @('samples.csv','process_samples.csv','summary.json','summary.txt','process_list_start.txt','process_list_end.txt','environment.txt')
            $missing = @($required | Where-Object {-not $_.Exists})
            Add-Validation -Severity BLOCKER -Scope $scope -Check 'Siete ficheros benchmark' -Pass ($missing.Count -eq 0) -Evidence ("missing={0}" -f ((Get-SafeNameList -Items $missing) -join ',')) -Path $summaryFile.Directory.FullName

            $processCsv = Join-Path $summaryFile.Directory.FullName 'process_samples.csv'
            $processRows = if(Test-Path -LiteralPath $processCsv){@(Import-Csv -LiteralPath $processCsv)}else{@()}
            $notepadRows = @($processRows | Where-Object {
                $processName = [string](Get-SafePropertyValue -Object $_ -Name 'ProcessName' -Default '')
                $role = [string](Get-SafePropertyValue -Object $_ -Name 'Role' -Default '')
                $processName -match '^(?i:notepad)(\.exe)?$' -or $role -match '(?i)notepad'
            })
            Add-Validation -Severity BLOCKER -Scope $scope -Check 'notepad.exe no clasificado como runner/VR' -Pass ($notepadRows.Count -eq 0) -Evidence "notepadRows=$($notepadRows.Count)" -Path $processCsv
            $clientRows = @($processRows | Where-Object {[string](Get-SafePropertyValue -Object $_ -Name 'Role' -Default '') -eq 'CLIENT_SERVICE'})
            $clientRowsPass = if($scenario -eq 'BASELINE_NO_VR'){$clientRows.Count -eq 0}else{$clientRows.Count -gt 0}
            Add-Validation -Severity BLOCKER -Scope $scope -Check 'process_samples CLIENT_SERVICE coherente' -Pass $clientRowsPass -Evidence "CLIENT_SERVICE_rows=$($clientRows.Count)" -Path $processCsv

            $samplesCsv = Join-Path $summaryFile.Directory.FullName 'samples.csv'
            $sampleRows = if(Test-Path -LiteralPath $samplesCsv){@(Import-Csv -LiteralPath $samplesCsv)}else{@()}
            $summarySamples = [int](Get-SafePropertyValue -Object $summary -Name 'Samples' -Default -1)
            $sampleCountPass = $sampleRows.Count -eq $summarySamples -and $sampleRows.Count -gt 0
            Add-Validation -Severity MAJOR -Scope $scope -Check 'Muestras CSV coinciden con summary' -Pass $sampleCountPass -Evidence ("csv={0}; summary={1}" -f $sampleRows.Count,$summarySamples) -Path $samplesCsv
        }
        catch {
            Add-Validation -Severity BLOCKER -Scope $scope -Check 'Excepción interna de validación controlada' -Pass $false -Evidence ("{0}: {1}" -f $_.Exception.GetType().FullName,$_.Exception.Message) -Path $summaryFile.FullName
            continue
        }
    }
    foreach($scenario in @('BASELINE_NO_VR','VR_IDLE','VR_TEC_RUNNER')){
        $count = if($scenarioCounts.ContainsKey($scenario)){$scenarioCounts[$scenario]}else{0}
        Add-Validation -Severity BLOCKER -Scope BENCHMARK -Check ("Tres repeticiones $scenario") -Pass ($count -eq 3) -Evidence "count=$count" -Path $benchRoot
    }
}

$resultArray = $results.ToArray()
$blockingFailures = @($resultArray | Where-Object {-not $_.Pass -and $_.Severity -in @('BLOCKER','MAJOR')})
$overall = if($blockingFailures.Count -eq 0){'PASS'}else{'FAIL'}
$resultArray | Export-Csv -LiteralPath (Join-Path $validationRoot 'VALIDATION_RESULTS.csv') -NoTypeInformation -Encoding UTF8
Write-CRJson -Object ([pscustomobject]@{
    ValidationId=$validationId
    Timestamp=(Get-Date).ToString('o')
    TecCampaignId=$TecCampaignId
    FpCampaignId=$FpCampaignId
    BenchmarkCampaignId=$BenchmarkCampaignId
    OverallStatus=$overall
    TotalChecks=$resultArray.Count
    BlockingFailures=$blockingFailures.Count
    Results=$resultArray
}) -Path (Join-Path $validationRoot 'VALIDATION_REPORT.json')

$md = New-Object System.Collections.Generic.List[string]
[void]$md.Add('# Validación de repetición controlada')
[void]$md.Add('')
[void]$md.Add(("- Estado: **{0}**" -f $overall))
[void]$md.Add(('- TEC: `{0}`' -f $TecCampaignId))
[void]$md.Add(('- FP: `{0}`' -f $FpCampaignId))
[void]$md.Add(('- Benchmark: `{0}`' -f $BenchmarkCampaignId))
[void]$md.Add(("- Checks: {0}" -f $resultArray.Count))
[void]$md.Add(("- Fallos BLOCKER/MAJOR: {0}" -f $blockingFailures.Count))
[void]$md.Add('')
[void]$md.Add('| Severidad | Ambito | Check | Resultado | Evidencia |')
[void]$md.Add('|---|---|---|---|---|')
foreach($item in $resultArray){
    $ev = ([string]$item.Evidence).Replace('|','/').Replace("`r",' ').Replace("`n",' ')
    [void]$md.Add(("| {0} | {1} | {2} | {3} | {4} |" -f $item.Severity,$item.Scope,$item.Check,$(if($item.Pass){'PASS'}else{'FAIL'}),$ev))
}
Write-CRUtf8NoBom -Path (Join-Path $validationRoot 'VALIDATION_REPORT.md') -Text ($md -join [Environment]::NewLine)
Export-CRHashManifest -Root $validationRoot -OutputPath (Join-Path $validationRoot 'HASHES_SHA256.csv')

if($overall -ne 'PASS'){
    Write-CRStatus -Message "VALIDACION FAIL: $($blockingFailures.Count) fallos. Informe: $validationRoot" -Level FAIL
    exit 2
}
Write-CRStatus -Message "VALIDACION PASS. Los datos pueden usarse para reconstruir CSV/Excel. Informe: $validationRoot" -Level OK
