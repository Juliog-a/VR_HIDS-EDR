<#
Collect final local evidence after running TFM_Run_All_TEC_Tests_v5.ps1.

This script is read-only over telemetry and campaign outputs. It does not run
techniques and does not modify receiver, router, artifacts or attack scripts.
#>

param(
    [string]$ProjectRoot = "${PSScriptRoot}\..",
    [string]$CampaignLogDir = "C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas\Logs_Pruebas_TFM",
    [string]$Tec009Dir = "C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas\TEC-009_Exfiltracion",
    [string]$ReceiverIP = "192.168.1.129",
    [int]$ReceiverPort = 8088,
    [int]$LookbackHours = 4
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outRoot = Join-Path $ProjectRoot ("05_LOGS\FINAL_RUN_{0}" -f $stamp)
New-Item -ItemType Directory -Force -Path $outRoot | Out-Null

$summaryDir = Join-Path $outRoot "campaign_summaries"
$tec009OutDir = Join-Path $outRoot "tec009"
$checksDir = Join-Path $outRoot "telemetry_checks"
New-Item -ItemType Directory -Force -Path $summaryDir | Out-Null
New-Item -ItemType Directory -Force -Path $tec009OutDir | Out-Null
New-Item -ItemType Directory -Force -Path $checksDir | Out-Null

function Write-Text {
    param([string]$Path, [string[]]$Lines)
    $Lines | Set-Content -LiteralPath $Path -Encoding UTF8
}

function Copy-Latest {
    param(
        [string]$Directory,
        [string]$Pattern,
        [string]$DestinationDirectory
    )
    if (-not (Test-Path -LiteralPath $Directory)) { return $null }
    $file = Get-ChildItem -LiteralPath $Directory -Filter $Pattern -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if ($null -eq $file) { return $null }
    $dest = Join-Path $DestinationDirectory $file.Name
    Copy-Item -LiteralPath $file.FullName -Destination $dest -Force
    return $dest
}

function Get-JsonSafe {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    try { return (Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json) }
    catch { return $null }
}

function Status-FromBool {
    param([bool]$Value, [string]$WarnInsteadOfFail = "")
    if ($Value) { return "OK" }
    if ($WarnInsteadOfFail) { return "WARN" }
    return "FAIL"
}

$latestTxt = Copy-Latest -Directory $CampaignLogDir -Pattern "*v5*_summary.txt" -DestinationDirectory $summaryDir
$latestJson = Copy-Latest -Directory $CampaignLogDir -Pattern "*v5*_summary.json" -DestinationDirectory $summaryDir
$latestCsv = Copy-Latest -Directory $CampaignLogDir -Pattern "*v5*_summary.csv" -DestinationDirectory $summaryDir

if ($null -eq $latestJson) {
    $latestTxt = Copy-Latest -Directory $CampaignLogDir -Pattern "*summary.txt" -DestinationDirectory $summaryDir
    $latestJson = Copy-Latest -Directory $CampaignLogDir -Pattern "*summary.json" -DestinationDirectory $summaryDir
    $latestCsv = Copy-Latest -Directory $CampaignLogDir -Pattern "*summary.csv" -DestinationDirectory $summaryDir
}

$tec009SummarySource = $null
if (Test-Path -LiteralPath $Tec009Dir) {
    $tec009SummarySource = Get-ChildItem -LiteralPath $Tec009Dir -Recurse -Filter "tec009_transfer_summary.json" -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
}

$tec009SummaryCopy = $null
$tec009Summary = $null
if ($null -ne $tec009SummarySource) {
    $tec009SummaryCopy = Join-Path $tec009OutDir "tec009_transfer_summary.json"
    Copy-Item -LiteralPath $tec009SummarySource.FullName -Destination $tec009SummaryCopy -Force
    $tec009Summary = Get-JsonSafe -Path $tec009SummaryCopy
}

$zipPath = ""
$copyPath = ""
$shaFromSummary = ""
$httpStatus = ""
$uploadSucceeded = ""
$receiverUrl = ""
if ($null -ne $tec009Summary) {
    if ($tec009Summary.PSObject.Properties.Name -contains "ZipFile") { $zipPath = [string]$tec009Summary.ZipFile }
    if ($tec009Summary.PSObject.Properties.Name -contains "CopyFile") { $copyPath = [string]$tec009Summary.CopyFile }
    if ($tec009Summary.PSObject.Properties.Name -contains "LocalSHA256") { $shaFromSummary = [string]$tec009Summary.LocalSHA256 }
    if ($tec009Summary.PSObject.Properties.Name -contains "HttpStatusCode") { $httpStatus = [string]$tec009Summary.HttpStatusCode }
    if ($tec009Summary.PSObject.Properties.Name -contains "UploadSucceeded") { $uploadSucceeded = [string]$tec009Summary.UploadSucceeded }
    if ($tec009Summary.PSObject.Properties.Name -contains "ReceiverUrl") { $receiverUrl = [string]$tec009Summary.ReceiverUrl }
}

$zipHash = ""
if (-not [string]::IsNullOrWhiteSpace($zipPath) -and (Test-Path -LiteralPath $zipPath)) {
    $zipCopyDest = Join-Path $tec009OutDir (Split-Path -Leaf $zipPath)
    Copy-Item -LiteralPath $zipPath -Destination $zipCopyDest -Force
    $zipHash = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash
}

if (-not [string]::IsNullOrWhiteSpace($copyPath) -and (Test-Path -LiteralPath $copyPath)) {
    Copy-Item -LiteralPath $copyPath -Destination (Join-Path $tec009OutDir (Split-Path -Leaf $copyPath)) -Force
}

$receiverCheck = [ordered]@{
    ReceiverIP = $ReceiverIP
    ReceiverPort = $ReceiverPort
    Reachable = $false
    Error = ""
}
try {
    $receiverCheck.Reachable = [bool](Test-NetConnection -ComputerName $ReceiverIP -Port $ReceiverPort -InformationLevel Quiet -WarningAction SilentlyContinue)
}
catch {
    $receiverCheck.Error = $_.Exception.Message
}
$receiverCheck | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $checksDir "receiver_connectivity.json") -Encoding UTF8

$startTime = (Get-Date).AddHours(-1 * $LookbackHours)

$ps4104Rows = @()
try {
    $psEvents = Get-WinEvent -FilterHashtable @{
        LogName = 'Microsoft-Windows-PowerShell/Operational'
        Id = 4104
        StartTime = $startTime
    } -ErrorAction Stop
    $ps4104Rows = @($psEvents | Where-Object {
        $m = [string]$_.Message
        $m -match 'Invoke-WebRequest|iwr|Invoke-RestMethod|WebClient|UploadFile|-InFile|application/zip|ReceiverUrl|\.zip|upload'
    } | Select-Object TimeCreated, Id, ProviderName, Message)
}
catch {
    Write-Text -Path (Join-Path $checksDir "powershell_4104_error.txt") -Lines @($_.Exception.Message)
}
$ps4104Rows | Export-Csv -LiteralPath (Join-Path $checksDir "powershell_4104_tec009_candidates.csv") -NoTypeInformation -Encoding UTF8

$sysmonRows = @()
try {
    $sysEvents = Get-WinEvent -FilterHashtable @{
        LogName = 'Microsoft-Windows-Sysmon/Operational'
        Id = 1,3,11
        StartTime = $startTime
    } -ErrorAction Stop
    $sysmonRows = @($sysEvents | Where-Object {
        $m = [string]$_.Message
        $m -match 'powershell|pwsh|ReceiverUrl|EnableUpload|Invoke-WebRequest|InFile|application/zip|\.zip|upload|8088'
    } | Select-Object TimeCreated, Id, ProviderName, Message)
}
catch {
    Write-Text -Path (Join-Path $checksDir "sysmon_id1_3_11_error.txt") -Lines @($_.Exception.Message)
}
$sysmonRows | Export-Csv -LiteralPath (Join-Path $checksDir "sysmon_id1_3_11_tec009_candidates.csv") -NoTypeInformation -Encoding UTF8

$hashLines = New-Object System.Collections.Generic.List[string]
Get-ChildItem -LiteralPath $outRoot -Recurse -File | ForEach-Object {
    $hash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
    $rel = $_.FullName.Substring($outRoot.Length).TrimStart('\')
    $hashLines.Add(("{0}  {1}" -f $hash, $rel)) | Out-Null
}
Write-Text -Path (Join-Path $outRoot "SHA256SUMS.txt") -Lines $hashLines

$tec009ZipOk = (-not [string]::IsNullOrWhiteSpace($zipPath)) -and (Test-Path -LiteralPath $zipPath)
$tec009CopyOk = (-not [string]::IsNullOrWhiteSpace($copyPath)) -and (Test-Path -LiteralPath $copyPath)
$tec009SummaryOk = $null -ne $tec009Summary
$tec009HashOk = $shaFromSummary -match '^[0-9a-fA-F]{64}$'
$tec009HttpOk = $false
[int]$httpInt = 0
if ([int]::TryParse($httpStatus, [ref]$httpInt)) { $tec009HttpOk = ($httpInt -ge 200 -and $httpInt -lt 300) }
$tec009UploadOk = $uploadSucceeded -match '(?i)^true$'
$ps4104Ok = @($ps4104Rows).Count -gt 0
$sysmonOk = @($sysmonRows).Count -gt 0
$tec009EvidenceStatus = if ($tec009ZipOk -and $tec009CopyOk -and $tec009SummaryOk -and $tec009HashOk -and $tec009HttpOk -and $tec009UploadOk -and $ps4104Ok) {
    "OK"
} elseif ($tec009ZipOk -and $tec009UploadOk) {
    "WARN"
} else {
    "FAIL"
}

$evidenceRows = @(
    [pscustomobject]@{ Tecnica="TEC-001"; EvidenciaLocal="Runner summary TXT/JSON/CSV"; EvidenciaVelociraptorEsperada="P3 TEC001_PowerShell_4104"; Estado=(Status-FromBool -Value ($latestJson -ne $null)) },
    [pscustomobject]@{ Tecnica="TEC-002"; EvidenciaLocal="Runner summary TXT/JSON/CSV"; EvidenciaVelociraptorEsperada="P3 TEC002_CMD_Sysmon"; Estado=(Status-FromBool -Value ($latestJson -ne $null)) },
    [pscustomobject]@{ Tecnica="TEC-003"; EvidenciaLocal="Runner summary TXT/JSON/CSV"; EvidenciaVelociraptorEsperada="P3 TEC003_ScheduledTask_Sysmon"; Estado=(Status-FromBool -Value ($latestJson -ne $null)) },
    [pscustomobject]@{ Tecnica="TEC-004"; EvidenciaLocal="Runner summary TXT/JSON/CSV"; EvidenciaVelociraptorEsperada="P3 TEC004_RunKey_Sysmon"; Estado=(Status-FromBool -Value ($latestJson -ne $null)) },
    [pscustomobject]@{ Tecnica="TEC-005"; EvidenciaLocal="Runner summary TXT/JSON/CSV"; EvidenciaVelociraptorEsperada="P3 TEC005_Service_System7045"; Estado=(Status-FromBool -Value ($latestJson -ne $null)) },
    [pscustomobject]@{ Tecnica="TEC-006"; EvidenciaLocal="Runner summary + 4104/Sysmon discovery"; EvidenciaVelociraptorEsperada="P3 TEC006_SecuritySoftwareDiscovery_4104"; Estado=(Status-FromBool -Value ($latestJson -ne $null)) },
    [pscustomobject]@{ Tecnica="TEC-007"; EvidenciaLocal="Runner summary .aesCount + 4104 crypto"; EvidenciaVelociraptorEsperada="P3 TEC007_DataEncryptedForImpact_4104"; Estado=(Status-FromBool -Value ($latestJson -ne $null)) },
    [pscustomobject]@{ Tecnica="TEC-008"; EvidenciaLocal="Runner summary FilesBefore/FilesAfter + 4104 deletion"; EvidenciaVelociraptorEsperada="P3 TEC008_DataDestruction_4104"; Estado=(Status-FromBool -Value ($latestJson -ne $null)) },
    [pscustomobject]@{ Tecnica="TEC-009"; EvidenciaLocal=("ZIP={0}; Copy={1}; SHA256={2}; HTTP={3}; Upload={4}; 4104Rows={5}; SysmonRows={6}" -f $tec009ZipOk,$tec009CopyOk,$shaFromSummary,$httpStatus,$uploadSucceeded,@($ps4104Rows).Count,@($sysmonRows).Count); EvidenciaVelociraptorEsperada="P3 TEC009_HTTP_ZIP_Upload_4104 HIGH + Sysmon context"; Estado=$tec009EvidenceStatus }
)

$md = New-Object System.Collections.Generic.List[string]
$md.Add("# RESUMEN_EVIDENCIA_FINAL") | Out-Null
$md.Add("") | Out-Null
$md.Add(("Output: {0}" -f $outRoot)) | Out-Null
$md.Add(("Generated: {0}" -f (Get-Date).ToString("o"))) | Out-Null
$md.Add(("CampaignLogDir: {0}" -f $CampaignLogDir)) | Out-Null
$md.Add(("Tec009Dir: {0}" -f $Tec009Dir)) | Out-Null
$md.Add("") | Out-Null
$md.Add("## TEC-009") | Out-Null
$md.Add("") | Out-Null
$md.Add(("- Summary JSON: {0}" -f $tec009SummaryCopy)) | Out-Null
$md.Add(("- ZipFile: {0}" -f $zipPath)) | Out-Null
$md.Add(("- CopyFile: {0}" -f $copyPath)) | Out-Null
$md.Add(("- SHA256 summary: {0}" -f $shaFromSummary)) | Out-Null
$md.Add(("- SHA256 recalculated: {0}" -f $zipHash)) | Out-Null
$md.Add(("- ReceiverUrl: {0}" -f $receiverUrl)) | Out-Null
$md.Add(("- HTTP status: {0}" -f $httpStatus)) | Out-Null
$md.Add(("- UploadSucceeded: {0}" -f $uploadSucceeded)) | Out-Null
$md.Add(("- 4104 candidate rows: {0}" -f @($ps4104Rows).Count)) | Out-Null
$md.Add(("- Sysmon ID1/3/11 candidate rows: {0}" -f @($sysmonRows).Count)) | Out-Null
$md.Add("") | Out-Null
$md.Add("## Matriz") | Out-Null
$md.Add("") | Out-Null
$md.Add("| Tecnica | Evidencia local | Evidencia Velociraptor esperada | Estado |") | Out-Null
$md.Add("|---|---|---|---|") | Out-Null
foreach ($row in $evidenceRows) {
    $md.Add(("| {0} | {1} | {2} | {3} |" -f $row.Tecnica, $row.EvidenciaLocal, $row.EvidenciaVelociraptorEsperada, $row.Estado)) | Out-Null
}
$md.Add("") | Out-Null
$md.Add("## Archivos generados") | Out-Null
$md.Add("") | Out-Null
$md.Add("- campaign_summaries/") | Out-Null
$md.Add("- tec009/") | Out-Null
$md.Add("- telemetry_checks/powershell_4104_tec009_candidates.csv") | Out-Null
$md.Add("- telemetry_checks/sysmon_id1_3_11_tec009_candidates.csv") | Out-Null
$md.Add("- SHA256SUMS.txt") | Out-Null

Write-Text -Path (Join-Path $outRoot "RESUMEN_EVIDENCIA_FINAL.md") -Lines $md

Write-Host ("Evidence output: {0}" -f $outRoot)
