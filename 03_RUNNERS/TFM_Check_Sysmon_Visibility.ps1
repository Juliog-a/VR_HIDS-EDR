<#
TFM_Check_Sysmon_Visibility.ps1

Read-only visibility checker for the TFM Velociraptor/Sysmon lab.

Purpose:
- Do not execute attacks.
- Do not modify the lab.
- Query Windows event logs and summarize whether telemetry exists for TEC-001 to TEC-009.
- Generate log, TXT summary, JSON summary and CSV event samples.

Usage:
  cd C:\Users\seguridad\Desktop\TFM\runner_candidate
  Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
  .\TFM_Check_Sysmon_Visibility.ps1 `
    -LookbackHours 24 `
    -IncludeSamples `
    -MaxSamplesPerQuery 10
#>

param(
    [int]$LookbackHours = 24,
    [string]$OutputDir = "C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas\Logs_Pruebas_TFM",
    [switch]$IncludeSamples,
    [int]$MaxSamplesPerQuery = 10
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$RunStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile = Join-Path $OutputDir ("TFM_Sysmon_Visibility_{0}.log" -f $RunStamp)
$SummaryTxt = Join-Path $OutputDir ("TFM_Sysmon_Visibility_{0}_summary.txt" -f $RunStamp)
$SummaryJson = Join-Path $OutputDir ("TFM_Sysmon_Visibility_{0}_summary.json" -f $RunStamp)
$EventsCsv = Join-Path $OutputDir ("TFM_Sysmon_Visibility_{0}_events.csv" -f $RunStamp)

$StartTime = (Get-Date).AddHours(-1 * $LookbackHours)
$EndTime = Get-Date
$Rows = New-Object System.Collections.Generic.List[object]
$Samples = New-Object System.Collections.Generic.List[object]

function Write-Log {
    param([AllowEmptyString()][string]$Message = "", [string]$Level = "INFO")
    if ([string]::IsNullOrWhiteSpace($Message)) {
        Write-Host ""
        Add-Content -Path $LogFile -Value ""
        return
    }
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Write-Host $line
    Add-Content -Path $LogFile -Value $line
}

function Test-LogExists {
    param([Parameter(Mandatory=$true)][string]$LogName)
    try {
        $null = Get-WinEvent -ListLog $LogName -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

function Convert-EventMessage {
    param([AllowEmptyString()][string]$Message = "")
    if ([string]::IsNullOrWhiteSpace($Message)) { return "" }
    return (($Message -replace "`r?`n", " | ").Trim())
}

function Add-VisibilityQuery {
    param(
        [Parameter(Mandatory=$true)][string]$Source,
        [Parameter(Mandatory=$true)][string]$LogName,
        [Parameter(Mandatory=$true)][int[]]$EventIds,
        [Parameter(Mandatory=$true)][string]$Description,
        [Parameter(Mandatory=$true)][string]$Relevance,
        [string]$MessageRegex = "",
        [string]$Technique = ""
    )

    Write-Log ("Query: {0} / {1} / IDs {2}" -f $Source, $Description, ($EventIds -join ","))

    if (-not (Test-LogExists -LogName $LogName)) {
        $Rows.Add([pscustomobject]@{
            Source = $Source
            LogName = $LogName
            EventID = ($EventIds -join ",")
            Description = $Description
            Count = 0
            Status = "MISSING"
            Relevance = $Relevance
            Technique = $Technique
            Comment = "Log not available."
        }) | Out-Null
        return
    }

    try {
        $filter = @{
            LogName = $LogName
            Id = $EventIds
            StartTime = $StartTime
            EndTime = $EndTime
        }
        $events = @(Get-WinEvent -FilterHashtable $filter -ErrorAction Stop)
        if (-not [string]::IsNullOrWhiteSpace($MessageRegex)) {
            $events = @($events | Where-Object { $_.Message -match $MessageRegex })
        }
        $count = $events.Count
        $status = if ($count -gt 0) { "OK" } else { "WARN" }
        $comment = if ($count -gt 0) { "Telemetry observed in lookback window." } else { "No matching events in lookback window." }

        $Rows.Add([pscustomobject]@{
            Source = $Source
            LogName = $LogName
            EventID = ($EventIds -join ",")
            Description = $Description
            Count = $count
            Status = $status
            Relevance = $Relevance
            Technique = $Technique
            Comment = $comment
        }) | Out-Null

        if ($IncludeSamples -and $count -gt 0) {
            foreach ($event in ($events | Sort-Object TimeCreated -Descending | Select-Object -First $MaxSamplesPerQuery)) {
                $Samples.Add([pscustomobject]@{
                    Source = $Source
                    Technique = $Technique
                    LogName = $LogName
                    EventID = $event.Id
                    TimeCreated = $event.TimeCreated.ToString("o")
                    ProviderName = $event.ProviderName
                    RecordId = $event.RecordId
                    Message = Convert-EventMessage -Message $event.Message
                }) | Out-Null
            }
        }
    }
    catch {
        $Rows.Add([pscustomobject]@{
            Source = $Source
            LogName = $LogName
            EventID = ($EventIds -join ",")
            Description = $Description
            Count = 0
            Status = "WARN"
            Relevance = $Relevance
            Technique = $Technique
            Comment = "Query failed: $($_.Exception.Message)"
        }) | Out-Null
    }
}

Write-Log "Starting TFM Sysmon/Windows visibility check"
Write-Log ("LookbackHours: {0}" -f $LookbackHours)
Write-Log ("WindowStart: {0}" -f $StartTime.ToString("o"))
Write-Log ("WindowEnd: {0}" -f $EndTime.ToString("o"))
Write-Log ("OutputDir: {0}" -f $OutputDir)
Write-Log ("IncludeSamples: {0}" -f $IncludeSamples)
Write-Log ("MaxSamplesPerQuery: {0}" -f $MaxSamplesPerQuery)

# Sysmon process creation and execution coverage.
Add-VisibilityQuery -Source "Sysmon" -LogName "Microsoft-Windows-Sysmon/Operational" -EventIds @(1) -Description "Process Create - PowerShell/cmd/script interpreters" -Relevance "TEC-001 TEC-002 CU execution visibility" -Technique "TEC-001/TEC-002" -MessageRegex '(?i)(powershell|pwsh|cmd\.exe|wscript|cscript|rundll32|regsvr32|mshta)'
Add-VisibilityQuery -Source "Sysmon" -LogName "Microsoft-Windows-Sysmon/Operational" -EventIds @(1) -Description "Process Create - scheduled task/service/registry tooling" -Relevance "TEC-003 TEC-004 TEC-005" -Technique "TEC-003/TEC-004/TEC-005" -MessageRegex '(?i)(schtasks|sc\.exe|reg\.exe)'
Add-VisibilityQuery -Source "Sysmon" -LogName "Microsoft-Windows-Sysmon/Operational" -EventIds @(1) -Description "Process Create - impact/destructive tooling" -Relevance "TEC-007 TEC-008 forensic support" -Technique "TEC-007/TEC-008" -MessageRegex '(?i)(vssadmin|wbadmin|bcdedit|wevtutil|powershell|pwsh|Remove-Item|CryptoStream|CreateEncryptor|\.aes)'
Add-VisibilityQuery -Source "Sysmon" -LogName "Microsoft-Windows-Sysmon/Operational" -EventIds @(1) -Description "Process Create - staging/upload/archive indicators" -Relevance "TEC-009 process command visibility" -Technique "TEC-009" -MessageRegex '(?i)(Compress-Archive|Invoke-WebRequest|Invoke-RestMethod|-InFile|application/zip|\.zip|curl|wget|http)'

# Sysmon network and file telemetry.
Add-VisibilityQuery -Source "Sysmon" -LogName "Microsoft-Windows-Sysmon/Operational" -EventIds @(3) -Description "Network Connection - outbound" -Relevance "Network visibility for exfiltration context" -Technique "TEC-009" -MessageRegex '(?i)(Initiated:\s*true|Destination)'
Add-VisibilityQuery -Source "Sysmon" -LogName "Microsoft-Windows-Sysmon/Operational" -EventIds @(3) -Description "Network Connection - receiver 192.168.1.129:8088 / PowerShell HTTP" -Relevance "Controlled HTTP receiver visibility" -Technique "TEC-009" -MessageRegex '(?i)(192\.168\.1\.129|8088|powershell|pwsh|http)'
Add-VisibilityQuery -Source "Sysmon" -LogName "Microsoft-Windows-Sysmon/Operational" -EventIds @(7) -Description "Image Loaded - optional" -Relevance "Optional DLL/image telemetry; high volume if enabled" -Technique "General"
Add-VisibilityQuery -Source "Sysmon" -LogName "Microsoft-Windows-Sysmon/Operational" -EventIds @(10) -Description "Process Access - optional" -Relevance "Optional process access telemetry" -Technique "General"
Add-VisibilityQuery -Source "Sysmon" -LogName "Microsoft-Windows-Sysmon/Operational" -EventIds @(11) -Description "File Create - ZIP/AES/lab staging" -Relevance "TEC-007 .aes and TEC-009 .zip file evidence" -Technique "TEC-007/TEC-009" -MessageRegex '(?i)(\.zip|\.aes|DUMB_LAB|staging|TEC-009)'
Add-VisibilityQuery -Source "Sysmon" -LogName "Microsoft-Windows-Sysmon/Operational" -EventIds @(12,13,14) -Description "Registry events - Run Keys" -Relevance "TEC-004 persistence visibility" -Technique "TEC-004" -MessageRegex '(?i)(CurrentVersion\\Run|RunOnce|TFM_T1547_001_test)'
Add-VisibilityQuery -Source "Sysmon" -LogName "Microsoft-Windows-Sysmon/Operational" -EventIds @(15) -Description "FileCreateStreamHash - ADS optional" -Relevance "Optional alternate data stream visibility" -Technique "General"
Add-VisibilityQuery -Source "Sysmon" -LogName "Microsoft-Windows-Sysmon/Operational" -EventIds @(22) -Description "DNS Query" -Relevance "DNS visibility; optional for IP-based TEC-009" -Technique "TEC-009"
Add-VisibilityQuery -Source "Sysmon" -LogName "Microsoft-Windows-Sysmon/Operational" -EventIds @(23,26) -Description "File Delete - forensic evidence" -Relevance "TEC-008 sabotage forensic support; not individual alert" -Technique "TEC-008"

# PowerShell logging.
Add-VisibilityQuery -Source "PowerShell" -LogName "Microsoft-Windows-PowerShell/Operational" -EventIds @(4104) -Description "ScriptBlock logging - execution/encoding" -Relevance "CU-001 PowerShell behavior" -Technique "TEC-001" -MessageRegex '(?i)(ExecutionPolicy|FromBase64String|IEX|Invoke-Expression|DownloadString)'
Add-VisibilityQuery -Source "PowerShell" -LogName "Microsoft-Windows-PowerShell/Operational" -EventIds @(4104) -Description "ScriptBlock logging - controlled exfiltration shape" -Relevance "CU-009 strong shape: HTTP tool + POST + InFile + application/zip" -Technique "TEC-009" -MessageRegex '(?i)(Invoke-WebRequest|Invoke-RestMethod|Compress-Archive|-InFile|Method\s+POST|application/zip|exfiltracion_v3)'
Add-VisibilityQuery -Source "PowerShell" -LogName "Microsoft-Windows-PowerShell/Operational" -EventIds @(4104) -Description "ScriptBlock logging - crypto/destruction" -Relevance "TEC-007/TEC-008 strong PowerShell evidence" -Technique "TEC-007/TEC-008" -MessageRegex '(?i)(System\.Security\.Cryptography|CryptoStream|CreateEncryptor|\.aes|Get-ChildItem|Remove-Item|ForEach-Object)'
Add-VisibilityQuery -Source "PowerShell" -LogName "Microsoft-Windows-PowerShell/Operational" -EventIds @(4103) -Description "Module logging - optional command detail" -Relevance "Optional PowerShell module logging" -Technique "General" -MessageRegex '(?i)(Invoke-WebRequest|Compress-Archive|Remove-Item|Get-CimInstance|SecurityCenter2)'
Add-VisibilityQuery -Source "PowerShell" -LogName "Microsoft-Windows-PowerShell/Operational" -EventIds @(400,403,600) -Description "PowerShell engine/provider lifecycle" -Relevance "Execution context support" -Technique "General"

# Security auditing.
Add-VisibilityQuery -Source "Security" -LogName "Security" -EventIds @(4688) -Description "Process Creation auditing" -Relevance "Alternative process creation visibility if enabled" -Technique "General" -MessageRegex '(?i)(powershell|cmd|schtasks|reg\.exe|sc\.exe)'
Add-VisibilityQuery -Source "Security" -LogName "Security" -EventIds @(4698,4702,4699) -Description "Scheduled Task auditing" -Relevance "TEC-003 Security auditing if enabled" -Technique "TEC-003"
Add-VisibilityQuery -Source "Security" -LogName "Security" -EventIds @(4657) -Description "Registry value modified auditing" -Relevance "TEC-004 registry auditing if enabled" -Technique "TEC-004"
Add-VisibilityQuery -Source "Security" -LogName "Security" -EventIds @(4663) -Description "File access auditing" -Relevance "Optional file access visibility" -Technique "TEC-007/TEC-008/TEC-009"
Add-VisibilityQuery -Source "Security" -LogName "Security" -EventIds @(4624,4625) -Description "Logon success/failure counts" -Relevance "Context only; not primary for TFM techniques" -Technique "Context"

# System and Application context.
Add-VisibilityQuery -Source "System" -LogName "System" -EventIds @(7045) -Description "Service Control Manager - service created" -Relevance "TEC-005 primary Windows evidence" -Technique "TEC-005"
Add-VisibilityQuery -Source "System" -LogName "System" -EventIds @(7036) -Description "Service state changes" -Relevance "TEC-005 supporting context" -Technique "TEC-005"
Add-VisibilityQuery -Source "Application" -LogName "Application" -EventIds @(1000,1001) -Description "Application errors/crashes" -Relevance "Operational context only" -Technique "Context"

$techniqueCoverage = [ordered]@{
    "TEC-001" = (@($Rows | Where-Object { $_.Technique -match "TEC-001" -and $_.Count -gt 0 }).Count -gt 0)
    "TEC-002" = (@($Rows | Where-Object { $_.Technique -match "TEC-002" -and $_.Count -gt 0 }).Count -gt 0)
    "TEC-003" = (@($Rows | Where-Object { $_.Technique -match "TEC-003" -and $_.Count -gt 0 }).Count -gt 0)
    "TEC-004" = (@($Rows | Where-Object { $_.Technique -match "TEC-004" -and $_.Count -gt 0 }).Count -gt 0)
    "TEC-005" = (@($Rows | Where-Object { $_.Technique -match "TEC-005" -and $_.Count -gt 0 }).Count -gt 0)
    "TEC-006" = (@($Rows | Where-Object { $_.Technique -match "TEC-006" -and $_.Count -gt 0 }).Count -gt 0)
    "TEC-007" = (@($Rows | Where-Object { $_.Technique -match "TEC-007" -and $_.Count -gt 0 }).Count -gt 0)
    "TEC-008" = (@($Rows | Where-Object { $_.Technique -match "TEC-008" -and $_.Count -gt 0 }).Count -gt 0)
    "TEC-009" = (@($Rows | Where-Object { $_.Technique -match "TEC-009" -and $_.Count -gt 0 }).Count -gt 0)
}

$missing = @()
if (-not (Test-LogExists "Microsoft-Windows-Sysmon/Operational")) { $missing += "Sysmon log not available." }
if ((@($Rows | Where-Object { $_.LogName -eq "Microsoft-Windows-PowerShell/Operational" -and $_.EventID -eq "4104" -and $_.Count -gt 0 }).Count) -lt 1) { $missing += "PowerShell 4104 not observed in lookback window." }
if ((@($Rows | Where-Object { $_.LogName -eq "Security" -and $_.EventID -eq "4688" -and $_.Count -gt 0 }).Count) -lt 1) { $missing += "Security 4688 not observed; process creation auditing may be disabled." }
if ((@($Rows | Where-Object { $_.LogName -eq "Microsoft-Windows-Sysmon/Operational" -and $_.EventID -eq "3" -and $_.Count -gt 0 }).Count) -lt 1) { $missing += "Sysmon Event ID 3 not observed; network telemetry may be unavailable or absent in window." }

$summaryObject = [ordered]@{
    Start = $StartTime.ToString("o")
    End = $EndTime.ToString("o")
    LookbackHours = $LookbackHours
    Hostname = $env:COMPUTERNAME
    User = ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
    OutputDir = $OutputDir
    LogFile = $LogFile
    SummaryTxt = $SummaryTxt
    SummaryJson = $SummaryJson
    EventsCsv = $EventsCsv
    TechniqueCoverage = $techniqueCoverage
    Gaps = $missing
    Rows = @($Rows)
}

$lines = New-Object System.Collections.Generic.List[string]
$lines.Add("TFM SYSMON/WINDOWS VISIBILITY SUMMARY") | Out-Null
$lines.Add("============================================================") | Out-Null
$lines.Add(("Window: {0} -> {1}" -f $StartTime.ToString("o"), $EndTime.ToString("o"))) | Out-Null
$lines.Add(("Hostname: {0}" -f $env:COMPUTERNAME)) | Out-Null
$lines.Add(("User: {0}" -f ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name))) | Out-Null
$lines.Add(("OutputDir: {0}" -f $OutputDir)) | Out-Null
$lines.Add("") | Out-Null
$lines.Add("Technique visibility") | Out-Null
$lines.Add("------------------------------------------------------------") | Out-Null
foreach ($key in $techniqueCoverage.Keys) {
    $lines.Add(("{0}: {1}" -f $key, $(if ($techniqueCoverage[$key]) { "VISIBLE" } else { "NO_MATCH_IN_WINDOW" }))) | Out-Null
}
$lines.Add("") | Out-Null
$lines.Add("Gaps") | Out-Null
$lines.Add("------------------------------------------------------------") | Out-Null
if ($missing.Count -gt 0) {
    foreach ($gap in $missing) { $lines.Add(("- {0}" -f $gap)) | Out-Null }
}
else {
    $lines.Add("- No critical visibility gaps detected by this read-only check.") | Out-Null
}
$lines.Add("") | Out-Null
$lines.Add("Source/EventID summary") | Out-Null
$lines.Add("------------------------------------------------------------") | Out-Null
foreach ($row in $Rows) {
    $lines.Add(("{0} | ID={1} | Count={2} | Status={3} | Technique={4} | {5}" -f $row.Source, $row.EventID, $row.Count, $row.Status, $row.Technique, $row.Description)) | Out-Null
    $lines.Add(("  Relevance: {0}" -f $row.Relevance)) | Out-Null
    $lines.Add(("  Comment: {0}" -f $row.Comment)) | Out-Null
}

$lines | Set-Content -LiteralPath $SummaryTxt -Encoding UTF8
$summaryObject | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $SummaryJson -Encoding UTF8
if ($Samples.Count -gt 0) {
    $Samples | Export-Csv -LiteralPath $EventsCsv -NoTypeInformation -Encoding UTF8
}
else {
    @([pscustomobject]@{ Source=""; Technique=""; LogName=""; EventID=""; TimeCreated=""; ProviderName=""; RecordId=""; Message="No samples exported. Use -IncludeSamples or widen -LookbackHours." }) |
        Export-Csv -LiteralPath $EventsCsv -NoTypeInformation -Encoding UTF8
}

Write-Log ("Summary TXT => {0}" -f $SummaryTxt)
Write-Log ("Summary JSON => {0}" -f $SummaryJson)
Write-Log ("Events CSV => {0}" -f $EventsCsv)

Write-Host ""
Write-Host "VISIBILITY CHECK FINISHED"
Write-Host ("Log: {0}" -f $LogFile)
Write-Host ("Summary TXT: {0}" -f $SummaryTxt)
Write-Host ("Summary JSON: {0}" -f $SummaryJson)
Write-Host ("Events CSV: {0}" -f $EventsCsv)
Write-Host ""
Write-Host "Resumen por fuente/EventID:"
$Rows | Sort-Object Source, EventID | Format-Table Source, EventID, Description, Count, Status, Technique -AutoSize
