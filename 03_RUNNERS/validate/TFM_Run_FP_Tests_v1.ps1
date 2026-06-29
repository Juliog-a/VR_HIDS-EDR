[CmdletBinding()]
param(
    [bool]$ClearJsonlBeforeRun = $false,
    [bool]$KeepWorkspace = $true,
    [string]$JsonlPath = "\\VBOXSVR\Carpeta_Compartida_TFM\inbox\soc_alerts.jsonl",
    [string]$OutputDir = "C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas\Logs_Pruebas_TFM\FPs",
    [string]$WorkspaceRoot = "C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas\FP_WORKSPACE",
    [int]$PostRunWaitSeconds = 15
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$RunId = "TFM_FP_{0}" -f (Get-Date -Format "yyyyMMdd_HHmmss")
$RunWorkspace = Join-Path $WorkspaceRoot $RunId
$StartTimeLocal = Get-Date
$StartTimeUtc = $StartTimeLocal.ToUniversalTime()

[System.IO.Directory]::CreateDirectory($OutputDir) | Out-Null
[System.IO.Directory]::CreateDirectory($RunWorkspace) | Out-Null

$LogFile = Join-Path $OutputDir ("{0}.log" -f $RunId)
$SummaryTxt = Join-Path $OutputDir ("{0}_summary.txt" -f $RunId)
$SummaryJson = Join-Path $OutputDir ("{0}_summary.json" -f $RunId)
$SummaryCsv = Join-Path $OutputDir ("{0}_summary.csv" -f $RunId)
$VrHitsCsv = Join-Path $OutputDir ("{0}_vr_hits.csv" -f $RunId)
$VrHitsTxt = Join-Path $OutputDir ("{0}_vr_hits.txt" -f $RunId)

$Script:Results = New-Object "System.Collections.Generic.List[object]"
$Script:Warnings = New-Object "System.Collections.Generic.List[string]"
$Script:JsonlWarnings = New-Object "System.Collections.Generic.List[string]"

# FP runner v1.1 finalization fix: keep post-processing tolerant of empty JSONL,
# singleton PSCustomObject values and Windows PowerShell collection conversion.

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"), $Level, $Message
    [System.IO.File]::AppendAllText($LogFile, $line + [Environment]::NewLine, [System.Text.Encoding]::UTF8)
    Write-Host $line
}

function Add-RunnerWarning {
    param([string]$Message)
    [void]$Script:Warnings.Add($Message)
    Write-Log $Message "WARN"
}

function Convert-ToObjectArray {
    param($Value)
    if ($null -eq $Value) { return @() }
    if ($Value -is [string]) { return @(,$Value) }
    if ($Value -is [System.Management.Automation.PSCustomObject]) { return @(,$Value) }
    if ($Value -is [System.Collections.IEnumerable]) {
        $items = New-Object "System.Collections.ArrayList"
        foreach ($item in $Value) {
            [void]$items.Add($item)
        }
        return $items.ToArray()
    }
    return @(,$Value)
}

function Convert-ToStringArray {
    param($Value)
    return [string[]](Convert-ToObjectArray -Value $Value | ForEach-Object { [string]$_ })
}

function Write-StringList {
    param(
        [string]$Path,
        [System.Collections.Generic.List[string]]$Lines
    )
    [System.IO.File]::WriteAllLines($Path, (Convert-ToStringArray -Value $Lines), [System.Text.Encoding]::UTF8)
}

function Convert-ToFlatString {
    param($Value)
    if ($null -eq $Value) { return "" }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [System.Array]) {
        return (@($Value) | ForEach-Object { Convert-ToFlatString -Value $_ }) -join " "
    }
    if ($Value -is [System.Management.Automation.PSCustomObject]) {
        try {
            return ($Value | ConvertTo-Json -Compress -Depth 6)
        } catch {
            return [string]$Value
        }
    }
    return [string]$Value
}

function Convert-ToPSLiteral {
    param([AllowEmptyString()][string]$Value)
    return "'" + ($Value -replace "'", "''") + "'"
}

function Invoke-NativeCommand {
    param(
        [string]$FilePath,
        [string[]]$Arguments
    )
    $output = @()
    $exitCode = $null
    try {
        $output = & $FilePath @Arguments 2>&1
        $exitCode = $LASTEXITCODE
    } catch {
        $output = @($_.Exception.Message)
        $exitCode = 9999
    }
    return [pscustomobject]@{
        ExitCode = $exitCode
        Output = ($output | ForEach-Object { [string]$_ }) -join "`n"
    }
}

function Invoke-ChildPowerShell {
    param([string]$Command)
    return Invoke-NativeCommand -FilePath "powershell.exe" -Arguments @("-NoProfile", "-Command", $Command)
}

function New-TestReturn {
    param(
        [ValidateSet("OK", "WARN", "FAIL", "SKIPPED")]
        [string]$Status,
        [string[]]$Artifacts = @(),
        [string]$Details = ""
    )
    return [pscustomobject]@{
        Status = $Status
        Artifacts = $Artifacts
        Details = $Details
    }
}

function Invoke-FPCase {
    param(
        [string]$Id,
        [string]$Description,
        [scriptblock]$Action
    )
    $caseStart = Get-Date
    Write-Log ("START {0} - {1}" -f $Id, $Description)

    $status = "OK"
    $details = ""
    $artifacts = @()

    try {
        $result = & $Action
        if ($null -ne $result) {
            if ($result.PSObject.Properties.Name -contains "Status" -and $result.Status) {
                $status = [string]$result.Status
            }
            if ($result.PSObject.Properties.Name -contains "Details" -and $result.Details) {
                $details = [string]$result.Details
            }
            if ($result.PSObject.Properties.Name -contains "Artifacts" -and $result.Artifacts) {
                $artifacts = @($result.Artifacts | ForEach-Object { [string]$_ })
            }
        }
    } catch {
        $status = "FAIL"
        $details = $_.Exception.Message
    }

    $caseEnd = Get-Date
    $row = [pscustomobject]@{
        RunId = $RunId
        FP_ID = $Id
        Description = $Description
        Status = $status
        StartTimeLocal = $caseStart.ToString("o")
        EndTimeLocal = $caseEnd.ToString("o")
        DurationSec = [math]::Round(($caseEnd - $caseStart).TotalSeconds, 2)
        Artifacts = ($artifacts -join "; ")
        Details = $details
    }
    [void]$Script:Results.Add($row)

    if ($status -eq "WARN" -or $status -eq "FAIL" -or $status -eq "SKIPPED") {
        [void]$Script:Warnings.Add(("{0} {1}: {2}" -f $Id, $status, $details))
    }

    Write-Log ("END {0} => {1} ({2}s)" -f $Id, $status, $row.DurationSec)
}

function Convert-ToUtcDate {
    param($Value)
    if ($null -eq $Value) { return $null }
    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    try {
        $dt = [datetime]$text
        if ($dt.Kind -eq [System.DateTimeKind]::Utc) { return $dt }
        return $dt.ToUniversalTime()
    } catch {
        return $null
    }
}

function Get-ObjectField {
    param(
        $Object,
        [string[]]$Names
    )
    foreach ($name in $Names) {
        if ($Object.PSObject.Properties.Name -contains $name) {
            $value = $Object.$name
            $text = Convert-ToFlatString -Value $value
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                return $text
            }
        }
    }
    return ""
}

function Get-ObjectSearchText {
    param($Object)
    $parts = New-Object "System.Collections.Generic.List[string]"
    foreach ($property in $Object.PSObject.Properties) {
        if ($null -ne $property.Value) {
            [void]$parts.Add((Convert-ToFlatString -Value $property.Value))
        }
    }
    return ($parts -join " ")
}

function Find-FPId {
    param([string]$Text)
    $matches = [regex]::Matches($Text, "FP-\d{3}") | ForEach-Object { $_.Value } | Select-Object -Unique
    $items = @($matches)
    if ($items.Count -eq 1) { return $items[0] }
    if ($items.Count -gt 1) { return "MULTIPLE" }
    return "UNKNOWN"
}

function Get-FPImpact {
    param([string]$Profile)
    if ($Profile -match "(?i)P1") { return "FP grave: alerta critica por actividad benigna" }
    if ($Profile -match "(?i)P2") { return "FP forense relevante" }
    if ($Profile -match "(?i)P3") { return "FP medio/comportamental" }
    if ($Profile -match "(?i)P4") { return "FP leve/visibilidad" }
    return "FP sin perfil clasificado"
}

function Read-VelociraptorHits {
    param(
        [datetime]$WindowStartUtc,
        [datetime]$WindowEndUtc
    )

    $hits = New-Object "System.Collections.Generic.List[object]"
    if (-not (Test-Path -LiteralPath $JsonlPath)) {
        Add-RunnerWarning ("JSONL no disponible: {0}" -f $JsonlPath)
        return (Convert-ToObjectArray -Value $hits)
    }

    $lineNumber = 0
    $reader = $null
    try {
        $reader = [System.IO.File]::OpenText($JsonlPath)
        while ($true) {
            $line = $reader.ReadLine()
            if ($null -eq $line) { break }
            $lineNumber++
            if ([string]::IsNullOrWhiteSpace($line)) { continue }

            $row = $null
            try {
                $row = $line | ConvertFrom-Json -ErrorAction Stop
            } catch {
                $warn = "JSONL corrupto saltado en linea {0}: {1}" -f $lineNumber, $_.Exception.Message
                [void]$Script:JsonlWarnings.Add($warn)
                Add-RunnerWarning $warn
                continue
            }

            $timestamp = Get-ObjectField -Object $row -Names @("Timestamp", "DetectionTime", "EventTime", "Time")
            $eventTimeUtc = Convert-ToUtcDate -Value $timestamp
            if (-not [string]::IsNullOrWhiteSpace($timestamp) -and $null -eq $eventTimeUtc) {
                $warn = "Timestamp no parseable saltado para ventana temporal en linea {0}: {1}" -f $lineNumber, $timestamp
                [void]$Script:JsonlWarnings.Add($warn)
                Add-RunnerWarning $warn
            }

            $text = Get-ObjectSearchText -Object $row
            $matchedByRunId = $text -match [regex]::Escape($RunId)
            $matchedByTime = $false
            if ($null -ne $eventTimeUtc) {
                $matchedByTime = ($eventTimeUtc -ge $WindowStartUtc -and $eventTimeUtc -le $WindowEndUtc)
            }

            if (-not ($matchedByTime -or $matchedByRunId)) {
                continue
            }

            $profile = Get-ObjectField -Object $row -Names @("Profile")
            $tec = Get-ObjectField -Object $row -Names @("TEC", "ID_Tecnica_Interna")
            $detectionName = Get-ObjectField -Object $row -Names @("DetectionName", "AlertTitle", "AlertName")
            $severity = Get-ObjectField -Object $row -Names @("Severity")
            $confidence = Get-ObjectField -Object $row -Names @("Confidence")
            $artifact = Get-ObjectField -Object $row -Names @("Artifact")
            $source = Get-ObjectField -Object $row -Names @("Source", "Channel")
            $evidence = Get-ObjectField -Object $row -Names @("IOA", "Evidence", "CU_Evidence", "CommandLine", "RawEventSummary")
            $fpId = Find-FPId -Text $text
            $excerpt = $evidence
            if ($excerpt.Length -gt 300) {
                $excerpt = $excerpt.Substring(0, 300)
            }

            [void]$hits.Add([pscustomobject]@{
                RunId = $RunId
                EventTime = $timestamp
                MatchedByTime = $matchedByTime
                MatchedByRunId = $matchedByRunId
                FP_ID = $fpId
                Profile = $profile
                TEC = $tec
                DetectionName = $detectionName
                Severity = $severity
                Confidence = $confidence
                Artifact = $artifact
                Source = $source
                PotentialFP = Get-FPImpact -Profile $profile
                EvidenceExcerpt = $excerpt
            })
        }
    } catch {
        Add-RunnerWarning ("Error leyendo JSONL; se continua con los hits ya parseados: {0}" -f $_.Exception.Message)
    } finally {
        if ($null -ne $reader) {
            $reader.Close()
        }
    }
    return (Convert-ToObjectArray -Value $hits)
}

function New-GroupSummary {
    param(
        [object[]]$Rows,
        [string]$PropertyName
    )
    return @(@($Rows) | Group-Object -Property $PropertyName | Sort-Object Count -Descending | Select-Object Name, Count)
}

function Write-EmptyVrHitsCsv {
    param([string]$Path)
    [System.IO.File]::WriteAllText($Path, "RunId,EventTime,MatchedByTime,MatchedByRunId,FP_ID,Profile,TEC,DetectionName,Severity,Confidence,Artifact,Source,PotentialFP,EvidenceExcerpt`r`n", [System.Text.Encoding]::UTF8)
}

function Export-RunnerCsv {
    param(
        [object[]]$Rows,
        [string]$Path,
        [string]$EmptyHeader
    )
    if (@($Rows).Count -gt 0) {
        @($Rows) | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
    } else {
        [System.IO.File]::WriteAllText($Path, $EmptyHeader, [System.Text.Encoding]::UTF8)
    }
}

function Write-TextReport {
    param(
        [object[]]$VrHits,
        [datetime]$EndTimeLocal,
        [datetime]$EndTimeUtc,
        [string]$GlobalStatus
    )

    $results = @(Convert-ToObjectArray -Value $Script:Results)
    $hits = @($VrHits)
    $statusGroups = @($results | Group-Object Status | Sort-Object Name)
    $profileGroups = @($hits | Group-Object Profile | Sort-Object Count -Descending)
    $fpGroups = @($hits | Group-Object FP_ID | Sort-Object Count -Descending)
    $detGroups = @($hits | Group-Object DetectionName | Sort-Object Count -Descending)

    $lines = New-Object "System.Collections.Generic.List[string]"
    [void]$lines.Add("# TFM FP Runner Summary")
    [void]$lines.Add("")
    [void]$lines.Add(("RunId: {0}" -f $RunId))
    [void]$lines.Add(("GlobalStatus: {0}" -f $GlobalStatus))
    [void]$lines.Add(("StartTimeLocal: {0}" -f $StartTimeLocal.ToString("o")))
    [void]$lines.Add(("EndTimeLocal: {0}" -f $EndTimeLocal.ToString("o")))
    [void]$lines.Add(("StartTimeUtc: {0}" -f $StartTimeUtc.ToString("o")))
    [void]$lines.Add(("EndTimeUtc: {0}" -f $EndTimeUtc.ToString("o")))
    [void]$lines.Add(("Workspace: {0}" -f $RunWorkspace))
    [void]$lines.Add(("JsonlPath: {0}" -f $JsonlPath))
    [void]$lines.Add("")
    [void]$lines.Add("## Test status")
    foreach ($group in $statusGroups) {
        [void]$lines.Add(("- {0}: {1}" -f $group.Name, $group.Count))
    }
    [void]$lines.Add("")
    [void]$lines.Add("## Tests")
    foreach ($r in $Script:Results) {
        [void]$lines.Add(("- {0} [{1}] {2} | {3}" -f $r.FP_ID, $r.Status, $r.Description, $r.Details))
    }
    [void]$lines.Add("")
    [void]$lines.Add(("## Velociraptor hits: {0}" -f $VrHits.Count))
    [void]$lines.Add("")
    [void]$lines.Add("### Hits by Profile")
    foreach ($group in $profileGroups) {
        [void]$lines.Add(("- {0}: {1}" -f $group.Name, $group.Count))
    }
    [void]$lines.Add("")
    [void]$lines.Add("### Hits by FP_ID")
    foreach ($group in $fpGroups) {
        [void]$lines.Add(("- {0}: {1}" -f $group.Name, $group.Count))
    }
    [void]$lines.Add("")
    [void]$lines.Add("### Hits by DetectionName")
    foreach ($group in $detGroups) {
        [void]$lines.Add(("- {0}: {1}" -f $group.Name, $group.Count))
    }
    [void]$lines.Add("")
    [void]$lines.Add("## Interpretation")
    [void]$lines.Add("- P4/P3 hits indicate possible low/medium false positives or expected broad visibility.")
    [void]$lines.Add("- P2 hits indicate relevant forensic false positives to review.")
    [void]$lines.Add("- P1_CRITICAL hits indicate severe false positives for benign activity.")
    if ($Script:JsonlWarnings.Count -gt 0) {
        [void]$lines.Add("")
        [void]$lines.Add("## JSONL warnings")
        foreach ($warning in @(Convert-ToObjectArray -Value $Script:JsonlWarnings)) {
            [void]$lines.Add(("- {0}" -f $warning))
        }
    }

    Write-StringList -Path $SummaryTxt -Lines $lines

    $hitLines = New-Object "System.Collections.Generic.List[string]"
    [void]$hitLines.Add("# TFM FP Velociraptor Hits")
    [void]$hitLines.Add("")
    [void]$hitLines.Add(("RunId: {0}" -f $RunId))
    [void]$hitLines.Add(("TotalHits: {0}" -f $VrHits.Count))
    [void]$hitLines.Add("")
    foreach ($h in $VrHits) {
        [void]$hitLines.Add(("[{0}] {1} {2} {3}/{4} {5} FP_ID={6}" -f $h.EventTime, $h.Profile, $h.TEC, $h.Severity, $h.Confidence, $h.DetectionName, $h.FP_ID))
        [void]$hitLines.Add(("  Artifact: {0}" -f $h.Artifact))
        [void]$hitLines.Add(("  Source: {0}" -f $h.Source))
        [void]$hitLines.Add(("  PotentialFP: {0}" -f $h.PotentialFP))
        [void]$hitLines.Add(("  Evidence: {0}" -f $h.EvidenceExcerpt))
        [void]$hitLines.Add("")
    }
    Write-StringList -Path $VrHitsTxt -Lines $hitLines
}

Write-Log ("RunId: {0}" -f $RunId)
Write-Log ("Workspace: {0}" -f $RunWorkspace)
Write-Log ("OutputDir: {0}" -f $OutputDir)
Write-Log ("JsonlPath: {0}" -f $JsonlPath)

if ($ClearJsonlBeforeRun) {
    if (Test-Path -LiteralPath $JsonlPath) {
        [System.IO.File]::WriteAllText($JsonlPath, "", [System.Text.Encoding]::UTF8)
        Write-Log "JSONL cleared before run"
    } else {
        Write-Log "ClearJsonlBeforeRun requested but JSONL does not exist yet" "WARN"
    }
}

Invoke-FPCase -Id "FP-001" -Description "PowerShell administrativo legitimo" -Action {
    $id = "FP-001"
    $sample = Join-Path $RunWorkspace ("{0}_benign.txt" -f $id)
    [System.IO.File]::WriteAllText($sample, "$RunId $id benign content`nsecond benign line", [System.Text.Encoding]::UTF8)
    $cmdGetItems = "Get" + "-ChildItem"
    $sampleLit = Convert-ToPSLiteral $sample
    $workspaceLit = Convert-ToPSLiteral $RunWorkspace
    $cmd = @"
`$ErrorActionPreference = 'Stop'
`$rid = '$RunId'
`$fp = '$id'
Write-Output "`$rid `$fp start"
Get-Process | Select-Object -First 5 | Out-Null
Get-Service | Select-Object -First 5 | Out-Null
& '$cmdGetItems' -LiteralPath $workspaceLit | Out-Null
Select-String -Path $sampleLit -Pattern 'benign' | Out-Null
Write-Output "`$rid `$fp ok"
"@
    $r = Invoke-ChildPowerShell -Command $cmd
    if ($r.ExitCode -ne 0) {
        return New-TestReturn -Status "WARN" -Artifacts @($sample) -Details $r.Output
    }
    return New-TestReturn -Status "OK" -Artifacts @($sample) -Details "PowerShell administrative commands completed"
}

Invoke-FPCase -Id "FP-002" -Description "CMD legitimo" -Action {
    $id = "FP-002"
    $out = New-Object "System.Collections.Generic.List[string]"
    $workspaceQuoted = '"' + $RunWorkspace + '"'
    foreach ($cmdLine in @("dir $workspaceQuoted", "whoami", "hostname", "echo $RunId $id")) {
        $r = Invoke-NativeCommand -FilePath "cmd.exe" -Arguments @("/c", $cmdLine)
        [void]$out.Add(("cmd /c {0} => {1}" -f $cmdLine, $r.ExitCode))
        if ($r.ExitCode -ne 0) {
            return New-TestReturn -Status "WARN" -Details (($out -join " | ") + " | " + $r.Output)
        }
    }
    return New-TestReturn -Status "OK" -Details ($out -join " | ")
}

Invoke-FPCase -Id "FP-003" -Description "Scheduled Task benigno temporal" -Action {
    $id = "FP-003"
    $taskName = "\TFM_FP_{0}_{1}" -f $RunId, $id
    $taskOutput = Join-Path $RunWorkspace ("{0}_scheduled_task_ok.txt" -f $id)
    $runAt = (Get-Date).AddMinutes(5).ToString("HH:mm")
    $taskCommand = 'cmd.exe /c echo {0} {1} scheduled task ok > "{2}"' -f $RunId, $id, $taskOutput

    $create = Invoke-NativeCommand -FilePath "schtasks.exe" -Arguments @("/Create", "/TN", $taskName, "/TR", $taskCommand, "/SC", "ONCE", "/ST", $runAt, "/F")
    if ($create.ExitCode -ne 0) {
        if ($create.Output -match "(?i)(access is denied|acceso denegado|privileg)") {
            return New-TestReturn -Status "SKIPPED" -Details ("No permissions to create task: {0}" -f $create.Output)
        }
        return New-TestReturn -Status "WARN" -Details ("Task create returned {0}: {1}" -f $create.ExitCode, $create.Output)
    }

    $run = Invoke-NativeCommand -FilePath "schtasks.exe" -Arguments @("/Run", "/TN", $taskName)
    Start-Sleep -Seconds 3
    $delete = Invoke-NativeCommand -FilePath "schtasks.exe" -Arguments @("/Delete", "/TN", $taskName, "/F")

    $details = "create=$($create.ExitCode); run=$($run.ExitCode); delete=$($delete.ExitCode)"
    if ($run.ExitCode -ne 0 -or $delete.ExitCode -ne 0) {
        return New-TestReturn -Status "WARN" -Artifacts @($taskOutput) -Details ($details + " | " + $run.Output + " | " + $delete.Output)
    }
    if (-not (Test-Path -LiteralPath $taskOutput)) {
        return New-TestReturn -Status "WARN" -Artifacts @($taskOutput) -Details ($details + "; task output file not observed")
    }
    return New-TestReturn -Status "OK" -Artifacts @($taskOutput) -Details $details
}

Invoke-FPCase -Id "FP-004" -Description "Registro benigno no persistente" -Action {
    $id = "FP-004"
    $subKey = "Software\TFM_FP\$RunId"
    $key = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($subKey)
    try {
        $key.SetValue("FP004Value", "$RunId $id benign registry value", [Microsoft.Win32.RegistryValueKind]::String)
        $value = [string]$key.GetValue("FP004Value", "")
        if ($value -notmatch [regex]::Escape($RunId)) {
            return New-TestReturn -Status "WARN" -Details "Registry value was not read back correctly"
        }
    } finally {
        if ($null -ne $key) { $key.Close() }
        [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree($subKey, $false)
    }
    return New-TestReturn -Status "OK" -Details ("HKCU:\{0} created, read and removed" -f $subKey)
}

Invoke-FPCase -Id "FP-005" -Description "Consulta de servicios" -Action {
    $id = "FP-005"
    $cmd = @"
`$ErrorActionPreference = 'Stop'
`$rid = '$RunId'
`$fp = '$id'
Get-Service | Select-Object -First 10 | Out-Null
sc.exe query | Out-Null
Write-Output "`$rid `$fp service query ok"
"@
    $r = Invoke-ChildPowerShell -Command $cmd
    if ($r.ExitCode -ne 0) {
        return New-TestReturn -Status "WARN" -Details $r.Output
    }
    return New-TestReturn -Status "OK" -Details "Service query commands completed"
}

Invoke-FPCase -Id "FP-006" -Description "Discovery legitimo de seguridad" -Action {
    $id = "FP-006"
    $cmd = @"
`$ErrorActionPreference = 'Stop'
`$rid = '$RunId'
`$fp = '$id'
`$mp = Get-Command Get-MpComputerStatus -ErrorAction SilentlyContinue
if (`$null -ne `$mp) { Get-MpComputerStatus | Out-Null }
Get-Process -Name SecurityHealthSystray,MsMpEng -ErrorAction SilentlyContinue | Out-Null
Write-Output "`$rid `$fp security discovery ok"
"@
    $r = Invoke-ChildPowerShell -Command $cmd
    if ($r.ExitCode -ne 0) {
        return New-TestReturn -Status "WARN" -Details $r.Output
    }
    return New-TestReturn -Status "OK" -Details "Security status queries completed"
}

Invoke-FPCase -Id "FP-007" -Description "ZIP benigno local" -Action {
    $id = "FP-007"
    $zipDir = Join-Path $RunWorkspace $id
    [System.IO.Directory]::CreateDirectory($zipDir) | Out-Null
    for ($i = 1; $i -le 5; $i++) {
        $file = Join-Path $zipDir ("benign_{0}.txt" -f $i)
        [System.IO.File]::WriteAllText($file, "$RunId $id benign archive file $i", [System.Text.Encoding]::UTF8)
    }
    $zipPath = Join-Path $RunWorkspace ("{0}_benign_archive.zip" -f $id)
    $compress = Get-Command Compress-Archive -ErrorAction SilentlyContinue
    if ($null -eq $compress) {
        return New-TestReturn -Status "SKIPPED" -Artifacts @($zipDir) -Details "Compress-Archive not available"
    }
    Compress-Archive -Path (Join-Path $zipDir "*.txt") -DestinationPath $zipPath -Force
    if (-not (Test-Path -LiteralPath $zipPath)) {
        return New-TestReturn -Status "FAIL" -Artifacts @($zipDir, $zipPath) -Details "ZIP was not created"
    }
    return New-TestReturn -Status "OK" -Artifacts @($zipDir, $zipPath) -Details "Local ZIP created without web transfer"
}

Invoke-FPCase -Id "FP-008" -Description "Borrado benigno controlado" -Action {
    $id = "FP-008"
    $deleteDir = Join-Path $RunWorkspace $id
    [System.IO.Directory]::CreateDirectory($deleteDir) | Out-Null
    for ($i = 1; $i -le 3; $i++) {
        $file = Join-Path $deleteDir ("delete_me_{0}.txt" -f $i)
        [System.IO.File]::WriteAllText($file, "$RunId $id benign delete file $i", [System.Text.Encoding]::UTF8)
    }
    $cmdRemove = "Remove" + "-Item"
    $argRecurse = "-" + "Recurse"
    $argForce = "-" + "Force"
    $deleteDirLit = Convert-ToPSLiteral $deleteDir
    $cmd = @"
`$ErrorActionPreference = 'Stop'
`$rid = '$RunId'
`$fp = '$id'
& '$cmdRemove' -LiteralPath $deleteDirLit $argRecurse $argForce
Write-Output "`$rid `$fp controlled delete ok"
"@
    $r = Invoke-ChildPowerShell -Command $cmd
    if ($r.ExitCode -ne 0) {
        return New-TestReturn -Status "WARN" -Artifacts @($deleteDir) -Details $r.Output
    }
    if (Test-Path -LiteralPath $deleteDir) {
        return New-TestReturn -Status "WARN" -Artifacts @($deleteDir) -Details "Controlled delete directory still exists"
    }
    return New-TestReturn -Status "OK" -Artifacts @($deleteDir) -Details "Only the controlled workspace subfolder was deleted"
}

Invoke-FPCase -Id "FP-009" -Description "HTTP benigno local" -Action {
    $id = "FP-009"
    $tnc = Test-NetConnection -ComputerName "127.0.0.1" -Port 80 -WarningAction SilentlyContinue
    if (-not $tnc.TcpTestSucceeded) {
        return New-TestReturn -Status "SKIPPED" -Details "No local HTTP listener on 127.0.0.1:80; Test-NetConnection completed"
    }
    $cmdWeb = "Invoke" + "-WebRequest"
    $cmd = @"
`$ErrorActionPreference = 'Stop'
`$rid = '$RunId'
`$fp = '$id'
& '$cmdWeb' -Uri 'http://127.0.0.1/' -Method GET -UseBasicParsing -TimeoutSec 5 | Out-Null
Write-Output "`$rid `$fp http get ok"
"@
    $r = Invoke-ChildPowerShell -Command $cmd
    if ($r.ExitCode -ne 0) {
        return New-TestReturn -Status "WARN" -Details $r.Output
    }
    return New-TestReturn -Status "OK" -Details "Local HTTP GET completed"
}

Invoke-FPCase -Id "FP-010" -Description "Operaciones normales masivas" -Action {
    $id = "FP-010"
    $bulkDir = Join-Path $RunWorkspace $id
    $copyDir = Join-Path $bulkDir "copy"
    [System.IO.Directory]::CreateDirectory($bulkDir) | Out-Null
    [System.IO.Directory]::CreateDirectory($copyDir) | Out-Null
    for ($i = 1; $i -le 20; $i++) {
        $sourceFile = Join-Path $bulkDir ("normal_{0:D2}.txt" -f $i)
        [System.IO.File]::WriteAllText($sourceFile, "$RunId $id normal file operation $i", [System.Text.Encoding]::UTF8)
        $destFile = Join-Path $copyDir ("normal_{0:D2}.txt" -f $i)
        [System.IO.File]::Copy($sourceFile, $destFile, $true)
    }
    for ($i = 1; $i -le 5; $i++) {
        $oldPath = Join-Path $copyDir ("normal_{0:D2}.txt" -f $i)
        $newPath = Join-Path $copyDir ("renamed_{0:D2}.txt" -f $i)
        if (Test-Path -LiteralPath $oldPath) {
            [System.IO.File]::Move($oldPath, $newPath)
        }
    }
    return New-TestReturn -Status "OK" -Artifacts @($bulkDir, $copyDir) -Details "20 files created, copied and partially renamed"
}

if (-not $KeepWorkspace) {
    try {
        if (Test-Path -LiteralPath $RunWorkspace) {
            [System.IO.Directory]::Delete($RunWorkspace, $true)
            Write-Log "Workspace removed because KeepWorkspace=false"
        }
    } catch {
        Write-Log ("Workspace cleanup warning: {0}" -f $_.Exception.Message) "WARN"
    }
}

if ($PostRunWaitSeconds -gt 0) {
    Write-Log ("Waiting {0}s before reading SOC JSONL" -f $PostRunWaitSeconds)
    Start-Sleep -Seconds $PostRunWaitSeconds
}

$EndTimeLocal = Get-Date
$EndTimeUtc = $EndTimeLocal.ToUniversalTime()

$vrHits = @(Read-VelociraptorHits -WindowStartUtc $StartTimeUtc -WindowEndUtc $EndTimeUtc)
$results = @(Convert-ToObjectArray -Value $Script:Results)

$okCount = @($results | Where-Object { $_.Status -eq "OK" }).Count
$warnCount = @($results | Where-Object { $_.Status -eq "WARN" }).Count
$failCount = @($results | Where-Object { $_.Status -eq "FAIL" }).Count
$skippedCount = @($results | Where-Object { $_.Status -eq "SKIPPED" }).Count

if ($failCount -gt 0) {
    $globalStatus = "FAIL"
} elseif ($warnCount -gt 0 -or $skippedCount -gt 0 -or $vrHits.Count -gt 0) {
    $globalStatus = "WARN"
} else {
    $globalStatus = "OK"
}

Export-RunnerCsv -Rows $results -Path $SummaryCsv -EmptyHeader "RunId,FP_ID,Description,Status,StartTimeLocal,EndTimeLocal,DurationSec,Artifacts,Details`r`n"

if ($vrHits.Count -gt 0) {
    @($vrHits) | Export-Csv -LiteralPath $VrHitsCsv -NoTypeInformation -Encoding UTF8
} else {
    Write-EmptyVrHitsCsv -Path $VrHitsCsv
}

$summaryObject = [pscustomobject]@{
    RunId = $RunId
    GlobalStatus = $globalStatus
    StartTimeLocal = $StartTimeLocal.ToString("o")
    EndTimeLocal = $EndTimeLocal.ToString("o")
    StartTimeUtc = $StartTimeUtc.ToString("o")
    EndTimeUtc = $EndTimeUtc.ToString("o")
    Workspace = $RunWorkspace
    OutputDir = $OutputDir
    JsonlPath = $JsonlPath
    ClearJsonlBeforeRun = $ClearJsonlBeforeRun
    KeepWorkspace = $KeepWorkspace
    Tests = $results
    TestCounts = [pscustomobject]@{
        OK = $okCount
        WARN = $warnCount
        FAIL = $failCount
        SKIPPED = $skippedCount
    }
    Warnings = @(Convert-ToObjectArray -Value $Script:Warnings)
    JsonlWarnings = @(Convert-ToObjectArray -Value $Script:JsonlWarnings)
    VelociraptorHits = [pscustomobject]@{
        Total = $vrHits.Count
        ByProfile = New-GroupSummary -Rows $vrHits -PropertyName "Profile"
        ByFP_ID = New-GroupSummary -Rows $vrHits -PropertyName "FP_ID"
        ByTEC = New-GroupSummary -Rows $vrHits -PropertyName "TEC"
        ByDetectionName = New-GroupSummary -Rows $vrHits -PropertyName "DetectionName"
    }
    OutputFiles = [pscustomobject]@{
        Log = $LogFile
        SummaryTxt = $SummaryTxt
        SummaryJson = $SummaryJson
        SummaryCsv = $SummaryCsv
        VrHitsCsv = $VrHitsCsv
        VrHitsTxt = $VrHitsTxt
    }
}

$summaryObject | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $SummaryJson -Encoding UTF8
Write-TextReport -VrHits $vrHits -EndTimeLocal $EndTimeLocal -EndTimeUtc $EndTimeUtc -GlobalStatus $globalStatus

Write-Log ("GlobalStatus: {0}" -f $globalStatus)
Write-Log ("Tests OK/WARN/FAIL/SKIPPED: {0}/{1}/{2}/{3}" -f $okCount, $warnCount, $failCount, $skippedCount)
Write-Log ("Total VR hits: {0}" -f $vrHits.Count)

Write-Host ""
Write-Host "===== TFM FP RUN SUMMARY ====="
Write-Host ("RunId: {0}" -f $RunId)
Write-Host ("GlobalStatus: {0}" -f $globalStatus)
Write-Host ("OK/WARN/FAIL/SKIPPED: {0}/{1}/{2}/{3}" -f $okCount, $warnCount, $failCount, $skippedCount)
Write-Host ("Total VR hits: {0}" -f $vrHits.Count)
Write-Host ""
Write-Host "Hits by Profile:"
@($vrHits | Group-Object Profile | Sort-Object Count -Descending | Select-Object Count, Name) | Format-Table -AutoSize
Write-Host "Hits by FP_ID:"
@($vrHits | Group-Object FP_ID | Sort-Object Count -Descending | Select-Object Count, Name) | Format-Table -AutoSize
Write-Host "Generated files:"
Write-Host ("- Log: {0}" -f $LogFile)
Write-Host ("- Summary TXT: {0}" -f $SummaryTxt)
Write-Host ("- Summary JSON: {0}" -f $SummaryJson)
Write-Host ("- Summary CSV: {0}" -f $SummaryCsv)
Write-Host ("- VR hits CSV: {0}" -f $VrHitsCsv)
Write-Host ("- VR hits TXT: {0}" -f $VrHitsTxt)

if ($globalStatus -eq "FAIL") {
    exit 1
}
exit 0
