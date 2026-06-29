#Requires -RunAsAdministrator
<#
.SYNOPSIS
  Ejecuta TEC-001 a TEC-009 para pruebas TFM con Velociraptor/Sysmon/Hayabusa.

.DESCRIPTION
  Version corregida:
  - Corrige errores de Write-Step con cadena vacia.
  - Corrige quoting de schtasks.exe.
  - Corrige quoting de sc.exe.
  - Mantiene 10 segundos entre pruebas.
  - Restaura DUMB_LAB antes/despues de TEC-007, TEC-008 y TEC-009.
  - Ejecuta enumeracion.ps1 antes de TEC-007.
  - TEC-008 borra solo dentro de DUMB_LAB.
  - TEC-009 solo hace staging local: Compress-Archive + Copy-Item.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

# =========================
# CONFIGURACION
# =========================

$BasePath      = "C:\Users\seguridad\Desktop\TFM\Pruebas"
$DumbLab       = Join-Path $BasePath "DUMB_LAB"

$RansomDir     = Join-Path $BasePath "TEC-007_Ransomware"
$SabotageDir   = Join-Path $BasePath "TEC-008_Sabotaje"
$ExfilDir      = Join-Path $BasePath "TEC-009_Exfiltracion"

$RansomScript  = Join-Path $RansomDir "Scriptransom.ps1"
$EnumScript    = Join-Path $RansomDir "enumeracion.ps1"

$RestoreRansom = Join-Path $RansomDir "crear_directorio_dummy.ps1"
$RestoreSab    = Join-Path $SabotageDir "crear_directorio_dummy.ps1"
$RestoreExfil  = Join-Path $ExfilDir "crear_directorio_dummy.ps1"

$PauseBetweenTestsSeconds = 10
$ShortPauseSeconds = 2

$LogDir = Join-Path $BasePath "Logs_Pruebas_TFM"
New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
$LogFile = Join-Path $LogDir ("TFM_TEC_Run_FIXED_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

# =========================
# FUNCIONES
# =========================

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

function Invoke-CmdLine {
    param(
        [Parameter(Mandatory=$true)][string]$Command,
        [string]$Description = "",
        [switch]$IgnoreExitCode
    )

    if ($Description) {
        Write-Step $Description
    }

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

        if ($stdout.Trim()) {
            Write-Step ("STDOUT => {0}" -f ($stdout.Trim() -replace "`r?`n"," | "))
        }

        if ($stderr.Trim()) {
            Write-Step ("STDERR => {0}" -f ($stderr.Trim() -replace "`r?`n"," | ")) "WARN"
        }

        Write-Step ("ExitCode => {0}" -f $p.ExitCode)

        if (($p.ExitCode -ne 0) -and (-not $IgnoreExitCode)) {
            Write-Step ("Comando con ExitCode no cero: {0}" -f $p.ExitCode) "WARN"
        }

        return $p.ExitCode
    }
    catch {
        Write-Step ("ERROR ejecutando comando: {0}" -f $_.Exception.Message) "ERROR"
        return -9999
    }
}

function Invoke-PSFileIfExists {
    param(
        [Parameter(Mandatory=$true)][string]$ScriptPath,
        [string]$Description = ""
    )

    if (Test-Path $ScriptPath) {
        $cmd = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{0}"' -f $ScriptPath
        Invoke-CmdLine -Command $cmd -Description $Description | Out-Null
    }
    else {
        Write-Step ("No existe el script: {0}" -f $ScriptPath) "WARN"
    }
}

function Restore-DumbLab {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet("Ransomware","Sabotaje","Exfiltracion","All")]
        [string]$Mode
    )

    Write-Step "Restaurando entorno DUMB_LAB. Modo: $Mode"

    switch ($Mode) {
        "Ransomware"   { Invoke-PSFileIfExists -ScriptPath $RestoreRansom -Description "Restauracion DUMB_LAB desde Ransomware" }
        "Sabotaje"     { Invoke-PSFileIfExists -ScriptPath $RestoreSab    -Description "Restauracion DUMB_LAB desde Sabotaje" }
        "Exfiltracion" { Invoke-PSFileIfExists -ScriptPath $RestoreExfil  -Description "Restauracion DUMB_LAB desde Exfiltracion" }
        "All" {
            Invoke-PSFileIfExists -ScriptPath $RestoreRansom -Description "Restauracion DUMB_LAB desde Ransomware"
            Invoke-PSFileIfExists -ScriptPath $RestoreSab    -Description "Restauracion DUMB_LAB desde Sabotaje"
            Invoke-PSFileIfExists -ScriptPath $RestoreExfil  -Description "Restauracion DUMB_LAB desde Exfiltracion"
        }
    }

    if (-not (Test-Path $DumbLab)) {
        Write-Step "DUMB_LAB no existe tras restaurar. Creando directorio minimo seguro." "WARN"
        New-Item -ItemType Directory -Force -Path $DumbLab | Out-Null

        1..30 | ForEach-Object {
            Set-Content -Path (Join-Path $DumbLab ("dummy_{0}.txt" -f $_)) -Value "dummy"
        }
    }
}

function Assert-DumbLabSafe {
    $Expected = "C:\Users\seguridad\Desktop\TFM\Pruebas\DUMB_LAB"

    if ($DumbLab -ne $Expected) {
        throw "Ruta DUMB_LAB no coincide con la ruta segura esperada. Abortado."
    }

    if (-not (Test-Path $DumbLab)) {
        throw "DUMB_LAB no existe. Abortado para evitar operaciones fuera del laboratorio."
    }
}

function Pause-BetweenTests {
    Write-Step "Esperando $PauseBetweenTestsSeconds segundos antes de la siguiente prueba..."
    Start-Sleep -Seconds $PauseBetweenTestsSeconds
}

function Start-Test {
    param(
        [Parameter(Mandatory=$true)][string]$Tec,
        [Parameter(Mandatory=$true)][string]$Name
    )

    Write-Step
    Write-Step "============================================================"
    Write-Step "INICIO $Tec - $Name"
    Write-Step "============================================================"
}

function End-Test {
    param(
        [Parameter(Mandatory=$true)][string]$Tec
    )

    Write-Step "FIN $Tec"
    Write-Step "============================================================"
    Pause-BetweenTests
}

# =========================
# PRECHECKS
# =========================

Write-Step "Inicio de campaña TEC-001 a TEC-009"
Write-Step "Log: $LogFile"

if (-not (Test-Path $BasePath)) {
    throw "No existe BasePath: $BasePath"
}

Write-Step "BasePath: $BasePath"
Write-Step "DUMB_LAB: $DumbLab"

# Limpieza defensiva de restos previos
Invoke-CmdLine -Command 'schtasks.exe /delete /tn "TFM-TEC-003" /f' -Description "Limpieza previa tarea TEC-003" -IgnoreExitCode | Out-Null
Invoke-CmdLine -Command 'sc.exe delete "TFM TEC005 Test Service"' -Description "Limpieza previa servicio TEC-005" -IgnoreExitCode | Out-Null
Invoke-CmdLine -Command 'reg.exe delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "TFM_T1547_001_test" /f' -Description "Limpieza previa RunKey TEC-004" -IgnoreExitCode | Out-Null

Restore-DumbLab -Mode "All"

# =========================
# TEC-001 — PowerShell / T1059.001
# =========================

Start-Test -Tec "TEC-001" -Name "PowerShell / T1059.001"

Invoke-CmdLine -Command 'powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Write-Host TEC001_OK; Start-Sleep 3"' `
    -Description "TEC-001: PowerShell con NoProfile y ExecutionPolicy Bypass" | Out-Null

End-Test -Tec "TEC-001"

# =========================
# TEC-002 — CMD / T1059.003
# =========================

Start-Test -Tec "TEC-002" -Name "Windows Command Shell / T1059.003"

Invoke-CmdLine -Command 'cmd.exe /c "echo TEC002_OK & powershell.exe -NoProfile -Command Get-Process"' `
    -Description "TEC-002: cmd.exe con /c y ejecución de PowerShell" | Out-Null

End-Test -Tec "TEC-002"

# =========================
# TEC-003 — Scheduled Task / T1053.005
# =========================

Start-Test -Tec "TEC-003" -Name "Scheduled Task / T1053.005"

$taskTime = (Get-Date).AddMinutes(5).ToString("HH:mm")
$tec003Create = 'schtasks.exe /create /tn "TFM-TEC-003" /tr "cmd.exe /c echo TEC003_OK ^> C:\Users\Public\tec003_ok.txt" /sc once /st {0} /f' -f $taskTime

Invoke-CmdLine -Command $tec003Create -Description "TEC-003: crear tarea programada" | Out-Null
Start-Sleep -Seconds $ShortPauseSeconds

Invoke-CmdLine -Command 'schtasks.exe /run /tn "TFM-TEC-003"' -Description "TEC-003: ejecutar tarea programada" -IgnoreExitCode | Out-Null
Start-Sleep -Seconds $ShortPauseSeconds

Invoke-CmdLine -Command 'schtasks.exe /delete /tn "TFM-TEC-003" /f' -Description "TEC-003: eliminar tarea programada" -IgnoreExitCode | Out-Null

End-Test -Tec "TEC-003"

# =========================
# TEC-004 — Run Key / T1547.001
# =========================

Start-Test -Tec "TEC-004" -Name "Registry Run Key / T1547.001"

Invoke-CmdLine -Command 'reg.exe add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "TFM_T1547_001_test" /t REG_SZ /d "C:\Windows\System32\calc.exe" /f' `
    -Description "TEC-004: crear clave Run" | Out-Null

Start-Sleep -Seconds $ShortPauseSeconds

Invoke-CmdLine -Command 'reg.exe delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "TFM_T1547_001_test" /f' `
    -Description "TEC-004: eliminar clave Run" | Out-Null

End-Test -Tec "TEC-004"

# =========================
# TEC-005 — Service Execution / T1569.002
# =========================

Start-Test -Tec "TEC-005" -Name "Service Execution / T1569.002"

Invoke-CmdLine -Command 'sc.exe create "TFM TEC005 Test Service" binPath= "C:\Windows\System32\cmd.exe /c echo TEC005_OK ^> C:\Users\Public\tec005_ok.txt" start= demand' `
    -Description "TEC-005: crear servicio" | Out-Null

Start-Sleep -Seconds $ShortPauseSeconds

# Puede devolver error porque cmd.exe no es un binario de servicio real. La evidencia principal es EventID 7045.
Invoke-CmdLine -Command 'sc.exe start "TFM TEC005 Test Service"' `
    -Description "TEC-005: intentar arrancar servicio" -IgnoreExitCode | Out-Null

Start-Sleep -Seconds $ShortPauseSeconds

Invoke-CmdLine -Command 'sc.exe delete "TFM TEC005 Test Service"' `
    -Description "TEC-005: eliminar servicio" -IgnoreExitCode | Out-Null

End-Test -Tec "TEC-005"

# =========================
# TEC-006 — Security Software Discovery / T1518.001
# =========================

Start-Test -Tec "TEC-006" -Name "Security Software Discovery / T1518.001"

Invoke-CmdLine -Command 'powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntivirusProduct | Select-Object displayName,productState"' `
    -Description "TEC-006: discovery AV/EDR mediante PowerShell" | Out-Null

Start-Sleep -Seconds $ShortPauseSeconds

Invoke-CmdLine -Command 'cmd.exe /c "wmic /namespace:\\root\SecurityCenter2 path AntivirusProduct get displayName,productState"' `
    -Description "TEC-006: discovery AV/EDR mediante WMIC" | Out-Null

End-Test -Tec "TEC-006"

# =========================
# TEC-007 — Data Encrypted for Impact / T1486
# =========================

Start-Test -Tec "TEC-007" -Name "Data Encrypted for Impact / T1486"

Restore-DumbLab -Mode "Ransomware"

Invoke-PSFileIfExists -ScriptPath $EnumScript -Description "TEC-007: ejecutar enumeracion.ps1 antes del simulador"

if (Test-Path $RansomScript) {
    $cmd = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{0}"' -f $RansomScript
    Invoke-CmdLine -Command $cmd -Description "TEC-007: ejecutar simulador controlado Scriptransom.ps1" | Out-Null
}
else {
    Write-Step "No existe Scriptransom.ps1. Se omite TEC-007." "ERROR"
}

Start-Sleep -Seconds $ShortPauseSeconds
Restore-DumbLab -Mode "Ransomware"

End-Test -Tec "TEC-007"

# =========================
# TEC-008 — Data Destruction / T1485
# =========================

Start-Test -Tec "TEC-008" -Name "Data Destruction / T1485"

Restore-DumbLab -Mode "Sabotaje"
Assert-DumbLabSafe

$tec008Command = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$target = ''{0}''; Write-Host ''TEC008_OK''; Get-ChildItem $target -Recurse -File | ForEach-Object {{ Remove-Item $_.FullName -Force }}; Start-Sleep 3"' -f $DumbLab

Invoke-CmdLine -Command $tec008Command -Description "TEC-008: borrado masivo controlado en DUMB_LAB" | Out-Null

Start-Sleep -Seconds $ShortPauseSeconds
Restore-DumbLab -Mode "Sabotaje"

End-Test -Tec "TEC-008"

# =========================
# TEC-009 — Staging / Archive / Exfil contextual
# =========================

Start-Test -Tec "TEC-009" -Name "Staging / Archive / Exfiltration contextual"

Restore-DumbLab -Mode "Exfiltracion"
Assert-DumbLabSafe

$StagingDir = Join-Path $ExfilDir "datos_robados"
New-Item -ItemType Directory -Force -Path $StagingDir | Out-Null

$Zip1 = Join-Path $StagingDir "dump_exfil.zip"
$Zip2 = Join-Path $StagingDir "dump_exfil_copy.zip"

Remove-Item $Zip1 -Force -ErrorAction SilentlyContinue
Remove-Item $Zip2 -Force -ErrorAction SilentlyContinue

$tec009Command = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$src = ''{0}\*''; $zip1 = ''{1}''; $zip2 = ''{2}''; Write-Host ''TEC009_OK''; Compress-Archive -Path $src -DestinationPath $zip1 -Force; Copy-Item $zip1 $zip2 -Force; Start-Sleep 3"' -f $DumbLab, $Zip1, $Zip2

Invoke-CmdLine -Command $tec009Command -Description "TEC-009: compresion y copia local para staging" | Out-Null

Start-Sleep -Seconds $ShortPauseSeconds
Restore-DumbLab -Mode "Exfiltracion"

End-Test -Tec "TEC-009"

# =========================
# RESTAURACION FINAL
# =========================

Write-Step "Restauracion final de laboratorio"
Restore-DumbLab -Mode "All"

Write-Step "Campaña finalizada correctamente"
Write-Step "Log final: $LogFile"

Write-Host ""
Write-Host "Pruebas finalizadas. Log: $LogFile"
