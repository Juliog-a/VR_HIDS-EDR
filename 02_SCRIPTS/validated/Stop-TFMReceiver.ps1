param(
    [string]$PidFile = "",
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($PidFile)) {
    $PidFile = Join-Path $ProjectRoot "Pruebas\Receiver\received_controlled\receiver_tfm_v2.pid"
}

if (-not (Test-Path -LiteralPath $PidFile)) {
    Write-Host "PID file not found. Receiver may already be stopped: $PidFile"
    return
}

$pidText = (Get-Content -LiteralPath $PidFile -Raw).Trim()
if (-not ($pidText -match '^\d+$')) {
    throw "Invalid PID file content: $PidFile"
}

$receiverPid = [int]$pidText
$process = Get-Process -Id $receiverPid -ErrorAction SilentlyContinue
if ($null -eq $process) {
    Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue
    Write-Host "Process not found. Removed stale PID file: $PidFile"
    return
}

Stop-Process -Id $receiverPid -Force:$Force
Start-Sleep -Seconds 1
if (-not (Get-Process -Id $receiverPid -ErrorAction SilentlyContinue)) {
    Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue
    Write-Host "Receiver stopped. PID=$receiverPid"
}
else {
    Write-Warning "Receiver process still running. Re-run with -Force if needed. PID=$receiverPid"
}
