param(
    [string]$HostAddress = "0.0.0.0",
    [int]$Port = 8000,
    [string]$OutputDir = "",
    [int]$MaxSeconds = 600,
    [int]$MaxUploads = 5,
    [int]$MaxBytes = 104857600,
    [string]$PythonExe = "python.exe",
    [string]$ReceiverScript = "",
    [string]$PidFile = "",
    [string]$LogFile = "",
    [string]$ProcessStdoutLog = "",
    [string]$ProcessStderrLog = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($ReceiverScript)) {
    $ReceiverScript = Join-Path $PSScriptRoot "receiver_tfm_v4.py"
}
if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $OutputDir = Join-Path $PSScriptRoot "received"
}
if ([string]::IsNullOrWhiteSpace($LogFile)) {
    $LogFile = Join-Path $OutputDir "receiver_log.jsonl"
}
if ([string]::IsNullOrWhiteSpace($PidFile)) {
    $PidFile = Join-Path $OutputDir "receiver_tfm_v4.pid"
}
if ([string]::IsNullOrWhiteSpace($ProcessStdoutLog)) {
    $ProcessStdoutLog = Join-Path $OutputDir "receiver_process_v4_stdout.log"
}
if ([string]::IsNullOrWhiteSpace($ProcessStderrLog)) {
    $ProcessStderrLog = Join-Path $OutputDir "receiver_process_v4_stderr.log"
}

if (-not (Test-Path -LiteralPath $ReceiverScript)) {
    throw "Receiver script not found: $ReceiverScript"
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $PidFile) | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $LogFile) | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $ProcessStdoutLog) | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $ProcessStderrLog) | Out-Null

$argList = @(
    "`"$ReceiverScript`"",
    "--host", $HostAddress,
    "--port", "$Port",
    "--out-dir", "`"$OutputDir`"",
    "--log-file", "`"$LogFile`"",
    "--max-seconds", "$MaxSeconds",
    "--max-uploads", "$MaxUploads",
    "--max-bytes", "$MaxBytes"
)

$process = Start-Process `
    -FilePath $PythonExe `
    -ArgumentList $argList `
    -PassThru `
    -WindowStyle Hidden `
    -RedirectStandardOutput $ProcessStdoutLog `
    -RedirectStandardError $ProcessStderrLog

$process.Id | Set-Content -LiteralPath $PidFile -Encoding ASCII

$ips = @()
try {
    $ips = Get-NetIPAddress -AddressFamily IPv4 |
        Where-Object { $_.IPAddress -notmatch '^(127\.|169\.254\.)' } |
        Select-Object -ExpandProperty IPAddress
}
catch {
    $ips = @()
}
if (-not $ips -or $ips.Count -lt 1) {
    $ips = @((Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled = True").IPAddress |
        ForEach-Object { $_ } |
        Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' -and $_ -notmatch '^(127\.|169\.254\.)' })
}

Write-Host "Receiver v4 PID: $($process.Id)"
Write-Host "PID file: $PidFile"
Write-Host "Output dir: $OutputDir"
Write-Host "JSONL log file: $LogFile"
Write-Host "Process stdout log file: $ProcessStdoutLog"
Write-Host "Process stderr log file: $ProcessStderrLog"
Write-Host "Listening on: $HostAddress`:$Port"
Write-Host "Health URL local: http://127.0.0.1:$Port/health"
foreach ($ip in $ips) {
    Write-Host "Health URL: http://$ip`:$Port/health"
    Write-Host "Upload URL: http://$ip`:$Port/upload"
}
Write-Host "Auto-stop: max_seconds=$MaxSeconds max_uploads=$MaxUploads max_bytes=$MaxBytes"
