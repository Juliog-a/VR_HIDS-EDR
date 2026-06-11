# TFM candidate v2 - Pre-enumeration for controlled TEC-007 simulation.
# Writes list files next to this script by default.
# Detection must not rely on this script name, lab paths or comments.

param(
    [string]$TargetPath = "C:\Users\seguridad\Desktop\TFM\Pruebas\DUMB_LAB",
    [string]$OutputDir = $PSScriptRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $TargetPath)) {
    throw "TargetPath does not exist: $TargetPath"
}

if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
}

$TxtFile = Join-Path $OutputDir "lista_archivos.txt"
$CsvFile = Join-Path $OutputDir "lista_archivos.csv"

$files = Get-ChildItem -LiteralPath $TargetPath -Recurse -File -ErrorAction Stop

$files |
    Select-Object Name, FullName, Length, LastWriteTime |
    Format-Table -AutoSize |
    Out-File -LiteralPath $TxtFile -Encoding UTF8

$files |
    Select-Object Name, FullName |
    Export-Csv -LiteralPath $CsvFile -NoTypeInformation -Encoding UTF8

Write-Host "TXT saved: $TxtFile"
Write-Host "CSV saved: $CsvFile"
Write-Host "Files enumerated: $($files.Count)"
