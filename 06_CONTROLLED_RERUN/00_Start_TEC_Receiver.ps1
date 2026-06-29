[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][ValidatePattern('^TEC_[A-Za-z0-9_-]+$')][string]$CampaignId,
    [string]$ProjectRoot = '',
    [string]$ListenAddress = '0.0.0.0',
    [int]$Port = 8088,
    [int]$MaxSeconds = 1800
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$packageRoot = $PSScriptRoot
. (Join-Path $packageRoot 'lib\ControlledRerun.Common.ps1')
$root = Get-CRProjectRoot -RequestedRoot $ProjectRoot -PackageRoot $packageRoot
$receiver = Join-Path $root '02_SCRIPTS\candidate\receiver_tfm_v4.py'
$expectedHash = 'D61A86270EA822D2D4195E8C506738D9544D7964C23CB174D9F09EE52DD84CF6'
if ((Get-CRSha256 -Path $receiver) -ne $expectedHash) {
    throw "Receiver ausente o con hash distinto: $receiver"
}

$campaignRoot = Join-Path $packageRoot ("OUTPUT\RECEIVER\{0}" -f $CampaignId)
if ((Test-Path -LiteralPath $campaignRoot) -and @(Get-ChildItem -LiteralPath $campaignRoot -Force).Count -gt 0) {
    throw "La carpeta del receiver ya contiene datos. Use otro CampaignId: $campaignRoot"
}
$receivedDir = New-CRDirectory -Path (Join-Path $campaignRoot 'received')
$logFile = Join-Path $campaignRoot 'receiver_log.jsonl'

$portInUse = Test-NetConnection -ComputerName '127.0.0.1' -Port $Port -InformationLevel Quiet -WarningAction SilentlyContinue
if ($portInUse) {
    throw "El puerto $Port ya está ocupado. No se inicia un receiver ambiguo."
}

$python = Get-Command py.exe -ErrorAction SilentlyContinue
$pythonArgs = @('-3')
if (-not $python) {
    $python = Get-Command python.exe -ErrorAction SilentlyContinue
    $pythonArgs = @()
}
if (-not $python) {
    throw 'No se encontró py.exe ni python.exe en el host.'
}

$metadata = [pscustomobject]@{
    CampaignId = $CampaignId
    ReceiverScript = $receiver
    ReceiverSHA256 = $expectedHash
    ListenAddress = $ListenAddress
    Port = $Port
    MaxSeconds = $MaxSeconds
    MaxUploads = 1
    StartedLocal = (Get-Date).ToString('o')
    StartedUtc = (Get-Date).ToUniversalTime().ToString('o')
}
Write-CRJson -Object $metadata -Path (Join-Path $campaignRoot 'receiver_manifest.json')

Write-CRStatus -Message "Receiver aislado para $CampaignId" -Level INFO
Write-CRStatus -Message "No cierre esta consola hasta recibir el ZIP o expirar el tiempo." -Level WARN
Write-CRStatus -Message "Salida: $campaignRoot" -Level INFO

$args = @($pythonArgs + @(
    $receiver,
    '--host', $ListenAddress,
    '--port', [string]$Port,
    '--out-dir', $receivedDir,
    '--log-file', $logFile,
    '--max-seconds', [string]$MaxSeconds,
    '--max-uploads', '1',
    '--max-bytes', '104857600',
    '--socket-timeout', '120'
))
& $python.Source @args
$exitCode = $LASTEXITCODE
Export-CRHashManifest -Root $campaignRoot -OutputPath (Join-Path $campaignRoot 'HASHES_SHA256.csv')
if ($exitCode -ne 0) {
    throw "El receiver terminó con código $exitCode"
}
Write-CRStatus -Message "Receiver finalizado. Conserve íntegra la carpeta: $campaignRoot" -Level OK
