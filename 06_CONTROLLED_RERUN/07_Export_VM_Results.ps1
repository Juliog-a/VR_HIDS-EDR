[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$TransferId,
    [string]$TecCampaignId = '',
    [string]$FpCampaignId = '',
    [string]$BenchmarkCampaignId = '',
    [string]$DestinationRoot = '\\VBOXSVR\Carpeta_Compartida_TFM\controlled_rerun_results'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$packageRoot = $PSScriptRoot
. (Join-Path $packageRoot 'lib\ControlledRerun.Common.ps1')
$sourceRoot = Join-Path $packageRoot 'OUTPUT'
$destination = Join-Path $DestinationRoot $TransferId
if (-not (Test-Path -LiteralPath $sourceRoot -PathType Container)) {
    throw "No existen resultados locales: $sourceRoot"
}
if ((Test-Path -LiteralPath $destination) -and @(Get-ChildItem -LiteralPath $destination -Force).Count -gt 0) {
    throw "El destino ya contiene datos y no se sobrescribe: $destination"
}

$ids = @($TecCampaignId,$FpCampaignId,$BenchmarkCampaignId)
$selectedCount = @($ids | Where-Object {-not [string]::IsNullOrWhiteSpace($_)}).Count
if($selectedCount -notin @(0,3)){
    throw 'Para una transferencia selectiva debe indicar TecCampaignId, FpCampaignId y BenchmarkCampaignId.'
}
foreach($id in @($ids | Where-Object {-not [string]::IsNullOrWhiteSpace($_)})){
    if([System.IO.Path]::GetFileName($id) -ne $id -or $id -match '[\\/]'){
        throw "Identificador de campaña no válido: $id"
    }
}

New-CRDirectory -Path $destination | Out-Null
if($selectedCount -eq 3){
    $selection = @(
        [pscustomobject]@{Section='TEC';CampaignId=$TecCampaignId},
        [pscustomobject]@{Section='FP';CampaignId=$FpCampaignId},
        [pscustomobject]@{Section='BENCHMARK';CampaignId=$BenchmarkCampaignId}
    )
    foreach($item in $selection){
        $source = Join-Path (Join-Path $sourceRoot $item.Section) $item.CampaignId
        if(-not (Test-Path -LiteralPath $source -PathType Container)){
            throw "No existe la campaña seleccionada: $source"
        }
        $sectionDestination = New-CRDirectory -Path (Join-Path $destination $item.Section)
        Copy-Item -LiteralPath $source -Destination $sectionDestination -Recurse
        Write-CRStatus -Message ("Seleccionada {0}/{1}" -f $item.Section,$item.CampaignId) -Level OK
    }
    $preflight = Join-Path $sourceRoot 'PREFLIGHT'
    if(Test-Path -LiteralPath $preflight -PathType Container){
        Copy-Item -LiteralPath $preflight -Destination $destination -Recurse
    }
} else {
    Write-CRStatus -Message 'Transferencia completa de secciones: puede incluir campañas históricas. Use los tres CampaignId para una entrega final aislada.' -Level WARN
    foreach($section in @('PREFLIGHT','TEC','FP','BENCHMARK')) {
        $source = Join-Path $sourceRoot $section
        if (Test-Path -LiteralPath $source -PathType Container) {
            Copy-Item -LiteralPath $source -Destination $destination -Recurse
        }
    }
}
Export-CRHashManifest -Root $destination -OutputPath (Join-Path $destination 'TRANSFER_HASHES_SHA256.csv')
Write-CRStatus -Message "Resultados copiados a: $destination" -Level OK
