[CmdletBinding()]
param(
    [string]$DestinationRoot = "${PSScriptRoot}\..\Carpeta_Compartida_TFM\controlled_rerun_bundle"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$packageRoot = $PSScriptRoot
$projectRoot = Split-Path -Parent $packageRoot
. (Join-Path $packageRoot 'lib\ControlledRerun.Common.ps1')

if ((Test-Path -LiteralPath $DestinationRoot) -and @(Get-ChildItem -LiteralPath $DestinationRoot -Force).Count -gt 0) {
    throw "El bundle ya existe. No se sobrescribe: $DestinationRoot"
}
New-CRDirectory -Path $DestinationRoot | Out-Null

function Copy-RelativeFile {
    param([string]$RelativePath)
    $source = Join-Path $projectRoot $RelativePath
    $destination = Join-Path $DestinationRoot $RelativePath
    New-CRDirectory -Path (Split-Path -Parent $destination) | Out-Null
    Copy-Item -LiteralPath $source -Destination $destination -Force
}

foreach($file in @(Get-ChildItem -LiteralPath $packageRoot -Recurse -File | Where-Object { $_.FullName -notmatch '\\OUTPUT\\' -and $_.FullName -notmatch '\\_audit_tmp\\' })) {
    $relativeInsidePackage = $file.FullName.Substring($packageRoot.Length).TrimStart('\')
    Copy-RelativeFile -RelativePath (Join-Path '06_CONTROLLED_RERUN' $relativeInsidePackage)
}
foreach($definition in @(Get-CRConfigDefinitions)) {
    Copy-RelativeFile -RelativePath $definition.RelativePath
}

Export-CRHashManifest -Root $DestinationRoot -OutputPath (Join-Path $DestinationRoot 'BUNDLE_HASHES_SHA256.csv')
Write-CRStatus -Message "Bundle VM preparado sin modificar validated: $DestinationRoot" -Level OK
