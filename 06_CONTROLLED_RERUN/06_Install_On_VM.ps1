[CmdletBinding()]
param(
    [string]$VmProjectRoot = 'C:\Users\seguridad\Desktop\TFM',
    [switch]$AllowPackageUpdate
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$packageInBundle = $PSScriptRoot
$bundleRoot = Split-Path -Parent $packageInBundle
. (Join-Path $packageInBundle 'lib\ControlledRerun.Common.ps1')
Assert-CRAdministrator
$patchStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$packageBackupRoot = Join-Path $VmProjectRoot ("06_CONTROLLED_RERUN\PATCH_BACKUPS\{0}" -f $patchStamp)

$bundleHashes = Join-Path $bundleRoot 'BUNDLE_HASHES_SHA256.csv'
if (-not (Test-Path -LiteralPath $bundleHashes -PathType Leaf)) {
    throw "No existe el manifest del bundle: $bundleHashes"
}
foreach($record in @(Import-Csv -LiteralPath $bundleHashes)) {
    $source = Join-Path $bundleRoot $record.RelativePath
    if ((Get-CRSha256 -Path $source) -ne [string]$record.SHA256) {
        throw "Bundle corrupto o incompleto: $($record.RelativePath)"
    }
}

function Install-RelativeFile {
    param([string]$RelativePath)
    $source = Join-Path $bundleRoot $RelativePath
    $destination = Join-Path $VmProjectRoot $RelativePath
    if (Test-Path -LiteralPath $destination -PathType Leaf) {
        if ((Get-CRSha256 -Path $destination) -ne (Get-CRSha256 -Path $source)) {
            $isPackageFile = $RelativePath.StartsWith('06_CONTROLLED_RERUN\',[System.StringComparison]::OrdinalIgnoreCase)
            if (-not $AllowPackageUpdate -or -not $isPackageFile) {
                throw "Existe un fichero distinto y no se sobrescribirá sin -AllowPackageUpdate: $destination"
            }
            $insidePackage = $RelativePath.Substring('06_CONTROLLED_RERUN\'.Length)
            $backup = Join-Path $packageBackupRoot $insidePackage
            New-CRDirectory -Path (Split-Path -Parent $backup) | Out-Null
            Copy-Item -LiteralPath $destination -Destination $backup -Force
            Copy-Item -LiteralPath $source -Destination $destination -Force
            Write-CRStatus -Message "Actualizado con backup: $RelativePath" -Level WARN
        }
        return
    }
    New-CRDirectory -Path (Split-Path -Parent $destination) | Out-Null
    Copy-Item -LiteralPath $source -Destination $destination
}

foreach($file in @(Get-ChildItem -LiteralPath (Join-Path $bundleRoot '06_CONTROLLED_RERUN') -Recurse -File)) {
    $relative = $file.FullName.Substring($bundleRoot.Length).TrimStart('\')
    Install-RelativeFile -RelativePath $relative
}
foreach($definition in @(Get-CRConfigDefinitions)) {
    Install-RelativeFile -RelativePath $definition.RelativePath
}
Write-CRStatus -Message "Paquete instalado/verificado en $VmProjectRoot. No se ha escrito en 01_ARTIFACTS\validated." -Level OK
if ($AllowPackageUpdate) {
    Write-CRStatus -Message "Backups de package previos, si hubo cambios: $packageBackupRoot" -Level INFO
}
