[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$packageRoot = Split-Path -Parent $PSScriptRoot
$projectRoot = Split-Path -Parent $packageRoot
$runner = Join-Path $projectRoot '03_RUNNERS\TFM_Run_All_TEC_Tests_v6_PATCHED_4104.ps1'
$basePath = Join-Path $projectRoot '01_ACTIVE_TESTS\Pruebas'
if (-not (Test-Path -LiteralPath $runner -PathType Leaf)) {
    throw "No existe runner TEC v6 patched: $runner"
}
if (-not (Test-Path -LiteralPath $basePath -PathType Container)) {
    throw "No existe BasePath: $basePath"
}

# Carga de trabajo para benchmark. La exfiltración real pertenece exclusivamente
# a la campaña TEC canónica y se desactiva aquí para no mezclar evidencias.
& $runner `
    -BasePath $basePath `
    -EnableExfilUpload $false `
    -KeepExfilArtifacts $true `
    -PauseBetweenTestsSeconds 10 `
    -ShortPauseSeconds 2
