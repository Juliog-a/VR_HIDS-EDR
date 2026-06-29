[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$packageRoot = $PSScriptRoot
$projectRoot = Split-Path -Parent $packageRoot
. (Join-Path $packageRoot 'lib\ControlledRerun.Common.ps1')

$failures = @()
$psFiles = @(Get-ChildItem -LiteralPath $packageRoot -Recurse -Filter '*.ps1' -File)
foreach($file in $psFiles){
    $tokens = $null
    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($file.FullName,[ref]$tokens,[ref]$errors)
    if($errors.Count -gt 0){
        $failures += "Parser $($file.Name): $($errors.Message -join '; ')"
    }
}

$hashChecks = @(Test-CRConfigHashes -ProjectRoot $projectRoot)
foreach($check in $hashChecks){
    if(-not $check.Match){$failures += "Hash $($check.Role): $($check.RelativePath)"}
}

$csvManifest = @(Import-Csv -LiteralPath (Join-Path $packageRoot 'CONFIG_SOURCE_MANIFEST.csv'))
foreach($record in $csvManifest){
    $definition = @($hashChecks | Where-Object {$_.Role -eq $record.Role})
    if($definition.Count -ne 1 -or [string]$definition[0].ExpectedSHA256 -ne [string]$record.SHA256){
        $failures += "Manifest inconsistente: $($record.Role)"
    }
}

$testMatrix = Join-Path $projectRoot '00_CONTEXT\TEST_MATRIX.md'
if(-not (Test-Path -LiteralPath $testMatrix -PathType Leaf)){
    $failures += 'Falta 00_CONTEXT/TEST_MATRIX.md'
} else {
    $matrixText = Get-Content -LiteralPath $testMatrix -Raw -Encoding UTF8
    foreach($tec in (1..9 | ForEach-Object {'TEC-{0:D3}' -f $_})){
        if($matrixText -notmatch [regex]::Escape($tec)){$failures += "TEST_MATRIX sin $tec"}
    }
}

foreach($required in @('README.md','01_Run_TEC_Canonical.ps1','02_Run_FP_Controlled.ps1','03_Run_Benchmark_3x.ps1','04_Validate_Controlled_Rerun.ps1','10_Build_ClientEvent_Exports_v2.ps1','11_Verify_ClientEvent_Exports.ps1')){
    if(-not (Test-Path -LiteralPath (Join-Path $packageRoot $required) -PathType Leaf)){$failures += "Falta $required"}
}

if($failures.Count -gt 0){
    $failures | ForEach-Object { Write-CRStatus -Message $_ -Level FAIL }
    exit 2
}
Write-CRStatus -Message ("SELFTEST PASS: {0} scripts parseables, {1} hashes fuente y TEST_MATRIX 9/9." -f $psFiles.Count,$hashChecks.Count) -Level OK
