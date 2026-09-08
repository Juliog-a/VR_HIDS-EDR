[CmdletBinding()]
param(
    [string]$ProjectRoot = 'C:\Users\seguridad\Desktop\TFM'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$runner = Join-Path $ProjectRoot '03_RUNNERS\validate\TFM_Run_All_TEC_Tests_v6_PATCHED_4104.ps1'
$expectedHash = 'C023F8F7DC4D90977C4F6C76C8147A90C3558D62E14D9E1A265358DCDFF68367'
if (-not (Test-Path -LiteralPath $runner -PathType Leaf)) {
    throw "No existe runner patched: $runner"
}
$actualHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $runner).Hash
if ($actualHash -ne $expectedHash) {
    throw "Hash runner patched incorrecto. Expected=$expectedHash Actual=$actualHash"
}

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($runner,[ref]$tokens,[ref]$errors)
if ($errors.Count -gt 0) {
    throw "Runner patched no parseable: $($errors.Message -join '; ')"
}
$requiredFunctions = @('Get-ScriptBlockIdFromMessage','Write-TEC009PowerShell4104Summary')
foreach ($name in $requiredFunctions) {
    $functionAst = @($ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
    },$true))
    if ($functionAst.Count -ne 1) {
        throw "No se encontró exactamente una función $name"
    }
    Invoke-Expression $functionAst[0].Extent.Text
}

function Write-Step {
    param([string]$Message,[string]$Level='INFO')
    Write-Host ("[{0}] {1}" -f $Level,$Message)
}

$script:Mock4104Events = @(
    foreach ($i in 1..49) {
        [pscustomobject]@{
            TimeCreated = (Get-Date).AddSeconds(-$i)
            Id = 4104
            Message = "Invoke-WebRequest -Uri http://127.0.0.1:8088/upload -Method POST -InFile file.zip -ContentType application/zip ScriptBlock ID: 00000000-0000-0000-0000-$('{0:D12}' -f $i)"
        }
    }
)

function Get-WinEvent {
    [CmdletBinding()]
    param([hashtable]$FilterHashtable)
    return $script:Mock4104Events
}

$script:CurrentTechnique = [pscustomobject]@{TEC_ID='TEC-009'}
$result = Write-TEC009PowerShell4104Summary -RunStart (Get-Date).AddMinutes(-5)
if ([string]$result.QueryStatus -ne 'OK') {
    throw "QueryStatus no OK: $($result.QueryStatus) $($result.QueryErrorFull)"
}
if ([int]$result.RawCount -ne 49 -or [int]$result.MatchingCount -ne 49 -or [int]$result.Count -ne 49) {
    throw "Conteos incorrectos: Raw=$($result.RawCount) Matching=$($result.MatchingCount) Count=$($result.Count)"
}
if (@($result.Samples).Count -ne 15) {
    throw "Samples normalizadas incorrectas: $(@($result.Samples).Count)"
}

$techniqueList = New-Object System.Collections.Generic.List[object]
foreach ($i in 1..9) {
    [void]$techniqueList.Add([pscustomobject]@{TEC_ID=('TEC-{0:D3}' -f $i);Status='OK'})
}
$summaryProbe = [ordered]@{Techniques=$techniqueList.ToArray()}
$roundTrip = $summaryProbe | ConvertTo-Json -Depth 4 | ConvertFrom-Json
if (@($roundTrip.Techniques).Count -ne 9) {
    throw 'La prueba de serialización del summary no conserva nueve técnicas.'
}

Write-Host '[OK] PATCHED_4104 SELFTEST PASS: Raw=49 Matching=49 Samples=15 SummaryTechniques=9' -ForegroundColor Green
