[CmdletBinding()]
param(
    [string]$ProjectRoot = '',
    [string]$JsonlPath = '\\VBOXSVR\Carpeta_Compartida_TFM\inbox\soc_alerts.jsonl'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$packageRoot = $PSScriptRoot
. (Join-Path $packageRoot 'lib\ControlledRerun.Common.ps1')
$root = Get-CRProjectRoot -RequestedRoot $ProjectRoot -PackageRoot $packageRoot
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$output = New-CRDirectory -Path (Join-Path $packageRoot ("OUTPUT\PREFLIGHT\PREFLIGHT_{0}" -f $stamp))

$checks = New-Object System.Collections.Generic.List[object]
function Add-Check {
    param([string]$Check,[bool]$Pass,[string]$Evidence,[string]$RequiredAction='')
    [void]$checks.Add([pscustomobject]@{Check=$Check;Pass=$Pass;Evidence=$Evidence;RequiredAction=$RequiredAction})
    Write-CRStatus -Message ("{0}: {1}" -f $Check, $Evidence) -Level $(if($Pass){'OK'}else{'FAIL'})
}

Add-Check -Check 'ProjectRoot' -Pass (Test-Path -LiteralPath $root -PathType Container) -Evidence $root -RequiredAction 'Copiar el paquete dentro de la raíz TFM o usar -ProjectRoot.'

$hashChecks = @(Test-CRConfigHashes -ProjectRoot $root)
foreach ($item in $hashChecks) {
    Add-Check -Check ("Hash {0}" -f $item.Role) -Pass $item.Match -Evidence ("expected={0}; actual={1}; path={2}" -f $item.ExpectedSHA256,$item.ActualSHA256,$item.RelativePath) -RequiredAction 'Volver a copiar el fichero exacto indicado por CONFIG_SOURCE_MANIFEST.csv.'
}

$jsonlParent = Split-Path -Parent $JsonlPath
Add-Check -Check 'Directorio JSONL accesible' -Pass (Test-Path -LiteralPath $jsonlParent -PathType Container) -Evidence $jsonlParent -RequiredAction 'Montar la carpeta compartida VBOXSVR.'

$service = Get-Service -Name 'Velociraptor' -ErrorAction SilentlyContinue
Add-Check -Check 'Servicio Velociraptor existe' -Pass ($null -ne $service) -Evidence $(if($service){"$($service.Name) $($service.Status)"}else{'No encontrado'}) -RequiredAction 'Instalar o identificar el servicio cliente Velociraptor.'

$serviceRecord = Get-CimInstance Win32_Service -Filter "Name='Velociraptor'" -ErrorAction SilentlyContinue
$clientShape = $false
$serviceEvidence = 'No disponible'
if ($serviceRecord) {
    $serviceEvidence = [string]$serviceRecord.PathName
    $clientShape = ($serviceEvidence -match '(?i)client\.config\.yaml') -and ($serviceEvidence -match '(?i)service\s+run')
}
Add-Check -Check 'Servicio con forma CLIENT_SERVICE' -Pass $clientShape -Evidence $serviceEvidence -RequiredAction 'El servicio debe usar client.config.yaml y service run.'

$basePath = Join-Path $root '01_ACTIVE_TESTS\Pruebas'
Add-Check -Check 'BasePath activo' -Pass (Test-Path -LiteralPath $basePath -PathType Container) -Evidence $basePath -RequiredAction 'Copiar/sincronizar el árbol activo de pruebas en la VM.'

$allPass = (@($checks | Where-Object { -not $_.Pass }).Count -eq 0)
$report = [pscustomobject]@{
    SchemaVersion = '1.0'
    TimestampLocal = (Get-Date).ToString('o')
    TimestampUtc = (Get-Date).ToUniversalTime().ToString('o')
    ComputerName = $env:COMPUTERNAME
    ProjectRoot = $root
    JsonlPath = $JsonlPath
    AllAutomatedChecksPass = $allPass
    ManualChecksStillRequired = @(
        'P1/P2/P3/P4 están activos como Client Event Monitoring.',
        'SOC_v3 está activo como Server Event Monitoring.',
        'EnableJSONL=true, EnableDiscord=false y JsonlPath coincide.',
        'P1 legacy, SOC_v1 y SOC_v2 están desactivados.',
        'Reloj de VM y host sincronizado.'
    )
    Checks = $checks.ToArray()
}
Write-CRJson -Object $report -Path (Join-Path $output 'preflight_report.json')
$checks.ToArray() | Export-Csv -LiteralPath (Join-Path $output 'preflight_checks.csv') -NoTypeInformation -Encoding UTF8
Export-CRHashManifest -Root $output -OutputPath (Join-Path $output 'HASHES_SHA256.csv')

if (-not $allPass) {
    Write-CRStatus -Message "Preflight NO APTO. Revise $output" -Level FAIL
    exit 2
}
Write-CRStatus -Message "Preflight automático APTO. Faltan los controles manuales de GUI. Salida: $output" -Level OK
