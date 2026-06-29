[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$TransferId,
    [string]$SourceRoot = 'C:\Users\julio\Desktop\TFM\Carpeta_Compartida_TFM\controlled_rerun_results'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$packageRoot = $PSScriptRoot
. (Join-Path $packageRoot 'lib\ControlledRerun.Common.ps1')

function Test-CRDirectoryIdentical {
    param(
        [Parameter(Mandatory=$true)][string]$Left,
        [Parameter(Mandatory=$true)][string]$Right
    )
    if(-not (Test-Path -LiteralPath $Left -PathType Container) -or -not (Test-Path -LiteralPath $Right -PathType Container)){
        return $false
    }
    $leftRoot = $Left.TrimEnd('\')
    $rightRoot = $Right.TrimEnd('\')
    $leftFiles = @(Get-ChildItem -LiteralPath $leftRoot -Recurse -File | ForEach-Object {
        [pscustomobject]@{RelativePath=$_.FullName.Substring($leftRoot.Length).TrimStart('\');Length=$_.Length;SHA256=Get-CRSha256 -Path $_.FullName}
    })
    $rightFiles = @(Get-ChildItem -LiteralPath $rightRoot -Recurse -File | ForEach-Object {
        [pscustomobject]@{RelativePath=$_.FullName.Substring($rightRoot.Length).TrimStart('\');Length=$_.Length;SHA256=Get-CRSha256 -Path $_.FullName}
    })
    if($leftFiles.Count -ne $rightFiles.Count){return $false}
    $rightByPath = @{}
    foreach($record in $rightFiles){$rightByPath[[string]$record.RelativePath]=$record}
    foreach($record in $leftFiles){
        $key=[string]$record.RelativePath
        if(-not $rightByPath.ContainsKey($key)){return $false}
        $other=$rightByPath[$key]
        if([long]$record.Length -ne [long]$other.Length -or [string]$record.SHA256 -ne [string]$other.SHA256){return $false}
    }
    return $true
}

function Merge-CRDirectoryNoOverwrite {
    param(
        [Parameter(Mandatory=$true)][string]$Source,
        [Parameter(Mandatory=$true)][string]$Destination
    )
    $sourceRoot = $Source.TrimEnd('\')
    foreach($file in @(Get-ChildItem -LiteralPath $sourceRoot -Recurse -File)){
        $relative = $file.FullName.Substring($sourceRoot.Length).TrimStart('\')
        $target = Join-Path $Destination $relative
        if(Test-Path -LiteralPath $target -PathType Leaf){
            if((Get-CRSha256 -Path $file.FullName) -ne (Get-CRSha256 -Path $target)){
                throw "Conflicto de hash; no se sobrescribe: $target"
            }
            continue
        }
        $parent = Split-Path -Parent $target
        New-CRDirectory -Path $parent | Out-Null
        Copy-Item -LiteralPath $file.FullName -Destination $target
    }
}

$source = Join-Path $SourceRoot $TransferId
$manifest = Join-Path $source 'TRANSFER_HASHES_SHA256.csv'
if (-not (Test-Path -LiteralPath $manifest -PathType Leaf)) {
    throw "Transferencia sin manifest: $manifest"
}
foreach($record in @(Import-Csv -LiteralPath $manifest)) {
    $file = Join-Path $source $record.RelativePath
    if ((Get-CRSha256 -Path $file) -ne [string]$record.SHA256) {
        throw "Hash de transferencia incorrecto: $($record.RelativePath)"
    }
}
$output = New-CRDirectory -Path (Join-Path $packageRoot 'OUTPUT')
foreach($section in @('TEC','FP','BENCHMARK','PREFLIGHT')) {
    $sectionSource = Join-Path $source $section
    if (-not (Test-Path -LiteralPath $sectionSource -PathType Container)) { continue }
    $sectionDestination = New-CRDirectory -Path (Join-Path $output $section)
    foreach($entry in @(Get-ChildItem -LiteralPath $sectionSource -Force)){
        $target = Join-Path $sectionDestination $entry.Name
        if($entry.PSIsContainer){
            if(Test-Path -LiteralPath $target){
                if(Test-CRDirectoryIdentical -Left $entry.FullName -Right $target){
                    Write-CRStatus -Message ("Ya importado e idéntico; se conserva: {0}/{1}" -f $section,$entry.Name) -Level OK
                    continue
                }
                Merge-CRDirectoryNoOverwrite -Source $entry.FullName -Destination $target
                Write-CRStatus -Message ("Campaña existente ampliada solo con ficheros nuevos: {0}/{1}" -f $section,$entry.Name) -Level OK
                continue
            }
            Copy-Item -LiteralPath $entry.FullName -Destination $sectionDestination -Recurse
            Write-CRStatus -Message ("Importado: {0}/{1}" -f $section,$entry.Name) -Level OK
        } else {
            if(Test-Path -LiteralPath $target){
                if((Get-CRSha256 -Path $entry.FullName) -eq (Get-CRSha256 -Path $target)){continue}
                throw "Existe un fichero distinto y no se sobrescribe: $target"
            }
            Copy-Item -LiteralPath $entry.FullName -Destination $target
        }
    }
}
Write-CRStatus -Message "Resultados VM importados y verificados en $output" -Level OK
