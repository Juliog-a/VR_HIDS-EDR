[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$CampaignDirectory,
    [string]$SourceDirectory = '',
    [string]$OutputDirectory = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$packageRoot = $PSScriptRoot
. (Join-Path $packageRoot 'lib\ControlledRerun.Common.ps1')

function Get-RowValue {
    param($Row,[string]$Name)
    if($null -eq $Row){return $null}
    $property = $Row.PSObject.Properties[$Name]
    if($null -eq $property){return $null}
    return $property.Value
}

function Get-EventUtc {
    param($Row)
    $culture = [Globalization.CultureInfo]::InvariantCulture
    $number = 0.0
    $raw = [string](Get-RowValue -Row $Row -Name 'DetectionTime')
    if(-not [string]::IsNullOrWhiteSpace($raw) -and [double]::TryParse($raw,[Globalization.NumberStyles]::Float,$culture,[ref]$number)){
        return [DateTimeOffset]::FromUnixTimeMilliseconds([long][Math]::Round($number * 1000.0)).UtcDateTime
    }
    $rawTs = [string](Get-RowValue -Row $Row -Name '_ts')
    if(-not [string]::IsNullOrWhiteSpace($rawTs) -and [double]::TryParse($rawTs,[Globalization.NumberStyles]::Float,$culture,[ref]$number)){
        $milliseconds = if([Math]::Abs($number) -ge 100000000000){[long][Math]::Round($number)}else{[long][Math]::Round($number * 1000.0)}
        return [DateTimeOffset]::FromUnixTimeMilliseconds($milliseconds).UtcDateTime
    }
    foreach($name in @('Timestamp','EventTime','TimeCreated')){
        $candidate = [string](Get-RowValue -Row $Row -Name $name)
        $parsed = [DateTimeOffset]::MinValue
        if(-not [string]::IsNullOrWhiteSpace($candidate) -and [DateTimeOffset]::TryParse($candidate,$culture,[Globalization.DateTimeStyles]::AssumeUniversal,[ref]$parsed)){
            return $parsed.UtcDateTime
        }
    }
    return $null
}

$campaignDirectoryFull = [System.IO.Path]::GetFullPath($CampaignDirectory)
if(-not (Test-Path -LiteralPath $campaignDirectoryFull -PathType Container)){
    throw "No existe CampaignDirectory: $campaignDirectoryFull"
}
$manifestPath = @(
    Join-Path $campaignDirectoryFull 'campaign_manifest.json'
    Join-Path $campaignDirectoryFull 'repetition_manifest.json'
) | Where-Object {Test-Path -LiteralPath $_ -PathType Leaf} | Select-Object -First 1
if([string]::IsNullOrWhiteSpace([string]$manifestPath)){
    throw 'No existe campaign_manifest.json ni repetition_manifest.json.'
}
$manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$startUtc = [DateTimeOffset]::Parse([string]$manifest.StartUtc).UtcDateTime
$endUtc = [DateTimeOffset]::Parse([string]$manifest.EndUtc).UtcDateTime
if($endUtc -le $startUtc){throw 'Ventana UTC inválida en el manifest.'}

if([string]::IsNullOrWhiteSpace($OutputDirectory)){$OutputDirectory=Join-Path $campaignDirectoryFull 'velociraptor_exports'}
$outputDirectoryFull = New-CRDirectory -Path ([System.IO.Path]::GetFullPath($OutputDirectory))

$sourceRoots = @()
if(-not [string]::IsNullOrWhiteSpace($SourceDirectory)){
    $sourceRoots = @([System.IO.Path]::GetFullPath($SourceDirectory))
} else {
    $sourceRoots = @($campaignDirectoryFull)
    $rawSources = Join-Path $campaignDirectoryFull 'velociraptor_exports\raw_sources'
    if(Test-Path -LiteralPath $rawSources -PathType Container){$sourceRoots += $rawSources}
}
foreach($root in $sourceRoots){if(-not (Test-Path -LiteralPath $root -PathType Container)){throw "No existe SourceDirectory: $root"}}

$profiles = @(
    [pscustomobject]@{Profile='P1_CRITICAL';Pattern='^P1_EVENT_';Artifact='Custom.TFM.HIDS.P1.Critical.Priority.Event_v1';Output='P1_CRITICAL_CLIENT_EVENT.csv'},
    [pscustomobject]@{Profile='P2_EVENT';Pattern='^Forensic_EVENT_';Artifact='Custom.TFM.HIDS.P2.High.Forensic.Event_v1';Output='P2_EVENT_CLIENT_EVENT.csv'},
    [pscustomobject]@{Profile='P3_EVENT';Pattern='^TEC\d{3}_';Artifact='Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1';Output='P3_EVENT_CLIENT_EVENT.csv'},
    [pscustomobject]@{Profile='P4';Pattern='^CU\d{3}_';Artifact='Custom.TFM.HIDS.P4.Low.Basic_v2';Output='P4_CLIENT_EVENT.csv'}
)

$allCsv = @($sourceRoots | ForEach-Object {Get-ChildItem -LiteralPath $_ -Recurse -File -Filter '*.csv'}) |
    Sort-Object FullName -Unique
$candidateSourceCsv = @($allCsv | Where-Object {$_.BaseName -match '^(P1_EVENT_|Forensic_EVENT_|TEC\d{3}_|CU\d{3}_)' -and $_.Name -notin @('P1_CRITICAL_CLIENT_EVENT.csv','P2_EVENT_CLIENT_EVENT.csv','P3_EVENT_CLIENT_EVENT.csv','P4_CLIENT_EVENT.csv')})
$latestInputUtc = [datetime]::MinValue
foreach($file in $candidateSourceCsv){
    foreach($row in @(Import-Csv -LiteralPath $file.FullName)){
        $eventUtc = Get-EventUtc -Row $row
        if($null -ne $eventUtc -and $eventUtc -gt $latestInputUtc){$latestInputUtc=$eventUtc}
    }
}
if($latestInputUtc -eq [datetime]::MinValue -or $latestInputUtc -lt $startUtc){
    throw ("Los exports fuente no acreditan datos posteriores al inicio de campaña. latest_input={0}; campaign_start={1}. Exporte CLIENT_EVENT después de finalizar la campaña." -f $(if($latestInputUtc -eq [datetime]::MinValue){'NONE'}else{$latestInputUtc.ToString('o')}),$startUtc.ToString('o'))
}
$records = New-Object System.Collections.ArrayList
foreach($profile in $profiles){
    $outputPath = Join-Path $outputDirectoryFull $profile.Output
    if(Test-Path -LiteralPath $outputPath){throw "La salida ya existe y no se sobrescribe: $outputPath"}
    $sourceFiles = @($allCsv | Where-Object {$_.BaseName -match $profile.Pattern -and $_.Name -ne $profile.Output})
    if($sourceFiles.Count -eq 0){throw "No hay exports por source para $($profile.Profile). Exporte al menos un CSV original, aunque no tenga filas."
    }
    $outRows = New-Object System.Collections.ArrayList
    $inputRows = 0
    $unparseable = 0
    $outside = 0
    $firstHeader = ''
    foreach($file in $sourceFiles){
        if([string]::IsNullOrWhiteSpace($firstHeader)){$firstHeader=[string](Get-Content -LiteralPath $file.FullName -TotalCount 1)}
        $sourceName = $file.BaseName -replace '-\d{4}-\d{2}-\d{2}T.*$',''
        foreach($row in @(Import-Csv -LiteralPath $file.FullName)){
            $inputRows++
            $eventUtc = Get-EventUtc -Row $row
            if($null -eq $eventUtc){$unparseable++;continue}
            if($eventUtc -lt $startUtc -or $eventUtc -gt $endUtc){$outside++;continue}
            $ordered = [ordered]@{Artifact=$profile.Artifact;Source=$sourceName}
            foreach($property in $row.PSObject.Properties){
                if(-not $ordered.Contains($property.Name)){$ordered[$property.Name]=$property.Value}
            }
            [void]$outRows.Add([pscustomobject]$ordered)
        }
    }
    if($outRows.Count -gt 0){
        $outRows.ToArray() | Export-Csv -LiteralPath $outputPath -NoTypeInformation -Encoding UTF8
    } else {
        if([string]::IsNullOrWhiteSpace($firstHeader)){throw "No se pudo obtener cabecera CSV para $($profile.Profile)."}
        Write-CRUtf8NoBom -Path $outputPath -Text ("Artifact,Source,{0}`r`n" -f $firstHeader)
    }
    [void]$records.Add([pscustomobject]@{
        Profile=$profile.Profile
        Artifact=$profile.Artifact
        OutputFile=$profile.Output
        SourceFiles=$sourceFiles.Count
        InputRows=$inputRows
        RowsInWindow=$outRows.Count
        RowsOutsideWindow=$outside
        UnparseableTimestamps=$unparseable
        StartUtc=$startUtc.ToString('o')
        EndUtc=$endUtc.ToString('o')
        SHA256=Get-CRSha256 -Path $outputPath
    })
    Write-CRStatus -Message ("{0}: sources={1}; in_window={2}; outside={3}; bad_time={4}" -f $profile.Profile,$sourceFiles.Count,$outRows.Count,$outside,$unparseable) -Level OK
}

$summaryCsv = Join-Path $outputDirectoryFull 'CLIENT_EVENT_EXPORT_BUILD.csv'
$records.ToArray() | Export-Csv -LiteralPath $summaryCsv -NoTypeInformation -Encoding UTF8
Write-CRJson -Object ([ordered]@{
    SchemaVersion='1.0'
    Manifest=$manifestPath
    StartUtc=$startUtc.ToString('o')
    EndUtc=$endUtc.ToString('o')
    Profiles=$records.ToArray()
}) -Path (Join-Path $outputDirectoryFull 'CLIENT_EVENT_EXPORT_BUILD.json') -Depth 8
Write-CRStatus -Message "Exports CLIENT_EVENT construidos sin modificar fuentes: $outputDirectoryFull" -Level OK
