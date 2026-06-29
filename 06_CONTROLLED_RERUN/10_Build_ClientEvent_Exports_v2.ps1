[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$CampaignDirectory,
    [Parameter(Mandatory=$true)][string]$SourceDirectory,
    [string]$OutputDirectory = '',
    [switch]$ReplaceDerivedOutputs
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$packageRoot = $PSScriptRoot
. (Join-Path $packageRoot 'lib\ControlledRerun.Common.ps1')
Add-Type -AssemblyName Microsoft.VisualBasic

function Get-RowValue {
    param($Row,[string]$Name)
    if($null -eq $Row){return $null}
    $property=$Row.PSObject.Properties[$Name]
    if($null -eq $property){return $null}
    return $property.Value
}

function ConvertTo-UtcValue {
    param($Value)
    if($null -eq $Value){return $null}
    $raw=[string]$Value
    if([string]::IsNullOrWhiteSpace($raw)){return $null}
    $culture=[Globalization.CultureInfo]::InvariantCulture
    $number=0.0
    if([double]::TryParse($raw,[Globalization.NumberStyles]::Float,$culture,[ref]$number)){
        try{
            $milliseconds=if([Math]::Abs($number) -ge 100000000000){[long][Math]::Round($number)}else{[long][Math]::Round($number*1000.0)}
            return [DateTimeOffset]::FromUnixTimeMilliseconds($milliseconds).UtcDateTime
        }catch{return $null}
    }
    $parsed=[DateTimeOffset]::MinValue
    if([DateTimeOffset]::TryParse($raw,$culture,[Globalization.DateTimeStyles]::AssumeUniversal,[ref]$parsed)){
        return $parsed.UtcDateTime
    }
    return $null
}

function Get-EventTimeInfo {
    param($Row)
    foreach($name in @('DetectionTime','Timestamp','_ts','EventTime','TimeCreated')){
        $value=Get-RowValue -Row $Row -Name $name
        $utc=ConvertTo-UtcValue -Value $value
        if($null -ne $utc){return [pscustomobject]@{Utc=$utc;Field=$name;Raw=[string]$value}}
    }
    return $null
}

function Get-FallbackTimeInfo {
    param($Row)
    foreach($name in @('Timestamp','_ts','EventTime','TimeCreated')){
        $value=Get-RowValue -Row $Row -Name $name
        $utc=ConvertTo-UtcValue -Value $value
        if($null -ne $utc){return [pscustomobject]@{Utc=$utc;Field=$name;Raw=[string]$value}}
    }
    return $null
}

function Read-RobustCsv {
    param([Parameter(Mandatory=$true)][string]$Path)
    $parser=New-Object Microsoft.VisualBasic.FileIO.TextFieldParser -ArgumentList $Path,([System.Text.Encoding]::UTF8)
    $parser.TextFieldType=[Microsoft.VisualBasic.FileIO.FieldType]::Delimited
    $parser.SetDelimiters(',')
    $parser.HasFieldsEnclosedInQuotes=$true
    $parser.TrimWhiteSpace=$false
    $rows=New-Object System.Collections.ArrayList
    $malformed=New-Object System.Collections.ArrayList
    $duplicateHeaders=New-Object System.Collections.ArrayList
    try{
        if($parser.EndOfData){throw "CSV vacío: $Path"}
        $rawHeaders=@($parser.ReadFields())
        $headers=New-Object System.Collections.ArrayList
        $seen=@{}
        for($i=0;$i -lt $rawHeaders.Count;$i++){
            $name=([string]$rawHeaders[$i]).TrimStart([char]0xFEFF).Trim()
            if([string]::IsNullOrWhiteSpace($name)){$name="Column_$($i+1)"}
            $key=$name.ToLowerInvariant()
            if($seen.ContainsKey($key)){
                $seen[$key]++
                [void]$duplicateHeaders.Add($name)
                $name="{0}__DUP{1}" -f $name,$seen[$key]
            }else{$seen[$key]=1}
            [void]$headers.Add($name)
        }
        $logicalRow=1
        while(-not $parser.EndOfData){
            $logicalRow++
            try{$fields=@($parser.ReadFields())}
            catch [Microsoft.VisualBasic.FileIO.MalformedLineException]{
                [void]$malformed.Add([pscustomobject]@{LogicalRow=$logicalRow;Error=$_.Exception.Message;Raw=$parser.ErrorLine})
                continue
            }
            $ordered=[ordered]@{}
            for($i=0;$i -lt $headers.Count;$i++){
                $ordered[[string]$headers[$i]]=if($i -lt $fields.Count){[string]$fields[$i]}else{''}
            }
            if($fields.Count -gt $headers.Count){
                for($i=$headers.Count;$i -lt $fields.Count;$i++){$ordered["ExtraColumn_$($i+1)"]=[string]$fields[$i]}
            }
            [void]$rows.Add([pscustomobject]$ordered)
        }
        return [pscustomobject]@{
            Path=$Path
            Headers=$headers.ToArray()
            Rows=$rows.ToArray()
            Malformed=$malformed.ToArray()
            DuplicateHeaders=$duplicateHeaders.ToArray()
        }
    }finally{$parser.Close();$parser.Dispose()}
}

function ConvertTo-FlatText {
    param($Value,[string]$ColumnName='')
    if($null -eq $Value){return ''}
    $text=[string]$Value
    if($ColumnName -in @('EventDataJSON','SystemJSON')){
        try{$text=($text|ConvertFrom-Json -ErrorAction Stop|ConvertTo-Json -Compress -Depth 50)}catch{}
    }
    $text=$text -replace "`r`n|`n|`r",' '
    $text=$text -replace "`0",''
    return $text.Trim()
}

function ConvertTo-CsvField {
    param($Value,[string]$ColumnName='')
    $text=ConvertTo-FlatText -Value $Value -ColumnName $ColumnName
    return '"'+$text.Replace('"','""')+'"'
}

function Write-StrictCsv {
    param([string]$Path,[string[]]$Columns,[object[]]$Rows)
    $encoding=New-Object System.Text.UTF8Encoding($false)
    $writer=New-Object System.IO.StreamWriter($Path,$false,$encoding)
    try{
        $writer.WriteLine((@($Columns|ForEach-Object {ConvertTo-CsvField -Value $_}) -join ','))
        foreach($row in @($Rows)){
            $fields=foreach($column in $Columns){ConvertTo-CsvField -Value (Get-RowValue -Row $row -Name $column) -ColumnName $column}
            $writer.WriteLine((@($fields)-join ','))
        }
    }finally{$writer.Flush();$writer.Dispose()}
    try{$imported=@(Import-Csv -LiteralPath $Path -ErrorAction Stop)}catch{throw "Import-Csv no puede leer $Path : $($_.Exception.Message)"}
    if($imported.Count -ne @($Rows).Count){throw "Recuento Import-Csv incorrecto en $Path : expected=$(@($Rows).Count) actual=$($imported.Count)"}
    $headerCheck=Read-RobustCsv -Path $Path
    $headerNames=@($headerCheck.Headers)
    $headerUnique=@($headerNames|Sort-Object -Unique)
    if($headerNames.Count -ne $headerUnique.Count){throw "Cabecera duplicada generada en $Path"}
    return $imported.Count
}

function Get-CanonicalTechnique {
    param($Row)
    foreach($name in @('ID_Tecnica_Interna','TEC','TEC_ID')){
        $value=[string](Get-RowValue -Row $Row -Name $name)
        if($value -match '^TEC-00[1-9]$'){return $value}
    }
    return ''
}

function Get-FpAdjudication {
    param($Row,[datetime]$EventUtc,[object[]]$Tests)
    $existing=[string](Get-RowValue -Row $Row -Name 'FP_ID')
    if($existing -match '^FP-(00[1-9]|010)$'){
        return [pscustomobject]@{FP_ID=$matches[0].ToUpperInvariant();Status='EXACT';Basis='FP_ID column';WasUnknown=$false;Adjudicated=$false}
    }
    $parts=New-Object System.Collections.ArrayList
    foreach($property in $Row.PSObject.Properties){if($null -ne $property.Value){[void]$parts.Add([string]$property.Value)}}
    $text=$parts.ToArray() -join ' '
    $marker=[regex]::Match($text,'(?i)\bFP-(?:00[1-9]|010)\b')
    if($marker.Success){return [pscustomobject]@{FP_ID=$marker.Value.ToUpperInvariant();Status='ADJUDICATED';Basis='Explicit marker in CLIENT_EVENT';WasUnknown=$true;Adjudicated=$true}}
    $tec=Get-CanonicalTechnique -Row $Row
    $artifact=[string](Get-RowValue -Row $Row -Name 'Artifact')
    $source=[string](Get-RowValue -Row $Row -Name 'Source')
    if($tec -eq 'TEC-006' -or ($artifact+' '+$source) -match '(?i)SecuritySoftwareDiscovery|PowerShell4104_Strong_By_Technique'){
        return [pscustomobject]@{FP_ID='FP-006';Status='ADJUDICATED';Basis='TEC-006 security discovery within repetition';WasUnknown=$true;Adjudicated=$true}
    }
    $matching=@($Tests|Where-Object {$EventUtc -ge $_.StartUtc -and $EventUtc -le $_.EndUtc})
    if($matching.Count -eq 1){return [pscustomobject]@{FP_ID=[string]$matching[0].FP_ID;Status='ADJUDICATED';Basis='Exact runner test window';WasUnknown=$true;Adjudicated=$true}}
    return [pscustomobject]@{FP_ID='UNKNOWN';Status='UNRESOLVED';Basis='No explicit marker or unique test window';WasUnknown=$true;Adjudicated=$false}
}

function Get-MappedRawColumn {
    param([string]$Name,[hashtable]$Map,[System.Collections.ArrayList]$Columns)
    if($Map.ContainsKey($Name)){return [string]$Map[$Name]}
    $fixed=@('Artifact','Source','Profile','EventTimeUtc','TimestampFieldUsed','FP_ID','Adjudication','AdjudicationBasis')
    $candidate=if($Name -in $fixed){"Raw_$Name"}else{$Name}
    $base=$candidate
    $suffix=1
    while(@($Columns|Where-Object {$_ -ieq $candidate}).Count -gt 0){$suffix++;$candidate="${base}__${suffix}"}
    $Map[$Name]=$candidate
    [void]$Columns.Add($candidate)
    return $candidate
}

$campaignDirectoryFull=[System.IO.Path]::GetFullPath($CampaignDirectory)
$sourceDirectoryFull=[System.IO.Path]::GetFullPath($SourceDirectory)
if(-not(Test-Path -LiteralPath $campaignDirectoryFull -PathType Container)){throw "No existe CampaignDirectory: $campaignDirectoryFull"}
if(-not(Test-Path -LiteralPath $sourceDirectoryFull -PathType Container)){throw "No existe SourceDirectory: $sourceDirectoryFull"}
$manifestPath=Join-Path $campaignDirectoryFull 'repetition_manifest.json'
if(-not(Test-Path -LiteralPath $manifestPath -PathType Leaf)){
    $manifestPath=Join-Path $campaignDirectoryFull 'campaign_manifest.json'
}
if(-not(Test-Path -LiteralPath $manifestPath -PathType Leaf)){throw 'No existe repetition_manifest.json ni campaign_manifest.json en CampaignDirectory.'}
$manifest=Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8|ConvertFrom-Json
$startUtc=[DateTimeOffset]::Parse([string]$manifest.StartUtc).UtcDateTime
$endUtc=[DateTimeOffset]::Parse([string]$manifest.EndUtc).UtcDateTime
if($endUtc -le $startUtc){throw 'Ventana UTC inválida en el manifest.'}
$repetition=[string](Get-RowValue -Row $manifest -Name 'RepetitionName')
if([string]::IsNullOrWhiteSpace($repetition)){$repetition=[System.IO.Path]::GetFileName($campaignDirectoryFull)}
$isFpCampaign=([string](Get-RowValue -Row $manifest -Name 'Mode') -eq 'FP_CONTROLLED') -or (Test-Path -LiteralPath (Join-Path $campaignDirectoryFull 'repetition_manifest.json'))

$preflightCsv=@(Get-ChildItem -LiteralPath $sourceDirectoryFull -Recurse -File -Filter '*.csv'|Where-Object {$_.BaseName -match '^(P1_EVENT_|Forensic_EVENT_|TEC\d{3}_|CU\d{3}_)'})
if($preflightCsv.Count -eq 0){throw "SourceDirectory no contiene exports CLIENT_EVENT raw: $sourceDirectoryFull"}
$latestSourceUtc=[datetime]::MinValue
foreach($file in $preflightCsv){
    $preflightParsed=Read-RobustCsv -Path $file.FullName
    foreach($row in @($preflightParsed.Rows)){
        $timeInfo=Get-EventTimeInfo -Row $row
        if($null -ne $timeInfo -and $timeInfo.Utc -gt $latestSourceUtc){$latestSourceUtc=$timeInfo.Utc}
    }
}
if($latestSourceUtc -eq [datetime]::MinValue -or $latestSourceUtc -lt $startUtc){
    throw ("Los raw no acreditan eventos posteriores al inicio de {0}: latest={1}; start={2}" -f $repetition,$(if($latestSourceUtc -eq [datetime]::MinValue){'NONE'}else{$latestSourceUtc.ToString('o')}),$startUtc.ToString('o'))
}

if([string]::IsNullOrWhiteSpace($OutputDirectory)){$OutputDirectory=Join-Path $campaignDirectoryFull 'velociraptor_exports'}
$outputDirectoryFull=New-CRDirectory -Path ([System.IO.Path]::GetFullPath($OutputDirectory))
$derivedNames=@('P1_CRITICAL_CLIENT_EVENT.csv','P2_EVENT_CLIENT_EVENT.csv','P3_EVENT_CLIENT_EVENT.csv','P4_CLIENT_EVENT.csv','CLIENT_EVENT_FILTERING_REPORT.md','CLIENT_EVENT_FILTERING_REPORT.csv')
$existingDerived=@($derivedNames|ForEach-Object {Join-Path $outputDirectoryFull $_}|Where-Object {Test-Path -LiteralPath $_})
if($existingDerived.Count -gt 0){
    if(-not $ReplaceDerivedOutputs){throw "Ya existen salidas derivadas. Use -ReplaceDerivedOutputs para respaldarlas, nunca para tocar raw: $($existingDerived -join ', ')"}
    $backup=New-CRDirectory -Path (Join-Path $outputDirectoryFull ("previous_builds\BUILD_{0}" -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
    foreach($file in $existingDerived){Move-Item -LiteralPath $file -Destination (Join-Path $backup ([System.IO.Path]::GetFileName($file)))}
}

$tests=@()
$summaryFile=Get-ChildItem -LiteralPath (Join-Path $campaignDirectoryFull 'runner_native') -Filter '*_summary.json' -File -ErrorAction SilentlyContinue|Sort-Object LastWriteTime -Descending|Select-Object -First 1
if($null -ne $summaryFile){
    $runnerSummary=Get-Content -LiteralPath $summaryFile.FullName -Raw -Encoding UTF8|ConvertFrom-Json
    $tests=@($runnerSummary.Tests|ForEach-Object {[pscustomobject]@{FP_ID=[string]$_.FP_ID;StartUtc=[DateTimeOffset]::Parse([string]$_.StartTimeLocal).UtcDateTime;EndUtc=[DateTimeOffset]::Parse([string]$_.EndTimeLocal).UtcDateTime}})
}

$profiles=@(
    [pscustomobject]@{Profile='P1_CRITICAL';Pattern='^P1_EVENT_';Artifact='Custom.TFM.HIDS.P1.Critical.Priority.Event_v1';Output='P1_CRITICAL_CLIENT_EVENT.csv'},
    [pscustomobject]@{Profile='P2_EVENT';Pattern='^Forensic_EVENT_';Artifact='Custom.TFM.HIDS.P2.High.Forensic.Event_v1';Output='P2_EVENT_CLIENT_EVENT.csv'},
    [pscustomobject]@{Profile='P3_EVENT';Pattern='^TEC\d{3}_';Artifact='Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1';Output='P3_EVENT_CLIENT_EVENT.csv'},
    [pscustomobject]@{Profile='P4';Pattern='^CU\d{3}_';Artifact='Custom.TFM.HIDS.P4.Low.Basic_v2';Output='P4_CLIENT_EVENT.csv'}
)
$allCsv=@(Get-ChildItem -LiteralPath $sourceDirectoryFull -Recurse -File -Filter '*.csv'|Sort-Object FullName -Unique)
$reportRows=New-Object System.Collections.ArrayList
$totalInside=0;$totalOutside=0;$totalRaw=0;$totalUnparseable=0;$totalLate=0;$totalUnknownBefore=0;$totalUnknownAdj=0;$totalUnknownRemaining=0
foreach($profile in $profiles){
    $sourceFiles=@($allCsv|Where-Object {$_.BaseName -match $profile.Pattern})
    if($sourceFiles.Count -eq 0){throw "No hay exports raw para $($profile.Profile) en $sourceDirectoryFull"}
    $columns=New-Object System.Collections.ArrayList
    foreach($fixed in @('Artifact','Source','Profile','EventTimeUtc','TimestampFieldUsed','FP_ID','Adjudication','AdjudicationBasis')){[void]$columns.Add($fixed)}
    $columnMap=@{}
    $outRows=New-Object System.Collections.ArrayList
    $rawRows=0;$outside=0;$unparseable=0;$late=0;$unknownBefore=0;$unknownAdj=0;$unknownRemaining=0;$malformed=0;$duplicateHeaderCount=0
    $hashes=New-Object System.Collections.ArrayList
    foreach($file in $sourceFiles){
        [void]$hashes.Add(("{0}:{1}" -f $file.Name,(Get-CRSha256 -Path $file.FullName)))
        $parsed=Read-RobustCsv -Path $file.FullName
        $malformed+=@($parsed.Malformed).Count
        $duplicateHeaderCount+=@($parsed.DuplicateHeaders).Count
        foreach($header in @($parsed.Headers)){[void](Get-MappedRawColumn -Name ([string]$header) -Map $columnMap -Columns $columns)}
        $sourceName=$file.BaseName -replace '-\d{4}-\d{2}-\d{2}T.*$',''
        foreach($row in @($parsed.Rows)){
            $rawRows++
            $timeInfo=Get-EventTimeInfo -Row $row
            if($null -eq $timeInfo){$unparseable++;continue}
            if($timeInfo.Utc -lt $startUtc -or $timeInfo.Utc -gt $endUtc){$outside++;continue}
            $fallback=Get-FallbackTimeInfo -Row $row
            $isLate=$timeInfo.Field -eq 'DetectionTime' -and $null -ne $fallback -and ($fallback.Utc -lt $startUtc -or $fallback.Utc -gt $endUtc)
            if($isLate){$late++}
            $adj=Get-FpAdjudication -Row $row -EventUtc $timeInfo.Utc -Tests $tests
            if($adj.WasUnknown){$unknownBefore++}
            if($adj.Adjudicated){$unknownAdj++}
            if($adj.FP_ID -eq 'UNKNOWN'){$unknownRemaining++}
            $ordered=[ordered]@{
                Artifact=$profile.Artifact
                Source=$sourceName
                Profile=$profile.Profile
                EventTimeUtc=$timeInfo.Utc.ToString('o')
                TimestampFieldUsed=$timeInfo.Field
                FP_ID=$adj.FP_ID
                Adjudication=$adj.Status
                AdjudicationBasis=$adj.Basis
            }
            foreach($property in $row.PSObject.Properties){
                $mapped=Get-MappedRawColumn -Name $property.Name -Map $columnMap -Columns $columns
                $ordered[$mapped]=$property.Value
            }
            [void]$outRows.Add([pscustomobject]$ordered)
        }
    }
    if($malformed -gt 0){throw "Hay $malformed registros CSV malformados en raw de $($profile.Profile); no se genera evidencia silenciosamente."}
    $outputPath=Join-Path $outputDirectoryFull $profile.Output
    $verifiedCount=Write-StrictCsv -Path $outputPath -Columns $columns.ToArray() -Rows $outRows.ToArray()
    $classification=if($isFpCampaign -and $profile.Profile -eq 'P1_CRITICAL' -and $outRows.Count -gt 0){'FAIL_P1_CRITICAL'}elseif($isFpCampaign -and $profile.Profile -eq 'P1_CRITICAL'){'OK_ZERO_P1'}elseif($isFpCampaign -and $profile.Profile -eq 'P2_EVENT' -and $outRows.Count -gt 0){'WARN_VISIBILITY_P2'}elseif($isFpCampaign -and $profile.Profile -eq 'P2_EVENT'){'OK_NO_P2'}else{'INFO_VISIBILITY'}
    [void]$reportRows.Add([pscustomobject]@{
        Repetition=$repetition;Profile=$profile.Profile;StartUtc=$startUtc.ToString('o');EndUtc=$endUtc.ToString('o');SourceFiles=$sourceFiles.Count
        RawRowsRead=$rawRows;RowsOutsideWindow=$outside;RowsInsideWindow=$outRows.Count;UnparseableTimestamps=$unparseable
        LateEventsAdjudicated=$late;UnknownInitially=$unknownBefore;UnknownAdjudicated=$unknownAdj;UnknownRemaining=$unknownRemaining
        DuplicateRawHeadersRenamed=$duplicateHeaderCount;ImportCsvVerified=($verifiedCount -eq $outRows.Count);Classification=$classification
        OutputFile=$profile.Output;OutputSHA256=Get-CRSha256 -Path $outputPath;RawSourceHashes=($hashes.ToArray()-join ';')
    })
    $totalRaw+=$rawRows;$totalOutside+=$outside;$totalInside+=$outRows.Count;$totalUnparseable+=$unparseable;$totalLate+=$late;$totalUnknownBefore+=$unknownBefore;$totalUnknownAdj+=$unknownAdj;$totalUnknownRemaining+=$unknownRemaining
    Write-CRStatus -Message ("{0}/{1}: raw={2}; inside={3}; outside={4}; late={5}; unknown_remaining={6}; Import-Csv=OK" -f $repetition,$profile.Profile,$rawRows,$outRows.Count,$outside,$late,$unknownRemaining) -Level $(if($classification -like 'FAIL*'){'FAIL'}elseif($classification -like 'WARN*'){'WARN'}else{'OK'})
}

$reportCsv=Join-Path $outputDirectoryFull 'CLIENT_EVENT_FILTERING_REPORT.csv'
$reportRows.ToArray()|Export-Csv -LiteralPath $reportCsv -NoTypeInformation -Encoding UTF8
$md=New-Object System.Collections.ArrayList
[void]$md.Add("# CLIENT_EVENT filtering report - $repetition")
[void]$md.Add('')
[void]$md.Add("- Manifest: ``$manifestPath``")
[void]$md.Add("- Ventana UTC usada: ``$($startUtc.ToString('o'))`` a ``$($endUtc.ToString('o'))``")
[void]$md.Add("- Raw directory: ``$sourceDirectoryFull``")
[void]$md.Add("- Filas raw leídas: **$totalRaw**")
[void]$md.Add("- Filas fuera de ventana: **$totalOutside**")
[void]$md.Add("- Filas dentro de ventana: **$totalInside**")
[void]$md.Add("- Timestamps no parseables: **$totalUnparseable**")
[void]$md.Add("- Eventos tardíos adjudicados por DetectionTime: **$totalLate**")
[void]$md.Add("- UNKNOWN iniciales/adjudicados/restantes: **$totalUnknownBefore / $totalUnknownAdj / $totalUnknownRemaining**")
[void]$md.Add('')
[void]$md.Add('| Perfil | Raw | Fuera | Dentro | Tardíos adjudicados | UNKNOWN adjudicados | UNKNOWN restantes | Clasificación | Import-Csv |')
[void]$md.Add('|---|---:|---:|---:|---:|---:|---:|---|---|')
foreach($row in $reportRows){[void]$md.Add("| $($row.Profile) | $($row.RawRowsRead) | $($row.RowsOutsideWindow) | $($row.RowsInsideWindow) | $($row.LateEventsAdjudicated) | $($row.UnknownAdjudicated) | $($row.UnknownRemaining) | $($row.Classification) | $($row.ImportCsvVerified) |")}
[void]$md.Add('')
[void]$md.Add('## Interpretación')
[void]$md.Add('')
[void]$md.Add('- `DetectionTime` prevalece. `Timestamp`, `_ts`, `EventTime` y `TimeCreated` solo son fallback.')
[void]$md.Add('- P2 se clasifica `WARN_VISIBILITY_P2`: visibilidad/forense dual-use, no detección fuerte ni fallo automático.')
[void]$md.Add('- Los UNKNOWN solo se adjudican con marcador explícito, contexto TEC-006 o una ventana única del runner; los restantes no se maquillan.')
[void]$md.Add('- Los raw originales no se modificaron.')
Write-CRUtf8NoBom -Path (Join-Path $outputDirectoryFull 'CLIENT_EVENT_FILTERING_REPORT.md') -Text (($md.ToArray()-join "`r`n")+"`r`n")
Write-CRStatus -Message "Builder v2 completado para $repetition. Reportes: $outputDirectoryFull" -Level OK
