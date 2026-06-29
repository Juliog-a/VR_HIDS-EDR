[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$FpCampaignDirectory,
    [string]$ManifestCampaignDirectory=''
)

Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$root=[System.IO.Path]::GetFullPath($FpCampaignDirectory)
if([string]::IsNullOrWhiteSpace($ManifestCampaignDirectory)){$ManifestCampaignDirectory=$root}
$manifestRoot=[System.IO.Path]::GetFullPath($ManifestCampaignDirectory)
$expected=@('P1_CRITICAL_CLIENT_EVENT.csv','P2_EVENT_CLIENT_EVENT.csv','P3_EVENT_CLIENT_EVENT.csv','P4_CLIENT_EVENT.csv')
$failures=New-Object System.Collections.ArrayList
$verified=0
foreach($number in 1..3){
    $rep='REP_{0:D2}' -f $number
    $repRoot=Join-Path $root $rep
    $manifestPath=Join-Path (Join-Path $manifestRoot $rep) 'repetition_manifest.json'
    try{$manifest=Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8|ConvertFrom-Json}
    catch{[void]$failures.Add("$rep manifest: $($_.Exception.Message)");continue}
    $start=[DateTimeOffset]::Parse([string]$manifest.StartUtc).UtcDateTime
    $end=[DateTimeOffset]::Parse([string]$manifest.EndUtc).UtcDateTime
    $exports=Join-Path $repRoot 'velociraptor_exports'
    foreach($name in $expected){
        $path=Join-Path $exports $name
        if(-not(Test-Path -LiteralPath $path -PathType Leaf)){[void]$failures.Add("$rep/$name ausente");continue}
        try{$rows=@(Import-Csv -LiteralPath $path -ErrorAction Stop)}
        catch{[void]$failures.Add("$rep/$name Import-Csv: $($_.Exception.Message)");continue}
        $outside=0;$badTime=0
        foreach($row in $rows){
            $raw=[string]$row.EventTimeUtc
            $event=[DateTimeOffset]::MinValue
            if(-not[DateTimeOffset]::TryParse($raw,[ref]$event)){$badTime++;continue}
            if($event.UtcDateTime -lt $start -or $event.UtcDateTime -gt $end){$outside++}
        }
        $physicalLines=@(Get-Content -LiteralPath $path).Count
        if($badTime -gt 0 -or $outside -gt 0){[void]$failures.Add("$rep/$name bad_time=$badTime outside=$outside");continue}
        if($physicalLines -ne ($rows.Count+1)){[void]$failures.Add("$rep/$name contiene saltos multilínea: physical=$physicalLines logical=$($rows.Count+1)");continue}
        $verified++
        Write-Host ("[OK] {0}/{1}: rows={2}; Import-Csv=OK; window=OK; one-line=OK" -f $rep,$name,$rows.Count) -ForegroundColor Green
    }
    foreach($report in @('CLIENT_EVENT_FILTERING_REPORT.md','CLIENT_EVENT_FILTERING_REPORT.csv')){
        if(-not(Test-Path -LiteralPath (Join-Path $exports $report) -PathType Leaf)){[void]$failures.Add("$rep/$report ausente")}
    }
}
if($failures.Count -gt 0){
    foreach($failure in $failures){Write-Host "[FAIL] $failure" -ForegroundColor Red}
    exit 2
}
Write-Host "[OK] VERIFY PASS: $verified/12 CSV parseables y limitados a su repetition_manifest.json." -ForegroundColor Green
