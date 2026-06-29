<#
TFM Velociraptor resource benchmark v1

Measures Velociraptor process and system resource usage for controlled TFM
scenarios. This script is intended to be copied to the VM and executed from:

  C:\Users\seguridad\Desktop\TFM\03_RUNNERS
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ScenarioName,

    [int]$DurationSec = 180,

    [int]$IntervalSec = 2,

    [string]$RunnerPath = "",

    [string]$OutputDir = "C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas\Logs_Pruebas_TFM\BENCHMARKS",

    [bool]$StopVRBeforeRun = $false,

    [bool]$StartVRAfterStop = $false,

    [string]$ServiceName = "",

    [int]$WarmupSec = 30
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$script:Warnings = New-Object System.Collections.ArrayList
$script:WarnKeys = @{}
$script:PreviousProc = @{}

function Add-Warn {
    param(
        [string]$Key,
        [string]$Message
    )

    if (-not $script:WarnKeys.ContainsKey($Key)) {
        $script:WarnKeys[$Key] = $true
        [void]$script:Warnings.Add($Message)
        Write-Host "[WARN] $Message"
    }
}

function Write-Info {
    param([string]$Message)
    Write-Host "[INFO] $Message"
}

function Write-Ok {
    param([string]$Message)
    Write-Host "[OK] $Message"
}

function ConvertTo-SafeName {
    param([string]$Value)
    $safe = $Value -replace '[^A-Za-z0-9_-]', '_'
    if ([string]::IsNullOrWhiteSpace($safe)) {
        return "UNKNOWN"
    }
    return $safe
}

function Get-Number {
    param(
        [object]$Value,
        [double]$Default = 0
    )
    try {
        if ($null -eq $Value) { return $Default }
        if ($Value -is [double] -or $Value -is [int] -or $Value -is [long] -or $Value -is [decimal]) {
            return [double]$Value
        }
        $text = [string]$Value
        if ([string]::IsNullOrWhiteSpace($text)) { return $Default }
        return [double]::Parse($text, [System.Globalization.CultureInfo]::InvariantCulture)
    } catch {
        return $Default
    }
}

function Get-LogicalProcessorCount {
    try {
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        $count = [int]$cs.NumberOfLogicalProcessors
        if ($count -lt 1) { return 1 }
        return $count
    } catch {
        Add-Warn -Key "logical_processors" -Message "No se pudo obtener NumberOfLogicalProcessors; se usa 1."
        return 1
    }
}

function Get-OsInfo {
    try {
        return Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    } catch {
        Add-Warn -Key "os_info" -Message "No se pudo leer Win32_OperatingSystem."
        return $null
    }
}

function Get-VRServiceRecord {
    param([string]$PreferredName)

    $services = @()
    try {
        if (-not [string]::IsNullOrWhiteSpace($PreferredName)) {
            $services = @(Get-Service -Name $PreferredName -ErrorAction Stop)
        }
    } catch {
        Add-Warn -Key "service_preferred_$PreferredName" -Message "No se encontro el servicio indicado '$PreferredName'. Se intentara autodeteccion."
        $services = @()
    }

    if ($services.Count -eq 0) {
        try {
            $services = @(Get-Service -ErrorAction Stop | Where-Object {
                ($_.Name -like "*Velociraptor*") -or ($_.DisplayName -like "*Velociraptor*")
            })
        } catch {
            Add-Warn -Key "service_list" -Message "No se pudo enumerar servicios para autodetectar Velociraptor."
            $services = @()
        }
    }

    if ($services.Count -eq 0) {
        return [pscustomobject]@{
            Found       = $false
            Name        = ""
            DisplayName = ""
            Status      = "NOT_FOUND"
        }
    }

    $selected = @($services | Sort-Object `
        @{ Expression = { if ($_.Status -eq "Running") { 0 } else { 1 } } }, `
        @{ Expression = { $_.Name } } | Select-Object -First 1)[0]

    return [pscustomobject]@{
        Found       = $true
        Name        = [string]$selected.Name
        DisplayName = [string]$selected.DisplayName
        Status      = [string]$selected.Status
    }
}

function Invoke-ServiceStop {
    param([pscustomobject]$ServiceRecord)

    if (-not $ServiceRecord.Found) {
        Add-Warn -Key "stop_no_service" -Message "StopVRBeforeRun solicitado, pero no se encontro servicio Velociraptor."
        return
    }

    try {
        $svc = Get-Service -Name $ServiceRecord.Name -ErrorAction Stop
        if ($svc.Status -ne "Stopped") {
            Write-Info "Stopping service '$($ServiceRecord.Name)'."
            Stop-Service -Name $ServiceRecord.Name -Force -ErrorAction Stop
            $svc.WaitForStatus("Stopped", [TimeSpan]::FromSeconds(30))
        }
    } catch {
        Add-Warn -Key "stop_service_$($ServiceRecord.Name)" -Message "No se pudo parar el servicio '$($ServiceRecord.Name)': $($_.Exception.Message)"
    }
}

function Invoke-ServiceStart {
    param([pscustomobject]$ServiceRecord)

    if (-not $ServiceRecord.Found) {
        Add-Warn -Key "start_no_service" -Message "StartVRAfterStop solicitado, pero no se encontro servicio Velociraptor."
        return
    }

    try {
        $svc = Get-Service -Name $ServiceRecord.Name -ErrorAction Stop
        if ($svc.Status -ne "Running") {
            Write-Info "Starting service '$($ServiceRecord.Name)'."
            Start-Service -Name $ServiceRecord.Name -ErrorAction Stop
            $svc.WaitForStatus("Running", [TimeSpan]::FromSeconds(30))
        }
    } catch {
        Add-Warn -Key "start_service_$($ServiceRecord.Name)" -Message "No se pudo arrancar el servicio '$($ServiceRecord.Name)': $($_.Exception.Message)"
    }
}

function Get-VRProcesses {
    $items = @()
    try {
        $items = @(Get-CimInstance -ClassName Win32_Process -ErrorAction Stop | Where-Object {
            ($_.Name -like "*velociraptor*") -or
            ($_.ExecutablePath -like "*velociraptor*") -or
            ($_.CommandLine -like "*velociraptor*")
        })
    } catch {
        Add-Warn -Key "cim_process_list" -Message "No se pudo consultar Win32_Process; se intentara Get-Process por nombre."
        try {
            $items = @(Get-Process -ErrorAction Stop | Where-Object { $_.ProcessName -like "*velociraptor*" } | ForEach-Object {
                [pscustomobject]@{
                    ProcessId      = $_.Id
                    Name           = $_.ProcessName
                    ExecutablePath = ""
                    CommandLine    = ""
                }
            })
        } catch {
            Add-Warn -Key "get_process_list" -Message "No se pudo consultar procesos Velociraptor."
            $items = @()
        }
    }

    return @($items | ForEach-Object {
        [pscustomobject]@{
            ProcessId      = [int]$_.ProcessId
            Name           = [string]$_.Name
            ExecutablePath = [string]$_.ExecutablePath
            CommandLine    = [string]$_.CommandLine
        }
    })
}

function Get-CounterValue {
    param(
        [string]$CounterPath,
        [string]$Key,
        [switch]$Sum,
        [string]$ExcludeInstanceRegex = ""
    )

    try {
        $samples = @((Get-Counter -Counter $CounterPath -ErrorAction Stop).CounterSamples)
        if (-not [string]::IsNullOrWhiteSpace($ExcludeInstanceRegex)) {
            $samples = @($samples | Where-Object { $_.InstanceName -notmatch $ExcludeInstanceRegex })
        }
        if ($samples.Count -eq 0) { return 0 }
        if ($Sum) {
            return [math]::Round((($samples | Measure-Object -Property CookedValue -Sum).Sum), 2)
        }
        return [math]::Round([double]$samples[0].CookedValue, 2)
    } catch {
        Add-Warn -Key "counter_$Key" -Message "No se pudo leer contador '$CounterPath'. Se usara 0 para esa metrica."
        return 0
    }
}

function Get-ProcessPropertyNumber {
    param(
        [System.Diagnostics.Process]$Process,
        [string]$PropertyName
    )

    try {
        $prop = $Process.PSObject.Properties[$PropertyName]
        if ($null -ne $prop -and $null -ne $prop.Value) {
            return (Get-Number -Value $prop.Value)
        }
    } catch {
    }
    return 0
}

function Initialize-Csv {
    param(
        [string]$Path,
        [string[]]$Columns
    )
    $quoted = @($Columns | ForEach-Object { '"' + ($_ -replace '"', '""') + '"' })
    Set-Content -LiteralPath $Path -Value ($quoted -join ",") -Encoding UTF8
}

function Add-CsvObject {
    param(
        [string]$Path,
        [pscustomobject]$Object,
        [string[]]$Columns
    )

    $line = @($Object | Select-Object -Property $Columns | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1)
    if ($line.Count -gt 0) {
        Add-Content -LiteralPath $Path -Value $line[0] -Encoding UTF8
    }
}

function Get-Stats {
    param(
        [object[]]$Rows,
        [string]$PropertyName
    )

    $values = @($Rows | ForEach-Object { Get-Number -Value $_.$PropertyName } | Sort-Object)
    if ($values.Count -eq 0) {
        return [pscustomobject]@{ Avg = 0; Max = 0; P95 = 0 }
    }

    $avg = (($values | Measure-Object -Average).Average)
    $max = (($values | Measure-Object -Maximum).Maximum)
    $index = [int]([math]::Ceiling($values.Count * 0.95) - 1)
    if ($index -lt 0) { $index = 0 }
    if ($index -ge $values.Count) { $index = $values.Count - 1 }

    return [pscustomobject]@{
        Avg = [math]::Round([double]$avg, 2)
        Max = [math]::Round([double]$max, 2)
        P95 = [math]::Round([double]$values[$index], 2)
    }
}

function Save-ProcessList {
    param(
        [string]$Path,
        [string]$Label
    )

    $now = Get-Date
    $vr = @(Get-VRProcesses)
    $content = New-Object System.Collections.Generic.List[string]
    $content.Add("TFM Velociraptor process list")
    $content.Add("Label: $Label")
    $content.Add("Time: $($now.ToString('o'))")
    $content.Add("")

    if ($vr.Count -eq 0) {
        $content.Add("No Velociraptor processes detected.")
    } else {
        foreach ($item in $vr) {
            $content.Add("PID: $($item.ProcessId)")
            $content.Add("Name: $($item.Name)")
            $content.Add("ExecutablePath: $($item.ExecutablePath)")
            $content.Add("CommandLine: $($item.CommandLine)")
            $content.Add("")
        }
    }

    Set-Content -LiteralPath $Path -Value $content -Encoding UTF8
}

$safeScenario = ConvertTo-SafeName -Value $ScenarioName
$runId = "TFM_BENCH_{0}_{1}" -f $safeScenario, (Get-Date -Format "yyyyMMdd_HHmmss")
$runDir = Join-Path -Path $OutputDir -ChildPath $runId

New-Item -ItemType Directory -Path $runDir -Force | Out-Null

$samplesCsv = Join-Path $runDir "samples.csv"
$processSamplesCsv = Join-Path $runDir "process_samples.csv"
$summaryJson = Join-Path $runDir "summary.json"
$summaryTxt = Join-Path $runDir "summary.txt"
$processListStart = Join-Path $runDir "process_list_start.txt"
$processListEnd = Join-Path $runDir "process_list_end.txt"
$environmentTxt = Join-Path $runDir "environment.txt"

$sampleColumns = @(
    "RunId", "Scenario", "Timestamp", "ElapsedSec", "ServiceName", "ServiceStatus",
    "VR_ProcessCount", "VR_PIDs", "VR_ProcessNames", "VR_ExecutablePaths",
    "VR_CPU_Percent", "VR_RAM_MB", "VR_WorkingSet_MB", "VR_PrivateMemory_MB",
    "VR_Handles", "VR_Threads", "VR_IO_Read_BytesSec", "VR_IO_Write_BytesSec",
    "VR_IO_Total_BytesSec", "System_CPU_Percent", "System_RAM_Used_MB",
    "System_RAM_Free_MB", "System_Disk_BytesSec", "System_Net_BytesSec",
    "RunnerActive"
)

$processColumns = @(
    "RunId", "Scenario", "Timestamp", "ElapsedSec", "ServiceStatus", "PID",
    "ProcessName", "ExecutablePath", "CommandLine", "CPU_Percent", "RAM_MB",
    "WorkingSet_MB", "PrivateMemory_MB", "Handles", "Threads",
    "IO_Read_BytesSec", "IO_Write_BytesSec", "IO_Total_BytesSec", "RunnerActive"
)

Initialize-Csv -Path $samplesCsv -Columns $sampleColumns
Initialize-Csv -Path $processSamplesCsv -Columns $processColumns

$logicalProcessors = Get-LogicalProcessorCount
$osInfo = Get-OsInfo
$serviceInitial = Get-VRServiceRecord -PreferredName $ServiceName
$serviceStatusStart = $serviceInitial.Status
$runnerProc = $null
$runnerStarted = $false
$runnerStartError = ""

Write-Info "RunId: $runId"
Write-Info "Scenario: $ScenarioName"
Write-Info "OutputDir: $runDir"

if ($StopVRBeforeRun) {
    Invoke-ServiceStop -ServiceRecord $serviceInitial
}

if ($StartVRAfterStop) {
    $serviceForStart = Get-VRServiceRecord -PreferredName $ServiceName
    Invoke-ServiceStart -ServiceRecord $serviceForStart
}

if ($WarmupSec -gt 0) {
    Write-Info "WarmupSec: $WarmupSec"
    Start-Sleep -Seconds $WarmupSec
}

$serviceAfterWarmup = Get-VRServiceRecord -PreferredName $ServiceName
$serviceStatusStart = $serviceAfterWarmup.Status

Save-ProcessList -Path $processListStart -Label "START"

$envLines = @(
    "TFM Velociraptor resource benchmark environment",
    "RunId: $runId",
    "Scenario: $ScenarioName",
    "Time: $((Get-Date).ToString('o'))",
    "ComputerName: $env:COMPUTERNAME",
    "UserName: $env:USERNAME",
    "PowerShellVersion: $($PSVersionTable.PSVersion.ToString())",
    "LogicalProcessors: $logicalProcessors",
    "DurationSec: $DurationSec",
    "IntervalSec: $IntervalSec",
    "WarmupSec: $WarmupSec",
    "RunnerPath: $RunnerPath",
    "OutputDir: $OutputDir",
    "StopVRBeforeRun: $StopVRBeforeRun",
    "StartVRAfterStop: $StartVRAfterStop",
    "ServiceNameParameter: $ServiceName",
    "ServiceDetectedName: $($serviceAfterWarmup.Name)",
    "ServiceDetectedDisplayName: $($serviceAfterWarmup.DisplayName)",
    "ServiceStatusAfterWarmup: $($serviceAfterWarmup.Status)"
)

if ($null -ne $osInfo) {
    $envLines += @(
        "OSCaption: $($osInfo.Caption)",
        "OSVersion: $($osInfo.Version)",
        "TotalVisibleMemoryMB: $([math]::Round((Get-Number $osInfo.TotalVisibleMemorySize) / 1024, 2))"
    )
}

Set-Content -LiteralPath $environmentTxt -Value $envLines -Encoding UTF8

if (-not [string]::IsNullOrWhiteSpace($RunnerPath)) {
    if (Test-Path -LiteralPath $RunnerPath) {
        try {
            Write-Info "Starting runner: $RunnerPath"
            $runnerArgs = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ('"{0}"' -f $RunnerPath))
            $runnerProc = Start-Process -FilePath "powershell.exe" -ArgumentList $runnerArgs -PassThru -WindowStyle Hidden -ErrorAction Stop
            $runnerStarted = $true
        } catch {
            $runnerStartError = $_.Exception.Message
            Add-Warn -Key "runner_start" -Message "No se pudo lanzar RunnerPath '$RunnerPath': $runnerStartError"
        }
    } else {
        Add-Warn -Key "runner_missing" -Message "RunnerPath no existe: $RunnerPath"
    }
}

$start = Get-Date
$samples = New-Object System.Collections.ArrayList

while (((Get-Date) - $start).TotalSeconds -lt $DurationSec) {
    $now = Get-Date
    $elapsed = [math]::Round((($now - $start).TotalSeconds), 2)
    $svc = Get-VRServiceRecord -PreferredName $ServiceName
    $vrCim = @(Get-VRProcesses)
    $cimByPid = @{}
    foreach ($item in $vrCim) {
        $cimByPid[[int]$item.ProcessId] = $item
    }

    $processRows = New-Object System.Collections.ArrayList
    $vrCpuTotal = 0.0
    $vrRamTotal = 0.0
    $vrWsTotal = 0.0
    $vrPrivateTotal = 0.0
    $vrHandlesTotal = 0
    $vrThreadsTotal = 0
    $vrReadRateTotal = 0.0
    $vrWriteRateTotal = 0.0
    $vrIoRateTotal = 0.0

    foreach ($vrItem in $vrCim) {
        $procObj = $null
        try {
            $procObj = Get-Process -Id ([int]$vrItem.ProcessId) -ErrorAction Stop
        } catch {
            continue
        }

        $procCpu = Get-Number -Value $procObj.CPU
        $readBytes = Get-ProcessPropertyNumber -Process $procObj -PropertyName "IOReadBytes"
        $writeBytes = Get-ProcessPropertyNumber -Process $procObj -PropertyName "IOWriteBytes"
        $prev = $null
        if ($script:PreviousProc.ContainsKey($procObj.Id)) {
            $prev = $script:PreviousProc[$procObj.Id]
        }

        $cpuPercent = 0.0
        $readRate = 0.0
        $writeRate = 0.0
        if ($null -ne $prev) {
            $deltaSec = (($now - $prev.Timestamp).TotalSeconds)
            if ($deltaSec -gt 0) {
                $cpuPercent = (($procCpu - $prev.CpuSeconds) / $deltaSec / $logicalProcessors) * 100
                if ($cpuPercent -lt 0) { $cpuPercent = 0 }
                $readRate = (($readBytes - $prev.ReadBytes) / $deltaSec)
                $writeRate = (($writeBytes - $prev.WriteBytes) / $deltaSec)
                if ($readRate -lt 0) { $readRate = 0 }
                if ($writeRate -lt 0) { $writeRate = 0 }
            }
        }

        $script:PreviousProc[$procObj.Id] = [pscustomobject]@{
            Timestamp  = $now
            CpuSeconds = $procCpu
            ReadBytes  = $readBytes
            WriteBytes = $writeBytes
        }

        $workingSetMb = [math]::Round((Get-Number -Value $procObj.WorkingSet64) / 1MB, 2)
        $privateMb = [math]::Round((Get-Number -Value $procObj.PrivateMemorySize64) / 1MB, 2)
        $handles = [int](Get-Number -Value $procObj.HandleCount)
        $threads = 0
        try { $threads = @($procObj.Threads).Count } catch { $threads = 0 }

        $row = [pscustomobject]@{
            RunId             = $runId
            Scenario          = $ScenarioName
            Timestamp         = $now.ToString("o")
            ElapsedSec        = $elapsed
            ServiceStatus     = $svc.Status
            PID               = $procObj.Id
            ProcessName       = $vrItem.Name
            ExecutablePath    = $vrItem.ExecutablePath
            CommandLine       = $vrItem.CommandLine
            CPU_Percent       = [math]::Round($cpuPercent, 2)
            RAM_MB            = $workingSetMb
            WorkingSet_MB     = $workingSetMb
            PrivateMemory_MB  = $privateMb
            Handles           = $handles
            Threads           = $threads
            IO_Read_BytesSec  = [math]::Round($readRate, 2)
            IO_Write_BytesSec = [math]::Round($writeRate, 2)
            IO_Total_BytesSec = [math]::Round(($readRate + $writeRate), 2)
            RunnerActive      = if ($runnerProc) { -not $runnerProc.HasExited } else { $false }
        }

        [void]$processRows.Add($row)
        Add-CsvObject -Path $processSamplesCsv -Object $row -Columns $processColumns

        $vrCpuTotal += $cpuPercent
        $vrRamTotal += $workingSetMb
        $vrWsTotal += $workingSetMb
        $vrPrivateTotal += $privateMb
        $vrHandlesTotal += $handles
        $vrThreadsTotal += $threads
        $vrReadRateTotal += $readRate
        $vrWriteRateTotal += $writeRate
        $vrIoRateTotal += ($readRate + $writeRate)
    }

    $currentOs = Get-OsInfo
    $freeMemMb = 0.0
    $usedMemMb = 0.0
    if ($null -ne $currentOs) {
        $freeMemMb = [math]::Round((Get-Number -Value $currentOs.FreePhysicalMemory) / 1024, 2)
        $totalMemMb = [math]::Round((Get-Number -Value $currentOs.TotalVisibleMemorySize) / 1024, 2)
        $usedMemMb = [math]::Round(($totalMemMb - $freeMemMb), 2)
    }

    $runnerActive = if ($runnerProc) { -not $runnerProc.HasExited } else { $false }
    $pids = @($processRows | ForEach-Object { [string]$_.PID })
    $names = @($processRows | ForEach-Object { [string]$_.ProcessName })
    $paths = @($processRows | ForEach-Object { [string]$_.ExecutablePath })

    $sampleRow = [pscustomobject]@{
        RunId                  = $runId
        Scenario               = $ScenarioName
        Timestamp              = $now.ToString("o")
        ElapsedSec             = $elapsed
        ServiceName            = $svc.Name
        ServiceStatus          = $svc.Status
        VR_ProcessCount        = $processRows.Count
        VR_PIDs                = ($pids -join ";")
        VR_ProcessNames        = ($names -join ";")
        VR_ExecutablePaths     = ($paths -join ";")
        VR_CPU_Percent         = [math]::Round($vrCpuTotal, 2)
        VR_RAM_MB              = [math]::Round($vrRamTotal, 2)
        VR_WorkingSet_MB       = [math]::Round($vrWsTotal, 2)
        VR_PrivateMemory_MB    = [math]::Round($vrPrivateTotal, 2)
        VR_Handles             = $vrHandlesTotal
        VR_Threads             = $vrThreadsTotal
        VR_IO_Read_BytesSec    = [math]::Round($vrReadRateTotal, 2)
        VR_IO_Write_BytesSec   = [math]::Round($vrWriteRateTotal, 2)
        VR_IO_Total_BytesSec   = [math]::Round($vrIoRateTotal, 2)
        System_CPU_Percent     = Get-CounterValue -CounterPath "\Processor(_Total)\% Processor Time" -Key "system_cpu"
        System_RAM_Used_MB     = $usedMemMb
        System_RAM_Free_MB     = $freeMemMb
        System_Disk_BytesSec   = Get-CounterValue -CounterPath "\PhysicalDisk(_Total)\Disk Bytes/sec" -Key "disk_total"
        System_Net_BytesSec    = Get-CounterValue -CounterPath "\Network Interface(*)\Bytes Total/sec" -Key "net_total" -Sum -ExcludeInstanceRegex "isatap|teredo|loopback"
        RunnerActive           = $runnerActive
    }

    [void]$samples.Add($sampleRow)
    Add-CsvObject -Path $samplesCsv -Object $sampleRow -Columns $sampleColumns

    Start-Sleep -Seconds $IntervalSec
}

$serviceFinal = Get-VRServiceRecord -PreferredName $ServiceName
Save-ProcessList -Path $processListEnd -Label "END"

$runnerStillRunning = $false
if ($runnerProc -and -not $runnerProc.HasExited) {
    $runnerStillRunning = $true
    Add-Warn -Key "runner_still_running" -Message "El runner sigue activo al finalizar la ventana de benchmark. No se mata el proceso."
}

$vrCpuStats = Get-Stats -Rows @($samples) -PropertyName "VR_CPU_Percent"
$vrRamStats = Get-Stats -Rows @($samples) -PropertyName "VR_RAM_MB"
$vrIoStats = Get-Stats -Rows @($samples) -PropertyName "VR_IO_Total_BytesSec"
$sysCpuStats = Get-Stats -Rows @($samples) -PropertyName "System_CPU_Percent"
$sysRamStats = Get-Stats -Rows @($samples) -PropertyName "System_RAM_Used_MB"
$diskStats = Get-Stats -Rows @($samples) -PropertyName "System_Disk_BytesSec"
$netStats = Get-Stats -Rows @($samples) -PropertyName "System_Net_BytesSec"
$vrCountStats = Get-Stats -Rows @($samples) -PropertyName "VR_ProcessCount"

$end = Get-Date
$summaryObj = [pscustomobject]@{
    SchemaVersion              = "1.0"
    ScriptVersion              = "TFM_Benchmark_VR_Resource_Usage_v1"
    RunId                      = $runId
    Scenario                   = $ScenarioName
    StartTime                  = $start.ToString("o")
    EndTime                    = $end.ToString("o")
    DurationSec                = $DurationSec
    ActualDurationSec          = [math]::Round((($end - $start).TotalSeconds), 2)
    IntervalSec                = $IntervalSec
    WarmupSec                  = $WarmupSec
    Samples                    = $samples.Count
    ServiceName                = $serviceFinal.Name
    ServiceStatusStart         = $serviceStatusStart
    ServiceStatusEnd           = $serviceFinal.Status
    StopVRBeforeRun            = $StopVRBeforeRun
    StartVRAfterStop           = $StartVRAfterStop
    RunnerPath                 = $RunnerPath
    RunnerStarted              = $runnerStarted
    RunnerStartError           = $runnerStartError
    RunnerStillRunningAtEnd    = $runnerStillRunning
    Avg_VR_CPU_Percent         = $vrCpuStats.Avg
    Max_VR_CPU_Percent         = $vrCpuStats.Max
    P95_VR_CPU_Percent         = $vrCpuStats.P95
    Avg_VR_RAM_MB              = $vrRamStats.Avg
    Max_VR_RAM_MB              = $vrRamStats.Max
    P95_VR_RAM_MB              = $vrRamStats.P95
    Avg_VR_IO_BytesSec         = $vrIoStats.Avg
    Max_VR_IO_BytesSec         = $vrIoStats.Max
    Avg_System_CPU_Percent     = $sysCpuStats.Avg
    Max_System_CPU_Percent     = $sysCpuStats.Max
    Avg_System_RAM_Used_MB     = $sysRamStats.Avg
    Max_System_RAM_Used_MB     = $sysRamStats.Max
    Avg_Disk_BytesSec          = $diskStats.Avg
    Max_Disk_BytesSec          = $diskStats.Max
    Avg_Net_BytesSec           = $netStats.Avg
    Max_Net_BytesSec           = $netStats.Max
    VRProcessCountAvg          = $vrCountStats.Avg
    VRProcessCountMax          = $vrCountStats.Max
    SamplesCsv                 = $samplesCsv
    ProcessSamplesCsv          = $processSamplesCsv
    ProcessListStart           = $processListStart
    ProcessListEnd             = $processListEnd
    EnvironmentTxt             = $environmentTxt
    Warnings                   = @($script:Warnings)
}

$summaryObj | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $summaryJson -Encoding UTF8

$summaryLines = @(
    "TFM Velociraptor Resource Benchmark",
    "",
    "RunId: $runId",
    "Scenario: $ScenarioName",
    "StartTime: $($start.ToString('o'))",
    "EndTime: $($end.ToString('o'))",
    "DurationSec configured: $DurationSec",
    "ActualDurationSec: $($summaryObj.ActualDurationSec)",
    "IntervalSec: $IntervalSec",
    "Samples: $($samples.Count)",
    "",
    "ServiceName: $($summaryObj.ServiceName)",
    "ServiceStatusStart: $($summaryObj.ServiceStatusStart)",
    "ServiceStatusEnd: $($summaryObj.ServiceStatusEnd)",
    "",
    "Avg_VR_CPU_Percent: $($summaryObj.Avg_VR_CPU_Percent)",
    "Max_VR_CPU_Percent: $($summaryObj.Max_VR_CPU_Percent)",
    "P95_VR_CPU_Percent: $($summaryObj.P95_VR_CPU_Percent)",
    "Avg_VR_RAM_MB: $($summaryObj.Avg_VR_RAM_MB)",
    "Max_VR_RAM_MB: $($summaryObj.Max_VR_RAM_MB)",
    "P95_VR_RAM_MB: $($summaryObj.P95_VR_RAM_MB)",
    "Avg_VR_IO_BytesSec: $($summaryObj.Avg_VR_IO_BytesSec)",
    "Max_VR_IO_BytesSec: $($summaryObj.Max_VR_IO_BytesSec)",
    "",
    "Avg_System_CPU_Percent: $($summaryObj.Avg_System_CPU_Percent)",
    "Max_System_CPU_Percent: $($summaryObj.Max_System_CPU_Percent)",
    "Avg_System_RAM_Used_MB: $($summaryObj.Avg_System_RAM_Used_MB)",
    "Max_System_RAM_Used_MB: $($summaryObj.Max_System_RAM_Used_MB)",
    "Avg_Disk_BytesSec: $($summaryObj.Avg_Disk_BytesSec)",
    "Max_Disk_BytesSec: $($summaryObj.Max_Disk_BytesSec)",
    "Avg_Net_BytesSec: $($summaryObj.Avg_Net_BytesSec)",
    "Max_Net_BytesSec: $($summaryObj.Max_Net_BytesSec)",
    "",
    "VRProcessCountAvg: $($summaryObj.VRProcessCountAvg)",
    "VRProcessCountMax: $($summaryObj.VRProcessCountMax)",
    "",
    "RunnerPath: $RunnerPath",
    "RunnerStarted: $runnerStarted",
    "RunnerStillRunningAtEnd: $runnerStillRunning",
    "",
    "Outputs:",
    "samples.csv: $samplesCsv",
    "process_samples.csv: $processSamplesCsv",
    "summary.json: $summaryJson",
    "process_list_start.txt: $processListStart",
    "process_list_end.txt: $processListEnd",
    "environment.txt: $environmentTxt",
    "",
    "Warnings:"
)

if ($script:Warnings.Count -eq 0) {
    $summaryLines += "- none"
} else {
    foreach ($warning in $script:Warnings) {
        $summaryLines += "- $warning"
    }
}

Set-Content -LiteralPath $summaryTxt -Value $summaryLines -Encoding UTF8

Write-Ok "Benchmark completed"
Write-Ok "RunDir: $runDir"
Write-Ok "samples.csv: $samplesCsv"
Write-Ok "process_samples.csv: $processSamplesCsv"
Write-Ok "summary.json: $summaryJson"
