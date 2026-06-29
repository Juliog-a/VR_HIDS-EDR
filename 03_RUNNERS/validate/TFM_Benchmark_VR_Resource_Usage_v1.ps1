<#
TFM Velociraptor resource benchmark v1.

Purpose:
- Measure Velociraptor Windows client resource usage.
- Keep client HIDS metrics separated from local server/GUI processes.
- Produce reproducible CSV/JSON/TXT evidence for the TFM performance section.

Copy and run in VM from:
  C:\Users\seguridad\Desktop\TFM\03_RUNNERS
#>

[CmdletBinding()]
param(
    [string]$ScenarioName = "VR_IDLE",
    [int]$DurationSec = 180,
    [int]$IntervalSec = 2,
    [string]$RunnerPath = "",
    [string]$OutputDir = "C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas\Logs_Pruebas_TFM\BENCHMARKS",
    [bool]$StopVRBeforeRun = $false,
    [bool]$StartVRAfterStop = $false,
    [bool]$SkipServiceControl = $false,
    [int]$WarmupSec = 30,
    [bool]$IncludeServerGuiInTotal = $false
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$script:Warnings = New-Object System.Collections.ArrayList
$script:WarnKeys = @{}

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
    if ([string]::IsNullOrWhiteSpace($safe)) { return "UNKNOWN" }
    return $safe
}

function Convert-ToNullableNumber {
    param([object]$Value)
    try {
        if ($null -eq $Value) { return $null }
        if ($Value -is [double] -or $Value -is [int] -or $Value -is [long] -or $Value -is [decimal]) {
            return [double]$Value
        }
        $text = [string]$Value
        if ([string]::IsNullOrWhiteSpace($text) -or $text -eq "NA") { return $null }
        return [double]::Parse($text, [System.Globalization.CultureInfo]::InvariantCulture)
    } catch {
        return $null
    }
}

function Convert-ToCsvValue {
    param([object]$Value)
    if ($null -eq $Value) { return "NA" }
    if ($Value -is [bool]) {
        if ($Value) { return "true" }
        return "false"
    }
    if ($Value -is [double] -or $Value -is [decimal] -or $Value -is [float]) {
        return ([double]$Value).ToString("0.####", [System.Globalization.CultureInfo]::InvariantCulture)
    }
    return [string]$Value
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

    $csvObject = [ordered]@{}
    foreach ($column in $Columns) {
        $prop = $Object.PSObject.Properties[$column]
        if ($null -eq $prop) {
            $csvObject[$column] = "NA"
        } else {
            $csvObject[$column] = Convert-ToCsvValue -Value $prop.Value
        }
    }

    $line = @([pscustomobject]$csvObject | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1)
    if ($line.Count -gt 0) {
        Add-Content -LiteralPath $Path -Value $line[0] -Encoding UTF8
    }
}

function Get-Stats {
    param(
        [object[]]$Rows,
        [string]$PropertyName
    )

    $values = @($Rows | ForEach-Object {
        $prop = $_.PSObject.Properties[$PropertyName]
        if ($null -ne $prop) { Convert-ToNullableNumber -Value $prop.Value }
    } | Where-Object { $null -ne $_ } | Sort-Object)

    if ($values.Count -eq 0) {
        return [pscustomobject]@{ Avg = $null; Max = $null; P95 = $null; Count = 0 }
    }

    $avg = (($values | Measure-Object -Average).Average)
    $max = (($values | Measure-Object -Maximum).Maximum)
    $index = [int]([math]::Ceiling($values.Count * 0.95) - 1)
    if ($index -lt 0) { $index = 0 }
    if ($index -ge $values.Count) { $index = $values.Count - 1 }

    return [pscustomobject]@{
        Avg   = [math]::Round([double]$avg, 2)
        Max   = [math]::Round([double]$max, 2)
        P95   = [math]::Round([double]$values[$index], 2)
        Count = $values.Count
    }
}

function Get-VRServiceRecord {
    $service = $null

    try {
        $services = @(Get-CimInstance -ClassName Win32_Service -ErrorAction Stop | Where-Object {
            ($_.Name -like "*Velociraptor*") -or ($_.DisplayName -like "*Velociraptor*")
        })

        if ($services.Count -gt 0) {
            $service = @($services | Sort-Object `
                @{ Expression = { if ($_.Name -eq "Velociraptor") { 0 } else { 1 } } }, `
                @{ Expression = { if ($_.State -eq "Running") { 0 } else { 1 } } }, `
                @{ Expression = { $_.Name } } | Select-Object -First 1)[0]
        }
    } catch {
        Add-Warn -Key "service_cim" -Message "No se pudo consultar Win32_Service: $($_.Exception.Message)"
    }

    if ($null -ne $service) {
        return [pscustomobject]@{
            Found       = $true
            Name        = [string]$service.Name
            DisplayName = [string]$service.DisplayName
            Status      = [string]$service.State
            StartType   = [string]$service.StartMode
            PathName    = [string]$service.PathName
        }
    }

    try {
        $svc = @(Get-Service -ErrorAction Stop | Where-Object {
            ($_.Name -like "*Velociraptor*") -or ($_.DisplayName -like "*Velociraptor*")
        } | Sort-Object Name | Select-Object -First 1)

        if ($svc.Count -gt 0) {
            return [pscustomobject]@{
                Found       = $true
                Name        = [string]$svc[0].Name
                DisplayName = [string]$svc[0].DisplayName
                Status      = [string]$svc[0].Status
                StartType   = ""
                PathName    = ""
            }
        }
    } catch {
        Add-Warn -Key "service_getservice" -Message "No se pudo consultar Get-Service: $($_.Exception.Message)"
    }

    return [pscustomobject]@{
        Found       = $false
        Name        = ""
        DisplayName = ""
        Status      = "NOT_FOUND"
        StartType   = ""
        PathName    = ""
    }
}

function Invoke-VRServiceStop {
    param([pscustomobject]$ServiceRecord)

    if (-not $ServiceRecord.Found) {
        Add-Warn -Key "stop_no_service" -Message "StopVRBeforeRun=true, pero no se encontro servicio Velociraptor."
        return $false
    }

    try {
        $svc = Get-Service -Name $ServiceRecord.Name -ErrorAction Stop
        if ($svc.Status -ne "Stopped") {
            Write-Info "Stopping service '$($ServiceRecord.Name)'."
            Stop-Service -Name $ServiceRecord.Name -Force -ErrorAction Stop
            $svc.WaitForStatus("Stopped", [TimeSpan]::FromSeconds(30))
        }
        return $true
    } catch {
        Add-Warn -Key "stop_service_failed" -Message "No se pudo parar '$($ServiceRecord.Name)': $($_.Exception.Message). Ejecuta PowerShell como Administrador o usa -SkipServiceControl `$true tras parar manualmente."
        return $false
    }
}

function Invoke-VRServiceStart {
    param([pscustomobject]$ServiceRecord)

    if (-not $ServiceRecord.Found) {
        Add-Warn -Key "start_no_service" -Message "StartVRAfterStop=true, pero no se encontro servicio Velociraptor."
        return $false
    }

    try {
        $svc = Get-Service -Name $ServiceRecord.Name -ErrorAction Stop
        if ($svc.Status -ne "Running") {
            Write-Info "Starting service '$($ServiceRecord.Name)'."
            Start-Service -Name $ServiceRecord.Name -ErrorAction Stop
            $svc.WaitForStatus("Running", [TimeSpan]::FromSeconds(30))
        }
        return $true
    } catch {
        Add-Warn -Key "start_service_failed" -Message "No se pudo arrancar '$($ServiceRecord.Name)': $($_.Exception.Message). Ejecuta PowerShell como Administrador o usa -SkipServiceControl `$true tras arrancar manualmente."
        return $false
    }
}

function Get-ProcessRole {
    param(
        [string]$Name,
        [string]$ExecutablePath,
        [string]$CommandLine
    )

    $nameText = if ($null -eq $Name) { "" } else { $Name }
    $exeText = if ($null -eq $ExecutablePath) { "" } else { $ExecutablePath }
    $cmdText = if ($null -eq $CommandLine) { "" } else { $CommandLine }
    $combined = "$nameText $exeText $cmdText"

    if ($nameText -ieq "notepad.exe") {
        return "EXCLUDED_NOTEPAD"
    }

    if ($combined -notmatch '(?i)velociraptor') {
        return "NOT_VELOCIRAPTOR"
    }

    $isClientPath = ($exeText -match '(?i)^C:\\Program Files\\Velociraptor\\Velociraptor\.exe$')
    $isClientCmd = (($cmdText -match '(?i)client\.config\.yaml') -and ($cmdText -match '(?i)service\s+run'))
    if ($isClientPath -or $isClientCmd) {
        return "CLIENT_SERVICE"
    }

    $isServerPath = ($exeText -match '(?i)^C:\\Users\\seguridad\\Desktop\\TFM\\Velociraptor\\')
    $isServerCmd = (($cmdText -match '(?i)server\.config\.yaml') -and ($cmdText -match '(?i)\bgui\b'))
    if ($isServerPath -or $isServerCmd) {
        return "SERVER_GUI"
    }

    return "OTHER_VELOCIRAPTOR"
}

function Get-VRProcessRecords {
    $records = New-Object System.Collections.ArrayList
    try {
        $items = @(Get-CimInstance -ClassName Win32_Process -ErrorAction Stop)
    } catch {
        Add-Warn -Key "process_cim" -Message "No se pudo consultar Win32_Process: $($_.Exception.Message)"
        $items = @()
    }

    foreach ($item in $items) {
        $name = [string]$item.Name
        $exe = [string]$item.ExecutablePath
        $cmd = [string]$item.CommandLine
        $role = Get-ProcessRole -Name $name -ExecutablePath $exe -CommandLine $cmd

        if ($role -in @("CLIENT_SERVICE", "SERVER_GUI", "OTHER_VELOCIRAPTOR")) {
            [void]$records.Add([pscustomobject]@{
                PID            = [int]$item.ProcessId
                ProcessName    = $name
                ExecutablePath = $exe
                CommandLine    = $cmd
                Role           = $role
            })
        }
    }

    return @($records)
}

function Get-PerfProcByPid {
    $map = @{}
    try {
        $perfItems = @(Get-CimInstance -ClassName Win32_PerfFormattedData_PerfProc_Process -ErrorAction Stop | Where-Object {
            $_.IDProcess -gt 0
        })
        foreach ($item in $perfItems) {
            $map[[int]$item.IDProcess] = $item
        }
    } catch {
        Add-Warn -Key "perfproc" -Message "No se pudo consultar Win32_PerfFormattedData_PerfProc_Process. CPU/IO por proceso pueden quedar como NA."
    }
    return $map
}

function Get-SystemCpuMetric {
    try {
        $total = @(Get-CimInstance -ClassName Win32_PerfFormattedData_PerfOS_Processor -ErrorAction Stop | Where-Object { $_.Name -eq "_Total" })
        if ($total.Count -gt 0) {
            return [pscustomobject]@{ Value = [double]$total[0].PercentProcessorTime; Available = $true; Source = "Win32_PerfFormattedData_PerfOS_Processor" }
        }
    } catch {
        Add-Warn -Key "sys_cpu_perf" -Message "No se pudo leer CPU sistema por PerfOS_Processor."
    }

    try {
        $cpus = @(Get-CimInstance -ClassName Win32_Processor -ErrorAction Stop | Where-Object { $null -ne $_.LoadPercentage })
        if ($cpus.Count -gt 0) {
            $avg = (($cpus | Measure-Object -Property LoadPercentage -Average).Average)
            return [pscustomobject]@{ Value = [math]::Round([double]$avg, 2); Available = $true; Source = "Win32_Processor.LoadPercentage" }
        }
    } catch {
        Add-Warn -Key "sys_cpu_wmi" -Message "No se pudo leer CPU sistema por Win32_Processor."
    }

    return [pscustomobject]@{ Value = $null; Available = $false; Source = "NA" }
}

function Get-SystemDiskMetric {
    try {
        $disk = @(Get-CimInstance -ClassName Win32_PerfFormattedData_PerfDisk_PhysicalDisk -ErrorAction Stop | Where-Object { $_.Name -eq "_Total" })
        if ($disk.Count -gt 0) {
            return [pscustomobject]@{ Value = [double]$disk[0].DiskBytesPersec; Available = $true; Source = "Win32_PerfFormattedData_PerfDisk_PhysicalDisk" }
        }
    } catch {
        Add-Warn -Key "sys_disk_perf" -Message "No se pudo leer disco sistema por PerfDisk_PhysicalDisk."
    }

    return [pscustomobject]@{ Value = $null; Available = $false; Source = "NA" }
}

function Get-SystemNetMetric {
    try {
        $interfaces = @(Get-CimInstance -ClassName Win32_PerfFormattedData_Tcpip_NetworkInterface -ErrorAction Stop | Where-Object {
            $_.Name -notmatch '(?i)isatap|teredo|loopback'
        })
        if ($interfaces.Count -gt 0) {
            $sum = (($interfaces | Measure-Object -Property BytesTotalPersec -Sum).Sum)
            return [pscustomobject]@{ Value = [double]$sum; Available = $true; Source = "Win32_PerfFormattedData_Tcpip_NetworkInterface" }
        }
    } catch {
        Add-Warn -Key "sys_net_perf" -Message "No se pudo leer red sistema por Tcpip_NetworkInterface."
    }

    return [pscustomobject]@{ Value = $null; Available = $false; Source = "NA" }
}

function Get-SystemMemoryMetric {
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $free = [math]::Round(([double]$os.FreePhysicalMemory / 1024), 2)
        $total = [math]::Round(([double]$os.TotalVisibleMemorySize / 1024), 2)
        $used = [math]::Round(($total - $free), 2)
        return [pscustomobject]@{ UsedMB = $used; FreeMB = $free; TotalMB = $total; Available = $true; Source = "Win32_OperatingSystem" }
    } catch {
        Add-Warn -Key "sys_mem" -Message "No se pudo leer memoria sistema por Win32_OperatingSystem."
        return [pscustomobject]@{ UsedMB = $null; FreeMB = $null; TotalMB = $null; Available = $false; Source = "NA" }
    }
}

function New-ZeroAggregate {
    return [pscustomobject]@{
        ProcessCount        = 0
        PIDs                = ""
        ProcessNames        = ""
        ExecutablePaths     = ""
        CPU_Percent         = 0.0
        RAM_MB              = 0.0
        WorkingSet_MB       = 0.0
        PrivateMemory_MB    = 0.0
        Handles             = 0
        Threads             = 0
        IO_Read_BytesSec    = 0.0
        IO_Write_BytesSec   = 0.0
        IO_Total_BytesSec   = 0.0
        ProcessPerfAvailable = $true
    }
}

function Add-ToAggregate {
    param(
        [pscustomobject]$Aggregate,
        [pscustomobject]$ProcessRow
    )

    $Aggregate.ProcessCount += 1
    $Aggregate.PIDs = (@($Aggregate.PIDs, [string]$ProcessRow.PID) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ";"
    $Aggregate.ProcessNames = (@($Aggregate.ProcessNames, [string]$ProcessRow.ProcessName) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ";"
    $Aggregate.ExecutablePaths = (@($Aggregate.ExecutablePaths, [string]$ProcessRow.ExecutablePath) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ";"

    foreach ($metric in @("CPU_Percent", "RAM_MB", "WorkingSet_MB", "PrivateMemory_MB", "IO_Read_BytesSec", "IO_Write_BytesSec", "IO_Total_BytesSec")) {
        $value = Convert-ToNullableNumber -Value $ProcessRow.$metric
        if ($null -ne $value) {
            $Aggregate.$metric = [math]::Round(([double]$Aggregate.$metric + $value), 2)
        } elseif ($metric -in @("CPU_Percent", "IO_Read_BytesSec", "IO_Write_BytesSec", "IO_Total_BytesSec")) {
            $Aggregate.ProcessPerfAvailable = $false
        }
    }

    $Aggregate.Handles += [int]($ProcessRow.Handles)
    $Aggregate.Threads += [int]($ProcessRow.Threads)
}

function Save-ProcessList {
    param(
        [string]$Path,
        [string]$Label
    )

    $rows = @(Get-VRProcessRecords)
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("TFM Velociraptor process list")
    $lines.Add("Label: $Label")
    $lines.Add("Time: $((Get-Date).ToString('o'))")
    $lines.Add("")

    if ($rows.Count -eq 0) {
        $lines.Add("No Velociraptor processes detected.")
    } else {
        foreach ($row in $rows) {
            $lines.Add("Role: $($row.Role)")
            $lines.Add("PID: $($row.PID)")
            $lines.Add("Name: $($row.ProcessName)")
            $lines.Add("ExecutablePath: $($row.ExecutablePath)")
            $lines.Add("CommandLine: $($row.CommandLine)")
            $lines.Add("")
        }
    }

    Set-Content -LiteralPath $Path -Value $lines -Encoding UTF8
}

function Get-ScenarioValidity {
    param(
        [string]$Scenario,
        [int]$ClientMax,
        [int]$RunnerObservedSamples
    )

    $reasons = New-Object System.Collections.ArrayList
    $validity = "VALID"

    if ($Scenario -eq "BASELINE_NO_VR") {
        if ($ClientMax -gt 0) {
            $validity = "INVALID"
            [void]$reasons.Add("BASELINE_NO_VR invalido: VR_Client_ProcessCountMax > 0.")
        }
    } elseif ($Scenario -eq "VR_IDLE") {
        if ($ClientMax -eq 0) {
            $validity = "INVALID"
            [void]$reasons.Add("VR_IDLE invalido: no se observo proceso cliente Velociraptor.")
        }
    } elseif ($Scenario -eq "VR_TEC_RUNNER") {
        if ($ClientMax -eq 0) {
            $validity = "INVALID"
            [void]$reasons.Add("VR_TEC_RUNNER invalido: no se observo proceso cliente Velociraptor.")
        }
        if ($RunnerObservedSamples -eq 0) {
            $validity = "INVALID"
            [void]$reasons.Add("VR_TEC_RUNNER invalido: no se observo runner activo durante el muestreo.")
        }
    } else {
        $validity = "NOT_ASSESSED"
        [void]$reasons.Add("Escenario auxiliar/no canonico; no se evalua contra criterios finales.")
    }

    if ($reasons.Count -eq 0) {
        [void]$reasons.Add("Criterios de validez satisfechos para el escenario.")
    }

    return [pscustomobject]@{
        Validity = $validity
        Reasons  = @($reasons)
    }
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
    "VR_Client_ProcessCount", "VR_Client_PIDs", "VR_Client_CPU_Percent", "VR_Client_RAM_MB",
    "VR_Client_WorkingSet_MB", "VR_Client_PrivateMemory_MB", "VR_Client_Handles",
    "VR_Client_Threads", "VR_Client_IO_Read_BytesSec", "VR_Client_IO_Write_BytesSec",
    "VR_Client_IO_Total_BytesSec", "VR_ServerGUI_ProcessCount", "VR_ServerGUI_PIDs",
    "VR_ServerGUI_CPU_Percent", "VR_ServerGUI_RAM_MB", "VR_Other_ProcessCount",
    "VR_Total_IncludingServer_ProcessCount", "VR_Total_IncludingServer_CPU_Percent",
    "VR_Total_IncludingServer_RAM_MB", "System_CPU_Percent", "System_CPU_Available",
    "System_RAM_Used_MB", "System_RAM_Free_MB", "System_RAM_Available",
    "System_Disk_BytesSec", "System_Disk_Available", "System_Net_BytesSec",
    "System_Net_Available", "ProcessPerf_Available", "RunnerActive"
)

$processColumns = @(
    "RunId", "Scenario", "Timestamp", "ElapsedSec", "ServiceStatus", "PID", "Role",
    "ProcessName", "ExecutablePath", "CommandLine", "CPU_Percent", "RAM_MB",
    "WorkingSet_MB", "PrivateMemory_MB", "Handles", "Threads", "IO_Read_BytesSec",
    "IO_Write_BytesSec", "IO_Total_BytesSec", "ProcessPerfAvailable", "RunnerActive"
)

Initialize-Csv -Path $samplesCsv -Columns $sampleColumns
Initialize-Csv -Path $processSamplesCsv -Columns $processColumns

Write-Info "RunId: $runId"
Write-Info "Scenario: $ScenarioName"
Write-Info "OutputDir: $runDir"
Write-Info "SkipServiceControl: $SkipServiceControl"

$serviceControlAvailable = $true
$serviceInitial = Get-VRServiceRecord

if ($SkipServiceControl) {
    Write-Info "Service control skipped by parameter. Only service status will be read."
} else {
    if ($StopVRBeforeRun) {
        if (-not (Invoke-VRServiceStop -ServiceRecord $serviceInitial)) { $serviceControlAvailable = $false }
    }

    if ($StartVRAfterStop) {
        $serviceForStart = Get-VRServiceRecord
        if (-not (Invoke-VRServiceStart -ServiceRecord $serviceForStart)) { $serviceControlAvailable = $false }
    }
}

if ($WarmupSec -gt 0) {
    Write-Info "WarmupSec: $WarmupSec"
    Start-Sleep -Seconds $WarmupSec
}

$serviceStart = Get-VRServiceRecord
Save-ProcessList -Path $processListStart -Label "START"
$startProcesses = @(Get-VRProcessRecords)
$clientStartCount = @($startProcesses | Where-Object { $_.Role -eq "CLIENT_SERVICE" }).Count
$serverStartCount = @($startProcesses | Where-Object { $_.Role -eq "SERVER_GUI" }).Count

$memAtStart = Get-SystemMemoryMetric
$envLines = @(
    "TFM Velociraptor resource benchmark environment",
    "RunId: $runId",
    "Scenario: $ScenarioName",
    "Time: $((Get-Date).ToString('o'))",
    "ComputerName: $env:COMPUTERNAME",
    "UserName: $env:USERNAME",
    "PowerShellVersion: $($PSVersionTable.PSVersion.ToString())",
    "DurationSec: $DurationSec",
    "IntervalSec: $IntervalSec",
    "WarmupSec: $WarmupSec",
    "RunnerPath: $RunnerPath",
    "OutputDir: $OutputDir",
    "StopVRBeforeRun: $StopVRBeforeRun",
    "StartVRAfterStop: $StartVRAfterStop",
    "SkipServiceControl: $SkipServiceControl",
    "IncludeServerGuiInTotal: $IncludeServerGuiInTotal",
    "ClientServiceName: $($serviceStart.Name)",
    "ClientServiceStatusStart: $($serviceStart.Status)",
    "ClientServicePath: $($serviceStart.PathName)",
    "VR_Client_ProcessCountStart: $clientStartCount",
    "VR_ServerGUI_ProcessCountStart: $serverStartCount",
    "SystemRAMTotalMB: $($memAtStart.TotalMB)"
)
Set-Content -LiteralPath $environmentTxt -Value $envLines -Encoding UTF8

$runnerProc = $null
$runnerStarted = $false
$runnerStartError = ""
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
        $runnerStartError = "RunnerPath does not exist."
        Add-Warn -Key "runner_missing" -Message "RunnerPath no existe: $RunnerPath"
    }
}

$samples = New-Object System.Collections.ArrayList
$runnerObservedSamples = 0
$start = Get-Date

while (((Get-Date) - $start).TotalSeconds -lt $DurationSec) {
    $now = Get-Date
    $elapsed = [math]::Round((($now - $start).TotalSeconds), 2)
    $serviceNow = Get-VRServiceRecord
    $vrProcesses = @(Get-VRProcessRecords)
    $perfByPid = Get-PerfProcByPid
    $processPerfAvailable = ($perfByPid.Count -gt 0)
    $runnerActive = if ($runnerProc) { -not $runnerProc.HasExited } else { $false }
    if ($runnerActive) { $runnerObservedSamples++ }

    $clientAgg = New-ZeroAggregate
    $serverAgg = New-ZeroAggregate
    $otherAgg = New-ZeroAggregate

    foreach ($vr in $vrProcesses) {
        $gp = $null
        try {
            $gp = Get-Process -Id $vr.PID -ErrorAction Stop
        } catch {
            continue
        }

        $perf = $null
        if ($perfByPid.ContainsKey($vr.PID)) {
            $perf = $perfByPid[$vr.PID]
        }

        $procPerfAvailable = ($null -ne $perf)
        $cpu = if ($procPerfAvailable) { [math]::Round([double]$perf.PercentProcessorTime, 2) } else { $null }
        $ioRead = if ($procPerfAvailable) { [math]::Round([double]$perf.IOReadBytesPersec, 2) } else { $null }
        $ioWrite = if ($procPerfAvailable) { [math]::Round([double]$perf.IOWriteBytesPersec, 2) } else { $null }
        $ioTotal = if ($procPerfAvailable) { [math]::Round([double]$perf.IODataBytesPersec, 2) } else { $null }
        $workingSetMb = [math]::Round(([double]$gp.WorkingSet64 / 1MB), 2)
        $privateMb = [math]::Round(([double]$gp.PrivateMemorySize64 / 1MB), 2)
        $handles = 0
        try { $handles = [int]$gp.HandleCount } catch { $handles = 0 }
        $threads = 0
        try { $threads = @($gp.Threads).Count } catch { $threads = 0 }

        $processRow = [pscustomobject]@{
            RunId                = $runId
            Scenario             = $ScenarioName
            Timestamp            = $now.ToString("o")
            ElapsedSec           = $elapsed
            ServiceStatus        = $serviceNow.Status
            PID                  = $vr.PID
            Role                 = $vr.Role
            ProcessName          = $vr.ProcessName
            ExecutablePath       = $vr.ExecutablePath
            CommandLine          = $vr.CommandLine
            CPU_Percent          = $cpu
            RAM_MB               = $workingSetMb
            WorkingSet_MB        = $workingSetMb
            PrivateMemory_MB     = $privateMb
            Handles              = $handles
            Threads              = $threads
            IO_Read_BytesSec     = $ioRead
            IO_Write_BytesSec    = $ioWrite
            IO_Total_BytesSec    = $ioTotal
            ProcessPerfAvailable = $procPerfAvailable
            RunnerActive         = $runnerActive
        }

        Add-CsvObject -Path $processSamplesCsv -Object $processRow -Columns $processColumns

        if ($vr.Role -eq "CLIENT_SERVICE") {
            Add-ToAggregate -Aggregate $clientAgg -ProcessRow $processRow
        } elseif ($vr.Role -eq "SERVER_GUI") {
            Add-ToAggregate -Aggregate $serverAgg -ProcessRow $processRow
        } elseif ($vr.Role -eq "OTHER_VELOCIRAPTOR") {
            Add-ToAggregate -Aggregate $otherAgg -ProcessRow $processRow
        }
    }

    $sysCpu = Get-SystemCpuMetric
    $sysRam = Get-SystemMemoryMetric
    $sysDisk = Get-SystemDiskMetric
    $sysNet = Get-SystemNetMetric

    $totalCount = $clientAgg.ProcessCount + $otherAgg.ProcessCount
    $totalCpu = [math]::Round(($clientAgg.CPU_Percent + $otherAgg.CPU_Percent), 2)
    $totalRam = [math]::Round(($clientAgg.RAM_MB + $otherAgg.RAM_MB), 2)
    if ($IncludeServerGuiInTotal) {
        $totalCount += $serverAgg.ProcessCount
        $totalCpu = [math]::Round(($totalCpu + $serverAgg.CPU_Percent), 2)
        $totalRam = [math]::Round(($totalRam + $serverAgg.RAM_MB), 2)
    }

    $sampleRow = [pscustomobject]@{
        RunId                                = $runId
        Scenario                             = $ScenarioName
        Timestamp                            = $now.ToString("o")
        ElapsedSec                           = $elapsed
        ServiceName                          = $serviceNow.Name
        ServiceStatus                        = $serviceNow.Status
        VR_Client_ProcessCount               = $clientAgg.ProcessCount
        VR_Client_PIDs                       = $clientAgg.PIDs
        VR_Client_CPU_Percent                = $clientAgg.CPU_Percent
        VR_Client_RAM_MB                     = $clientAgg.RAM_MB
        VR_Client_WorkingSet_MB              = $clientAgg.WorkingSet_MB
        VR_Client_PrivateMemory_MB           = $clientAgg.PrivateMemory_MB
        VR_Client_Handles                    = $clientAgg.Handles
        VR_Client_Threads                    = $clientAgg.Threads
        VR_Client_IO_Read_BytesSec           = $clientAgg.IO_Read_BytesSec
        VR_Client_IO_Write_BytesSec          = $clientAgg.IO_Write_BytesSec
        VR_Client_IO_Total_BytesSec          = $clientAgg.IO_Total_BytesSec
        VR_ServerGUI_ProcessCount            = $serverAgg.ProcessCount
        VR_ServerGUI_PIDs                    = $serverAgg.PIDs
        VR_ServerGUI_CPU_Percent             = $serverAgg.CPU_Percent
        VR_ServerGUI_RAM_MB                  = $serverAgg.RAM_MB
        VR_Other_ProcessCount                = $otherAgg.ProcessCount
        VR_Total_IncludingServer_ProcessCount = $totalCount
        VR_Total_IncludingServer_CPU_Percent  = $totalCpu
        VR_Total_IncludingServer_RAM_MB       = $totalRam
        System_CPU_Percent                   = $sysCpu.Value
        System_CPU_Available                 = $sysCpu.Available
        System_RAM_Used_MB                   = $sysRam.UsedMB
        System_RAM_Free_MB                   = $sysRam.FreeMB
        System_RAM_Available                 = $sysRam.Available
        System_Disk_BytesSec                 = $sysDisk.Value
        System_Disk_Available               = $sysDisk.Available
        System_Net_BytesSec                  = $sysNet.Value
        System_Net_Available                = $sysNet.Available
        ProcessPerf_Available               = $processPerfAvailable
        RunnerActive                         = $runnerActive
    }

    [void]$samples.Add($sampleRow)
    Add-CsvObject -Path $samplesCsv -Object $sampleRow -Columns $sampleColumns
    Start-Sleep -Seconds $IntervalSec
}

$serviceEnd = Get-VRServiceRecord
Save-ProcessList -Path $processListEnd -Label "END"
$endProcesses = @(Get-VRProcessRecords)
$clientEndCount = @($endProcesses | Where-Object { $_.Role -eq "CLIENT_SERVICE" }).Count
$serverDetected = ((@($samples | Where-Object { $_.VR_ServerGUI_ProcessCount -gt 0 }).Count) -gt 0)
$runnerStillRunning = $false
if ($runnerProc -and -not $runnerProc.HasExited) {
    $runnerStillRunning = $true
    Add-Warn -Key "runner_still_running" -Message "El runner sigue activo al finalizar la ventana de benchmark. No se mata el proceso."
}

$clientCpuStats = Get-Stats -Rows @($samples) -PropertyName "VR_Client_CPU_Percent"
$clientRamStats = Get-Stats -Rows @($samples) -PropertyName "VR_Client_RAM_MB"
$clientIoStats = Get-Stats -Rows @($samples) -PropertyName "VR_Client_IO_Total_BytesSec"
$clientCountStats = Get-Stats -Rows @($samples) -PropertyName "VR_Client_ProcessCount"
$serverCountStats = Get-Stats -Rows @($samples) -PropertyName "VR_ServerGUI_ProcessCount"
$sysCpuStats = Get-Stats -Rows @($samples) -PropertyName "System_CPU_Percent"
$sysRamStats = Get-Stats -Rows @($samples) -PropertyName "System_RAM_Used_MB"
$sysDiskStats = Get-Stats -Rows @($samples) -PropertyName "System_Disk_BytesSec"
$sysNetStats = Get-Stats -Rows @($samples) -PropertyName "System_Net_BytesSec"
$validity = Get-ScenarioValidity -Scenario $ScenarioName -ClientMax ([int]($clientCountStats.Max)) -RunnerObservedSamples $runnerObservedSamples

$metricsAvailability = [pscustomobject]@{
    SystemCPUAvailable       = ((@($samples | Where-Object { $_.System_CPU_Available -eq $true }).Count) -gt 0)
    SystemRAMAvailable       = ((@($samples | Where-Object { $_.System_RAM_Available -eq $true }).Count) -gt 0)
    SystemDiskAvailable      = ((@($samples | Where-Object { $_.System_Disk_Available -eq $true }).Count) -gt 0)
    SystemNetAvailable       = ((@($samples | Where-Object { $_.System_Net_Available -eq $true }).Count) -gt 0)
    ProcessPerfAvailable     = ((@($samples | Where-Object { $_.ProcessPerf_Available -eq $true }).Count) -gt 0)
    SystemCPUSource          = "Win32_PerfFormattedData_PerfOS_Processor fallback Win32_Processor"
    SystemDiskSource         = "Win32_PerfFormattedData_PerfDisk_PhysicalDisk"
    SystemNetSource          = "Win32_PerfFormattedData_Tcpip_NetworkInterface"
    ProcessPerfSource        = "Win32_PerfFormattedData_PerfProc_Process"
    UnavailableMetricsPolicy = "NA/null; excluded from averages"
}

$end = Get-Date
$summaryObj = [pscustomobject]@{
    SchemaVersion                       = "1.1"
    ScriptVersion                       = "TFM_Benchmark_VR_Resource_Usage_v1_role_fix"
    RunId                               = $runId
    Scenario                            = $ScenarioName
    ScenarioValidity                    = $validity.Validity
    ValidityReasons                     = @($validity.Reasons)
    StartTime                           = $start.ToString("o")
    EndTime                             = $end.ToString("o")
    DurationSec                         = $DurationSec
    ActualDurationSec                   = [math]::Round((($end - $start).TotalSeconds), 2)
    IntervalSec                         = $IntervalSec
    WarmupSec                           = $WarmupSec
    Samples                             = $samples.Count
    ClientServiceName                   = $serviceEnd.Name
    ClientServiceStatusStart            = $serviceStart.Status
    ClientServiceStatusEnd              = $serviceEnd.Status
    ClientServicePath                   = $serviceEnd.PathName
    ServiceControlSkipped               = $SkipServiceControl
    ServiceControlAvailable             = $serviceControlAvailable
    StopVRBeforeRun                     = $StopVRBeforeRun
    StartVRAfterStop                    = $StartVRAfterStop
    IncludeServerGuiInTotal             = $IncludeServerGuiInTotal
    VR_Client_ProcessCountStart         = $clientStartCount
    VR_Client_ProcessCountEnd           = $clientEndCount
    VR_Client_ProcessCountMax           = [int]($clientCountStats.Max)
    VR_ServerGUI_ProcessCountStart      = $serverStartCount
    VR_ServerGUI_ProcessCountMax        = [int]($serverCountStats.Max)
    ServerGuiDetected                   = $serverDetected
    ServerGuiExcludedFromClientMetrics  = $true
    RunnerPath                          = $RunnerPath
    RunnerStarted                       = $runnerStarted
    RunnerStartError                    = $runnerStartError
    RunnerObservedSamples               = $runnerObservedSamples
    RunnerStillRunningAtEnd             = $runnerStillRunning
    MetricsAvailability                 = $metricsAvailability
    Avg_VR_Client_CPU_Percent           = $clientCpuStats.Avg
    Max_VR_Client_CPU_Percent           = $clientCpuStats.Max
    P95_VR_Client_CPU_Percent           = $clientCpuStats.P95
    Avg_VR_Client_RAM_MB                = $clientRamStats.Avg
    Max_VR_Client_RAM_MB                = $clientRamStats.Max
    P95_VR_Client_RAM_MB                = $clientRamStats.P95
    Avg_VR_Client_IO_BytesSec           = $clientIoStats.Avg
    Max_VR_Client_IO_BytesSec           = $clientIoStats.Max
    Avg_System_CPU_Percent              = $sysCpuStats.Avg
    Max_System_CPU_Percent              = $sysCpuStats.Max
    Avg_System_RAM_Used_MB              = $sysRamStats.Avg
    Max_System_RAM_Used_MB              = $sysRamStats.Max
    Avg_Disk_BytesSec                   = $sysDiskStats.Avg
    Max_Disk_BytesSec                   = $sysDiskStats.Max
    Avg_Net_BytesSec                    = $sysNetStats.Avg
    Max_Net_BytesSec                    = $sysNetStats.Max
    SamplesCsv                          = $samplesCsv
    ProcessSamplesCsv                   = $processSamplesCsv
    ProcessListStart                    = $processListStart
    ProcessListEnd                      = $processListEnd
    EnvironmentTxt                      = $environmentTxt
    Warnings                            = @($script:Warnings)
}

$summaryObj | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $summaryJson -Encoding UTF8

$summaryLines = @(
    "TFM Velociraptor Resource Benchmark",
    "",
    "RunId: $runId",
    "Scenario: $ScenarioName",
    "ScenarioValidity: $($summaryObj.ScenarioValidity)",
    "ValidityReasons: $(@($summaryObj.ValidityReasons) -join ' | ')",
    "StartTime: $($start.ToString('o'))",
    "EndTime: $($end.ToString('o'))",
    "DurationSec configured: $DurationSec",
    "ActualDurationSec: $($summaryObj.ActualDurationSec)",
    "IntervalSec: $IntervalSec",
    "Samples: $($samples.Count)",
    "",
    "ClientServiceName: $($summaryObj.ClientServiceName)",
    "ClientServiceStatusStart: $($summaryObj.ClientServiceStatusStart)",
    "ClientServiceStatusEnd: $($summaryObj.ClientServiceStatusEnd)",
    "ClientServicePath: $($summaryObj.ClientServicePath)",
    "ServiceControlSkipped: $($summaryObj.ServiceControlSkipped)",
    "ServiceControlAvailable: $($summaryObj.ServiceControlAvailable)",
    "",
    "VR_Client_ProcessCountStart: $($summaryObj.VR_Client_ProcessCountStart)",
    "VR_Client_ProcessCountEnd: $($summaryObj.VR_Client_ProcessCountEnd)",
    "VR_Client_ProcessCountMax: $($summaryObj.VR_Client_ProcessCountMax)",
    "ServerGuiDetected: $($summaryObj.ServerGuiDetected)",
    "ServerGuiExcludedFromClientMetrics: $($summaryObj.ServerGuiExcludedFromClientMetrics)",
    "",
    "Avg_VR_Client_CPU_Percent: $(Convert-ToCsvValue $summaryObj.Avg_VR_Client_CPU_Percent)",
    "Max_VR_Client_CPU_Percent: $(Convert-ToCsvValue $summaryObj.Max_VR_Client_CPU_Percent)",
    "P95_VR_Client_CPU_Percent: $(Convert-ToCsvValue $summaryObj.P95_VR_Client_CPU_Percent)",
    "Avg_VR_Client_RAM_MB: $(Convert-ToCsvValue $summaryObj.Avg_VR_Client_RAM_MB)",
    "Max_VR_Client_RAM_MB: $(Convert-ToCsvValue $summaryObj.Max_VR_Client_RAM_MB)",
    "P95_VR_Client_RAM_MB: $(Convert-ToCsvValue $summaryObj.P95_VR_Client_RAM_MB)",
    "Avg_VR_Client_IO_BytesSec: $(Convert-ToCsvValue $summaryObj.Avg_VR_Client_IO_BytesSec)",
    "Max_VR_Client_IO_BytesSec: $(Convert-ToCsvValue $summaryObj.Max_VR_Client_IO_BytesSec)",
    "",
    "RunnerPath: $RunnerPath",
    "RunnerStarted: $runnerStarted",
    "RunnerObservedSamples: $runnerObservedSamples",
    "RunnerStillRunningAtEnd: $runnerStillRunning",
    "",
    "MetricsAvailability:",
    ($metricsAvailability | ConvertTo-Json -Compress),
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
