# created script using Cursor AI
#
#Requires -Version 5.1
<#
.SYNOPSIS
    Collects CPU, memory, network, disk, and NUMA performance metrics on Windows Server 2022+.

.DESCRIPTION
    Uses Get-Counter (Performance Monitor API) to sample metrics and writes timestamped CSV
    files suitable for Excel analysis. Supports snapshot and continuous collection modes,
    with summary and/or detailed per-instance breakdowns.

.PARAMETER Mode
    Snapshot: collect one sample. Continuous: poll at IntervalSeconds until duration or Ctrl+C.

.PARAMETER IntervalSeconds
    Seconds between samples in Continuous mode.

.PARAMETER DurationSeconds
    Stop after this many seconds in Continuous mode. 0 runs until Ctrl+C or MaxSamples.

.PARAMETER DetailLevel
    Summary, Detailed, or Both (writes separate CSV files when Both).

.PARAMETER OutputDirectory
    Directory for CSV and topology output files.

.PARAMETER OutputPrefix
    Prefix for output filenames.

.PARAMETER IncludeNuma
    Include NUMA performance counters when available. Pass -IncludeNuma $false to disable.

.PARAMETER MaxSamples
    Maximum number of samples in Continuous mode. 0 means unlimited.

.EXAMPLE
    .\Collect-WindowsPerfMetrics.ps1 -Mode Snapshot -DetailLevel Both -OutputDirectory C:\perf

.EXAMPLE
    .\Collect-WindowsPerfMetrics.ps1 -Mode Continuous -IntervalSeconds 5 -DurationSeconds 600 -DetailLevel Summary
#>
[CmdletBinding()]
param(
    [ValidateSet('Snapshot', 'Continuous')]
    [string]$Mode = 'Snapshot',

    [ValidateRange(1, [int]::MaxValue)]
    [int]$IntervalSeconds = 5,

    [ValidateRange(0, [int]::MaxValue)]
    [int]$DurationSeconds = 0,

    [ValidateSet('Summary', 'Detailed', 'Both')]
    [string]$DetailLevel = 'Both',

    [string]$OutputDirectory = '.',

    [string]$OutputPrefix = 'winperf',

    [bool]$IncludeNuma = $true,

    [ValidateRange(0, [int]::MaxValue)]
    [int]$MaxSamples = 0
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Helpers

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $formatted = "[$timestamp] [$Level] $Message"
    switch ($Level) {
        'Warning' { Write-Warning $Message }
        'Error'   { Write-Error $Message }
        default   { Write-Host $formatted }
    }
}

function ConvertTo-SafeColumnName {
    param([string]$Name)
    $safe = $Name -replace '[\\()/\s%\.]', '_'
    $safe = $safe -replace '_+', '_'
    $safe.Trim('_')
}

function Get-FormattedTimestamp {
    param([datetime]$DateTime = (Get-Date))
    $DateTime.ToString('yyyy-MM-dd HH:mm:ss.fff')
}

function Normalize-CounterPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $Path
    }
    # Get-Counter returns \\HOSTNAME\Set(instance)\Counter; map keys use \Set(instance)\Counter
    if ($Path -match '^\\\\[^\\]+(\\.+)$') {
        return $Matches[1]
    }
    if ($Path -notmatch '^\\') {
        return "\$Path"
    }
    return $Path
}

function Get-CounterPathKey {
    param([string]$Path)
    (Normalize-CounterPath -Path $Path).ToLowerInvariant()
}

function Get-CollectionCount {
    param(
        [AllowNull()]
        $Items
    )
    if ($null -eq $Items) {
        return 0
    }
    @($Items).Count
}

function Ensure-Array {
    param(
        [AllowNull()]
        $Items
    )
    if ($null -eq $Items) {
        return @()
    }
    @($Items)
}

function Test-NetworkInterfaceExcluded {
    param([string]$InstanceName)
    $lower = $InstanceName.ToLowerInvariant()
    $excludePatterns = @(
        'loopback',
        'isatap',
        'teredo',
        '6to4',
        'microsoft wi-fi direct',
        'bluetooth',
        'qos packet scheduler',
        'wfp'
    )
    foreach ($pattern in $excludePatterns) {
        if ($lower -like "*$pattern*") {
            return $true
        }
    }
    return $false
}

function Get-AvailableCounterPaths {
    param([string[]]$Paths)
    $Paths = Ensure-Array -Items $Paths
    if ((Get-CollectionCount -Items $Paths) -eq 0) {
        return @()
    }

    $available = [System.Collections.Generic.List[string]]::new()
    $batchSize = 50
    for ($i = 0; $i -lt (Get-CollectionCount -Items $Paths); $i += $batchSize) {
        $batch = $Paths[$i..([Math]::Min($i + $batchSize - 1, (Get-CollectionCount -Items $Paths) - 1))]
        try {
            $null = Get-Counter -Counter $batch -SampleInterval 1 -MaxSamples 1 -ErrorAction Stop
            foreach ($path in $batch) {
                $available.Add($path) | Out-Null
            }
        }
        catch {
            foreach ($path in $batch) {
                try {
                    $null = Get-Counter -Counter $path -SampleInterval 1 -MaxSamples 1 -ErrorAction Stop
                    $available.Add($path) | Out-Null
                }
                catch {
                    Write-Log "Counter unavailable, skipping: $path" -Level Warning
                }
            }
        }
    }
    return @($available)
}

function Get-CounterResultSamples {
    param(
        [AllowNull()]
        $CounterResult
    )
    if ($null -eq $CounterResult) {
        return @()
    }
    if ($null -ne $CounterResult.PSObject.Properties['CounterSamples']) {
        return Ensure-Array -Items $CounterResult.CounterSamples
    }
    if ($null -ne $CounterResult.PSObject.Properties['Path']) {
        return @($CounterResult)
    }
    return @()
}

function Test-CounterInstanceIncluded {
    param(
        [string]$SetName,
        [string]$InstanceName,
        [string]$InstancePattern,
        [string[]]$ExcludeInstances
    )
    if ($InstancePattern -ne '*' -and $InstanceName -notlike $InstancePattern) {
        return $false
    }
    if ($ExcludeInstances -contains $InstanceName) {
        return $false
    }
    if ($SetName -eq 'Network Interface' -and (Test-NetworkInterfaceExcluded -InstanceName $InstanceName)) {
        return $false
    }
    return $true
}

function Get-CounterInstances {
    param(
        [string]$SetName,
        [string]$ProbeCounter,
        [string]$InstancePattern = '*',
        [string[]]$ExcludeInstances = @()
    )
    $instances = [System.Collections.Generic.List[string]]::new()
    $wildcardPath = ('\{0}(*)\{1}' -f $SetName, $ProbeCounter)
    try {
        $result = Get-Counter -Counter $wildcardPath -SampleInterval 1 -MaxSamples 1 -ErrorAction Stop
        foreach ($sample in (Get-CounterResultSamples -CounterResult $result)) {
            $instanceName = $sample.InstanceName
            if (-not (Test-CounterInstanceIncluded -SetName $SetName -InstanceName $instanceName `
                    -InstancePattern $InstancePattern -ExcludeInstances $ExcludeInstances)) {
                continue
            }
            if ($instances -notcontains $instanceName) {
                $instances.Add($instanceName) | Out-Null
            }
        }
    }
    catch {
        Write-Log "Counter set unavailable or not enumerable: $SetName ($ProbeCounter) - $_" -Level Warning
    }
    return @($instances)
}

function Get-CounterPathsFromWildcard {
    param(
        [string]$SetName,
        [string]$InstancePattern,
        [string[]]$Counters,
        [string[]]$ExcludeInstances = @()
    )
    $Counters = Ensure-Array -Items $Counters
    if ((Get-CollectionCount -Items $Counters) -eq 0) {
        return @()
    }

    # ListSet.Paths may contain wildcard paths like \Processor(*)\% Processor Time,
    # not instance names. Discover real instances via a probe counter query.
    $instances = Get-CounterInstances -SetName $SetName -ProbeCounter $Counters[0] `
        -InstancePattern $InstancePattern -ExcludeInstances $ExcludeInstances
    if ((Get-CollectionCount -Items $instances) -eq 0) {
        return @()
    }

    $paths = [System.Collections.Generic.List[string]]::new()
    foreach ($instance in $instances) {
        foreach ($counter in $Counters) {
            $paths.Add(('\{0}({1})\{2}' -f $SetName, $instance, $counter)) | Out-Null
        }
    }
    return @($paths)
}

#endregion

#region Counter Path Builders

function Get-SummaryCounterDefinitions {
    param([bool]$NumaEnabled)

    $direct = @(
        @{ Path = '\Processor(_Total)\% Processor Time';           Column = 'CPU_Total_ProcessorTime' },
        @{ Path = '\Processor(_Total)\% Privileged Time';           Column = 'CPU_Total_PrivilegedTime' },
        @{ Path = '\Processor(_Total)\% User Time';               Column = 'CPU_Total_UserTime' },
        @{ Path = '\Processor(_Total)\% Idle Time';                Column = 'CPU_Total_IdleTime' },
        @{ Path = '\System\Processor Queue Length';               Column = 'CPU_ProcessorQueueLength' },
        @{ Path = '\System\Context Switches/sec';                 Column = 'CPU_ContextSwitchesPerSec' },
        @{ Path = '\Memory\Available MBytes';                     Column = 'Mem_AvailableMB' },
        @{ Path = '\Memory\Committed Bytes';                      Column = 'Mem_CommittedBytes' },
        @{ Path = '\Memory\Commit Limit';                          Column = 'Mem_CommitLimit' },
        @{ Path = '\Memory\Pages/sec';                             Column = 'Mem_PagesPerSec' },
        @{ Path = '\Memory\Page Faults/sec';                       Column = 'Mem_PageFaultsPerSec' },
        @{ Path = '\Memory\Pool Nonpaged Bytes';                   Column = 'Mem_PoolNonpagedBytes' },
        @{ Path = '\Memory\Pool Paged Bytes';                      Column = 'Mem_PoolPagedBytes' },
        @{ Path = '\PhysicalDisk(_Total)\Disk Reads/sec';          Column = 'Disk_Total_ReadsPerSec' },
        @{ Path = '\PhysicalDisk(_Total)\Disk Writes/sec';          Column = 'Disk_Total_WritesPerSec' },
        @{ Path = '\PhysicalDisk(_Total)\Disk Bytes/sec';           Column = 'Disk_Total_BytesPerSec' },
        @{ Path = '\PhysicalDisk(_Total)\% Disk Time';              Column = 'Disk_Total_DiskTimePct' },
        @{ Path = '\PhysicalDisk(_Total)\Avg. Disk sec/Read';       Column = 'Disk_Total_AvgSecPerRead' },
        @{ Path = '\PhysicalDisk(_Total)\Avg. Disk sec/Write';      Column = 'Disk_Total_AvgSecPerWrite' },
        @{ Path = '\PhysicalDisk(_Total)\Current Disk Queue Length'; Column = 'Disk_Total_QueueLength' }
    )

    if ($NumaEnabled) {
        $direct += @(
            @{ Path = '\NUMA Node(_Total)\% Processor Time';  Column = 'NUMA_Total_ProcessorTime' },
            @{ Path = '\NUMA Node(_Total)\Processor Utility'; Column = 'NUMA_Total_ProcessorUtility' },
            @{ Path = '\NUMA Node(_Total)\Remote Frees/sec';  Column = 'NUMA_Total_RemoteFreesPerSec' }
        )
    }

    $networkCounters = @(
        'Bytes Received/sec',
        'Bytes Sent/sec',
        'Bytes Total/sec',
        'Packets/sec'
    )
    $networkPaths = Get-CounterPathsFromWildcard -SetName 'Network Interface' -InstancePattern '*' -Counters $networkCounters

    return @{
        Direct   = $direct
        Network  = $networkPaths
        AggregateNetwork = @{
            'Bytes Received/sec' = 'Net_Total_BytesReceivedPerSec'
            'Bytes Sent/sec'     = 'Net_Total_BytesSentPerSec'
            'Bytes Total/sec'    = 'Net_Total_BytesTotalPerSec'
            'Packets/sec'        = 'Net_Total_PacketsPerSec'
        }
    }
}

function Get-DetailedCounterDefinitions {
    param([bool]$NumaEnabled)

    $cpuPaths = Get-CounterPathsFromWildcard `
        -SetName 'Processor' `
        -InstancePattern '*' `
        -Counters @('% Processor Time') `
        -ExcludeInstances @('_Total', 'Idle')

    $networkCounters = @(
        'Bytes Received/sec',
        'Bytes Sent/sec',
        'Packets/sec',
        'Output Queue Length'
    )
    $networkPaths = Get-CounterPathsFromWildcard `
        -SetName 'Network Interface' `
        -InstancePattern '*' `
        -Counters $networkCounters

    $diskCounters = @(
        'Disk Reads/sec',
        'Disk Writes/sec',
        'Disk Bytes/sec',
        '% Disk Time',
        'Current Disk Queue Length'
    )
    $diskPaths = Get-CounterPathsFromWildcard `
        -SetName 'PhysicalDisk' `
        -InstancePattern '*' `
        -Counters $diskCounters `
        -ExcludeInstances @('_Total')

    $numaPaths = @()
    if ($NumaEnabled) {
        $numaPaths = Get-CounterPathsFromWildcard `
            -SetName 'NUMA Node' `
            -InstancePattern '*' `
            -Counters @('% Processor Time', 'Processor Utility', 'Remote Frees/sec') `
            -ExcludeInstances @('_Total')
    }

    return @{
        Cpu     = $cpuPaths
        Network = $networkPaths
        Disk    = $diskPaths
        Numa    = $numaPaths
    }
}

function Get-DetailedColumnName {
    param(
        [string]$Path
    )
    $Path = Normalize-CounterPath -Path $Path
    if ($Path -match '\\([^\\]+)\(([^)]+)\)\\(.+)$') {
        $setName = $Matches[1]
        $instance = ConvertTo-SafeColumnName -Name $Matches[2]
        $counter = ConvertTo-SafeColumnName -Name $Matches[3]

        switch ($setName) {
            'Processor'         { return "CPU_${instance}_${counter}" }
            'Network Interface' { return "Net_${instance}_${counter}" }
            'PhysicalDisk'      { return "Disk_${instance}_${counter}" }
            'NUMA Node'         { return "NUMA_${instance}_${counter}" }
            default             { return "$(ConvertTo-SafeColumnName -Name $setName)_${instance}_${counter}" }
        }
    }
    return ConvertTo-SafeColumnName -Name $Path
}

function Build-SummaryColumnMap {
    param($Definitions)
    $map = @{}
    foreach ($entry in $Definitions.Direct) {
        $map[(Get-CounterPathKey -Path $entry.Path)] = $entry.Column
    }
    foreach ($path in $Definitions.Network) {
        if ($path -match '\\([^\\]+)$') {
            $counterName = $Matches[1]
            if ($Definitions.AggregateNetwork.ContainsKey($counterName)) {
                $map[(Get-CounterPathKey -Path $path)] = "__AGGREGATE__$counterName"
            }
        }
    }
    return $map
}

#endregion

#region Sampling and Row Building

function Get-CounterSampleBatch {
    param([string[]]$Paths)
    $Paths = Ensure-Array -Items $Paths
    if ((Get-CollectionCount -Items $Paths) -eq 0) {
        return @()
    }

    # Get-Counter has a limit on paths per call; batch if needed
    $batchSize = 100
    $allSamples = [System.Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt (Get-CollectionCount -Items $Paths); $i += $batchSize) {
        $batch = $Paths[$i..([Math]::Min($i + $batchSize - 1, (Get-CollectionCount -Items $Paths) - 1))]
        try {
            $result = Get-Counter -Counter $batch -SampleInterval 1 -MaxSamples 1 -ErrorAction Stop
            foreach ($sample in (Get-CounterResultSamples -CounterResult $result)) {
                $allSamples.Add($sample) | Out-Null
            }
        }
        catch {
            Write-Log "Batch counter read failed (batch starting at index $i): $_" -Level Warning
            foreach ($path in $batch) {
                try {
                    $result = Get-Counter -Counter $path -SampleInterval 1 -MaxSamples 1 -ErrorAction Stop
                    foreach ($sample in (Get-CounterResultSamples -CounterResult $result)) {
                        $allSamples.Add($sample) | Out-Null
                    }
                }
                catch {
                    Write-Log "Counter read failed: $path" -Level Warning
                }
            }
        }
    }
    return @($allSamples)
}

function ConvertTo-SummaryRow {
    param(
        $Samples,
        [hashtable]$ColumnMap,
        [hashtable]$AggregateNetworkColumns,
        [datetime]$SampleTime
    )

    $row = [ordered]@{ Timestamp = Get-FormattedTimestamp -DateTime $SampleTime }
    $aggregates = @{}
    foreach ($key in $AggregateNetworkColumns.Values) {
        $aggregates[$key] = 0.0
    }

    foreach ($sample in (Ensure-Array -Items $Samples)) {
        $pathKey = Get-CounterPathKey -Path $sample.Path
        if (-not $ColumnMap.ContainsKey($pathKey)) {
            continue
        }
        $column = $ColumnMap[$pathKey]
        $value = [math]::Round($sample.CookedValue, 4)

        if ($column -like '__AGGREGATE__*') {
            $counterName = $column -replace '^__AGGREGATE__', ''
            $aggColumn = $AggregateNetworkColumns[$counterName]
            if ($aggColumn) {
                $aggregates[$aggColumn] += $value
            }
        }
        else {
            $row[$column] = $value
        }
    }

    foreach ($aggColumn in $aggregates.Keys) {
        $row[$aggColumn] = [math]::Round($aggregates[$aggColumn], 4)
    }

    # Ensure all expected direct columns exist (fill missing with empty/null)
    foreach ($col in ($ColumnMap.Values | Where-Object { $_ -notlike '__AGGREGATE__*' } | Select-Object -Unique)) {
        if (-not $row.Contains($col)) {
            $row[$col] = $null
        }
    }
    foreach ($aggCol in $AggregateNetworkColumns.Values) {
        if (-not $row.Contains($aggCol)) {
            $row[$aggCol] = 0.0
        }
    }

    return [PSCustomObject]$row
}

function ConvertTo-DetailedRow {
    param(
        $Samples,
        [hashtable]$ColumnMap,
        [datetime]$SampleTime
    )

    $row = [ordered]@{ Timestamp = Get-FormattedTimestamp -DateTime $SampleTime }

    foreach ($column in ($ColumnMap.Values | Sort-Object)) {
        $row[$column] = $null
    }

    foreach ($sample in (Ensure-Array -Items $Samples)) {
        $pathKey = Get-CounterPathKey -Path $sample.Path
        if (-not $ColumnMap.ContainsKey($pathKey)) {
            continue
        }
        $column = $ColumnMap[$pathKey]
        $row[$column] = [math]::Round($sample.CookedValue, 4)
    }

    return [PSCustomObject]$row
}

#endregion

#region NUMA Topology

function Export-NumaTopology {
    param(
        [string]$OutputDirectory,
        [string]$OutputPrefix,
        [string]$RunTimestamp
    )

    $lines = @()
    $lines += "NUMA Topology Report"
    $lines += "Generated: $(Get-FormattedTimestamp)"
    $lines += ""

    try {
        $numaNodes = @(Get-CimInstance -ClassName Win32_NumaNode -ErrorAction Stop)
        $lines += "NUMA Node Count: $(Get-CollectionCount -Items $numaNodes)"
        $lines += ""
        foreach ($node in ($numaNodes | Sort-Object NodeId)) {
            $lines += "Node ID: $($node.NodeId)"
            $lines += "  Reliability: $($node.Reliability)"
            $lines += "  Capacity: $($node.Capacity) bytes"
            $lines += "  Available: $($node.Available) bytes"
            $lines += "  Processor Count: $($node.NumberOfProcessors)"
            $lines += ""
        }
    }
    catch {
        $lines += "Win32_NumaNode unavailable: $_"
        $lines += ""
    }

    try {
        $processors = @(Get-CimInstance -ClassName Win32_Processor -ErrorAction Stop)
        $lines += "Processor Topology:"
        foreach ($proc in $processors) {
            $lines += "  $($proc.DeviceID): $($proc.Name)"
            $lines += "    Socket: $($proc.SocketDesignation), Cores: $($proc.NumberOfCores), Logical: $($proc.NumberOfLogicalProcessors)"
            $lines += "    Max Clock: $($proc.MaxClockSpeed) MHz, Current: $($proc.CurrentClockSpeed) MHz"
        }
        $lines += ""
    }
    catch {
        $lines += "Win32_Processor unavailable: $_"
        $lines += ""
    }

    $topologyText = $lines -join [Environment]::NewLine
    Write-Host $topologyText

    $topologyPath = Join-Path $OutputDirectory "${OutputPrefix}_numa_topology_${RunTimestamp}.txt"
    $topologyText | Out-File -FilePath $topologyPath -Encoding UTF8
    Write-Log "NUMA topology written to: $topologyPath"
}

#endregion

#region CSV Output

function Get-SummaryColumnOrder {
    param(
        [hashtable]$ColumnMap,
        [hashtable]$AggregateNetworkColumns
    )
    $columns = [System.Collections.Generic.List[string]]::new()
    $columns.Add('Timestamp') | Out-Null
    foreach ($col in ($ColumnMap.Values | Where-Object { $_ -notlike '__AGGREGATE__*' } | Select-Object -Unique | Sort-Object)) {
        $columns.Add($col) | Out-Null
    }
    foreach ($col in ($AggregateNetworkColumns.Values | Sort-Object)) {
        if ($columns -notcontains $col) {
            $columns.Add($col) | Out-Null
        }
    }
    return @($columns)
}

function Get-DetailedColumnOrder {
    param([hashtable]$ColumnMap)
    @('Timestamp') + @($ColumnMap.Values | Select-Object -Unique | Sort-Object)
}

function ConvertTo-CsvDataLine {
    param(
        $Row,
        [string[]]$ColumnOrder
    )
    $obj = [ordered]@{}
    foreach ($name in $ColumnOrder) {
        $value = $null
        if ($null -ne $Row -and $null -ne $Row.PSObject.Properties[$name]) {
            $value = $Row.$name
        }
        $obj[$name] = $value
    }
    return ([PSCustomObject]$obj | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 -First 1)
}

function Initialize-MetricsCsvStream {
    param(
        [string]$FilePath,
        [string[]]$ColumnOrder
    )
    $headerObj = [ordered]@{}
    foreach ($name in $ColumnOrder) {
        $headerObj[$name] = $null
    }
    $headerLine = ([PSCustomObject]$headerObj | ConvertTo-Csv -NoTypeInformation | Select-Object -First 1)

    $utf8Bom = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllText($FilePath, $headerLine + [Environment]::NewLine, $utf8Bom)

    return @{
        FilePath       = $FilePath
        ColumnOrder    = $ColumnOrder
        RowCount       = 0
        AppendEncoding = New-Object System.Text.UTF8Encoding $false
    }
}

function Append-MetricsCsvRow {
    param(
        [hashtable]$Stream,
        $Row
    )
    $line = ConvertTo-CsvDataLine -Row $Row -ColumnOrder $Stream.ColumnOrder
    [System.IO.File]::AppendAllText(
        $Stream.FilePath,
        $line + [Environment]::NewLine,
        $Stream.AppendEncoding
    )
    $Stream.RowCount++
}

function Write-MetricsCsv {
    param(
        [object[]]$Rows,
        [string]$FilePath
    )

    $Rows = Ensure-Array -Items $Rows
    if ((Get-CollectionCount -Items $Rows) -eq 0) {
        Write-Log "No rows to write for: $FilePath" -Level Warning
        return
    }

    $propertyNames = @(
        $Rows |
            ForEach-Object { $_.PSObject.Properties.Name } |
            Select-Object -Unique
    )
    $orderedProps = @('Timestamp') + @($propertyNames | Where-Object { $_ -ne 'Timestamp' } | Sort-Object)

    $normalizedRows = foreach ($row in $Rows) {
        $obj = [ordered]@{}
        foreach ($name in $orderedProps) {
            $obj[$name] = $row.$name
        }
        [PSCustomObject]$obj
    }

    $csv = $normalizedRows | ConvertTo-Csv -NoTypeInformation
    $utf8Bom = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllLines($FilePath, $csv, $utf8Bom)
    Write-Log "Wrote $(Get-CollectionCount -Items $Rows) row(s) to: $FilePath"
}

function Wait-Interruptible {
    param([int]$Seconds)
    if ($Seconds -le 0) {
        return
    }
    $deadline = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $deadline -and -not $script:StopRequested) {
        Start-Sleep -Milliseconds 200
    }
}

function Write-BufferedMetricsCsv {
    param(
        [bool]$WriteSummary,
        [bool]$WriteDetailed,
        [string]$OutputDirectory,
        [string]$OutputPrefix,
        [string]$RunTimestamp
    )

    if ($WriteSummary -and $script:summaryRows.Count -gt 0) {
        $summaryFile = Join-Path $OutputDirectory "${OutputPrefix}_summary_${RunTimestamp}.csv"
        Write-MetricsCsv -Rows @($script:summaryRows) -FilePath $summaryFile
    }

    if ($WriteDetailed -and $script:detailedRows.Count -gt 0) {
        $detailedFile = Join-Path $OutputDirectory "${OutputPrefix}_detailed_${RunTimestamp}.csv"
        Write-MetricsCsv -Rows @($script:detailedRows) -FilePath $detailedFile
    }
}

#endregion

#region Validation

function Test-CounterAvailability {
    param(
        [string[]]$SummaryPaths,
        [string[]]$DetailedPaths
    )

    $testPaths = @(
        '\Processor(_Total)\% Processor Time',
        '\Memory\Available MBytes',
        '\PhysicalDisk(_Total)\Disk Reads/sec'
    )

    $failures = 0
    foreach ($path in $testPaths) {
        try {
            $null = Get-Counter -Counter $path -SampleInterval 1 -MaxSamples 1 -ErrorAction Stop
            Write-Log "Self-check OK: $path"
        }
        catch {
            Write-Log "Self-check FAILED: $path - $_" -Level Warning
            $failures++
        }
    }

    if ((Get-CollectionCount -Items $SummaryPaths) -eq 0 -and (Get-CollectionCount -Items $DetailedPaths) -eq 0) {
        throw 'No performance counter paths available. Ensure you are running on Windows Server 2022+ with sufficient privileges.'
    }

    if ($failures -eq (Get-CollectionCount -Items $testPaths)) {
        throw 'Core performance counters are unavailable. Run as Administrator or check PerfMon service.'
    }
}

#endregion

#region Main

$script:StopRequested = $false

if (-not (Test-Path -LiteralPath $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}
$OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path

$runTimestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
Write-Log "Windows Performance Metrics Collection starting (Mode=$Mode, DetailLevel=$DetailLevel)"

if ($IncludeNuma) {
    Export-NumaTopology -OutputDirectory $OutputDirectory -OutputPrefix $OutputPrefix -RunTimestamp $runTimestamp
}

$collectSummary = $DetailLevel -in @('Summary', 'Both')
$collectDetailed = $DetailLevel -in @('Detailed', 'Both')

$summaryDefs = $null
$detailedDefs = $null
$summaryColumnMap = $null
$detailedColumnMap = $null
$summaryPaths = @()
$detailedPaths = @()

if ($collectSummary) {
    $summaryDefs = Get-SummaryCounterDefinitions -NumaEnabled:$IncludeNuma
    $directPaths = $summaryDefs.Direct | ForEach-Object { $_.Path }
    $validatedDirect = Get-AvailableCounterPaths -Paths $directPaths
    $summaryDefs.Direct = @($summaryDefs.Direct | Where-Object { $validatedDirect -contains $_.Path })
    $validatedNetwork = Get-AvailableCounterPaths -Paths $summaryDefs.Network
    $summaryDefs.Network = $validatedNetwork
    $summaryColumnMap = Build-SummaryColumnMap -Definitions $summaryDefs
    $summaryPaths = @(Ensure-Array -Items ($summaryDefs.Direct | ForEach-Object { $_.Path })) + @(Ensure-Array -Items $summaryDefs.Network)
    Write-Log "Summary counters: $(Get-CollectionCount -Items $summaryPaths) paths"
}

if ($collectDetailed) {
    $detailedDefs = Get-DetailedCounterDefinitions -NumaEnabled:$IncludeNuma
    $detailedPaths = @($detailedDefs.Cpu) + @($detailedDefs.Network) + @($detailedDefs.Disk) + @($detailedDefs.Numa)
    $detailedPaths = Get-AvailableCounterPaths -Paths $detailedPaths
    $detailedColumnMap = @{}
    foreach ($path in $detailedPaths) {
        $detailedColumnMap[(Get-CounterPathKey -Path $path)] = Get-DetailedColumnName -Path $path
    }
    Write-Log "Detailed counters: $(Get-CollectionCount -Items $detailedPaths) paths"
}

Test-CounterAvailability -SummaryPaths $summaryPaths -DetailedPaths $detailedPaths

$script:collectSummary = $collectSummary
$script:collectDetailed = $collectDetailed
$script:summaryDefs = $summaryDefs
$script:summaryColumnMap = $summaryColumnMap
$script:summaryPaths = $summaryPaths
$script:detailedColumnMap = $detailedColumnMap
$script:detailedPaths = $detailedPaths
$script:summaryRows = [System.Collections.Generic.List[object]]::new()
$script:detailedRows = [System.Collections.Generic.List[object]]::new()

$script:SampleCount = 0
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {
    [Console]::TreatControlCAsInput = $false
}
catch {
    # Not available in all hosts
}

$cancelHandler = {
    param($sender, $eventArgs)
    $eventArgs.Cancel = $true
    $script:StopRequested = $true
    Write-Warning 'Stop requested (Ctrl+C). Finishing current sample; collected data is already saved to CSV.'
}
try {
    [Console]::CancelKeyPress += $cancelHandler
}
catch {
    Write-Log 'Unable to register Ctrl+C handler in this host.' -Level Warning
}

function Invoke-SampleCollection {
    $sampleTime = Get-Date

    if ($script:collectSummary -and (Get-CollectionCount -Items $script:summaryPaths) -gt 0) {
        $samples = Get-CounterSampleBatch -Paths $script:summaryPaths
        $matched = Get-CollectionCount -Items @(
            Ensure-Array -Items $samples | ForEach-Object { Get-CounterPathKey -Path $_.Path } |
                Where-Object { $script:summaryColumnMap.ContainsKey($_) }
        )
        Write-Log "Summary sample: $(Get-CollectionCount -Items $samples) counters read, $matched matched"
        $row = ConvertTo-SummaryRow `
            -Samples $samples `
            -ColumnMap $script:summaryColumnMap `
            -AggregateNetworkColumns $script:summaryDefs.AggregateNetwork `
            -SampleTime $sampleTime
        $script:summaryRows.Add($row) | Out-Null
        if ($null -ne $script:summaryCsvStream) {
            Append-MetricsCsvRow -Stream $script:summaryCsvStream -Row $row
        }
    }

    if ($script:collectDetailed -and (Get-CollectionCount -Items $script:detailedPaths) -gt 0) {
        $samples = Get-CounterSampleBatch -Paths $script:detailedPaths
        $matched = Get-CollectionCount -Items @(
            Ensure-Array -Items $samples | ForEach-Object { Get-CounterPathKey -Path $_.Path } |
                Where-Object { $script:detailedColumnMap.ContainsKey($_) }
        )
        Write-Log "Detailed sample: $(Get-CollectionCount -Items $samples) counters read, $matched matched"
        $row = ConvertTo-DetailedRow `
            -Samples $samples `
            -ColumnMap $script:detailedColumnMap `
            -SampleTime $sampleTime
        $script:detailedRows.Add($row) | Out-Null
        if ($null -ne $script:detailedCsvStream) {
            Append-MetricsCsvRow -Stream $script:detailedCsvStream -Row $row
        }
    }

    $script:SampleCount++
}

$script:summaryCsvStream = $null
$script:detailedCsvStream = $null

try {
    if ($Mode -eq 'Snapshot') {
        Invoke-SampleCollection
        Write-Log 'Snapshot collection complete.'
    }
    else {
        Write-Log "Continuous collection: interval=${IntervalSeconds}s, duration=$(
            if ($DurationSeconds -gt 0) { "${DurationSeconds}s" } else { 'unlimited' }
        ), maxSamples=$(
            if ($MaxSamples -gt 0) { $MaxSamples } else { 'unlimited' }
        )"

        if ($collectSummary -and (Get-CollectionCount -Items $summaryPaths) -gt 0) {
            $summaryFile = Join-Path $OutputDirectory "${OutputPrefix}_summary_${runTimestamp}.csv"
            $summaryColumnOrder = Get-SummaryColumnOrder -ColumnMap $summaryColumnMap `
                -AggregateNetworkColumns $summaryDefs.AggregateNetwork
            $script:summaryCsvStream = Initialize-MetricsCsvStream -FilePath $summaryFile `
                -ColumnOrder $summaryColumnOrder
            Write-Log "Streaming summary metrics to: $summaryFile"
        }

        if ($collectDetailed -and (Get-CollectionCount -Items $detailedPaths) -gt 0) {
            $detailedFile = Join-Path $OutputDirectory "${OutputPrefix}_detailed_${runTimestamp}.csv"
            $detailedColumnOrder = Get-DetailedColumnOrder -ColumnMap $detailedColumnMap
            $script:detailedCsvStream = Initialize-MetricsCsvStream -FilePath $detailedFile `
                -ColumnOrder $detailedColumnOrder
            Write-Log "Streaming detailed metrics to: $detailedFile"
        }

        while (-not $script:StopRequested) {
            Invoke-SampleCollection

            if ($MaxSamples -gt 0 -and $script:SampleCount -ge $MaxSamples) {
                Write-Log "Reached MaxSamples ($MaxSamples). Stopping."
                break
            }
            if ($DurationSeconds -gt 0 -and $stopwatch.Elapsed.TotalSeconds -ge $DurationSeconds) {
                Write-Log "Reached DurationSeconds ($DurationSeconds). Stopping."
                break
            }
            if ($script:StopRequested) {
                break
            }

            Wait-Interruptible -Seconds $IntervalSeconds
        }
        Write-Log "Continuous collection complete. Samples collected: $script:SampleCount"
    }
}
finally {
    try {
        [Console]::CancelKeyPress -= $cancelHandler
    }
    catch {
        # ignore
    }

    if ($Mode -eq 'Snapshot') {
        Write-BufferedMetricsCsv -WriteSummary $collectSummary -WriteDetailed $collectDetailed `
            -OutputDirectory $OutputDirectory -OutputPrefix $OutputPrefix -RunTimestamp $runTimestamp
    }
    else {
        if ($null -ne $script:summaryCsvStream -and $script:summaryCsvStream.RowCount -gt 0) {
            Write-Log "Summary CSV contains $($script:summaryCsvStream.RowCount) row(s): $($script:summaryCsvStream.FilePath)"
        }
        if ($null -ne $script:detailedCsvStream -and $script:detailedCsvStream.RowCount -gt 0) {
            Write-Log "Detailed CSV contains $($script:detailedCsvStream.RowCount) row(s): $($script:detailedCsvStream.FilePath)"
        }
    }

    Write-Log 'Done.'
}
#endregion
