[CmdletBinding()]
param(
    [Parameter()]
    [string[]]$Targets,

    [Parameter()]
    [ValidateRange(1, 65535)]
    [int[]]$Ports,

    [Parameter()]
    [string]$ExportPath,

    [Parameter()]
    [ValidateSet('Json', 'Csv')]
    [string]$ExportFormat = 'Json',

    [Parameter()]
    [switch]$PromptForTargets,

    [Parameter()]
    [switch]$SkipRemoteChecks,

    [Parameter()]
    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-IsAdministrator {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-InteractiveTargets {
    $response = Read-Host 'Enter hostnames or IP addresses separated by commas (leave blank for local-only diagnostics)'
    if ([string]::IsNullOrWhiteSpace($response)) {
        return @()
    }

    return @(
        $response.Split(',') |
        ForEach-Object { $_.Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
}

function Test-IsApipaAddress {
    param(
        [Parameter(Mandatory)]
        [string]$Address
    )

    return $Address -match '^169\.254\.'
}

function Resolve-TargetAddresses {
    param(
        [Parameter(Mandatory)]
        [string]$Target
    )

    try {
        $addresses = [System.Net.Dns]::GetHostAddresses($Target) |
            ForEach-Object { $_.IPAddressToString } |
            Select-Object -Unique

        return [pscustomobject]@{
            Status = 'Resolved'
            Addresses = @($addresses)
            Error = $null
        }
    }
    catch {
        return [pscustomobject]@{
            Status = 'Failed'
            Addresses = @()
            Error = $_.Exception.Message
        }
    }
}

function Get-PingResponseTime {
    param(
        [Parameter(Mandatory)]
        [object]$Reply
    )

    if ($Reply.PSObject.Properties.Name -contains 'ResponseTime') {
        return [double]$Reply.ResponseTime
    }

    if ($Reply.PSObject.Properties.Name -contains 'Latency') {
        return [double]$Reply.Latency
    }

    return $null
}

function Test-TcpPort {
    param(
        [Parameter(Mandatory)]
        [string]$Target,

        [Parameter(Mandatory)]
        [ValidateRange(1, 65535)]
        [int]$Port,

        [Parameter()]
        [ValidateRange(500, 30000)]
        [int]$TimeoutMs = 2000
    )

    $client = [System.Net.Sockets.TcpClient]::new()
    try {
        $asyncHandle = $client.BeginConnect($Target, $Port, $null, $null)
        $connected = $asyncHandle.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
        if (-not $connected) {
            return [pscustomobject]@{
                Port = $Port
                Reachable = $false
                Error = 'Timed out'
            }
        }

        $client.EndConnect($asyncHandle)
        return [pscustomobject]@{
            Port = $Port
            Reachable = $true
            Error = $null
        }
    }
    catch {
        return [pscustomobject]@{
            Port = $Port
            Reachable = $false
            Error = $_.Exception.Message
        }
    }
    finally {
        $client.Dispose()
    }
}

function Get-EthernetAdapters {
    $allPhysicalAdapters = @()
    try {
        $allPhysicalAdapters = @(Get-NetAdapter -Physical -ErrorAction Stop)
    }
    catch {
        return @()
    }

    $ethernetAdapters = @(
        $allPhysicalAdapters |
        Where-Object {
            $_.Name -match 'Ethernet' -or
            $_.InterfaceDescription -match 'Ethernet' -or
            $_.NdisPhysicalMedium -eq '802.3'
        }
    )

    if ($ethernetAdapters.Count -eq 0) {
        $ethernetAdapters = $allPhysicalAdapters
    }

    return @(
        $ethernetAdapters | Sort-Object -Property ifIndex | ForEach-Object {
            $adapter = $_
            $ipConfig = Get-NetIPConfiguration -InterfaceIndex $adapter.ifIndex -ErrorAction SilentlyContinue
            $ipv4 = @($ipConfig.IPv4Address | ForEach-Object { $_.IPv4Address })
            $ipv6 = @($ipConfig.IPv6Address | ForEach-Object { $_.IPv6Address })
            $gateways = @(
                $ipConfig.IPv4DefaultGateway | ForEach-Object {
                    $gatewayEntry = $_
                    $nextHop = $null

                    try {
                        $nextHop = $gatewayEntry | Select-Object -ExpandProperty NextHop -ErrorAction Stop
                    }
                    catch {
                        $nextHop = $null
                    }

                    if (-not [string]::IsNullOrWhiteSpace([string]$nextHop)) {
                        $nextHop
                    }
                    elseif (-not [string]::IsNullOrWhiteSpace([string]$gatewayEntry)) {
                        [string]$gatewayEntry
                    }
                }
            )
            $dnsServers = @(
                try {
                    if ($null -ne $ipConfig.DnsServer.ServerAddresses) {
                        $ipConfig.DnsServer.ServerAddresses
                    }
                }
                catch {
                }
            )

            [pscustomobject]@{
                Name = $adapter.Name
                Status = $adapter.Status
                LinkSpeed = $adapter.LinkSpeed
                MacAddress = $adapter.MacAddress
                InterfaceIndex = $adapter.ifIndex
                IPv4Addresses = @($ipv4)
                IPv6Addresses = @($ipv6)
                Gateways = @($gateways)
                DnsServers = @($dnsServers)
            }
        }
    )
}

function Get-DiskSummary {
    $drives = @(Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType = 3')
    return @(
        $drives | Sort-Object -Property DeviceID | ForEach-Object {
            $totalBytes = [double]$_.Size
            $freeBytes = [double]$_.FreeSpace
            $usedBytes = $totalBytes - $freeBytes
            $usedPercent = if ($totalBytes -gt 0) {
                [math]::Round(($usedBytes / $totalBytes) * 100, 2)
            }
            else {
                0
            }

            [pscustomobject]@{
                Drive = $_.DeviceID
                VolumeName = $_.VolumeName
                FileSystem = $_.FileSystem
                TotalGB = [math]::Round($totalBytes / 1GB, 2)
                FreeGB = [math]::Round($freeBytes / 1GB, 2)
                UsedPercent = $usedPercent
            }
        }
    )
}

function Get-MemorySummary {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    $totalMB = [double]$os.TotalVisibleMemorySize / 1024
    $freeMB = [double]$os.FreePhysicalMemory / 1024
    $usedMB = $totalMB - $freeMB
    $usedPercent = if ($totalMB -gt 0) {
        [math]::Round(($usedMB / $totalMB) * 100, 2)
    }
    else {
        0
    }

    return [pscustomobject]@{
        TotalGB = [math]::Round($totalMB / 1024, 2)
        FreeGB = [math]::Round($freeMB / 1024, 2)
        UsedGB = [math]::Round($usedMB / 1024, 2)
        UsedPercent = $usedPercent
    }
}

function Get-CpuSummary {
    $processors = @(Get-CimInstance -ClassName Win32_Processor)
    $averageLoad = if ($processors.Count -gt 0) {
        [math]::Round((($processors | Measure-Object -Property LoadPercentage -Average).Average), 2)
    }
    else {
        0
    }

    return [pscustomobject]@{
        AverageLoadPercent = $averageLoad
        LogicalProcessors = [Environment]::ProcessorCount
        Name = (($processors | Select-Object -ExpandProperty Name -Unique) -join '; ')
    }
}

function Get-GpuSummary {
    $controllers = @(Get-CimInstance -ClassName Win32_VideoController -ErrorAction SilentlyContinue)
    $utilization = $null
    $counterAvailable = $false

    try {
        $samples = (Get-Counter '\GPU Engine(*)\Utilization Percentage' -SampleInterval 1 -MaxSamples 1 -ErrorAction Stop).CounterSamples
        $positiveSamples = @($samples | Where-Object { $_.CookedValue -gt 0 })
        if ($positiveSamples.Count -gt 0) {
            $utilization = [math]::Round((($positiveSamples | Measure-Object -Property CookedValue -Maximum).Maximum), 2)
        }
        else {
            $utilization = 0
        }
        $counterAvailable = $true
    }
    catch {
        $utilization = $null
    }

    return [pscustomobject]@{
        Controllers = @(
            $controllers | ForEach-Object {
                [pscustomobject]@{
                    Name = $_.Name
                    DriverVersion = $_.DriverVersion
                    AdapterRamGB = if ($_.AdapterRAM) { [math]::Round(([double]$_.AdapterRAM / 1GB), 2) } else { $null }
                }
            }
        )
        UtilizationPercent = $utilization
        CounterAvailable = $counterAvailable
    }
}

function Get-PowerSummary {
    $activePowerPlan = $null
    try {
        $powercfgOutput = & powercfg /getactivescheme 2>$null
        if ($LASTEXITCODE -eq 0 -and $powercfgOutput) {
            $line = ($powercfgOutput | Select-Object -First 1)
            if ($line -match '\((?<Name>.+)\)$') {
                $activePowerPlan = $Matches.Name
            }
            else {
                $activePowerPlan = $line
            }
        }
    }
    catch {
        $activePowerPlan = $null
    }

    $batteries = @(Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue)
    $averageCharge = if ($batteries.Count -gt 0) {
        [math]::Round((($batteries | Measure-Object -Property EstimatedChargeRemaining -Average).Average), 2)
    }
    else {
        $null
    }

    return [pscustomobject]@{
        ActivePowerPlan = $activePowerPlan
        HasBattery = $batteries.Count -gt 0
        BatteryChargeRemainingPercent = $averageCharge
        BatteryCount = $batteries.Count
    }
}

function Get-LocalDiagnostics {
    $ethernetAdapters = @(Get-EthernetAdapters)
    $ipEnabledAdapters = @($ethernetAdapters | Where-Object { $_.IPv4Addresses.Count -gt 0 -or $_.IPv6Addresses.Count -gt 0 })

    return [pscustomobject]@{
        ComputerName = $env:COMPUTERNAME
        UserName = $env:USERNAME
        IsAdministrator = Test-IsAdministrator
        EthernetAdapters = $ethernetAdapters
        ConnectedEthernetAdapters = @($ethernetAdapters | Where-Object { $_.Status -eq 'Up' })
        AddressedAdapters = $ipEnabledAdapters
        Disks = @(Get-DiskSummary)
        Memory = Get-MemorySummary
        Cpu = Get-CpuSummary
        Gpu = Get-GpuSummary
        Power = Get-PowerSummary
    }
}

function Get-RemoteDiagnostics {
    param(
        [Parameter(Mandatory)]
        [string[]]$TargetList,

        [Parameter()]
        [int[]]$PortList
    )

    return @(
        $TargetList |
        Select-Object -Unique |
        ForEach-Object {
            $target = $_
            $resolution = Resolve-TargetAddresses -Target $target
            $pingReplies = @()
            $status = 'Unknown'
            $notes = New-Object System.Collections.Generic.List[string]

            try {
                $pingReplies = @(Test-Connection -ComputerName $target -Count 2 -ErrorAction Stop)
            }
            catch {
                $notes.Add($_.Exception.Message)
            }

            $latencies = @(
                $pingReplies |
                ForEach-Object { Get-PingResponseTime -Reply $_ } |
                Where-Object { $null -ne $_ }
            )

            $averageLatency = if ($latencies.Count -gt 0) {
                [math]::Round((($latencies | Measure-Object -Average).Average), 2)
            }
            else {
                $null
            }

            if ($pingReplies.Count -gt 0) {
                $status = 'Reachable'
            }
            elseif ($resolution.Status -eq 'Resolved') {
                $status = 'NoPingReply'
            }
            else {
                $status = 'DnsFailed'
            }

            $portResults = @()
            if ($PortList.Count -gt 0) {
                if ($resolution.Status -eq 'Resolved') {
                    $portResults = @(
                        $PortList | ForEach-Object {
                            Test-TcpPort -Target $target -Port $_
                        }
                    )
                }
                else {
                    $portResults = @(
                        $PortList | ForEach-Object {
                            [pscustomobject]@{
                                Port = $_
                                Reachable = $false
                                Error = 'Skipped because DNS resolution failed'
                            }
                        }
                    )
                }
            }

            [pscustomobject]@{
                Target = $target
                Status = $status
                DnsStatus = $resolution.Status
                ResolvedAddresses = @($resolution.Addresses)
                AverageLatencyMs = $averageLatency
                PingReplies = $pingReplies.Count
                PortChecks = @($portResults)
                Notes = @($notes)
                ResolutionError = $resolution.Error
            }
        }
    )
}

function Convert-ReportToCsvRows {
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Report
    )

    $rows = New-Object System.Collections.Generic.List[object]

    foreach ($adapter in $Report.Local.EthernetAdapters) {
        $rows.Add([pscustomobject]@{
                Category = 'EthernetAdapter'
                Name = $adapter.Name
                Status = $adapter.Status
                Detail = ($adapter.IPv4Addresses -join ', ')
            })
    }

    foreach ($disk in $Report.Local.Disks) {
        $rows.Add([pscustomobject]@{
                Category = 'Disk'
                Name = $disk.Drive
                Status = if ($disk.UsedPercent -ge 90) { 'Fail' } elseif ($disk.UsedPercent -ge 80) { 'Warn' } else { 'Pass' }
                Detail = "Used $($disk.UsedPercent)% | Free $($disk.FreeGB) GB"
            })
    }

    $rows.Add([pscustomobject]@{
            Category = 'Memory'
            Name = 'PhysicalMemory'
            Status = if ($Report.Local.Memory.UsedPercent -ge 90) { 'Fail' } elseif ($Report.Local.Memory.UsedPercent -ge 80) { 'Warn' } else { 'Pass' }
            Detail = "Used $($Report.Local.Memory.UsedPercent)%"
        })

    $rows.Add([pscustomobject]@{
            Category = 'Cpu'
            Name = 'AverageLoad'
            Status = if ($Report.Local.Cpu.AverageLoadPercent -ge 90) { 'Fail' } elseif ($Report.Local.Cpu.AverageLoadPercent -ge 80) { 'Warn' } else { 'Pass' }
            Detail = "$($Report.Local.Cpu.AverageLoadPercent)%"
        })

    foreach ($remote in $Report.Remote) {
        $rows.Add([pscustomobject]@{
                Category = 'RemoteTarget'
                Name = $remote.Target
                Status = $remote.Status
                Detail = if ($null -ne $remote.AverageLatencyMs) { "$($remote.AverageLatencyMs) ms" } elseif ($remote.ResolutionError) { $remote.ResolutionError } else { ($remote.Notes -join ' | ') }
            })

        foreach ($portCheck in $remote.PortChecks) {
            $rows.Add([pscustomobject]@{
                    Category = 'PortCheck'
                    Name = "$($remote.Target):$($portCheck.Port)"
                    Status = if ($portCheck.Reachable) { 'Reachable' } else { 'ClosedOrFiltered' }
                    Detail = if ($portCheck.Error) { $portCheck.Error } else { 'Open' }
                })
        }
    }

    return $rows
}

function Export-Report {
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Report,

        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [ValidateSet('Json', 'Csv')]
        [string]$Format
    )

    $directory = Split-Path -Path $Path -Parent
    if ($directory -and -not (Test-Path -Path $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    switch ($Format) {
        'Json' {
            $Report | ConvertTo-Json -Depth 8 | Set-Content -Path $Path -Encoding ASCII
        }
        'Csv' {
            Convert-ReportToCsvRows -Report $Report | Export-Csv -Path $Path -NoTypeInformation -Encoding ASCII
        }
    }
}

function Write-ReportSummary {
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Report
    )

    Write-Host ''
    Write-Host 'System Diagnostics Summary' -ForegroundColor Cyan
    Write-Host "Timestamp: $($Report.Timestamp)"
    Write-Host "Computer: $($Report.Local.ComputerName)"
    Write-Host "Running as administrator: $($Report.Local.IsAdministrator)"

    Write-Host ''
    Write-Host 'Ethernet adapters' -ForegroundColor Cyan
    if ($Report.Local.EthernetAdapters.Count -eq 0) {
        Write-Host '  No physical Ethernet adapters were found.' -ForegroundColor Yellow
    }
    else {
        foreach ($adapter in $Report.Local.EthernetAdapters) {
            $addressList = @($adapter.IPv4Addresses + $adapter.IPv6Addresses) -join ', '
            if (-not $addressList) {
                $addressList = 'No IP addresses assigned'
            }

            $color = if ($adapter.Status -eq 'Up') { 'Green' } else { 'Yellow' }
            Write-Host "  [$($adapter.Status)] $($adapter.Name) | $($adapter.LinkSpeed) | $addressList" -ForegroundColor $color
        }
    }

    Write-Host ''
    Write-Host 'Storage' -ForegroundColor Cyan
    foreach ($disk in $Report.Local.Disks) {
        $diskColor = if ($disk.UsedPercent -ge 90) { 'Red' } elseif ($disk.UsedPercent -ge 80) { 'Yellow' } else { 'Green' }
        Write-Host "  [$($disk.Drive)] Used $($disk.UsedPercent)% | Free $($disk.FreeGB) GB of $($disk.TotalGB) GB" -ForegroundColor $diskColor
    }

    Write-Host ''
    Write-Host 'Compute' -ForegroundColor Cyan
    $memoryColor = if ($Report.Local.Memory.UsedPercent -ge 90) { 'Red' } elseif ($Report.Local.Memory.UsedPercent -ge 80) { 'Yellow' } else { 'Green' }
    $cpuColor = if ($Report.Local.Cpu.AverageLoadPercent -ge 90) { 'Red' } elseif ($Report.Local.Cpu.AverageLoadPercent -ge 80) { 'Yellow' } else { 'Green' }
    Write-Host "  Memory used: $($Report.Local.Memory.UsedPercent)% ($($Report.Local.Memory.UsedGB) GB / $($Report.Local.Memory.TotalGB) GB)" -ForegroundColor $memoryColor
    Write-Host "  CPU load: $($Report.Local.Cpu.AverageLoadPercent)%" -ForegroundColor $cpuColor
    if ($null -ne $Report.Local.Gpu.UtilizationPercent) {
        Write-Host "  GPU utilization: $($Report.Local.Gpu.UtilizationPercent)%" -ForegroundColor Green
    }
    else {
        Write-Host '  GPU utilization: unavailable on this system' -ForegroundColor Yellow
    }

    Write-Host ''
    Write-Host 'Power' -ForegroundColor Cyan
    if ($Report.Local.Power.ActivePowerPlan) {
        Write-Host "  Active power plan: $($Report.Local.Power.ActivePowerPlan)"
    }
    else {
        Write-Host '  Active power plan: unavailable' -ForegroundColor Yellow
    }

    if ($Report.Local.Power.HasBattery) {
        Write-Host "  Battery charge: $($Report.Local.Power.BatteryChargeRemainingPercent)%"
    }
    else {
        Write-Host '  Battery information: no battery detected' -ForegroundColor Yellow
    }

    if ($Report.Remote.Count -gt 0) {
        Write-Host ''
        Write-Host 'Remote targets' -ForegroundColor Cyan
        foreach ($remote in $Report.Remote) {
            $remoteColor = switch ($remote.Status) {
                'Reachable' { 'Green' }
                'NoPingReply' { 'Yellow' }
                default { 'Red' }
            }

            $addressText = if ($remote.ResolvedAddresses.Count -gt 0) {
                $remote.ResolvedAddresses -join ', '
            }
            else {
                'No resolved addresses'
            }

            $detail = if ($null -ne $remote.AverageLatencyMs) {
                "$($remote.AverageLatencyMs) ms"
            }
            elseif ($remote.ResolutionError) {
                $remote.ResolutionError
            }
            else {
                ($remote.Notes -join ' | ')
            }

            Write-Host "  [$($remote.Status)] $($remote.Target) | $addressText | $detail" -ForegroundColor $remoteColor

            foreach ($portCheck in $remote.PortChecks) {
                $portColor = if ($portCheck.Reachable) { 'Green' } else { 'Yellow' }
                $portDetail = if ($portCheck.Reachable) { 'Open' } else { $portCheck.Error }
                Write-Host "    Port $($portCheck.Port): $portDetail" -ForegroundColor $portColor
            }
        }
    }

    if ($Report.Warnings.Count -gt 0) {
        Write-Host ''
        Write-Host 'Warnings' -ForegroundColor Cyan
        foreach ($warning in $Report.Warnings) {
            Write-Host "  - $warning" -ForegroundColor Yellow
        }
    }
}

$targetList = @($Targets | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
if ($PromptForTargets -and $targetList.Count -eq 0 -and -not $SkipRemoteChecks) {
    $targetList = @(Get-InteractiveTargets)
}

$warnings = New-Object System.Collections.Generic.List[string]
$localDiagnostics = Get-LocalDiagnostics

if (-not $localDiagnostics.IsAdministrator) {
    $warnings.Add('Some counters and adapter details may be limited when the script is not run as administrator.')
}

if ($localDiagnostics.EthernetAdapters.Count -eq 0) {
    $warnings.Add('No physical Ethernet adapters were detected.')
}
elseif ($localDiagnostics.ConnectedEthernetAdapters.Count -eq 0) {
    $warnings.Add('Ethernet adapters were found, but none are currently up.')
}

$apipaAdapters = @(
    $localDiagnostics.ConnectedEthernetAdapters |
    Where-Object {
        @($_.IPv4Addresses | Where-Object { Test-IsApipaAddress -Address $_ }).Count -gt 0
    }
)

foreach ($adapter in $apipaAdapters) {
    $apipaAddresses = @($adapter.IPv4Addresses | Where-Object { Test-IsApipaAddress -Address $_ }) -join ', '
    $warnings.Add("Adapter '$($adapter.Name)' is up but only has APIPA/link-local IPv4 addressing ($apipaAddresses). This usually means DHCP or routed network connectivity is unavailable.")
}

if ($null -eq $localDiagnostics.Gpu.UtilizationPercent) {
    $warnings.Add('GPU utilization counters are not available on this system.')
}

if (-not $localDiagnostics.Power.HasBattery) {
    $warnings.Add('Battery information is unavailable or the device does not have a battery.')
}

$remoteDiagnostics = @()
if (-not $SkipRemoteChecks -and $targetList.Count -gt 0) {
    $remoteDiagnostics = @(Get-RemoteDiagnostics -TargetList $targetList -PortList $Ports)
}

$report = [pscustomobject]@{
    Timestamp = (Get-Date).ToString('s')
    Local = $localDiagnostics
    Remote = @($remoteDiagnostics)
    Warnings = @($warnings)
}

Write-ReportSummary -Report $report

if ($ExportPath) {
    Export-Report -Report $report -Path $ExportPath -Format $ExportFormat
    Write-Host ''
    Write-Host "Exported $ExportFormat report to $ExportPath" -ForegroundColor Cyan
}

if ($PassThru) {
    $report
}