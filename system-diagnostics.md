# System Diagnostics Script

## Overview

`system-diagnostics.ps1` is a standalone Windows diagnostics script for quick local health checks and optional remote connectivity debugging. It reports Ethernet adapter state, IP configuration, disk usage, memory pressure, CPU load, GPU availability, power status, and remote reachability for hostnames or IP addresses you provide.

This implementation is intentionally separate from `create-venv.ps1` and does not depend on repo-level README changes.

## What It Checks

- Physical Ethernet adapters and assigned IP addresses
- Default gateways and DNS server assignments exposed through adapter configuration
- Disk capacity, free space, and used percentage for local fixed drives
- Physical memory usage
- CPU load from Windows processor telemetry
- GPU controller inventory and utilization when the Windows GPU counter is available
- Power plan and battery charge data when Windows exposes it
- DNS resolution, ping reachability, and optional TCP port checks for remote targets

## Requirements

- Windows PowerShell 5.1 or PowerShell 7 on Windows
- Windows networking cmdlets such as `Get-NetAdapter` and `Get-NetIPConfiguration`
- Optional: run from an elevated PowerShell session for the most complete adapter and counter visibility

## Parameters

### `-Targets`

One or more hostnames or IP addresses to probe.

```powershell
.\system-diagnostics.ps1 -Targets printer01, 192.168.1.10
```

### `-Ports`

Optional TCP ports to test on each remote target.

```powershell
.\system-diagnostics.ps1 -Targets server01 -Ports 22, 80, 443
```

### `-ExportPath`

Optional file path for structured output.

```powershell
.\system-diagnostics.ps1 -ExportPath .\reports\system-report.json
```

### `-ExportFormat`

Optional export type. Supported values are `Json` and `Csv`. Default is `Json`.

```powershell
.\system-diagnostics.ps1 -ExportPath .\reports\system-report.csv -ExportFormat Csv
```

### `-PromptForTargets`

Prompts for a comma-separated target list if `-Targets` was omitted.

```powershell
.\system-diagnostics.ps1 -PromptForTargets
```

### `-SkipRemoteChecks`

Runs local diagnostics only.

```powershell
.\system-diagnostics.ps1 -SkipRemoteChecks
```

### `-PassThru`

Returns the report object to the pipeline in addition to the console summary.

```powershell
.\system-diagnostics.ps1 -SkipRemoteChecks -PassThru
```

## Recommended Usage

Run a local-only diagnostics pass:

```powershell
.\system-diagnostics.ps1 -SkipRemoteChecks
```

Run local diagnostics and probe a few devices:

```powershell
.\system-diagnostics.ps1 -Targets router.local, nas01, 192.168.1.50 -Ports 22, 80, 443
```

Export a machine-readable report for later review:

```powershell
.\system-diagnostics.ps1 -Targets workstation02 -ExportPath .\reports\diag.json -ExportFormat Json
```

## Output Notes

- The script always prints a console summary first.
- Use `-PassThru` if you want the full report object returned to the pipeline.
- JSON export preserves the full report object, including nested local and remote details.
- CSV export flattens key findings into category, name, status, and detail rows.
- Remote targets are reported as `Reachable`, `NoPingReply`, or `DnsFailed`.
- An adapter shown as `Up` with a `169.254.x.x` address is using APIPA/link-local addressing, which usually means it does not have working DHCP or routed network access.

## Limitations

- Per-device power draw for arbitrary remote devices is not included. That usually requires vendor APIs, SNMP, or device-specific tooling.
- GPU utilization depends on the Windows GPU Engine performance counter. Some systems expose controller inventory but not live utilization.
- Battery information is only available on systems where Windows reports a battery.
- ICMP ping can fail even when a device is healthy if the device or firewall blocks ping.

## Troubleshooting

### No Ethernet adapters shown

- Confirm the machine exposes adapters through `Get-NetAdapter -Physical`.
- Run the script in an elevated PowerShell session.
- On systems without standard Windows networking cmdlets, adapter visibility may be limited.

### An adapter is up but shows `169.254.x.x`

- That is an APIPA/link-local IPv4 address rather than a normal routed address.
- It usually means the adapter is physically up but could not obtain DHCP configuration.
- Check cabling, switch or dock state, VLAN assignment, or DHCP availability on that segment.

### GPU utilization says unavailable

- The script will still report GPU controllers when available.
- Some drivers do not expose the `GPU Engine` performance counter.

### A remote device resolves but is not reachable

- The device may be online but blocking ICMP.
- Add `-Ports` for service-level checks such as `22`, `80`, `443`, or another known application port.

### Export failed

- Make sure the target directory is writable.
- If you export CSV, remember that nested fields are flattened into summary rows rather than raw nested objects.

## Future Extensions

- Read targets from a CSV or JSON inventory file
- Add named port profiles for common service groups
- Add historical report comparison for repeated health checks