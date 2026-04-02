# XProtect-Prepare

PowerShell script to prepare and optimize Windows machines for Milestone XProtect installations.

## Version

**v2.0** (2026) - Major update with NTP fix, SSH, Copilot/Recall disable, 25H2 bloatware, and more.

## What it does

- Configures antivirus exceptions for Milestone processes, folders, and file extensions
- Checks storage drive block size (64KB) and disables indexing
- Removes Windows bloatware including 25H2 apps (Clipchamp, TikTok, Copilot, etc.)
- Disables Windows Copilot and Recall via Group Policy
- Disables telemetry scheduled tasks
- Optimizes power management for 24/7 recording (Never Sleep Mode)
- Disables unnecessary Windows services
- Enables and optimizes Remote Desktop
- Configures NTP Time Server for camera time synchronization (with correct peer format)
- Installs and configures OpenSSH Server
- Enables LargeSystemCache for file server performance
- Removes OneDrive, Microsoft Teams, Outlook App, Office Hub
- Installs Milestone PowerShell Tools
- Optional SNMP installation
- Windows Update toggle (standalone, not part of Complete Setup)
- Comprehensive logging with registry backup

## Menu Structure (v2.0)

```
=== XProtect-Prepare v2.0 ===
Detected: [Windows Server 2022 / Windows 11 etc]
[WARNING if NTP bug detected from v1.9x]

1. Complete Setup (without SSH)
2. Complete Setup (with SSH activation)
3. Individual optimizations (submenu)
4. Configure NTP Time Server
5. Install & Configure SSH Server
6. Toggle Windows Update On/Off
7. NTP Recovery (fix v1.9x bug)
8. Exit
```

## v2.0 Changes

- **CRITICAL FIX:** NTP peer format corrected - added `,0x9` flags (v1.9x bug caused silent sync failure)
- **NEW:** NTP auto-detects domain vs standalone for correct sync mode (AllSync/NTP)
- **NEW:** NTP Recovery function and standalone script for fixing deployed machines
- **NEW:** SSH Server install and configure option
- **NEW:** Copilot and Recall disabled via Group Policy registry keys
- **NEW:** Telemetry scheduled tasks disabled (Compatibility Appraiser, CEIP, DiskDiagnostic)
- **NEW:** LargeSystemCache registry optimization
- **NEW:** 25H2 bloatware additions (Clipchamp, Dev Home, TikTok, Spotify, Facebook, McAfee, Copilot, LinkedIn, etc.)
- **IMPROVED:** Windows Update toggle moved to standalone menu option (removed from Complete Setup flow)
- **FIXED:** All `Get-WmiObject` replaced with `Get-CimInstance`
- **FIXED:** ErrorAction placement on Where-Object pipelines
- **FIXED:** AV exclusions added regardless of folder existence (pre-install scenario)
- **FIXED:** Auto-detect Windows version at startup

## Files

| File | Description |
|------|-------------|
| `Xprotect-prepare-20.ps1` | Main script v2.0 |
| `Xprotect-prepare-19.ps1` | Previous version (v1.92) |
| `NTP-Recovery.ps1` | Standalone NTP fix for machines deployed with v1.9x |

## Usage

1. Clone the repository:
   ```powershell
   git clone https://github.com/Frimurare/XProtect-Prepare.git
   cd XProtect-Prepare
   ```

2. Set execution policy (required the first time):
   ```powershell
   Set-ExecutionPolicy RemoteSigned -Scope Process -Force
   ```

3. Run as Administrator:
   ```powershell
   .\Xprotect-prepare-20.ps1
   ```

### NTP Recovery (for machines with v1.9x)

Run on each affected machine as Administrator:
```powershell
.\NTP-Recovery.ps1
```

## NTP Bug Details

The v1.9x scripts set NTP peers as:
```
0.se.pool.ntp.org,1.se.pool.ntp.org,2.se.pool.ntp.org,3.se.pool.ntp.org
```

The correct format requires `,0x9` flags after each peer (space-separated, not comma-separated):
```
0.se.pool.ntp.org,0x9 1.se.pool.ntp.org,0x9 2.se.pool.ntp.org,0x9 3.se.pool.ntp.org,0x9
```

Without the flags, Windows Time Service cannot parse the peer list and time sync fails silently. The v2.0 script detects this bug at startup and shows a warning banner.

## Author

Ulf Holmstrom (Frimurare)
Technical Engineer - OpenEye / ex Solution Engineer at Milestone Systems

For questions contact: ulf@manvarg.se
