# Changelog — XProtect-Prepare

All notable changes to this project are documented here. Newest first.

## v2.1 "Swedish Summer Edition" (2026)

- **NEW:** Full support for **Milestone XProtect 2026 R1**.
- **NEW:** **VictoriaLogs** antivirus process exclusions (`victoria-logs-windows-amd64-prod.exe`, `VideoOS.ServiceWrapper.exe`). 2026 R1 replaced the SQL-based Log Server with VictoriaLogs.
- **NEW:** **XProtect Smart Client** antivirus exclusions (`Client.exe` + per-user `VideoOS` cache), added automatically **only when the client is installed** — fixes the very slow first launch after a (re)install, caused by the antivirus scanning thousands of fresh client files.
- **NEW:** VictoriaLogs / legacy Log Server **auto-detection** with an informational message.
- **CHANGED:** **MilestonePSTools** is now installed via the official one-line web installer (`milestonepstools.com/install.ps1`). The previous `Install-Module` path could leave a module too old to connect to a 2026 R1 server.
- **NEW:** `-Silent` **unattended mode** for golden images and deployment automation (`-WithSSH`, `-Drives`).
- **COMPATIBILITY:** the legacy Log Server process exclusion is kept, so the antivirus configuration is correct on **both** 2026 R1 and earlier Milestone versions.

## v2.0 (2026)

- Menu restructured with **Complete Setup** (with / without SSH).
- **CRITICAL FIX:** NTP peer format corrected — added `,0x9` flags. The v1.9x format caused **silent** time-sync failure.
- NTP auto-detects domain vs standalone for the correct sync mode (AllSync / NTP).
- **NTP Recovery** function and standalone `NTP-Recovery.ps1` to fix already-deployed machines.
- OpenSSH Server install & configure option.
- Copilot and Recall disabled via Group Policy.
- Telemetry scheduled tasks disabled (Compatibility Appraiser, CEIP, DiskDiagnostic).
- `LargeSystemCache` registry optimization.
- Windows 25H2 bloatware additions (Clipchamp, Dev Home, TikTok, Spotify, Facebook, McAfee, Copilot, LinkedIn, etc.).
- Windows Update toggle moved to a standalone menu option (removed from Complete Setup flow).
- All `Get-WmiObject` replaced with `Get-CimInstance`.
- Antivirus exclusions added regardless of folder existence (pre-install scenario).
- Auto-detect Windows version at startup.

## v1.92

- Added Milestone installation folder exclusions on C: (`Program Files\Milestone\`, `Program Files (x86)\Milestone\`, `ProgramData\Milestone\`, `ProgramData\VideoDeviceDrivers\`, `ProgramData\VideoOS\`).
- Added media file extensions (`.pic`, `.pqz`, `.sts`, `.ts`) for XProtect Enterprise.
- Aligned with Milestone's official antivirus best practices.

## v1.9

- NTP Time Server configuration for camera time synchronization.
- Automatic firewall configuration for NTP (UDP port 123).
- Reliable, forensically valid timestamps on recordings.
