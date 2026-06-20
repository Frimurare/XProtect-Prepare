# XProtect-Prepare

A PowerShell tool that prepares and optimizes Windows machines for **Milestone XProtect** installations — antivirus exclusions, storage tuning, bloatware removal, performance hardening, NTP time-server setup, and more, all from one menu (or fully unattended).

> **v2.1 "Swedish Summer Edition"** — now with full **Milestone XProtect 2026 R1** support (VictoriaLogs), XProtect Smart Client antivirus exclusions, the official MilestonePSTools web installer, and a `-Silent` unattended mode.

Built by **Ulf Holmström, Happy Problem Solver at Manvarg AB** — to make life easier for Milestone resellers and integrators.
It is provided as a free public resource and is **NOT** supported by Milestone Systems A/S.

---

## Compatibility

| Milestone version | Supported |
|---|---|
| **XProtect 2026 R1** (VictoriaLogs log engine) | ✅ Yes |
| XProtect 2025 R-x and earlier (SQL Log Server) | ✅ Yes |
| Windows Server 2019 / 2022 / 2025 | ✅ Yes |
| Windows 10 / 11 (incl. 24H2 / 25H2) | ✅ Yes |

The antivirus exclusions cover **both** the new VictoriaLogs engine (2026 R1) **and** the legacy Log Server, so the same tool is correct regardless of which Milestone version you run.

---

## What it does

- **Antivirus exclusions** following Milestone best practice — Milestone folders, media file extensions, server processes, **VictoriaLogs** (2026 R1), and the **XProtect Smart Client** (added automatically only if the client is installed).
- **Storage drive checks** — 64 KB allocation unit (block size) verification and indexing disable on recording drives.
- **Performance hardening** — High Performance / Never Sleep power plan, unnecessary services disabled, `LargeSystemCache` for file-server throughput.
- **Bloatware removal** — Windows 25H2 apps (Clipchamp, Copilot, TikTok, Spotify, etc.), OneDrive, Teams, Outlook App, Office Hub. (Xbox Game Bar is deliberately *kept* — it backs Windows screen recording.)
- **Privacy** — Copilot, Recall and telemetry scheduled tasks disabled via Group Policy.
- **Remote Desktop** — enabled and tuned for performance.
- **NTP Time Server** — turns the machine into a reliable NTP server for your cameras, with the **correct peer format** (the `,0x9` flag bug from older scripts is fixed and auto-detected).
- **OpenSSH Server** — optional install and configuration (password auth enabled for admins).
- **MilestonePSTools** — installed via the official web installer (required for 2026 R1).
- **SNMP** — optional install.
- **Windows Update** — standalone on/off toggle.
- **Logging** — every run writes a detailed log plus a registry backup.

---

## Quick start

### Option A — run the executable (recommended for technicians)

Download `XProtect-Prepare-2.1.exe`, right-click → **Run as administrator**, and follow the menu.

### Option B — run the script

```powershell
# Run PowerShell as Administrator
Set-ExecutionPolicy RemoteSigned -Scope Process -Force
.\XProtect-Prepare-2.1.ps1
```

---

## Silent / unattended mode

For golden images and automated deployment, run everything (Complete Setup) without any prompts:

```powershell
# Complete Setup, no SSH, no extra storage-drive AV exclusions
.\XProtect-Prepare-2.1.exe -Silent

# Include the SSH server, and add whole-drive AV exclusions for D: and E:
.\XProtect-Prepare-2.1.exe -Silent -WithSSH -Drives "D,E"

# Add AV exclusions for ALL non-OS storage drives automatically
.\XProtect-Prepare-2.1.exe -Silent -Drives "storage"
```

| Parameter | Meaning |
|---|---|
| `-Silent` | Run Complete Setup with no interactive prompts |
| `-WithSSH` | Also install and configure the OpenSSH server |
| `-Drives "D,E"` | Whole-drive AV exclusions for the listed drives (or `"storage"` for all non-OS drives). Omit for Milestone folders only. |

The Milestone folder, process and extension exclusions (including VictoriaLogs and the Smart Client) are always applied in silent mode.

---

## Menu overview

```
=== XProtect-Prepare v2.1 ===
Detected: [Windows Server 2022 / Windows 11 ...]
[WARNING banner if the v1.9x NTP bug is detected]

1. Complete Setup (without SSH)
2. Complete Setup (with SSH activation)
3. Individual optimizations (submenu)
4. Configure NTP Time Server
5. Install & Configure SSH Server
6. Toggle Windows Update On/Off
7. NTP Recovery (fix v1.9x bug)
8. Exit
```

---

## What's new in v2.1 "Swedish Summer Edition"

- **NEW:** Full **Milestone XProtect 2026 R1** support.
- **NEW:** **VictoriaLogs** antivirus process exclusions (2026 R1 replaced the SQL Log Server). The log engine is auto-detected and reported.
- **NEW:** **XProtect Smart Client** antivirus exclusions (`Client.exe` + per-user cache) — added automatically only when the client is installed. Fixes the slow first launch after a (re)install caused by the antivirus scanning thousands of fresh client files.
- **CHANGED:** **MilestonePSTools** is now installed via the official one-line web installer (`milestonepstools.com/install.ps1`) — required for 2026 R1 compatibility.
- **NEW:** `-Silent` unattended mode for golden images and deployment automation.
- **COMPATIBILITY:** the legacy Log Server process exclusion is kept, so the AV configuration is correct on both 2026 R1 and earlier Milestone versions.

(See `CHANGELOG.md` for the full history, including the v2.0 NTP fix.)

---

## The NTP `,0x9` bug (fixed since v2.0)

Older v1.9x scripts wrote NTP peers as a plain comma-separated list:

```
0.se.pool.ntp.org,1.se.pool.ntp.org,2.se.pool.ntp.org,3.se.pool.ntp.org
```

The correct format needs `,0x9` flags after each peer (space-separated):

```
0.se.pool.ntp.org,0x9 1.se.pool.ntp.org,0x9 2.se.pool.ntp.org,0x9 3.se.pool.ntp.org,0x9
```

Without the flags, Windows Time Service cannot parse the peer list and time sync fails **silently** — bad news for forensically valid timestamps. The tool **detects this at startup** (warning banner), and **menu option 7 (NTP Recovery)** — or the standalone `NTP-Recovery.ps1` — fixes already-deployed machines.

---

## Files

| File | Description |
|------|-------------|
| `XProtect-Prepare-2.1.ps1` | Main script (v2.1) |
| `XProtect-Prepare-2.1.exe` | Compiled executable (branded, run as administrator) |
| `NTP-Recovery.ps1` | Standalone NTP fix for machines deployed with v1.9x |
| `CHANGELOG.md` | Full version history |
| `manvarg.ico` | Application icon used when compiling |

---

## Disclaimer

This tool makes significant system changes (services, registry, power, app removal). Test it before using it in production, and review the menu prompts. It is provided **as-is**, with no warranty, and is **not** affiliated with or supported by Milestone Systems A/S.

## Author

**Ulf Holmström** — Happy Problem Solver at **Manvarg AB**
For questions or feedback: **ulf@manvarg.se**
