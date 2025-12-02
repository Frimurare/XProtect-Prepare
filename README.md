# XProtect-Prepare v2

PowerShell toolkit for preparing Windows servers for Milestone XProtect. Version 2 introduces a modular layout, submenu-driven navigation, and a new entry script name: `XProtect-Prepare-v2.ps1`.

## Author Tag
All code in this repository is authored by **Ulf Holmström, ex employee, Solution Engineer at Milestone Systems December 2025**.

## Layout
- `XProtect-Prepare-v2.ps1` – entrypoint that loads modules and hosts the main menu.
- `modules/Core.psm1` – banner, admin check, logging helper, and main menu text.
- `modules/Storage.psm1` – antivirus exclusions, block-size checks, and indexing controls.
- `modules/Updates.psm1` – Windows Update status and registry backup helper.
- `modules/Bloatware.psm1` – removal routines for standard apps, OneDrive, Teams, Outlook app, Office Hub, and optional Xbox apps.
- `modules/Optimization.psm1` – power plan, service trimming, and RDP configuration.
- `modules/Network.psm1` – SNMP installation plus split time-service controls (NTP server and taskbar seconds).
- `modules/MilestoneTools.psm1` – Milestone PowerShell Tools installation.
- `modules/Setup.psm1` – one-click deployment that runs everything (including SNMP) in sequence.

## Running the script
1. Open PowerShell **as Administrator**.
2. Allow script execution for the current session (if needed):
   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
   ```
3. Run the entry script from the repository root:
   ```powershell
   .\XProtect-Prepare-v2.ps1
   ```
4. Use the main menu to run individual areas or choose **Complete Deployment** to run the full sequence.

## Building an EXE
`ps2exe` works with the modular structure. Keep the modules folder beside `XProtect-Prepare-v2.ps1` and run:
```powershell
ps2exe .\XProtect-Prepare-v2.ps1 -output XProtect-Prepare-v2.exe
```

## Key Improvements in v2
- Submenus for storage/antivirus, bloatware, optimization, and network/time services for clearer navigation.
- Split time controls: configure the Windows NTP server and taskbar seconds independently or together.
- Complete deployment now runs every step (including SNMP) without prompts and leaves Windows Update enabled afterward.
- Unified author tag across all modules and log output.
