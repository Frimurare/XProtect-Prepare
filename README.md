\# XProtect Prepare



PowerShell script to prepare and optimize Milestone XProtect installations on Windows.



\## Purpose

Automates the manual pre-installation tasks normally done by system integrators and administrators.



\## What it does

\- Checks Windows components (IIS, .NET, VC++ runtimes)

\- Verifies hostname, time sync, and network configuration

\- Adds or confirms firewall rules for XProtect

\- Logs all actions for troubleshooting



\## Usage

1\. Clone the repository:

&nbsp;  ```powershell

&nbsp;  git clone https://github.com/Frimurare/XProtect-Prepare.git

&nbsp;  cd XProtect-Prepare

2\. Set execution policy (required the first time on a new system):

&nbsp;  Set-ExecutionPolicy RemoteSigned -Scope Process -Force

3\. Run the script as Administrator:

&nbsp;  .\\Xprotect-prepare-19.ps1



Author



Ulf Holmström (Frimurare)

Technical Engineer – OpenEye / ex Solution Engineer at Milestone Systems

