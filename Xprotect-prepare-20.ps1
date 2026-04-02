<#
   Enhanced Script for Milestone XProtect System Configuration
   By Ulf Holmstrom, Happy Problem Solver at Manvarg AB (Rev 2.0, 2026)
   For questions contact: ulf@manvarg.se

   IMPORTANT DISCLAIMER:
   This script has been developed by Ulf Holmstrom, Happy Problem Solver at Manvarg AB, to facilitate resellers.
   It is provided as a public resource and is NOT supported by Milestone Systems.

   CHANGELOG Rev 2.0:
   - NEW: Menu restructured with Complete Setup (with/without SSH)
   - CRITICAL FIX: NTP peer format corrected (added ,0x9 flags)
   - NEW: NTP auto-detects domain vs standalone for correct sync mode
   - NEW: NTP Recovery function to fix v1.9x bug on deployed machines
   - NEW: SSH Server install & configure option
   - NEW: Copilot and Recall disabled via Group Policy
   - NEW: Telemetry scheduled tasks disabled
   - NEW: LargeSystemCache registry optimization
   - NEW: Windows 25H2 bloatware additions (Clipchamp, Copilot, TikTok, etc.)
   - IMPROVED: Windows Update toggle moved to standalone menu option (not in Complete Setup)
   - FIXED: All Get-WmiObject replaced with Get-CimInstance
   - FIXED: ErrorAction placement on Where-Object pipelines
   - FIXED: AV exclusions added regardless of folder existence
   - FIXED: Auto-detect Windows version at startup

   CHANGELOG Rev 1.92:
   - CRITICAL FIX: Added Milestone installation folder exceptions on C: drive
   - ADDED: C:\Program Files\Milestone\ antivirus exception
   - ADDED: C:\Program Files (x86)\Milestone\ antivirus exception
   - ADDED: C:\ProgramData\Milestone\ antivirus exception
   - ADDED: C:\ProgramData\VideoDeviceDrivers\ antivirus exception
   - ADDED: C:\ProgramData\VideoOS\ antivirus exception
   - ADDED: Missing file extensions (.pic, .pqz, .sts, .ts) for XProtect Enterprise
   - IMPROVED: Now follows Milestone's official best practices completely
   - SECURITY: C: drive bulk exclusion still blocked - only specific Milestone folders

   CHANGELOG Rev 1.9:
   - ADDED: NTP Time Server configuration for camera time synchronization
   - CRITICAL: Enables Windows Time Service as reliable NTP server for cameras
   - ADDED: Automatic firewall configuration for NTP (UDP port 123)
   - IMPROVED: Cameras can now sync accurate time from this server
   - FORENSIC: Ensures legally valid timestamps on all recordings
#>

# Check for Administrator privileges
function Test-Administrator {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Administrator)) {
    Write-Host "You must run this script as an Administrator. Please re-run with elevated privileges." -ForegroundColor Red
    Read-Host "Press Enter to exit..."
    exit
}

# Auto-detect Windows version
$global:windowsVersion = (Get-CimInstance Win32_OperatingSystem).Caption

# NTP bug detection at startup
$global:ntpBugDetected = $false
try {
    $ntpServerValue = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters" -Name "NtpServer" -ErrorAction SilentlyContinue).NtpServer
    if ($ntpServerValue -and $ntpServerValue -match "," -and $ntpServerValue -notmatch ",0x") {
        $global:ntpBugDetected = $true
    }
} catch {
    # Ignore - NTP may not be configured
}

# Function: Play-Fanfare
function Play-Fanfare {
    $melody = @(
        @{ Note = 523; Duration = 300 },
        @{ Note = 523; Duration = 300 },
        @{ Note = 392; Duration = 300 },
        @{ Note = 440; Duration = 300 },
        @{ Note = 392; Duration = 300 }
    )
    foreach ($tone in $melody) {
        [Console]::Beep($tone.Note, $tone.Duration)
        Start-Sleep -Milliseconds 50
    }
}

# Function: Get Available Drives
function Get-AvailableDrives {
    return Get-CimInstance -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 } | Select-Object DeviceID, @{Name="Size(GB)";Expression={[math]::Round($_.Size/1GB,2)}}, @{Name="FreeSpace(GB)";Expression={[math]::Round($_.FreeSpace/1GB,2)}}
}

# Function: Check Drive Block Size
function Test-DriveBlockSize {
    param([string]$Drive)
    try {
        $fsInfo = fsutil fsinfo ntfsinfo $Drive 2>&1 | Out-String
        if ($fsInfo -match "Bytes Per Cluster\s*:\s*(\d+)") {
            return [int]$matches[1]
        } elseif ($fsInfo -match "Bytes per cluster\s*:\s*(\d+)") {
            return [int]$matches[1]
        } else {
            return $null
        }
    }
    catch {
        return $null
    }
}

# Function: Check Drive Indexing Status
function Test-DriveIndexing {
    param([string]$Drive)
    try {
        $driveLetter = $Drive.TrimEnd('\')
        $volume = Get-CimInstance -Class Win32_Volume | Where-Object { $_.DriveLetter -eq $driveLetter }
        if ($volume) {
            return $volume.IndexingEnabled
        }
        return $null
    }
    catch {
        return $null
    }
}

# Function: Disable Drive Indexing
function Disable-DriveIndexing {
    param([string]$Drive)
    try {
        $driveLetter = $Drive.TrimEnd('\')
        $volume = Get-CimInstance -Class Win32_Volume | Where-Object { $_.DriveLetter -eq $driveLetter }
        if ($volume) {
            $volume | Set-CimInstance -Property @{ IndexingEnabled = $false }
            return $true
        }
        return $false
    }
    catch {
        return $false
    }
}

# Function: Get Windows Update Status
function Get-WindowsUpdateStatus {
    try {
        $regPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
        if (Test-Path $regPath) {
            $noAutoUpdate = Get-ItemProperty -Path $regPath -Name "NoAutoUpdate" -ErrorAction SilentlyContinue
            return $noAutoUpdate.NoAutoUpdate -ne 1
        }
        return $true
    }
    catch {
        return $null
    }
}

# Function: Set Windows Update Status
function Set-WindowsUpdateStatus {
    param([bool]$Enable)
    try {
        $regPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }

        if ($Enable) {
            Remove-ItemProperty -Path $regPath -Name "NoAutoUpdate" -ErrorAction SilentlyContinue
            Remove-ItemProperty -Path $regPath -Name "AUOptions" -ErrorAction SilentlyContinue
        } else {
            Set-ItemProperty -Path $regPath -Name "NoAutoUpdate" -Value 1 -Type DWord
            Set-ItemProperty -Path $regPath -Name "AUOptions" -Value 1 -Type DWord
        }

        Restart-Service -Name "wuauserv" -Force -ErrorAction SilentlyContinue
        return $true
    }
    catch {
        return $false
    }
}

# Function: Create Registry Backup
function New-RegistryBackup {
    param([string]$BackupName)
    try {
        $backupFolder = "C:\Milestonecheck-ulfh\RegistryBackups"
        if (-not (Test-Path $backupFolder)) {
            New-Item -ItemType Directory -Path $backupFolder -Force | Out-Null
        }

        $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")

        $regKeys = @(
            "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate",
            "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services"
        )

        foreach ($key in $regKeys) {
            $keyFileName = $key.Replace("\", "_").Replace(":", "")
            $keyBackupFile = Join-Path $backupFolder "${BackupName}_${keyFileName}_${timestamp}.reg"
            reg export $key $keyBackupFile /y 2>$null
        }

        Write-Host "Registry backup created in: $backupFolder" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Failed to create registry backup: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function: Configure Power Management
function Set-PowerManagement {
    param([bool]$HighPerformance)
    try {
        if ($HighPerformance) {
            powercfg /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c
            powercfg /setacvalueindex SCHEME_CURRENT 238c9fa8-0aad-41ed-83f4-97be242c8f20 29f6c1db-86da-48c5-9fdb-f2b67b1f44da 0
            powercfg /setdcvalueindex SCHEME_CURRENT 238c9fa8-0aad-41ed-83f4-97be242c8f20 29f6c1db-86da-48c5-9fdb-f2b67b1f44da 0
            powercfg /setacvalueindex SCHEME_CURRENT 7516b95f-f776-4464-8c53-06167f40cc99 3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e 0
            powercfg /setdcvalueindex SCHEME_CURRENT 7516b95f-f776-4464-8c53-06167f40cc99 3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e 0
            powercfg /setacvalueindex SCHEME_CURRENT 0012ee47-9041-4b5d-9b77-535fba8b1442 6738e2c4-e8a5-4a42-b16a-e040e769756e 0
            powercfg /setdcvalueindex SCHEME_CURRENT 0012ee47-9041-4b5d-9b77-535fba8b1442 6738e2c4-e8a5-4a42-b16a-e040e769756e 0
            powercfg /setactive SCHEME_CURRENT
            Write-Host "Power management optimized: High Performance + Never Sleep Mode (Critical for 24/7 recording)" -ForegroundColor Green
            return $true
        } else {
            powercfg /setactive 381b4222-f694-41f0-9685-ff5bb260df2e
            Write-Host "Power management set to Balanced mode" -ForegroundColor Green
            return $true
        }
    }
    catch {
        Write-Host "Failed to configure power management: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function: Configure Windows Services
function Set-WindowsServices {
    param([bool]$OptimizeForMilestone)

    $servicesToDisable = @(
        @{ Name = "Fax"; Display = "Fax Service" },
        @{ Name = "WSearch"; Display = "Windows Search" },
        @{ Name = "Spooler"; Display = "Print Spooler" },
        @{ Name = "SysMain"; Display = "Superfetch/SysMain" },
        @{ Name = "HomeGroupListener"; Display = "HomeGroup Listener" },
        @{ Name = "HomeGroupProvider"; Display = "HomeGroup Provider" },
        @{ Name = "Themes"; Display = "Themes" },
        @{ Name = "TabletInputService"; Display = "Tablet PC Input Service" }
    )

    $disabledCount = 0
    foreach ($service in $servicesToDisable) {
        try {
            $svc = Get-Service -Name $service.Name -ErrorAction SilentlyContinue
            if ($svc) {
                if ($OptimizeForMilestone) {
                    if ($svc.Status -eq 'Running') {
                        Stop-Service -Name $service.Name -Force -ErrorAction SilentlyContinue
                    }
                    Set-Service -Name $service.Name -StartupType Disabled -ErrorAction Stop
                    Write-Host " - Disabled: $($service.Display)" -ForegroundColor Green
                    $disabledCount++
                } else {
                    Set-Service -Name $service.Name -StartupType Manual -ErrorAction Stop
                    Write-Host " - Restored: $($service.Display) to Manual" -ForegroundColor Green
                }
            }
        }
        catch {
            Write-Host " - Warning: Could not configure $($service.Display) - Service may not exist" -ForegroundColor Yellow
        }
    }
    return $disabledCount
}

# Function: Configure Remote Desktop
function Set-RemoteDesktopOptimization {
    param([bool]$OptimizeRDP)
    try {
        if ($OptimizeRDP) {
            Write-Host " - Enabling Remote Desktop..." -ForegroundColor Cyan
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -Value 0

            Write-Host " - Configuring Windows Firewall for RDP..." -ForegroundColor Cyan
            Enable-NetFirewallRule -DisplayGroup "Remote Desktop" -ErrorAction SilentlyContinue

            Write-Host " - Optimizing RDP for performance..." -ForegroundColor Cyan
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" -Name "fDisableWallpaper" -Value 1 -ErrorAction SilentlyContinue
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" -Name "fDisableFullWindowDrag" -Value 1 -ErrorAction SilentlyContinue
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" -Name "fDisableMenuAnims" -Value 1 -ErrorAction SilentlyContinue

            Write-Host "[OK] Remote Desktop ENABLED and optimized" -ForegroundColor Green
            return $true
        } else {
            Write-Host "[OK] Remote Desktop settings restored to defaults" -ForegroundColor Green
            return $true
        }
    }
    catch {
        Write-Host "[FAIL] Failed to configure Remote Desktop: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function: Remove Bloatware App
function Remove-BloatwareApp {
    param (
        [string]$AppName,
        [string]$FriendlyName
    )
    try {
        $packages = Get-AppxPackage -Name "*$AppName*" -AllUsers -ErrorAction SilentlyContinue
        $provisioned = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*$AppName*" }

        if ($packages -or $provisioned) {
            Write-Host " - Removing: $FriendlyName..." -ForegroundColor Yellow

            if ($packages) {
                foreach ($package in $packages) {
                    $package | Remove-AppxPackage -AllUsers -ErrorAction Stop
                    Write-Host " - $FriendlyName removed for users." -ForegroundColor Green
                }
            }

            if ($provisioned) {
                foreach ($prov in $provisioned) {
                    $prov | Remove-AppxProvisionedPackage -Online -ErrorAction Stop
                    Write-Host " - $FriendlyName removed from provisioned apps." -ForegroundColor Green
                }
            }
            return $true
        } else {
            Write-Host " - $FriendlyName not found." -ForegroundColor Gray
            return $false
        }
    } catch {
        Write-Host " - Error removing $FriendlyName - $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function: Remove OneDrive
function Remove-OneDrive {
    try {
        Write-Host " - Stopping OneDrive processes..." -ForegroundColor Yellow
        Stop-Process -Name "OneDrive" -Force -ErrorAction SilentlyContinue

        Write-Host " - Uninstalling OneDrive..." -ForegroundColor Yellow

        $onedriveSetupPaths = @(
            "$env:SYSTEMROOT\System32\OneDriveSetup.exe",
            "$env:SYSTEMROOT\SysWOW64\OneDriveSetup.exe",
            "$env:PROGRAMFILES\Microsoft OneDrive\OneDriveSetup.exe",
            "$env:PROGRAMFILES(x86)\Microsoft OneDrive\OneDriveSetup.exe"
        )

        $uninstalled = $false
        foreach ($path in $onedriveSetupPaths) {
            if (Test-Path $path) {
                Start-Process -FilePath $path -ArgumentList "/uninstall" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
                Write-Host " - OneDrive uninstalled using: $path" -ForegroundColor Green
                $uninstalled = $true
                break
            }
        }

        if (-not $uninstalled) {
            Write-Host " - OneDrive setup not found - may already be uninstalled" -ForegroundColor Yellow
        }

        $folders = @(
            "$env:LOCALAPPDATA\Microsoft\OneDrive",
            "$env:APPDATA\Microsoft\OneDrive",
            "$env:PROGRAMDATA\Microsoft OneDrive",
            "$env:USERPROFILE\OneDrive"
        )
        foreach ($folder in $folders) {
            if (Test-Path $folder) {
                Remove-Item -Path $folder -Recurse -Force -ErrorAction SilentlyContinue
                Write-Host " - Deleted OneDrive folder: $folder" -ForegroundColor Green
            }
        }
        return $true
    } catch {
        Write-Host " - Error uninstalling OneDrive: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function: Remove Microsoft Teams (All Versions)
function Remove-MicrosoftTeams {
    try {
        Write-Host " - Stopping Teams processes..." -ForegroundColor Yellow
        Get-Process -Name "*Teams*" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

        $teamsRemoved = $false

        Write-Host " - Removing Teams AppX packages (new Teams)..." -ForegroundColor Yellow
        $teamsPackages = Get-AppxPackage -Name "*Teams*" -AllUsers -ErrorAction SilentlyContinue
        if ($teamsPackages) {
            foreach ($package in $teamsPackages) {
                $package | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
                Write-Host " - Teams package removed: $($package.Name)" -ForegroundColor Green
                $teamsRemoved = $true
            }
        }

        Write-Host " - Removing Teams provisioned packages..." -ForegroundColor Yellow
        $teamsProvisioned = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*Teams*" }
        if ($teamsProvisioned) {
            foreach ($prov in $teamsProvisioned) {
                $prov | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
                Write-Host " - Teams provisioned package removed: $($prov.DisplayName)" -ForegroundColor Green
                $teamsRemoved = $true
            }
        }

        Write-Host " - Removing classic Teams installation..." -ForegroundColor Yellow
        $teamsUninstallerPaths = @(
            "$env:LOCALAPPDATA\Microsoft\Teams\Update.exe",
            "$env:APPDATA\Microsoft\Teams\Update.exe"
        )

        foreach ($path in $teamsUninstallerPaths) {
            if (Test-Path $path) {
                Start-Process -FilePath $path -ArgumentList "--uninstall" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
                Write-Host " - Classic Teams installation removed" -ForegroundColor Green
                $teamsRemoved = $true
                break
            }
        }

        $teamsFolders = @(
            "$env:LOCALAPPDATA\Microsoft\Teams",
            "$env:APPDATA\Microsoft\Teams",
            "$env:PROGRAMDATA\Microsoft\Teams"
        )

        foreach ($folder in $teamsFolders) {
            if (Test-Path $folder) {
                Remove-Item -Path $folder -Recurse -Force -ErrorAction SilentlyContinue
                Write-Host " - Deleted Teams folder: $folder" -ForegroundColor Green
            }
        }

        if (-not $teamsRemoved) {
            Write-Host " - Microsoft Teams not found or already removed." -ForegroundColor Gray
        }

        return $teamsRemoved
    } catch {
        Write-Host " - Error uninstalling Microsoft Teams: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function: Remove Outlook App
function Remove-OutlookApp {
    try {
        Write-Host " - Removing Outlook App..." -ForegroundColor Yellow

        $outlookRemoved = $false

        $outlookPackages = Get-AppxPackage -Name "*OutlookForWindows*" -AllUsers -ErrorAction SilentlyContinue
        if ($outlookPackages) {
            foreach ($package in $outlookPackages) {
                $package | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
                Write-Host " - Outlook App removed: $($package.Name)" -ForegroundColor Green
                $outlookRemoved = $true
            }
        }

        $outlookProvisioned = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*OutlookForWindows*" }
        if ($outlookProvisioned) {
            foreach ($prov in $outlookProvisioned) {
                $prov | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
                Write-Host " - Outlook App provisioned package removed" -ForegroundColor Green
                $outlookRemoved = $true
            }
        }

        if (-not $outlookRemoved) {
            Write-Host " - Outlook App not found or already removed." -ForegroundColor Gray
        }

        return $outlookRemoved
    } catch {
        Write-Host " - Error removing Outlook App: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function: Remove Office Hub/365 Installer
function Remove-OfficeHub {
    try {
        Write-Host " - Removing Office Hub/365 installer..." -ForegroundColor Yellow

        $officeRemoved = $false

        $officePackages = Get-AppxPackage -Name "*OfficeHub*" -AllUsers -ErrorAction SilentlyContinue
        if ($officePackages) {
            foreach ($package in $officePackages) {
                $package | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
                Write-Host " - Office Hub removed: $($package.Name)" -ForegroundColor Green
                $officeRemoved = $true
            }
        }

        $officeProvisioned = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*OfficeHub*" }
        if ($officeProvisioned) {
            foreach ($prov in $officeProvisioned) {
                $prov | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
                Write-Host " - Office Hub provisioned package removed" -ForegroundColor Green
                $officeRemoved = $true
            }
        }

        if (-not $officeRemoved) {
            Write-Host " - Office Hub not found or already removed." -ForegroundColor Gray
        }

        return $officeRemoved
    } catch {
        Write-Host " - Error removing Office Hub: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function: Remove Xbox Apps (Optional - with warning)
function Remove-XboxApps {
    Write-Host "`nWARNING: Xbox Apps Removal" -ForegroundColor Red
    Write-Host "==========================================" -ForegroundColor Red
    Write-Host "Xbox Game Bar is used by Windows for:" -ForegroundColor Yellow
    Write-Host " - Screen recording (Win+G)" -ForegroundColor Yellow
    Write-Host " - Game capture functionality" -ForegroundColor Yellow
    Write-Host " - Some system overlay features" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Removing these may cause:" -ForegroundColor Yellow
    Write-Host " - Popup messages about missing components" -ForegroundColor Yellow
    Write-Host " - Loss of built-in screen recording" -ForegroundColor Yellow
    Write-Host " - Issues with some applications expecting Game Bar" -ForegroundColor Yellow
    Write-Host "==========================================" -ForegroundColor Red
    Write-Host ""

    $confirm = Read-Host "Are you SURE you want to remove Xbox apps? (Type 'YES' to confirm)"

    if ($confirm -ne "YES") {
        Write-Host "Xbox apps removal cancelled." -ForegroundColor Green
        return $false
    }

    try {
        $xboxApps = @(
            @{ Name = "Microsoft.XboxApp"; Display = "Xbox App" },
            @{ Name = "Microsoft.XboxIdentityProvider"; Display = "Xbox Identity Provider" },
            @{ Name = "Microsoft.XboxGameOverlay"; Display = "Xbox Game Overlay" },
            @{ Name = "Microsoft.XboxGamingOverlay"; Display = "Xbox Gaming Overlay" },
            @{ Name = "Microsoft.XboxSpeechToTextOverlay"; Display = "Xbox Speech To Text" }
        )

        $removedCount = 0
        foreach ($app in $xboxApps) {
            if (Remove-BloatwareApp -AppName $app.Name -FriendlyName $app.Display) {
                $removedCount++
            }
        }

        Write-Host "`nXbox apps removal completed. $removedCount apps removed." -ForegroundColor Green
        return $true

    } catch {
        Write-Host " - Error removing Xbox apps: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function: Install Milestone PowerShell Tools
function Install-MilestonePSTools {
    try {
        Write-Host " - Checking for Milestone PowerShell Tools..." -ForegroundColor Cyan

        $existingModule = Get-Module -ListAvailable -Name "MilestonePSTools" -ErrorAction SilentlyContinue
        if ($existingModule) {
            Write-Host " - Milestone PowerShell Tools already installed (Version: $($existingModule.Version))" -ForegroundColor Green
            return $true
        }

        Write-Host " - Installing Milestone PowerShell Tools from PowerShell Gallery..." -ForegroundColor Cyan

        $nugetProvider = Get-PackageProvider -Name "NuGet" -ErrorAction SilentlyContinue
        if (-not $nugetProvider) {
            Write-Host " - Installing NuGet package provider..." -ForegroundColor Yellow
            Install-PackageProvider -Name "NuGet" -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser
        }

        $originalPolicy = Get-PSRepository -Name "PSGallery" | Select-Object -ExpandProperty InstallationPolicy
        Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted

        try {
            Write-Host " - Downloading and installing MilestonePSTools module..." -ForegroundColor Yellow
            Install-Module -Name "MilestonePSTools" -Scope CurrentUser -Force -AllowClobber

            $installedModule = Get-Module -ListAvailable -Name "MilestonePSTools" -ErrorAction SilentlyContinue
            if ($installedModule) {
                Write-Host " - Milestone PowerShell Tools successfully installed!" -ForegroundColor Green
                Write-Host " - Version: $($installedModule.Version)" -ForegroundColor Green
                return $true
            } else {
                Write-Host " - Installation verification failed" -ForegroundColor Red
                return $false
            }
        }
        finally {
            Set-PSRepository -Name "PSGallery" -InstallationPolicy $originalPolicy
        }

    } catch {
        Write-Host " - Error installing Milestone PowerShell Tools: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function: Enable SNMP
function Enable-SNMP {
    try {
        Write-Host " - Installing SNMP Client capability..." -ForegroundColor Cyan
        Write-Host "   NOTE: This may take a while - please be patient..." -ForegroundColor Yellow
        Add-WindowsCapability -Online -Name "SNMP.Client~~~~0.0.1.0" -ErrorAction Stop | Out-Null
        Write-Host " - Installing WMI SNMP Provider..." -ForegroundColor Cyan
        Add-WindowsCapability -Online -Name "WMI-SNMP-Provider.Client~~~~0.0.1.0" -ErrorAction Stop | Out-Null
        Write-Host "[OK] SNMP capabilities installed successfully" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "[FAIL] Failed to install SNMP: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function: Configure NTP Server for Cameras (v2.0 FIXED - correct peer format + domain detection)
function Enable-NTPServer {
    try {
        Write-Host "`n=== CONFIGURING NTP TIME SERVER FOR CAMERAS ===" -ForegroundColor Yellow
        Write-Host "This configures Windows Time Service to act as an NTP server" -ForegroundColor Cyan
        Write-Host "for your CCTV cameras. Accurate time synchronization is CRITICAL" -ForegroundColor Cyan
        Write-Host "for surveillance systems - timestamps must be forensically accurate." -ForegroundColor Cyan
        Write-Host ""

        # Stop the Windows Time service
        Write-Host " - Stopping Windows Time service..." -ForegroundColor Cyan
        Stop-Service w32time -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2

        # Configure Windows Time service as NTP server
        Write-Host " - Configuring time synchronization settings..." -ForegroundColor Cyan

        # Set this server as a reliable time source
        # AnnounceFlags=10 (0x0A): Act as NTP server for cameras but ALSO sync from external sources
        # Value 5 made the server think it IS the primary source and refuse to sync upstream
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Config" -Name "AnnounceFlags" -Value 10 -Type DWord -ErrorAction Stop

        # Enable NTP Server
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\NtpServer" -Name "Enabled" -Value 1 -Type DWord -ErrorAction Stop

        # Configure NTP client to sync from Swedish NTP pool (FIXED: correct peer format with ,0x9 flags)
        Write-Host " - Setting Swedish NTP pool servers as time source..." -ForegroundColor Cyan
        $ntpServers = "0.se.pool.ntp.org,0x9 1.se.pool.ntp.org,0x9 2.se.pool.ntp.org,0x9 3.se.pool.ntp.org,0x9"
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters" -Name "NtpServer" -Value $ntpServers -Type String -ErrorAction Stop

        # Auto-detect domain membership for correct sync mode
        $isDomain = (Get-CimInstance Win32_ComputerSystem).PartOfDomain
        if ($isDomain) {
            Write-Host " - Domain member detected - using AllSync mode" -ForegroundColor Cyan
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters" -Name "Type" -Value "AllSync" -Type String -ErrorAction Stop
            $syncFlags = "ALL"
        } else {
            Write-Host " - Standalone server detected - using NTP mode" -ForegroundColor Cyan
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters" -Name "Type" -Value "NTP" -Type String -ErrorAction Stop
            $syncFlags = "manual"
        }

        # Set w32time to Automatic startup
        Set-Service w32time -StartupType Automatic

        # Configure time synchronization via w32tm
        Write-Host " - Applying w32tm configuration..." -ForegroundColor Cyan
        & w32tm /config /manualpeerlist:$ntpServers /syncfromflags:$syncFlags /reliable:yes /update 2>&1 | Out-Null

        # Re-apply AnnounceFlags=10 AFTER w32tm (reliable:yes resets it to 5)
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Config" -Name "AnnounceFlags" -Value 10 -Type DWord -ErrorAction Stop

        # Configure firewall rule for NTP (UDP 123)
        Write-Host " - Opening firewall for NTP (UDP port 123)..." -ForegroundColor Cyan

        # Remove existing rule if it exists
        netsh advfirewall firewall delete rule name="NTP Server (XProtect)" 2>&1 | Out-Null

        # Add new firewall rule
        $firewallResult = netsh advfirewall firewall add rule name="NTP Server (XProtect)" dir=in action=allow protocol=UDP localport=123 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host " - Firewall rule created successfully" -ForegroundColor Green
        } else {
            Write-Host " - Warning: Firewall rule creation returned code $LASTEXITCODE" -ForegroundColor Yellow
        }

        # Enable seconds display in Windows taskbar clock
        Write-Host " - Enabling seconds display in taskbar clock..." -ForegroundColor Cyan
        $registryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        try {
            if (-not (Test-Path $registryPath)) {
                New-Item -Path $registryPath -Force | Out-Null
            }

            Set-ItemProperty -Path $registryPath -Name "ShowSecondsInSystemClock" -Value 1 -Type DWord -ErrorAction Stop
            Write-Host " - Seconds display enabled in taskbar clock" -ForegroundColor Green
            Write-Host " - NOTE: Explorer must be restarted for clock change to take effect" -ForegroundColor Yellow

            $restartExplorer = Read-Host " - Restart Windows Explorer now to show seconds? (Y/N)"
            if ($restartExplorer -match "^[Yy]") {
                Write-Host " - Restarting Windows Explorer..." -ForegroundColor Cyan
                Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 2
                Start-Process explorer
                Write-Host " - Windows Explorer restarted - seconds should now be visible!" -ForegroundColor Green
            } else {
                Write-Host " - Seconds will be visible after next Explorer restart or system reboot" -ForegroundColor Yellow
            }
        }
        catch {
            Write-Host " - Warning: Could not enable seconds display: $($_.Exception.Message)" -ForegroundColor Yellow
        }

        # Start the Windows Time service
        Write-Host " - Starting Windows Time service..." -ForegroundColor Cyan
        Start-Service w32time -ErrorAction Stop
        Start-Sleep -Seconds 3

        # Force immediate synchronization
        Write-Host " - Forcing immediate time synchronization..." -ForegroundColor Cyan
        & w32tm /resync /rediscover 2>&1 | Out-Null
        Start-Sleep -Seconds 2

        # Ensure DST is enabled
        Write-Host " - Checking Daylight Saving Time..." -ForegroundColor Cyan
        $dstReg = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\TimeZoneInformation" -Name "DynamicDaylightTimeDisabled" -ErrorAction SilentlyContinue
        if ($dstReg.DynamicDaylightTimeDisabled -eq 1) {
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\TimeZoneInformation" -Name "DynamicDaylightTimeDisabled" -Value 0 -Type DWord
            Write-Host " - DST was disabled - FIXED" -ForegroundColor Yellow
        } else {
            Write-Host " - DST enabled (OK)" -ForegroundColor Green
        }

        # Re-apply timezone to force DST recalculation
        $currentTZ = (Get-TimeZone).Id
        Set-TimeZone -Id $currentTZ

        # Verify configuration
        Write-Host "`n - Verifying NTP server configuration..." -ForegroundColor Cyan
        $w32tmStatus = & w32tm /query /status 2>&1
        $w32tmConfig = & w32tm /query /configuration 2>&1

        # Check if service is running
        $timeService = Get-Service w32time -ErrorAction SilentlyContinue
        if ($timeService.Status -eq 'Running') {
            Write-Host "[OK] Windows Time service is running" -ForegroundColor Green
        } else {
            Write-Host "[WARN] Windows Time service status: $($timeService.Status)" -ForegroundColor Yellow
        }

        $isDST = [System.TimeZoneInfo]::Local.IsDaylightSavingTime((Get-Date))
        $utcOffset = (Get-TimeZone).GetUtcOffset((Get-Date))

        # Display configuration summary
        Write-Host "`n=== NTP SERVER CONFIGURATION SUMMARY ===" -ForegroundColor Green
        Write-Host "NTP Server Status: ENABLED" -ForegroundColor Green
        Write-Host "Time Source: Swedish NTP Pool (se.pool.ntp.org)" -ForegroundColor Green
        Write-Host "Peer Format: Correct (with ,0x9 flags)" -ForegroundColor Green
        Write-Host "Sync Mode: $(if($isDomain){'AllSync (domain member)'}else{'NTP (standalone)'})" -ForegroundColor Green
        Write-Host "Startup Type: Automatic" -ForegroundColor Green
        Write-Host "Firewall: UDP Port 123 opened for incoming connections" -ForegroundColor Green
        Write-Host "Server Mode: Reliable NTP server for camera synchronization" -ForegroundColor Green
        Write-Host "Taskbar Clock: Seconds display ENABLED" -ForegroundColor Green
        Write-Host "DST: $(if($isDST){'Active (summer time)'}else{'Inactive (winter time)'}) UTC+$($utcOffset.Hours)" -ForegroundColor Green
        Write-Host ""
        Write-Host "Your cameras can now sync time from this server!" -ForegroundColor Green
        Write-Host "Configure your cameras to use this server's IP address as NTP server." -ForegroundColor Cyan
        Write-Host ""
        Write-Host "To verify NTP is working, run these commands:" -ForegroundColor Yellow
        Write-Host "  w32tm /query /status" -ForegroundColor Gray
        Write-Host "  w32tm /query /configuration" -ForegroundColor Gray
        Write-Host ""

        return $true

    } catch {
        Write-Host "[FAIL] Failed to configure NTP server: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Attempting to start Windows Time service anyway..." -ForegroundColor Yellow
        Start-Service w32time -ErrorAction SilentlyContinue
        return $false
    }
}

# Function: NTP Recovery (fix v1.9x bug)
function Invoke-NTPRecovery {
    Write-Host "`n=== NTP RECOVERY - FIX v1.9x BUG ===" -ForegroundColor Yellow
    Write-Host "This fixes the NTP peer format bug introduced in v1.9x scripts." -ForegroundColor Cyan
    Write-Host "The bug used commas without ,0x9 flags, causing sync failures." -ForegroundColor Cyan
    Write-Host ""

    # Check current state
    try {
        $currentPeers = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters" -Name "NtpServer" -ErrorAction Stop).NtpServer
        $currentType = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters" -Name "Type" -ErrorAction Stop).Type

        Write-Host "Current NTP configuration:" -ForegroundColor Cyan
        Write-Host "  NtpServer: $currentPeers" -ForegroundColor White
        Write-Host "  Type: $currentType" -ForegroundColor White
        Write-Host ""

        if ($currentPeers -match ",0x") {
            Write-Host "[OK] NTP peer format appears correct (has ,0x flags)." -ForegroundColor Green
            Write-Host "No recovery needed unless you want to force re-apply." -ForegroundColor Yellow
            $force = Read-Host "Force re-apply anyway? (Y/N)"
            if ($force -notmatch "^[Yy]") {
                return $true
            }
        } else {
            Write-Host "[BUG DETECTED] NTP peers are missing ,0x9 flags!" -ForegroundColor Red
            Write-Host "This causes Windows Time service to fail NTP synchronization." -ForegroundColor Red
        }
    } catch {
        Write-Host "Could not read current NTP configuration." -ForegroundColor Yellow
    }

    Write-Host ""
    $confirm = Read-Host "Apply NTP recovery fix? (Y/N)"
    if ($confirm -notmatch "^[Yy]") {
        Write-Host "NTP recovery cancelled." -ForegroundColor Yellow
        return $false
    }

    try {
        Write-Host "`n - Stopping Windows Time service..." -ForegroundColor Cyan
        Stop-Service w32time -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2

        # Fix peer format
        $ntpServers = "0.se.pool.ntp.org,0x9 1.se.pool.ntp.org,0x9 2.se.pool.ntp.org,0x9 3.se.pool.ntp.org,0x9"
        Write-Host " - Setting correct NTP peer format..." -ForegroundColor Cyan
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters" -Name "NtpServer" -Value $ntpServers -Type String -ErrorAction Stop

        # Auto-detect domain membership
        $isDomain = (Get-CimInstance Win32_ComputerSystem).PartOfDomain
        if ($isDomain) {
            Write-Host " - Domain member detected - setting AllSync mode" -ForegroundColor Cyan
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters" -Name "Type" -Value "AllSync" -Type String -ErrorAction Stop
            $syncFlags = "ALL"
        } else {
            Write-Host " - Standalone server detected - setting NTP mode" -ForegroundColor Cyan
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters" -Name "Type" -Value "NTP" -Type String -ErrorAction Stop
            $syncFlags = "manual"
        }

        # Set w32time to Automatic
        Write-Host " - Setting w32time startup to Automatic..." -ForegroundColor Cyan
        Set-Service w32time -StartupType Automatic

        # Set AnnounceFlags=10 (serve time to cameras AND sync from external)
        Write-Host " - Setting AnnounceFlags to 10 (server + client)..." -ForegroundColor Cyan
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Config" -Name "AnnounceFlags" -Value 10 -Type DWord -ErrorAction SilentlyContinue

        # Apply via w32tm
        Write-Host " - Applying w32tm configuration..." -ForegroundColor Cyan
        & w32tm /config /manualpeerlist:$ntpServers /syncfromflags:$syncFlags /reliable:yes /update 2>&1 | Out-Null

        # Re-apply AnnounceFlags (w32tm /reliable:yes resets it to 5)
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Config" -Name "AnnounceFlags" -Value 10 -Type DWord -ErrorAction SilentlyContinue

        # Start service
        Write-Host " - Starting Windows Time service..." -ForegroundColor Cyan
        Start-Service w32time -ErrorAction Stop
        Start-Sleep -Seconds 3

        # Force resync
        Write-Host " - Forcing time resynchronization..." -ForegroundColor Cyan
        & w32tm /resync /rediscover 2>&1 | Out-Null
        Start-Sleep -Seconds 2

        # Ensure DST is enabled
        Write-Host " - Checking Daylight Saving Time..." -ForegroundColor Cyan
        $dstReg = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\TimeZoneInformation" -Name "DynamicDaylightTimeDisabled" -ErrorAction SilentlyContinue
        if ($dstReg.DynamicDaylightTimeDisabled -eq 1) {
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\TimeZoneInformation" -Name "DynamicDaylightTimeDisabled" -Value 0 -Type DWord
            Write-Host " - DST was disabled - FIXED" -ForegroundColor Yellow
        } else {
            Write-Host " - DST enabled (OK)" -ForegroundColor Green
        }

        # Re-apply timezone to force DST recalculation
        $currentTZ = (Get-TimeZone).Id
        Set-TimeZone -Id $currentTZ

        # Verify
        $timeService = Get-Service w32time -ErrorAction SilentlyContinue
        $newPeers = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters" -Name "NtpServer" -ErrorAction SilentlyContinue).NtpServer
        $isDST = [System.TimeZoneInfo]::Local.IsDaylightSavingTime((Get-Date))
        $utcOffset = (Get-TimeZone).GetUtcOffset((Get-Date))

        Write-Host "`n=== NTP RECOVERY COMPLETE ===" -ForegroundColor Green
        Write-Host "Service Status: $($timeService.Status)" -ForegroundColor Green
        Write-Host "NTP Peers: $newPeers" -ForegroundColor Green
        Write-Host "Sync Mode: $(if($isDomain){'AllSync'}else{'NTP'})" -ForegroundColor Green
        Write-Host "Startup: Automatic" -ForegroundColor Green
        Write-Host "DST: $(if($isDST){'Active (summer time)'}else{'Inactive (winter time)'})" -ForegroundColor Green
        Write-Host "UTC Offset: +$($utcOffset.Hours)h" -ForegroundColor Green
        Write-Host ""

        return $true
    } catch {
        Write-Host "[FAIL] NTP recovery failed: $($_.Exception.Message)" -ForegroundColor Red
        Start-Service w32time -ErrorAction SilentlyContinue
        return $false
    }
}

# Function: Install and Configure SSH Server
function Install-SSHServer {
    try {
        Write-Host "`n=== INSTALLING SSH SERVER ===" -ForegroundColor Yellow

        # Check if already installed
        $sshCapability = Get-WindowsCapability -Online -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "OpenSSH.Server*" }
        if ($sshCapability -and $sshCapability.State -eq "Installed") {
            Write-Host " - OpenSSH Server is already installed" -ForegroundColor Green
        } else {
            Write-Host " - Installing OpenSSH Server..." -ForegroundColor Cyan
            Add-WindowsCapability -Online -Name OpenSSH.Server~~~~0.0.1.0 -ErrorAction Stop | Out-Null
            Write-Host " - OpenSSH Server installed successfully" -ForegroundColor Green
        }

        # Start once to generate config files
        Write-Host " - Starting sshd to generate configuration..." -ForegroundColor Cyan
        Start-Service sshd -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        Stop-Service sshd -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1

        # Enable password authentication
        Write-Host " - Enabling password authentication..." -ForegroundColor Cyan
        $configPath = "$env:ProgramData\ssh\sshd_config"
        if (Test-Path $configPath) {
            $config = Get-Content $configPath
            $config = $config -replace '#PasswordAuthentication yes','PasswordAuthentication yes' -replace 'PasswordAuthentication no','PasswordAuthentication yes'
            # Fix: Allow password auth for admin users (Windows blocks it by default via Match Group)
            $config = $config -replace '^Match Group administrators','#Match Group administrators'
            $config = $config -replace '^\s+AuthorizedKeysFile __PROGRAMDATA__','#       AuthorizedKeysFile __PROGRAMDATA__'
            $config | Set-Content $configPath
            Write-Host " - Password authentication enabled (including admin accounts)" -ForegroundColor Green
        } else {
            Write-Host " - Warning: sshd_config not found at $configPath" -ForegroundColor Yellow
        }

        # Start and enable service
        Write-Host " - Starting SSH service..." -ForegroundColor Cyan
        Start-Service sshd -ErrorAction Stop
        Set-Service sshd -StartupType Automatic
        Write-Host " - SSH service started and set to Automatic" -ForegroundColor Green

        # Configure firewall
        Write-Host " - Configuring firewall for SSH (TCP port 22)..." -ForegroundColor Cyan
        netsh advfirewall firewall delete rule name="SSH Server (XProtect)" 2>&1 | Out-Null
        netsh advfirewall firewall add rule name="SSH Server (XProtect)" dir=in action=allow protocol=TCP localport=22 2>&1 | Out-Null
        Write-Host " - Firewall rule created for SSH" -ForegroundColor Green

        Write-Host "`n=== SSH SERVER CONFIGURED ===" -ForegroundColor Green
        Write-Host "SSH is now accessible on port 22" -ForegroundColor Green
        Write-Host "Use any SSH client to connect with your Windows credentials" -ForegroundColor Cyan
        Write-Host ""

        return $true
    } catch {
        Write-Host "[FAIL] SSH Server installation failed: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function: Disable Copilot and Recall
function Disable-CopilotAndRecall {
    try {
        Write-Host " - Disabling Windows Copilot via Group Policy..." -ForegroundColor Cyan
        New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot" -Force | Out-Null
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot" -Name "TurnOffWindowsCopilot" -Value 1 -Type DWord
        Write-Host "[OK] Windows Copilot disabled" -ForegroundColor Green

        Write-Host " - Disabling Windows Recall via Group Policy..." -ForegroundColor Cyan
        New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsAI" -Force | Out-Null
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsAI" -Name "DisableAIDataAnalysis" -Value 1 -Type DWord
        Write-Host "[OK] Windows Recall disabled" -ForegroundColor Green

        return $true
    } catch {
        Write-Host "[FAIL] Error disabling Copilot/Recall: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function: Disable Telemetry Scheduled Tasks
function Disable-TelemetryTasks {
    try {
        Write-Host " - Disabling telemetry scheduled tasks..." -ForegroundColor Cyan

        $tasks = @(
            "Microsoft\Windows\Application Experience\Microsoft Compatibility Appraiser",
            "Microsoft\Windows\Customer Experience Improvement Program\Consolidator",
            "Microsoft\Windows\Customer Experience Improvement Program\UsbCeip",
            "Microsoft\Windows\DiskDiagnostic\Microsoft-Windows-DiskDiagnosticDataCollector"
        )

        $disabledCount = 0
        foreach ($task in $tasks) {
            try {
                $taskObj = Get-ScheduledTask -TaskPath "\$($task.Substring(0, $task.LastIndexOf('\') + 1))" -TaskName $task.Substring($task.LastIndexOf('\') + 1) -ErrorAction SilentlyContinue
                if ($taskObj) {
                    $taskObj | Disable-ScheduledTask -ErrorAction Stop | Out-Null
                    Write-Host "   [OK] Disabled: $task" -ForegroundColor Green
                    $disabledCount++
                } else {
                    Write-Host "   [SKIP] Not found: $task" -ForegroundColor Gray
                }
            } catch {
                Write-Host "   [WARN] Could not disable: $task" -ForegroundColor Yellow
            }
        }

        Write-Host "[OK] Telemetry tasks: $disabledCount disabled" -ForegroundColor Green
        return $true
    } catch {
        Write-Host "[FAIL] Error disabling telemetry tasks: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function: Set LargeSystemCache
function Set-LargeSystemCache {
    try {
        Write-Host " - Enabling LargeSystemCache for file server performance..." -ForegroundColor Cyan
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management" -Name "LargeSystemCache" -Value 1 -Type DWord
        Write-Host "[OK] LargeSystemCache enabled" -ForegroundColor Green
        return $true
    } catch {
        Write-Host "[FAIL] Error setting LargeSystemCache: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# BONUS: Standalone function to just enable/disable seconds in taskbar
function Set-TaskbarSecondsDisplay {
    param(
        [Parameter(Mandatory=$true)]
        [bool]$ShowSeconds
    )

    try {
        $registryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"

        if (-not (Test-Path $registryPath)) {
            New-Item -Path $registryPath -Force | Out-Null
        }

        $value = if ($ShowSeconds) { 1 } else { 0 }
        Set-ItemProperty -Path $registryPath -Name "ShowSecondsInSystemClock" -Value $value -Type DWord -ErrorAction Stop

        $status = if ($ShowSeconds) { "ENABLED" } else { "DISABLED" }
        Write-Host "Taskbar seconds display: $status" -ForegroundColor Green
        Write-Host "Restart Windows Explorer or reboot for changes to take effect." -ForegroundColor Yellow

        $restart = Read-Host "Restart Windows Explorer now? (Y/N)"
        if ($restart -match "^[Yy]") {
            Stop-Process -Name explorer -Force
            Start-Sleep -Seconds 2
            Start-Process explorer
            Write-Host "Windows Explorer restarted!" -ForegroundColor Green
        }

        return $true
    }
    catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Initialize global variables
$global:logContent = ""

$global:bloatwareApps = @(
    @{ Name = "Microsoft.3DBuilder"; Display = "3D Builder" },
    @{ Name = "Microsoft.WindowsAlarms"; Display = "Alarms and Clock" },
    @{ Name = "Microsoft.WindowsCommunicationsApps"; Display = "Mail and Calendar" },
    @{ Name = "Microsoft.GetHelp"; Display = "Get Help" },
    @{ Name = "Microsoft.Getstarted"; Display = "Microsoft Tips" },
    @{ Name = "Microsoft.SkypeApp"; Display = "Skype" },
    @{ Name = "Microsoft.ZuneMusic"; Display = "Groove Music" },
    @{ Name = "Microsoft.WindowsMaps"; Display = "Maps" },
    @{ Name = "Microsoft.MicrosoftSolitaireCollection"; Display = "Microsoft Solitaire Collection" },
    @{ Name = "Microsoft.BingFinance"; Display = "Money" },
    @{ Name = "Microsoft.ZuneVideo"; Display = "Movies and TV" },
    @{ Name = "Microsoft.BingNews"; Display = "News" },
    @{ Name = "Microsoft.Office.OneNote"; Display = "OneNote" },
    @{ Name = "Microsoft.People"; Display = "People" },
    @{ Name = "Microsoft.BingSports"; Display = "Sports" },
    @{ Name = "Microsoft.BingWeather"; Display = "Weather" },
    @{ Name = "Microsoft.MicrosoftStickyNotes"; Display = "Sticky Notes" },
    @{ Name = "Microsoft.OneConnect"; Display = "Mobile Plans" },
    @{ Name = "Microsoft.549981C3F5F10"; Display = "Cortana" },
    @{ Name = "Microsoft.WindowsFeedbackHub"; Display = "Feedback Hub" },
    @{ Name = "Microsoft.YourPhone"; Display = "Your Phone" },
    @{ Name = "Microsoft.MixedReality.Portal"; Display = "Mixed Reality Portal" },
    # 25H2 additions
    @{ Name = "Clipchamp.Clipchamp"; Display = "Clipchamp" },
    @{ Name = "Microsoft.Windows.DevHome"; Display = "Dev Home" },
    @{ Name = "MicrosoftCorporationII.MicrosoftFamily"; Display = "Microsoft Family" },
    @{ Name = "Microsoft.PowerAutomateDesktop"; Display = "Power Automate Desktop" },
    @{ Name = "Microsoft.BingSearch"; Display = "Bing Search" },
    @{ Name = "Disney.37853FC22B2CE"; Display = "Disney+" },
    @{ Name = "SpotifyAB.SpotifyMusic"; Display = "Spotify" },
    @{ Name = "BytedancePte.Ltd.TikTok"; Display = "TikTok" },
    @{ Name = "FACEBOOK.FACEBOOK"; Display = "Facebook" },
    @{ Name = "AmazonVideo.PrimeVideo"; Display = "Amazon Prime Video" },
    @{ Name = "5A894077.McAfeeSecurity"; Display = "McAfee Security" },
    @{ Name = "Microsoft.Windows.Ai.Copilot.Provider"; Display = "Copilot Provider" },
    @{ Name = "Microsoft.Copilot"; Display = "Microsoft Copilot" },
    @{ Name = "Microsoft.LinkedInForWindows"; Display = "LinkedIn" }
)

# Function: Display Banner
function Show-Banner {
    Clear-Host
    Write-Host "=== XProtect-Prepare v2.0 ===" -ForegroundColor Cyan
    Write-Host "By Ulf Holmstrom, Happy Problem Solver at Manvarg AB (2026)" -ForegroundColor Cyan
    Write-Host "DISCLAIMER: Developed privately to facilitate resellers. NOT supported by Milestone Systems A/S." -ForegroundColor Magenta
    Write-Host "For questions contact: ulf@manvarg.se" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Detected: $global:windowsVersion" -ForegroundColor White
    if ($global:ntpBugDetected) {
        Write-Host ""
        Write-Host "!!! WARNING: NTP v1.9x BUG DETECTED !!!" -ForegroundColor Red -BackgroundColor Yellow
        Write-Host "NTP peers are missing ,0x9 flags - time sync may be broken!" -ForegroundColor Red
        Write-Host "Use menu option 7 (NTP Recovery) to fix this." -ForegroundColor Red
    }
    Write-Host ""
}

# Function: Main Menu
function Show-MainMenu {
    Write-Host "=== MAIN MENU ===" -ForegroundColor Yellow
    Write-Host "1. Complete Setup (without SSH)"
    Write-Host "2. Complete Setup (with SSH activation)"
    Write-Host "3. Individual optimizations (submenu)"
    Write-Host "4. Configure NTP Time Server"
    Write-Host "5. Install & Configure SSH Server"
    Write-Host "6. Toggle Windows Update On/Off"
    Write-Host "7. NTP Recovery (fix v1.9x bug)"
    Write-Host "8. Exit"
    Write-Host ""
}

# Function: Individual Optimizations Submenu
function Show-IndividualMenu {
    Write-Host "`n=== INDIVIDUAL OPTIMIZATIONS ===" -ForegroundColor Yellow
    Write-Host "1. Antivirus and Storage Configuration"
    Write-Host "2. Windows Bloatware Removal"
    Write-Host "3. System Performance Optimization"
    Write-Host "4. Enable SNMP Capabilities (optional - for network monitoring)"
    Write-Host "5. Install Milestone PowerShell Tools"
    Write-Host "6. Disable Copilot and Recall"
    Write-Host "7. Disable Telemetry Tasks"
    Write-Host "8. Enable LargeSystemCache"
    Write-Host "9. Generate Log File"
    Write-Host "10. Return to Main Menu"
    Write-Host ""
}

# Function: Antivirus and Storage Configuration
function Invoke-AntivirusStorageConfig {
    Write-Host "`n=== ANTIVIRUS AND STORAGE CONFIGURATION ===" -ForegroundColor Yellow

    Write-Host "`nAvailable drives:" -ForegroundColor Cyan
    $drives = Get-AvailableDrives
    $drives | Format-Table -AutoSize

    Write-Host "Select drives for antivirus exceptions:" -ForegroundColor Green
    $driveInput = Read-Host "Enter drive letters (comma separated, e.g., d,e) or 'all' for all drives"

    if ($driveInput.ToLower() -eq 'all') {
        $selectedDrives = $drives.DeviceID
    } else {
        $selectedDrives = $driveInput -split "," | ForEach-Object { ($_.Trim().ToUpper() + ":").Replace("::", ":") }
    }

    Write-Host "`n--- Adding Antivirus Exceptions ---" -ForegroundColor Yellow
    foreach ($drive in $selectedDrives) {
        $drivePath = $drive + "\"
        try {
            Add-MpPreference -ExclusionPath $drivePath -ErrorAction Stop
            Write-Host "Antivirus exception added for drive ${drivePath}" -ForegroundColor Green
            $global:logContent += "Antivirus exception added for drive ${drivePath}`r`n"
        }
        catch {
            Write-Host "Error adding antivirus exception for drive ${drivePath}: $_" -ForegroundColor Red
            $global:logContent += "Error adding antivirus exception for drive ${drivePath}: $_`r`n"
        }
    }

    # Add Milestone installation folder exceptions on C: drive (added regardless of existence)
    Write-Host "`n--- Adding Milestone Installation Folder Exceptions (C: drive) ---" -ForegroundColor Yellow
    Write-Host "Following Milestone best practices - adding specific folder exceptions..." -ForegroundColor Cyan

    $milestoneFolders = @(
        "C:\Program Files\Milestone\",
        "C:\Program Files (x86)\Milestone\",
        "C:\ProgramData\Milestone\",
        "C:\ProgramData\VideoDeviceDrivers\",
        "C:\ProgramData\VideoOS\"
    )

    foreach ($folder in $milestoneFolders) {
        try {
            Add-MpPreference -ExclusionPath $folder -ErrorAction Stop
            Write-Host "Milestone folder exception added: ${folder}" -ForegroundColor Green
            $global:logContent += "Milestone folder exception added: ${folder}`r`n"
        }
        catch {
            Write-Host "Error adding exception for ${folder}: $_" -ForegroundColor Red
            $global:logContent += "Error adding exception for ${folder}: $_`r`n"
        }
    }

    # Add file extension exceptions
    Write-Host "`n--- Adding File Extension Exceptions ---" -ForegroundColor Yellow
    $exclusionExtensions = @("blk", "idx", "pic", "pqz", "sts", "ts")
    foreach ($ext in $exclusionExtensions) {
        try {
            Add-MpPreference -ExclusionExtension $ext -ErrorAction Stop
            Write-Host "Antivirus exception added for file extension: ${ext}" -ForegroundColor Green
            $global:logContent += "Antivirus exception added for file extension: ${ext}`r`n"
        }
        catch {
            Write-Host "Error adding antivirus exception for file extension ${ext}: $_" -ForegroundColor Red
            $global:logContent += "Error adding antivirus exception for file extension ${ext}: $_`r`n"
        }
    }

    $exclusionProcesses = @(
        "VideoOS.Recorder.Service.exe",
        "VideoOS.Server.Service.exe",
        "VideoOS.Administration.exe",
        "VideoOS.Event.Server.exe",
        "VideoOS.Failover.Service.exe",
        "VideoOS.MobileServer.Service.exe",
        "VideoOS.LPR.Server.exe"
    )
    foreach ($proc in $exclusionProcesses) {
        try {
            Add-MpPreference -ExclusionProcess $proc -ErrorAction Stop
            Write-Host "Antivirus exception added for process: ${proc}" -ForegroundColor Green
            $global:logContent += "Antivirus exception added for process: ${proc}`r`n"
        }
        catch {
            Write-Host "Error adding antivirus exception for process ${proc}: $_" -ForegroundColor Red
            $global:logContent += "Error adding antivirus exception for process ${proc}: $_`r`n"
        }
    }

    Write-Host "`n--- Storage Drive Configuration Check (Storage Drives Only) ---" -ForegroundColor Yellow
    foreach ($drive in $selectedDrives) {
        $driveForCheck = $drive.TrimEnd(':') + ":"
        $drivePath = $driveForCheck + "\"

        if ($driveForCheck.ToUpper() -eq "C:") {
            Write-Host "Skipping OS drive ${driveForCheck} - storage optimization not needed" -ForegroundColor Gray
            continue
        }

        Write-Host "`nChecking storage drive ${driveForCheck}..." -ForegroundColor Cyan

        $blockSize = Test-DriveBlockSize $drivePath
        if ($blockSize -eq 65536) {
            Write-Host "Storage Drive ${driveForCheck}: Correct block size - 64 KB." -ForegroundColor Green
            $global:logContent += "Storage Drive ${driveForCheck}: Correct block size - 64 KB.`r`n"
        } elseif ($blockSize) {
            $blockSizeKB = $blockSize / 1024
            Write-Host "Storage Drive ${driveForCheck}: Incorrect block size - ${blockSizeKB} KB. Expected 64 KB." -ForegroundColor Red
            Write-Host "NOTE: Drive needs to be reformatted with 64KB allocation unit size." -ForegroundColor Yellow
            $global:logContent += "Storage Drive ${driveForCheck}: Incorrect block size - ${blockSizeKB} KB.`r`n"
        } else {
            Write-Host "Storage Drive ${driveForCheck}: Could not determine block size." -ForegroundColor Yellow
            $global:logContent += "Storage Drive ${driveForCheck}: Block size check failed.`r`n"
        }

        $indexingEnabled = Test-DriveIndexing $drivePath
        if ($indexingEnabled -eq $false) {
            Write-Host "Storage Drive ${driveForCheck}: Indexing is OFF - correct for storage drives." -ForegroundColor Green
            $global:logContent += "Storage Drive ${driveForCheck}: Indexing is OFF.`r`n"
        } elseif ($indexingEnabled -eq $true) {
            Write-Host "Storage Drive ${driveForCheck}: Indexing is ON. Should be disabled for storage drives." -ForegroundColor Yellow
            $disableChoice = Read-Host "Disable indexing for storage drive ${driveForCheck}? (Y/N)"
            if ($disableChoice -match "^[Yy]") {
                if (Disable-DriveIndexing $drivePath) {
                    Write-Host "Storage Drive ${driveForCheck}: Indexing disabled." -ForegroundColor Green
                    $global:logContent += "Storage Drive ${driveForCheck}: Indexing disabled.`r`n"
                } else {
                    Write-Host "Storage Drive ${driveForCheck}: Failed to disable indexing." -ForegroundColor Red
                    $global:logContent += "Storage Drive ${driveForCheck}: Failed to disable indexing.`r`n"
                }
            }
        }
    }

    Read-Host "`nAntivirus and Storage configuration complete. Press Enter to return to menu..."
}

# Function: Windows Update Management
function Invoke-WindowsUpdateManagement {
    Write-Host "`n=== WINDOWS UPDATE MANAGEMENT ===" -ForegroundColor Yellow

    $currentStatus = Get-WindowsUpdateStatus
    if ($currentStatus -eq $true) {
        Write-Host "Current Status: Windows Updates are ENABLED" -ForegroundColor Green
        $global:logContent += "Windows Updates: Currently ENABLED`r`n"
    } elseif ($currentStatus -eq $false) {
        Write-Host "Current Status: Windows Updates are DISABLED" -ForegroundColor Red
        $global:logContent += "Windows Updates: Currently DISABLED`r`n"
    } else {
        Write-Host "Current Status: Unable to determine Windows Update status" -ForegroundColor Yellow
        $global:logContent += "Windows Updates: Status unknown`r`n"
    }

    Write-Host "`nOptions:" -ForegroundColor Cyan
    Write-Host "1. Enable Windows Updates"
    Write-Host "2. Disable Windows Updates"
    Write-Host "3. Return to Main Menu"

    $choice = Read-Host "`nEnter your choice (1-3)"
    switch ($choice) {
        "1" {
            if (Set-WindowsUpdateStatus -Enable $true) {
                Write-Host "Windows Updates have been ENABLED." -ForegroundColor Green
                $global:logContent += "Windows Updates: ENABLED by script`r`n"
            } else {
                Write-Host "Failed to enable Windows Updates." -ForegroundColor Red
            }
        }
        "2" {
            if (Set-WindowsUpdateStatus -Enable $false) {
                Write-Host "Windows Updates have been DISABLED." -ForegroundColor Red
                Write-Host "NOTE: Remember to manually check for critical security updates." -ForegroundColor Yellow
                $global:logContent += "Windows Updates: DISABLED by script`r`n"
            } else {
                Write-Host "Failed to disable Windows Updates." -ForegroundColor Red
            }
        }
        "3" {
            return
        }
    }

    Read-Host "`nPress Enter to return to main menu..."
}

# Function: Bloatware Removal
function Invoke-BloatwareRemoval {
    Write-Host "`n=== WINDOWS BLOATWARE REMOVAL ===" -ForegroundColor Yellow
    Write-Host "Remove unnecessary Windows apps for optimal performance." -ForegroundColor Cyan
    Write-Host ""

    Write-Host "Options:" -ForegroundColor Cyan
    Write-Host "1. Remove ALL bloatware apps automatically (excludes Xbox - see option 6)"
    Write-Host "2. Select specific apps to remove"
    Write-Host "3. Remove OneDrive only"
    Write-Host "4. Remove Microsoft Teams (all versions) + Outlook App + Office Hub"
    Write-Host "5. Full cleanup - All apps + OneDrive + Teams + Outlook + Office Hub"
    Write-Host "6. Remove Xbox Apps (WARNING: May cause issues - see details)"
    Write-Host "7. Return to Menu"

    $choice = Read-Host "`nEnter your choice (1-7)"

    switch ($choice) {
        "1" {
            Write-Host "`nRemoving ALL standard bloatware apps (Xbox apps excluded)..." -ForegroundColor Yellow
            $removedCount = 0
            foreach ($app in $global:bloatwareApps) {
                if (Remove-BloatwareApp -AppName $app.Name -FriendlyName $app.Display) {
                    $global:logContent += "Bloatware removed: $($app.Display)`r`n"
                    $removedCount++
                }
            }
            Write-Host "`nBloatware removal completed. $removedCount apps were removed." -ForegroundColor Green
            Write-Host "NOTE: Xbox apps were NOT removed (can cause screen recording issues)" -ForegroundColor Cyan
        }
        "2" {
            Write-Host "`nAvailable apps to remove:" -ForegroundColor Cyan
            for ($i = 0; $i -lt $global:bloatwareApps.Count; $i++) {
                Write-Host "$($i + 1). $($global:bloatwareApps[$i].Display)"
            }
            $appChoice = Read-Host "`nEnter app numbers to remove (comma separated, e.g., 1,3,5)"
            $appIndices = $appChoice -split "," | ForEach-Object { [int]$_.Trim() - 1 }

            $removedCount = 0
            foreach ($index in $appIndices) {
                if ($index -ge 0 -and $index -lt $global:bloatwareApps.Count) {
                    $app = $global:bloatwareApps[$index]
                    if (Remove-BloatwareApp -AppName $app.Name -FriendlyName $app.Display) {
                        $global:logContent += "Bloatware removed: $($app.Display)`r`n"
                        $removedCount++
                    }
                }
            }
            Write-Host "`nSelected apps removal completed. $removedCount apps were removed." -ForegroundColor Green
        }
        "3" {
            Write-Host "`nRemoving OneDrive..." -ForegroundColor Yellow
            if (Remove-OneDrive) {
                Write-Host "OneDrive removal completed." -ForegroundColor Green
                $global:logContent += "OneDrive: Successfully removed`r`n"
            }
        }
        "4" {
            Write-Host "`nRemoving Microsoft Teams, Outlook App, and Office Hub..." -ForegroundColor Yellow
            $teamsResult = Remove-MicrosoftTeams
            $outlookResult = Remove-OutlookApp
            $officeResult = Remove-OfficeHub

            Write-Host "`nRemoval completed:" -ForegroundColor Green
            Write-Host " - Teams: $(if($teamsResult){'Removed'}else{'Not found or already removed'})" -ForegroundColor $(if($teamsResult){'Green'}else{'Gray'})
            Write-Host " - Outlook App: $(if($outlookResult){'Removed'}else{'Not found or already removed'})" -ForegroundColor $(if($outlookResult){'Green'}else{'Gray'})
            Write-Host " - Office Hub: $(if($officeResult){'Removed'}else{'Not found or already removed'})" -ForegroundColor $(if($officeResult){'Green'}else{'Gray'})

            $global:logContent += "Microsoft Teams: $(if($teamsResult){'Removed'}else{'Not found'})`r`n"
            $global:logContent += "Outlook App: $(if($outlookResult){'Removed'}else{'Not found'})`r`n"
            $global:logContent += "Office Hub: $(if($officeResult){'Removed'}else{'Not found'})`r`n"
        }
        "5" {
            Write-Host "`nPerforming FULL CLEANUP (Xbox apps excluded)..." -ForegroundColor Yellow
            $confirm = Read-Host "This will remove ALL bloatware apps, OneDrive, Teams, Outlook, and Office Hub. Continue? (Y/N)"
            if ($confirm -match "^[Yy]") {
                $removedCount = 0
                foreach ($app in $global:bloatwareApps) {
                    if (Remove-BloatwareApp -AppName $app.Name -FriendlyName $app.Display) {
                        $removedCount++
                    }
                }
                Remove-OneDrive | Out-Null
                Remove-MicrosoftTeams | Out-Null
                Remove-OutlookApp | Out-Null
                Remove-OfficeHub | Out-Null

                Write-Host "`nFULL CLEANUP completed! $removedCount standard apps removed." -ForegroundColor Green
                Write-Host "Plus: OneDrive, Teams (all versions), Outlook App, Office Hub removed" -ForegroundColor Green
                Write-Host "NOTE: Xbox apps were NOT removed (can cause screen recording issues)" -ForegroundColor Cyan
                $global:logContent += "Full cleanup completed: $removedCount apps, OneDrive, Teams, Outlook, Office Hub removed`r`n"
            }
        }
        "6" {
            Remove-XboxApps
        }
        "7" {
            return
        }
    }

    Read-Host "`nPress Enter to return to menu..."
}

# Function: System Performance Optimization
function Invoke-SystemOptimization {
    Write-Host "`n=== SYSTEM PERFORMANCE OPTIMIZATION ===" -ForegroundColor Yellow
    Write-Host "Optimize Windows for Milestone XProtect performance." -ForegroundColor Cyan
    Write-Host ""

    Write-Host "Options:" -ForegroundColor Cyan
    Write-Host "1. Configure Power Management (Never Sleep Mode - Critical for 24/7 recording)"
    Write-Host "2. Optimize Windows Services"
    Write-Host "3. Enable and Optimize Remote Desktop"
    Write-Host "4. FULL OPTIMIZATION - All above"
    Write-Host "5. Return to Menu"

    $choice = Read-Host "`nEnter your choice (1-5)"

    if ($choice -match "^[1-4]$") {
        Write-Host "`nCreating registry backup..." -ForegroundColor Cyan
        New-RegistryBackup -BackupName "SystemOptimization"
    }

    switch ($choice) {
        "1" {
            Write-Host "`nConfiguring Power Management..." -ForegroundColor Yellow
            Write-Host "Setting Never Sleep Mode - CRITICAL for 24/7 video recording" -ForegroundColor Cyan
            if (Set-PowerManagement -HighPerformance $true) {
                $global:logContent += "Power Management: Never Sleep Mode activated for 24/7 recording`r`n"
            }
        }
        "2" {
            Write-Host "`nOptimizing Windows Services..." -ForegroundColor Yellow
            $disabledCount = Set-WindowsServices -OptimizeForMilestone $true
            Write-Host "Service optimization completed. $disabledCount services disabled." -ForegroundColor Green
            $global:logContent += "Windows Services: $disabledCount services disabled`r`n"
        }
        "3" {
            Write-Host "`nEnabling and Optimizing Remote Desktop..." -ForegroundColor Yellow
            if (Set-RemoteDesktopOptimization -OptimizeRDP $true) {
                $global:logContent += "Remote Desktop: ENABLED and optimized`r`n"
            }
        }
        "4" {
            Write-Host "`nPERFORMING FULL OPTIMIZATION..." -ForegroundColor Yellow
            $confirm = Read-Host "This will optimize power, services, and enable RDP. Continue? (Y/N)"
            if ($confirm -match "^[Yy]") {
                $powerResult = Set-PowerManagement -HighPerformance $true
                $servicesDisabled = Set-WindowsServices -OptimizeForMilestone $true
                $rdpResult = Set-RemoteDesktopOptimization -OptimizeRDP $true

                Write-Host "`nFULL OPTIMIZATION COMPLETED!" -ForegroundColor Green
                Write-Host " - Power Management: $(if($powerResult){'Never Sleep Active (Critical for 24/7 recording)'}else{'Failed'})" -ForegroundColor Green
                Write-Host " - Services Disabled: $servicesDisabled" -ForegroundColor Green
                Write-Host " - Remote Desktop: $(if($rdpResult){'Enabled and Optimized'}else{'Failed'})" -ForegroundColor Green

                $global:logContent += "FULL OPTIMIZATION: Power $(if($powerResult){'optimized'}else{'failed'}), $servicesDisabled services disabled, RDP $(if($rdpResult){'enabled'}else{'failed'})`r`n"
            }
        }
        "5" {
            return
        }
    }

    Write-Host "`nTIP: A system restart is recommended to apply all changes." -ForegroundColor Yellow
    Read-Host "`nPress Enter to return to menu..."
}

# Function: SNMP Configuration
function Invoke-SNMPConfiguration {
    Write-Host "`n=== SNMP CAPABILITIES INSTALLATION ===" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "What is SNMP?" -ForegroundColor Cyan
    Write-Host "SNMP (Simple Network Management Protocol) allows network monitoring" -ForegroundColor White
    Write-Host "systems to collect performance data from your servers." -ForegroundColor White
    Write-Host ""
    Write-Host "Do you need SNMP?" -ForegroundColor Cyan
    Write-Host "Only install if you have network monitoring tools like:" -ForegroundColor White
    Write-Host " - PRTG, Nagios, Zabbix, or similar monitoring systems" -ForegroundColor White
    Write-Host " - Enterprise management platforms that use SNMP" -ForegroundColor White
    Write-Host ""
    Write-Host "If you're unsure, you probably DON'T need it for basic XProtect operation." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "NOTE: Windows Update must be enabled for this to work!" -ForegroundColor Red
    Write-Host ""

    $updateStatus = Get-WindowsUpdateStatus
    if ($updateStatus -eq $false) {
        Write-Host "WARNING: Windows Update is currently DISABLED!" -ForegroundColor Red
        Write-Host "SNMP installation requires Windows Update to be enabled." -ForegroundColor Yellow
        $enableUpdate = Read-Host "Enable Windows Update temporarily? (Y/N)"
        if ($enableUpdate -match "^[Yy]") {
            Set-WindowsUpdateStatus -Enable $true
            Write-Host "Windows Update enabled. Proceeding with SNMP installation..." -ForegroundColor Green
        } else {
            Write-Host "SNMP installation cancelled. Please enable Windows Update first." -ForegroundColor Yellow
            Read-Host "`nPress Enter to return to menu..."
            return
        }
    }

    $confirm = Read-Host "Install SNMP capabilities? (Y/N)"
    if ($confirm -match "^[Yy]") {
        if (Enable-SNMP) {
            Write-Host "SNMP installation completed successfully." -ForegroundColor Green
            $global:logContent += "SNMP: Client and WMI Provider installed successfully`r`n"
        } else {
            Write-Host "SNMP installation failed." -ForegroundColor Red
            $global:logContent += "SNMP: Installation failed`r`n"
        }
    } else {
        Write-Host "SNMP installation cancelled." -ForegroundColor Yellow
    }

    Read-Host "`nPress Enter to return to menu..."
}

# Function: Milestone PowerShell Tools Installation
function Invoke-MilestonePSToolsInstallation {
    Write-Host "`n=== MILESTONE POWERSHELL TOOLS INSTALLATION ===" -ForegroundColor Yellow
    Write-Host "Install the Milestone PowerShell Tools for advanced XProtect management." -ForegroundColor Cyan
    Write-Host "These tools provide powerful cmdlets for automating XProtect tasks." -ForegroundColor Cyan
    Write-Host ""

    $confirm = Read-Host "Install Milestone PowerShell Tools? (Y/N)"
    if ($confirm -match "^[Yy]") {
        if (Install-MilestonePSTools) {
            Write-Host "Milestone PowerShell Tools installation completed successfully." -ForegroundColor Green
            $global:logContent += "Milestone PowerShell Tools: Successfully installed from PowerShell Gallery`r`n"
        } else {
            Write-Host "Milestone PowerShell Tools installation failed." -ForegroundColor Red
            $global:logContent += "Milestone PowerShell Tools: Installation failed`r`n"
        }
    } else {
        Write-Host "Milestone PowerShell Tools installation cancelled." -ForegroundColor Yellow
    }

    Read-Host "`nPress Enter to return to menu..."
}

# Function: NTP Server Configuration (standalone menu item)
function Invoke-NTPServerConfiguration {
    Write-Host "`n=== NTP TIME SERVER FOR CAMERAS ===" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Why do you need this?" -ForegroundColor Cyan
    Write-Host "CCTV cameras MUST have accurate time synchronization for:" -ForegroundColor White
    Write-Host " - Forensically valid video timestamps" -ForegroundColor White
    Write-Host " - Synchronized recording across multiple cameras" -ForegroundColor White
    Write-Host " - Accurate event correlation in XProtect" -ForegroundColor White
    Write-Host " - Legal compliance (timestamps must be accurate in court)" -ForegroundColor White
    Write-Host ""
    Write-Host "This configuration will:" -ForegroundColor Cyan
    Write-Host " - Configure Windows Time service as a reliable NTP server" -ForegroundColor White
    Write-Host " - Sync the server time from Swedish NTP pool (se.pool.ntp.org)" -ForegroundColor White
    Write-Host " - Open firewall UDP port 123 for camera connections" -ForegroundColor White
    Write-Host " - Enable your cameras to sync time from this server" -ForegroundColor White
    Write-Host ""
    Write-Host "IMPORTANT: After enabling this, configure your cameras to use" -ForegroundColor Yellow
    Write-Host "this server's IP address as their NTP time server." -ForegroundColor Yellow
    Write-Host ""

    $confirm = Read-Host "Configure NTP Time Server for cameras? (Y/N)"
    if ($confirm -match "^[Yy]") {
        if (Enable-NTPServer) {
            Write-Host "`nNTP Time Server configuration completed successfully." -ForegroundColor Green
            $global:logContent += "NTP Server: Configured successfully for camera time synchronization`r`n"
        } else {
            Write-Host "`nNTP Time Server configuration encountered issues." -ForegroundColor Red
            $global:logContent += "NTP Server: Configuration failed or incomplete`r`n"
        }
    } else {
        Write-Host "NTP Server configuration cancelled." -ForegroundColor Yellow
    }

    Read-Host "`nPress Enter to return to main menu..."
}

# Function: Complete System Setup
function Invoke-CompleteSystemSetup {
    param(
        [bool]$IncludeSSH = $false
    )

    $sshLabel = if ($IncludeSSH) { "WITH SSH" } else { "WITHOUT SSH" }

    Write-Host "`n=== COMPLETE MILESTONE XPROTECT SYSTEM SETUP ($sshLabel) ===" -ForegroundColor Yellow
    Write-Host "This will perform ALL essential optimizations for Milestone XProtect:" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Execution Order:" -ForegroundColor White
    Write-Host " 1. Registry backup" -ForegroundColor Cyan
    Write-Host " 2. Performance optimizations (power, services, indexing)" -ForegroundColor Cyan
    Write-Host " 3. RDP configuration" -ForegroundColor Cyan
    Write-Host " 4. Bloatware removal (existing + 25H2 apps - Xbox excluded)" -ForegroundColor Cyan
    Write-Host " 5. Copilot/Recall disable" -ForegroundColor Cyan
    Write-Host " 6. Telemetry tasks disable" -ForegroundColor Cyan
    Write-Host " 7. LargeSystemCache optimization" -ForegroundColor Cyan
    Write-Host " 8. OneDrive removal" -ForegroundColor Cyan
    Write-Host " 9. Teams/Outlook/Office Hub removal" -ForegroundColor Cyan
    Write-Host "10. AV exclusions for Milestone" -ForegroundColor Cyan
    Write-Host "11. NTP Time Server configuration (FIXED v2.0)" -ForegroundColor Cyan
    if ($IncludeSSH) {
        Write-Host "12. SSH Server installation and configuration" -ForegroundColor Cyan
    }
    Write-Host ""
    Write-Host "NOTE: Windows Update toggle is NOT included - use menu option 6 separately." -ForegroundColor Yellow
    Write-Host "NOTE: Xbox apps will NOT be removed (prevents screen recording issues)" -ForegroundColor Yellow
    Write-Host "WARNING: This will make significant system changes!" -ForegroundColor Red
    Write-Host ""

    $confirm = Read-Host "Proceed with COMPLETE system setup? (Y/N)"
    if ($confirm -notmatch "^[Yy]") {
        Write-Host "Complete system setup cancelled." -ForegroundColor Yellow
        Read-Host "Press Enter to return to main menu..."
        return
    }

    Write-Host "`n=== STARTING COMPLETE SYSTEM CONFIGURATION ===" -ForegroundColor Green
    $startTime = Get-Date

    # Phase 1: Registry Backup
    Write-Host "`n--- PHASE 1: REGISTRY BACKUP ---" -ForegroundColor Yellow
    New-RegistryBackup -BackupName "CompleteSetup_v2"

    # Phase 2: Performance Optimizations
    Write-Host "`n--- PHASE 2: PERFORMANCE OPTIMIZATIONS ---" -ForegroundColor Yellow
    $powerResult = Set-PowerManagement -HighPerformance $true
    $servicesDisabled = Set-WindowsServices -OptimizeForMilestone $true
    Write-Host "[OK] Performance optimization completed" -ForegroundColor Green
    $global:logContent += "COMPLETE SETUP - Power: $(if($powerResult){'optimized'}else{'failed'}), $servicesDisabled services disabled`r`n"

    # Phase 3: RDP Configuration
    Write-Host "`n--- PHASE 3: RDP CONFIGURATION ---" -ForegroundColor Yellow
    $rdpResult = Set-RemoteDesktopOptimization -OptimizeRDP $true
    $global:logContent += "COMPLETE SETUP - RDP: $(if($rdpResult){'enabled'}else{'failed'})`r`n"

    # Phase 4: Bloatware Removal
    Write-Host "`n--- PHASE 4: BLOATWARE REMOVAL ---" -ForegroundColor Yellow
    Write-Host "Removing standard bloatware including 25H2 apps (Xbox apps excluded)..." -ForegroundColor Cyan
    $removedCount = 0
    foreach ($app in $global:bloatwareApps) {
        if (Remove-BloatwareApp -AppName $app.Name -FriendlyName $app.Display) {
            $removedCount++
        }
    }
    Write-Host "[OK] Bloatware removal: $removedCount apps removed" -ForegroundColor Green
    $global:logContent += "COMPLETE SETUP - Bloatware: $removedCount apps removed (Xbox preserved)`r`n"

    # Phase 5: Copilot/Recall Disable
    Write-Host "`n--- PHASE 5: COPILOT/RECALL DISABLE ---" -ForegroundColor Yellow
    $copilotResult = Disable-CopilotAndRecall
    $global:logContent += "COMPLETE SETUP - Copilot/Recall: $(if($copilotResult){'disabled'}else{'failed'})`r`n"

    # Phase 6: Telemetry Tasks Disable
    Write-Host "`n--- PHASE 6: TELEMETRY TASKS DISABLE ---" -ForegroundColor Yellow
    $telemetryResult = Disable-TelemetryTasks
    $global:logContent += "COMPLETE SETUP - Telemetry tasks: $(if($telemetryResult){'disabled'}else{'failed'})`r`n"

    # Phase 7: LargeSystemCache
    Write-Host "`n--- PHASE 7: LARGESYSTEMCACHE ---" -ForegroundColor Yellow
    $cacheResult = Set-LargeSystemCache
    $global:logContent += "COMPLETE SETUP - LargeSystemCache: $(if($cacheResult){'enabled'}else{'failed'})`r`n"

    # Phase 8: OneDrive Removal
    Write-Host "`n--- PHASE 8: ONEDRIVE REMOVAL ---" -ForegroundColor Yellow
    Remove-OneDrive | Out-Null
    $global:logContent += "COMPLETE SETUP - OneDrive: removed`r`n"

    # Phase 9: Teams/Outlook/Office Hub Removal
    Write-Host "`n--- PHASE 9: TEAMS/OUTLOOK/OFFICE HUB REMOVAL ---" -ForegroundColor Yellow
    Remove-MicrosoftTeams | Out-Null
    Remove-OutlookApp | Out-Null
    Remove-OfficeHub | Out-Null
    Write-Host "[OK] Teams, Outlook, Office Hub removal completed" -ForegroundColor Green
    $global:logContent += "COMPLETE SETUP - Teams/Outlook/OfficeHub: removed`r`n"

    # Phase 10: AV Exclusions
    Write-Host "`n--- PHASE 10: ANTIVIRUS EXCLUSIONS ---" -ForegroundColor Yellow
    Write-Host "Select storage drives for antivirus exceptions (C: excluded for security):" -ForegroundColor Green

    $drives = Get-AvailableDrives
    Write-Host "Available drives:" -ForegroundColor Cyan
    $drives | Where-Object { $_.DeviceID -ne "C:" } | Format-Table -AutoSize

    $driveInput = Read-Host "Enter drive letters (comma separated, e.g., d,e) or 'storage' for all non-OS drives"

    if ($driveInput.ToLower() -eq 'storage') {
        $selectedDrives = $drives | Where-Object { $_.DeviceID -ne "C:" } | Select-Object -ExpandProperty DeviceID
        Write-Host "Selected storage drives: $($selectedDrives -join ', ')" -ForegroundColor Cyan
    } else {
        $selectedDrives = $driveInput -split "," | ForEach-Object {
            $drive = ($_.Trim().ToUpper() + ":").Replace("::", ":")
            if ($drive -eq "C:") {
                Write-Host "WARNING: C: drive excluded for security" -ForegroundColor Red
                return $null
            }
            return $drive
        } | Where-Object { $_ -ne $null }

        if ($selectedDrives.Count -eq 0) {
            Write-Host "No valid storage drives selected. Skipping drive antivirus exceptions." -ForegroundColor Yellow
            $selectedDrives = @()
        }
    }

    # Add storage drive exclusions
    foreach ($drive in $selectedDrives) {
        $drivePath = $drive + "\"
        try {
            Add-MpPreference -ExclusionPath $drivePath -ErrorAction Stop
            Write-Host "[OK] Antivirus exception added for storage drive ${drivePath}" -ForegroundColor Green
            $global:logContent += "COMPLETE SETUP - AV exception (STORAGE): ${drivePath}`r`n"
        }
        catch {
            Write-Host "[FAIL] Error adding exception for ${drivePath}" -ForegroundColor Red
        }
    }

    # Add Milestone folder exceptions (regardless of existence)
    Write-Host "`nAdding Milestone installation folder exceptions..." -ForegroundColor Cyan
    $milestoneFolders = @(
        "C:\Program Files\Milestone\",
        "C:\Program Files (x86)\Milestone\",
        "C:\ProgramData\Milestone\",
        "C:\ProgramData\VideoDeviceDrivers\",
        "C:\ProgramData\VideoOS\"
    )

    foreach ($folder in $milestoneFolders) {
        try {
            Add-MpPreference -ExclusionPath $folder -ErrorAction Stop
            Write-Host "[OK] Milestone folder exception: ${folder}" -ForegroundColor Green
            $global:logContent += "COMPLETE SETUP - Milestone folder exception: ${folder}`r`n"
        }
        catch {
            Write-Host "[FAIL] Error adding exception for ${folder}" -ForegroundColor Red
        }
    }

    # File extension exceptions
    Write-Host "`nAdding file extension exceptions..." -ForegroundColor Cyan
    $exclusionExtensions = @("blk", "idx", "pic", "pqz", "sts", "ts")
    foreach ($ext in $exclusionExtensions) {
        try {
            Add-MpPreference -ExclusionExtension $ext -ErrorAction Stop
            Write-Host "[OK] Extension exception: ${ext}" -ForegroundColor Green
            $global:logContent += "COMPLETE SETUP - Extension exception: ${ext}`r`n"
        }
        catch {
            Write-Host "[FAIL] Error adding extension exception: ${ext}" -ForegroundColor Red
        }
    }

    # Process exceptions
    $exclusionProcesses = @(
        "VideoOS.Recorder.Service.exe",
        "VideoOS.Server.Service.exe",
        "VideoOS.Administration.exe",
        "VideoOS.Event.Server.exe",
        "VideoOS.Failover.Service.exe",
        "VideoOS.MobileServer.Service.exe",
        "VideoOS.LPR.Server.exe"
    )
    foreach ($proc in $exclusionProcesses) {
        try {
            Add-MpPreference -ExclusionProcess $proc -ErrorAction Stop
            Write-Host "[OK] Process exception: ${proc}" -ForegroundColor Green
            $global:logContent += "COMPLETE SETUP - Process exception: ${proc}`r`n"
        }
        catch {
            Write-Host "[FAIL] Error adding process exception: ${proc}" -ForegroundColor Red
        }
    }

    # Storage drive optimization
    $storageDrives = $selectedDrives | Where-Object { $_ -ne "C:" }
    if ($storageDrives -and $storageDrives.Count -gt 0) {
        Write-Host "`nOptimizing selected storage drives..." -ForegroundColor Cyan
        foreach ($drive in $storageDrives) {
            $drivePath = $drive + "\"

            $blockSize = Test-DriveBlockSize $drivePath
            if ($blockSize -eq 65536) {
                Write-Host "[OK] ${drive}: Correct 64KB block size" -ForegroundColor Green
            } else {
                Write-Host "[WARN] ${drive}: Incorrect block size - should be 64KB for optimal video recording" -ForegroundColor Yellow
            }

            $indexingEnabled = Test-DriveIndexing $drivePath
            if ($indexingEnabled -eq $true) {
                if (Disable-DriveIndexing $drivePath) {
                    Write-Host "[OK] ${drive}: Indexing disabled" -ForegroundColor Green
                    $global:logContent += "COMPLETE SETUP - ${drive}: Indexing disabled`r`n"
                }
            }
        }
    }

    # Phase 11: NTP Time Server
    Write-Host "`n--- PHASE 11: NTP TIME SERVER (FIXED v2.0) ---" -ForegroundColor Yellow
    Write-Host "Configuring NTP server with correct peer format and domain detection..." -ForegroundColor Cyan
    $ntpResult = Enable-NTPServer
    if ($ntpResult) {
        Write-Host "[OK] NTP Time Server configured successfully" -ForegroundColor Green
        $global:logContent += "COMPLETE SETUP - NTP Server: Configured (v2.0 fixed format)`r`n"
    } else {
        Write-Host "[WARN] NTP Time Server configuration had issues" -ForegroundColor Yellow
        $global:logContent += "COMPLETE SETUP - NTP Server: Configuration failed`r`n"
    }

    # Phase 12: SSH (only if requested)
    $sshResult = $false
    if ($IncludeSSH) {
        Write-Host "`n--- PHASE 12: SSH SERVER ---" -ForegroundColor Yellow
        $sshResult = Install-SSHServer
        $global:logContent += "COMPLETE SETUP - SSH Server: $(if($sshResult){'installed'}else{'failed'})`r`n"
    }

    # Final Summary
    $endTime = Get-Date
    $duration = $endTime - $startTime

    Write-Host "`n=== COMPLETE SYSTEM SETUP FINISHED ===" -ForegroundColor Green
    Write-Host "Setup completed in $($duration.Minutes) minutes and $($duration.Seconds) seconds" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "SUMMARY:" -ForegroundColor White
    Write-Host " - Power: Never Sleep Mode activated (CRITICAL for 24/7 recording)" -ForegroundColor Green
    Write-Host " - Services: $servicesDisabled services disabled" -ForegroundColor Green
    Write-Host " - RDP: $(if($rdpResult){'Enabled and configured'}else{'Failed'})" -ForegroundColor Green
    Write-Host " - Bloatware: $removedCount apps removed (Xbox preserved)" -ForegroundColor Green
    Write-Host " - Copilot/Recall: $(if($copilotResult){'Disabled'}else{'Failed'})" -ForegroundColor Green
    Write-Host " - Telemetry tasks: $(if($telemetryResult){'Disabled'}else{'Failed'})" -ForegroundColor Green
    Write-Host " - LargeSystemCache: $(if($cacheResult){'Enabled'}else{'Failed'})" -ForegroundColor Green
    Write-Host " - OneDrive/Teams/Outlook/OfficeHub: Removed" -ForegroundColor Green
    Write-Host " - AV Exclusions: Milestone folders + extensions + processes" -ForegroundColor Green
    Write-Host " - Storage: $(if($storageDrives -and $storageDrives.Count -gt 0){"$($storageDrives.Count) drives optimized"}else{"No storage drives selected"})" -ForegroundColor Green
    Write-Host " - NTP Server: $(if($ntpResult){'CONFIGURED (v2.0 fixed format)'}else{'Had issues'})" -ForegroundColor $(if($ntpResult){'Green'}else{'Yellow'})
    if ($IncludeSSH) {
        Write-Host " - SSH Server: $(if($sshResult){'Installed and configured'}else{'Failed'})" -ForegroundColor $(if($sshResult){'Green'}else{'Yellow'})
    }
    Write-Host " - Windows Update: NOT changed (use menu option 6)" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "NEXT STEPS:" -ForegroundColor Yellow
    Write-Host " 1. RESTART the system to apply all changes" -ForegroundColor Yellow
    Write-Host " 2. Toggle Windows Update on/off as needed (menu option 6)" -ForegroundColor Yellow
    Write-Host " 3. Install Milestone XProtect software" -ForegroundColor Yellow
    Write-Host " 4. Configure cameras to use this server's IP as NTP server" -ForegroundColor Yellow
    Write-Host " 5. Configure camera connections and storage" -ForegroundColor Yellow
    Write-Host ""

    # Auto-generate log
    $logFolderPath = "C:\Milestonecheck-ulfh"
    if (-not (Test-Path $logFolderPath)) {
        New-Item -ItemType Directory -Path $logFolderPath | Out-Null
    }
    $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $logFilePath = Join-Path $logFolderPath "CompleteSetupLog_v2_$timestamp.txt"
    $header = "Milestone XProtect COMPLETE SETUP Log - $(Get-Date)`r`nXProtect-Prepare v2.0`r`nBy Ulf Holmstrom, Happy Problem Solver at Manvarg AB (2026)`r`nFor questions contact: ulf@manvarg.se`r`n===========================================`r`n"
    $fullLogContent = $header + $global:logContent
    $fullLogContent | Out-File -FilePath $logFilePath -Encoding UTF8
    Write-Host "[OK] Complete setup log saved: $logFilePath" -ForegroundColor Green

    Play-Fanfare

    Read-Host "`nPress Enter to return to main menu..."
}

# Function: Generate Log File
function Invoke-LogGeneration {
    Write-Host "`n=== LOG FILE GENERATION ===" -ForegroundColor Yellow

    if ($global:logContent -eq "") {
        Write-Host "No configuration changes have been made yet. Nothing to log." -ForegroundColor Yellow
        Read-Host "Press Enter to return to menu..."
        return
    }

    $writeLog = Read-Host "Save a log file with configuration details? (Y/N)"
    if ($writeLog -match "^[Yy]") {
        $logFolderPath = "C:\Milestonecheck-ulfh"
        if (-not (Test-Path $logFolderPath)) {
            New-Item -ItemType Directory -Path $logFolderPath | Out-Null
        }
        $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
        $logFilePath = Join-Path $logFolderPath "MilestoneXProtectConfigLog_$timestamp.txt"
        $header = "Milestone XProtect Configuration Log - $(Get-Date)`r`nXProtect-Prepare v2.0`r`nBy Ulf Holmstrom, Happy Problem Solver at Manvarg AB (2026)`r`nFor questions contact: ulf@manvarg.se`r`n===========================================`r`n"
        $fullLogContent = $header + $global:logContent
        $fullLogContent | Out-File -FilePath $logFilePath -Encoding UTF8
        Write-Host "Log file saved at: $logFilePath" -ForegroundColor Green
        Play-Fanfare
    } else {
        Write-Host "No log file will be created." -ForegroundColor Yellow
    }

    Read-Host "Press Enter to return to menu..."
}

# Main Script Execution
Show-Banner
Write-Host "XProtect-Prepare v2.0 - Complete system preparation for Milestone XProtect"
Write-Host ""
Write-Host "This script will help you configure:" -ForegroundColor Cyan
Write-Host "  - Antivirus exceptions and storage drive optimization"
Write-Host "  - Bloatware removal (including 25H2 apps, Copilot, TikTok, etc.)"
Write-Host "  - System performance optimization (services, power, RDP)"
Write-Host "  - Copilot/Recall/Telemetry disable"
Write-Host "  - NTP Time Server for camera synchronization (FIXED in v2.0)"
Write-Host "  - SSH Server installation and configuration"
Write-Host "  - Windows Update management (standalone toggle)"
Write-Host ""
Write-Host "v2.0 HIGHLIGHTS:" -ForegroundColor Cyan
Write-Host "  - CRITICAL FIX: NTP peer format corrected (,0x9 flags)" -ForegroundColor Green
Write-Host "  - NEW: NTP Recovery for machines with v1.9x bug" -ForegroundColor Green
Write-Host "  - NEW: SSH Server install option" -ForegroundColor Green
Write-Host "  - NEW: Copilot, Recall, Telemetry disabled" -ForegroundColor Green
Write-Host "  - NEW: 25H2 bloatware (Clipchamp, TikTok, Copilot, etc.)" -ForegroundColor Green
Write-Host "  - NEW: LargeSystemCache optimization" -ForegroundColor Green
Write-Host "  - IMPROVED: Windows Update moved out of Complete Setup" -ForegroundColor Green
Write-Host ""
Read-Host "Press Enter to continue..."

do {
    Show-Banner
    Show-MainMenu
    $choice = Read-Host "Enter your choice (1-8)"

    switch ($choice) {
        "1" { Invoke-CompleteSystemSetup -IncludeSSH $false }
        "2" { Invoke-CompleteSystemSetup -IncludeSSH $true }
        "3" {
            do {
                Show-Banner
                Show-IndividualMenu
                $subChoice = Read-Host "Enter your choice (1-10)"

                switch ($subChoice) {
                    "1" { Invoke-AntivirusStorageConfig }
                    "2" { Invoke-BloatwareRemoval }
                    "3" { Invoke-SystemOptimization }
                    "4" { Invoke-SNMPConfiguration }
                    "5" { Invoke-MilestonePSToolsInstallation }
                    "6" {
                        Write-Host "`n=== DISABLE COPILOT AND RECALL ===" -ForegroundColor Yellow
                        $confirm = Read-Host "Disable Windows Copilot and Recall? (Y/N)"
                        if ($confirm -match "^[Yy]") {
                            Disable-CopilotAndRecall
                            $global:logContent += "Copilot/Recall: Disabled`r`n"
                        }
                        Read-Host "`nPress Enter to return to menu..."
                    }
                    "7" {
                        Write-Host "`n=== DISABLE TELEMETRY TASKS ===" -ForegroundColor Yellow
                        $confirm = Read-Host "Disable telemetry scheduled tasks? (Y/N)"
                        if ($confirm -match "^[Yy]") {
                            Disable-TelemetryTasks
                            $global:logContent += "Telemetry tasks: Disabled`r`n"
                        }
                        Read-Host "`nPress Enter to return to menu..."
                    }
                    "8" {
                        Write-Host "`n=== LARGESYSTEMCACHE ===" -ForegroundColor Yellow
                        $confirm = Read-Host "Enable LargeSystemCache for file server performance? (Y/N)"
                        if ($confirm -match "^[Yy]") {
                            Set-LargeSystemCache
                            $global:logContent += "LargeSystemCache: Enabled`r`n"
                        }
                        Read-Host "`nPress Enter to return to menu..."
                    }
                    "9" { Invoke-LogGeneration }
                    "10" { break }
                    default {
                        Write-Host "Invalid choice. Please select 1-10." -ForegroundColor Red
                        Start-Sleep -Seconds 2
                    }
                }
            } while ($subChoice -ne "10")
        }
        "4" { Invoke-NTPServerConfiguration }
        "5" {
            Write-Host "`n=== SSH SERVER INSTALLATION ===" -ForegroundColor Yellow
            $confirm = Read-Host "Install and configure SSH Server? (Y/N)"
            if ($confirm -match "^[Yy]") {
                Install-SSHServer
                $global:logContent += "SSH Server: Installed and configured`r`n"
            }
            Read-Host "`nPress Enter to return to main menu..."
        }
        "6" { Invoke-WindowsUpdateManagement }
        "7" {
            Invoke-NTPRecovery
            Read-Host "`nPress Enter to return to main menu..."
        }
        "8" {
            Write-Host "Exiting script. Thank you for using XProtect-Prepare v2.0!" -ForegroundColor Green
            Write-Host "For questions or feedback, contact: ulf@manvarg.se" -ForegroundColor Cyan
            break
        }
        default {
            Write-Host "Invalid choice. Please select 1-8." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
} while ($choice -ne "8")
