<#
   Enhanced Script for Milestone XProtect System Configuration
   By Ulf Holmstrom, Happy Problem Solver at Manvarg AB (Rev 2.0, 2025)
   For questions contact: ulf@manvarg.se

   IMPORTANT DISCLAIMER:
   This script has been developed by Ulf Holmstrom, Happy Problem Solver at Manvarg AB, to facilitate resellers.
   It is provided as a public resource and is NOT supported by Milestone Systems.

   CHANGELOG Rev 2.0:
   - RESTRUCTURED: Better menu separation - each function is now in its own logical category
   - IMPROVED: Antivirus exceptions separated from storage optimization
   - IMPROVED: Remote Desktop moved to its own "Remote Access & Network" category
   - IMPROVED: NTP and SNMP grouped together under network services
   - ADDED: Region selection for NTP servers (Nordic, Europe, Global)
   - FIXED: Consistent C: drive exclusion from antivirus exceptions (security)
   - FIXED: Removed obsolete services (HomeGroup) that don't exist in Windows 10/11
   - ADDED: Restore from backup option
   - IMPROVED: Complete setup now asks about all optional services consistently
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
    return Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 } | Select-Object DeviceID, @{Name="Size(GB)";Expression={[math]::Round($_.Size/1GB,2)}}, @{Name="FreeSpace(GB)";Expression={[math]::Round($_.FreeSpace/1GB,2)}}
}

# Function: Get Storage Drives (excludes C:)
function Get-StorageDrives {
    return Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 -and $_.DeviceID -ne "C:" } | Select-Object DeviceID, @{Name="Size(GB)";Expression={[math]::Round($_.Size/1GB,2)}}, @{Name="FreeSpace(GB)";Expression={[math]::Round($_.FreeSpace/1GB,2)}}
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
        $volume = Get-WmiObject -Class Win32_Volume | Where-Object { $_.DriveLetter -eq $driveLetter }
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
        $volume = Get-WmiObject -Class Win32_Volume | Where-Object { $_.DriveLetter -eq $driveLetter }
        if ($volume) {
            $volume.IndexingEnabled = $false
            $volume.Put() | Out-Null
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

# Function: Restore Registry Backup
function Restore-RegistryBackup {
    $backupFolder = "C:\Milestonecheck-ulfh\RegistryBackups"

    if (-not (Test-Path $backupFolder)) {
        Write-Host "No backup folder found at: $backupFolder" -ForegroundColor Yellow
        return $false
    }

    $backupFiles = Get-ChildItem -Path $backupFolder -Filter "*.reg" | Sort-Object LastWriteTime -Descending

    if ($backupFiles.Count -eq 0) {
        Write-Host "No backup files found." -ForegroundColor Yellow
        return $false
    }

    Write-Host "`nAvailable backups:" -ForegroundColor Cyan
    for ($i = 0; $i -lt [Math]::Min($backupFiles.Count, 10); $i++) {
        Write-Host "$($i + 1). $($backupFiles[$i].Name) - $($backupFiles[$i].LastWriteTime)"
    }

    $choice = Read-Host "`nSelect backup to restore (1-$([Math]::Min($backupFiles.Count, 10))) or 'C' to cancel"

    if ($choice -eq 'C') {
        return $false
    }

    $index = [int]$choice - 1
    if ($index -ge 0 -and $index -lt $backupFiles.Count) {
        $selectedBackup = $backupFiles[$index]
        Write-Host "Restoring: $($selectedBackup.Name)..." -ForegroundColor Yellow

        try {
            reg import $selectedBackup.FullName 2>&1 | Out-Null
            Write-Host "Backup restored successfully." -ForegroundColor Green
            Write-Host "NOTE: A system restart may be required for changes to take effect." -ForegroundColor Yellow
            return $true
        }
        catch {
            Write-Host "Failed to restore backup: $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
    }

    return $false
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

# Function: Configure Windows Services (Updated - removed obsolete services)
function Set-WindowsServices {
    param([bool]$OptimizeForMilestone)

    # Removed HomeGroupListener and HomeGroupProvider - they don't exist in Windows 10/11
    $servicesToDisable = @(
        @{ Name = "Fax"; Display = "Fax Service" },
        @{ Name = "WSearch"; Display = "Windows Search" },
        @{ Name = "Spooler"; Display = "Print Spooler" },
        @{ Name = "SysMain"; Display = "Superfetch/SysMain" },
        @{ Name = "Themes"; Display = "Themes" },
        @{ Name = "TabletInputService"; Display = "Tablet PC Input Service" },
        @{ Name = "DiagTrack"; Display = "Connected User Experiences and Telemetry" },
        @{ Name = "dmwappushservice"; Display = "Device Management WAP Push" }
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
        $provisioned = Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -like "*$AppName*" } -ErrorAction SilentlyContinue

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
        $teamsProvisioned = Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -like "*Teams*" } -ErrorAction SilentlyContinue
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

        $outlookProvisioned = Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -like "*OutlookForWindows*" } -ErrorAction SilentlyContinue
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

        $officeProvisioned = Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -like "*OfficeHub*" } -ErrorAction SilentlyContinue
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

# Function: Get NTP Server Region
function Get-NTPServerRegion {
    Write-Host "`nSelect NTP server region:" -ForegroundColor Cyan
    Write-Host "1. Nordic (Sweden, Norway, Denmark, Finland)"
    Write-Host "2. Europe (General European pool)"
    Write-Host "3. Global (Worldwide pool)"
    Write-Host "4. Custom (Enter your own NTP servers)"

    $choice = Read-Host "`nEnter choice (1-4)"

    switch ($choice) {
        "1" {
            return "0.se.pool.ntp.org,0.no.pool.ntp.org,0.dk.pool.ntp.org,0.fi.pool.ntp.org"
        }
        "2" {
            return "0.europe.pool.ntp.org,1.europe.pool.ntp.org,2.europe.pool.ntp.org,3.europe.pool.ntp.org"
        }
        "3" {
            return "0.pool.ntp.org,1.pool.ntp.org,2.pool.ntp.org,3.pool.ntp.org"
        }
        "4" {
            $custom = Read-Host "Enter NTP servers (comma separated)"
            return $custom
        }
        default {
            Write-Host "Invalid choice, using Nordic servers." -ForegroundColor Yellow
            return "0.se.pool.ntp.org,0.no.pool.ntp.org,0.dk.pool.ntp.org,0.fi.pool.ntp.org"
        }
    }
}

# Function: Configure NTP Server for Cameras
function Enable-NTPServer {
    param([string]$NtpServers = "")

    try {
        Write-Host "`n=== CONFIGURING NTP TIME SERVER FOR CAMERAS ===" -ForegroundColor Yellow
        Write-Host "This configures Windows Time Service to act as an NTP server" -ForegroundColor Cyan
        Write-Host "for your CCTV cameras. Accurate time synchronization is CRITICAL" -ForegroundColor Cyan
        Write-Host "for surveillance systems - timestamps must be forensically accurate." -ForegroundColor Cyan
        Write-Host ""

        # Get NTP servers if not provided
        if ($NtpServers -eq "") {
            $NtpServers = Get-NTPServerRegion
        }

        # Stop the Windows Time service
        Write-Host " - Stopping Windows Time service..." -ForegroundColor Cyan
        Stop-Service w32time -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2

        # Configure Windows Time service as NTP server
        Write-Host " - Configuring time synchronization settings..." -ForegroundColor Cyan

        # Set this server as a reliable time source
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Config" -Name "AnnounceFlags" -Value 5 -Type DWord -ErrorAction Stop

        # Enable NTP Server
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\NtpServer" -Name "Enabled" -Value 1 -Type DWord -ErrorAction Stop

        # Configure NTP client to sync from selected pool
        Write-Host " - Setting NTP pool servers as time source..." -ForegroundColor Cyan
        Write-Host "   Servers: $NtpServers" -ForegroundColor Gray
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters" -Name "NtpServer" -Value $NtpServers -Type String -ErrorAction Stop
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters" -Name "Type" -Value "NTP" -Type String -ErrorAction Stop

        # Configure time synchronization via w32tm
        Write-Host " - Applying w32tm configuration..." -ForegroundColor Cyan
        & w32tm /config /manualpeerlist:$NtpServers /syncfromflags:manual /reliable:yes /update 2>&1 | Out-Null

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

        # Start the Windows Time service
        Write-Host " - Starting Windows Time service..." -ForegroundColor Cyan
        Start-Service w32time -ErrorAction Stop
        Start-Sleep -Seconds 3

        # Force immediate synchronization
        Write-Host " - Forcing immediate time synchronization..." -ForegroundColor Cyan
        & w32tm /resync /rediscover 2>&1 | Out-Null
        Start-Sleep -Seconds 2

        # Check if service is running
        $timeService = Get-Service w32time -ErrorAction SilentlyContinue
        if ($timeService.Status -eq 'Running') {
            Write-Host "[OK] Windows Time service is running" -ForegroundColor Green
        } else {
            Write-Host "[WARN] Windows Time service status: $($timeService.Status)" -ForegroundColor Yellow
        }

        # Display configuration summary
        Write-Host "`n=== NTP SERVER CONFIGURATION SUMMARY ===" -ForegroundColor Green
        Write-Host "NTP Server Status: ENABLED" -ForegroundColor Green
        Write-Host "Time Source: $NtpServers" -ForegroundColor Green
        Write-Host "Firewall: UDP Port 123 opened for incoming connections" -ForegroundColor Green
        Write-Host "Server Mode: Reliable NTP server for camera synchronization" -ForegroundColor Green
        Write-Host ""
        Write-Host "Your cameras can now sync time from this server!" -ForegroundColor Green
        Write-Host "Configure your cameras to use this server's IP address as NTP server." -ForegroundColor Cyan
        Write-Host ""

        return $true

    } catch {
        Write-Host "[FAIL] Failed to configure NTP server: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Attempting to start Windows Time service anyway..." -ForegroundColor Yellow
        Start-Service w32time -ErrorAction SilentlyContinue
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
    @{ Name = "Microsoft.MixedReality.Portal"; Display = "Mixed Reality Portal" }
)

# Function: Display Banner
function Show-Banner {
    Clear-Host
    Write-Host "Enhanced Milestone XProtect System Configuration Script" -ForegroundColor Cyan
    Write-Host "By Ulf Holmstrom, Happy Problem Solver at Manvarg AB (Rev 2.0, 2025)" -ForegroundColor Cyan
    Write-Host "IMPORTANT DISCLAIMER: This script has been developed privately by Ulf Holmstrom, to facilitate resellers. It is provided as a public resource and is NOT supported by Milestone Systems A/S." -ForegroundColor Magenta
    Write-Host "For questions contact: ulf@manvarg.se" -ForegroundColor Cyan
    Write-Host ""
}

# Function: Main Menu
function Show-MainMenu {
    Write-Host "=== MAIN MENU ===" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "STORAGE & SECURITY" -ForegroundColor Cyan
    Write-Host "  1. Antivirus Exceptions (for XProtect processes and storage)"
    Write-Host "  2. Storage Drive Optimization (block size, indexing)"
    Write-Host ""
    Write-Host "WINDOWS CONFIGURATION" -ForegroundColor Cyan
    Write-Host "  3. Windows Update Management"
    Write-Host "  4. Remove Bloatware Apps"
    Write-Host ""
    Write-Host "SYSTEM PERFORMANCE" -ForegroundColor Cyan
    Write-Host "  5. Power Management (Never Sleep Mode)"
    Write-Host "  6. Optimize Windows Services"
    Write-Host ""
    Write-Host "REMOTE ACCESS & NETWORK" -ForegroundColor Cyan
    Write-Host "  7. Enable Remote Desktop (RDP)"
    Write-Host "  8. Configure NTP Time Server (for cameras)"
    Write-Host "  9. Install SNMP Capabilities (for monitoring)"
    Write-Host ""
    Write-Host "MILESTONE TOOLS" -ForegroundColor Cyan
    Write-Host "  10. Install Milestone PowerShell Tools"
    Write-Host ""
    Write-Host "AUTOMATION & LOGGING" -ForegroundColor Cyan
    Write-Host "  11. COMPLETE SYSTEM SETUP (Guided Wizard)"
    Write-Host "  12. Generate Log File"
    Write-Host "  13. Restore from Backup"
    Write-Host ""
    Write-Host "  14. Exit"
    Write-Host ""
}

# Function: Antivirus Exceptions Only
function Invoke-AntivirusConfig {
    Write-Host "`n=== ANTIVIRUS EXCEPTIONS CONFIGURATION ===" -ForegroundColor Yellow
    Write-Host "This adds Windows Defender exceptions for XProtect processes and storage drives." -ForegroundColor Cyan
    Write-Host ""
    Write-Host "SECURITY NOTE: The C: drive (OS) will be EXCLUDED from antivirus exceptions" -ForegroundColor Red
    Write-Host "for security reasons. Only storage drives should have exceptions." -ForegroundColor Red
    Write-Host ""

    $drives = Get-StorageDrives
    if ($drives) {
        Write-Host "Available storage drives (C: excluded):" -ForegroundColor Cyan
        $drives | Format-Table -AutoSize

        Write-Host "Select drives for antivirus exceptions:" -ForegroundColor Green
        $driveInput = Read-Host "Enter drive letters (comma separated, e.g., d,e) or 'all' for all storage drives"

        if ($driveInput.ToLower() -eq 'all') {
            $selectedDrives = $drives.DeviceID
        } else {
            $selectedDrives = $driveInput -split "," | ForEach-Object {
                $drive = ($_.Trim().ToUpper() + ":").Replace("::", ":")
                if ($drive -eq "C:") {
                    Write-Host "WARNING: C: drive excluded for security" -ForegroundColor Red
                    return $null
                }
                return $drive
            } | Where-Object { $_ -ne $null }
        }

        Write-Host "`n--- Adding Antivirus Exceptions ---" -ForegroundColor Yellow
        foreach ($drive in $selectedDrives) {
            $drivePath = $drive + "\"
            try {
                Add-MpPreference -ExclusionPath $drivePath -ErrorAction Stop
                Write-Host "[OK] Antivirus exception added for drive ${drivePath}" -ForegroundColor Green
                $global:logContent += "Antivirus exception added for drive ${drivePath}`r`n"
            }
            catch {
                Write-Host "[FAIL] Error adding antivirus exception for drive ${drivePath}: $_" -ForegroundColor Red
                $global:logContent += "Error adding antivirus exception for drive ${drivePath}: $_`r`n"
            }
        }
    } else {
        Write-Host "No storage drives found (only C: drive present)." -ForegroundColor Yellow
    }

    # Add file extensions
    $exclusionExtensions = @("blk", "idx")
    foreach ($ext in $exclusionExtensions) {
        try {
            Add-MpPreference -ExclusionExtension $ext -ErrorAction Stop
            Write-Host "[OK] Antivirus exception added for file extension: ${ext}" -ForegroundColor Green
            $global:logContent += "Antivirus exception added for file extension: ${ext}`r`n"
        }
        catch {
            Write-Host "[FAIL] Error adding antivirus exception for file extension ${ext}: $_" -ForegroundColor Red
        }
    }

    # Add XProtect processes
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
            $global:logContent += "Antivirus exception added for process: ${proc}`r`n"
        }
        catch {
            Write-Host "[FAIL] Error adding antivirus exception for process ${proc}: $_" -ForegroundColor Red
        }
    }

    Read-Host "`nAntivirus exceptions configuration complete. Press Enter to return to main menu..."
}

# Function: Storage Drive Optimization Only
function Invoke-StorageOptimization {
    Write-Host "`n=== STORAGE DRIVE OPTIMIZATION ===" -ForegroundColor Yellow
    Write-Host "This checks and optimizes storage drives for video recording." -ForegroundColor Cyan
    Write-Host ""

    $drives = Get-StorageDrives
    if (-not $drives) {
        Write-Host "No storage drives found (only C: drive present)." -ForegroundColor Yellow
        Read-Host "Press Enter to return to main menu..."
        return
    }

    Write-Host "Available storage drives:" -ForegroundColor Cyan
    $drives | Format-Table -AutoSize

    $driveInput = Read-Host "Enter drive letters to optimize (comma separated, e.g., d,e) or 'all'"

    if ($driveInput.ToLower() -eq 'all') {
        $selectedDrives = $drives.DeviceID
    } else {
        $selectedDrives = $driveInput -split "," | ForEach-Object { ($_.Trim().ToUpper() + ":").Replace("::", ":") }
    }

    Write-Host "`n--- Storage Drive Configuration Check ---" -ForegroundColor Yellow
    foreach ($drive in $selectedDrives) {
        $driveForCheck = $drive.TrimEnd(':') + ":"
        $drivePath = $driveForCheck + "\"

        Write-Host "`nChecking storage drive ${driveForCheck}..." -ForegroundColor Cyan

        # Check block size
        $blockSize = Test-DriveBlockSize $drivePath
        if ($blockSize -eq 65536) {
            Write-Host "[OK] ${driveForCheck}: Correct block size - 64 KB" -ForegroundColor Green
            $global:logContent += "Storage Drive ${driveForCheck}: Correct block size - 64 KB.`r`n"
        } elseif ($blockSize) {
            $blockSizeKB = $blockSize / 1024
            Write-Host "[WARN] ${driveForCheck}: Block size is ${blockSizeKB} KB (expected 64 KB)" -ForegroundColor Yellow
            Write-Host "       NOTE: Drive needs to be reformatted with 64KB allocation unit size for optimal performance." -ForegroundColor Yellow
            $global:logContent += "Storage Drive ${driveForCheck}: Incorrect block size - ${blockSizeKB} KB.`r`n"
        } else {
            Write-Host "[?] ${driveForCheck}: Could not determine block size" -ForegroundColor Yellow
        }

        # Check indexing
        $indexingEnabled = Test-DriveIndexing $drivePath
        if ($indexingEnabled -eq $false) {
            Write-Host "[OK] ${driveForCheck}: Indexing is OFF" -ForegroundColor Green
            $global:logContent += "Storage Drive ${driveForCheck}: Indexing is OFF.`r`n"
        } elseif ($indexingEnabled -eq $true) {
            Write-Host "[WARN] ${driveForCheck}: Indexing is ON (should be disabled for storage)" -ForegroundColor Yellow
            $disableChoice = Read-Host "       Disable indexing for ${driveForCheck}? (Y/N)"
            if ($disableChoice -match "^[Yy]") {
                if (Disable-DriveIndexing $drivePath) {
                    Write-Host "[OK] ${driveForCheck}: Indexing disabled" -ForegroundColor Green
                    $global:logContent += "Storage Drive ${driveForCheck}: Indexing disabled.`r`n"
                } else {
                    Write-Host "[FAIL] ${driveForCheck}: Failed to disable indexing" -ForegroundColor Red
                }
            }
        }
    }

    Read-Host "`nStorage optimization complete. Press Enter to return to main menu..."
}

# Function: Windows Update Management
function Invoke-WindowsUpdateManagement {
    Write-Host "`n=== WINDOWS UPDATE MANAGEMENT ===" -ForegroundColor Yellow
    Write-Host "NOTE: Keep updates enabled until after installing SNMP or PowerShell Tools!" -ForegroundColor Cyan

    $currentStatus = Get-WindowsUpdateStatus
    if ($currentStatus -eq $true) {
        Write-Host "Current Status: Windows Updates are ENABLED" -ForegroundColor Green
        $global:logContent += "Windows Updates: Currently ENABLED`r`n"
    } elseif ($currentStatus -eq $false) {
        Write-Host "Current Status: Windows Updates are DISABLED" -ForegroundColor Red
        $global:logContent += "Windows Updates: Currently DISABLED`r`n"
    } else {
        Write-Host "Current Status: Unable to determine Windows Update status" -ForegroundColor Yellow
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

    Write-Host "CONSUMER APPS" -ForegroundColor Cyan
    Write-Host "  1. Remove ALL standard bloatware (games, weather, news, etc.)"
    Write-Host "  2. Select specific apps to remove"
    Write-Host ""
    Write-Host "MICROSOFT CLOUD SERVICES" -ForegroundColor Cyan
    Write-Host "  3. Remove OneDrive only"
    Write-Host "  4. Remove Teams + Outlook App + Office Hub"
    Write-Host ""
    Write-Host "SYSTEM COMPONENTS" -ForegroundColor Cyan
    Write-Host "  5. Remove Xbox Apps (WARNING: May cause issues)"
    Write-Host ""
    Write-Host "COMPLETE" -ForegroundColor Cyan
    Write-Host "  6. Full cleanup - Everything above (except Xbox)"
    Write-Host ""
    Write-Host "  7. Return to Main Menu"

    $choice = Read-Host "`nEnter your choice (1-7)"

    switch ($choice) {
        "1" {
            Write-Host "`nRemoving ALL standard bloatware apps..." -ForegroundColor Yellow
            $removedCount = 0
            foreach ($app in $global:bloatwareApps) {
                if (Remove-BloatwareApp -AppName $app.Name -FriendlyName $app.Display) {
                    $global:logContent += "Bloatware removed: $($app.Display)`r`n"
                    $removedCount++
                }
            }
            Write-Host "`nBloatware removal completed. $removedCount apps were removed." -ForegroundColor Green
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
            Write-Host " - Teams: $(if($teamsResult){'Removed'}else{'Not found'})" -ForegroundColor $(if($teamsResult){'Green'}else{'Gray'})
            Write-Host " - Outlook App: $(if($outlookResult){'Removed'}else{'Not found'})" -ForegroundColor $(if($outlookResult){'Green'}else{'Gray'})
            Write-Host " - Office Hub: $(if($officeResult){'Removed'}else{'Not found'})" -ForegroundColor $(if($officeResult){'Green'}else{'Gray'})

            $global:logContent += "Microsoft Teams: $(if($teamsResult){'Removed'}else{'Not found'})`r`n"
            $global:logContent += "Outlook App: $(if($outlookResult){'Removed'}else{'Not found'})`r`n"
            $global:logContent += "Office Hub: $(if($officeResult){'Removed'}else{'Not found'})`r`n"
        }
        "5" {
            Remove-XboxApps
        }
        "6" {
            Write-Host "`nPerforming FULL CLEANUP (Xbox excluded)..." -ForegroundColor Yellow
            $confirm = Read-Host "This will remove ALL bloatware, OneDrive, Teams, Outlook, Office Hub. Continue? (Y/N)"
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
                Write-Host "Plus: OneDrive, Teams, Outlook App, Office Hub removed" -ForegroundColor Green
                $global:logContent += "Full cleanup: $removedCount apps, OneDrive, Teams, Outlook, Office Hub removed`r`n"
            }
        }
        "7" {
            return
        }
    }

    Read-Host "`nPress Enter to return to main menu..."
}

# Function: Power Management Only
function Invoke-PowerManagement {
    Write-Host "`n=== POWER MANAGEMENT CONFIGURATION ===" -ForegroundColor Yellow
    Write-Host "Configure power settings for 24/7 video recording operation." -ForegroundColor Cyan
    Write-Host ""
    Write-Host "CRITICAL: XProtect servers MUST be configured to NEVER sleep!" -ForegroundColor Red
    Write-Host "Sleep/hibernate will cause recording interruptions and camera disconnections." -ForegroundColor Red
    Write-Host ""

    Write-Host "Options:" -ForegroundColor Cyan
    Write-Host "1. Enable High Performance + Never Sleep (RECOMMENDED)"
    Write-Host "2. Restore to Balanced mode"
    Write-Host "3. Return to Main Menu"

    $choice = Read-Host "`nEnter your choice (1-3)"

    switch ($choice) {
        "1" {
            Write-Host "`nConfiguring Never Sleep Mode..." -ForegroundColor Yellow
            New-RegistryBackup -BackupName "PowerManagement"
            if (Set-PowerManagement -HighPerformance $true) {
                $global:logContent += "Power Management: Never Sleep Mode activated`r`n"
            }
        }
        "2" {
            if (Set-PowerManagement -HighPerformance $false) {
                $global:logContent += "Power Management: Restored to Balanced`r`n"
            }
        }
        "3" {
            return
        }
    }

    Read-Host "`nPress Enter to return to main menu..."
}

# Function: Windows Services Optimization Only
function Invoke-ServicesOptimization {
    Write-Host "`n=== WINDOWS SERVICES OPTIMIZATION ===" -ForegroundColor Yellow
    Write-Host "Disable unnecessary Windows services to free up system resources." -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Services to be disabled:" -ForegroundColor Cyan
    Write-Host " - Fax Service"
    Write-Host " - Windows Search (WSearch)"
    Write-Host " - Print Spooler"
    Write-Host " - Superfetch/SysMain"
    Write-Host " - Themes"
    Write-Host " - Tablet PC Input Service"
    Write-Host " - Connected User Experiences and Telemetry"
    Write-Host " - Device Management WAP Push"
    Write-Host ""

    $confirm = Read-Host "Disable these services? (Y/N)"
    if ($confirm -match "^[Yy]") {
        Write-Host "`nCreating registry backup..." -ForegroundColor Cyan
        New-RegistryBackup -BackupName "ServicesOptimization"

        Write-Host "Optimizing Windows Services..." -ForegroundColor Yellow
        $disabledCount = Set-WindowsServices -OptimizeForMilestone $true
        Write-Host "`nService optimization completed. $disabledCount services disabled." -ForegroundColor Green
        $global:logContent += "Windows Services: $disabledCount services disabled`r`n"
    }

    Read-Host "`nPress Enter to return to main menu..."
}

# Function: Remote Desktop Configuration
function Invoke-RemoteDesktopConfig {
    Write-Host "`n=== REMOTE DESKTOP CONFIGURATION ===" -ForegroundColor Yellow
    Write-Host "Enable and optimize Remote Desktop for server management." -ForegroundColor Cyan
    Write-Host ""

    Write-Host "This will:" -ForegroundColor Cyan
    Write-Host " - Enable Remote Desktop connections"
    Write-Host " - Configure Windows Firewall for RDP"
    Write-Host " - Optimize RDP for performance (disable wallpaper, animations)"
    Write-Host ""

    $confirm = Read-Host "Enable and optimize Remote Desktop? (Y/N)"
    if ($confirm -match "^[Yy]") {
        Write-Host "`nCreating registry backup..." -ForegroundColor Cyan
        New-RegistryBackup -BackupName "RemoteDesktop"

        if (Set-RemoteDesktopOptimization -OptimizeRDP $true) {
            $global:logContent += "Remote Desktop: ENABLED and optimized`r`n"
        }
    }

    Read-Host "`nPress Enter to return to main menu..."
}

# Function: NTP Server Configuration
function Invoke-NTPServerConfiguration {
    Write-Host "`n=== NTP TIME SERVER FOR CAMERAS ===" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Why do you need this?" -ForegroundColor Cyan
    Write-Host "CCTV cameras MUST have accurate time synchronization for:" -ForegroundColor White
    Write-Host " - Forensically valid video timestamps"
    Write-Host " - Synchronized recording across multiple cameras"
    Write-Host " - Accurate event correlation in XProtect"
    Write-Host " - Legal compliance (timestamps must be accurate in court)"
    Write-Host ""
    Write-Host "This configuration will:" -ForegroundColor Cyan
    Write-Host " - Configure Windows Time service as a reliable NTP server"
    Write-Host " - Sync server time from regional NTP pool"
    Write-Host " - Open firewall UDP port 123 for camera connections"
    Write-Host ""

    $confirm = Read-Host "Configure NTP Time Server for cameras? (Y/N)"
    if ($confirm -match "^[Yy]") {
        if (Enable-NTPServer) {
            Write-Host "`nNTP Time Server configuration completed successfully." -ForegroundColor Green
            $global:logContent += "NTP Server: Configured successfully`r`n"
        } else {
            Write-Host "`nNTP Time Server configuration encountered issues." -ForegroundColor Red
            $global:logContent += "NTP Server: Configuration failed`r`n"
        }
    } else {
        Write-Host "NTP Server configuration cancelled." -ForegroundColor Yellow
    }

    Read-Host "`nPress Enter to return to main menu..."
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
    Write-Host ""
    Write-Host "If you're unsure, you probably DON'T need it." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "NOTE: Windows Update must be enabled for this to work!" -ForegroundColor Red
    Write-Host ""

    # Check if Windows Update is disabled
    $updateStatus = Get-WindowsUpdateStatus
    if ($updateStatus -eq $false) {
        Write-Host "WARNING: Windows Update is currently DISABLED!" -ForegroundColor Red
        $enableUpdate = Read-Host "Enable Windows Update temporarily? (Y/N)"
        if ($enableUpdate -match "^[Yy]") {
            Set-WindowsUpdateStatus -Enable $true
            Write-Host "Windows Update enabled." -ForegroundColor Green
        } else {
            Write-Host "SNMP installation cancelled. Enable Windows Update first." -ForegroundColor Yellow
            Read-Host "`nPress Enter to return to main menu..."
            return
        }
    }

    $confirm = Read-Host "Install SNMP capabilities? (Y/N)"
    if ($confirm -match "^[Yy]") {
        if (Enable-SNMP) {
            Write-Host "SNMP installation completed successfully." -ForegroundColor Green
            $global:logContent += "SNMP: Installed successfully`r`n"
        } else {
            Write-Host "SNMP installation failed." -ForegroundColor Red
            $global:logContent += "SNMP: Installation failed`r`n"
        }
    }

    Read-Host "`nPress Enter to return to main menu..."
}

# Function: Milestone PowerShell Tools Installation
function Invoke-MilestonePSToolsInstallation {
    Write-Host "`n=== MILESTONE POWERSHELL TOOLS INSTALLATION ===" -ForegroundColor Yellow
    Write-Host "Install the Milestone PowerShell Tools for advanced XProtect management." -ForegroundColor Cyan
    Write-Host ""

    $confirm = Read-Host "Install Milestone PowerShell Tools? (Y/N)"
    if ($confirm -match "^[Yy]") {
        if (Install-MilestonePSTools) {
            Write-Host "Milestone PowerShell Tools installation completed successfully." -ForegroundColor Green
            $global:logContent += "Milestone PowerShell Tools: Installed successfully`r`n"
        } else {
            Write-Host "Milestone PowerShell Tools installation failed." -ForegroundColor Red
            $global:logContent += "Milestone PowerShell Tools: Installation failed`r`n"
        }
    }

    Read-Host "`nPress Enter to return to main menu..."
}

# Function: Restore from Backup
function Invoke-RestoreBackup {
    Write-Host "`n=== RESTORE FROM BACKUP ===" -ForegroundColor Yellow
    Write-Host "Restore registry settings from a previous backup." -ForegroundColor Cyan
    Write-Host ""
    Write-Host "WARNING: This will overwrite current registry settings!" -ForegroundColor Red
    Write-Host ""

    if (Restore-RegistryBackup) {
        $global:logContent += "Registry restored from backup`r`n"
    }

    Read-Host "`nPress Enter to return to main menu..."
}

# Function: Complete System Setup (Guided Wizard)
function Invoke-CompleteSystemSetup {
    Write-Host "`n=== COMPLETE MILESTONE XPROTECT SYSTEM SETUP ===" -ForegroundColor Yellow
    Write-Host "This guided wizard will configure your system for optimal XProtect performance." -ForegroundColor Cyan
    Write-Host ""
    Write-Host "The wizard will ask you about each configuration step:" -ForegroundColor White
    Write-Host "  1. Antivirus exceptions for storage drives"
    Write-Host "  2. Storage drive optimization"
    Write-Host "  3. Bloatware removal"
    Write-Host "  4. Power management (Never Sleep)"
    Write-Host "  5. Windows services optimization"
    Write-Host "  6. Remote Desktop"
    Write-Host "  7. NTP Time Server"
    Write-Host "  8. Milestone PowerShell Tools"
    Write-Host "  9. SNMP capabilities"
    Write-Host "  10. Windows Update configuration"
    Write-Host ""
    Write-Host "WARNING: This will make significant system changes!" -ForegroundColor Red
    Write-Host ""

    $confirm = Read-Host "Start the guided setup wizard? (Y/N)"
    if ($confirm -notmatch "^[Yy]") {
        Write-Host "Complete system setup cancelled." -ForegroundColor Yellow
        Read-Host "Press Enter to return to main menu..."
        return
    }

    Write-Host "`nCreating comprehensive registry backup..." -ForegroundColor Cyan
    New-RegistryBackup -BackupName "CompleteSetup"

    # Ensure Windows Update is enabled for installations
    Write-Host "`nEnsuring Windows Update is enabled for installations..." -ForegroundColor Cyan
    Set-WindowsUpdateStatus -Enable $true

    Write-Host "`n=== STARTING GUIDED SETUP WIZARD ===" -ForegroundColor Green
    $startTime = Get-Date

    # Step 1: Antivirus Exceptions
    Write-Host "`n--- STEP 1/10: ANTIVIRUS EXCEPTIONS ---" -ForegroundColor Yellow
    $avChoice = Read-Host "Configure antivirus exceptions for XProtect? (Y/N)"
    if ($avChoice -match "^[Yy]") {
        $drives = Get-StorageDrives
        if ($drives) {
            Write-Host "Available storage drives:" -ForegroundColor Cyan
            $drives | Format-Table -AutoSize
            $driveInput = Read-Host "Enter drive letters (comma separated) or 'all'"

            if ($driveInput.ToLower() -eq 'all') {
                $selectedDrives = $drives.DeviceID
            } else {
                $selectedDrives = $driveInput -split "," | ForEach-Object { ($_.Trim().ToUpper() + ":").Replace("::", ":") } | Where-Object { $_ -ne "C:" }
            }

            foreach ($drive in $selectedDrives) {
                $drivePath = $drive + "\"
                try {
                    Add-MpPreference -ExclusionPath $drivePath -ErrorAction Stop
                    Write-Host "[OK] Exception: ${drivePath}" -ForegroundColor Green
                    $global:logContent += "WIZARD - AV exception: ${drivePath}`r`n"
                } catch { }
            }
        }

        # Extensions and processes
        @("blk", "idx") | ForEach-Object { Add-MpPreference -ExclusionExtension $_ -ErrorAction SilentlyContinue }
        @("VideoOS.Recorder.Service.exe", "VideoOS.Server.Service.exe", "VideoOS.Administration.exe", "VideoOS.Event.Server.exe", "VideoOS.Failover.Service.exe", "VideoOS.MobileServer.Service.exe", "VideoOS.LPR.Server.exe") | ForEach-Object { Add-MpPreference -ExclusionProcess $_ -ErrorAction SilentlyContinue }
        Write-Host "[OK] XProtect process and extension exceptions added" -ForegroundColor Green
    }

    # Step 2: Storage Optimization
    Write-Host "`n--- STEP 2/10: STORAGE OPTIMIZATION ---" -ForegroundColor Yellow
    $storageChoice = Read-Host "Optimize storage drives (check block size, disable indexing)? (Y/N)"
    if ($storageChoice -match "^[Yy]") {
        $drives = Get-StorageDrives
        if ($drives) {
            foreach ($drive in $drives.DeviceID) {
                $drivePath = $drive + "\"
                $indexingEnabled = Test-DriveIndexing $drivePath
                if ($indexingEnabled -eq $true) {
                    Disable-DriveIndexing $drivePath | Out-Null
                    Write-Host "[OK] ${drive}: Indexing disabled" -ForegroundColor Green
                    $global:logContent += "WIZARD - ${drive}: Indexing disabled`r`n"
                }
            }
        }
    }

    # Step 3: Bloatware Removal
    Write-Host "`n--- STEP 3/10: BLOATWARE REMOVAL ---" -ForegroundColor Yellow
    $bloatChoice = Read-Host "Remove bloatware (Teams, OneDrive, etc.)? (Y/N)"
    if ($bloatChoice -match "^[Yy]") {
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
        Write-Host "[OK] Bloatware removal: $removedCount apps removed" -ForegroundColor Green
        $global:logContent += "WIZARD - Bloatware: $removedCount apps removed`r`n"
    }

    # Step 4: Power Management
    Write-Host "`n--- STEP 4/10: POWER MANAGEMENT ---" -ForegroundColor Yellow
    $powerChoice = Read-Host "Configure Never Sleep Mode (CRITICAL for 24/7 recording)? (Y/N)"
    if ($powerChoice -match "^[Yy]") {
        Set-PowerManagement -HighPerformance $true | Out-Null
        $global:logContent += "WIZARD - Power: Never Sleep Mode activated`r`n"
    }

    # Step 5: Services Optimization
    Write-Host "`n--- STEP 5/10: SERVICES OPTIMIZATION ---" -ForegroundColor Yellow
    $servicesChoice = Read-Host "Disable unnecessary Windows services? (Y/N)"
    if ($servicesChoice -match "^[Yy]") {
        $disabledCount = Set-WindowsServices -OptimizeForMilestone $true
        Write-Host "[OK] $disabledCount services disabled" -ForegroundColor Green
        $global:logContent += "WIZARD - Services: $disabledCount disabled`r`n"
    }

    # Step 6: Remote Desktop
    Write-Host "`n--- STEP 6/10: REMOTE DESKTOP ---" -ForegroundColor Yellow
    $rdpChoice = Read-Host "Enable and optimize Remote Desktop? (Y/N)"
    if ($rdpChoice -match "^[Yy]") {
        Set-RemoteDesktopOptimization -OptimizeRDP $true | Out-Null
        $global:logContent += "WIZARD - RDP: Enabled and optimized`r`n"
    }

    # Step 7: NTP Server
    Write-Host "`n--- STEP 7/10: NTP TIME SERVER ---" -ForegroundColor Yellow
    Write-Host "NTP is CRITICAL for camera time synchronization and forensic timestamps." -ForegroundColor Cyan
    $ntpChoice = Read-Host "Configure NTP Time Server for cameras? (Y/N)"
    if ($ntpChoice -match "^[Yy]") {
        $ntpServers = Get-NTPServerRegion
        Enable-NTPServer -NtpServers $ntpServers | Out-Null
        $global:logContent += "WIZARD - NTP: Configured`r`n"
    }

    # Step 8: Milestone PowerShell Tools
    Write-Host "`n--- STEP 8/10: MILESTONE POWERSHELL TOOLS ---" -ForegroundColor Yellow
    $psToolsChoice = Read-Host "Install Milestone PowerShell Tools? (Y/N)"
    if ($psToolsChoice -match "^[Yy]") {
        if (Install-MilestonePSTools) {
            $global:logContent += "WIZARD - MilestonePSTools: Installed`r`n"
        }
    }

    # Step 9: SNMP
    Write-Host "`n--- STEP 9/10: SNMP CAPABILITIES ---" -ForegroundColor Yellow
    Write-Host "Only needed for network monitoring systems (PRTG, Nagios, Zabbix)." -ForegroundColor Cyan
    $snmpChoice = Read-Host "Install SNMP capabilities? (Y/N)"
    if ($snmpChoice -match "^[Yy]") {
        if (Enable-SNMP) {
            $global:logContent += "WIZARD - SNMP: Installed`r`n"
        }
    }

    # Step 10: Windows Update
    Write-Host "`n--- STEP 10/10: WINDOWS UPDATE ---" -ForegroundColor Yellow
    Write-Host "All installations complete. Configure Windows Update:" -ForegroundColor Cyan
    $updateChoice = Read-Host "Windows Updates: (E)nable, (D)isable, or (S)kip? (E/D/S)"
    switch ($updateChoice.ToUpper()) {
        "D" {
            Set-WindowsUpdateStatus -Enable $false
            Write-Host "[OK] Windows Updates DISABLED" -ForegroundColor Red
            $global:logContent += "WIZARD - Updates: DISABLED`r`n"
        }
        default {
            Write-Host "[OK] Windows Updates remain ENABLED" -ForegroundColor Green
            $global:logContent += "WIZARD - Updates: ENABLED`r`n"
        }
    }

    # Final Summary
    $endTime = Get-Date
    $duration = $endTime - $startTime

    Write-Host "`n=== SETUP WIZARD COMPLETED ===" -ForegroundColor Green
    Write-Host "Completed in $($duration.Minutes) minutes and $($duration.Seconds) seconds" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "NEXT STEPS:" -ForegroundColor Yellow
    Write-Host " 1. RESTART the system to apply all changes"
    Write-Host " 2. Install Milestone XProtect software"
    Write-Host " 3. Configure cameras to use this server's IP as NTP server"
    Write-Host ""

    # Auto-generate log
    $logFolderPath = "C:\Milestonecheck-ulfh"
    if (-not (Test-Path $logFolderPath)) {
        New-Item -ItemType Directory -Path $logFolderPath | Out-Null
    }
    $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $logFilePath = Join-Path $logFolderPath "SetupWizardLog_$timestamp.txt"
    $header = "Milestone XProtect SETUP WIZARD Log - $(Get-Date)`r`nBy Ulf Holmstrom, Manvarg AB (Rev 2.0, 2025)`r`nContact: ulf@manvarg.se`r`n===========================================`r`n"
    $fullLogContent = $header + $global:logContent
    $fullLogContent | Out-File -FilePath $logFilePath -Encoding UTF8
    Write-Host "[OK] Log saved: $logFilePath" -ForegroundColor Green

    Play-Fanfare

    Read-Host "`nPress Enter to return to main menu..."
}

# Function: Generate Log File
function Invoke-LogGeneration {
    Write-Host "`n=== LOG FILE GENERATION ===" -ForegroundColor Yellow

    if ($global:logContent -eq "") {
        Write-Host "No configuration changes have been made yet. Nothing to log." -ForegroundColor Yellow
        Read-Host "Press Enter to return to main menu..."
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
        $header = "Milestone XProtect Configuration Log - $(Get-Date)`r`nBy Ulf Holmstrom, Manvarg AB (Rev 2.0, 2025)`r`nContact: ulf@manvarg.se`r`n===========================================`r`n"
        $fullLogContent = $header + $global:logContent
        $fullLogContent | Out-File -FilePath $logFilePath -Encoding UTF8
        Write-Host "Log file saved at: $logFilePath" -ForegroundColor Green
        Play-Fanfare
    }

    Read-Host "Press Enter to return to main menu..."
}

# Main Script Execution
Show-Banner
Write-Host "This script will help you configure Milestone XProtect system settings."
Write-Host ""
Write-Host "Rev 2.0 IMPROVEMENTS:" -ForegroundColor Cyan
Write-Host "  - RESTRUCTURED: Better menu organization with logical categories" -ForegroundColor Green
Write-Host "  - SEPARATED: Antivirus exceptions now separate from storage optimization" -ForegroundColor Green
Write-Host "  - SEPARATED: Remote Desktop moved to Remote Access & Network category" -ForegroundColor Green
Write-Host "  - ADDED: Region selection for NTP servers (Nordic/Europe/Global)" -ForegroundColor Green
Write-Host "  - FIXED: C: drive always excluded from AV exceptions (security)" -ForegroundColor Green
Write-Host "  - FIXED: Removed obsolete services (HomeGroup)" -ForegroundColor Green
Write-Host "  - ADDED: Restore from backup option" -ForegroundColor Green
Write-Host "  - IMPROVED: Guided wizard asks about ALL optional services" -ForegroundColor Green
Write-Host ""
Read-Host "Press Enter to continue..."

do {
    Show-Banner
    Show-MainMenu
    $choice = Read-Host "Enter your choice (1-14)"

    switch ($choice) {
        "1" { Invoke-AntivirusConfig }
        "2" { Invoke-StorageOptimization }
        "3" { Invoke-WindowsUpdateManagement }
        "4" { Invoke-BloatwareRemoval }
        "5" { Invoke-PowerManagement }
        "6" { Invoke-ServicesOptimization }
        "7" { Invoke-RemoteDesktopConfig }
        "8" { Invoke-NTPServerConfiguration }
        "9" { Invoke-SNMPConfiguration }
        "10" { Invoke-MilestonePSToolsInstallation }
        "11" { Invoke-CompleteSystemSetup }
        "12" { Invoke-LogGeneration }
        "13" { Invoke-RestoreBackup }
        "14" {
            Write-Host "Exiting script. Thank you for using XProtect Prepare!" -ForegroundColor Green
            Write-Host "For questions or feedback, contact: ulf@manvarg.se" -ForegroundColor Cyan
            break
        }
        default {
            Write-Host "Invalid choice. Please select 1-14." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
} while ($choice -ne "14")
