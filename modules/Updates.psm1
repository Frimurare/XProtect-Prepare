<#
    XProtect-Prepare-v2 Windows Update Module
    Authored by Ulf Holmström, ex employee, Solution Engineer at Milestone Systems December 2025
#>

function Get-WindowsUpdateStatus {
    try {
        $regPath = "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\AU"
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

function Set-WindowsUpdateStatus {
    param([bool]$Enable)
    try {
        $regPath = "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\AU"
        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }
        if ($Enable) {
            Set-ItemProperty -Path $regPath -Name "NoAutoUpdate" -Value 0 -Type DWord
            Write-Host "Windows Updates ENABLED" -ForegroundColor Green
            return $true
        } else {
            Set-ItemProperty -Path $regPath -Name "NoAutoUpdate" -Value 1 -Type DWord
            Write-Host "Windows Updates DISABLED" -ForegroundColor Yellow
            return $true
        }
    }
    catch {
        Write-Host "Failed to configure Windows Update: $_" -ForegroundColor Red
        return $false
    }
}

function New-RegistryBackup {
    param(
        [string]$BackupName = "XProtect_Backup"
    )

    $backupPath = "C:\\Milestonecheck-ulfh\\RegistryBackups"
    if (-not (Test-Path $backupPath)) {
        New-Item -ItemType Directory -Path $backupPath | Out-Null
    }

    $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $fileName = "${BackupName}_${timestamp}.reg"
    $fullPath = Join-Path $backupPath $fileName

    Write-Host "Creating registry backup: $fullPath" -ForegroundColor Cyan
    reg export HKLM $fullPath /y | Out-Null
    $Global:logContent += "Registry backup created: ${fullPath}`r`n"
}

function Invoke-WindowsUpdateMenu {
    Write-Host "`n=== WINDOWS UPDATE MANAGEMENT ===" -ForegroundColor Yellow
    Write-Host "NOTE: Keep updates enabled for SNMP or PowerShell Tools installation." -ForegroundColor Cyan

    $currentStatus = Get-WindowsUpdateStatus
    if ($currentStatus -eq $true) {
        Write-Host "Current Status: Windows Updates are ENABLED" -ForegroundColor Green
        $Global:logContent += "Windows Updates: Currently ENABLED`r`n"
    } elseif ($currentStatus -eq $false) {
        Write-Host "Current Status: Windows Updates are DISABLED" -ForegroundColor Red
        $Global:logContent += "Windows Updates: Currently DISABLED`r`n"
    } else {
        Write-Host "Current Status: Unable to determine Windows Update status" -ForegroundColor Yellow
        $Global:logContent += "Windows Updates: Status unknown`r`n"
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
                $Global:logContent += "Windows Updates: ENABLED by script`r`n"
            }
        }
        "2" {
            if (Set-WindowsUpdateStatus -Enable $false) {
                Write-Host "Windows Updates have been DISABLED." -ForegroundColor Red
                $Global:logContent += "Windows Updates: DISABLED by script`r`n"
            }
        }
        default { return }
    }

    Read-Host "`nPress Enter to return to main menu..." | Out-Null
}

Export-ModuleMember -Function *
