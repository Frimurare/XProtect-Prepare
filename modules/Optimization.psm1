<#
    XProtect-Prepare-v2 Optimization Module
    Authored by Ulf Holmström, ex employee, Solution Engineer at Milestone Systems December 2025
#>

function Set-PowerManagement {
    try {
        Write-Host "Configuring power settings for 24/7 recording (Never Sleep)..." -ForegroundColor Cyan
        powercfg /change /standby-timeout-ac 0 | Out-Null
        powercfg /change /hibernate-timeout-ac 0 | Out-Null
        powercfg /hibernate off | Out-Null
        return $true
    }
    catch {
        Write-Host "Failed to configure power settings: $_" -ForegroundColor Red
        return $false
    }
}

function Set-WindowsServices {
    try {
        Write-Host "Disabling non-essential Windows services for performance..." -ForegroundColor Cyan
        $servicesToDisable = @(
            "DiagTrack",
            "dmwappushservice",
            "WSearch"
        )
        foreach ($service in $servicesToDisable) {
            Stop-Service -Name $service -ErrorAction SilentlyContinue
            Set-Service -Name $service -StartupType Disabled -ErrorAction SilentlyContinue
        }
        return $servicesToDisable.Count
    }
    catch {
        Write-Host "Error optimizing services: $_" -ForegroundColor Red
        return 0
    }
}

function Set-RemoteDesktopOptimization {
    try {
        Write-Host "Enabling and optimizing Remote Desktop..." -ForegroundColor Cyan
        Set-ItemProperty -Path "HKLM:\\System\\CurrentControlSet\\Control\\Terminal Server" -Name "fDenyTSConnections" -Value 0
        Enable-NetFirewallRule -DisplayGroup "Remote Desktop" | Out-Null
        return $true
    }
    catch {
        Write-Host "Failed to configure Remote Desktop: $_" -ForegroundColor Red
        return $false
    }
}

function Invoke-OptimizationMenu {
    Write-Host "`n=== SYSTEM PERFORMANCE OPTIMIZATION ===" -ForegroundColor Yellow
    Write-Host "1. Configure power settings" -ForegroundColor White
    Write-Host "2. Disable non-essential services" -ForegroundColor White
    Write-Host "3. Enable and optimize Remote Desktop" -ForegroundColor White
    Write-Host "4. Run full optimization (power + services + RDP)" -ForegroundColor White
    Write-Host "5. Back to main menu" -ForegroundColor White

    $choice = Read-Host "Select an option (1-5)"
    switch ($choice) {
        "1" {
            if (Set-PowerManagement) {
                $Global:logContent += "Power Management: Never Sleep Mode activated for 24/7 recording`r`n"
            }
            Read-Host "Press Enter to return..." | Out-Null
        }
        "2" {
            $disabledCount = Set-WindowsServices
            $Global:logContent += "Windows Services: $disabledCount services disabled`r`n"
            Read-Host "Press Enter to return..." | Out-Null
        }
        "3" {
            if (Set-RemoteDesktopOptimization) {
                $Global:logContent += "Remote Desktop: ENABLED and optimized`r`n"
            }
            Read-Host "Press Enter to return..." | Out-Null
        }
        "4" {
            $powerResult = Set-PowerManagement
            $servicesDisabled = Set-WindowsServices
            $rdpResult = Set-RemoteDesktopOptimization
            $Global:logContent += "FULL OPTIMIZATION: Power $(if($powerResult){'optimized'}else{'failed'}), $servicesDisabled services disabled, RDP $(if($rdpResult){'enabled'}else{'failed'})`r`n"
            Read-Host "Press Enter to return..." | Out-Null
        }
        default { return }
    }
}

Export-ModuleMember -Function *
