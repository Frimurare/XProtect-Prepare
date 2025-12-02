<#
    XProtect-Prepare-v2 Network and Time Services Module
    Authored by Ulf Holmström, ex employee, Solution Engineer at Milestone Systems December 2025
#>

function Enable-SNMP {
    try {
        Write-Host "Installing SNMP Client and WMI Provider..." -ForegroundColor Cyan
        Add-WindowsCapability -Online -Name "SNMP.Client~~~~0.0.1.0" -ErrorAction Stop | Out-Null
        Add-WindowsCapability -Online -Name "SNMP.WMI~~~~0.0.1.0" -ErrorAction Stop | Out-Null
        Write-Host "SNMP capabilities installed successfully." -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "SNMP installation failed: $_" -ForegroundColor Red
        return $false
    }
}

function Enable-NTPServer {
    try {
        Write-Host "`n=== CONFIGURING NTP TIME SERVER FOR CAMERAS ===" -ForegroundColor Yellow
        Write-Host "This configures Windows Time Service to act as an NTP server" -ForegroundColor Cyan

        Stop-Service w32time -Force -ErrorAction SilentlyContinue

        Write-Host " - Setting Swedish NTP pool servers as time source..." -ForegroundColor Cyan
        w32tm /config /syncfromflags:manual /manualpeerlist:"0.se.pool.ntp.org 1.se.pool.ntp.org 2.se.pool.ntp.org 3.se.pool.ntp.org" /reliable:yes /update | Out-Null
        Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Services\\W32Time\\Parameters" -Name "Type" -Value "NTP" -Type String -ErrorAction Stop
        Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Services\\W32Time\\Parameters" -Name "NtpServer" -Value "0.se.pool.ntp.org,0x8 1.se.pool.ntp.org,0x8 2.se.pool.ntp.org,0x8 3.se.pool.ntp.org,0x8" -Type String -ErrorAction Stop

        Write-Host " - Opening firewall for NTP (UDP port 123)..." -ForegroundColor Cyan
        netsh advfirewall firewall delete rule name="NTP Server (XProtect)" 2>&1 | Out-Null
        netsh advfirewall firewall add rule name="NTP Server (XProtect)" dir=in action=allow protocol=UDP localport=123 2>&1 | Out-Null

        Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Services\\W32Time\\TimeProviders\\NtpServer" -Name "Enabled" -Value 1 -Type DWord -ErrorAction Stop
        Start-Service w32time -ErrorAction Stop
        w32tm /config /update | Out-Null
        w32tm /resync /force | Out-Null

        Write-Host "`n - Verifying NTP server configuration..." -ForegroundColor Cyan
        $timeService = Get-Service w32time -ErrorAction SilentlyContinue
        if ($timeService.Status -eq 'Running') {
            Write-Host "NTP Server Status: ENABLED" -ForegroundColor Green
            Write-Host "Time Source: Swedish NTP Pool (se.pool.ntp.org)" -ForegroundColor Green
            Write-Host "Server Mode: Reliable NTP server for camera synchronization" -ForegroundColor Green
        }

        $Global:logContent += "NTP Server: Configured successfully for camera time synchronization`r`n"
        return $true
    }
    catch {
        Write-Host "[FAIL] Failed to configure NTP server: $($_.Exception.Message)" -ForegroundColor Red
        Start-Service w32time -ErrorAction SilentlyContinue
        $Global:logContent += "NTP Server: Configuration failed or incomplete`r`n"
        return $false
    }
}

function Enable-TaskbarSeconds {
    try {
        Write-Host "Enabling taskbar seconds for clear time visibility..." -ForegroundColor Cyan
        $regPath = "HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced"
        Set-ItemProperty -Path $regPath -Name "ShowSecondsInSystemClock" -Value 1 -Type DWord -ErrorAction Stop
        Write-Host "Taskbar seconds enabled. Sign out and back in to see the change." -ForegroundColor Green
        $Global:logContent += "Taskbar seconds enabled for Windows clock`r`n"
        return $true
    }
    catch {
        Write-Host "Failed to enable taskbar seconds: $_" -ForegroundColor Red
        return $false
    }
}

function Invoke-NetworkMenu {
    Write-Host "`n=== NETWORK & TIME SERVICES ===" -ForegroundColor Yellow
    Write-Host "1. Configure NTP Time Server" -ForegroundColor White
    Write-Host "2. Enable taskbar seconds" -ForegroundColor White
    Write-Host "3. Configure NTP and enable taskbar seconds" -ForegroundColor White
    Write-Host "4. Install SNMP capabilities" -ForegroundColor White
    Write-Host "5. Back to main menu" -ForegroundColor White

    $choice = Read-Host "Select an option (1-5)"
    switch ($choice) {
        "1" { Enable-NTPServer | Out-Null; Read-Host "Press Enter to return..." | Out-Null }
        "2" { Enable-TaskbarSeconds | Out-Null; Read-Host "Press Enter to return..." | Out-Null }
        "3" {
            $ntpResult = Enable-NTPServer
            $secondsResult = Enable-TaskbarSeconds
            if ($ntpResult -and $secondsResult) {
                $Global:logContent += "Time services: NTP server configured and taskbar seconds enabled`r`n"
            }
            Read-Host "Press Enter to return..." | Out-Null
        }
        "4" {
            $snmpResult = Enable-SNMP
            if ($snmpResult) {
                $Global:logContent += "SNMP: Client and WMI Provider installed successfully`r`n"
            } else {
                $Global:logContent += "SNMP: Installation failed`r`n"
            }
            Read-Host "Press Enter to return..." | Out-Null
        }
        default { return }
    }
}

Export-ModuleMember -Function *
