<#
   NTP Recovery Script - Fix v1.9x XProtect-Prepare Bug
   By Ulf Holmstrom, Happy Problem Solver at Manvarg AB (2026)
   For questions contact: ulf@manvarg.se

   PURPOSE:
   The XProtect-Prepare v1.9x scripts set NTP peers with incorrect format:
     BROKEN: "0.se.pool.ntp.org,1.se.pool.ntp.org,2.se.pool.ntp.org,3.se.pool.ntp.org"
     FIXED:  "0.se.pool.ntp.org,0x9 1.se.pool.ntp.org,0x9 2.se.pool.ntp.org,0x9 3.se.pool.ntp.org,0x9"

   The ,0x9 flag tells Windows Time Service to use NTP client mode.
   Without it, w32time cannot parse the peer list and time sync fails silently.

   This script can be deployed to all affected machines to fix the issue.

   USAGE:
   Run as Administrator on each affected machine:
     Set-ExecutionPolicy RemoteSigned -Scope Process -Force
     .\NTP-Recovery.ps1
#>

# Check for Administrator privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "ERROR: This script must be run as Administrator." -ForegroundColor Red
    Read-Host "Press Enter to exit..."
    exit 1
}

Write-Host "=== NTP Recovery Script ===" -ForegroundColor Cyan
Write-Host "Fixes the v1.9x XProtect-Prepare NTP peer format bug" -ForegroundColor Cyan
Write-Host ""

# Check current NTP configuration
$regPath = "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters"
try {
    $currentPeers = (Get-ItemProperty -Path $regPath -Name "NtpServer" -ErrorAction Stop).NtpServer
    $currentType = (Get-ItemProperty -Path $regPath -Name "Type" -ErrorAction Stop).Type
} catch {
    Write-Host "Could not read current NTP configuration." -ForegroundColor Red
    Write-Host "Windows Time service may not be configured." -ForegroundColor Yellow
    $currentPeers = ""
    $currentType = "unknown"
}

Write-Host "Current configuration:" -ForegroundColor White
Write-Host "  NtpServer: $currentPeers" -ForegroundColor Gray
Write-Host "  Type: $currentType" -ForegroundColor Gray
Write-Host ""

# Detect the bug
$bugDetected = $false
if ($currentPeers -and $currentPeers -match "," -and $currentPeers -notmatch ",0x") {
    Write-Host "[BUG DETECTED] NTP peers are missing ,0x9 flags!" -ForegroundColor Red
    Write-Host "This causes time synchronization to fail silently." -ForegroundColor Red
    $bugDetected = $true
} elseif ($currentPeers -match ",0x") {
    Write-Host "[OK] NTP peer format appears correct." -ForegroundColor Green
    Write-Host "This machine may not need the fix." -ForegroundColor Yellow
} else {
    Write-Host "[INFO] NTP may not be configured yet." -ForegroundColor Yellow
}

Write-Host ""
$confirm = Read-Host "Apply NTP fix? (Y/N)"
if ($confirm -notmatch "^[Yy]") {
    Write-Host "Cancelled." -ForegroundColor Yellow
    exit 0
}

# Apply the fix
Write-Host ""
Write-Host "Applying fix..." -ForegroundColor Cyan

try {
    # Stop service
    Write-Host " - Stopping Windows Time service..." -ForegroundColor Cyan
    Stop-Service w32time -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2

    # Set correct peer format
    $ntpServers = "0.se.pool.ntp.org,0x9 1.se.pool.ntp.org,0x9 2.se.pool.ntp.org,0x9 3.se.pool.ntp.org,0x9"
    Write-Host " - Setting correct NTP peer format..." -ForegroundColor Cyan
    Set-ItemProperty -Path $regPath -Name "NtpServer" -Value $ntpServers -Type String -ErrorAction Stop

    # Auto-detect domain membership
    $isDomain = (Get-CimInstance Win32_ComputerSystem).PartOfDomain
    if ($isDomain) {
        Write-Host " - Domain member detected - setting AllSync mode" -ForegroundColor Cyan
        Set-ItemProperty -Path $regPath -Name "Type" -Value "AllSync" -Type String -ErrorAction Stop
        $syncFlags = "ALL"
    } else {
        Write-Host " - Standalone server - setting NTP mode" -ForegroundColor Cyan
        Set-ItemProperty -Path $regPath -Name "Type" -Value "NTP" -Type String -ErrorAction Stop
        $syncFlags = "manual"
    }

    # Set startup to Automatic
    Write-Host " - Setting w32time to Automatic startup..." -ForegroundColor Cyan
    Set-Service w32time -StartupType Automatic

    # Apply via w32tm
    Write-Host " - Applying w32tm configuration..." -ForegroundColor Cyan
    & w32tm /config /manualpeerlist:$ntpServers /syncfromflags:$syncFlags /reliable:yes /update 2>&1 | Out-Null

    # Fix AnnounceFlags AFTER w32tm (reliable:yes resets it to 5)
    Write-Host " - Setting AnnounceFlags to 10 (server + client)..." -ForegroundColor Cyan
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Config" -Name "AnnounceFlags" -Value 10 -Type DWord -ErrorAction SilentlyContinue

    # Start service
    Write-Host " - Starting Windows Time service..." -ForegroundColor Cyan
    Start-Service w32time -ErrorAction Stop
    Start-Sleep -Seconds 3

    # Force resync
    Write-Host " - Forcing time resynchronization..." -ForegroundColor Cyan
    & w32tm /resync /rediscover 2>&1 | Out-Null
    Start-Sleep -Seconds 2

    # Verify
    $newPeers = (Get-ItemProperty -Path $regPath -Name "NtpServer" -ErrorAction SilentlyContinue).NtpServer
    $newType = (Get-ItemProperty -Path $regPath -Name "Type" -ErrorAction SilentlyContinue).Type
    $serviceStatus = (Get-Service w32time -ErrorAction SilentlyContinue).Status

    # Ensure DST (Daylight Saving Time) is enabled
    Write-Host " - Checking Daylight Saving Time settings..." -ForegroundColor Cyan
    $dstDisabled = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\TimeZoneInformation" -Name "DynamicDaylightTimeDisabled" -ErrorAction SilentlyContinue
    if ($dstDisabled.DynamicDaylightTimeDisabled -eq 1) {
        Write-Host " - WARNING: DST was DISABLED - enabling now..." -ForegroundColor Yellow
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\TimeZoneInformation" -Name "DynamicDaylightTimeDisabled" -Value 0 -Type DWord
        Write-Host " - DST re-enabled" -ForegroundColor Green
    } else {
        Write-Host " - DST is enabled (OK)" -ForegroundColor Green
    }

    # Re-apply timezone to force DST recalculation
    $currentTZ = (Get-TimeZone).Id
    Write-Host " - Re-applying timezone ($currentTZ) to force DST update..." -ForegroundColor Cyan
    Set-TimeZone -Id $currentTZ

    # Verify current time
    $isDST = [System.TimeZoneInfo]::Local.IsDaylightSavingTime((Get-Date))
    $utcOffset = (Get-TimeZone).GetUtcOffset((Get-Date))
    Write-Host " - Current UTC offset: +$($utcOffset.Hours)h (DST active: $isDST)" -ForegroundColor $(if($isDST){'Green'}else{'Yellow'})

    Write-Host ""
    Write-Host "=== NTP RECOVERY COMPLETE ===" -ForegroundColor Green
    Write-Host "  Service: $serviceStatus" -ForegroundColor Green
    Write-Host "  NtpServer: $newPeers" -ForegroundColor Green
    Write-Host "  Type: $newType" -ForegroundColor Green
    Write-Host "  Startup: Automatic" -ForegroundColor Green
    Write-Host "  DST: $(if($isDST){'Active (summer time)'}else{'Inactive (winter time)'})" -ForegroundColor Green
    Write-Host "  UTC Offset: +$($utcOffset.Hours)h" -ForegroundColor Green
    Write-Host ""
    Write-Host "Verify with: w32tm /query /status" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "========================================================" -ForegroundColor Red
    Write-Host "  IMPORTANT: A REBOOT IS REQUIRED!" -ForegroundColor Red
    Write-Host "========================================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "  The Windows Time service caches its internal state." -ForegroundColor Yellow
    Write-Host "  Without a reboot, it may continue using the old" -ForegroundColor Yellow
    Write-Host "  (broken) configuration and refuse to sync." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Please reboot this machine as soon as possible." -ForegroundColor Yellow
    Write-Host "  After reboot, verify with: w32tm /query /source" -ForegroundColor Yellow
    Write-Host "  It should show a .se.pool.ntp.org address." -ForegroundColor Yellow
    Write-Host ""
    $rebootNow = Read-Host "  Reboot now? (Y/N)"
    if ($rebootNow -eq 'Y' -or $rebootNow -eq 'y') {
        Write-Host "  Rebooting in 30 seconds..." -ForegroundColor Cyan
        shutdown /r /t 30 /c "NTP Recovery - reboot required for time sync fix"
    } else {
        Write-Host "  Remember to reboot this machine!" -ForegroundColor Red
    }
    Write-Host ""

} catch {
    Write-Host ""
    Write-Host "[FAIL] NTP recovery failed: $($_.Exception.Message)" -ForegroundColor Red
    Start-Service w32time -ErrorAction SilentlyContinue
}

Read-Host "Press Enter to exit..."
