<#
    XProtect-Prepare-v2 Complete Deployment Module
    Authored by Ulf Holmström, ex employee, Solution Engineer at Milestone Systems December 2025
#>

function Invoke-CompleteDeployment {
    Write-Host "`n=== COMPLETE XPROTECT PREPARE DEPLOYMENT ===" -ForegroundColor Yellow
    Write-Host "This will run all configurations automatically, including SNMP." -ForegroundColor Cyan

    New-RegistryBackup -BackupName "CompleteDeployment"

    Write-Host "Ensuring Windows Update is enabled for prerequisite installations..." -ForegroundColor Cyan
    Set-WindowsUpdateStatus -Enable $true | Out-Null
    $Global:logContent += "Windows Updates: Enabled for deployment prerequisites`r`n"

    $drives = Get-AvailableDrives
    $storageDrives = $drives | Where-Object { $_.DeviceID -ne "C:" } | Select-Object -ExpandProperty DeviceID
    if (-not $storageDrives) {
        Write-Host "No storage drives found beyond OS drive. Antivirus exclusions will target processes only." -ForegroundColor Yellow
    }

    Add-XProtectAntivirusExclusions -SelectedDrives $storageDrives
    Test-StorageDriveHealth -SelectedDrives $storageDrives

    Write-Host "Removing Windows bloatware..." -ForegroundColor Cyan
    $bloatwareApps = @(
        @{ Name = "Microsoft.3DBuilder"; Display = "3D Builder" },
        @{ Name = "Microsoft.GetHelp"; Display = "Get Help" },
        @{ Name = "Microsoft.Getstarted"; Display = "Get Started" },
        @{ Name = "Microsoft.ZuneMusic"; Display = "Groove Music" },
        @{ Name = "Microsoft.ZuneVideo"; Display = "Movies & TV" },
        @{ Name = "Microsoft.MinecraftUWP"; Display = "Minecraft" },
        @{ Name = "Microsoft.MicrosoftSolitaireCollection"; Display = "Solitaire" },
        @{ Name = "Microsoft.People"; Display = "People" },
        @{ Name = "Microsoft.BingWeather"; Display = "Weather" },
        @{ Name = "Microsoft.WindowsMaps"; Display = "Maps" },
        @{ Name = "Microsoft.Messaging"; Display = "Messaging" },
        @{ Name = "Microsoft.SkypeApp"; Display = "Skype" },
        @{ Name = "Microsoft.Office.Sway"; Display = "Sway" },
        @{ Name = "Microsoft.WindowsFeedbackHub"; Display = "Feedback Hub" },
        @{ Name = "Microsoft.YourPhone"; Display = "Your Phone" },
        @{ Name = "Microsoft.BingNews"; Display = "News" },
        @{ Name = "Microsoft.Microsoft3DViewer"; Display = "3D Viewer" },
        @{ Name = "Microsoft.MicrosoftOfficeHub"; Display = "Office Hub" },
        @{ Name = "MicrosoftTeams"; Display = "Teams (AppX)" }
    )
    $removedCount = 0
    foreach ($app in $bloatwareApps) {
        if (Remove-BloatwareApp -PackageName $app.Name) {
            $removedCount++
            $Global:logContent += "Bloatware removed: $($app.Display)`r`n"
        }
    }

    if (Remove-OneDrive) {
        $Global:logContent += "OneDrive: Successfully removed`r`n"
    }

    $teamsResult = Remove-MicrosoftTeams
    $outlookResult = Remove-OutlookApp
    $officeResult = Remove-OfficeHub

    $Global:logContent += "Microsoft Teams: $(if($teamsResult){'Removed'}else{'Not found'})`r`n"
    $Global:logContent += "Outlook App: $(if($outlookResult){'Removed'}else{'Not found'})`r`n"
    $Global:logContent += "Office Hub: $(if($officeResult){'Removed'}else{'Not found'})`r`n"
    $Global:logContent += "Full cleanup completed: $removedCount apps, OneDrive, Teams, Outlook, Office Hub removed`r`n"

    Write-Host "Applying performance optimizations..." -ForegroundColor Cyan
    $powerResult = Set-PowerManagement
    $servicesDisabled = Set-WindowsServices
    $rdpResult = Set-RemoteDesktopOptimization
    $Global:logContent += "FULL OPTIMIZATION: Power $(if($powerResult){'optimized'}else{'failed'}), $servicesDisabled services disabled, RDP $(if($rdpResult){'enabled'}else{'failed'})`r`n"

    Write-Host "Configuring NTP time server and taskbar seconds..." -ForegroundColor Cyan
    $ntpResult = Enable-NTPServer
    $secondsResult = Enable-TaskbarSeconds
    if ($ntpResult -and $secondsResult) {
        $Global:logContent += "Time services: NTP server configured and taskbar seconds enabled`r`n"
    }

    Write-Host "Installing Milestone PowerShell Tools..." -ForegroundColor Cyan
    if (Install-MilestonePSTools) {
        $Global:logContent += "Milestone PowerShell Tools: Successfully installed from PowerShell Gallery`r`n"
    } else {
        $Global:logContent += "Milestone PowerShell Tools: Installation failed`r`n"
    }

    Write-Host "Installing SNMP capabilities..." -ForegroundColor Cyan
    if (Enable-SNMP) {
        $Global:logContent += "SNMP: Client and WMI Provider installed successfully`r`n"
    } else {
        $Global:logContent += "SNMP: Installation failed`r`n"
    }

    Write-Host "Leaving Windows Updates ENABLED for security." -ForegroundColor Green
    $Global:logContent += "Windows Updates: Left enabled after deployment`r`n"

    Write-Host "`nComplete deployment finished. Generate a log from the main menu to save details." -ForegroundColor Green
    Play-Fanfare
    Read-Host "Press Enter to return to main menu..." | Out-Null
}

Export-ModuleMember -Function *
