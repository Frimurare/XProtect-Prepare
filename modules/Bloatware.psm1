<#
    XProtect-Prepare-v2 Bloatware Removal Module
    Authored by Ulf Holmström, ex employee, Solution Engineer at Milestone Systems December 2025
#>

function Remove-BloatwareApp {
    param([string]$PackageName)
    try {
        Get-AppxPackage -Name $PackageName -AllUsers | Remove-AppxPackage -ErrorAction SilentlyContinue
        Get-AppXProvisionedPackage -Online | Where-Object { $_.PackageName -like "$PackageName*" } | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

function Remove-OneDrive {
    try {
        Write-Host "Removing OneDrive..." -ForegroundColor Cyan
        $onedriveSetup = Join-Path $env:SystemRoot "System32\\OneDriveSetup.exe"
        if (Test-Path $onedriveSetup) {
            Start-Process $onedriveSetup "/uninstall" -Wait
        }
        Stop-Process -Name OneDrive -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        Remove-Item -Recurse -Force "$env:LOCALAPPDATA\\Microsoft\\OneDrive" -ErrorAction SilentlyContinue
        Remove-Item -Recurse -Force "$env:ProgramData\\Microsoft OneDrive" -ErrorAction SilentlyContinue
        Remove-Item -Recurse -Force "$env:SystemDrive\\OneDriveTemp" -ErrorAction SilentlyContinue
        return $true
    }
    catch {
        return $false
    }
}

function Remove-MicrosoftTeams {
    $teamsRemoved = $false
    try {
        Write-Host "Removing Microsoft Teams (new app)..." -ForegroundColor Cyan
        $teamsRemoved = Remove-BloatwareApp -PackageName "MSTeams"
    }
    catch {}

    try {
        Write-Host "Removing Microsoft Teams (classic)..." -ForegroundColor Cyan
        $uninstallCmd = Get-ChildItem -Path "$env:LOCALAPPDATA\\Microsoft\\Teams\\Update.exe" -ErrorAction SilentlyContinue
        if ($uninstallCmd) {
            & $uninstallCmd.FullName --uninstall -s | Out-Null
            $teamsRemoved = $true
        }
    }
    catch {}

    return $teamsRemoved
}

function Remove-OutlookApp {
    try {
        Write-Host "Removing Outlook App (new Windows Outlook)..." -ForegroundColor Cyan
        return Remove-BloatwareApp -PackageName "OutlookForWindows"
    }
    catch {
        return $false
    }
}

function Remove-OfficeHub {
    try {
        Write-Host "Removing Office Hub / 365 installer..." -ForegroundColor Cyan
        return Remove-BloatwareApp -PackageName "Microsoft.MicrosoftOfficeHub"
    }
    catch {
        return $false
    }
}

function Remove-XboxApps {
    $xboxPackages = @(
        "Microsoft.Xbox.TCUI",
        "Microsoft.XboxApp",
        "Microsoft.XboxGameOverlay",
        "Microsoft.XboxGamingOverlay",
        "Microsoft.XboxIdentityProvider",
        "Microsoft.XboxSpeechToTextOverlay",
        "Microsoft.GamingApp"
    )

    $removedCount = 0
    foreach ($package in $xboxPackages) {
        if (Remove-BloatwareApp -PackageName $package) {
            $removedCount++
            Write-Host "Removed: $package" -ForegroundColor Green
        }
    }

    return $removedCount
}

function Invoke-BloatwareMenu {
    Write-Host "`n=== WINDOWS BLOATWARE REMOVAL ===" -ForegroundColor Yellow
    Write-Host "1. Remove standard bloatware" -ForegroundColor White
    Write-Host "2. Remove OneDrive" -ForegroundColor White
    Write-Host "3. Remove Teams, Outlook App, and Office Hub" -ForegroundColor White
    Write-Host "4. Remove Xbox apps (optional)" -ForegroundColor White
    Write-Host "5. Run full cleanup (standard + OneDrive + Teams/Outlook/Office)" -ForegroundColor White
    Write-Host "6. Back to main menu" -ForegroundColor White

    $choice = Read-Host "Select an option (1-6)"
    switch ($choice) {
        "1" {
            $bloatwareApps = @(
                @{ Name = "Microsoft.3DBuilder"; Display = "3D Builder" },
                @{ Name = "Microsoft.XboxApp"; Display = "Xbox App" },
                @{ Name = "Microsoft.XboxGameOverlay"; Display = "Xbox Game Overlay" },
                @{ Name = "Microsoft.XboxGamingOverlay"; Display = "Xbox Gaming Overlay" },
                @{ Name = "Microsoft.XboxIdentityProvider"; Display = "Xbox Identity Provider" },
                @{ Name = "Microsoft.XboxSpeechToTextOverlay"; Display = "Xbox Speech To Text" },
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

            foreach ($app in $bloatwareApps) {
                if (Remove-BloatwareApp -PackageName $app.Name) {
                    Write-Host "Removed: $($app.Display)" -ForegroundColor Green
                    $Global:logContent += "Bloatware removed: $($app.Display)`r`n"
                }
            }
            Read-Host "Press Enter to return..." | Out-Null
        }
        "2" {
            if (Remove-OneDrive) {
                Write-Host "OneDrive removed." -ForegroundColor Green
                $Global:logContent += "OneDrive: Successfully removed`r`n"
            }
            Read-Host "Press Enter to return..." | Out-Null
        }
        "3" {
            $teamsResult = Remove-MicrosoftTeams
            $outlookResult = Remove-OutlookApp
            $officeResult = Remove-OfficeHub
            $Global:logContent += "Microsoft Teams: $(if($teamsResult){'Removed'}else{'Not found'})`r`n"
            $Global:logContent += "Outlook App: $(if($outlookResult){'Removed'}else{'Not found'})`r`n"
            $Global:logContent += "Office Hub: $(if($officeResult){'Removed'}else{'Not found'})`r`n"
            Read-Host "Press Enter to return..." | Out-Null
        }
        "4" {
            $removedCount = Remove-XboxApps
            $Global:logContent += "Xbox apps removed: $removedCount`r`n"
            Read-Host "Press Enter to return..." | Out-Null
        }
        "5" {
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
                    Write-Host "Removed: $($app.Display)" -ForegroundColor Green
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
            Read-Host "Press Enter to return..." | Out-Null
        }
        default { return }
    }
}

Export-ModuleMember -Function *
