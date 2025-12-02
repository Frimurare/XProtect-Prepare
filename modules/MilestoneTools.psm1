<#
    XProtect-Prepare-v2 Milestone Tools Module
    Authored by Ulf Holmström, ex employee, Solution Engineer at Milestone Systems December 2025
#>

function Install-MilestonePSTools {
    try {
        Write-Host "Installing Milestone PowerShell Tools from PowerShell Gallery..." -ForegroundColor Cyan
        Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted -ErrorAction SilentlyContinue
        Install-Module -Name 'MilestonePSTools' -Force -AllowClobber -ErrorAction Stop
        Write-Host "Milestone PowerShell Tools installed successfully." -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Failed to install Milestone PowerShell Tools: $_" -ForegroundColor Red
        return $false
    }
}

function Invoke-MilestoneToolsMenu {
    Write-Host "`n=== MILESTONE POWERSHELL TOOLS ===" -ForegroundColor Yellow
    Write-Host "1. Install Milestone PowerShell Tools" -ForegroundColor White
    Write-Host "2. Back to main menu" -ForegroundColor White

    $choice = Read-Host "Select an option (1-2)"
    switch ($choice) {
        "1" {
            if (Install-MilestonePSTools) {
                $Global:logContent += "Milestone PowerShell Tools: Successfully installed from PowerShell Gallery`r`n"
            } else {
                $Global:logContent += "Milestone PowerShell Tools: Installation failed`r`n"
            }
            Read-Host "Press Enter to return..." | Out-Null
        }
        default { return }
    }
}

Export-ModuleMember -Function *
