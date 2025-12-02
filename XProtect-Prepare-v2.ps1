<#
    XProtect-Prepare-v2 Entry Script
    Authored by Ulf Holmström, ex employee, Solution Engineer at Milestone Systems December 2025
#>

$ErrorActionPreference = 'Continue'

$moduleRoot = Join-Path $PSScriptRoot 'modules'
$coreModule = Join-Path $moduleRoot 'Core.psm1'
Import-Module $coreModule -Force
Get-ChildItem $moduleRoot -Filter '*.psm1' | Where-Object { $_.Name -ne 'Core.psm1' } | ForEach-Object { Import-Module $_.FullName -Force }

Initialize-XProtectPrepare -Version "2.0"

if (-not (Test-Administrator)) {
    Write-Host "You must run this script as an Administrator. Please re-run with elevated privileges." -ForegroundColor Red
    Read-Host "Press Enter to exit..." | Out-Null
    exit
}

Show-Banner
Write-Host "This script configures Milestone XProtect prerequisites with modular menus." -ForegroundColor White
Write-Host "Author tag: $Global:AuthorTag" -ForegroundColor Cyan
Write-Host "Features: antivirus exclusions, storage checks, updates, bloatware removal, optimization, SNMP, NTP, and Milestone PowerShell Tools." -ForegroundColor White
Write-Host "Use the main menu to run individual areas or the complete deployment to run everything." -ForegroundColor White
Write-Host "" 
Read-Host "Press Enter to continue..." | Out-Null

$choice = ""
do {
    Show-Banner
    Show-MainMenu
    $choice = Read-Host "Enter your choice (1-9)"

    switch ($choice) {
        "1" { Invoke-StorageMenu }
        "2" { Invoke-WindowsUpdateMenu }
        "3" { Invoke-BloatwareMenu }
        "4" { Invoke-OptimizationMenu }
        "5" { Invoke-NetworkMenu }
        "6" { Invoke-MilestoneToolsMenu }
        "7" { Invoke-CompleteDeployment }
        "8" { Invoke-LogGeneration }
        "9" {
            Write-Host "Exiting XProtect-Prepare-v2. Thank you for using the script!" -ForegroundColor Green
            Write-Host "Authored by $Global:AuthorTag" -ForegroundColor Cyan
            break
        }
        default {
            Write-Host "Invalid choice. Please select 1-9." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
} while ($choice -ne "9")
