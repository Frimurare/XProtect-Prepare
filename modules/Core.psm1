<#
    XProtect-Prepare-v2 Core Functions
    Authored by Ulf Holmström, ex employee, Solution Engineer at Milestone Systems December 2025
#>

$Global:AuthorTag = "Ulf Holmström, ex employee, Solution Engineer at Milestone Systems December 2025"

function Test-Administrator {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

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

function Initialize-XProtectPrepare {
    param(
        [string]$Version = "2.0"
    )

    $Global:ScriptVersion = $Version
    $Global:logContent = ""
    $Global:LogFolderPath = "C:\\Milestonecheck-ulfh"
}

function Show-Banner {
    Write-Host "==========================================================" -ForegroundColor Cyan
    Write-Host "      XProtect-Prepare-v2 (Rev $Global:ScriptVersion)" -ForegroundColor White
    Write-Host "      Authored by $Global:AuthorTag" -ForegroundColor Cyan
    Write-Host "      For questions contact: ulf@manvarg.se" -ForegroundColor White
    Write-Host "==========================================================" -ForegroundColor Cyan
    Write-Host "" 
}

function Show-MainMenu {
    Write-Host "=== MAIN MENU ===" -ForegroundColor Yellow
    Write-Host "1. Antivirus & Storage" -ForegroundColor White
    Write-Host "2. Windows Update" -ForegroundColor White
    Write-Host "3. Bloatware Removal" -ForegroundColor White
    Write-Host "4. Performance Optimization" -ForegroundColor White
    Write-Host "5. Network & Time Services" -ForegroundColor White
    Write-Host "6. Milestone PowerShell Tools" -ForegroundColor White
    Write-Host "7. COMPLETE DEPLOYMENT (runs everything, including SNMP)" -ForegroundColor Yellow
    Write-Host "8. Generate Log File" -ForegroundColor White
    Write-Host "9. Exit" -ForegroundColor White
    Write-Host ""
}

function Invoke-LogGeneration {
    Write-Host "`n=== LOG FILE GENERATION ===" -ForegroundColor Yellow

    if ($Global:logContent -eq "") {
        Write-Host "No configuration changes have been made yet. Nothing to log." -ForegroundColor Yellow
        Read-Host "Press Enter to return to main menu..." | Out-Null
        return
    }

    $writeLog = Read-Host "Save a log file with configuration details? (Y/N)"
    if ($writeLog -match "^[Yy]") {
        if (-not (Test-Path $Global:LogFolderPath)) {
            New-Item -ItemType Directory -Path $Global:LogFolderPath | Out-Null
        }
        $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
        $logFilePath = Join-Path $Global:LogFolderPath "XProtect-Prepare-v2_${timestamp}.txt"
        $header = "XProtect-Prepare-v2 Log - $(Get-Date)`r`n$Global:AuthorTag`r`n===========================================`r`n"
        $fullLogContent = $header + $Global:logContent
        $fullLogContent | Out-File -FilePath $logFilePath -Encoding UTF8
        Write-Host "Log file saved at: $logFilePath" -ForegroundColor Green
        Play-Fanfare
    } else {
        Write-Host "No log file will be created." -ForegroundColor Yellow
    }

    Read-Host "Press Enter to return to main menu..." | Out-Null
}

Export-ModuleMember -Function *
