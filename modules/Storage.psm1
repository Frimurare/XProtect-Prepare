<#
    XProtect-Prepare-v2 Storage and Antivirus Module
    Authored by Ulf Holmström, ex employee, Solution Engineer at Milestone Systems December 2025
#>

function Get-AvailableDrives {
    Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 } | Select-Object DeviceID, @{Name="Size(GB)";Expression={[math]::Round($_.Size/1GB,2)}}, @{Name="FreeSpace(GB)";Expression={[math]::Round($_.FreeSpace/1GB,2)}}
}

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

function Test-DriveIndexing {
    param([string]$Drive)
    try {
        $driveLetter = $Drive.TrimEnd('\\')
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

function Disable-DriveIndexing {
    param([string]$Drive)
    try {
        $driveLetter = $Drive.TrimEnd('\\')
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

function Add-XProtectAntivirusExclusions {
    param(
        [string[]]$SelectedDrives
    )

    Write-Host "`n--- Adding Antivirus Exceptions ---" -ForegroundColor Yellow
    foreach ($drive in $SelectedDrives) {
        if ($drive -eq "C:") {
            Write-Host "Skipping antivirus exception for OS drive C:" -ForegroundColor Yellow
            continue
        }
        $drivePath = $drive + "\\"
        try {
            Add-MpPreference -ExclusionPath $drivePath -ErrorAction Stop
            Write-Host "Antivirus exception added for drive ${drivePath}" -ForegroundColor Green
            $Global:logContent += "Antivirus exception added for drive ${drivePath}`r`n"
        }
        catch {
            Write-Host "Error adding antivirus exception for drive ${drivePath}: $_" -ForegroundColor Red
            $Global:logContent += "Error adding antivirus exception for drive ${drivePath}: $_`r`n"
        }
    }

    $exclusionExtensions = @("blk", "idx")
    foreach ($ext in $exclusionExtensions) {
        try {
            Add-MpPreference -ExclusionExtension $ext -ErrorAction Stop
            Write-Host "Antivirus exception added for file extension: ${ext}" -ForegroundColor Green
            $Global:logContent += "Antivirus exception added for file extension: ${ext}`r`n"
        }
        catch {
            Write-Host "Error adding antivirus exception for file extension ${ext}: $_" -ForegroundColor Red
            $Global:logContent += "Error adding antivirus exception for file extension ${ext}: $_`r`n"
        }
    }

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
            Write-Host "Antivirus exception added for process: ${proc}" -ForegroundColor Green
            $Global:logContent += "Antivirus exception added for process: ${proc}`r`n"
        }
        catch {
            Write-Host "Error adding antivirus exception for process ${proc}: $_" -ForegroundColor Red
            $Global:logContent += "Error adding antivirus exception for process ${proc}: $_`r`n"
        }
    }
}

function Test-StorageDriveHealth {
    param([string[]]$SelectedDrives)

    Write-Host "`n--- Storage Drive Configuration Check (Storage Drives Only) ---" -ForegroundColor Yellow
    foreach ($drive in $SelectedDrives) {
        $driveForCheck = $drive.TrimEnd(':') + ":"
        $drivePath = $driveForCheck + "\\"

        if ($driveForCheck.ToUpper() -eq "C:") {
            Write-Host "Skipping OS drive ${driveForCheck} - storage optimization not needed" -ForegroundColor Gray
            continue
        }

        Write-Host "`nChecking storage drive ${driveForCheck}..." -ForegroundColor Cyan

        $blockSize = Test-DriveBlockSize $drivePath
        if ($blockSize -eq 65536) {
            Write-Host "Storage Drive ${driveForCheck}: Correct block size - 64 KB." -ForegroundColor Green
            $Global:logContent += "Storage Drive ${driveForCheck}: Correct block size - 64 KB.`r`n"
        } elseif ($blockSize) {
            $blockSizeKB = $blockSize / 1024
            Write-Host "Storage Drive ${driveForCheck}: Incorrect block size - ${blockSizeKB} KB. Expected 64 KB." -ForegroundColor Red
            Write-Host "NOTE: Drive needs to be reformatted with 64KB allocation unit size." -ForegroundColor Yellow
            $Global:logContent += "Storage Drive ${driveForCheck}: Incorrect block size - ${blockSizeKB} KB.`r`n"
        } else {
            Write-Host "Storage Drive ${driveForCheck}: Could not determine block size." -ForegroundColor Yellow
            $Global:logContent += "Storage Drive ${driveForCheck}: Block size check failed.`r`n"
        }

        $indexingEnabled = Test-DriveIndexing $drivePath
        if ($indexingEnabled -eq $false) {
            Write-Host "Storage Drive ${driveForCheck}: Indexing is OFF." -ForegroundColor Green
            $Global:logContent += "Storage Drive ${driveForCheck}: Indexing is OFF.`r`n"
        } elseif ($indexingEnabled -eq $true) {
            Write-Host "Storage Drive ${driveForCheck}: Indexing is ON. Should be disabled for storage drives." -ForegroundColor Yellow
            $disableChoice = Read-Host "Disable indexing for storage drive ${driveForCheck}? (Y/N)"
            if ($disableChoice -match "^[Yy]") {
                if (Disable-DriveIndexing $drivePath) {
                    Write-Host "Storage Drive ${driveForCheck}: Indexing disabled." -ForegroundColor Green
                    $Global:logContent += "Storage Drive ${driveForCheck}: Indexing disabled.`r`n"
                } else {
                    Write-Host "Storage Drive ${driveForCheck}: Failed to disable indexing." -ForegroundColor Red
                    $Global:logContent += "Storage Drive ${driveForCheck}: Failed to disable indexing.`r`n"
                }
            }
        }
    }
}

function Invoke-StorageAndAntivirus {
    Write-Host "`n=== ANTIVIRUS AND STORAGE CONFIGURATION ===" -ForegroundColor Yellow

    Write-Host "`nAvailable drives:" -ForegroundColor Cyan
    $drives = Get-AvailableDrives
    $drives | Format-Table -AutoSize

    Write-Host "Select drives for antivirus exceptions:" -ForegroundColor Green
    $driveInput = Read-Host "Enter drive letters (comma separated, e.g., d,e) or 'all' for all drives"

    if ($driveInput.ToLower() -eq 'all') {
        $selectedDrives = $drives.DeviceID
    } else {
        $selectedDrives = $driveInput -split "," | ForEach-Object { ($_.Trim().ToUpper() + ":").Replace("::", ":") }
    }

    Add-XProtectAntivirusExclusions -SelectedDrives $selectedDrives
    Test-StorageDriveHealth -SelectedDrives $selectedDrives

    Read-Host "`nAntivirus and Storage configuration complete. Press Enter to return to menu..." | Out-Null
}

function Invoke-StorageMenu {
    Write-Host "`n=== ANTIVIRUS & STORAGE MENU ===" -ForegroundColor Yellow
    Write-Host "1. Add antivirus exclusions" -ForegroundColor White
    Write-Host "2. Check block size and indexing" -ForegroundColor White
    Write-Host "3. Run full antivirus + storage workflow" -ForegroundColor White
    Write-Host "4. Back to main menu" -ForegroundColor White

    $choice = Read-Host "Select an option (1-4)"
    switch ($choice) {
        "1" {
            $drives = Get-AvailableDrives
            $drives | Format-Table -AutoSize
            $input = Read-Host "Enter drive letters (comma separated, e.g., d,e) or 'all' for all drives"
            $selected = if ($input.ToLower() -eq 'all') { $drives.DeviceID } else { $input -split "," | ForEach-Object { ($_.Trim().ToUpper() + ":").Replace("::", ":") } }
            Add-XProtectAntivirusExclusions -SelectedDrives $selected
            Read-Host "Press Enter to return..." | Out-Null
        }
        "2" {
            $drives = Get-AvailableDrives
            $drives | Format-Table -AutoSize
            $input = Read-Host "Enter drive letters (comma separated, e.g., d,e) or 'all' for all drives"
            $selected = if ($input.ToLower() -eq 'all') { $drives.DeviceID } else { $input -split "," | ForEach-Object { ($_.Trim().ToUpper() + ":").Replace("::", ":") } }
            Test-StorageDriveHealth -SelectedDrives $selected
            Read-Host "Press Enter to return..." | Out-Null
        }
        "3" { Invoke-StorageAndAntivirus }
        default { return }
    }
}

Export-ModuleMember -Function *
