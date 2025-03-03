# tested on Windows 11 with VMware Workstation 17, powershell 7.5, admin rights are required

# Ensure script runs as Administrator
if (-not ([System.Security.Principal.WindowsPrincipal] [System.Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([System.Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "This script must be run as Administrator." -ForegroundColor Red
    Exit
}

# Get script location for logging
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogFile = Join-Path $ScriptDir "ActionsLog.txt"
$ErrorLogFile = Join-Path $ScriptDir "ErrorLog.txt"

# Function to log actions
function Write-Log {
    param ([string]$Message)
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$Timestamp - $Message" | Out-File -Append -FilePath $LogFile
}

# Function to log errors
function Write-ErrorLog {
    param ([string]$Message)
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$Timestamp - ERROR: $Message" | Out-File -Append -FilePath $ErrorLogFile
}

# Ask user for VMDK folder
$VMDKFolder = Read-Host "Enter the full path to the VMDK folder"
if (-Not (Test-Path $VMDKFolder)) {
    Write-Host "Invalid path or access denied: $VMDKFolder" -ForegroundColor Red
    Write-ErrorLog "Invalid path or access denied: $VMDKFolder"
    Exit
}

Write-Log "User provided VMDK folder: $VMDKFolder"

# Create a subfolder for old lock files
$LockBackupFolder = Join-Path $VMDKFolder "Old_Locks"
if (-Not (Test-Path $LockBackupFolder)) {
    try {
        New-Item -Path $LockBackupFolder -ItemType Directory -Force | Out-Null
        Write-Log "Created lock backup folder: $LockBackupFolder"
    } catch {
        Write-Host "Failed to create lock backup folder." -ForegroundColor Red
        Write-ErrorLog "Failed to create lock backup folder: $_"
        Exit
    }
}

# Move .lck files/folders into the subfolder
$LCKItems = Get-ChildItem -Path $VMDKFolder -Filter "*.lck" -Force -ErrorAction SilentlyContinue
if ($LCKItems.Count -gt 0) {
    Write-Log "Found $($LCKItems.Count) .lck files/folders in $VMDKFolder"
    foreach ($Item in $LCKItems) {
        $NewLocation = Join-Path $LockBackupFolder $Item.Name
        try {
            Move-Item -Path $Item.FullName -Destination $NewLocation -Force
            Write-Log "Moved: $($Item.FullName) -> $NewLocation"
        } catch {
            Write-Host "Failed to move $($Item.FullName)" -ForegroundColor Red
            Write-ErrorLog "Failed to move $($Item.FullName): $_"
        }
    }
} else {
    Write-Log "No .lck files found in $VMDKFolder"
}

# Stop all VMware processes
Write-Host "Stopping VMware processes..." -ForegroundColor Yellow
try {
    Get-Process | Where-Object { $_.Name -like "vmware*" } -ErrorAction Stop | Stop-Process -Force -ErrorAction Stop
    Write-Log "Successfully stopped all VMware processes."
} catch {
    Write-Host "Failed to stop VMware processes." -ForegroundColor Red
    Write-ErrorLog "Failed to stop VMware processes: $_"
}

# Restart VMware services (Ignore disabled ones, log all errors)
Write-Host "Restarting VMware services..." -ForegroundColor Yellow
try {
    $VmwareServices = Get-Service -Name "VMware*" -ErrorAction SilentlyContinue
    if (-Not $VmwareServices) {
        Write-ErrorLog "No VMware services found or unable to query VMware services."
        Write-Host "No VMware services found." -ForegroundColor Yellow
    } else {
        foreach ($Service in $VmwareServices) {
            if ($Service.StartType -eq "Disabled") {
                Write-Log "Skipping disabled service: $($Service.Name)"
                Continue
            }
            try {
                Restart-Service -Name $Service.Name -Force -ErrorAction Stop
                Write-Log "Restarted service: $($Service.Name)"
            } catch {
                Write-Host "Failed to restart service: $($Service.Name)" -ForegroundColor Red
                Write-ErrorLog "Failed to restart service: $($Service.Name) - $_"
            }
        }
    }
} catch {
    Write-Host "Failed to query VMware services." -ForegroundColor Red
    Write-ErrorLog "Failed to query VMware services: $_"
}

Write-Host "Script execution completed. Check logs in $ScriptDir" -ForegroundColor Green
