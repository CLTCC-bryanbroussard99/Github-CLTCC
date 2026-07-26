# Ensure script runs as Administrator
If (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "Please run this script as an Administrator!"
    Exit
}


# --- CONFIGURATION FOR EMAIL NOTIFICATIONS ---
$EnableEmail = $false  # Set to $true to turn on email alerts
$SMTPServer  = "://yourcompany.com"
$SMTPPort    = 587
$EmailTo     = "admin@yourcompany.com"
$EmailFrom   = "$env:COMPUTERNAME@yourcompany.com"
$EmailSubject= "Maintenance Report - $env:COMPUTERNAME"
# ---------------------------------------------

# Define log file path on desktop
$LogPath = "$env:USERPROFILE\Desktop\MaintenanceLog.txt"
"--- Maintenance Log Started: $(Get-Date) ---" | Out-File -FilePath $LogPath -Append

Function Log-Output ($Message, $Color = "Cyan") {
    Write-Host $Message -ForegroundColor $Color
    $Message | Out-File -FilePath $LogPath -Append
}

# 1. Disk Cleanup (Wipe Temporary Cache Files)
Log-Output "1. Starting temporary cache and disk cleanup..." "Yellow"
# Clean up user temp folder
Get-ChildItem -Path "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
# Clean up system temp folder
Get-ChildItem -Path "C:\Windows\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
# Run Windows Delivery Optimization and system cleanup
cleanmgr /sagerun:1 | Out-File -FilePath $LogPath -Append
Log-Output "Disk cleanup completed."

#DISM (/Restorehealth): Fixes the underlying Windows system image source files using Windows Update.
Write-Host "1. Starting DISM health restoration..." -ForegroundColor Cyan
Log-Output "1. Starting DISM health restoration..."
DISM.exe /Online /Cleanup-image /Restorehealth

#SFC (/scannow): Scans and repairs missing or broken individual protected operating system files.
Write-Host "2. Starting System File Checker (SFC) scan..." -ForegroundColor Cyan
Log-Output "2. Starting System File Checker (SFC) scan..."
sfc /scannow

# WinGet (upgrade --all): Automatically upgrades all installed software programs on your PC to their latest versions.
Write-Host "3. Updating apps via WinGet..." -ForegroundColor Cyan
Log-Output "3. Updating apps via WinGet..."
winget upgrade --all --accept-source-agreements --accept-package-agreements

Write-Host "All maintenance tasks are complete!" -ForegroundColor Green

# Finalize log and restart
Log-Output "All tasks complete. Saving log and restarting PC in 60 seconds..." "Green"
"--- Maintenance Log Finished: $(Get-Date) ---" | Out-File -FilePath $LogPath -Append

shutdown /r /t 60 /c "System maintenance complete. Restarting in 60 seconds."


