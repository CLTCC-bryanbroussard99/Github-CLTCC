# Ensure script runs silently as SYSTEM or Administrator
If (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Exit
}

# --- NON-INTERACTIVE CONFIGURATION ---
$EnableEmail   = $true  
$SMTPServer    = "smtp.college.edu"
$SMTPPort      = 587
$EmailTo       = "cybersecurity-admin@college.edu"
$EmailFrom     = "$env:COMPUTERNAME@college.edu"
$EmailSubject  = "Endpoint Maintenance Complete - $env:COMPUTERNAME"

# Log locally to ProgramData (accessible by SYSTEM account)
$LogPath = "C:\ProgramData\EndpointMaintenanceLog.txt"
"--- Maintenance Log Started: $(Get-Date) ---" | Out-File -FilePath $LogPath -Force

Function Log-Output ($Message) {
    "$Message" | Out-File -FilePath $LogPath -Append
}

# 1. Silent Cache Wipe
Log-Output "1. Wiping temporary cache files..."
Get-ChildItem -Path "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
Get-ChildItem -Path "C:\Windows\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
cleanmgr /sagerun:1 | Out-File -FilePath $LogPath -Append

# 2. DISM Scan
Log-Output "2. Running DISM Health Restoration..."
DISM.exe /Online /Cleanup-image /Restorehealth | Out-File -FilePath $LogPath -Append

# 3. SFC Scan
Log-Output "3. Running SFC File Verification..."
sfc /scannow | Out-File -FilePath $LogPath -Append

# 4. Silent WinGet Updates
Log-Output "4. Upgrading third-party software via WinGet..."
winget upgrade --all --silent --accept-source-agreements --accept-package-agreements | Out-File -FilePath $LogPath -Append

"--- Maintenance Log Finished: $(Get-Date) ---" | Out-File -FilePath $LogPath -Append

# 5. Automated Notification
if ($EnableEmail) {
    try {
        $Body = Get-Content -Path $LogPath -Raw
        Send-MailMessage -SmtpServer $SMTPServer -Port $SMTPPort -UseSsl -From $EmailFrom -To $EmailTo -Subject $EmailSubject -Body $Body
    } catch {
        "Email delivery failed: $_" | Out-File -FilePath $LogPath -Append
    }
}

# Force reboot immediately without interactive delays
shutdown /r /t 10 /f /c "Automated system maintenance complete. Rebooting endpoint."

