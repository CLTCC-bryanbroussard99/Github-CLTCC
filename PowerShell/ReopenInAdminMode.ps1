# ==============================================================================
# SCRIPT RE-LAUNCHER: ENFORCE ADMINISTRATOR RIGHTS & TARGET NEWEST POWERSHELL ENGINE
# ==============================================================================

# 1. CHECK PRIVILEGES: Verify if the current session already has Administrator rights
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = [Security.Principal.WindowsPrincipal]$identity
$isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

# 2. CHECK ENGINE VERSION: Determine if we are currently running modern PowerShell 6+ (pwsh)
# Windows PowerShell is version 5.1 and below; PowerShell Core is 6, 7, etc.
$isModernPowerShell = $PSVersionTable.PSVersion.Major -ge 6

# 3. SCAN SYSTEM: Check if the modern PowerShell executable (pwsh.exe) exists on this machine
$hasPwshInstalled = [bool](Get-Command pwsh.exe -ErrorAction SilentlyContinue)


# ==============================================================================
# CONDITION EVALUATION
# Trigger relaunch if:
#   A) We are missing Administrator rights OR
#   B) We are stuck in legacy PowerShell (v5.1) but modern PowerShell (v7+) is installed
# ==============================================================================
if (-not $isAdmin -or (-not $isModernPowerShell -and $hasPwshInstalled)) {
    
    Write-Warning "Optimizing environment... Relaunching script with highest privileges and newest engine."

    # Decide which engine to target (prefer pwsh.exe over legacy powershell.exe)
    $targetEngine = if ($hasPwshInstalled) { "pwsh.exe" } else { "powershell.exe" }

    # Define execution arguments for the new window:
    #   -NoProfile: Prevents user profiles from loading (speeds up boot and prevents profile conflicts)
    #   -ExecutionPolicy Bypass: Bypasses execution restrictions so this script can run
    #   -File "$PSCommandPath": Dynamically targets the exact file path of THIS current script
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""

    # Execute the new process:
    #   -Verb RunAs: Triggers the Windows User Account Control (UAC) prompt to grant Admin rights
    Start-Process $targetEngine -ArgumentList $arguments -Verb RunAs

    # Kill the current, sub-optimal PowerShell session immediately
    Exit
}


# ==============================================================================
# ACTUAL SCRIPT LOGIC STARTS HERE
# This section is only reached if the session is confirmed Admin AND using the best engine
# ==============================================================================

Write-Host "--- SUCCESS ---" -ForegroundColor Green
Write-Host "Running with elevated Administrator privileges." -ForegroundColor Green
Write-Host "Active Engine Engine Version: $($PSVersionTable.PSVersion)" -ForegroundColor Cyan

# Keep the window open so you can see the success message
Pause
