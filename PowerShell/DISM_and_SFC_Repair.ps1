# Requires -RunAsAdministrator

# Set execution policy to allow script execution ########################################################
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

# Check for elevation ########################################################
if (-not ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator")) {

    Write-Host "Restarting script as administrator..."
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Master list of all available deployment targets
$MasterList = @(
    "t02pc01", "t02pc02", 
    "t04pc01", "t04pc012", 
    "t06pc01", "t06pc02", 
    "t08pc01", "t08pc02", 
    "t10pc01", "t10pc02"
)

Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "REMOTE DEPLOYMENT: MULTI-PC WINDOWS REPAIR" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan

# 1. Target Computer Selection Menu
Write-Host "`n[STEP 1] Choose target computer(s):" -ForegroundColor Cyan
Write-Host "1) Target ALL 10 computers"
Write-Host "2) Target a SINGLE computer from the list"
Write-Host "3) Enter a CUSTOM list of computers (comma-separated)"

$TargetChoice = Read-Host "`nEnter option number (1, 2, or 3)"
$Computers = @()

switch ($TargetChoice) {
    "1" {
        $Computers = $MasterList
        Write-Host "`nTargeting all computers: $($Computers -join ', ')" -ForegroundColor Green
    }
    "2" {
        Write-Host "`nAvailable computers:" -ForegroundColor Gray
        for ($i = 0; $i -lt $MasterList.Count; $i++) {
            Write-Host "$($i + 1)) $($MasterList[$i])"
        }
        $Selection = Read-Host "`nEnter the number of the computer"
        if ($Selection -match '^\d+$' -and [int]$Selection -ge 1 -and [int]$Selection -le $MasterList.Count) {
            $Computers = @($MasterList[[int]$Selection - 1])
            Write-Host "`nTargeting single machine: $Computers" -ForegroundColor Green
        } else {
            Write-Host "Invalid selection. Aborting deployment." -ForegroundColor Red
            Exit
        }
    }
    "3" {
        $CustomInput = Read-Host "`nEnter computer names separated by commas (e.g., t02pc01,t04pc01)"
        $Computers = $CustomInput.Split(',').ForEach({ $_.Trim() }) | Where-Object { $_ -ne "" }
        if ($Computers.Count -eq 0) {
            Write-Host "No computer names entered. Aborting deployment." -ForegroundColor Red
            Exit
        }
        Write-Host "`nTargeting custom list: $($Computers -join ', ')" -ForegroundColor Green
    }
    Default {
        Write-Host "Invalid option. Aborting deployment." -ForegroundColor Red
        Exit
    }
}

# 2. Interactive Restart Selection
Write-Host "`n[STEP 2] Choose restart policy for targets:" -ForegroundColor Cyan
Write-Host "1) Restart with 60-second delay"
Write-Host "2) Restart Immediately"
Write-Host "3) Execute repairs only (Do NOT restart)"

$RestartChoice = Read-Host "`nEnter option number (1, 2, or 3)"

if ($RestartChoice -notin @("1", "2", "3")) {
    Write-Host "Invalid selection. Aborting deployment." -ForegroundColor Red
    Exit
}

# 3. Define the script block that will execute locally on each target PC
$ScriptBlock = {
    param($RestartOption)
    
    $Computer = $env:COMPUTERNAME
    $LogPath = "C:\Windows\Logs\DISM\TotalRepair.log"
    
    function Log-Message([string]$Message) {
        $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        "[$TimeStamp] $Message" | Out-File -FilePath $LogPath -Append
    }
    
    Log-Message "=== Starting repair sequence on $Computer ==="
    
    # Step 1: Deep Scan
    Log-Message "Running DISM ScanHealth..."
    dism /Online /Cleanup-Image /ScanHealth *>$null
    
    # Step 2: Online Restoration
    Log-Message "Running DISM RestoreHealth..."
    dism /Online /Cleanup-Image /RestoreHealth *>$null
    
    # Step 3: Component Store Cleanup
    Log-Message "Running DISM StartComponentCleanup..."
    dism /Online /Cleanup-Image /StartComponentCleanup *>$null
    
    # Step 4: System File Checker
    Log-Message "Running SFC Scannow..."
    sfc /scannow *>$null
    
    Log-Message "Repairs completed on $Computer."
    
    # Step 5: Execute chosen restart policy on the remote computer
    switch ($RestartOption) {
        "1" {
            Log-Message "Initiating 60-second delayed reboot."
            Restart-Computer -Force -Delay 60
        }
        "2" {
            Log-Message "Initiating immediate reboot."
            Restart-Computer -Force
        }
        "3" {
            Log-Message "Sequence complete. No reboot requested."
        }
    }
}

# 4. Parallel Execution across the selected targets
Write-Host "`nDeploying repair sequence to chosen targets in parallel..." -ForegroundColor Yellow

# Invoke-Command runs asynchronously across all targets simultaneously
Invoke-Command -ComputerName $Computers -ScriptBlock $ScriptBlock -ArgumentList $RestartChoice -ThrottleLimit 10

Write-Host "`n[✓] Deployment commands sent to all reachable targets." -ForegroundColor Green
Write-Host "Check local log files at 'C:\Windows\Logs\DISM\TotalRepair.log' on target machines for details." -ForegroundColor Gray