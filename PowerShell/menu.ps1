# Requires -RunAsAdministrator

# 1. Define your Active Directory Search Target
# Replace this with the actual Distinguished Name (DN) of your target OU
$TargetOU = "OU=Domain Clients,OU=Cyber312.local Accounts,DC=cyber312,DC=local"
# Requires -RunAsAdministrator

# 1. Define your Active Directory Search Target
# Replace this with the actual Distinguished Name (DN) of your target OU
$TargetOU = "OU=Computers,OU=TechDept,DC=yourdomain,DC=com"

Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "REMOTE DEPLOYMENT: OU-INTEGRATED WINDOWS REPAIR" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan

# Check if Active Directory Module is available
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "[!] Error: Active Directory module is not installed on this machine." -ForegroundColor Red
    Write-Host "Please install RSAT AD tools or run this script from a Domain Controller." -ForegroundColor Yellow
    Exit
}

# 2. Dynamic Target Querying
Write-Host "`nFetching active computer accounts from OU..." -ForegroundColor Yellow
try {
    # Queries the OU, filters for enabled computers, and extracts just their names
    $MasterList = Get-ADComputer -SearchBase $TargetOU -Filter 'Enabled -eq $true' | Select-Object -ExpandProperty Name | Sort-Object
    
    if (-not $MasterList) {
        Write-Host "[!] No active computers found in the specified OU." -ForegroundColor Red
        Exit
    }
}
catch {
    Write-Host "[!] AD Query Failed: Ensure the Distinguished Name (DN) is correct and you have read rights." -ForegroundColor Red
    Write-Error $_
    Exit
}

# 3. Numbered Selection Menu
Write-Host "`nDiscovered Systems in OU:" -ForegroundColor Cyan
for ($i = 0; $i -lt $MasterList.Count; $i++) {
    # Formats index padding for clean vertical alignment (e.g., [ 1], [12])
    $IndexStr = ($i + 1).ToString().PadLeft(($MasterList.Count.ToString().Length))
    Write-Host " [$IndexStr] $($MasterList[$i])" -ForegroundColor Gray
}

Write-Host "`n[STEP 1] Choose target computer(s):" -ForegroundColor Cyan
Write-Host " - Type 'A' to target ALL computers"
Write-Host " - Type a single number to target one computer (e.g., 3)"
Write-Host " - Type comma-separated numbers to target a subset (e.g., 1,3,5)"

$UserSelection = Read-Host "`nEnter your selection"
$Computers = @()

# Evaluate user input
if ($UserSelection.Trim().ToUpper() -eq "A") {
    $Computers = $MasterList
    Write-Host "`nTargeting ALL computers: $($Computers -join ', ')" -ForegroundColor Green
} else {
    # Split input by commas, trim excess spaces, and filter out non-numeric values
    $Indices = $UserSelection.Split(',').ForEach({ $_.Trim() }) | Where-Object { $_ -match '^\d+$' }
    
    foreach ($Index in $Indices) {
        $IntIndex = [int]$Index
        if ($IntIndex -ge 1 -and $IntIndex -le $MasterList.Count) {
            $Computers += $MasterList[$IntIndex - 1]
        }
    }
    
    # Validation step to ensure array populated correctly
    if ($Computers.Count -eq 0) {
        Write-Host "[!] Invalid selection or numbers out of bounds. Aborting deployment." -ForegroundColor Red
        Exit
    }
    
    Write-Host "`nTargeting selection: $($Computers -join ', ')" -ForegroundColor Green
}

# 4. Interactive Restart Selection
Write-Host "`n[STEP 2] Choose restart policy for targets:" -ForegroundColor Cyan
Write-Host "1) Restart with 60-second delay"
Write-Host "2) Restart Immediately"
Write-Host "3) Execute repairs only (Do NOT restart)"

$RestartChoice = Read-Host "`nEnter option number (1, 2, or 3)"

if ($RestartChoice -notin @("1", "2", "3")) {
    Write-Host "Invalid selection. Aborting deployment." -ForegroundColor Red
    Exit
}

# 5. Remote Script Block Execution Setup
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

# 6. Parallel Execution across the selected targets
Write-Host "`nDeploying repair sequence to chosen targets in parallel..." -ForegroundColor Yellow

Invoke-Command -ComputerName $Computers -ScriptBlock $ScriptBlock -ArgumentList $RestartChoice -ThrottleLimit 15

Write-Host "`n[✓] Deployment commands sent to all reachable targets." -ForegroundColor Green
Write-Host "Check local log files at 'C:\Windows\Logs\DISM\TotalRepair.log' on target machines for details." -ForegroundColor Gray