# Menu.ps1
# Persistent PowerShell menu to run other scripts in the same folder

Clear-Host

# --- Determine script directory reliably ---
try {
    # $PSScriptRoot is an automatic variable in PowerShell that contains the full path to the directory of the currently executing script.
    if ($PSScriptRoot) {
        $ScriptDir = $PSScriptRoot
    # $MyInvocation is an automatic variable that provides information about the current command, specifically when that command is a function, a script, or a script block.    
    } elseif ($MyInvocation.MyCommand.Path) {
        $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    } else {
        $ScriptDir = Get-Location
    }
}
catch {
    $ScriptDir = Get-Location
}

# --- Get all .ps1 scripts except this one ---
$ThisScript = Split-Path -Leaf $MyInvocation.MyCommand.Path
$Scripts = Get-ChildItem -Path $ScriptDir -Filter *.ps1 | Where-Object { $_.Name -ne $ThisScript }

if (-not $Scripts) {
    Write-Host "No other PowerShell scripts found in $ScriptDir" -ForegroundColor Yellow
    exit
}

# --- Main menu loop ---
do {
    Clear-Host
    Write-Host "===================================" -ForegroundColor Cyan
    Write-Host "   PowerShell Script Menu" -ForegroundColor Green
    Write-Host "===================================" -ForegroundColor Cyan

    for ($i = 0; $i -lt $Scripts.Count; $i++) {
        Write-Host ("[{0}] {1}" -f ($i + 1), $Scripts[$i].Name)
    }

    Write-Host "[0] Exit" -ForegroundColor Yellow
    Write-Host "==================================="
    $choice = Read-Host "Select a script number to run"

    switch ($choice) {
        '0' {
            Write-Host "Exiting..." -ForegroundColor Yellow
            break
        }

        { $_ -match '^\d+$' -and $_ -le $Scripts.Count -and $_ -gt 0 } {
            $selectedScript = $Scripts[$choice - 1].FullName
            Write-Host "`nRunning: $($selectedScript)`n" -ForegroundColor Cyan
            try {
                & $selectedScript
            }
            catch {
                Write-Host "Error running script: $($_.Exception.Message)" -ForegroundColor Red
            }
            Write-Host "`nScript finished. Press Enter to return to the menu..."
            Read-Host
        }

        default {
            Write-Host "Invalid selection!" -ForegroundColor Red
            Start-Sleep -Seconds 1.5
        }
    }
}
until ($choice -eq '0')
