# Menu.ps1
# Persistent PowerShell menu to run other scripts in the same folder

# Relaunch the script using the latest PowerShell (pwsh) if running in legacy PowerShell
if ($PSVersionTable.PSVersion.Major -le 5) {
    if (Get-Command pwsh -ErrorAction SilentlyContinue) {
        pwsh -File $MyInvocation.MyCommand.Definition
        exit
    } else {
        Write-Error "PowerShell 7+ is not installed on this system."
        exit
    }
}


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

# --- Main menu loop ---
do {
    Clear-Host
    Write-Host "===================================" -ForegroundColor Cyan
    Write-Host "   PowerShell Script Menu" -ForegroundColor Green
    Write-Host "===================================" -ForegroundColor Cyan

    # Display dynamic scripts if any exist
    if ($Scripts) {
        for ($i = 0; $i -lt $Scripts.Count; $i++) {
            Write-Host ("[{0}] {1}" -f ($i + 1), $Scripts[$i].Name)
        }
        Write-Host "-----------------------------------"
    } else {
        Write-Host "No local scripts found in $ScriptDir" -ForegroundColor Yellow
        Write-Host "-----------------------------------"
    }

    # Built-in utilities
    Write-Host "[G] Run GPUpdate (Force)" -ForegroundColor Magenta
    Write-Host "[0] Exit" -ForegroundColor Yellow
    Write-Host "==================================="
    $choice = Read-Host "Select an option"

    switch ($choice) {
        '0' {
            Write-Host "Exiting..." -ForegroundColor Yellow
            break
        }

        'G' {
            Write-Host "`nRunning Group Policy Update..." -ForegroundColor Cyan
            try {
                gpupdate /force
            }
            catch {
                Write-Host "Error running gpupdate: $($_.Exception.Message)" -ForegroundColor Red
            }
            Write-Host "`nProcess finished. Press Enter to return to the menu..."
            Read-Host
        }

        { $Scripts -and $_ -match '^\d+$' -and $_ -le $Scripts.Count -and $_ -gt 0 } {
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