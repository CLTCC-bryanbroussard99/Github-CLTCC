<#
.SYNOPSIS
    Interactive Windows Startup Troubleshooting Menu (Part 2) with Automatic Elevation.
.DESCRIPTION
    This script provides an interactive menu-driven interface to execute common 
    Windows startup, boot configuration, recovery environment, and system repair commands.
.NOTES
    Author: Senior PowerShell Automation Engineer
    Environment: Windows 11 / Windows Server
#>

[CmdletBinding()]
param()

# Ensure script runs with Administrator privileges
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Administrative privileges required. Relaunching script as Administrator..." -ForegroundColor Yellow
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

$ErrorActionPreference = 'Stop'

function Show-Menu {
    Clear-Host
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host "        WINDOWS STARTUP TROUBLESHOOTING TOOLKIT         " -ForegroundColor Yellow
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host " [1] bcdedit /enum all" -ForegroundColor White
    Write-Host "     -> Displays boot configuration entries." -ForegroundColor DarkGray
    Write-Host " [2] bcdedit /export C:\BCD_Backup" -ForegroundColor White
    Write-Host "     -> Backs up boot configuration data." -ForegroundColor DarkGray
    Write-Host " [3] reagentc /info" -ForegroundColor White
    Write-Host "     -> Checks Windows Recovery Environment status." -ForegroundColor DarkGray
    Write-Host " [4] reagentc /enable" -ForegroundColor White
    Write-Host "     -> Enables Windows Recovery Environment." -ForegroundColor DarkGray
    Write-Host " [5] DISM /Image:C:\ /Cleanup-Image /RestoreHealth" -ForegroundColor White
    Write-Host "     -> Repairs an offline Windows image." -ForegroundColor DarkGray
    Write-Host " [6] sfc /scannow /offbootdir=C:\ /offwindir=C:\Windows" -ForegroundColor White
    Write-Host "     -> Scans offline Windows system files." -ForegroundColor DarkGray
    Write-Host " [7] chkdsk C: /r" -ForegroundColor White
    Write-Host "     -> Checks for disk errors and bad sectors." -ForegroundColor DarkGray
    Write-Host " [8] wevtutil qe System /c:30 /rd:true /f:text" -ForegroundColor White
    Write-Host "     -> Displays the latest 30 System event log entries." -ForegroundColor DarkGray
    Write-Host " [9] powercfg /lastwake" -ForegroundColor White
    Write-Host "     -> Identifies the last device or event that woke the PC." -ForegroundColor DarkGray
    Write-Host "[10] explorer.exe" -ForegroundColor White
    Write-Host "     -> Starts the Windows Explorer shell if not running." -ForegroundColor DarkGray
    Write-Host "--------------------------------------------------------" -ForegroundColor Cyan
    Write-Host " [Q] Quit / Exit" -ForegroundColor Red
    Write-Host "========================================================" -ForegroundColor Cyan
}

do {
    Show-Menu
    $choice = Read-Host "Select an option (1-10 or Q)"

    switch ($choice.ToUpper()) {
        '1' {
            Write-Host "`n[+] Executing: bcdedit /enum all" -ForegroundColor Green
            bcdedit /enum all
        }
        '2' {
            Write-Host "`n[+] Executing: bcdedit /export C:\BCD_Backup" -ForegroundColor Green
            bcdedit /export "C:\BCD_Backup"
        }
        '3' {
            Write-Host "`n[+] Executing: reagentc /info" -ForegroundColor Green
            reagentc /info
        }
        '4' {
            Write-Host "`n[+] Executing: reagentc /enable" -ForegroundColor Green
            reagentc /enable
        }
        '5' {
            Write-Host "`n[+] Executing: DISM /Image:C:\ /Cleanup-Image /RestoreHealth" -ForegroundColor Green
            DISM /Image:C:\ /Cleanup-Image /RestoreHealth
        }
        '6' {
            Write-Host "`n[+] Executing: sfc /scannow /offbootdir=C:\ /offwindir=C:\Windows" -ForegroundColor Green
            sfc /scannow /offbootdir=C: /offwindir=C:\Windows
        }
        '7' {
            Write-Host "`n[+] Executing: chkdsk C: /r" -ForegroundColor Green
            Write-Host "Note: Running chkdsk on the active boot drive may prompt for a system restart." -ForegroundColor Yellow
            chkdsk C: /r
        }
        '8' {
            Write-Host "`n[+] Executing: wevtutil qe System /c:30 /rd:true /f:text" -ForegroundColor Green
            wevtutil qe System /c:30 /rd:true /f:text
        }
        '9' {
            Write-Host "`n[+] Executing: powercfg /lastwake" -ForegroundColor Green
            powercfg /lastwake
        }
        '10' {
            Write-Host "`n[+] Executing: explorer.exe" -ForegroundColor Green
            Start-Process explorer.exe
        }
        'Q' {
            Write-Host "Exiting troubleshooting toolkit. Goodbye!" -ForegroundColor Cyan
            exit
        }
        default {
            Write-Host "Invalid selection. Please choose a number between 1 and 10, or Q to exit." -ForegroundColor Red
        }
    }

    Write-Host ""
    Read-Host "Press Enter to return to the menu..."
} while ($true)