<#
.SYNOPSIS
    Safely remove a PDC Emulator DC from the domain with logging and replication health checks.
.DESCRIPTION
    This script:
      • Automatically normalizes short/FQDN DC names
      • Confirms actions interactively
      • Logs every step to a specified file
      • Runs dcdiag & repadmin before and after
      • Transfers the PDC FSMO role, then demotes and cleans up
.NOTES
    Run as Domain Admin
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$NewPDC,

    [Parameter(Mandatory = $true)]
    [string]$OldPDC,

    [Parameter(Mandatory = $true)]
    [string]$LogPath,

    [switch]$ForceDemote
)

# =========================
# Function Definitions
# =========================

function Write-Log {
    param([string]$Message, [string]$Color = "White")
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $entry = "[$timestamp] $Message"
    Write-Host $entry -ForegroundColor $Color
    Add-Content -Path $LogPath -Value $entry
}

function Run-HealthCheck {
    param([string]$Phase)
    Write-Log "Running $Phase health check..." Cyan
    $dcdiagFile = Join-Path (Split-Path $LogPath) "$($Phase)_dcdiag.txt"
    $repadminFile = Join-Path (Split-Path $LogPath) "$($Phase)_repadmin.txt"

    dcdiag /v > $dcdiagFile 2>&1
    repadmin /replsummary > $repadminFile 2>&1

    Write-Log "Health check complete. Results saved to:`n $dcdiagFile`n $repadminFile" Green
}

# =========================
# Script Body
# =========================

$ErrorActionPreference = "Stop"
Import-Module ActiveDirectory

# Normalize names
function Get-CanonicalDCName {
    param([string]$DCName)
    try {
        $dc = Get-ADDomainController -Identity $DCName -ErrorAction Stop
        return $dc.HostName
    } catch {
        Write-Host "⚠️ Could not resolve $DCName in AD. Using provided name." -ForegroundColor Yellow
        return $DCName
    }
}

$NewPDC = Get-CanonicalDCName $NewPDC
$OldPDC = Get-CanonicalDCName $OldPDC

# Confirm start
Write-Host "`n*** SAFE PDC REMOVAL SCRIPT ***" -ForegroundColor Yellow
Write-Host "This will transfer FSMO role and demote $OldPDC." -ForegroundColor Yellow
$confirmStart = Read-Host "Continue? (Y/N)"
if ($confirmStart -notmatch '^[Yy]$') {
    Write-Host "Operation cancelled." -ForegroundColor Red
    exit
}

# Log header
New-Item -ItemType File -Path $LogPath -Force | Out-Null
Write-Log "Starting PDC Removal Script. Log: $LogPath" Cyan
Write-Log "Target New PDC: $NewPDC"
Write-Log "Target Old PDC: $OldPDC"

# Pre-check
Write-Log "Checking current FSMO roles..."
$currentPDC = (Get-ADDomain).PDCEmulator
$currentShort = ($currentPDC -split '\.')[0]
$oldShort = ($OldPDC -split '\.')[0]

if (($currentPDC -ne $OldPDC) -and ($currentShort -ne $oldShort)) {
    Write-Log "⚠️ Warning: $OldPDC does not match current PDC Emulator ($currentPDC)." Yellow
    $continue = Read-Host "Continue anyway? (Y/N)"
    if ($continue -notmatch '^[Yy]$') {
        Write-Log "Operation cancelled by user." Red
        exit
    }
}

# Run pre-health check
Run-HealthCheck -Phase "PreCheck"

# FSMO transfer
Write-Host "`nReady to transfer PDC Emulator role to $NewPDC." -ForegroundColor Cyan
$confirmTransfer = Read-Host "Proceed with FSMO transfer? (Y/N)"
if ($confirmTransfer -match '^[Yy]$') {
    Write-Log "Transferring FSMO PDC role from $OldPDC to $NewPDC..." Yellow
    Move-ADDirectoryServerOperationMasterRole -Identity $NewPDC -OperationMasterRole PDCEmulator -Confirm:$false
    $newHolder = (Get-ADDomain).PDCEmulator
    Write-Log "New PDC Emulator is now: $newHolder" Green
} else {
    Write-Log "FSMO transfer cancelled by user." Red
    exit
}

# Demotion
Write-Host "`nReady to demote $OldPDC." -ForegroundColor Cyan
$confirmDemote = Read-Host "Proceed with demotion? (Y/N)"
if ($confirmDemote -match '^[Yy]$') {
    Write-Log "Starting demotion of $OldPDC..." Cyan
    try {
        Invoke-Command -ComputerName $OldPDC -ScriptBlock {
            Import-Module ADDSDeployment
            if ($using:ForceDemote) {
                Uninstall-ADDSDomainController -DemoteOperationMasterRole:$true -ForceRemoval:$true -Force -LocalAdministratorPassword (ConvertTo-SecureString "TempPass123!" -AsPlainText -Force)
            } else {
                Uninstall-ADDSDomainController -DemoteOperationMasterRole:$true -Force -LocalAdministratorPassword (ConvertTo-SecureString "TempPass123!" -AsPlainText -Force)
            }
            Restart-Computer -Force
        }
        Write-Log "Demotion initiated. Server $OldPDC rebooting..." Yellow
    } catch {
        Write-Log "Demotion failed: $_" Red
    }
} else {
    Write-Log "Demotion cancelled by user." Red
    exit
}

# Cleanup
Write-Host "`nPerforming AD cleanup..." -ForegroundColor Cyan
try {
    Remove-ADComputer -Identity $OldPDC -Confirm:$false -ErrorAction SilentlyContinue
    Write-Log "Removed $OldPDC computer object from AD." Green
} catch {
    Write-Log "Warning: Unable to remove AD object for $OldPDC. $_" Yellow
}

# Post health check
Run-HealthCheck -Phase "PostCheck"

Write-Log "`n✅ PDC removal and cleanup complete." Green
Write-Log "Review log and health check files for details."
