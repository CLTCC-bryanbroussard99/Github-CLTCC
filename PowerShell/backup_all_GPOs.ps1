<#
.SYNOPSIS
    Backup all GPOs from a remote domain controller.

.DESCRIPTION
    This script connects to a remote computer, verifies access to the Group Policy module,
    and exports all or selected GPOs to a specified backup path.

.EXAMPLE
    .\Backup-GPOs.ps1 -RemoteServer DC02 -BackupPath "\\dc02"
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$RemoteServer,

    [Parameter(Mandatory = $true)]
    [string]$BackupPath
)

# Ensure the backup path exists locally or remotely
if (!(Test-Path $BackupPath)) {
    Write-Host "Creating backup directory at $BackupPath..." -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $BackupPath -Force | Out-Null
}

# Establish a remote session
Write-Host "Connecting to $RemoteServer..." -ForegroundColor Cyan
$session = New-PSSession -ComputerName $RemoteServer -ErrorAction Stop

try {
    Write-Host "Starting GPO backup on $RemoteServer..." -ForegroundColor Green
    Invoke-Command -Session $session -ScriptBlock {
        param($BackupPath)

        Import-Module GroupPolicy

        $date = Get-Date -Format "yyyyMMdd_HHmmss"
        $backupDir = Join-Path $BackupPath "GPO_Backup_$date"

        New-Item -ItemType Directory -Path $backupDir -Force | Out-Null

        $gpos = Get-GPO -All
        foreach ($gpo in $gpos) {
            Write-Host "Backing up GPO: $($gpo.DisplayName)" -ForegroundColor Cyan
            Backup-GPO -Name $gpo.DisplayName -Path $backupDir -Comment "Automated backup $(Get-Date)"
        }

        Write-Host "`nAll GPOs backed up to: $backupDir" -ForegroundColor Green
    } -ArgumentList $BackupPath
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Remove-PSSession $session
    Write-Host "Session closed." -ForegroundColor Yellow
}
