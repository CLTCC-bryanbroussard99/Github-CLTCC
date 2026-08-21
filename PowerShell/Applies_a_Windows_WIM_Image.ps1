<#
.SYNOPSIS
    Applies a Windows WIM image and configures boot/recovery settings.

.DESCRIPTION
    Deploys a specified Windows image file to the target Windows partition (W:\),
    configures boot files on the System partition (S:\), and sets up Windows RE (R:\).
    Based on Microsoft documentation:
    https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/capture-and-apply-windows-using-a-single-wim?view=windows-11

.PARAMETER ImagePath
    The path to the WIM image file to apply (e.g., E:\Images\ThinImage.wim).

.EXAMPLE
    .\ApplyImage.ps1 -ImagePath "E:\Images\ThinImage.wim"
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, Position = 0, HelpMessage = "Path to the .wim file.")]
    [string]$ImagePath
)

# Set execution safety options
$ErrorActionPreference = "Stop"

# Define standard partition drive letters
$windowsDrive  = "W:\"
$systemDrive   = "S:"
$recoveryDrive = "R:\Recovery\WindowsRE"

# Set high-performance power scheme to speed up the deployment process
Write-Host "Setting power scheme to High Performance..." -ForegroundColor Green
powercfg.exe /s 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c

# Apply the Windows image to the target partition using DISM
Write-Host "Applying Windows image from '$ImagePath' to $windowsDrive..." -ForegroundColor Green
Expand-WindowsImage -ImagePath $ImagePath -Index 1 -ApplyPath $windowsDrive

# Copy boot files from the newly applied Windows image to the System partition
Write-Host "Configuring boot files on system partition $systemDrive..." -ForegroundColor Green
& "$windowsDrive\Windows\System32\bcdboot.exe" "$windowsDrive\Windows" /s $systemDrive

# ==============================================================================
# Optional Recovery Partition Configuration (Uncomment if using WinRE)
# ==============================================================================

# Create the Windows RE directory on the recovery partition
# if (-not (Test-Path -Path $recoveryDrive)) {
#     New-Item -Path $recoveryDrive -ItemType Directory -Force | Out-Null
# }

# Copy the WinRE image to the Recovery partition (including hidden files)
# Copy-Item -Path "$windowsDrive\Windows\System32\Recovery\Winre.wim" -Destination $recoveryDrive -Force

# Register the location of the recovery tools with Windows
# & "$windowsDrive\Windows\System32\Reagentc.exe" /Setreimage /Path $recoveryDrive /Target "$windowsDrive\Windows"

# Verify the configuration status of the recovery environment
# & "$windowsDrive\Windows\System32\Reagentc.exe" /Info /Target "$windowsDrive\Windows"