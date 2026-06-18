<#
    Description: Turns your Windows hosts file into a DNS-based ad and malware filter.
    Source: Uses Steven Black’s unified hosts list (ads + malware + tracking)
    https://github.com/StevenBlack/hosts
#>

# Run as admin check
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "⚠️  Please run this script as Administrator." -ForegroundColor Yellow
    exit
}

$HostsPath = "$env:SystemRoot\System32\drivers\etc\hosts"
$BackupPath = "$env:SystemRoot\System32\drivers\etc\hosts.bak_$(Get-Date -Format yyyyMMdd_HHmmss)"
$TempFile = "$env:TEMP\hosts_download.txt"
$BlocklistURL = "https://raw.githubusercontent.com/StevenBlack/hosts/master/hosts"

Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "   Windows Hosts DNS Filter Installer" -ForegroundColor Green
Write-Host "=========================================" -ForegroundColor Cyan

# --- Backup existing hosts file ---
Write-Host "Backing up current hosts file to:`n$BackupPath" -ForegroundColor Yellow
Copy-Item -Path $HostsPath -Destination $BackupPath -Force

# --- Download blocklist ---
Write-Host "`nDownloading latest DNS filter list..." -ForegroundColor Cyan
try {
    Invoke-WebRequest -Uri $BlocklistURL -OutFile $TempFile -UseBasicParsing -ErrorAction Stop
}
catch {
    Write-Host "❌ Failed to download blocklist: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# --- Preserve custom user entries (lines not starting with 0.0.0.0 or 127.0.0.1) ---
Write-Host "Merging your local entries with the blocklist..." -ForegroundColor Cyan
$UserEntries = Get-Content $HostsPath | Where-Object { ($_ -match '^\s*#') -or ($_ -notmatch '^\s*(0\.0\.0\.0|127\.0\.0\.1)\s+') }

# Combine header + user entries + blocklist
$Header = @(
    "# ====================================================================="
    "# Custom Hosts DNS Filter"
    "# Generated on $(Get-Date)"
    "# Source: $BlocklistURL"
    "# Backup: $BackupPath"
    "# ====================================================================="
)
$Combined = $Header + "" + $UserEntries + "" + (Get-Content $TempFile)

# --- Write combined file ---
Write-Host "Updating hosts file..." -ForegroundColor Yellow
Set-Content -Path $HostsPath -Value $Combined -Force -Encoding ASCII

# --- Flush DNS ---
Write-Host "`nFlushing DNS cache..." -ForegroundColor Cyan
ipconfig /flushdns | Out-Null

Write-Host "`n✅ Hosts file successfully updated!"
Write-Host "Blocked domains active immediately."
Write-Host "Backup saved to: $BackupPath" -ForegroundColor Green

# --- Optional cleanup ---
Remove-Item $TempFile -ErrorAction SilentlyContinue
