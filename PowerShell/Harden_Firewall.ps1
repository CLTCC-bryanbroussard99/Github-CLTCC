Write-Host "Starting Windows Firewall hardening..."

# Enable firewall and block incoming traffic
$response = Read-Host -Prompt "Do you want to Enable firewall and block incoming traffic? (Y/N)"
if ($response -eq 'Y' -or $response -eq 'y') {
    Write-Host "You chose Yes. Proceeding..." -ForegroundColor Green
    Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled True
    Set-NetFirewallProfile -Profile Domain,Public,Private -DefaultInboundAction Block -DefaultOutboundAction Allow
}

# Hide system on the network (block ICMP responses)
$response = Read-Host -Prompt "Do you want to Hide system on the network (block ICMP responses)? (Y/N)"
if ($response -eq 'Y' -or $response -eq 'y') {
    Write-Host "You chose Yes. Proceeding..." -ForegroundColor Green
    Set-NetFirewallProfile -Profile Domain,Public,Private -AllowInboundEchoRequest False
}

# Disable Remote Desktop
$response = Read-Host -Prompt "Do you want Disable Remote Desktop? (Y/N)"
if ($response -eq 'Y' -or $response -eq 'y') {
    Write-Host "You chose Yes. Proceeding..." -ForegroundColor Green
    Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Terminal Server" -Name "fDenyTSConnections" -Value 1
    Set-Service -Name "TermService" -StartupType Disabled
    Stop-Service -Name "TermService" -Force
}

# Disable Remote Registry
$response = Read-Host -Prompt "Do you want to Disable Remote Registry? (Y/N)"
if ($response -eq 'Y' -or $response -eq 'y') {
    Write-Host "You chose Yes. Proceeding..." -ForegroundColor Green
    Set-Service -Name "RemoteRegistry" -StartupType Disabled
    Stop-Service -Name "RemoteRegistry" -Force
}

# Disable WinRM (remote PowerShell)
$response = Read-Host -Prompt "Do you want to Disable WinRM (remote PowerShell)? (Y/N)"
if ($response -eq 'Y' -or $response -eq 'y') {
    Write-Host "You chose Yes. Proceeding..." -ForegroundColor Green
    Disable-PSRemoting -Force  
}

# Disable WMI (remote management)
$response = Read-Host -Prompt "Do you want to WMI (remote management)? (Y/N)"
if ($response -eq 'Y' -or $response -eq 'y') {
    Write-Host "You chose Yes. Proceeding..." -ForegroundColor Green
    Set-NetFirewallRule -DisplayName "Windows Management Instrumentation (WMI-In)" -Enabled False
}

# Block common scanning ports
$response = Read-Host -Prompt "Do you want to block common scanning ports (21, 23, 135, 137, 138, 139, 445, 1433, 3389)? (Y/N)"
if ($response -eq 'Y' -or $response -eq 'y') {
    Write-Host "You chose Yes. Blocking common scanning ports ()" -ForegroundColor Green
    $ports = @(21, 23, 135, 137, 138, 139, 445, 1433, 3389)

    foreach ($port in $ports) {
        New-NetFirewallRule -DisplayName "Block port $port" -Direction Inbound -Protocol TCP -LocalPort $port -Action Block
        New-NetFirewallRule -DisplayName "Block port UDP $port" -Direction Inbound -Protocol UDP -LocalPort $port -Action Block
    }
}

# Disable SMBv1
$response = Read-Host -Prompt "Do you want to Disable SMBv1 (HIGHLY RECOMMENDED)? (Y/N)"
if ($response -eq 'Y' -or $response -eq 'y') {
    Write-Host "You chose Yes. Disabling SMBv1" -ForegroundColor Green
    Disable-WindowsOptionalFeature -Online -FeatureName SMB1Protocol -NoRestart
}

# Disable Telnet Client (if installed)
$response = Read-Host -Prompt "Do you want to Disable Telnet Client? (Y/N)"
if ($response -eq 'Y' -or $response -eq 'y') {
    Write-Host "You chose Yes. Disabling Telnet Client" -ForegroundColor Green
    Disable-WindowsOptionalFeature -Online -FeatureName TelnetClient -NoRestart
}

Write-Host "Hardening complete. System restart recommended."

# Recommend a restart to apply all changes
Write-Host "Please restart your computer to apply all changes."
