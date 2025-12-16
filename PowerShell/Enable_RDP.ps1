# Enable Remote Desktop
# This script enables Remote Desktop on a Windows machine and allows it through the firewall.
# It modifies the registry to allow Remote Desktop connections and enables the necessary firewall rules.
Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server' -Name "fDenyTSConnections" -Value 0

# This command enables the firewall rule for Remote Desktop
Enable-NetFirewallRule -DisplayGroup "Remote Desktop"

# This command restarts the Remote Desktop service to apply the changes
Restart-Service -Name TermService

# This command checks if the Remote Desktop service is running
Get-Service -Name TermService

# This command checks if the firewall rule for Remote Desktop is enabled
Get-NetFirewallRule -DisplayGroup "Remote Desktop" | Select-Object Name, Enabled