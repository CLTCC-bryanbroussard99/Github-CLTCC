# This script is designed to update the Secure Boot policy on a Windows system.
# It uses PowerShell commands to modify the registry and start a scheduled task.
# This command adds or updates a registry value that tells Windows to stage the certificate updates.
# It will also start the Secure Boot update task to apply the changes.
Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot" -Name "AvailableUpdates" -Value 0x40

# Start the Secure Boot update task to apply the changes
# Note: The task name may vary depending on the version of Windows you are using. 
# You can use the Get-ScheduledTask cmdlet to find the correct task name.
Start-ScheduledTask -TaskName "\Microsoft\Windows\PI\Secure-Boot-Update"

# Normally this task runs automatically every 12 hours, but this command forces it to run immediately.

# Reboot your computer twice to allow the changes to the UEFI firmware to take effect.
# After the second reboot, the changes should be reflected in the UEFI firmware.
# To Verify that the Secure Boot policy has been updated
#   This step is optional but recommended to ensure the changes have been applied successfully.
# You can use the following command to check the status of the Secure Boot policy:
# Get-Scheduled-S -TaskName "\Microsoft\Windows\PI\Secure-Boot-Update" | Select-Object -Property State, LastRunTime, NextRunTime    