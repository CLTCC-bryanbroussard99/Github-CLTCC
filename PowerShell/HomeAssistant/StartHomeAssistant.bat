
# Path to VBoxManage.exe
$VBoxManage = "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe"

# Name of your virtual machine
$VMName = "HomeAssistant"

# Start VM in headless mode
& $VBoxManage startvm $VMName --type headless
