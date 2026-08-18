<#
Automating at System Startup
To run this script automatically when your computer turns on without requiring a manual user login,
use the Windows Task Scheduler:
 
- Open Task Scheduler from the Windows Start menu.
- Select Create Task on the right action panel.
- Give the task a name (e.g., Autostart Headless VM).
- Under the General tab, select Run whether user is logged on or not and check Run with highest privileges.
- Go to the Triggers tab, click New, and set "Begin the task" to At startup.
- Go to the Actions tab, click New, and choose Start a program:
    - Program/script: powershell.exe
    - Add arguments: -ExecutionPolicy Bypass -File "C:\Path\To\Your\Start-HeadlessVM.ps1"
- Click OK and enter your Windows account credentials when prompted to save the privileged background task.

Key Requirements for Execution

- Path Verification: Ensure VBoxManage.exe is actually installed at 
    "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe". 
    If VirtualBox is installed in a different directory or custom drive, 
    update $VBoxManage accordingly.

- VM Name Matching: Ensure the target VirtualBox VM is named exactly "HomeAssistant" 
   (case-insensitive in VirtualBox, but exact naming prevents issues).

- Execution Policy: If running manually via PowerShell, your local execution policy 
    must allow local scripts (e.g., Set-ExecutionPolicy RemoteSigned -Scope CurrentUser). 
    If running via Task Scheduler as described in the comments, 
    the -ExecutionPolicy Bypass argument handles this automatically.

#>


# Path to VBoxManage.exe
$VBoxManage = "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe"

# Name of your virtual machine
$VMName = "HomeAssistant"

# Start VM in headless mode
& $VBoxManage startvm $VMName --type headless
