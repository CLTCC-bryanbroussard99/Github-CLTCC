<#
    Explaining Your Choice When Running:
        * Pick Option 1 if you managed to install a normal production ISO. It uses slmgr.vbs to slide the key straight into the local license engine.
        * Pick Option 2 if you are on an Evaluation desktop environment. It forces the system to strip the "Eval" limits away and register your core system image with your backup key.
#>

# Ensure script runs with Admin rights
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "Please re-run this script as an Administrator!"
    Exit
}

Clear-Host
Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host "          Windows Server 2019 VM Activation             " -ForegroundColor Cyan
Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host ""

# 1. Capture and validate the backup product key
$KeyPattern = "^([A-Z0-9]{5}-){4}[A-Z0-9]{5}$"
$ValidKey = $false

while (-not $ValidKey) {
    $ProductKey = Read-Host "Enter your 25-character backup product key (XXXXX-XXXXX-XXXXX-XXXXX-XXXXX)"
    $ProductKey = $ProductKey.Trim().ToUpper()

    if ($ProductKey -match $KeyPattern) {
        $ValidKey = $true
        Write-Host "✔ Key format looks valid!" -ForegroundColor Green
    } else {
        Write-Host "❌ Invalid format. Please make sure to include the hyphens." -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host " Choose your activation path based on your VM installation:" -ForegroundColor Cyan
Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host " [1] Standard Activation (Use if you installed a clean Retail/OEM ISO)"
Write-Host " [2] Evaluation Upgrade + Activate (Use if your ISO says 'Evaluation Edition')"
Write-Host " [3] Exit"
Write-Host ""

$Choice = Read-Host "Select an option [1-3]"

switch ($Choice) {
    "1" {
        Write-Host "`n[*] Starting Standard Key Installation..." -ForegroundColor Yellow
        # Install the product key
        cscript //nologo C:\Windows\System32\slmgr.vbs /ipk $ProductKey
        
        Write-Host "[*] Contacting Microsoft servers to force online activation..." -ForegroundColor Yellow
        # Force online activation
        cscript //nologo C:\Windows\System32\slmgr.vbs /ato
        
        Write-Host "`n✔ Process finished. Check system activation properties to confirm." -ForegroundColor Green
    }
    
    "2" {
        Write-Host "`n[*] Starting DISM Evaluation Conversion to ServerStandard..." -ForegroundColor Yellow
        Write-Host "ℹ Note: Your VM will automatically reboot to finalize changes." -ForegroundColor DarkYellow
        Start-Sleep -Seconds 3
        
        # Upgrade edition and inject product key simultaneously
        DISM /online /Set-Edition:ServerStandard /ProductKey:$ProductKey /AcceptEula
    }
    
    "3" {
        Write-Host "`nExiting script. No changes made." -ForegroundColor Gray
        Exit
    }
    
    Default {
        Write-Host "`n❌ Invalid choice selection. Please run the script again." -ForegroundColor Red
    }
}

<#
Yes, you can use that real mixed backup key to reinstall your server in a virtual machine (VM), but because it is an OEM or Retail license being moved into a virtual environment, you must handle the installation carefully to avoid activation failures. [1] 
Follow this workflow to move your physical server into a VM using that backup key:
## Step 1: Verify Your ISO Media Type (Critical)
If you download a standard ISO from Microsoft to build your new VM, it will install as an Evaluation Edition by default. An Evaluation Edition will reject your real backup product key if you try to activate it normally. [2] 
Instead of using standard activation commands (slmgr /ipk), you must use the Deployment Image Servicing and Management (DISM) tool inside your new VM to simultaneously convert the OS to the retail version and inject your backup key: [2, 3] 

   1. Open PowerShell or Command Prompt as an Administrator inside the new VM.
   2. Run this command to elevate the edition and inject your key:
   
   DISM /online /Set-Edition:ServerStandard /ProductKey:XXXXX-XXXXX-XXXXX-XXXXX-XXXXX /AcceptEula
   
   (Replace the X's with your real mixed backup key. If your original server was Datacenter edition, change ServerStandard to ServerDatacenter).
   3. The VM will prompt you to reboot to finalize the change. [2, 3, 4, 5, 6] 

## Step 2: Handle the Activation Fallback
Because Windows detects that the underlying hardware signature has changed (from a physical motherboard to a hypervisor virtual motherboard), online activation via the internet might return a hardware mismatch error (such as error 0xC004C003 or 0xC004F050). [7, 8, 9, 10, 11] 
If automated activation fails after running the DISM command, force Microsoft's phone activation clearinghouse to manually authorize the transfer to the VM:

   1. Inside the VM, open the Run dialog (Win + R), type slui 4, and press Enter.
   2. Select your country to display the Microsoft automated activation phone number and your unique Installation ID.
   3. Call the number, follow the automated prompts, and provide your Installation ID.
   4. When asked "How many computers has this key been installed on?", answer 1 (signifying it is only actively running inside this specific VM, and the old physical install is being wiped/retired).
   5. The system will give you a Confirmation ID to type into your screen, which fully activates the VM. [2, 12, 13, 14, 15] 

## Licensing Best Practices

* 
* Decommission the Host: To remain legally compliant with Microsoft's EULA, ensure that the physical instance of the OS is wiped or shut down so that two environments are not actively utilizing the exact same key at the same time. [16] 
* Hyper-V Exceptions: If you are using Hyper-V on the physical host to run this VM, note that a physical Windows Server Standard license officially grants you rights to run two virtual instances on that same hardware host using the matching key structure. [17, 18] 
* 

To ensure the transition goes smoothly, let me know:

* 
* What hypervisor platform are you building the VM on (Hyper-V, VMware ESXi, or Proxmox)?
* Is your backup key for Windows Server 2019 Standard or Datacenter edition?
* 


[1] [https://learn.microsoft.com](https://learn.microsoft.com/en-nz/answers/questions/2194858/reusing-a-oem-key-for-windows-server-2019)
[2] [https://industrialmonitordirect.com](https://industrialmonitordirect.com/blogs/knowledgebase/resolving-windows-server-2019-hyper-v-vm-activation-non-genuine-error)
[3] [https://learn.microsoft.com](https://learn.microsoft.com/en-us/answers/questions/710462/how-to-activete-virtual-machine-with-widows-server)
[4] [https://support.lenovo.com](https://support.lenovo.com/my/en/solutions/ht515750-how-to-activate-a-windows-server-2019-image-reinstalled-on-a-vm-using-oem-keys-for-the-lenovo-sr650-server)
[5] [https://learn.microsoft.com](https://learn.microsoft.com/en-us/answers/questions/1391223/after-reinstalling-win-server-2019-on-new-hardware)
[6] [https://learn.microsoft.com](https://learn.microsoft.com/en-us/windows-server/virtualization/hyper-v/replication-virtual-machines)
[7] [https://learn.microsoft.com](https://learn.microsoft.com/en-us/answers/questions/5584245/how-to-install-windows-key-for-vms-if-i-face-an-er)
[8] [https://support.lenovo.com](https://support.lenovo.com/my/en/solutions/ht515750-how-to-activate-a-windows-server-2019-image-reinstalled-on-a-vm-using-oem-keys-for-the-lenovo-sr650-server)
[9] [https://learn.microsoft.com](https://learn.microsoft.com/en-us/answers/questions/3898356/purchased-windows-10-pro-to-run-inside-a-virtual-m)
[10] [https://support.rockwellautomation.com](https://support.rockwellautomation.com/app/answers/answer_view/a_id/1076314/~/factorytalk-activation-manager%3A-activations-on-a-dongle-do-not-work-)
[11] [https://www.nakivo.com](https://www.nakivo.com/blog/convert-physical-machine-hyper-v-vm/)
[12] [https://learn.microsoft.com](https://learn.microsoft.com/en-us/answers/questions/1496207/server-2019-standard-activation)
[13] [https://superuser.com](https://superuser.com/questions/821251/can-i-transfer-my-windows-7-license-to-a-virtual-machine-running-on-the-same-c)
[14] [https://learn.microsoft.com](https://learn.microsoft.com/en-us/answers/questions/5510596/how-can-i-re-activate-windows-11-pro-after-a-hardw)
[15] [https://learn.microsoft.com](https://learn.microsoft.com/en-us/answers/questions/2280984/windows-server-standard-license-key-says-already-u)
[16] [https://community.broadcom.com](https://community.broadcom.com/vmware-cloud-foundation/discussion/oem-windows-activation-got-you-down)
[17] [https://protekitsolutions.com](https://protekitsolutions.com/what-license-key-to-use-when-activating-microsoft-windows-server-on-virtual-machines-when-the-hypervisor-host-has-an-oem-copy-of-windows/)
[18] [https://learn.microsoft.com](https://learn.microsoft.com/en-us/answers/questions/2280984/windows-server-standard-license-key-says-already-u)

#>
<#
Moving a Windows Server 2019 license from physical hardware into Proxmox VE (KVM) works perfectly, but because Proxmox emulates a clean, generic motherboard, Windows will immediately recognize the hardware change. [1] 
To prevent performance issues and ensure your backup key activates properly, use this optimized setup and activation blueprint for Proxmox.
------------------------------
## Step 1: Optimize the Proxmox VM Settings
Before installing Windows, configure your VM hardware to match Microsoft's expectations and allow VirtIO high-performance drivers. [2] 

* OS Tab: Set Type to Microsoft Windows and Version to 10/2016/2019. [3] 
* System Tab:
* Check the box for Qemu Agent (crucial for clean reboots and backups).
   * Set the Machine type to q35.
   * Set the BIOS to OVMF (UEFI) rather than SeaBIOS. (Windows Server 2019 boots faster and manages GPT disks better under UEFI). [4, 5, 6, 7, 8] 
* CPU Tab: Set the Type to host. This passes your actual physical CPU instructions directly to the VM, which helps Windows stability and performance. [9, 10, 11, 12] 
* Network/Disks: Use VirtIO Block for disks and VirtIO (paravirtualized) for network to get native speed.

Note: You will need to download and attach the official [Fedora VirtIO ISO](https://pve.proxmox.com/wiki/Windows_VirtIO_Drivers) as a second CD-ROM drive to load the storage drivers during the Windows installation wizard. [13, 14] 
------------------------------
## Step 2: Convert and Inject the Key (DISM Method)
If your installation media defaults to an Evaluation ISO, your backup key will be rejected by the standard settings menu. You must force the version upgrade and inject your key at the same time using PowerShell inside the VM.

   1. Right-click the Windows Start button inside the Proxmox VM and open Windows PowerShell (Admin).
   2. Run the elevation command matching your license type:For Windows Server 2019 Standard:
   
   DISM /online /Set-Edition:ServerStandard /ProductKey:XXXXX-XXXXX-XXXXX-XXXXX-XXXXX /AcceptEula
   
   For Windows Server 2019 Datacenter:
   
   DISM /online /Set-Edition:ServerDatacenter /ProductKey:XXXXX-XXXXX-XXXXX-XXXXX-XXXXX /AcceptEula
   
   (Replace the X's with your real mixed backup key).
   3. The VM will process the activation files and ask you to press Y to reboot. [15] 

------------------------------
## Step 3: Bypass Proxmox Hardware Activation Errors
Because Proxmox does not replicate your old physical motherboard's UUID or serial numbers, automated internet activation will likely fail with error 0xC004C003. You must bypass this via Microsoft's telephone system.

   1. Once the VM reboots, press Win + R, type slui 4, and hit Enter.
   2. Select your country to generate a toll-free activation number and a 63-digit Installation ID.
   3. Call the automated number (or use the provided smartphone link if offered by the system).
   4. When the system asks how many computers are using this key, type 1. If you say 0 or 2+, the system will block the key.
   5. Enter the provided Confirmation ID blocks back into the Proxmox VM window to permanently bind the license to your virtual environment.

If you run into any issues, let me know:

* Have you already installed the Windows ISO on the VM, or are you preparing to build it now?
* Did the DISM conversion command complete successfully, or did it return an error code?


[1] [https://www.reddit.com](https://www.reddit.com/r/homeassistant/comments/1rq9wlj/upgrading_from_ha_green_to_mini_pc_with_proxmox/)
[2] [https://www.facebook.com](https://www.facebook.com/groups/proxmox/posts/7908161252545871/)
[3] [https://forum.proxmox.com](https://forum.proxmox.com/threads/how-to-convert-windows-10-generation-2-hyper-v-vm-to-a-proxmox-vm.107511/)
[4] [https://www.it-connect.tech](https://www.it-connect.tech/how-to-create-a-windows-11-virtual-machine-in-proxmox-ve/)
[5] [https://medium.com](https://medium.com/@antosanto/installing-trunas-scale-on-proxmox-0fd3fb77fd8a)
[6] [https://docs.siderolabs.com](https://docs.siderolabs.com/talos/v1.13/platform-specific-installations/virtualized-platforms/proxmox)
[7] [https://bobcares.com](https://bobcares.com/blog/proxmox-uefi-not-booting/)
[8] [https://www.reddit.com](https://www.reddit.com/r/Proxmox/comments/1trztwg/windows_server_2025_fails_to_boot_after_pve_92/)
[9] [https://support.checkpoint.com](https://support.checkpoint.com/results/sk/sk180399)
[10] [https://www.saturnme.com](https://www.saturnme.com/proxmox-vm-cpu-type-explained-host-vs-x86-64-v2-aes-performance-vs-portability-deep-dive/)
[11] [https://www.servereasy.it](https://www.servereasy.it/en/cloud-and-virtualization-en/business-vps/)
[12] [https://www.vinchin.com](https://www.vinchin.com/vm-backup/proxmox-for-windows.html)
[13] [https://forum.proxmox.com](https://forum.proxmox.com/threads/p2v-windows-11-to-proxmox.139934/)
[14] [https://yetiops.net](https://yetiops.net/posts/proxmox-terraform-cloudinit-windows/)
[15] [https://www.starwindsoftware.com](https://www.starwindsoftware.com/blog/proxmox-import-vmware-vms/)

#>