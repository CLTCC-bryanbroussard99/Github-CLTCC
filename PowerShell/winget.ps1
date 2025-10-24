
# Winget Documentation - https://learn.microsoft.com/en-us/windows/package-manager/winget/install
# Browse winget packages - https://winget.run/pkg/Microsoft or https://winget.run/ or https://winstall.app/


# Set execution policy to allow script execution ########################################################
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

# Check for elevation ########################################################
if (-not ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator")) {

    Write-Host "Restarting script as administrator..."
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Install Winget if not installed and ensure it is the latest version ########################################################
$wingetInstaller = "$env:TEMP\Microsoft.DesktopAppInstaller.msixbundle"
if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    Write-Host "Installing Winget..."
    Invoke-WebRequest -Uri "https://aka.ms/getwinget" -OutFile $wingetInstaller
    Add-AppxPackage -Path $wingetInstaller
    Start-Sleep -Seconds 10
}

# Check if Winget is installed ########################################################
if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    Write-Host "Winget installation failed. Please check the installer." -ForegroundColor Red
    exit 1
}

# Ensure the script runs with the latest version of PowerShell ########################################################
$PSVersionTable.PSVersion.Major -lt 5.1 | Out-Null
if ($?) {
    Write-Host "PowerShell version is sufficient." -ForegroundColor Green
} else {
    Write-Host "Please update PowerShell to version 5.1 or higher." -ForegroundColor Red
    exit 1
}

# Check if the script is running on Windows 10 or later ########################################################
if ([System.Environment]::OSVersion.Version.Major -lt 10) {
    Write-Host "This script requires Windows 10 or later." -ForegroundColor Red
    exit 1
} else {
    Write-Host "Operating system is compatible." -ForegroundColor Green
}
# Check if the script is running on a 64-bit architecture ########################################################
if ([System.Environment]::Is64BitOperatingSystem -eq $false) {
    Write-Host "This script requires a 64-bit operating system." -ForegroundColor Red
    exit 1
} else {
    Write-Host "Running on a 64-bit operating system." -ForegroundColor Green
}

# Ensure the script is running with administrative privileges ########################################################
if (-not ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "This script must be run as an administrator." -ForegroundColor Red
    exit 1
} else {
    Write-Host "Running with administrative privileges." -ForegroundColor Green
}
# Ensure the script is running in a PowerShell environment ########################################################
if ($PSVersionTable.PSEdition -ne "Core" -and $PSVersionTable.PSEdition -ne "Desktop") {
    Write-Host "This script must be run in a PowerShell environment." -ForegroundColor Red
    exit 1
} else {
    Write-Host "Running in a PowerShell environment." -ForegroundColor Green
}


# Allow remoting and unrestricted script execution temporarily
Enable-PSRemoting -Force -SkipNetworkProfileCheck


# Update winget sources  ########################################################
winget source update

# array of programs to install  ########################################################
$programs =@(   
# Browsers
'Microsoft.Edge', # World-class performance with more privacy, more productivity, and more value while you browse.
'Mozilla.Firefox', # Firefox Browser, also known as Mozilla Firefox or simply Firefox, is a free and open-source web browser developed by the Mozilla Foundation and its subsidiary, the Mozilla Corporation. Firefox uses the Gecko layout engine to render web pages, which implements current and anticipated web standards.
'VivaldiTechnologies.Vivaldi', # The new Vivaldi browser protects you from trackers, blocks unwanted ads, and puts you in control with unique built-in features. Get Vivaldi and browse fast.
'Google.Chrome', # Google Chrome - A more simple, secure, and faster web browser than ever, with Google’s smarts built-in.
'Opera.Opera', # Opera Stable - is a multi-platform web browser developed by Opera Software. Opera is a Chromium-based browser. It distinguishes itself from other browsers through its user interface, functionality, and other features

#Tools
'Git.Git', # Git - is a distributed version control system that tracks changes in any set of computer files, usually used for coordinating work among programmer
'GitHub.GitHubDesktop', #Focus on what matters instead of fighting with Git. Whether you're new to Git or a seasoned user, GitHub Desktop simplifies your development workflow.
'Google.GoogleDrive', #Mounts Google Drive(s) as a share drive and streams files as needed from the cloud. 
'Microsoft.VCRedist.2015+.x64'. #
'Oracle.MySQLWorkbench', #
'PowerSoftware.PowerISO', #Provides an all-in-one solution. You can do every thing with your ISO files and disc image files.
'Microsoft.PowerToys', #Microsoft PowerToys is a set of utilities for power users to tune and streamline their Windows experience for greater productivity
'RevoUninstaller.RevoUninstaller'. #Revo Uninstaller is a freeware uninstall utility. It has powerful features to uninstall programs scanning for leftover files,
'VideoLAN.VLC', #VLC is a free and open source cross-platform multimedia player and framework that plays most multimedia files as well as DVDs, Audio CDs, VCDs, and various streaming protocols.
'RustDesk.RustDesk', #RustDesk is a full-featured open source remote control alternative for self-hosting and security with minimal configuration.

# Communication
'Microsoft.Teams', #Make amazing things happen together at home, work, and school.
'tomlm.electron-outlook-365', #Outlook Web Application as a dedicated application. This hosts mail.office365.com and is suitable for using with enterprise work or school environmen

# Various 
'PDFgear.PDFgear', #Read, edit, convert, merge, and sign PDF files completely free
'Git.Git', # Git - is a distributed version control system that tracks changes in any set of computer files, usually used for coordinating work among programmer
'Microsoft.LAPS', # Local Administrator Password Solution - Provides management of local account passwords of domain joined computers. Passwords are stored in Active Directory
'Microsoft.DeploymentToolkit', # Microsoft Deployment Toolkit (MDT) - provides a unified collection of tools, processes, and guidance for automating desktop and server deployments.
'Microsoft.VisualStudioCode', # Install VSCode - 
'Famatech.AdvancedIPScanner', # Advanced IP Scanner - shows all network devices, gives you access to shared folders, and can even remotely switch computers off.
'Oracle.MySQL', # MySQL - delivers a very fast, multithreaded, multi-user, and robust SQL (Structured Query Language) database server.
'Oracle.JavaRuntimeEnvironment', # Its also integral to the intranet applications and other e-business solutions that are the foundation of corporate computing.
'Python.Python.3.12', # Python - is a programming language that lets you work more quickly and integrate your systems more effectively.
'7zip.7zip', # 7-zip - Free and open source file archiver with a high compression ratio.
'PuTTY.PuTTY', # PuTTY - A free implementation of SSH and Telnet, along with an xterm terminal emulator.
'Notepad++.Notepad++', # Notepad++ - is a free (as in “free speech” and also as in “free beer”) source code editor and Notepad replacement that supports several languages. Running in the MS Windows environment, its use is governed by GNU General Public License.
'RevoUninstaller.RevoUninstaller', # Revo Uninstaller - is a freeware uninstall utility. It has powerful features to uninstall programs scanning for leftover files, folders and registry entries after uninstall. With its unique 'Hunter mode' it offers you some simple, easy to use but effective and powerful approaches to manage (uninstall, stop, delete, disable from auto starting) and to get information about your installed and/or running application. Many cleaning tools included!
'CodecGuide.K-LiteCodecPack.Standard', # K-Lite Codec Pack Standard - is a collection of DirectShow filters, VFW/ACM codecs, and tools. Codecs and DirectShow filters are needed for encoding and decoding audio and video formats. The K-Lite Codec Pack is designed as a user-friendly solution for playing all your audio and movie files. With the K-Lite Codec Pack you should be able to play all the popular audio and video formats and even several less common formats.
'Microsoft.DotNet.DesktopRuntime.7', # Microsoft .NET Windows Desktop Runtime 7.0 - .NET is a free, cross-platform, open-source developer platform for building many different types of applications.
'dotPDNLLC.paintdotnet', # Paint.NET - is image and photo editing software for PCs that run Windows.
'Microsoft.PowerShell', # PowerShell - is a cross-platform (Windows, Linux, and macOS) automation and configuration tool/framework that works well with your existing tools and is optimized for dealing with structured data (e.g. JSON, CSV, XML, etc.), REST APIs, and object models. It includes a command-line shell, an associated scripting language and a framework for processing cmdlets.
'Oracle.VirtualBox', # VirtualBox - is a powerful x86 and AMD64/Intel64 virtualization product for enterprise as well as home use. Not only is VirtualBox an extremely feature rich, high performance product for enterprise customers, it is also the only professional solution that is freely available as Open Source Software under the terms of the GNU General Public License (GPL) version 3.
'WiresharkFoundation.Wireshark' # Wireshark - is the world's foremost and widely-used network protocol analyzer. It lets you see what's happening on your network at a microscopic level and is the de facto (and often de jure) standard across many commercial and non-profit enterprises, government agencies, and educational institutions.
'Bitwarden.Bitwarden' #A secure and free password manager for all of your devices.
)

# Unistall Microsoft Office if installed ########################################################

<#
.SYNOPSIS
    Detects Microsoft Office installations and attempts to uninstall them.

.DESCRIPTION
    This script searches common registry uninstall locations (HKLM/HKCU and Wow6432Node)
    for products whose DisplayName contains 'Microsoft Office' or 'Office' and/or detects
    Click-to-Run (Office 365) via the ClickToRun configuration key. For each discovered
    product the script will attempt to run the product's UninstallString. MSI-based
    products are uninstalled using msiexec. Executable uninstallers are called with
    common quiet arguments when -Force is provided.

    The script prompts before uninstalling each product unless -AutoConfirm is used.

.PARAMETER Force
    Try to perform non-interactive/silent uninstall where possible.

.PARAMETER AutoConfirm
    Skip confirmation prompts and proceed with uninstall attempts.

.EXAMPLES
    .\officetest.ps1
    .\officetest.ps1 -Force -AutoConfirm
#>

param(
    [switch]$Force,
    [switch]$AutoConfirm
)

function Get-OfficeProducts {
    $results = @()

    $regPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    )

    foreach ($path in $regPaths) {
        if (Test-Path $path) {
            Get-ChildItem -Path $path -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    $props = Get-ItemProperty -Path $_.PSPath -ErrorAction Stop
                } catch { return }

                if ($null -ne $props.DisplayName) {
                    if ($props.DisplayName -like '*Microsoft Office*' -or $props.DisplayName -match '\bOffice\b') {
                        $results += [pscustomobject]@{
                            DisplayName    = $props.DisplayName
                            DisplayVersion = $props.DisplayVersion
                            Publisher      = $props.Publisher
                            UninstallString= $props.UninstallString
                            RegPath        = $_.PSPath
                        }
                    }
                }
            }
        }
    }

    # Detect Click-to-Run (Office 365 / Microsoft 365)
    $c2r = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    if (Test-Path $c2r) {
        try {
            $cprops = Get-ItemProperty -Path $c2r -ErrorAction Stop
            $version = $cprops.VersionToReport
            $productIds = $cprops.ProductReleaseIds -join ', '
            $results += [pscustomobject]@{
                DisplayName    = "Microsoft 365 / Office Click-to-Run"
                DisplayVersion = $version
                Publisher      = 'Microsoft Corporation'
                UninstallString= $null
                RegPath        = $c2r
                ProductIds     = $productIds
            }
        } catch { }
    }

    return $results | Sort-Object DisplayName -Unique
}

function Invoke-Uninstall {
    param(
        [Parameter(Mandatory=$true)] [pscustomobject]$Product,
        [switch]$Force
    )

    Write-Host "\n-> Uninstalling: $($Product.DisplayName) $($Product.DisplayVersion)"

    if ($Product.UninstallString) {
        $us = $Product.UninstallString.Trim()

        # If it's an msiexec string, normalize and call msiexec with /x
        if ($us -match '(?i)msiexec') {
            # Extract any GUID or /I{GUID} or /X{GUID}
            $guidMatch = ($us -match '\{[0-9A-Fa-f\-]{36}\}') | Out-Null; $guid = $matches[0]
            if (-not $guid) {
                # Try to find product code-like token
                $guid = ($us -split ' ') | Where-Object { $_ -match '\{[0-9A-Fa-f\-]{36}\}' } | Select-Object -First 1
            }

            if ($guid) {
                $args = "/x $guid"
                if ($Force) { $args += ' /qn /norestart' } else { $args += ' /qb' }
                Write-Host "Running: msiexec $args"
                $p = Start-Process -FilePath 'msiexec.exe' -ArgumentList $args -Wait -PassThru
                return $p.ExitCode
            } else {
                # Fallback: run the original uninstall string but try to inject silent flags
                $exe, $exeArgs = $us -split ' ', 2
                $attemptArgs = $exeArgs
                if ($Force) { $attemptArgs += ' /qn /norestart' }
                Write-Host "Running: $exe $attemptArgs"
                $p = Start-Process -FilePath $exe -ArgumentList $attemptArgs -Wait -PassThru
                return $p.ExitCode
            }
        } else {
            # Executable based uninstaller. Try common silent flags when forced.
            $exe, $exeArgs = $us -split ' ', 2
            if ($Force) {
                $silentFlags = @('/quiet','/qn','/S','/silent','/uninstall')
                foreach ($flag in $silentFlags) {
                    try {
                        Write-Host "Trying: $exe $flag"
                        $p = Start-Process -FilePath $exe -ArgumentList $flag -Wait -PassThru -ErrorAction Stop
                        if ($p.ExitCode -eq 0) { return 0 }
                    } catch {
                        # ignore and try next
                    }
                }

                # Last resort: run original uninstall string
                try {
                    Write-Host "Running original uninstall string: $us"
                    $p = Start-Process -FilePath $exe -ArgumentList $exeArgs -Wait -PassThru
                    return $p.ExitCode
                } catch {
                    Write-Warning "Failed to execute uninstall: $_"
                    return 1
                }
            } else {
                # Not forced: run the original uninstallstring so user sees UI
                try {
                    Write-Host "Running uninstall: $us"
                    $p = Start-Process -FilePath $exe -ArgumentList $exeArgs -Wait -PassThru
                    return $p.ExitCode
                } catch {
                    Write-Warning "Failed to execute uninstall: $_"
                    return 1
                }
            }
        }
    } else {
        # No UninstallString - special handling for Click-to-Run
        if ($Product.RegPath -like '*ClickToRun*') {
            Write-Host "Detected Click-to-Run Office. Automatic uninstall is not supported by this script in all cases."
            Write-Host "Recommended: use the Office Deployment Tool (ODT) with an XML that removes all products, or run Microsoft's Support and Recovery Assistant."
            return 2
        }

        Write-Warning "No UninstallString found for $($Product.DisplayName). Skipping."
        return 1
    }
}

# Main
$found = Get-OfficeProducts

if (-not $found -or $found.Count -eq 0) {
    Write-Host "No Microsoft Office installations detected on this system."
    return 0
}

Write-Host "Detected the following Office-related products:\n"
$i = 1
foreach ($p in $found) {
    Write-Host "[$i] $($p.DisplayName) $($p.DisplayVersion)  (Publisher: $($p.Publisher))"
    $i++
}

foreach ($product in $found) {
    $proceed = $false
    if ($AutoConfirm) { $proceed = $true } else {
        $yn = Read-Host "Do you want to uninstall '$($product.DisplayName)'? (Y/N)"
        if ($yn -match '^[Yy]') { $proceed = $true }
    }

    if ($proceed) {
        $exit = Invoke-Uninstall -Product $product -Force:$Force
        switch ($exit) {
            0 { Write-Host "Uninstall completed successfully for $($product.DisplayName)." }
            1 { Write-Warning "Uninstall reported an error for $($product.DisplayName)." }
            2 { Write-Warning "Special case: Click-to-Run or manual removal required for $($product.DisplayName)." }
            default { Write-Warning "Uninstall finished with exit code $exit for $($product.DisplayName)." }
        }
    } else {
        Write-Host "Skipping $($product.DisplayName)."
    }
}

Write-Host "\nDone. Review messages above for results and next steps."


# Install all programs in the $programs array ########################################################
Foreach ($program in $programs)
    {
    Write-Host "Installing or Updating: ",$program  -ForegroundColor Yellow -BackgroundColor DarkGreen
    winget install -e --id $program -h --silent --accept-package-agreements --accept-source-agreements
    }

$RSATTools = @(
    "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0",
    "Rsat.Dns.Tools~~~~0.0.1.0",
    "Rsat.FileServices.Tools~~~~0.0.1.0"
)

foreach ($tool in $RSATTools) {
    Add-WindowsCapability -Online -Name $tool
}

# Verify RSAT installation ########################################################
Get-WindowsCapability -Name RSAT* -Online | Where-Object State -eq 'Installed'

# Upgrade all installed apps ########################################################
winget upgrade --all

# Clean up Winget cache ########################################################
winget cache purge

# Remove the installer file ########################################################
Remove-Item -Path $wingetInstaller -Force -ErrorAction SilentlyContinue

# Final message ########################################################
Write-Host "Winget installation and updates completed successfully!" -ForegroundColor Green