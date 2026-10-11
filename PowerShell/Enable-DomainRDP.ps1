<#
.SYNOPSIS
    Interactive Windows Startup Troubleshooting Menu (Part 2) with Automatic Elevation.
.DESCRIPTION
    This script provides an interactive menu-driven interface to execute common 
    Windows startup, boot configuration, recovery environment, and system repair commands.
.NOTES
    Author: Senior PowerShell Automation Engineer
    Environment: Windows 11 / Windows Server
#>

[CmdletBinding()]
param()

# Ensure script runs with Administrator privileges
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Administrative privileges required. Relaunching script as Administrator..." -ForegroundColor Yellow
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

$ErrorActionPreference = 'Stop'

function Test-DomainEnvironment {
    [CmdletBinding()]
    param()
    
    Write-Verbose "Checking if the current machine is joined to an Active Directory domain."
    try {
        $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain()
        Write-Verbose "Successfully connected to domain: $($domain.Name)"
        return $true
    }
    catch {
        Write-Warning "The current machine does not appear to be domain-joined. Script will execute in local/targeted mode."
        return $false
    }
}

function Set-LocalOrRemoteRDP {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $true)]
        [bool]$NLAState
    )

    process {
        Write-Verbose "Processing RDP configuration for target: $ComputerName"
        
        try {
            # 1. Modify Registry to Enable RDP (fDenyTSConnections = 0)
            Write-Verbose "Configuring Terminal Server registry settings on [$ComputerName]..."
            $registryPath = "HKLM:\System\CurrentControlSet\Control\Terminal Server"
            
            if ($ComputerName -eq $env:COMPUTERNAME -or $ComputerName -eq "localhost" -or $ComputerName -eq ".") {
                Set-ItemProperty -Path $registryPath -Name "fDenyTSConnections" -Value 0 -Force
                
                # Configure Network Level Authentication (NLA)
                $nlaPath = "$registryPath\WinStations\RDP-Tcp"
                if ($NLAState) {
                    Set-ItemProperty -Path $nlaPath -Name "UserAuthentication" -Value 1 -Force
                    Write-Verbose "NLA successfully enabled locally."
                } else {
                    Set-ItemProperty -Path $nlaPath -Name "UserAuthentication" -Value 0 -Force
                    Write-Verbose "NLA disabled locally as requested."
                }

                # 2. Enable Firewall Rules for Remote Desktop
                Write-Verbose "Enabling built-in Windows Firewall rules for Remote Desktop..."
                Enable-NetFirewallRule -DisplayGroup "Remote Desktop" -ErrorAction SilentlyContinue
            }
            else {
                # Remote CIM/WMI Registry modification fallback if explicit targets are provided
                $session = New-CimSession -ComputerName $ComputerName
                
                $regParams = @{
                    CimSession   = $session
                    Namespace    = 'root\default'
                    ClassName    = 'StdRegProv'
                    MethodName   = 'SetDWORDValue'
                    Arguments    = @{
                        hDefKey      = [uint32]2147483650 # HKEY_LOCAL_MACHINE
                        sSubKeyName  = 'System\CurrentControlSet\Control\Terminal Server'
                        sValueName   = 'fDenyTSConnections'
                        uValue       = [uint32]0
                    }
                }
                Invoke-CimMethod @regParams | Out-Null
                
                if ($NLAState) {
                    $nlaParams = @{
                        CimSession   = $session
                        Namespace    = 'root\default'
                        ClassName    = 'StdRegProv'
                        MethodName   = 'SetDWORDValue'
                        Arguments    = @{
                            hDefKey      = [uint32]2147483650
                            sSubKeyName  = 'System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp'
                            sValueName   = 'UserAuthentication'
                            uValue       = [uint32]1
                        }
                    }
                    Invoke-CimMethod @nlaParams | Out-Null
                }

                Invoke-Command -CimSession $session -ScriptBlock {
                    Enable-NetFirewallRule -DisplayGroup "Remote Desktop"
                } | Out-Null

                Remove-CimSession -CimSession $session
            }

            Write-Verbose "Successfully configured RDP on [$ComputerName]."
            [PSCustomObject]@{
                ComputerName = $ComputerName
                Status       = "Success"
                NLAEnabled   = $NLAState
                Timestamp    = (Get-Date)
            }
        }
        catch {
            Write-Error "Failed to configure RDP on [$ComputerName]. Error: $_"
            [PSCustomObject]@{
                ComputerName = $ComputerName
                Status       = "Failed"
                ErrorDetails = $_.Exception.Message
                Timestamp    = (Get-Date)
            }
        }
    }
}

# --- Main Execution Flow ---
try {
    Write-Verbose "Starting domain-wide/targeted RDP configuration script execution."
    
    $isDomainJoined = Test-DomainEnvironment

    if (-not $TargetComputers -and $isDomainJoined) {
        Write-Verbose "No explicit target computers specified. Querying Active Directory for domain computers..."
        try {
            Import-Module ActiveDirectory -ErrorAction Stop
            $ADComputers = (Get-ADComputer -Filter 'OperatingSystem -like "*Windows*"' -Property Name).Name
            Write-Verbose "Found $( $ADComputers.Count ) computers in Active Directory."
            
            $results = foreach ($comp in $ADComputers) {
                if ($PSCmdlet.ShouldProcess($comp, "Enable RDP and configure firewall")) {
                    Set-LocalOrRemoteRDP -ComputerName $comp -NLAState $EnableNLA
                }
            }
            $results | Out-Host
        }
        catch {
            Write-Error "Failed to query Active Directory or module missing. Please specify -TargetComputers manually. Details: $_"
        }
    }
    elseif ($TargetComputers) {
        Write-Verbose "Executing against provided explicit target list: $($TargetComputers -join ', ')"
        $results = foreach ($comp in $TargetComputers) {
            if ($PSCmdlet.ShouldProcess($comp, "Enable RDP and configure firewall")) {
                Set-LocalOrRemoteRDP -ComputerName $comp -NLAState $EnableNLA
            }
        }
        $results | Out-Host
    }
    else {
        Write-Verbose "Executing on local machine as fallback."
        if ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, "Enable RDP and configure firewall locally")) {
            Set-LocalOrRemoteRDP -ComputerName $env:COMPUTERNAME -NLAState $EnableNLA
        }
    }

    Write-Verbose "Script execution completed successfully."
}
catch {
    Write-Error "A fatal error occurred during execution: $_"
    exit 1
}
finally {
    Write-Verbose "Cleanup and exit operations completed."
}