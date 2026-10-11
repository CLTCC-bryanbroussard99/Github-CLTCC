<#
.SYNOPSIS
    Enables Remote Desktop (RDP) domain-wide across target Active Directory computers.

.DESCRIPTION
    This script queries Active Directory for domain-joined Windows computers (Servers and/or Workstations),
    tests network connectivity, and remotely configures each target to enable Remote Desktop.
    
    It performs the following actions on each target:
    1. Sets fDenyTSConnections to 0 in the Registry (Enables RDP).
    2. Configures the 'TermService' service startup type to Automatic and ensures it is running.
    3. Enables the predefined Windows Defender Firewall rules for Remote Desktop.

.PARAMETER TargetOU
    Optional Distinguished Name (DN) of an Organizational Unit to limit the scope of target computers.
    If omitted, the script queries all enabled domain computers.

.PARAMETER IncludeWorkstations
    Switch parameter. If specified, the script targets both Windows Servers and Windows Workstations.
    By default, it targets only Windows Servers for safety.

.PARAMETER ThrottleLimit
    The maximum number of concurrent remote connections allowed. Default is 32.

.EXAMPLE
    PS C:\> .\Enable-DomainRDP.ps1 -Verbose
    Enables RDP across all enabled domain servers in the Active Directory domain.

.EXAMPLE
    PS C:\> .\Enable-DomainRDP.ps1 -TargetOU "OU=Servers,DC=contoso,DC=com" -IncludeWorkstations -Verbose
    Enables RDP for all enabled computers within the specified OU.

.NOTES
    Environment: Windows 11, Windows Server
    Prerequisites:
        - Must be run from an elevated PowerShell session (Run as Administrator).
        - Must be run by a user with Domain Admin or local administrator privileges on target machines.
        - ActiveDirectory PowerShell module must be installed.
        - PowerShell Remoting (WinRM) must be enabled on target machines.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false, HelpMessage = "Distinguished Name of an OU to target specific computers.")]
    [ValidateNotNullOrEmpty()]
    [string]$TargetOU,

    [Parameter(Mandatory = $false, HelpMessage = "Include Windows Workstations (e.g., Windows 10/11) along with Servers.")]
    [switch]$IncludeWorkstations,

    [Parameter(Mandatory = $false, HelpMessage = "Maximum concurrent connections for Invoke-Command.")]
    [ValidateRange(1, 128)]
    [int]$ThrottleLimit = 32
)

# Set strict execution mode and default error action
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ==========================================================
# MODULE DEPENDENCY CHECK
# ==========================================================
Write-Verbose "Checking for ActiveDirectory module installation..."
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Error "The 'ActiveDirectory' module is required to run this script. Please install RSAT / ActiveDirectory module." -Category ResourceUnavailable
    return
}

# Ensure the module is loaded
Import-Module ActiveDirectory -ErrorAction Stop

# ==========================================================
# SCRIPT BLOCK FOR REMOTE EXECUTION
# ==========================================================
# This script block will execute remotely on each target host.
$RDPEnableScriptBlock = {
    [CmdletBinding()]
    param()

    $Result = [PSCustomObject]@{
        ComputerName   = $env:COMPUTERNAME
        RegistryUpdate = "Failed"
        ServiceConfig  = "Failed"
        FirewallConfig = "Failed"
        Status         = "Success"
        ErrorDetails   = ""
    }

    try {
        # 1. Enable RDP in the Windows Registry
        # Path: HKLM:\System\CurrentControlSet\Control\Terminal Server
        # Value: fDenyTSConnections (0 = Allow RDP, 1 = Deny RDP)
        Write-Verbose "[$env:COMPUTERNAME] Setting fDenyTSConnections registry entry to 0..."
        $RegPath = "HKLM:\System\CurrentControlSet\Control\Terminal Server"
        Set-ItemProperty -Path $RegPath -Name "fDenyTSConnections" -Value 0 -Type DWord -ErrorAction Stop
        $Result.RegistryUpdate = "Success"

        # 2. Configure and Start Remote Desktop Service (TermService)
        Write-Verbose "[$env:COMPUTERNAME] Setting TermService to Automatic and starting..."
        Set-Service -Name "TermService" -StartupType Automatic -ErrorAction Stop
        $TermService = Get-Service -Name "TermService" -ErrorAction Stop
        if ($TermService.Status -ne 'Running') {
            Start-Service -Name "TermService" -ErrorAction Stop
        }
        $Result.ServiceConfig = "Success"

        # 3. Enable Inbound Firewall Rules for Remote Desktop
        Write-Verbose "[$env:COMPUTERNAME] Enabling Windows Defender Firewall rules for Remote Desktop..."
        # Enable predefined Firewall Rule Group for RDP across all profiles
        Enable-NetFirewallRule -DisplayGroup "Remote Desktop" -Enabled True -ErrorAction Stop
        $Result.FirewallConfig = "Success"

    }
    catch {
        $Result.Status = "Failed"
        $Result.ErrorDetails =$_.Exception.Message
        Write-Error "[$env:COMPUTERNAME] Failed to configure RDP: $($_.Exception.Message)"
    }

    return $Result
}

# ==========================================================
# ACTIVE DIRECTORY TARGET RETRIEVAL
# ==========================================================
try {
    Write-Verbose "Building Active Directory computer query parameters..."
    
    # Filter only enabled computer accounts
    $ADFilter = 'Enabled -eq$true'
    if (-not $IncludeWorkstations) {
        # Default behavior: Limit to Windows Server OS
        $ADFilter += ' -and OperatingSystem -like "*Server*"'
        Write-Verbose "Scope: Windows Server operating systems only."
    } else {
        Write-Verbose "Scope: All Windows operating systems (Servers + Workstations)."
    }

    $GetADParams = @{
        Filter     = $ADFilter
        Properties = 'Name', 'DNSHostName', 'OperatingSystem'
    }

    if ($PSBoundParameters.ContainsKey('TargetOU')) {
        Write-Verbose "Filtering query scope to OU: $TargetOU"
        $GetADParams['SearchBase'] =$TargetOU
    }

    Write-Verbose "Querying Active Directory for target computers..."
    $TargetComputers = Get-ADComputer @GetADParams -ErrorAction Stop

    if (-not $TargetComputers or$TargetComputers.Count -eq 0) {
        Write-Warning "No active computer accounts were returned matching the specified criteria."
        return
    }

    Write-Verbose "Found $($TargetComputers.Count) target computer(s) in Active Directory."

}
catch {
    Write-Error "Failed to query Active Directory: $($_.Exception.Message)"
    return
}

# Extract valid DNS hostnames/names
$ComputerList =$TargetComputers | ForEach-Object {
    if ($_.DNSHostName) { $_.DNSHostName } else {$_.Name }
}

# ==========================================================
# PRE-FLIGHT CONNECTIVITY CHECK
# ==========================================================
Write-Verbose "Performing network ping checks on target hosts..."
$OnlineComputers = [System.Collections.Generic.List[string]]::new()$OfflineComputers = [System.Collections.Generic.List[string]]::new()

foreach ($Computer in$ComputerList) {
    if (Test-Connection -ComputerName $Computer -Count 1 -Quiet) {
        $OnlineComputers.Add($Computer)
    } else {
        Write-Verbose "Host unreachable via ICMP Ping: $Computer"
        $OfflineComputers.Add($Computer)
    }
}

Write-Verbose "Connectivity Summary: $($OnlineComputers.Count) Online | $($OfflineComputers.Count) Offline"

if ($OfflineComputers.Count -gt 0) {
    Write-Warning "The following $($OfflineComputers.Count) host(s) were offline or unreachable and will be skipped:"
    $OfflineComputers \vert{} ForEach-Object { Write-Warning " - $_" }
}

if ($OnlineComputers.Count -eq 0) {
    Write-Warning "No target computers are currently reachable via network ping. Exiting."
    return
}

# ==========================================================
# REMOTE EXECUTION & DISPATCH
# ==========================================================
if ($PSCmdlet.ShouldProcess("Active Directory ($($OnlineComputers.Count) computers)", "Enable Remote Desktop (RDP) domain-wide")) {
    
    Write-Verbose "Initiating remote execution across $($OnlineComputers.Count) host(s)..."
    
    $ExecutionResults = [System.Collections.Generic.List[PSCustomObject]]::new()

    try {
        # Invoke-Command executes the script block across target machines concurrently
        $RemoteOutput = Invoke-Command -ComputerName$OnlineComputers `
                                       -ScriptBlock $RDPEnableScriptBlock `
                                       -ThrottleLimit $ThrottleLimit `
                                       -ErrorVariable RemoteErrors `
                                       -ErrorAction SilentlyContinue

        if ($RemoteOutput) {
            foreach ($Item in$RemoteOutput) {
                $ExecutionResults.Add($Item)
            }
        }

        # Catch WinRM/Connection level failures captured in ErrorVariable
        if ($RemoteErrors) {
            foreach ($Err in$RemoteErrors) {
                $FailedTarget =$Err.TargetObject
                Write-Warning "Failed to establish WinRM session with ${FailedTarget}: $($Err.Exception.Message)"
                $ExecutionResults.Add([PSCustomObject]@{
                    ComputerName   = $FailedTarget
                    RegistryUpdate = "N/A"
                    ServiceConfig  = "N/A"
                    FirewallConfig = "N/A"
                    Status         = "Connection Failed"
                    ErrorDetails   = $Err.Exception.Message
                })
            }
        }
    }
    catch {
        Write-Error "An unhandled exception occurred during remote invocation: $($_.Exception.Message)"
    }
    finally {
        # ==========================================================
        # REPORTING & OUTPUT SUMMARY
        # ==========================================================
        Write-Verbose "Processing deployment execution results..."
        
        # Display aggregated result summary table to stdout
        $ExecutionResults | Format-Table -Property ComputerName, Status, RegistryUpdate, ServiceConfig, FirewallConfig, ErrorDetails -AutoSize

        $SuccessCount = ($ExecutionResults \vert{} Where-Object {$_.Status -eq "Success" }).Count
        $FailureCount = ($ExecutionResults \vert{} Where-Object {$_.Status -ne "Success" }).Count

        Write-Host "==========================================================" -ForegroundColor Cyan
        Write-Host " RDP DOMAIN DEPLOYMENT SUMMARY" -ForegroundColor Cyan
        Write-Host "==========================================================" -ForegroundColor Cyan
        Write-Host " Total Targets Processed : $($ExecutionResults.Count)"
        Write-Host " Successfully Configured : $SuccessCount" -ForegroundColor Green
        Write-Host " Failed / Unreachable    : $($FailureCount +$OfflineComputers.Count)" -ForegroundColor Red
        Write-Host "==========================================================" -ForegroundColor Cyan
    }
}