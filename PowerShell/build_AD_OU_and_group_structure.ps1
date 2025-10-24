<#
.SYNOPSIS
    Creates a best-practice OU and group structure in Active Directory.

.DESCRIPTION
    This script sets up a logical OU hierarchy, role-based security groups,
    and sample domain local groups for file/resource access using the AGDLP model.

.EXAMPLE
    .\Setup-ADStructure.ps1 -RootOUName "Company" -Departments @("HR","Finance","IT")
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$RootOUName,

    [Parameter(Mandatory = $true)]
    [string[]]$Departments
)

Import-Module ActiveDirectory

Write-Host "Creating best-practice AD structure..." -ForegroundColor Cyan

# Create root OU
$rootOU = "OU=$RootOUName,DC=$(Get-ADDomain).DomainName"
if (-not (Get-ADOrganizationalUnit -LDAPFilter "(ou=$RootOUName)" -ErrorAction SilentlyContinue)) {
    New-ADOrganizationalUnit -Name $RootOUName -ProtectedFromAccidentalDeletion $true
    Write-Host "Created root OU: $RootOUName"
}

# Core OUs under root
$coreOUs = @("_Users","_Computers","_Servers","_Groups","_Admin")
foreach ($ou in $coreOUs) {
    $ouPath = "OU=$ou,$rootOU"
    if (-not (Get-ADOrganizationalUnit -LDAPFilter "(ou=$ou)" -SearchBase $rootOU -ErrorAction SilentlyContinue)) {
        New-ADOrganizationalUnit -Name $ou -Path $rootOU -ProtectedFromAccidentalDeletion $true
        Write-Host "Created OU: $ou"
    }
}

# Create department OUs under _Users and _Computers
foreach ($dept in $Departments) {
    foreach ($parent in "_Users","_Computers") {
        $path = "OU=$parent,OU=$RootOUName,DC=$(Get-ADDomain).DomainName"
        if (-not (Get-ADOrganizationalUnit -LDAPFilter "(ou=$dept)" -SearchBase $path -ErrorAction SilentlyContinue)) {
            New-ADOrganizationalUnit -Name $dept -Path $path -ProtectedFromAccidentalDeletion $false
            Write-Host "Created OU: $dept under $parent"
        }
    }
}

# Create group containers
$groupPath = "OU=_Groups,OU=$RootOUName,DC=$(Get-ADDomain).DomainName"
$subGroups = @("Security","Distribution")
foreach ($sg in $subGroups) {
    if (-not (Get-ADOrganizationalUnit -LDAPFilter "(ou=$sg)" -SearchBase $groupPath -ErrorAction SilentlyContinue)) {
        New-ADOrganizationalUnit -Name $sg -Path $groupPath
        Write-Host "Created OU: $sg under _Groups"
    }
}

# Create AGDLP-style groups for each department
foreach ($dept in $Departments) {
    # Global (role-based) group
    $globalGroup = "GG_${dept}_Users"
    if (-not (Get-ADGroup -Filter "Name -eq '$globalGroup'" -ErrorAction SilentlyContinue)) {
        New-ADGroup -Name $globalGroup -GroupScope Global -GroupCategory Security -Path "OU=Security,$groupPath"
        Write-Host "Created Global group: $globalGroup"
    }

    # Domain Local (resource-based) group
    $localGroup = "DL_${dept}_FileShare_RW"
    if (-not (Get-ADGroup -Filter "Name -eq '$localGroup'" -ErrorAction SilentlyContinue)) {
        New-ADGroup -Name $localGroup -GroupScope DomainLocal -GroupCategory Security -Path "OU=Security,$groupPath"
        Write-Host "Created Domain Local group: $localGroup"
    }

    # Add global → domain local membership (AGDLP principle)
    Add-ADGroupMember -Identity $localGroup -Members $globalGroup -ErrorAction SilentlyContinue
}

Write-Host "`n✅ AD structure created successfully!" -ForegroundColor Green
Write-Host "------------------------------------------------"
Write-Host "Root OU: $RootOUName"
Write-Host "Departments: $($Departments -join ', ')"
Write-Host "Structure includes:"
Write-Host " - User, Computer, Server, Admin, and Group OUs"
Write-Host " - Department sub-OUs"
Write-Host " - AGDLP-based Security Groups (Global + Domain Local)"
Write-Host "------------------------------------------------`n"
