<# 

    Adds NTFS permissions with validation, user listing, and confirmation.

    - Lists available local and/or domain users/groups
    - Validates path and account resolution
    - Confirms changes before applying ACL
#>

# -----------------------------
# List Available Users / Groups
# -----------------------------
Write-Host "`n==== AVAILABLE ACCOUNTS ====" -ForegroundColor Cyan

# List local users
try {
    Write-Host "`nLocal Users:" -ForegroundColor Yellow
    Get-LocalUser | Select-Object -ExpandProperty Name
}
catch {
    Write-Host "Unable to list local users." -ForegroundColor DarkYellow
}

# List local groups
try {
    Write-Host "`nLocal Groups:" -ForegroundColor Yellow
    Get-LocalGroup | Select-Object -ExpandProperty Name
}
catch {
    Write-Host "Unable to list local groups." -ForegroundColor DarkYellow
}

# Attempt to list domain users/groups if AD module is available
if (Get-Module -ListAvailable -Name ActiveDirectory) {
    Import-Module ActiveDirectory -ErrorAction SilentlyContinue

    try {
        Write-Host "`nDomain Users (sample):" -ForegroundColor Yellow
        Get-ADUser -Filter * |
            Select-Object -ExpandProperty SamAccountName
    }
    catch {
        Write-Host "Unable to list domain users." -ForegroundColor DarkYellow
    }

    try {
        Write-Host "`nDomain Groups (sample):" -ForegroundColor Yellow
        Get-ADGroup -Filter * -ResultSetSize 10 |
            Select-Object -ExpandProperty Name
        Write-Host "(Showing first 10 groups)" -ForegroundColor DarkGray
    }
    catch {
        Write-Host "Unable to list domain groups." -ForegroundColor DarkYellow
    }
}
else {
    Write-Host "`nActive Directory module not found. Domain accounts not listed." -ForegroundColor DarkYellow
}

Write-Host "============================`n" -ForegroundColor Cyan

# -----------------------------
# Validate Folder Path
# -----------------------------
do {
    $Path = Read-Host "Enter the folder path (UNC or local)"
    if (-not (Test-Path -Path $Path)) {
        Write-Host "❌ Path does not exist. Please try again." -ForegroundColor Red
    }
} until (Test-Path -Path $Path)

# -----------------------------
# Validate User or Group
# -----------------------------
do {
    $User = Read-Host "Enter the user or group (Domain\User or Computer\User)"
    try {
        $null = New-Object System.Security.Principal.NTAccount($User).Translate(
            [System.Security.Principal.SecurityIdentifier]
        )
        $UserValid = $true
    }
    catch {
        Write-Host "❌ User or group cannot be resolved. Please try again." -ForegroundColor Red
        $UserValid = $false
    }
} until ($UserValid)

# -----------------------------
# Validate Access Rights
# -----------------------------
do {
    $AccessRights = Read-Host "Enter access rights (Read, Modify, FullControl)"
    try {
        $AccessRightsEnum = [System.Security.AccessControl.FileSystemRights]::$AccessRights
        $RightsValid = $true
    }
    catch {
        Write-Host "❌ Invalid access right." -ForegroundColor Red
        $RightsValid = $false
    }
} until ($RightsValid)

# -----------------------------
# Optional Settings (Defaults)
# -----------------------------
$Inheritance = (Read-Host "Inheritance flags [ContainerInherit,ObjectInherit]") `
    ?? "ContainerInherit,ObjectInherit"

$Propagation = (Read-Host "Propagation flags [None]") ?? "None"
$Type        = (Read-Host "Access type [Allow or Deny]") ?? "Allow"

$InheritanceEnum = [System.Security.AccessControl.InheritanceFlags]::$Inheritance
$PropagationEnum = [System.Security.AccessControl.PropagationFlags]::$Propagation
$TypeEnum        = [System.Security.AccessControl.AccessControlType]::$Type

# -----------------------------
# Summary & Confirmation
# -----------------------------
Write-Host "`n==== ACL CHANGE SUMMARY ====" -ForegroundColor Cyan
Write-Host "Path          : $Path"
Write-Host "User/Group    : $User"
Write-Host "Access Rights : $AccessRights"
Write-Host "Inheritance   : $Inheritance"
Write-Host "Propagation   : $Propagation"
Write-Host "Access Type   : $Type"
Write-Host "============================`n" -ForegroundColor Cyan

$Confirm = Read-Host "Apply these permissions? (Y/N)"
if ($Confirm -notin @("Y","y")) {
    Write-Host "⚠ Operation cancelled. No changes were made." -ForegroundColor Yellow
    exit
}

# -----------------------------
# Apply ACL
# -----------------------------
$acl = Get-Acl -Path $Path
$rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
    $User,
    $AccessRightsEnum,
    $InheritanceEnum,
    $PropagationEnum,
    $TypeEnum
)

$acl.AddAccessRule($rule)
Set-Acl -Path $Path -AclObject $acl

Write-Host "✅ ACL successfully applied." -ForegroundColor Green
