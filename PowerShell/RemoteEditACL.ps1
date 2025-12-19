<# 
.SYNOPSIS
    Safely adds an NTFS permission to a folder with validation and confirmation.

.DESCRIPTION
    This script:
    - Prompts for folder path, user/group, and permissions
    - Validates the path exists
    - Validates the user or group resolves to a SID
    - Confirms the change before applying the ACL
    - Applies NTFS permissions using Set-Acl

.NOTES
    Run in an elevated PowerShell session.
#>

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
        # Attempt to translate the account name to a SID
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
        # Convert string input to FileSystemRights enum
        $AccessRightsEnum = [System.Security.AccessControl.FileSystemRights]::$AccessRights
        $RightsValid = $true
    }
    catch {
        Write-Host "❌ Invalid access right. Valid examples: Read, Modify, FullControl" -ForegroundColor Red
        $RightsValid = $false
    }
} until ($RightsValid)

# -----------------------------
# Optional ACL Settings (Defaults Provided)
# -----------------------------
$InheritanceInput = Read-Host "Enter inheritance flags [ContainerInherit,ObjectInherit]"
$Inheritance = if ($InheritanceInput) { $InheritanceInput } else { "ContainerInherit,ObjectInherit" }

$PropagationInput = Read-Host "Enter propagation flags [None]"
$Propagation = if ($PropagationInput) { $PropagationInput } else { "None" }

$TypeInput = Read-Host "Enter access type [Allow or Deny]"
$Type = if ($TypeInput) { $TypeInput } else { "Allow" }

# Convert optional values to enums
try {
    $InheritanceEnum = [System.Security.AccessControl.InheritanceFlags]::$Inheritance
    $PropagationEnum = [System.Security.AccessControl.PropagationFlags]::$Propagation
    $TypeEnum        = [System.Security.AccessControl.AccessControlType]::$Type
}
catch {
    Write-Error "One or more ACL flag values are invalid. Script cannot continue."
    exit 1
}

# -----------------------------
# Display Summary and Confirm
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
if ($Confirm -notin @("Y", "y")) {
    Write-Host "⚠ Operation cancelled. No changes were made." -ForegroundColor Yellow
    exit
}

# -----------------------------
# Retrieve and Modify ACL
# -----------------------------
$acl = Get-Acl -Path $Path

$accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
    $User,
    $AccessRightsEnum,
    $InheritanceEnum,
    $PropagationEnum,
    $TypeEnum
)

# Add the new rule to the ACL
$acl.AddAccessRule($accessRule)

# -----------------------------
# Apply ACL to the Folder
# -----------------------------
Set-Acl -Path $Path -AclObject $acl

Write-Host "✅ ACL successfully applied to $Path for $User" -ForegroundColor Green
