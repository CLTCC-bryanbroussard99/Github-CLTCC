<#
.SYNOPSIS
    Generates a NIST-compliant cryptographically secure password.
.DESCRIPTION
    Uses System.Security.Cryptography.RandomNumberGenerator to produce 
    passwords meeting NIST SP 800-63B recommendations (length-focused, 
    high entropy, non-deterministic).
.PARAMETER Length
    The length of the password. Default is 16 characters (NIST minimum recommended for user passwords).
.PARAMETER Count
    The number of passwords to generate. Default is 1.
#>
[CmdletBinding()]
param (
    [ValidateRange(8, 128)]
    [int]$Length = 16,

    [ValidateRange(1, 100)]
    [int]$Count = 1
)

function Get-SecureNistPassword {
    param (
        [int]$PasswordLength
    )

    # NIST SP 800-63B encourages broad character set support including printable ASCII and spaces.
    $charSet = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()_+-=[]{}|;:,.<>?'
    $charArray = $charSet.ToCharArray()
    $result = New-Object char[] $PasswordLength

    for ($i = 0; $i -lt $PasswordLength; $i++) {
        # Generate a cryptographically secure random integer index
        $randIndex = [System.Security.Cryptography.RandomNumberGenerator]::GetInt32(0, $charArray.Length)
        $result[$i] = $charArray[$randIndex]
    }

    return [string]::new($result)
}

# Output requested count of passwords
1..$Count | ForEach-Object {
    Get-SecureNistPassword -PasswordLength $Length
}