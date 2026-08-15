# Requires ActiveDirectory Module
Import-Module ActiveDirectory

# Configuration

$CsvPath       = "\\dc02\DistroFolder\New Account Creation\users.csv"                           # Path the csv file containing user data
$DomainController = "dc02.cyber312.local"                                                       # Domain Controller FQDN or IP
$DomainName    = "cyber312.local"                                                               # Replace with your actual domain
$TargetOU      = "OU=Students,OU=Domain Users,OU=Cyber312.local Accounts,DC=cyber312,DC=local"  # Replace with your target OU
$DefaultPass   = ConvertTo-SecureString "def@ultP@55w0rd" -AsPlainText -Force                   # Default password for new users (must meet domain password policy)

# Optional: Uncomment below if running under a non-admin workstation login
# $AdminCred = Get-Credential

if (-not (Test-Path -Path $CsvPath)) {
    Write-Error "CSV file not found at $CsvPath" -ErrorAction Stop
}

$Users = Import-Csv -Path $CsvPath
     Write-Host "user: $Users" -ForegroundColor Green

foreach ($User in $Users) {
# Safe string conversion prevents calling .Trim() on null
     $FirstName = ([string]$User.'first name').Trim()
     $LastName  = ([string]$User.'last name').Trim()


    if ([string]::IsNullOrWhiteSpace($FirstName) -or [string]::IsNullOrWhiteSpace($LastName)) {
        Write-Warning "Skipping invalid row with missing name fields."
        continue
    }

    $Username          = "$FirstName.$LastName".ToLower().Replace(" ", "")
    $UserPrincipalName = "$Username@$DomainName"
    $DisplayName       = "$FirstName $LastName"

    # Query the Domain Controller remotely using -Server
    $ADParams = @{ Server = $DomainController }
    if ($AdminCred) { $ADParams['Credential'] = $AdminCred }

    $ExistingUser = Get-ADUser -Filter "SamAccountName -eq '$Username'" @ADParams -ErrorAction SilentlyContinue

    if ($ExistingUser) {
        Write-Warning "User '$Username' already exists in AD. Skipping."
    } else {
        try {
            $NewUserParams = @{
                Server                = $DomainController
                SamAccountName        = $Username
                UserPrincipalName     = $UserPrincipalName
                Name                  = $DisplayName
                GivenName             = $FirstName
                Surname               = $LastName
                DisplayName           = $DisplayName
                AccountPassword       = $DefaultPass
                Enabled               = $true
                ChangePasswordAtLogon = $true
                Path                  = $TargetOU
                ErrorAction           = "Stop"
            }
            if ($AdminCred) { $NewUserParams['Credential'] = $AdminCred }

            New-ADUser @NewUserParams

            Write-Host "Successfully created user remotely: $Username" -ForegroundColor Green
        }
        catch {
            Write-Error "Failed to create user ${Username}: $_"
        }
    }
}