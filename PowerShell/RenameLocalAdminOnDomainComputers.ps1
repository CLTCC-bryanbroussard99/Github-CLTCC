# 1. Define your new credentials
$NewAdminName = "NewAdminName"
$NewPassword  = ConvertTo-SecureString "YourSecurePassword123!" -AsPlainText -Force

# 2. Get target computer names from Active Directory
# (Modify the -SearchBase or -Filter to target specific OUs if needed)
$Computers = Get-ADComputer -Filter 'Enabled -eq $true' | Select-Object -ExpandProperty Name

# 3. Execute the changes remotely in parallel
Invoke-Command -ComputerName $Computers -ScriptBlock {
    param($NewName, $Password)
    
    try {
        # Find the built-in local admin account by its unique SID suffix (-500)
        $AdminAccount = Get-LocalUser | Where-Object { $_.SID -like "S-1-5-21-*-500" }
        
        if ($AdminAccount) {
            $CurrentName = $AdminAccount.Name
            Write-Output "Processing computer $env:COMPUTERNAME (Current Admin Name: $CurrentName)..."
            
            # Change the password
            Set-LocalUser -Name $CurrentName -Password $Password
            
            # Rename the account
            Rename-LocalUser -Name $CurrentName -NewName $NewName
            
            Write-Output "Successfully updated admin account on $env:COMPUTERNAME"
        } else {
            Write-Error "Could not find built-in local administrator account on $env:COMPUTERNAME"
        }
    } catch {
        Write-Error "Failed to update $env:COMPUTERNAME : $_"
    }
} -ArgumentList $NewAdminName, $NewPassword
