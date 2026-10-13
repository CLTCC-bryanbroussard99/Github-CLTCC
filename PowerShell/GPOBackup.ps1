# Import Active Directory module to scan for Domain Controllers
Import-Module ActiveDirectory

# 1. Scan the domain for all available Domain Controllers
Write-Host "Scanning domain for Domain Controllers..." -ForegroundColor Cyan
$DCs = Get-ADDomainController -Filter * | Select-Object Name, HostName, IPv4Address

if ($DCs.Count -eq 0) {
    Write-Error "No Domain Controllers found. Ensure you are running this from a domain-joined machine with AD tools installed."
    return
}

# 2. Present a menu to select the target DC
Write-Host "`nAvailable Domain Controllers:" -ForegroundColor Yellow
for ($i = 0; $i -lt $DCs.Count; $i++) {
    Write-Host "[$($i + 1)] $($DCs[$i].Name) ($($DCs[$i].IPv4Address))"
}

# Prompt the user for their choice and validate it
do {
    $selection = Read-Host "`nSelect a Domain Controller by number (1-(DCs.Count))"
    $index = [int]($selection - 1)
} while ($selection -notmatch '^\d+' -or $index -lt 0 -or $index -ge $DCs.Count)

# Assign the chosen DC Name to a variable
\$TargetDC = \(DCs[\)index].HostName
Write-Host "Targeting Domain Controller: \$TargetDC`n" -ForegroundColor Green

# 3. Define dates and folder paths
$CurrentDate  = Get-Date -Format "yyyy-MM-dd"
$BaseFolder   = "C:\GPOBackups"
$TargetFolder = Join-Path -Path $BaseFolder -ChildPath $CurrentDate

# Create the date-specific subfolder
New-Item -Path $TargetFolder -ItemType Directory -Force | Out-Null

# 4. Execute GPO exports targeting the chosen DC
Write-Host "Backing up GPOs..." -ForegroundColor Cyan
Backup-GPO -All -Path $TargetFolder -Server $TargetDC -Comment "Full GPO backup on $CurrentDate via $TargetDC"

Write-Host "Generating reports..." -ForegroundColor Cyan
Get-GPOReport -All -ReportType HTML -Server $TargetDC -Path (Join-Path $TargetFolder "AllGPOsReport.html")
Get-GPOReport -All -ReportType XML -Server $TargetDC -Path (Join-Path $TargetFolder "AllGPOsReport.xml")

Write-Host "Exporting GPO inventory CSV..." -ForegroundColor Cyan
Get-GPO -All -Server $TargetDC | 
    Select-Object DisplayName, Id, GpoStatus, CreationTime, ModificationTime | 
    Export-Csv -Path (Join-Path $TargetFolder "GPO_Inventory.csv") -NoTypeInformation

Write-Host "`nProcess complete. All files saved to: \$TargetFolder" -ForegroundColor Green
