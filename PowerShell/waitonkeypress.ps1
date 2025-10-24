
Write-Host ([Environment]::NewLine)
Write-Host "Press any key to continue..."

# Wait for a keypress without echoing the key or requiring Enter
$null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
clear-host
.\menu.ps1