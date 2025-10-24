# PowerShell Menu System

function Show-Menu {
    Clear-Host
    Write-Host "==============================="
    Write-Host "   PowerShell Menu System"
    Write-Host "==============================="
    Write-Host "1. Show Date and Time"
    Write-Host "2. List Running Processes"
    Write-Host "3. Show Disk Usage"
    Write-Host "4. Run Speedtest"
    Write-Host "99. Restart Computer"
    Write-Host "100. Exit"
    Write-Host "==============================="
}

function RunMenu {
    do {
        Show-Menu
        $choice = Read-Host "Enter your choice"

        switch ($choice) {
            "1" {
                Clear-Host
                Write-Host "Current Date and Time:" -ForegroundColor Cyan
                Get-Date
                .\waitonkeypress.ps1
            }
            "2" {
                Clear-Host
                Write-Host "Running Processes:" -ForegroundColor Cyan
                Get-Process | Sort-Object CPU -Descending | Select-Object -First 10
                .\waitonkeypress.ps1
            }
            "3" {
                Clear-Host
                Write-Host "Disk Usage:" -ForegroundColor Cyan
                Get-PSDrive -PSProvider FileSystem | Select-Object Name,Free,Used,@{Name="Used(%)";Expression={[math]::Round($_.Used/$_.Maximum*100,2)}}
                .\waitonkeypress.ps1
            }
            "4" {
                Clear-Host
                Write-Host "Running Speedtest" -ForegroundColor Red
                .\speedtest.ps1
                Pause
            }
            "99" {
                Clear-Host
                Write-Host "Restarting computer in 10 seconds... (Press Ctrl+C to cancel)" -ForegroundColor Red
                Start-Sleep -Seconds 10
                Restart-Computer -Force
            }
            "100" {
                Write-Host "Exiting menu..." -ForegroundColor Yellow
                break
            }
            Default {
                Write-Host "Invalid selection. Try again." -ForegroundColor Red
                Pause
            }
        }
    } while ($choice -ne "5")
}

# Run the menu
RunMenu
