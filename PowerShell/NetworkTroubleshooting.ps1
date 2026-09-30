function Show-Menu {
    Clear-Host
    Write-Host "=====================================" -ForegroundColor Green
    Write-Host "    ADVANCED NETWORK TOOLKIT" -ForegroundColor Green
    Write-Host "=====================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "1. Show Full IP Configuration"
    Write-Host "2. Ping Google"
    Write-Host "3. Flush DNS Cache"
    Write-Host "4. Release IP Address"
    Write-Host "5. Renew IP Address"
    Write-Host "6. Reset Winsock"
    Write-Host "7. Reset TCP/IP Stack"
    Write-Host "8. Open Network Connections"
    Write-Host "9. Open WiFi Settings"
    Write-Host "10. Open Device Manager"
    Write-Host "11. Open Network Troubleshooter"
    Write-Host "12. Speed Test Ping"
    Write-Host "13. Restart Explorer"
    Write-Host "14. Exit"
    Write-Host ""
}

$host.UI.RawUI.WindowTitle = "Advanced Windows Network Toolkit"

do {
    Show-Menu
    $choice = Read-Host "Enter Option"

    switch ($choice) {
        "1" { 
            Get-NetIPConfiguration -All
            # Or use native command: ipconfig /all
        }
        "2" { 
            Test-Connection google.com
        }
        "3" { 
            Clear-DnsClientCache
            Write-Host "DNS cache flushed." -ForegroundColor Green
        }
        "4" { 
            ipconfig /release
        }
        "5" { 
            ipconfig /renew
        }
        "6" { 
            netsh winsock reset
        }
        "7" { 
            netsh int ip reset
        }
        "8" { 
            ncpa.cpl
        }
        "9" { 
            Start-Process "ms-settings:network-wifi"
        }
        "10" { 
            devmgmt.msc
        }
        "11" { 
            msdt.exe /id NetworkDiagnosticsNetworkAdapter
        }
        "12" { 
            ping.exe 8.8.8.8 -t
        }
        "13" { 
            Stop-Process -Name explorer -Force
            Start-Process explorer
        }
        "14" { 
            break
        }
        default { 
            Write-Host "Invalid option. Please try again." -ForegroundColor Red
        }
    }

    if ($choice -ne "14") {
        Write-Host ""
        Pause
    }
} while ($choice -ne "14")