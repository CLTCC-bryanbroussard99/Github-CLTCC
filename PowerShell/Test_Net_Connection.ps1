while ($true) {
    # Prompt with defaults
    $InputHost = Read-Host "Enter hostname [google.com] to test (or type 'q' to quit): "
    if ($InputHost -eq 'q') {
        Write-Host "Exiting..." -ForegroundColor Yellow
        break
    }

    $Hostname = if ($InputHost) { $InputHost } else { "google.com" }

    $InputPort = Read-Host "Enter port [443]: "
    $Port = if ($InputPort) { $InputPort } else { 443 }

    Write-Host "Testing connection to $Hostname on port $Port..." -ForegroundColor Cyan

    # Execute
    Test-NetConnection -ComputerName $Hostname -Port $Port
}

