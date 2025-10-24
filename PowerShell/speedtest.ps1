######### Absolute monitoring values ########## 
$maxpacketloss = 2 #how much % packetloss until we alert. 
$MinimumDownloadSpeed = 100 #What is the minimum expected download speed in Mbps
$MinimumUploadSpeed = 20 #What is the minimum expected upload speed in Mbps
######### End absolute monitoring values ######
clear-host
write-host "Starting Speedtest..."
#Replace $DownloadURL to latest version of the file's URL. We will only download this file once. 
#Latest version can be found at: https://www.speedtest.net/apps/cli
$DownloadURL = "https://install.speedtest.net/app/cli/ookla-speedtest-1.2.0-win64.zip"
$DownloadLocation = "$($Env:ProgramData)\SpeedtestCLI"
write-host "Downloading SpeedtestCLI if not already present..."
try {
    # Check if download location exists
    $TestDownloadLocation = Test-Path $DownloadLocation
    # If not, create the directory, download and extract the SpeedtestCLI
    if (!$TestDownloadLocation) {
        # Create directory
        new-item $DownloadLocation -ItemType Directory -force
        # Download the zip file
        Invoke-WebRequest -Uri $DownloadURL -OutFile "$($DownloadLocation)\speedtest.zip"
        # Expand zip files and name the file speedtest.zip
        Expand-Archive "$($DownloadLocation)\speedtest.zip" -DestinationPath $DownloadLocation -Force
    }
write-host "SpeedtestCLI is ready."
}
# Catch any errors during download and extraction
catch {
    # Output error message and exit script
    write-host "The download and extraction of SpeedtestCLI failed. Error: $($_.Exception.Message)"
    exit 1
}
# Running speedtest and storing previous results
write-host "Running Speedtest..."
$PreviousResults = if (test-path "$($DownloadLocation)\LastResults.txt") { get-content "$($DownloadLocation)\LastResults.txt" | ConvertFrom-Json }
# Running speedtest
$SpeedtestResults = & "$($DownloadLocation)\speedtest.exe" --format=json --accept-license --accept-gdpr
# Storing last results
$SpeedtestResults | Out-File "$($DownloadLocation)\LastResults.txt" -Force
# Converting results from JSON
$SpeedtestResults = $SpeedtestResults | ConvertFrom-Json
 write-host "Speedtest completed successfully."
#creating object
[PSCustomObject]$SpeedtestObj = @{
    # Converting from bytes per second to megabits per second (Mbps)
    downloadspeed = [math]::Round($SpeedtestResults.download.bandwidth / 1000000 * 8, 2)
    # Converting from bytes per second to megabits per second (Mbps)
    uploadspeed   = [math]::Round($SpeedtestResults.upload.bandwidth / 1000000 * 8, 2)
    # Rounding packetloss to whole number
    packetloss    = [math]::Round($SpeedtestResults.packetLoss)
    # isp
    isp           = $SpeedtestResults.isp
    # Exteranl IP
    ExternalIP    = $SpeedtestResults.interface.externalIp
    # Internal IP
    InternalIP    = $SpeedtestResults.interface.internalIp
    # Server used for test
    UsedServer    = $SpeedtestResults.server.host
    # Results URL
    ResultsURL    = $SpeedtestResults.result.url
    # Jitter in ms
    Jitter        = [math]::Round($SpeedtestResults.ping.jitter)
    # Latency in ms
    Latency       = [math]::Round($SpeedtestResults.ping.latency)
}

# Output results
Write-Output "Speedtest Results:"
# Displaying the results in a table format
Format-Table -InputObject $SpeedtestObj

#Health check
$SpeedtestHealth = @()
#Comparing against previous result. Alerting is download or upload differs more than 20%.
if ($PreviousResults) {
    # Calculating if the previous download or upload speed differs more than 20%
    if ($PreviousResults.download.bandwidth / $SpeedtestResults.download.bandwidth * 100 -le 80) { $SpeedtestHealth += "Download speed difference is more than 20%" }
    # Calculating if the previous download or upload speed differs more than 20%
    if ($PreviousResults.upload.bandwidth / $SpeedtestResults.upload.bandwidth * 100 -le 80) { $SpeedtestHealth += "Upload speed difference is more than 20%" }
}
 
#Comparing against preset variables.
# Alerting if download or upload speed is lower than expected or packetloss is higher than expected.
if ($SpeedtestObj.downloadspeed -lt $MinimumDownloadSpeed) { $SpeedtestHealth += "Download speed is lower than $MinimumDownloadSpeed Mbit/ps" }
# Alerting if upload speed is lower than expected
if ($SpeedtestObj.uploadspeed -lt $MinimumUploadSpeed) { $SpeedtestHealth += "Upload speed is lower than $MinimumUploadSpeed Mbit/ps" }
# Alerting if packetloss is higher than expected
if ($SpeedtestObj.packetloss -gt $MaxPacketLoss) { $SpeedtestHealth += "Packetloss is higher than $maxpacketloss%" }

# 
if (!$SpeedtestHealth) {
    # If no issues found, setting health to healthy
    $SpeedtestHealth = "Healthy"
}
write-host "Speedtest Health Check Result:" 
$SpeedtestHealth | ForEach-Object { Write-Output "- $_" }
.\waitonkeypress.ps1
