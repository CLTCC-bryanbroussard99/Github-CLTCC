# Define the path to your downloaded installer
$InstallerPath = "\\dc02\DistroFolder\Packet Tracer\Packet_Tracer822_64bit_setup_signed.exe"

# Execute silent installation parameters
Start-Process -FilePath $InstallerPath -ArgumentList "/VERYSILENT /NORESTART" -Wait -NoNewWindow
