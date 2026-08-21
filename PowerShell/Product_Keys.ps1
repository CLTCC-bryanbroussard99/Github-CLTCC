function Get-ServerProductKeys {
    [CmdletBinding()]
    param()

    process {
        Write-Host "=============================================" -ForegroundColor Cyan
        Write-Host "   Windows Server Product Key Auditor        " -ForegroundColor Cyan
        Write-Host "=============================================" -ForegroundColor Cyan
        Write-Host ""

        # --- OPTION 1: UEFI / BIOS Motherboard OEM Key ---
        Write-Host "[*] Checking motherboard firmware (OEM Key)..." -ForegroundColor Yellow
        try {
            $OemKey = (Get-CimInstance -Query 'Select * from SoftwareLicensingService' -ErrorAction Stop).OA3xOriginalProductKey
            if (-not [string]::IsNullOrWhiteSpace($OemKey)) {
                Write-Host "-> Found OEM Key: $OemKey" -ForegroundColor Green
            } else {
                Write-Host "-> No OEM key embedded in firmware." -ForegroundColor Gray
            }
        } catch {
            Write-Host "-> Failed to read firmware licensing service." -ForegroundColor Red
        }
        Write-Host ""

        # --- OPTION 2: Clear-text Registry Backup ---
        Write-Host "[*] Checking registry for plain-text backup key..." -ForegroundColor Yellow
        $BackupPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform'
        if (Test-Path $BackupPath) {
            $BackupKey = Get-ItemProperty -Path $BackupPath -Name 'BackupProductKeyDefault' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty BackupProductKeyDefault -ErrorAction SilentlyContinue
            if (-not [string]::IsNullOrWhiteSpace($BackupKey) -and $BackupKey -ne "BBBBB-BBBBB-BBBBB-BBBBB-BBBBB") {
                Write-Host "-> Found Backup Key: $BackupKey" -ForegroundColor Green
            } else {
                Write-Host "-> No plain-text backup key found (or placeholder detected)." -ForegroundColor Gray
            }
        } else {
            Write-Host "-> Registry path for backup key does not exist." -ForegroundColor Gray
        }
        Write-Host ""

        # --- OPTION 3: Encrypted Registry Key Decoder ---
        Write-Host "[*] Decoding primary registry DigitalProductId..." -ForegroundColor Yellow
        $RegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
        if (Test-Path $RegPath) {
            $DigitalProductId = (Get-ItemProperty -Path $RegPath -Name DigitalProductId -ErrorAction SilentlyContinue).DigitalProductId
            
            if ($DigitalProductId) {
                # Setup decoding arrays and structures
                $KeyOffset = 52
                $Characters = "BCDFGHJKMPQRTVWXY2346789"
                
                # Check for Windows 8 / Server 2012 and higher flag
                $IsWin8OrHigher = ([int]($DigitalProductId[66] / 6) -band 1)
                $DigitalProductId[66] = [byte] (($DigitalProductId[66] -and 247) -or (($IsWin8OrHigher -and 2) -shl 4))
                
                $i = 24
                $KeyString = ""
                
                while ($i -ge 0) {
                    $Current = 0
                    $x = 14
                    while ($x -ge 0) {
                        $Current = $Current * 256
                        $Current = $DigitalProductId[$x + $KeyOffset] + $Current
                        $DigitalProductId[$x + $KeyOffset] = [byte]([math]::Floor($Current / 24))
                        $Current = $Current % 24
                        $x--
                    }
                    $KeyString = $Characters[$Current] + $KeyString
                    $i--
                }
                
                # Format string injection for modern Windows kernels
                if ($IsWin8OrHigher -eq 1) {
                    $FirstCharOffset = $Current
                    $KeyString = $KeyString.Substring(1, $FirstCharOffset) + "N" + $KeyString.Substring($FirstCharOffset + 1)
                }
                
                # Inject standard 5x5 product key hyphens
                $FinalKey = ""
                for ($Position = 0; $Position -lt 25; $Position += 5) {
                    $FinalKey += $KeyString.Substring($Position, 5) + "-"
                }
                $DecodedKey = $FinalKey.TrimEnd("-")

                if ($DecodedKey -eq "BBBBB-BBBBB-BBBBB-BBBBB-BBBBB") {
                    Write-Host "-> Decoded Key: BBBBB-BBBBB-BBBBB-BBBBB-BBBBB" -ForegroundColor DarkYellow
                    Write-Host "   ℹ️ Note: This server is activated via a Volume License (KMS / ADBA)." -ForegroundColor Gray
                    Write-Host "     No unique retail/OEM key is physically stored on this machine." -ForegroundColor Gray
                } else {
                    Write-Host "-> Decoded Active Key: $DecodedKey" -ForegroundColor Green
                }
            } else {
                Write-Host "-> DigitalProductId binary entry missing from registry." -ForegroundColor Red
            }
        } else {
            Write-Host "-> Core registry path not found." -ForegroundColor Red
        }
        Write-Host ""
        Write-Host "=============================================" -ForegroundColor Cyan
    }
}

# Run the unified tool
Get-ServerProductKeys
