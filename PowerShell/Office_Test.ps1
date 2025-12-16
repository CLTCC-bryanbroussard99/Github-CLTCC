<#
.SYNOPSIS
    Detects Microsoft Office installations and attempts to uninstall them.

.DESCRIPTION
    This script searches common registry uninstall locations (HKLM/HKCU and Wow6432Node)
    for products whose DisplayName contains 'Microsoft Office' or 'Office' and/or detects
    Click-to-Run (Office 365) via the ClickToRun configuration key. For each discovered
    product the script will attempt to run the product's UninstallString. MSI-based
    products are uninstalled using msiexec. Executable uninstallers are called with
    common quiet arguments when -Force is provided.

    The script prompts before uninstalling each product unless -AutoConfirm is used.

.PARAMETER Force
    Try to perform non-interactive/silent uninstall where possible.

.PARAMETER AutoConfirm
    Skip confirmation prompts and proceed with uninstall attempts.

.EXAMPLES
    .\officetest.ps1
    .\officetest.ps1 -Force -AutoConfirm
#>

param(
    [switch]$Force,
    [switch]$AutoConfirm
)

function Get-OfficeProducts {
    $results = @()

    $regPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    )

    foreach ($path in $regPaths) {
        if (Test-Path $path) {
            Get-ChildItem -Path $path -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    $props = Get-ItemProperty -Path $_.PSPath -ErrorAction Stop
                } catch { return }

                if ($null -ne $props.DisplayName) {
                    if ($props.DisplayName -like '*Microsoft Office*' -or $props.DisplayName -match '\bOffice\b') {
                        $results += [pscustomobject]@{
                            DisplayName    = $props.DisplayName
                            DisplayVersion = $props.DisplayVersion
                            Publisher      = $props.Publisher
                            UninstallString= $props.UninstallString
                            RegPath        = $_.PSPath
                        }
                    }
                }
            }
        }
    }

    # Detect Click-to-Run (Office 365 / Microsoft 365)
    $c2r = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    if (Test-Path $c2r) {
        try {
            $cprops = Get-ItemProperty -Path $c2r -ErrorAction Stop
            $version = $cprops.VersionToReport
            $productIds = $cprops.ProductReleaseIds -join ', '
            $results += [pscustomobject]@{
                DisplayName    = "Microsoft 365 / Office Click-to-Run"
                DisplayVersion = $version
                Publisher      = 'Microsoft Corporation'
                UninstallString= $null
                RegPath        = $c2r
                ProductIds     = $productIds
            }
        } catch { }
    }

    return $results | Sort-Object DisplayName -Unique
}

function Invoke-Uninstall {
    param(
        [Parameter(Mandatory=$true)] [pscustomobject]$Product,
        [switch]$Force
    )

    Write-Host "\n-> Uninstalling: $($Product.DisplayName) $($Product.DisplayVersion)"

    if ($Product.UninstallString) {
        $us = $Product.UninstallString.Trim()

        # If it's an msiexec string, normalize and call msiexec with /x
        if ($us -match '(?i)msiexec') {
            # Extract any GUID or /I{GUID} or /X{GUID}
            $guidMatch = ($us -match '\{[0-9A-Fa-f\-]{36}\}') | Out-Null; $guid = $matches[0]
            if (-not $guid) {
                # Try to find product code-like token
                $guid = ($us -split ' ') | Where-Object { $_ -match '\{[0-9A-Fa-f\-]{36}\}' } | Select-Object -First 1
            }

            if ($guid) {
                $args = "/x $guid"
                if ($Force) { $args += ' /qn /norestart' } else { $args += ' /qb' }
                Write-Host "Running: msiexec $args"
                $p = Start-Process -FilePath 'msiexec.exe' -ArgumentList $args -Wait -PassThru
                return $p.ExitCode
            } else {
                # Fallback: run the original uninstall string but try to inject silent flags
                $exe, $exeArgs = $us -split ' ', 2
                $attemptArgs = $exeArgs
                if ($Force) { $attemptArgs += ' /qn /norestart' }
                Write-Host "Running: $exe $attemptArgs"
                $p = Start-Process -FilePath $exe -ArgumentList $attemptArgs -Wait -PassThru
                return $p.ExitCode
            }
        } else {
            # Executable based uninstaller. Try common silent flags when forced.
            $exe, $exeArgs = $us -split ' ', 2
            if ($Force) {
                $silentFlags = @('/quiet','/qn','/S','/silent','/uninstall')
                foreach ($flag in $silentFlags) {
                    try {
                        Write-Host "Trying: $exe $flag"
                        $p = Start-Process -FilePath $exe -ArgumentList $flag -Wait -PassThru -ErrorAction Stop
                        if ($p.ExitCode -eq 0) { return 0 }
                    } catch {
                        # ignore and try next
                    }
                }

                # Last resort: run original uninstall string
                try {
                    Write-Host "Running original uninstall string: $us"
                    $p = Start-Process -FilePath $exe -ArgumentList $exeArgs -Wait -PassThru
                    return $p.ExitCode
                } catch {
                    Write-Warning "Failed to execute uninstall: $_"
                    return 1
                }
            } else {
                # Not forced: run the original uninstallstring so user sees UI
                try {
                    Write-Host "Running uninstall: $us"
                    $p = Start-Process -FilePath $exe -ArgumentList $exeArgs -Wait -PassThru
                    return $p.ExitCode
                } catch {
                    Write-Warning "Failed to execute uninstall: $_"
                    return 1
                }
            }
        }
    } else {
        # No UninstallString - special handling for Click-to-Run
        if ($Product.RegPath -like '*ClickToRun*') {
            Write-Host "Detected Click-to-Run Office. Automatic uninstall is not supported by this script in all cases."
            Write-Host "Recommended: use the Office Deployment Tool (ODT) with an XML that removes all products, or run Microsoft's Support and Recovery Assistant."
            return 2
        }

        Write-Warning "No UninstallString found for $($Product.DisplayName). Skipping."
        return 1
    }
}

# Main
$found = Get-OfficeProducts

if (-not $found -or $found.Count -eq 0) {
    Write-Host "No Microsoft Office installations detected on this system."
    return 0
}

Write-Host "Detected the following Office-related products:\n"
$i = 1
foreach ($p in $found) {
    Write-Host "[$i] $($p.DisplayName) $($p.DisplayVersion)  (Publisher: $($p.Publisher))"
    $i++
}

foreach ($product in $found) {
    $proceed = $false
    if ($AutoConfirm) { $proceed = $true } else {
        $yn = Read-Host "Do you want to uninstall '$($product.DisplayName)'? (Y/N)"
        if ($yn -match '^[Yy]') { $proceed = $true }
    }

    if ($proceed) {
        $exit = Invoke-Uninstall -Product $product -Force:$Force
        switch ($exit) {
            0 { Write-Host "Uninstall completed successfully for $($product.DisplayName)." }
            1 { Write-Warning "Uninstall reported an error for $($product.DisplayName)." }
            2 { Write-Warning "Special case: Click-to-Run or manual removal required for $($product.DisplayName)." }
            default { Write-Warning "Uninstall finished with exit code $exit for $($product.DisplayName)." }
        }
    } else {
        Write-Host "Skipping $($product.DisplayName)."
    }
}

Write-Host "\nDone. Review messages above for results and next steps."

s = Get-Content -Raw -Path 'c:\Users\bryan.broussard.CYBER312\Documents\PowerShell\officetest.ps1'; try { [scriptblock]::Create($s) | Out-Null; Write-Host 'PARSE_OK' } catch { Write-Host 'PARSE_ERROR'; Write-Host $_.Exception.Message; exit 1 }