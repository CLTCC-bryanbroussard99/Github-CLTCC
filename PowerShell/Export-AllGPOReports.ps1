# Requires -Modules GroupPolicy
# Run this script as an Administrator on a Domain Controller or a management machine with RSAT installed.

$ExportPath = "C:\GPO_Backup_Reports"
$Timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$TargetDir  = Join-Path $ExportPath $Timestamp

# Create target directory if it doesn't exist
if (-not (Test-Path $TargetDir)) {
    New-Item -ItemType Directory -Path $TargetDir | Out-Null
    Write-Host "Created export directory at: $TargetDir" -ForegroundColor Green
}

Write-Host "Fetching all GPOs from the domain..." -ForegroundColor Cyan
$AllGPOs = Get-GPO -All

foreach ($GPO in $AllGPOs) {
    # Clean GPO name to remove invalid file characters
    $SafeName = $GPO.DisplayName -replace '[\\/:*?"<>|]', '_'
    $OutputFile = Join-Path $TargetDir "$($SafeName).html"
    
    try {
        Write-Host "Exporting report for: $($GPO.DisplayName)..." -ForegroundColor Yellow
        Get-GPOReport -Guid $GPO.Id -ReportType HTML -Path $OutputFile -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to export report for $($GPO.DisplayName). Error: $_"
    }
}

Write-Host "`nAll GPO HTML reports successfully exported to: $TargetDir" -ForegroundColor Green
