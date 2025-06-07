# UB500 Full Cleanup + Auto Driver Download Script
# Run this script as Administrator

$driverUrl = "https://static.tp-link.com/upload/download/2021/202108/20210805/UB500(EU)_V1_210526_Windows.zip"
$driverZip = "$env:TEMP\UB500_driver.zip"
$extractPath = "$env:TEMP\UB500_driver"

Write-Host "=== TP-Link UB500 Bluetooth Cleanup & Driver Prep ===" -ForegroundColor Cyan

# STEP 1: Stop Bluetooth Support Service
Write-Host "`n[1/7] Stopping Bluetooth services..." -ForegroundColor Yellow
Stop-Service bthserv -Force -ErrorAction SilentlyContinue

# STEP 2: Remove Hidden Bluetooth Devices
Write-Host "[2/7] Removing ghosted Bluetooth devices..." -ForegroundColor Yellow
pnputil /enum-devices /class Bluetooth | ForEach-Object {
    if ($_ -match "Instance ID: (.+)") {
        $id = $matches[1]
        Write-Host " → Removing: $id"
        pnputil /remove-device "$id" | Out-Null
    }
}

# STEP 3: Remove known Realtek/CSR stacks from Programs and Features
$oldStacks = @(
    "CSR Harmony Wireless Software Stack",
    "Realtek Bluetooth Driver",
    "Realtek Bluetooth"
)
Write-Host "[3/7] Uninstalling known Bluetooth driver stacks..." -ForegroundColor Yellow
Get-WmiObject -Class Win32_Product | Where-Object {
    $oldStacks -contains $_.Name
} | ForEach-Object {
    Write-Host " → Uninstalling: $($_.Name)"
    $_.Uninstall() | Out-Null
}

# STEP 4: Download Latest Driver ZIP
Write-Host "[4/7] Downloading latest driver package..." -ForegroundColor Yellow
Invoke-WebRequest -Uri $driverUrl -OutFile $driverZip -UseBasicParsing

# STEP 5: Extract Driver ZIP
Write-Host "[5/7] Extracting driver to temp folder..." -ForegroundColor Yellow
if (Test-Path $extractPath) { Remove-Item $extractPath -Recurse -Force }
Expand-Archive -Path $driverZip -DestinationPath $extractPath

# STEP 6: Reset Network/Bluetooth stack
Write-Host "[6/7] Resetting network stack (Bluetooth included)..." -ForegroundColor Yellow
netcfg -d

# STEP 7: Prompt for driver install
Write-Host "`n[7/7] Driver prepared at: $extractPath" -ForegroundColor Green
Write-Host " → RUN setup.exe from that folder after reboot to install the UB500 driver"
Start-Process "explorer.exe" -ArgumentList "`"$extractPath`""

# Final Message
Write-Host "`nAll cleanup tasks complete." -ForegroundColor Cyan
Write-Host "Please REBOOT your machine, then install the driver from:"
Write-Host "  → $extractPath\setup.exe" -ForegroundColor Green

Read-Host -Prompt "`nPress ENTER to finish"
