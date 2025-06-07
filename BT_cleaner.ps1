<#
.SYNOPSIS
    Unpair Bluetooth devices, then clean up all registry remnants (Devices, Keys, DeviceCache, Migration).

.DESCRIPTION
    • Auto-removes any “Unknown” devices.  
    • Lets you pick a remaining paired device to force-remove via pnputil.exe.  
    • After removal, cleans registry under:  
        – HKLM:\SYSTEM\CurrentControlSet\Services\BTHPORT\Parameters\Devices\<MAC>  
        – HKLM:\SYSTEM\CurrentControlSet\Services\BTHPORT\Parameters\Keys\<MAC>  
        – HKLM:\SOFTWARE\Microsoft\Bluetooth\DeviceCache\… (by Address)  
        – HKLM:\SYSTEM\Setup\Upgrade\PnP\…\DeviceMigration\Devices\BTHENUM\DEV_<MAC>  
    • Must be run “As Administrator.”
#>

#------------------------------------------
# Cleanup function: removes all registry entries for a given InstanceId
#------------------------------------------
Function Clean-BTRegistry {
    param (
        [string]$InstanceId
    )
    # Extract the MAC (hex) from "BTHENUM\DEV_<MAC>\..."
    if ($InstanceId -match 'BTHENUM\\DEV_([^\\]+)') {
        $mac = $Matches[1].ToUpper()

        # 1) BTHPORT\Parameters\Devices
        $devPath = "HKLM:\SYSTEM\CurrentControlSet\Services\BTHPORT\Parameters\Devices\$mac"
        if (Test-Path $devPath) {
            Remove-Item -Path $devPath -Recurse -Force
            Write-Host "Deleted registry key: $devPath"
        }

        # 2) BTHPORT\Parameters\Keys
        $keyPath = "HKLM:\SYSTEM\CurrentControlSet\Services\BTHPORT\Parameters\Keys\$mac"
        if (Test-Path $keyPath) {
            Remove-Item -Path $keyPath -Recurse -Force
            Write-Host "Deleted registry key: $keyPath"
        }

        # 3) SOFTWARE\Microsoft\Bluetooth\DeviceCache
        $cacheRoot = 'HKLM:\SOFTWARE\Microsoft\Bluetooth\DeviceCache'
        if (Test-Path $cacheRoot) {
            Get-ChildItem -Path $cacheRoot | ForEach-Object {
                $props = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
                if ($props.Address) {
                    # Normalize both addresses (strip separators)
                    $addrNorm = ($props.Address -replace '[:\-]','').ToUpper()
                    if ($addrNorm -eq $mac) {
                        Remove-Item -Path $_.PSPath -Recurse -Force
                        Write-Host "Deleted DeviceCache entry: $($_.PSPath)"
                    }
                }
            }
        }

        # 4) DeviceMigration leftovers
        $migRoot = 'HKLM:\SYSTEM\Setup\Upgrade\PnP\CurrentControlSet\Control\DeviceMigration\Devices\BTHENUM'
        $migKey  = "DEV_$mac"
        if (Test-Path "$migRoot\$migKey") {
            Remove-Item -Path "$migRoot\$migKey" -Recurse -Force
            Write-Host "Deleted migration entry: $migRoot\$migKey"
        }
    }
}

#------------------------------------------
# Function to list paired peripherals only
#------------------------------------------
Function Get-BTDevice {
    $all = Get-PnpDevice -Class Bluetooth -ErrorAction SilentlyContinue
    if (-not $all) { return @() }

    $idx = 0
    foreach ($dev in $all) {
        if (
            $dev.FriendlyName -and 
            $dev.FriendlyName.Trim().Length -gt 0 -and
            $dev.InstanceId -like 'BTHENUM\*'
        ) {
            $idx++
            [PSCustomObject]@{
                Index        = $idx
                FriendlyName = $dev.FriendlyName
                Status       = $dev.Status
                InstanceId   = $dev.InstanceId
            }
        }
    }
}

#------------------------------------------
# 1) Remove any "Unknown" devices silently
#------------------------------------------
$unknowns = Get-BTDevice | Where-Object Status -eq 'Unknown'
if ($unknowns.Count) {
    Write-Host "`n--- Cleaning up Unknown-status devices ---`n"
    foreach ($u in $unknowns) {
        Write-Host "Removing [Unknown]: $($u.FriendlyName) ($($u.InstanceId))" -ForegroundColor Yellow
        & pnputil.exe /remove-device $u.InstanceId /force 2>&1 | Out-Null
        Clean-BTRegistry -InstanceId $u.InstanceId
    }
    Write-Host "`nUnknown devices cleaned.`n"
}

#------------------------------------------
# 2) Main menu: show remaining devices, remove chosen one
#------------------------------------------
do {
    $devices = @( Get-BTDevice )
    if (-not $devices.Count) {
        Write-Host "`nNo paired Bluetooth devices found." -ForegroundColor DarkYellow
        break
    }

    Write-Host "`n**** Paired Bluetooth Devices ****`n"
    $devices | ForEach-Object {
        ("{0,3} - {1} ({2})  InstanceId: {3}" -f 
          $_.Index, $_.FriendlyName, $_.Status, $_.InstanceId
        ) | Write-Host
    }

    $sel = Read-Host "`nSelect a device to remove (0 = exit)"
    if ($sel -notmatch '^\d+$') {
        Write-Host "Enter a valid number." -ForegroundColor Red; continue
    }
    $i = [int]$sel
    if ($i -eq 0) { break }
    if ($i -lt 1 -or $i -gt $devices.Count) {
        Write-Host "Choose between 0 and $($devices.Count)." -ForegroundColor Red; continue
    }

    $dev = $devices | Where-Object Index -eq $i
    Write-Host "`nRemoving: $($dev.FriendlyName) ($($dev.InstanceId))" -ForegroundColor Yellow

    # Force-uninstall via pnputil
    & pnputil.exe /remove-device $dev.InstanceId /force 2>&1 > $null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "→ PnP node removed." -ForegroundColor Green
    } else {
        Write-Host "→ Pnputil exit code $LASTEXITCODE" -ForegroundColor Red
    }

    # Then scrub registry
    Clean-BTRegistry -InstanceId $dev.InstanceId

    Start-Sleep -Milliseconds 500
} while ($true)

Write-Host "`nCleanup complete. Press any key to exit..." -NoNewline
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
