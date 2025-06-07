<#
.SYNOPSIS
    List & force-remove ghosts/unknowns and chosen devices for Mouse or Keyboard.

.PARAMETER Class
    Which PnP class to operate on: "Mouse" or "Keyboard".

.DESCRIPTION
    • Auto-removes any Status="Unknown" entries in that class.  
    • Shows the rest (FriendlyName + InstanceId).  
    • Prompts you to pick one to force-uninstall via pnputil.exe.  
    • Falls back to disable+remove on Access-Denied.  
    • No Bluetooth-only registry cleanup here.  
    • Must run “As Administrator.”
#>

param(
    [ValidateSet('Mouse','Keyboard')]
    [string]$Class = 'Mouse'
)

function Get-InputDevice {
    <#
    .OUTPUTS
      PSCustomObject with properties:
        • Index        – menu index
        • FriendlyName – nonempty
        • Status       – OK, Unknown, etc.
        • InstanceId   – the string you pass to pnputil.exe
    #>
    $all = Get-PnpDevice -Class $Class -ErrorAction SilentlyContinue
    if (-not $all) { return @() }

    $i = 0
    foreach ($dev in $all) {
        if ($dev.FriendlyName -and $dev.FriendlyName.Trim().Length -gt 0) {
            $i++
            [PSCustomObject]@{
                Index        = $i
                FriendlyName = $dev.FriendlyName
                Status       = $dev.Status
                InstanceId   = $dev.InstanceId
            }
        }
    }
}

# 1) Auto-clean any Unknown-status entries
$unknowns = Get-InputDevice | Where-Object Status -eq 'Unknown'
if ($unknowns.Count) {
    Write-Host "`n--- Removing Unknown-status $Class devices ---`n" -ForegroundColor Yellow
    foreach ($u in $unknowns) {
        Write-Host "→ $($u.FriendlyName) ($($u.InstanceId))" -NoNewline
        & pnputil.exe /remove-device $u.InstanceId /force > $null 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host " removed." -ForegroundColor Green
        } else {
            Write-Host " failed (code $LASTEXITCODE)." -ForegroundColor Red
        }
    }
    Write-Host
}

# 2) Main menu
do {
    $devs = @( Get-InputDevice )
    if (-not $devs.Count) {
        Write-Host "`nNo $Class devices found." -ForegroundColor DarkYellow
        break
    }

    Write-Host "`n**** $Class Devices ****`n"
    $devs | ForEach-Object {
        "{0,3} - {1} ({2})  InstanceId: {3}" -f `
          $_.Index, $_.FriendlyName, $_.Status, $_.InstanceId |
        Write-Host
    }

    $sel = Read-Host "`nSelect a device to remove (0 = exit)"
    if ($sel -notmatch '^\d+$') {
        Write-Host "Enter a number." -ForegroundColor Red; continue
    }
    $idx = [int]$sel
    if ($idx -eq 0) { break }
    if ($idx -lt 1 -or $idx -gt $devs.Count) {
        Write-Host "Choose 1–$($devs.Count) or 0." -ForegroundColor Red; continue
    }

    $dev = $devs | Where-Object Index -eq $idx
    Write-Host "`nRemoving: $($dev.FriendlyName) ($($dev.InstanceId))" -ForegroundColor Yellow

    # Attempt #1: force-remove
    & pnputil.exe /remove-device $dev.InstanceId /force > $null 2>&1
    $ec = $LASTEXITCODE

    if ($ec -eq 0) {
        Write-Host "→ Removed." -ForegroundColor Green
    }
    elseif ($ec -eq 5) {
        # Access Denied → disable then retry
        Write-Host "→ Access Denied; disabling then retrying…" -ForegroundColor Yellow
        & pnputil.exe /disable-device $dev.InstanceId > $null 2>&1
        if ($LASTEXITCODE -eq 0) {
            & pnputil.exe /remove-device $dev.InstanceId /force > $null 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-Host "→ Removed after disable." -ForegroundColor Green
            } else {
                Write-Host "→ Still failed (code $LASTEXITCODE)." -ForegroundColor Red
            }
        }
        else {
            Write-Host "→ Disable failed (code $LASTEXITCODE)." -ForegroundColor Red
        }
    }
    else {
        Write-Host "→ Failed (pnputil exit code $ec)." -ForegroundColor Red
    }

    Start-Sleep -Milliseconds 500

} while ($true)

Write-Host "`nAll done. Press any key to exit…" -NoNewline
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
