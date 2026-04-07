#Requires -Version 5.1
$ErrorActionPreference = "Stop"

$comAvailable = $false
$exoConnected = $false

Write-Host "Connecting to Outlook to disable all rules..." -ForegroundColor Cyan
Write-Host "WARNING: This modifies active Outlook state. Uncomment the `$rule.Enabled = `$false line (or Disable-InboxRule) to actually apply." -ForegroundColor Yellow

try {
    $outlook = New-Object -ComObject Outlook.Application
    $store = $outlook.Session.DefaultStore
    $rules = $store.GetRules()
    $comAvailable = $true
} catch {
    Write-Host "Outlook COM rules unavailable." -ForegroundColor Yellow
}

if ($comAvailable) {
    $disabledCount = 0
    foreach ($rule in $rules) {
        if ($rule.Enabled) {
            # $rule.Enabled = $false
            $disabledCount++
            Write-Host "[Dry Run COM] Would disable: $($rule.Name)" -ForegroundColor Yellow
        }
    }
    if ($disabledCount -gt 0) {
        # $rules.Save()
        Write-Host "Successfully disabled $disabledCount rules via COM (Dry Run)." -ForegroundColor Green
    } else {
        Write-Host "All rules are already disabled (COM)." -ForegroundColor Green
    }
} else {
    Write-Host "Falling back to Exchange Online..." -ForegroundColor Cyan
    $userEmail = (whoami /upn 2>$null).Trim()
    if (-not $userEmail -or $userEmail -notmatch '@') {
        $userEmail = Read-Host "Enter your Exchange email address"
    }

    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Connect-ExchangeOnline -UserPrincipalName $userEmail -ShowBanner:$false
        $exoConnected = $true

        $exoRules = Get-InboxRule -Mailbox $userEmail -ErrorAction Stop
        $disabledCount = 0

        foreach ($rule in $exoRules) {
            if ($rule.Enabled) {
                # Disable-InboxRule -Identity $rule.Identity -Confirm:$false
                $disabledCount++
                Write-Host "[Dry Run EXO] Would disable: $($rule.Name)" -ForegroundColor Yellow
            }
        }
        if ($disabledCount -gt 0) {
            Write-Host "Successfully disabled $disabledCount rules via Exchange (Dry Run)." -ForegroundColor Green
        } else {
            Write-Host "All rules are already disabled (EXO)." -ForegroundColor Green
        }
    } catch {
        Write-Host "ERROR: Could not connect to Exchange Online." -ForegroundColor Red
    }
}

# Cleanup
if ($exoConnected) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue }
if ($null -ne $rules) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($rules) | Out-Null }
if ($null -ne $store) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($store) | Out-Null }
if ($null -ne $outlook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null }
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
