---
name: inboxcraft-outlook-disable-all-rules
description: Safely unchecks and pauses all active Outlook inbox rules without deleting them.
---

# InboxCraft Outlook Disable All Rules

This skill instructs the agent on how to write a PowerShell script that acts as a "Panic Button" for Outlook rules, disabling all of them immediately while preserving them so they can be re-enabled later. It falls back to Exchange Online if COM logic fails.

## When to Use

Use this skill whenever the user asks to "disable all my rules", "pause my inbox rules", or "stop email routing temporarily".

## Steps

1. **Inform the User:** Start by telling the user: "I will generate a script that connects to your Outlook or Exchange account, scans all your Inbox rules, and unchecks (disables) them."
2. **Generate Script:** Generate exactly the following PowerShell script.

### PowerShell Script Template

```powershell
#Requires -Version 5.1
$ErrorActionPreference = "Stop"

$comAvailable = $false
$exoConnected = $false

Write-Host "Connecting to Outlook to disable all rules..." -ForegroundColor Cyan

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
            $rule.Enabled = $false
            $disabledCount++
            Write-Host "Disabled rule: $($rule.Name)" -ForegroundColor Yellow
        }
    }
    if ($disabledCount -gt 0) {
        $rules.Save()
        Write-Host "Successfully disabled $disabledCount rules via COM." -ForegroundColor Green
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
                Disable-InboxRule -Identity $rule.Identity -Confirm:$false
                $disabledCount++
                Write-Host "Disabled rule: $($rule.Name)" -ForegroundColor Yellow
            }
        }
        if ($disabledCount -gt 0) {
            Write-Host "Successfully disabled $disabledCount rules via Exchange." -ForegroundColor Green
        } else {
            Write-Host "All rules are already disabled (EXO)." -ForegroundColor Green
        }
    } catch {
        Write-Host "ERROR: Could not connect to Exchange Online." -ForegroundColor Red
    }
}

if ($exoConnected) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue }
if ($null -ne $rules) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($rules) | Out-Null }
if ($null -ne $store) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($store) | Out-Null }
if ($null -ne $outlook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null }
[System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
```

3. **Output:** Deliver the fully synthesized script to the user.
