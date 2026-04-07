---
name: inboxcraft-outlook-show-rules
description: Fetches and displays all existing Inbox rules and their status in the console.
version: 1.0.0
---

# InboxCraft Outlook Show Rules

This skill instructs the agent on how to write a PowerShell script that inspects a user's Outlook profile to list rules. It uses the native COM object first, and safely falls back to Exchange Online if COM is unresponsive.

## When to Use

Use this skill whenever the user asks to "show my outlook rules", "list my rules", or asks to "see what rules I have in my inbox".

## Steps

1. **Inform the User:** Start by telling the user that you are generating a script to fetch their rules.
2. **Generate Script:** Generate exactly the following PowerShell script.

### PowerShell Script Template

```powershell
#Requires -Version 5.1
$ErrorActionPreference = "Stop"

$comAvailable = $false
$exoConnected = $false

Write-Host "Trying to fetch rules from Outlook COM..." -ForegroundColor Cyan
try {
    $outlook = New-Object -ComObject Outlook.Application
    $rules = $outlook.Session.DefaultStore.GetRules()
    $comAvailable = $true
} catch {
    Write-Host "Outlook COM rules unavailable." -ForegroundColor Yellow
}

if ($comAvailable) {
    if ($rules.Count -eq 0) {
        Write-Host "No rules found in the default mail store." -ForegroundColor Yellow
    } else {
        Write-Host "Found $($rules.Count) Rules (via COM):`n" -ForegroundColor Green
        foreach ($rule in $rules) {
            $enabled = if ($rule.Enabled) { "[ENABLED]" } else { "[DISABLED]" }
            Write-Host "$enabled $($rule.ExecutionOrder) - $($rule.Name)"
        }
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
        if ($exoRules.Count -eq 0) {
            Write-Host "No rules found in Exchange." -ForegroundColor Yellow
        } else {
            Write-Host "Found $($exoRules.Count) Rules (via Exchange):`n" -ForegroundColor Green
            foreach ($rule in $exoRules) {
                $enabled = if ($rule.Enabled) { "[ENABLED]" } else { "[DISABLED]" }
                Write-Host "$enabled $($rule.Priority) - $($rule.Name)"
            }
        }
    } catch {
        Write-Host "ERROR: Could not connect to Exchange Online." -ForegroundColor Red
    }
}

if ($exoConnected) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue }
if ($null -ne $rules) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($rules) | Out-Null }
if ($null -ne $outlook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null }
[System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
```

3. **Output:** Deliver the fully synthesized script to the user and prompt them to run it.
