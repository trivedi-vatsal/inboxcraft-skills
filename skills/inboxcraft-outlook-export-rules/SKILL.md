---
name: inboxcraft-outlook-export-rules
description: Gathers Outlook inbox rules and exports them to a JSON or CSV file for sharing or backup.
version: 1.0.0
---

# InboxCraft Outlook Export Rules

This skill instructs the agent on how to write a PowerShell script that exports a user's Outlook rules to a file, falling back to Exchange Online if COM is unavailable.

## When to Use

Use this skill whenever the user asks to "export my outlook rules", "back up my rules", or save them to a file.

## Steps

1. **Clarify Requirements:** Ask the user:
   - "Do you prefer the export in JSON or CSV format?"
   - "Where would you like to save the file? (If not specified, I will default to your Desktop)."

2. **Generate Script:** Generate the following PowerShell script. Modify the `$exportPath` based on their answer, and replace `Export-Csv` with `ConvertTo-Json | Out-File` if they chose JSON.

### PowerShell Script Template

```powershell
#Requires -Version 5.1
$ErrorActionPreference = "Stop"

$comAvailable = $false
$exoConnected = $false
$exportPath = Join-Path -Path [Environment]::GetFolderPath("Desktop") -ChildPath "OutlookRules_Export.csv"
$exportData = @()

Write-Host "Trying to fetch rules from Outlook COM to export to: $exportPath" -ForegroundColor Cyan
try {
    $outlook = New-Object -ComObject Outlook.Application
    $rules = $outlook.Session.DefaultStore.GetRules()
    $comAvailable = $true
} catch {
    Write-Host "Outlook COM rules unavailable." -ForegroundColor Yellow
}

if ($comAvailable) {
    foreach ($rule in $rules) {
        $exportData += [PSCustomObject]@{
            Name = $rule.Name
            ExecutionOrder = $rule.ExecutionOrder
            Enabled = $rule.Enabled
            IsLocalRule = $rule.IsLocalRule
            Source = "COM"
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
        foreach ($rule in $exoRules) {
            $exportData += [PSCustomObject]@{
                Name = $rule.Name
                ExecutionOrder = $rule.Priority
                Enabled = $rule.Enabled
                IsLocalRule = $false
                Source = "EXO"
            }
        }
    } catch {
        Write-Host "ERROR: Could not connect to Exchange Online." -ForegroundColor Red
    }
}

if ($exportData.Count -gt 0) {
    $exportData | Export-Csv -Path $exportPath -NoTypeInformation
    Write-Host "Successfully exported $($exportData.Count) rules." -ForegroundColor Green
} else {
    Write-Host "No rules found to export." -ForegroundColor Yellow
}

if ($exoConnected) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue }
if ($null -ne $rules) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($rules) | Out-Null }
if ($null -ne $outlook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null }
[System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
```

3. **Output:** Deliver the script to the user.
