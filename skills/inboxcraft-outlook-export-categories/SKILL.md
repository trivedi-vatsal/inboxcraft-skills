---
name: inboxcraft-outlook-export-categories
description: Extracts all custom Outlook Categories and color labels for backup or inspection.
license: MIT
metadata:
  author: "trivedi-vatsal<trivedivatsal005@gmail.com>"
  version: "1.0.0"
---

# InboxCraft Outlook Export Categories

This skill instructs the agent on how to write a PowerShell script that fetches all custom categories and their colors.

## When to Use

Use this skill whenever the user asks to "export my outlook categories", "list my labels and colors", or "back up my tags".

## Steps

1. **Clarify Requirements:** Ask if they want to see the list or export it to a CSV.
2. **Generate Script:** Generate exactly the following PowerShell script. Do not write a fallback to Exchange Online, as EXO does natively support extracting graphical colors without Graph API. If COM fails, let it warn the user.

### PowerShell Script Template

```powershell
#Requires -Version 5.1
$ErrorActionPreference = "Stop"

$comAvailable = $false

Write-Host "Fetching Categories from Outlook COM..." -ForegroundColor Cyan

try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $categories = $namespace.Categories
    $comAvailable = $true
    
    $exportData = @()
    foreach ($cat in $categories) {
        $exportData += [PSCustomObject]@{
            Name = $cat.Name
            Color = $cat.Color
            ShortcutKey = $cat.ShortcutKey
        }
    }

    if ($exportData.Count -gt 0) {
        $exportData | Format-Table -AutoSize
        Write-Host "Found $($exportData.Count) custom categories." -ForegroundColor Green
    } else {
        Write-Host "No custom categories found." -ForegroundColor Yellow
    }

} catch {
    Write-Host "Outlook COM unavailable. Exception: $_" -ForegroundColor Yellow
}

if (-not $comAvailable) {
    Write-Host "`nWARNING: Cannot fallback to Exchange Online for Categories." -ForegroundColor Red
    Write-Host "Exchange Online PowerShell natively blocks access to Custom Category colors" -ForegroundColor Red
    Write-Host "without Graph API tokens. Please open Classic Windows Outlook to extract Categories." -ForegroundColor Red
}

if ($null -ne $categories) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($categories) | Out-Null }
if ($null -ne $namespace) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) | Out-Null }
if ($null -ne $outlook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null }
[System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
```

3. **Output:** Deliver the fully synthesized script to the user.
