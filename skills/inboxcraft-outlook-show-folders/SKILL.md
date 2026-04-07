---
name: inboxcraft-outlook-show-folders
description: Generate a PowerShell script to list the Outlook inbox folder tree structure.
---

# InboxCraft Outlook Show Folders

This skill instructs the agent on how to write a PowerShell script that recursively inspects a user's Outlook profile to print a formatted tree of all their inbox folders safely falling back to Exchange Online if COM relies fails.

## When to Use

Use this skill whenever the user asks to "show my outlook folders", "list my directories", or "print my folder structure".

## Steps

1. **Inform the User:** Start by telling the user that you are generating a script to fetch their folder tree.
2. **Generate Script:** Generate exactly the following PowerShell script.

### PowerShell Script Template

```powershell
#Requires -Version 5.1
$ErrorActionPreference = "Stop"

$comAvailable = $false
$exoConnected = $false

Write-Host "Fetching Inbox folder structure from Outlook..." -ForegroundColor Cyan

function Print-FolderTree {
    param([System.__ComObject]$Folder, [int]$IndentLevel = 0)
    $indentStr = "  " * $IndentLevel + "  |-- "
    $count = $Folder.Items.Count
    Write-Host "$indentStr$($Folder.Name) ($count items)"
    foreach ($subFolder in $Folder.Folders) {
        Print-FolderTree -Folder $subFolder -IndentLevel ($IndentLevel + 1)
    }
}

try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    $testFolder = $inbox.Folders
    $comAvailable = $true
} catch {
    Write-Host "Outlook COM unavailable." -ForegroundColor Yellow
}

if ($comAvailable) {
    Write-Host "`nInbox Root (via COM):" -ForegroundColor Green
    Print-FolderTree -Folder $inbox -IndentLevel 0
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

        Write-Host "`nInbox Root (via Exchange):" -ForegroundColor Green
        $folders = Get-MailboxFolderStatistics -Identity $userEmail -ErrorAction Stop | 
            Where-Object { $_.FolderPath -match "^/Inbox" } | 
            Sort-Object FolderPath

        foreach ($folder in $folders) {
            $depth = ($folder.FolderPath -split "/").Count - 2
            if ($depth -lt 0) { $depth = 0 }
            $indentStr = "  " * $depth + "  |-- "
            Write-Host "$indentStr$($folder.Name) ($($folder.ItemsInFolder) items)"
        }
    } catch {
        Write-Host "ERROR: Could not connect to Exchange Online." -ForegroundColor Red
    }
}

if ($exoConnected) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue }
if ($null -ne $inbox) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($inbox) | Out-Null }
if ($null -ne $namespace) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) | Out-Null }
if ($null -ne $outlook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null }
[System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
```

3. **Output:** Deliver the fully synthesized script to the user.
