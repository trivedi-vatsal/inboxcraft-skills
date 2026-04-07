---
name: inboxcraft-outlook-clean-empty-folders
description: Tracks down and safely deletes completely empty Outlook subfolders (includes Dry Run mode).
version: 1.0.0
---

# InboxCraft Outlook Clean Empty Folders

This skill instructs the agent on how to write a PowerShell script that inspects a user's Outlook profile to find all folders that contain zero items and optionally gives the user the ability to delete them. Falls back to Exchange Online if COM is unresponsive.

## When to Use

Use this skill whenever the user asks to "clean up my outlook", "delete empty folders", or "find folders with nothing in them".

## Steps

1. **Warn the User:** Strongly warn the user about destructive actions. Run in 'Dry Run Mode' first.
2. **Generate Script:** Generate exactly the following PowerShell script.

### PowerShell Script Template

```powershell
#Requires -Version 5.1
$ErrorActionPreference = "Stop"

$DryRun = $true # Agent: modify this to $false ONLY if the user explicitly requested deletion
$comAvailable = $false
$exoConnected = $false
$emptyFolders = @()

Write-Host "Scanning for empty folders. $(if($DryRun){'(DRY RUN MODE)'}else{'[DELETION MODE]'})" -ForegroundColor Cyan

function Find-EmptyFolders {
    param([System.__ComObject]$Folder)
    $ignoreList = @("Inbox", "Drafts", "Sent Items", "Deleted Items", "Outbox", "Junk Email", "Archive")
    
    for ($i = $Folder.Folders.Count; $i -ge 1; $i--) {
        $subFolder = $Folder.Folders.Item($i)
        Find-EmptyFolders -Folder $subFolder
    }
    
    if ($Folder.Items.Count -eq 0 -and $Folder.Folders.Count -eq 0 -and $Folder.Name -notin $ignoreList) {
        $emptyFolders += $Folder.Name
        if (-not $DryRun) {
            Write-Host "Deleting empty folder: $($Folder.Name)" -ForegroundColor Yellow
            $Folder.Delete()
        } else {
            Write-Host "[Dry Run] Found empty folder: $($Folder.Name)" -ForegroundColor Green
        }
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
    Find-EmptyFolders -Folder $inbox
    Write-Host "`nTotal empty folders found (via COM): $($emptyFolders.Count)" -ForegroundColor Cyan
} else {
    Write-Host "Falling back to Exchange Online..." -ForegroundColor Cyan
    $userEmail = (whoami /upn 2>$null).Trim()
    if (-not $userEmail -or $userEmail -notmatch '@') { $userEmail = Read-Host "Enter your Exchange email address" }

    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Connect-ExchangeOnline -UserPrincipalName $userEmail -ShowBanner:$false
        $exoConnected = $true

        $ignoreList = @("Inbox", "Drafts", "Sent Items", "Deleted Items", "Outbox", "Junk Email", "Archive")
        $folders = Get-MailboxFolderStatistics -Identity $userEmail -ErrorAction Stop | 
            Where-Object { $_.ItemsInFolder -eq 0 -and $_.FolderSize -match "0 B" -and $_.Name -notin $ignoreList }

        foreach ($folder in $folders) {
            $emptyFolders += $folder.Name
            if (-not $DryRun) {
                Write-Host "Deleting empty folder: $($folder.FolderPath)" -ForegroundColor Yellow
                Remove-MailboxFolder -Identity "$userEmail`:$($folder.FolderPath.Replace('\','/'))" -Confirm:$false
            } else {
                Write-Host "[Dry Run EXO] Found empty folder: $($folder.FolderPath)" -ForegroundColor Green
            }
        }
        Write-Host "`nTotal empty folders found (via EXO): $($emptyFolders.Count)" -ForegroundColor Cyan
    } catch { Write-Host "ERROR: Could not connect to Exchange Online." -ForegroundColor Red }
}

if ($DryRun -and $emptyFolders.Count -gt 0) { Write-Host "Set `$DryRun = `$false to actually delete them." }

if ($exoConnected) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue }
if ($null -ne $inbox) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($inbox) | Out-Null }
if ($null -ne $namespace) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) | Out-Null }
if ($null -ne $outlook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null }
[System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
```

3. **Output:** Deliver the script to the user.
