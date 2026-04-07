#Requires -Version 5.1
$ErrorActionPreference = "Stop"

$comAvailable = $false
$exoConnected = $false
$exportPath = Join-Path -Path $PWD -ChildPath "OutlookRules_Export.csv"
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
        $details = ""
        try { if ($rule.Conditions.From.Enabled) { $details += "[If From] " } } catch {}
        try { if ($rule.Conditions.Subject.Enabled) { $details += "[If Subject] " } } catch {}
        try { if ($rule.Conditions.SentTo.Enabled) { $details += "[If SentTo] " } } catch {}
        try { if ($rule.Actions.MoveToFolder.Enabled) { $details += "[Move To: $($rule.Actions.MoveToFolder.Folder.Name)] " } } catch {}
        try { if ($rule.Actions.Delete.Enabled -or $rule.Actions.DeletePermanently.Enabled) { $details += "[Delete] " } } catch {}
        if ([string]::IsNullOrWhiteSpace($details)) { $details = "Complex logic - View in Outlook/Exchange" }

        $exportData += [PSCustomObject]@{
            Name = $rule.Name
            ExecutionOrder = $rule.ExecutionOrder
            Enabled = $rule.Enabled
            IsLocalRule = $rule.IsLocalRule
            Details = $details.Trim()
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
                Details = ($rule.Description -replace "`n", " " -replace "`r", "").Trim()
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

# Cleanup
if ($exoConnected) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue }
if ($null -ne $rules) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($rules) | Out-Null }
if ($null -ne $outlook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null }
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
