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
