#Requires -Version 5.1
$ErrorActionPreference = "Stop"

$comAvailable = $false
$exoConnected = $false
$global:folderStats = @()

Write-Host "Scanning Outlook folders for item counts. This may take a minute..." -ForegroundColor Cyan

function Scan-Folders {
    param([System.__ComObject]$Folder, [string]$ParentPath = "")
    $currentPath = if ($ParentPath -eq "") { $Folder.Name } else { "$ParentPath\$($Folder.Name)" }
    if ($Folder.Items.Count -gt 0) {
        $global:folderStats += [PSCustomObject]@{ FolderPath = $currentPath; ItemCount = $Folder.Items.Count }
    }
    foreach ($subFolder in $Folder.Folders) { Scan-Folders -Folder $subFolder -ParentPath $currentPath }
}

try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $rootFolders = $namespace.DefaultStore.GetRootFolder().Folders
    $testFolder = $rootFolders.Count
    $comAvailable = $true
} catch {
    Write-Host "Outlook COM unavailable." -ForegroundColor Yellow
}

if ($comAvailable) {
    foreach ($root in $rootFolders) { Scan-Folders -Folder $root }
} else {
    Write-Host "Falling back to Exchange Online..." -ForegroundColor Cyan
    $userEmail = (whoami /upn 2>$null).Trim()
    if (-not $userEmail -or $userEmail -notmatch '@') { $userEmail = Read-Host "Enter your Exchange email address" }

    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Connect-ExchangeOnline -UserPrincipalName $userEmail -ShowBanner:$false
        $exoConnected = $true
        
        $folders = Get-MailboxFolderStatistics -Identity $userEmail -ErrorAction Stop
        foreach ($folder in $folders) {
            if ($folder.ItemsInFolder -gt 0) {
                $global:folderStats += [PSCustomObject]@{
                    FolderPath = ($folder.FolderPath -replace "/", "\").TrimStart('\')
                    ItemCount = $folder.ItemsInFolder
                }
            }
        }
    } catch { Write-Host "ERROR: Could not connect to Exchange Online." -ForegroundColor Red }
}

Write-Host "`nTop 10 Largest Folders by Item Count:" -ForegroundColor Green
$global:folderStats | Sort-Object ItemCount -Descending | Select-Object -First 10 | Format-Table -AutoSize

if ($exoConnected) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue }
if ($null -ne $namespace) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) | Out-Null }
if ($null -ne $outlook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null }
[System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
