#Requires -Version 5.1
$ErrorActionPreference = "Stop"

$comAvailable = $false
$exoConnected = $false
$exportPath = Join-Path -Path $PWD -ChildPath "OutlookFolders_Export.csv"
$global:exportData = @()

Write-Host "Fetching Folder structure to export to: $exportPath" -ForegroundColor Cyan

function Gather-FolderNodes {
    param([System.__ComObject]$Folder, [string]$ParentPath = "")
    $currentPath = if ($ParentPath -eq "") { $Folder.Name } else { "$ParentPath\$($Folder.Name)" }
    $global:exportData += [PSCustomObject]@{ FolderPath = $currentPath; FolderName = $Folder.Name; ItemCount = $Folder.Items.Count; Source="COM" }
    foreach ($subFolder in $Folder.Folders) { Gather-FolderNodes -Folder $subFolder -ParentPath $currentPath }
}

try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $rootFolder = $namespace.DefaultStore.GetRootFolder()
    $testFolder = $rootFolder.Folders
    $comAvailable = $true
} catch {
    Write-Host "Outlook COM unavailable." -ForegroundColor Yellow
}

if ($comAvailable) {
    foreach ($folder in $rootFolder.Folders) { Gather-FolderNodes -Folder $folder }
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
            $global:exportData += [PSCustomObject]@{
                FolderPath = ($folder.FolderPath -replace "/", "\")
                FolderName = $folder.Name
                ItemCount = $folder.ItemsInFolder
                Source = "EXO"
            }
        }
    } catch { Write-Host "ERROR: Could not connect to Exchange Online." -ForegroundColor Red }
}

if ($global:exportData.Count -gt 0) {
    $global:exportData | Export-Csv -Path $exportPath -NoTypeInformation
    Write-Host "Successfully exported $($global:exportData.Count) folders." -ForegroundColor Green
} else { Write-Host "No folders found to export." -ForegroundColor Yellow }

if ($exoConnected) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue }
if ($null -ne $inbox) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($inbox) | Out-Null }
if ($null -ne $namespace) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) | Out-Null }
if ($null -ne $outlook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null }
[System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
