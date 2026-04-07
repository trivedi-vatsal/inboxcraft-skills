#Requires -Version 5.1
$ErrorActionPreference = "Continue"

$testsDir = $PSScriptRoot

Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "   InboxCraft Skills Test Suite Runner   " -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "This will run all READ-ONLY tests sequentially to verify COM object functionality.`n"

$readonlyTests = @(
    "test_show_rules.ps1",
    "test_show_folders.ps1",
    "test_find_large_folders.ps1",
    "test_export_categories.ps1"
)

foreach ($test in $readonlyTests) {
    Write-Host "`n>>> Running: $test" -ForegroundColor Magenta
    $testPath = Join-Path $testsDir $test
    
    if (Test-Path $testPath) {
        try {
            & $testPath
            Write-Host ">>> [PASS] $test executed successfully." -ForegroundColor Green
        } catch {
            Write-Host ">>> [FAIL] $test encountered an error: $_" -ForegroundColor Red
        }
    } else {
        Write-Host ">>> [WARN] $test not found." -ForegroundColor Yellow
    }
    
    Start-Sleep -Seconds 2
}

Write-Host "`n=========================================" -ForegroundColor Cyan
Write-Host "The following scripts require user parameters or are DESTRUCTIVE."
Write-Host "Please run them manually to test them:" -ForegroundColor Yellow
Write-Host " - test_export_rules.ps1"
Write-Host " - test_export_folders.ps1"
Write-Host " - test_clean_empty_folders.ps1 (Has Dry Run mode on by default)"
Write-Host " - test_disable_all_rules.ps1 (Has Dry Run mode on by default)"
Write-Host "=========================================" -ForegroundColor Cyan
