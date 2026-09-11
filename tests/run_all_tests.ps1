# ==============================================================================
# Master Test Runner for Excel SQL Connect Pro
# ==============================================================================
param(
    [switch]$Detailed
)

$PSScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $PSScriptRoot

Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host " EXCEL SQL CONNECT PRO - REGRESSION & SAFETY TEST SUITE" -ForegroundColor Cyan
Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host "Repository Root: $repoRoot"

. (Join-Path $PSScriptRoot "test_harness.ps1")

# 1. Pre-Run Corporate Hash Check
Write-Host "`n[PRE-CHECK] Verifying corporate fixture baseline hashes..." -ForegroundColor Yellow
$preCheck = Assert-CorporateFilesUntouched $repoRoot
if (-not $preCheck) {
    Write-Host "[FATAL] Corporate files have already been modified prior to testing! Aborting." -ForegroundColor Red
    exit 1
}
Write-Host "[PRE-CHECK] Corporate fixture hashes match baseline." -ForegroundColor Green

# 2. Run Unit Tests
$unitTestPath = Join-Path $PSScriptRoot "unit\test_unit.ps1"
if (Test-Path $unitTestPath) {
    Write-Host "`n[RUNNING] Unit Tests ($unitTestPath)..." -ForegroundColor Yellow
    . $unitTestPath
}

$phase3UnitPath = Join-Path $PSScriptRoot "unit\test_phase3_suite.ps1"
if (Test-Path $phase3UnitPath) {
    Write-Host "`n[RUNNING] Phase 3 Operational Unit Tests ($phase3UnitPath)..." -ForegroundColor Yellow
    . $phase3UnitPath
}

# 3. Run Integration Tests
$intTestPath = Join-Path $PSScriptRoot "integration\test_integration.ps1"
if (Test-Path $intTestPath) {
    Write-Host "`n[RUNNING] Integration Tests ($intTestPath)..." -ForegroundColor Yellow
    . $intTestPath
}

# 4. Run Remediation & Blocker Tests
$remTestPath = Join-Path $PSScriptRoot "remediation\test_remediation_suite.ps1"
if (Test-Path $remTestPath) {
    Write-Host "`n[RUNNING] Remediation & Blocker Tests ($remTestPath)..." -ForegroundColor Yellow
    . $remTestPath
}

# 5. Post-Run Corporate Hash Check
Write-Host "`n[POST-CHECK] Verifying corporate fixture baseline hashes..." -ForegroundColor Yellow
$postCheck = Assert-CorporateFilesUntouched $repoRoot
if (-not $postCheck) {
    Write-Host "[FATAL] Corporate fixture files WERE MODIFIED during test execution!" -ForegroundColor Red
    exit 2
}
Write-Host "[POST-CHECK] Corporate fixture hashes remain 100% untouched." -ForegroundColor Green

# 5. Summary
Write-Host "`n==================================================================" -ForegroundColor Cyan
Write-Host " TEST RESULTS SUMMARY" -ForegroundColor Cyan
Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host "Total Passed: $($global:TestResults.Passed)" -ForegroundColor Green
Write-Host "Total Failed: $($global:TestResults.Failed)" -ForegroundColor $(if ($global:TestResults.Failed -gt 0) { "Red" } else { "Green" })

if ($global:TestResults.Failed -gt 0) {
    Write-Host "`nFAILED TESTS DETAILS:" -ForegroundColor Red
    $global:TestResults.Tests | Where-Object { $_.Status -eq "FAIL" } | ForEach-Object {
        Write-Host " - $($_.Name): $($_.Details)" -ForegroundColor Red
    }
    exit 1
}

Write-Host "`nALL TESTS PASSED SUCCESSFULLY!" -ForegroundColor Green
exit 0
