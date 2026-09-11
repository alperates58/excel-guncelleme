# ==============================================================================
# REMEDIATION TEST SUITE: P0/P1 Production Blocker Regressions
# Covers BUG-AUDIT-01 to BUG-AUDIT-06
# ==============================================================================

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
if (-not (Test-Path (Join-Path $repoRoot "engine\excel_engine.ps1"))) {
    $repoRoot = (Get-Location).Path
}

$engineScript = Join-Path $repoRoot "engine\excel_engine.ps1"
. $engineScript

$remWorkDir = Join-Path $env:TEMP "excel_remediation_test_work"
if (Test-Path $remWorkDir) { Remove-Item $remWorkDir -Recurse -Force -ErrorAction SilentlyContinue }
New-Item -ItemType Directory -Path $remWorkDir -Force | Out-Null

$remResults = @{
    Total = 0
    Passed = 0
    Failed = 0
    Tests = @()
}

function Assert-Remediation ($name, $passed, $details = "") {
    $remResults.Total++
    if ($passed) {
        $remResults.Passed++
        $remResults.Tests += @{ Name = $name; Status = "PASS"; Details = $details }
        Write-Host "  [PASS] $name" -ForegroundColor Green
    } else {
        $remResults.Failed++
        $remResults.Tests += @{ Name = $name; Status = "FAIL"; Details = $details }
        Write-Host "  [FAIL] $name : $details" -ForegroundColor Red
    }
}

Write-Host "=================================================================="
Write-Host " RUNNING REMEDIATION TEST SUITE (POST-FIX VERIFICATION)"
Write-Host "=================================================================="

# ------------------------------------------------------------------------------
# 1. BUG-AUDIT-01: test_com_no_orphan_process
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 1] test_com_no_orphan_process --" -ForegroundColor Magenta
$pidsBefore = @(Get-Process excel -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$iso = New-IsolatedExcelInstance
$detectedPid = $iso.Pid

$pidNonZero = ($detectedPid -ne $null -and $detectedPid -gt 0)
Assert-Remediation "test_com_pid_detected_nonzero" $pidNonZero "Detected PID must be > 0 (was: $detectedPid)"

Close-IsolatedExcelInstance $iso
Start-Sleep -Milliseconds 500
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

$pidsAfter = @(Get-Process excel -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$orphanPids = @($pidsAfter | Where-Object { $_ -notin $pidsBefore })
Assert-Remediation "test_com_no_orphan_process" ($orphanPids.Count -eq 0) "Expected 0 orphan processes, found: $($orphanPids -join ', ')"

# ------------------------------------------------------------------------------
# 2. BUG-AUDIT-01: test_user_excel_process_preserved
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 2] test_user_excel_process_preserved --" -ForegroundColor Magenta
$userExcel = New-Object -ComObject Excel.Application
$userExcel.Visible = $false
$userPids = @(Get-Process excel -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$simulatedUserPid = $userPids[-1]

$iso2 = New-IsolatedExcelInstance
Close-IsolatedExcelInstance $iso2

$userProcStillAlive = (Get-Process -Id $simulatedUserPid -ErrorAction SilentlyContinue) -ne $null
Assert-Remediation "test_user_excel_process_preserved" $userProcStillAlive "User Excel PID ($simulatedUserPid) must remain alive"

$userExcel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($userExcel) | Out-Null
[System.GC]::Collect()

# ------------------------------------------------------------------------------
# 3. BUG-AUDIT-03: test_rollback_partial_failure
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 3] test_rollback_partial_failure --" -ForegroundColor Magenta
$rollbackTestDir = Join-Path $remWorkDir "rollback_partial_fail_dir"
New-Item -ItemType Directory -Path $rollbackTestDir -Force | Out-Null

$f1 = Join-Path $rollbackTestDir "file1.xlsx"
$f2 = Join-Path $rollbackTestDir "file2.xlsx"
"content1" | Set-Content $f1 -Encoding UTF8
"content2" | Set-Content $f2 -Encoding UTF8

$bRes = Create-ExcelBackup -DirectoryPath $rollbackTestDir

# Corrupt backup of file1 so its restore fails checksum verification
$backedF1 = Join-Path $bRes.backupDirectory "file1.xlsx"
"corrupted in backup" | Set-Content $backedF1 -Encoding UTF8

$restoreRes = Restore-BatchBackup -BackupDir $bRes.backupDirectory -TargetDir $rollbackTestDir

$restoreFailedCorrectly = (-not $restoreRes.Success)
Assert-Remediation "test_restore_batch_backup_detects_corrupted_file" $restoreFailedCorrectly "Restore-BatchBackup must return Success=false"

$hasFailedFiles = ($restoreRes.failedRestoreFiles -ne $null -and $restoreRes.failedRestoreFiles.Count -gt 0)
Assert-Remediation "test_restore_batch_backup_lists_failed_files" $hasFailedFiles "Restore-BatchBackup must return failedRestoreFiles list"

if (Test-Path $bRes.backupDirectory) { Remove-Item $bRes.backupDirectory -Recurse -Force -ErrorAction SilentlyContinue }

# ------------------------------------------------------------------------------
# 4. BUG-AUDIT-02: test_double_fault_manual_recovery
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 4] test_double_fault_manual_recovery --" -ForegroundColor Magenta
$dfDir = Join-Path $remWorkDir "df_manual_recovery"
New-Item -ItemType Directory -Path $dfDir -Force | Out-Null
$dfTarget = Join-Path $dfDir "original_data.xlsx"
"Secret original data" | Set-Content $dfTarget -Encoding UTF8
$dfInitialHash = Get-FileSha256 $dfTarget

$dfStage = Join-Path $dfDir "stage_data.xlsx"
"New modified data" | Set-Content $dfStage -Encoding UTF8

$dfOldExpected = "$dfTarget.DFOP.rename_bak"
Move-Item $dfTarget $dfOldExpected -Force

$lockStream = [System.IO.File]::Open($dfTarget, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)

try {
    $cRes = Commit-StagingWorkbook -StagingPath $dfStage -OriginalPath $dfTarget -InitialHash $dfInitialHash -OperationId "DFOP"
    
    $isCriticalRecovery = ($cRes.ErrorCode -eq "CRITICAL_MANUAL_RECOVERY_REQUIRED")
    Assert-Remediation "test_double_fault_returns_critical_code" $isCriticalRecovery "ErrorCode must be CRITICAL_MANUAL_RECOVERY_REQUIRED (was: $($cRes.ErrorCode))"
    
    $hasRecoveryPath = (-not [string]::IsNullOrWhiteSpace($cRes.recoveryFilePath) -and (Test-Path $cRes.recoveryFilePath))
    Assert-Remediation "test_double_fault_exposes_recovery_file" $hasRecoveryPath "recoveryFilePath must be populated and exist"
} finally {
    $lockStream.Close()
    $lockStream.Dispose()
    if (Test-Path $dfTarget) { Remove-Item $dfTarget -Force -ErrorAction SilentlyContinue }
    if (Test-Path $dfOldExpected) { Remove-Item $dfOldExpected -Force -ErrorAction SilentlyContinue }
}

# ------------------------------------------------------------------------------
# 5. BUG-AUDIT-04: test_open_folder_command_injection
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 5] test_open_folder_command_injection --" -ForegroundColor Magenta
$serverScript = Join-Path $repoRoot "server.ps1"
$serverContent = Get-Content $serverScript -Raw

$hasCmdExe = ($serverContent -match 'cmd\.exe\s+/c')
Assert-Remediation "test_open_folder_no_cmd_exe" (-not $hasCmdExe) "server.ps1 must NOT use 'cmd.exe /c' in open-folder"

$hasPathResolve = ($serverContent -match 'Resolve-Path')
Assert-Remediation "test_open_folder_uses_resolve_path" $hasPathResolve "server.ps1 must use Resolve-Path to sanitize folder path"

# ------------------------------------------------------------------------------
# 6. BUG-AUDIT-05: test_static_file_directory_traversal & encoded traversal
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 6-7] test_static_file_directory_traversal --" -ForegroundColor Magenta
$hasCanonicalContainment = ($serverContent -match '\[System\.IO\.Path\]::GetFullPath' -and $serverContent -match 'StartsWith')
Assert-Remediation "test_static_file_canonical_containment" $hasCanonicalContainment "server.ps1 must enforce canonical root containment check"

$hasUrlDecode = ($serverContent -match 'UrlDecode')
Assert-Remediation "test_static_file_url_decode" $hasUrlDecode "server.ps1 must URL-decode path before resolving"

# ------------------------------------------------------------------------------
# 8. BUG-AUDIT-06: test_cors_not_wildcard
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 8] test_cors_not_wildcard --" -ForegroundColor Magenta
$hasWildcardCors = ($serverContent -match 'Access-Control-Allow-Origin",\s*"\*"')
Assert-Remediation "test_cors_not_wildcard" (-not $hasWildcardCors) "server.ps1 must NOT use wildcard Access-Control-Allow-Origin: *"

# ------------------------------------------------------------------------------
# 9. BUG-AUDIT-06: test_http_listener_loopback_only
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 9] test_http_listener_loopback_only --" -ForegroundColor Magenta
$hasWildcardBinding = ($serverContent -match 'http://\+:' -or $serverContent -match 'http://\*:')
Assert-Remediation "test_http_listener_loopback_only" (-not $hasWildcardBinding) "server.ps1 must only bind to loopback (no http://+: or http://*:)"

Write-Host "`n=================================================================="
Write-Host " REMEDIATION TEST SUMMARY: Passed=$($remResults.Passed), Failed=$($remResults.Failed) / Total=$($remResults.Total)"
Write-Host "=================================================================="

if (Test-Path $remWorkDir) { Remove-Item $remWorkDir -Recurse -Force -ErrorAction SilentlyContinue }

return $remResults
