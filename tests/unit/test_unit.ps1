# ==============================================================================
# Unit & Regression Tests for Excel SQL Connect Pro (Safety Engine)
# ==============================================================================

$repoRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$engineScript = Join-Path $repoRoot "engine\excel_engine.ps1"
if (Test-Path $engineScript) {
    . $engineScript
}

$workDir = Join-Path $repoRoot "tests\work"
if (-not (Test-Path $workDir)) { New-Item -ItemType Directory -Path $workDir -Force | Out-Null }
$fixtureDir = Join-Path $repoRoot "tests\fixtures"

# ------------------------------------------------------------------------------
# 1. test_ip_exact_match
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 1] test_ip_exact_match --" -ForegroundColor Magenta
if (Get-Command "Invoke-SafeReplacement" -ErrorAction SilentlyContinue) {
    $inputStr = 'Sql.Database("192.168.1.1", "MikroDB")'
    $rules = @(@{ oldText = "192.168.1.1"; newText = "10.0.0.1" })
    $res = Invoke-SafeReplacement -InputText $inputStr -Rules $rules
    Assert-Equal $res.ResultText 'Sql.Database("10.0.0.1", "MikroDB")' "test_ip_exact_match"
} else {
    Write-Host "  [SKIP] Invoke-SafeReplacement not yet implemented" -ForegroundColor Yellow
}

# ------------------------------------------------------------------------------
# 2. test_ip_substring_collision
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 2] test_ip_substring_collision --" -ForegroundColor Magenta
if (Get-Command "Invoke-SafeReplacement" -ErrorAction SilentlyContinue) {
    # 192.168.1.1 should NOT replace inside 192.168.1.15 or 192.168.1.100 or 1192.168.1.1
    $inputStr = 'Data Source=192.168.1.15;Alt=192.168.1.100;Third=1192.168.1.1;'
    $rules = @(@{ oldText = "192.168.1.1"; newText = "10.0.0.1" })
    $res = Invoke-SafeReplacement -InputText $inputStr -Rules $rules
    Assert-Equal $res.ResultText $inputStr "test_ip_substring_collision (should not modify substrings)"
    Assert-Equal $res.ReplacementsCount 0 "test_ip_substring_collision count should be 0"
} else {
    Write-Host "  [SKIP] Invoke-SafeReplacement not yet implemented" -ForegroundColor Yellow
}

# ------------------------------------------------------------------------------
# 3. test_cascading_rules
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 3] test_cascading_rules --" -ForegroundColor Magenta
if (Get-Command "Invoke-SafeReplacement" -ErrorAction SilentlyContinue) {
    # If rule 1 is A -> B and rule 2 is B -> C, A should become B, NOT C!
    $inputStr = 'Server=192.168.1.1;'
    $rules = @(
        @{ oldText = "192.168.1.1"; newText = "192.168.1.2" },
        @{ oldText = "192.168.1.2"; newText = "10.0.0.50" }
    )
    $res = Invoke-SafeReplacement -InputText $inputStr -Rules $rules
    Assert-Equal $res.ResultText 'Server=192.168.1.2;' "test_cascading_rules (single-pass prevents cascade to C)"
} else {
    Write-Host "  [SKIP] Invoke-SafeReplacement not yet implemented" -ForegroundColor Yellow
}

# ------------------------------------------------------------------------------
# 4. test_duplicate_rules
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 4] test_duplicate_rules --" -ForegroundColor Magenta
if (Get-Command "Test-ReplacementRulesValid" -ErrorAction SilentlyContinue) {
    $rules = @(
        @{ oldText = "192.168.1.1"; newText = "10.0.0.1" },
        @{ oldText = "192.168.1.1"; newText = "10.0.0.2" }
    )
    $check = Test-ReplacementRulesValid -Rules $rules
    Assert-False $check.IsValid "test_duplicate_rules should fail validation"
} else {
    Write-Host "  [SKIP] Test-ReplacementRulesValid not yet implemented" -ForegroundColor Yellow
}

# ------------------------------------------------------------------------------
# 5. test_no_match
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 5] test_no_match --" -ForegroundColor Magenta
if (Get-Command "Invoke-SafeReplacement" -ErrorAction SilentlyContinue) {
    $inputStr = 'Server=172.16.0.1;Database=DbTest;'
    $rules = @(@{ oldText = "192.168.1.1"; newText = "10.0.0.1" })
    $res = Invoke-SafeReplacement -InputText $inputStr -Rules $rules
    Assert-Equal $res.ResultText $inputStr "test_no_match text untouched"
    Assert-Equal $res.ReplacementsCount 0 "test_no_match count is 0"
} else {
    Write-Host "  [SKIP] Invoke-SafeReplacement not yet implemented" -ForegroundColor Yellow
}

# ------------------------------------------------------------------------------
# 6. test_connection_name_immutable
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 6] test_connection_name_immutable --" -ForegroundColor Magenta
if (Get-Command "Get-WorkbookSnapshot" -ErrorAction SilentlyContinue) {
    # Will test that connection name does not mutate in snapshot
    Assert-True $true "test_connection_name_immutable harness ready"
} else {
    Write-Host "  [SKIP] Get-WorkbookSnapshot not yet implemented" -ForegroundColor Yellow
}

# ------------------------------------------------------------------------------
# 7. test_backup_hash
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 7] test_backup_hash --" -ForegroundColor Magenta
if (Get-Command "New-VerifiedBackup" -ErrorAction SilentlyContinue -and (Get-Command "Get-FileSha256" -ErrorAction SilentlyContinue)) {
    $dummySrc = Join-Path $workDir "dummy_for_backup.txt"
    "Dummy content for backup test $(Get-Date)" | Set-Content $dummySrc -Encoding UTF8
    $bakDir = Join-Path $workDir "backup_test_dir"
    
    $bakRes = New-VerifiedBackup -SourcePath $dummySrc -BackupDir $bakDir
    Assert-True $bakRes.Success "test_backup_hash creation success"
    
    $origHash = Get-FileSha256 $dummySrc
    $bakHash = Get-FileSha256 $bakRes.BackupPath
    Assert-Equal $origHash $bakHash "test_backup_hash hashes match exactly"
} else {
    Write-Host "  [SKIP] New-VerifiedBackup or Get-FileSha256 not yet implemented" -ForegroundColor Yellow
}

# ------------------------------------------------------------------------------
# 8. test_staging_original_untouched
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 8] test_staging_original_untouched --" -ForegroundColor Magenta
if (Get-Command "New-StagingFile" -ErrorAction SilentlyContinue -and (Get-Command "Get-FileSha256" -ErrorAction SilentlyContinue)) {
    $dummyOrig = Join-Path $workDir "dummy_orig.txt"
    "Original content $(Get-Date)" | Set-Content $dummyOrig -Encoding UTF8
    $origHashBefore = Get-FileSha256 $dummyOrig
    
    $stageRes = New-StagingFile -OriginalFilePath $dummyOrig -OperationId "TESTOP1"
    Assert-True $stageRes.Success "test_staging creation success"
    Assert-True (Test-Path $stageRes.StagingPath) "staging file exists"
    
    # Modify staging
    "Modified in staging" | Set-Content $stageRes.StagingPath -Encoding UTF8
    $origHashAfter = Get-FileSha256 $dummyOrig
    Assert-Equal $origHashBefore $origHashAfter "test_staging_original_untouched (original remains unchanged)"
    
    if (Test-Path $stageRes.StagingPath) { Remove-Item $stageRes.StagingPath -Force }
} else {
    Write-Host "  [SKIP] New-StagingFile not yet implemented" -ForegroundColor Yellow
}

# ------------------------------------------------------------------------------
# 9. test_locked_file
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 9] test_locked_file --" -ForegroundColor Magenta
if (Get-Command "Test-FileWritable" -ErrorAction SilentlyContinue) {
    $lockFile = Join-Path $workDir "locked_test.txt"
    "Lock test content" | Set-Content $lockFile -Encoding UTF8
    
    $stream = [System.IO.File]::Open($lockFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    $isWritable = Test-FileWritable $lockFile
    $stream.Close()
    $stream.Dispose()
    
    Assert-False $isWritable "test_locked_file correctly detected as non-writable"
} else {
    Write-Host "  [SKIP] Test-FileWritable not yet implemented" -ForegroundColor Yellow
}

# ------------------------------------------------------------------------------
# 10. test_original_hash_changed
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 10] test_original_hash_changed --" -ForegroundColor Magenta
if (Get-Command "Commit-StagingWorkbook" -ErrorAction SilentlyContinue -and (Get-Command "Get-FileSha256" -ErrorAction SilentlyContinue)) {
    $dummyTarget = Join-Path $workDir "target_for_concurrent.txt"
    "Original text" | Set-Content $dummyTarget -Encoding UTF8
    $initialHash = Get-FileSha256 $dummyTarget
    
    $dummyStage = Join-Path $workDir "staging_for_concurrent.txt"
    "Staging text" | Set-Content $dummyStage -Encoding UTF8
    
    # Simulate someone else modifying original!
    "Modified by someone else!" | Set-Content $dummyTarget -Encoding UTF8
    
    $commitRes = Commit-StagingWorkbook -StagingPath $dummyStage -OriginalPath $dummyTarget -InitialHash $initialHash
    Assert-False $commitRes.Success "test_original_hash_changed commit should abort"
    Assert-Equal $commitRes.ErrorCode "CONCURRENT_MODIFICATION_DETECTED" "test_original_hash_changed error code"
} else {
    Write-Host "  [SKIP] Commit-StagingWorkbook not yet implemented" -ForegroundColor Yellow
}

# ------------------------------------------------------------------------------
# 11. test_xlsx_format, test_xlsm_format, test_xlsb_format
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 11-13] Format checks on synthetic fixtures --" -ForegroundColor Magenta
$normFixture = Join-Path $fixtureDir "test_norm.xlsx"
$macroFixture = Join-Path $fixtureDir "test_macro.xlsm"
$binFixture = Join-Path $fixtureDir "test_binary.xlsb"

Assert-True (Test-Path $normFixture) "test_xlsx_format fixture exists"
Assert-True (Test-Path $macroFixture) "test_xlsm_format fixture exists"
Assert-True (Test-Path $binFixture) "test_xlsb_format fixture exists"
