# ==============================================================================
# Unit & Regression Tests for Excel SQL Connect Pro (Safety Engine)
# ==============================================================================

$repoRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$engineScript = Join-Path $repoRoot "engine\excel_engine.ps1"
if (Test-Path $engineScript) {
    . $engineScript
}

$workDir = Join-Path $env:TEMP "excel_updater_test_work"
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
Write-Host "`n-- [TEST 6] test_connection_name_immutable & semantic snapshots --" -ForegroundColor Magenta
if (Get-Command "Compare-WorkbookSnapshot" -ErrorAction SilentlyContinue) {
    $snapA = @{
        FileFormat = 52
        WorksheetsCount = 2
        WorksheetNames = @("Sheet1", "Sheet2")
        QueriesCount = 1
        QueryNames = @("Query1")
        ConnectionsCount = 1
        ConnectionNames = @("SQL_Conn_192.168.1.1")
        HasVba = $true
        VbaAccessible = $true
        VbaModulesCount = 1
    }
    # Clone snapshot
    $snapMatching = @{
        FileFormat = 52
        WorksheetsCount = 2
        WorksheetNames = @("Sheet1", "Sheet2")
        QueriesCount = 1
        QueryNames = @("Query1")
        ConnectionsCount = 1
        ConnectionNames = @("SQL_Conn_192.168.1.1")
        HasVba = $true
        VbaAccessible = $true
        VbaModulesCount = 1
    }
    $cmpIdentical = Compare-WorkbookSnapshot $snapA $snapMatching
    Assert-True $cmpIdentical.IsValid "Identical snapshot comparison is valid"

    # Test mutated connection name (should FAIL with VALIDATION_FAILED_CONNECTION_IDENTITY)
    $snapMutatedConn = @{
        FileFormat = 52
        WorksheetsCount = 2
        WorksheetNames = @("Sheet1", "Sheet2")
        QueriesCount = 1
        QueryNames = @("Query1")
        ConnectionsCount = 1
        ConnectionNames = @("SQL_Conn_10.0.0.1") # Mutated!
        HasVba = $true
        VbaAccessible = $true
        VbaModulesCount = 1
    }
    $cmpMutatedConn = Compare-WorkbookSnapshot $snapA $snapMutatedConn
    Assert-False $cmpMutatedConn.IsValid "Mutated connection name must fail validation"
    Assert-Equal $cmpMutatedConn.ErrorCode "VALIDATION_FAILED_CONNECTION_IDENTITY" "test_connection_name_immutable error code"

    # Test mutated file format (52 -> 51)
    $snapMutatedFormat = @{
        FileFormat = 51 # Stripped macros!
        WorksheetsCount = 2
        WorksheetNames = @("Sheet1", "Sheet2")
        QueriesCount = 1
        QueryNames = @("Query1")
        ConnectionsCount = 1
        ConnectionNames = @("SQL_Conn_192.168.1.1")
        HasVba = $true
        VbaAccessible = $true
        VbaModulesCount = 1
    }
    $cmpMutatedFormat = Compare-WorkbookSnapshot $snapA $snapMutatedFormat
    Assert-False $cmpMutatedFormat.IsValid "Mutated FileFormat must fail validation"
    Assert-Equal $cmpMutatedFormat.ErrorCode "VALIDATION_FAILED_FILE_FORMAT" "Format change error code"
} else {
    Write-Host "  [SKIP] Compare-WorkbookSnapshot not yet implemented" -ForegroundColor Yellow
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
    $lockFile = Join-Path $workDir ("locked_test_" + [Guid]::NewGuid().ToString("N") + ".txt")
    [System.IO.File]::WriteAllText($lockFile, "Lock test content", [System.Text.Encoding]::UTF8)
    
    $stream = $null
    try {
        $stream = [System.IO.File]::Open($lockFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $isWritable = Test-FileWritable $lockFile
        Assert-False $isWritable "test_locked_file correctly detected as non-writable"
    } finally {
        if ($stream) {
            $stream.Close()
            $stream.Dispose()
        }
        if (Test-Path $lockFile) { Remove-Item $lockFile -Force }
    }
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

# ------------------------------------------------------------------------------
# 14. test_commit_staging_success
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 14] test_commit_staging_success --" -ForegroundColor Magenta
$targetFile = Join-Path $workDir "target_commit_success.txt"
"Original data before commit" | Set-Content $targetFile -Encoding UTF8
$initHash = Get-FileSha256 $targetFile

$stageFile = Join-Path $workDir "stage_commit_success.txt"
"New data committed safely" | Set-Content $stageFile -Encoding UTF8
$expectedStageHash = Get-FileSha256 $stageFile

$cRes = Commit-StagingWorkbook -StagingPath $stageFile -OriginalPath $targetFile -InitialHash $initHash
Assert-True $cRes.Success "Commit-StagingWorkbook should succeed"
Assert-Equal $cRes.FinalHash $expectedStageHash "Committed file final hash must match staging hash"
Assert-Equal (Get-Content $targetFile -Raw).Trim() "New data committed safely" "Target content updated"
Assert-False (Test-Path $stageFile) "Staging file must be cleaned/consumed after commit"

# ------------------------------------------------------------------------------
# 15. test_batch_rollback
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 15] test_batch_rollback --" -ForegroundColor Magenta
$batchTestDir = Join-Path $workDir "batch_test_dir"
if (-not (Test-Path $batchTestDir)) { New-Item -ItemType Directory -Path $batchTestDir -Force | Out-Null }
$f1 = Join-Path $batchTestDir "f1.xlsx"
$f2 = Join-Path $batchTestDir "f2.xlsx"
"Dummy content 1" | Set-Content $f1 -Encoding UTF8
"Dummy content 2" | Set-Content $f2 -Encoding UTF8
$f1Hash = Get-FileSha256 $f1
$f2Hash = Get-FileSha256 $f2

$bRes = Create-ExcelBackup -DirectoryPath $batchTestDir
Assert-True $bRes.success "Batch backup created successfully"

# Simulate corruption / unwanted modification in batch
"Corrupted content 1" | Set-Content $f1 -Encoding UTF8
Remove-Item $f2 -Force

$rRes = Restore-BatchBackup -BackupDir $bRes.backupDirectory -TargetDir $batchTestDir
Assert-True $rRes.Success "Restore-BatchBackup should succeed"
Assert-Equal (Get-FileSha256 $f1) $f1Hash "f1 restored to exact initial hash"
Assert-True (Test-Path $f2) "f2 restored from backup"
Assert-Equal (Get-FileSha256 $f2) $f2Hash "f2 restored to exact initial hash"

# ------------------------------------------------------------------------------
# 16. test_update_excel_directory_synthetic
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 16] test_update_excel_directory_synthetic --" -ForegroundColor Magenta
$integDir = Join-Path $workDir "integ_test_dir"
if (Test-Path $integDir) { Remove-Item $integDir -Recurse -Force -ErrorAction SilentlyContinue }
New-Item -ItemType Directory -Path $integDir -Force | Out-Null

$testNormDst = Join-Path $integDir "test_norm.xlsx"
$excelGen = New-Object -ComObject Excel.Application
$excelGen.Visible = $false
$excelGen.DisplayAlerts = $false
$wbGen = $excelGen.Workbooks.Add()
$wbGen.Queries.Add("TestQuery", 'let Source = Sql.Database("192.168.1.50", "TestDB") in Source', "Test SQL Query") | Out-Null
$wbGen.SaveAs($testNormDst, 51)
$wbGen.Close($false)
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wbGen) | Out-Null
$excelGen.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelGen) | Out-Null

$integInitHash = Get-FileSha256 $testNormDst

$rules = @(
    @{ oldText = "192.168.1.50"; newText = "10.0.0.1" }
)
$options = @{
    autoBackup = $true
    atomicBatch = $true
    updateConnections = $true
    updateQueries = $true
    updateVba = $false
}

$upRes = Update-ExcelDirectory -DirectoryPath $integDir -Rules $rules -Options $options
Assert-True $upRes.success "Update-ExcelDirectory returns success"
Assert-Equal $upRes.batchStatus "COMMITTED" "Batch status is COMMITTED"
Assert-Equal $upRes.updatedFilesCount 1 "Exactly 1 file updated"
Assert-True ($upRes.totalReplacements -gt 0) "Replacements made > 0"

$integFinalHash = Get-FileSha256 $testNormDst
Assert-False ($integInitHash -eq $integFinalHash) "File hash changed after update"
Assert-True (Test-Path $upRes.backupDirectory) "Backup directory created"

# ------------------------------------------------------------------------------
# 17. test_atomic_batch_preflight_abort
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 17] test_atomic_batch_preflight_abort --" -ForegroundColor Magenta
$batchFailDir = Join-Path $workDir "batch_fail_test_dir"
if (Test-Path $batchFailDir) { Remove-Item $batchFailDir -Recurse -Force -ErrorAction SilentlyContinue }
New-Item -ItemType Directory -Path $batchFailDir -Force | Out-Null

$fGood = Join-Path $batchFailDir "1_good.xlsx"
$excelGen = New-Object -ComObject Excel.Application
$excelGen.Visible = $false
$excelGen.DisplayAlerts = $false
$wbGen = $excelGen.Workbooks.Add()
$wbGen.Queries.Add("GoodQuery", 'let Source = "192.168.1.50" in Source', "Test SQL Query") | Out-Null
$wbGen.SaveAs($fGood, 51)
$wbGen.Close($false)
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wbGen) | Out-Null
$excelGen.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelGen) | Out-Null

$fGoodInitHash = Get-FileSha256 $fGood

# Second file is locked so Test-FileWritable fails
$fBad = Join-Path $batchFailDir "2_bad.xlsx"
"Dummy bad content" | Set-Content $fBad -Encoding UTF8

$lockStream = [System.IO.File]::Open($fBad, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)

try {
    $rules = @(
        @{ oldText = "192.168.1.50"; newText = "10.0.0.1" }
    )
    $options = @{
        autoBackup = $true
        atomicBatch = $true
        updateConnections = $true
        updateQueries = $true
        updateVba = $false
    }

    $failRes = Update-ExcelDirectory -DirectoryPath $batchFailDir -Rules $rules -Options $options
    Assert-False $failRes.success "Batch with locked file must fail pre-flight"
    Assert-Equal (Get-FileSha256 $fGood) $fGoodInitHash "First file must remain untouched after pre-flight abort"
} finally {
    if ($lockStream) {
        $lockStream.Close()
        $lockStream.Dispose()
    }
}

# ------------------------------------------------------------------------------
# 18. test_db_name_extraction
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 18] test_db_name_extraction --" -ForegroundColor Magenta
if (Get-Command "Get-DatabaseNamesFromText" -ErrorAction SilentlyContinue) {
    $sOle = "Provider=SQLOLEDB;Data Source=10.0.0.1;Initial Catalog=LOGO_2023;"
    $dbsOle = Get-DatabaseNamesFromText $sOle
    Assert-True ($dbsOle -contains "LOGO_2023") "test_db_name_extraction OLEDB Initial Catalog"

    $sOdbc = "DRIVER={SQL Server};SERVER=10.0.0.1;DATABASE=MIKRO_DB;"
    $dbsOdbc = Get-DatabaseNamesFromText $sOdbc
    Assert-True ($dbsOdbc -contains "MIKRO_DB") "test_db_name_extraction ODBC DATABASE"

    $sPq = 'let Source = Sql.Database("10.0.0.5", "ERP_PROD") in Source'
    $dbsPq = Get-DatabaseNamesFromText $sPq
    Assert-True ($dbsPq -contains "ERP_PROD") "test_db_name_extraction Power Query M"

    $sSql = 'USE [FINANS_DB]; SELECT * FROM [STOK_DB].[dbo].[ITEMS]'
    $dbsSql = Get-DatabaseNamesFromText $sSql
    Assert-True ($dbsSql -contains "FINANS_DB") "test_db_name_extraction SQL USE"
    Assert-True ($dbsSql -contains "STOK_DB") "test_db_name_extraction SQL three-part naming"
} else {
    Write-Host "  [SKIP] Get-DatabaseNamesFromText not yet implemented" -ForegroundColor Yellow
}

# ------------------------------------------------------------------------------
# 19. test_db_name_safe_replacement
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 19] test_db_name_safe_replacement --" -ForegroundColor Magenta
if (Get-Command "Invoke-SafeReplacement" -ErrorAction SilentlyContinue) {
    $dbRules = @(
        @{ oldText = "LOGO_2023"; newText = "LOGO_2024" }
    )

    $inConn = "Provider=SQLOLEDB;Data Source=10.0.0.1;Initial Catalog=LOGO_2023;"
    $resConn = Invoke-SafeReplacement -InputText $inConn -Rules $dbRules
    Assert-Equal $resConn.ResultText "Provider=SQLOLEDB;Data Source=10.0.0.1;Initial Catalog=LOGO_2024;" "test_db_name_safe_replacement in ConnectionString"

    $inPq = 'Sql.Database("10.0.0.1", "LOGO_2023")'
    $resPq = Invoke-SafeReplacement -InputText $inPq -Rules $dbRules
    Assert-Equal $resPq.ResultText 'Sql.Database("10.0.0.1", "LOGO_2024")' "test_db_name_safe_replacement in Power Query"

    $inSql = "SELECT * FROM [LOGO_2023].[dbo].[FATURA] WHERE TARIH > '2023-01-01'"
    $resSql = Invoke-SafeReplacement -InputText $inSql -Rules $dbRules
    Assert-Equal $resSql.ResultText "SELECT * FROM [LOGO_2024].[dbo].[FATURA] WHERE TARIH > '2023-01-01'" "test_db_name_safe_replacement inside SQL square brackets"
} else {
    Write-Host "  [SKIP] Invoke-SafeReplacement not yet implemented" -ForegroundColor Yellow
}

# ------------------------------------------------------------------------------
# 20. test_db_and_ip_combined_replacement
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 20] test_db_and_ip_combined_replacement --" -ForegroundColor Magenta
if (Get-Command "Invoke-SafeReplacement" -ErrorAction SilentlyContinue) {
    $combinedRules = @(
        @{ oldText = "192.168.1.50"; newText = "10.0.0.1" },
        @{ oldText = "LOGO_2023"; newText = "LOGO_2024" }
    )

    $inCombined = 'Data Source=192.168.1.50;Initial Catalog=LOGO_2023;Provider=SQLOLEDB;'
    $resCombined = Invoke-SafeReplacement -InputText $inCombined -Rules $combinedRules
    Assert-Equal $resCombined.ResultText 'Data Source=10.0.0.1;Initial Catalog=LOGO_2024;Provider=SQLOLEDB;' "Combined IP and DB replacement in single pass"
    Assert-Equal $resCombined.ReplacementsCount 2 "Combined replacements count should be 2"
} else {
    Write-Host "  [SKIP] Invoke-SafeReplacement not yet implemented" -ForegroundColor Yellow
}



