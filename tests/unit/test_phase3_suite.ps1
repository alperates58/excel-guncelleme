# ==============================================================================
# Phase 3 Unit & Integration Regression Test Suite
# ==============================================================================

$repoRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path))
if (-not (Test-Path (Join-Path $repoRoot "engine\excel_engine.ps1"))) {
    $repoRoot = (Get-Location).Path
}

. (Join-Path $repoRoot "tests\test_harness.ps1")
$opMgrPath = Join-Path $repoRoot "engine\operation_manager.ps1"
if (Test-Path $opMgrPath) {
    . $opMgrPath
}
. (Join-Path $repoRoot "engine\excel_engine.ps1")

Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host " RUNNING PHASE 3 TEST SUITE" -ForegroundColor Cyan
Write-Host "==================================================================" -ForegroundColor Cyan

# ------------------------------------------------------------------------------
# 1. State Machine & Transition Validation Tests
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 1] State Machine Transition Rules --" -ForegroundColor Yellow

if (Get-Command "Test-ValidStateTransition" -ErrorAction SilentlyContinue) {
    Assert-True (Test-ValidStateTransition "QUEUED" "RUNNING") "Transition QUEUED -> RUNNING should be valid"
    Assert-True (Test-ValidStateTransition "RUNNING" "COMMITTING") "Transition RUNNING -> COMMITTING should be valid"
    Assert-True (Test-ValidStateTransition "COMMITTING" "COMPLETED") "Transition COMMITTING -> COMPLETED should be valid"
    Assert-True (Test-ValidStateTransition "RUNNING" "CANCELLATION_REQUESTED") "Transition RUNNING -> CANCELLATION_REQUESTED should be valid"
    Assert-True (Test-ValidStateTransition "CANCELLATION_REQUESTED" "CANCELLED") "Transition CANCELLATION_REQUESTED -> CANCELLED should be valid"
    Assert-True (Test-ValidStateTransition "CANCELLATION_REQUESTED" "ROLLING_BACK") "Transition CANCELLATION_REQUESTED -> ROLLING_BACK should be valid"
    Assert-True (Test-ValidStateTransition "ROLLING_BACK" "CANCELLED") "Transition ROLLING_BACK -> CANCELLED should be valid"
    Assert-True (Test-ValidStateTransition "ROLLING_BACK" "ROLLBACK_PARTIAL_FAILURE") "Transition ROLLING_BACK -> ROLLBACK_PARTIAL_FAILURE should be valid"
    Assert-True (Test-ValidStateTransition "RUNNING" "FAILED") "Transition RUNNING -> FAILED should be valid"
    Assert-True (Test-ValidStateTransition "RUNNING" "STALE_INTERRUPTED") "Transition RUNNING -> STALE_INTERRUPTED should be valid"

    # Invalid transitions
    Assert-False (Test-ValidStateTransition "COMPLETED" "RUNNING") "Transition COMPLETED -> RUNNING must be invalid"
    Assert-False (Test-ValidStateTransition "FAILED" "COMMITTING") "Transition FAILED -> COMMITTING must be invalid"
    Assert-False (Test-ValidStateTransition "CANCELLED" "RUNNING") "Transition CANCELLED -> RUNNING must be invalid"
    Assert-False (Test-ValidStateTransition "COMPLETED" "FAILED") "Transition COMPLETED -> FAILED must be invalid"
} else {
    Assert-True $false "Test-ValidStateTransition not implemented yet"
}

# ------------------------------------------------------------------------------
# 2. Secret Masking Fail-Closed Tests
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 2] Sensitive Data Masking --" -ForegroundColor Yellow

if (Get-Command "Protect-SensitiveData" -ErrorAction SilentlyContinue) {
    $rawConnStr = "Server=192.168.1.10;Database=Finance;User Id=sa;Password=UltraSecretPass123!;Trusted_Connection=False;"
    $maskedConnStr = Protect-SensitiveData $rawConnStr
    Assert-False ($maskedConnStr.Contains("UltraSecretPass123!")) "Raw password must not exist in masked string"
    Assert-True ($maskedConnStr.Contains("Password=***")) "Password must be masked with ***"
    Assert-True ($maskedConnStr.Contains("User Id=***")) "User Id must be masked"

    $rawPwd = "Provider=SQLOLEDB;Data Source=10.0.0.1;Pwd=HiddenPwd999;UID=admin;"
    $maskedPwd = Protect-SensitiveData $rawPwd
    Assert-False ($maskedPwd.Contains("HiddenPwd999")) "Pwd value must not exist in masked string"
    Assert-True ($maskedPwd.Contains("Pwd=***")) "Pwd must be masked with ***"

    # Hashtable masking
    $ht = @{
        Message = "Error connecting to db with Pwd=Secret999"
        ErrorDetails = @{
            Conn = "Server=srv;Password=TopSecret;"
        }
    }
    $maskedHt = Protect-SensitiveData $ht
    Assert-False ($maskedHt.Message.Contains("Secret999")) "Nested message password must be masked"
    Assert-False ($maskedHt.ErrorDetails.Conn.Contains("TopSecret")) "Nested hashtable password must be masked"
} else {
    Assert-True $false "Protect-SensitiveData not implemented yet"
}

# ------------------------------------------------------------------------------
# 3. Thread-Safe Registry & Snapshot Immutability Tests
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 3] Thread-Safe Operation Registry & Snapshot Immutability --" -ForegroundColor Yellow

if (Get-Command "Register-Operation" -ErrorAction SilentlyContinue) {
    Reset-OperationRegistry
    $regRes = Register-Operation -Type "UPDATE" -Directory "C:\Temp\Test" -Options @{ atomicBatch = $true } -Rules @(@{ oldText="A"; newText="B" })
    Assert-True $regRes.Success "Register-Operation should succeed"
    $testOpId = $regRes.OperationId
    Assert-True (-not [string]::IsNullOrWhiteSpace($testOpId)) "OperationId must be non-empty"

    $snap1 = Get-OperationSnapshot $testOpId
    Assert-Equal $snap1.status "QUEUED" "Initial status must be QUEUED"

    # Mutate through central mutator
    Set-OperationStatus $testOpId "RUNNING" "Preflight"
    $snap2 = Get-OperationSnapshot $testOpId
    Assert-Equal $snap2.status "RUNNING" "Status must transition to RUNNING"
    Assert-Equal $snap2.currentStage "Preflight" "CurrentStage must be Preflight"

    # Verify Snapshot Immutability: Mutating $snap2 should NOT mutate registry
    $snap2.status = "MUTATED_EXTERNALLY"
    $snap3 = Get-OperationSnapshot $testOpId
    Assert-Equal $snap3.status "RUNNING" "Registry must not be mutated by modifying returned snapshot"
    
    # Cleanup
    Reset-OperationRegistry
} else {
    Assert-True $false "Registry functions not implemented yet"
}

# ------------------------------------------------------------------------------
# 4. Single-Active COM Operation Mutex & 409 Conflict Tests
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 4] Concurrency Mutex & 409 Rejection --" -ForegroundColor Yellow

if (Get-Command "Register-Operation" -ErrorAction SilentlyContinue) {
    Reset-OperationRegistry
    $op1 = Register-Operation -Type "UPDATE" -Directory "C:\Temp\Test1"
    Set-OperationStatus $op1.OperationId "RUNNING" "Excel Update"

    # Attempt to register second heavy operation while op1 is RUNNING
    $op2 = Register-Operation -Type "UPDATE" -Directory "C:\Temp\Test2"
    Assert-False $op2.Success "Second update must be rejected while first is running"
    Assert-Equal $op2.ErrorCode "OPERATION_ALREADY_RUNNING" "Error code must be OPERATION_ALREADY_RUNNING"
    Assert-Equal $op2.ActiveOperationId $op1.OperationId "ActiveOperationId must reference op1"

    # Finish op1
    Set-OperationStatus $op1.OperationId "COMPLETED" "Done"
    
    # Now second operation should succeed
    $op3 = Register-Operation -Type "UPDATE" -Directory "C:\Temp\Test3"
    Assert-True $op3.Success "New update should succeed once previous is COMPLETED"

    # Cleanup
    Reset-OperationRegistry
} else {
    Assert-True $false "Concurrency mutex not implemented yet"
}

# ------------------------------------------------------------------------------
# 5. History Atomic Persistence & Corrupt Recovery Tests
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 5] History Atomic Persistence & Corrupt Recovery --" -ForegroundColor Yellow

$testHistoryDir = Join-Path $env:TEMP "phase3_test_history"
if (Test-Path $testHistoryDir) { Remove-Item $testHistoryDir -Recurse -Force }
New-Item -ItemType Directory -Path $testHistoryDir -Force | Out-Null

if (Get-Command "Save-OperationHistory" -ErrorAction SilentlyContinue) {
    $histFile = Join-Path $testHistoryDir "history.json"
    
    $mockOp = @{
        operationId = "op-test-1"
        type = "UPDATE"
        status = "COMPLETED"
        createdAt = (Get-Date).ToString("o")
        finishedAt = (Get-Date).ToString("o")
        directory = "C:\Test"
        totalFiles = 5
        updatedFiles = 3
        failedFiles = 0
        batchStatus = "COMMITTED"
        backupDirectory = "C:\Test\Backup"
    }

    $saveRes = Save-OperationHistory -OperationSummary $mockOp -HistoryFilePath $histFile
    Assert-True $saveRes "Save-OperationHistory should succeed"
    Assert-True (Test-Path $histFile) "history.json should exist"

    # Load history
    $histList = Get-OperationHistory -HistoryFilePath $histFile
    Assert-Equal $histList.Count 1 "History list should have 1 item"
    Assert-Equal $histList[0].operationId "op-test-1" "Operation ID must match"

    # Test Corrupt File Recovery
    [System.IO.File]::WriteAllText($histFile, "THIS IS CORRUPTED JSON {{{")
    $corruptRecoverList = Get-OperationHistory -HistoryFilePath $histFile
    Assert-True ($corruptRecoverList -ne $null) "Corrupt history should not throw and return safe list"
    Assert-True (Get-ChildItem $testHistoryDir -Filter "history.corrupt.*.json").Count -gt 0 "Corrupted history must be preserved with .corrupt extension"

    Remove-Item $testHistoryDir -Recurse -Force
} else {
    Assert-True $false "History persistence not implemented yet"
}

# ------------------------------------------------------------------------------
# 6. Structured Audit Log JSONL Tests
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 6] Structured JSONL Audit Logging --" -ForegroundColor Yellow

$testLogDir = Join-Path $env:TEMP "phase3_test_logs"
if (Test-Path $testLogDir) { Remove-Item $testLogDir -Recurse -Force }
New-Item -ItemType Directory -Path $testLogDir -Force | Out-Null

if (Get-Command "Write-AuditLogEvent" -ErrorAction SilentlyContinue) {
    $opLogId = "op-audit-test"
    Write-AuditLogEvent -LogDirectory $testLogDir -OperationId $opLogId -Level "INFO" -EventName "OPERATION_STARTED" -Message "Started operation" -Data @{ dir = "C:\Test" }
    Write-AuditLogEvent -LogDirectory $testLogDir -OperationId $opLogId -Level "WARNING" -EventName "FILE_WARNING" -Message "Connecting with Password=SuperSecret123" -Data @{ detail = "pwd" }

    $logFiles = Get-ChildItem -Path $testLogDir -Recurse -Filter "$opLogId.jsonl"
    Assert-Equal $logFiles.Count 1 "Audit log JSONL file must be created"
    
    if ($logFiles.Count -eq 1) {
        $lines = Get-Content $logFiles[0].FullName
        Assert-Equal $lines.Count 2 "There should be 2 JSONL lines"
        
        $line1 = $lines[0] | ConvertFrom-Json
        Assert-Equal $line1.event "OPERATION_STARTED" "Event name must match"
        
        $line2 = $lines[1] | ConvertFrom-Json
        Assert-Equal $line2.level "WARNING" "Level must match"
        Assert-False ($lines[1].Contains("SuperSecret123")) "Audit log must sanitize passwords"
        Assert-True ($lines[1].Contains("Password=***")) "Password must be replaced with ***"
    }

    Remove-Item $testLogDir -Recurse -Force
} else {
    Assert-True $false "Write-AuditLogEvent not implemented yet"
}

# ------------------------------------------------------------------------------
# 7. Real STA Runspace Verification
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 7] Real PowerShell STA Runspace Verification --" -ForegroundColor Yellow

if (Get-Command "New-StaRunspace" -ErrorAction SilentlyContinue) {
    $staContext = New-StaRunspace
    Assert-True ($staContext -ne $null) "STA Runspace must be created"
    
    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $staContext.Runspace
    $ps.AddScript({
        [System.Threading.Thread]::CurrentThread.GetApartmentState().ToString()
    }) | Out-Null
    
    $aptState = ($ps.Invoke())[0]
    $ps.Dispose()
    Close-StaRunspace $staContext
    
    Assert-Equal $aptState "STA" "Worker runspace apartment state must be strictly STA"
} else {
    Assert-True $false "New-StaRunspace not implemented yet"
}

# ------------------------------------------------------------------------------
# 8. Preview (Dry Run) Tests on Synthetic Fixtures
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 8] Preview (Dry Run) Leaves Files 100% Untouched --" -ForegroundColor Yellow

$testPreviewDir = Join-Path $env:TEMP "phase3_test_preview"
if (Test-Path $testPreviewDir) { Remove-Item $testPreviewDir -Recurse -Force }
New-Item -ItemType Directory -Path $testPreviewDir -Force | Out-Null

# Generate a synthetic workbook using isolated COM
$iso = New-IsolatedExcelInstance
$wbs = $iso.Excel.Workbooks
$wb = $wbs.Add()
$ws = $wb.ActiveSheet
$rA1 = $ws.Range("A1")
$rA1.Value2 = "Server=192.168.2.15;Database=Test;"
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($rA1) | Out-Null
$rA1 = $null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) | Out-Null
$ws = $null

$sampleXlsx = Join-Path $testPreviewDir "sample_preview.xlsx"
$wb.SaveAs($sampleXlsx, 51)
$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
$wb = $null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wbs) | Out-Null
$wbs = $null
Close-IsolatedExcelInstance $iso
$iso = $null

$initialSha = Get-FileSha256 $sampleXlsx

if (Get-Command "Preview-ExcelDirectory" -ErrorAction SilentlyContinue) {
    $rules = @(@{ oldText = "192.168.2.15"; newText = "10.0.0.1" })
    $prevRes = Preview-ExcelDirectory -DirectoryPath $testPreviewDir -Rules $rules
    
    Assert-True $prevRes.Success "Preview should succeed"
    Assert-Equal $prevRes.TotalFiles 1 "Preview should find 1 file"
    
    # Assert ZERO modifications to source file
    $postSha = Get-FileSha256 $sampleXlsx
    Assert-Equal $postSha $initialSha "Preview must leave source file hash 100% untouched"
    
    # Assert NO backup folder created
    $backups = Get-ChildItem $testPreviewDir -Directory -Filter "*Backup*"
    Assert-Equal $backups.Count 0 "Preview must NEVER create a backup directory"
    
    # Assert NO staging file left
    $stagings = Get-ChildItem $testPreviewDir -Filter "*.staging.*"
    Assert-Equal $stagings.Count 0 "Preview must NEVER leave staging files"
} else {
    Assert-True $false "Preview-ExcelDirectory not implemented yet"
}

Remove-Item $testPreviewDir -Recurse -Force

# ------------------------------------------------------------------------------
# 9. Diagnostics Endpoint Verification
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 9] Diagnostics System Health Check --" -ForegroundColor Yellow

if (Get-Command "Invoke-SystemDiagnostics" -ErrorAction SilentlyContinue) {
    $diag = Invoke-SystemDiagnostics -RepoRoot $repoRoot
    Assert-True $diag.Success "Diagnostics must succeed"
    Assert-True $diag.ExcelComAvailable "Excel COM must be available"
    Assert-True (-not [string]::IsNullOrWhiteSpace($diag.ExcelVersion)) "Excel Version must be detected"
    Assert-True $diag.WorkingDirWritable "Working directory must be writable"
    Assert-True $diag.BackupDirWritable "Backup directory must be writable"
    Assert-True ($diag.ActiveExcelPids.Count -ge 0) "Active Excel PIDs must be an array"
} else {
    Assert-True $false "Invoke-SystemDiagnostics not implemented yet"
}

# ------------------------------------------------------------------------------
# 10. Monotonic Progress Calculation Tests
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 10] Monotonic Progress Calculation --" -ForegroundColor Yellow

if (Get-Command "Calculate-WeightedProgress" -ErrorAction SilentlyContinue) {
    # 10 files total. File 1 at Preflight (stage 0) -> progress small
    $p1 = Calculate-WeightedProgress -TotalFiles 10 -CompletedFiles 0 -CurrentFileStage "Preflight"
    # File 1 at Excel Update (stage 3)
    $p2 = Calculate-WeightedProgress -TotalFiles 10 -CompletedFiles 0 -CurrentFileStage "Excel Update"
    # File 1 at Commit (stage 5)
    $p3 = Calculate-WeightedProgress -TotalFiles 10 -CompletedFiles 0 -CurrentFileStage "Commit"
    # 5 files completed, 6th at Excel Update
    $p4 = Calculate-WeightedProgress -TotalFiles 10 -CompletedFiles 5 -CurrentFileStage "Excel Update"
    # All 10 files completed
    $pFinal = Calculate-WeightedProgress -TotalFiles 10 -CompletedFiles 10 -CurrentFileStage "Cleanup"

    Assert-True ($p1 -le $p2) "Progress must be monotonic (p1 <= p2)"
    Assert-True ($p2 -le $p3) "Progress must be monotonic (p2 <= p3)"
    Assert-True ($p3 -le $p4) "Progress must be monotonic (p3 <= p4)"
    Assert-True ($pFinal -eq 100) "Completed progress must be exactly 100"
    Assert-True ($p4 -lt 100) "In-progress progress must be < 100"
} else {
    Assert-True $false "Calculate-WeightedProgress not implemented yet"
}

Write-Host "`n==================================================================" -ForegroundColor Cyan
Write-Host " PHASE 3 TEST SUITE COMPLETED" -ForegroundColor Cyan
Write-Host "==================================================================" -ForegroundColor Cyan
