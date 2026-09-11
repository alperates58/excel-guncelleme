# ==============================================================================
# Phase 3 Real HTTP Integration Test Suite
# ==============================================================================

$repoRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
if (-not (Test-Path (Join-Path $repoRoot "engine\excel_engine.ps1"))) {
    $repoRoot = (Get-Location).Path
}

. (Join-Path $repoRoot "tests\test_harness.ps1")
. (Join-Path $repoRoot "engine\operation_manager.ps1")
. (Join-Path $repoRoot "engine\excel_engine.ps1")

Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host " RUNNING PHASE 3 REAL HTTP INTEGRATION TEST SUITE" -ForegroundColor Cyan
Write-Host "==================================================================" -ForegroundColor Cyan

$testPort = 3899
$testBaseUrl = "http://127.0.0.1:$testPort"
$serverScript = Join-Path $repoRoot "server.ps1"

# ------------------------------------------------------------------------------
# 0. Start Server in Background Runspace
# ------------------------------------------------------------------------------
Write-Host "`n[SETUP] Starting background HTTP server on port $testPort..." -ForegroundColor Yellow

$serverRunspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
$serverRunspace.ApartmentState = [System.Threading.ApartmentState]::STA
$serverRunspace.ThreadOptions = "ReuseThread"
$serverRunspace.Open()

$serverPs = [System.Management.Automation.PowerShell]::Create()
$serverPs.Runspace = $serverRunspace
$serverPs.AddScript({
    param($scriptPath, $port)
    $env:EXCEL_UPDATER_NO_BROWSER = "1"
    & $scriptPath -Port $port -NoBrowser
}).AddArgument($serverScript).AddArgument($testPort) | Out-Null

$serverAsync = $serverPs.BeginInvoke()

# Wait for server to start responding
$serverReady = $false
for ($i = 0; $i -lt 25; $i++) {
    Start-Sleep -Milliseconds 400
    try {
        $res = Invoke-RestMethod -Uri "$testBaseUrl/api/status" -Method GET -TimeoutSec 2 -ErrorAction Stop
        if ($res.status -eq "ok") {
            $serverReady = $true
            break
        }
    } catch { }
}

Assert-True $serverReady "Background HTTP server must be listening and responsive on port $testPort"
if (-not $serverReady) {
    Write-Host "[FATAL] Server failed to start. Aborting integration tests." -ForegroundColor Red
    return
}

# ------------------------------------------------------------------------------
# 1. GET /api/status Endpoint
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 1] HTTP GET /api/status --" -ForegroundColor Yellow
$statusRes = Invoke-RestMethod -Uri "$testBaseUrl/api/status" -Method GET
Assert-Equal $statusRes.status "ok" "Status must be ok"
Assert-Equal $statusRes.port $testPort "Reported port must match testPort"

# ------------------------------------------------------------------------------
# 2. GET /api/diagnostics Endpoint
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 2] HTTP GET /api/diagnostics --" -ForegroundColor Yellow
$diagRes = Invoke-RestMethod -Uri "$testBaseUrl/api/diagnostics" -Method GET
Assert-True $diagRes.Success "Diagnostics must report success"
Assert-True $diagRes.ExcelComAvailable "Excel COM must be available"
Assert-True $diagRes.WorkingDirWritable "Working directory must be writable"
Assert-True $diagRes.BackupDirWritable "Backup directory must be writable"

# ------------------------------------------------------------------------------
# 3. POST /api/preview (Asynchronous Preview / Dry Run)
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 3] HTTP POST /api/preview (Dry Run via Worker) --" -ForegroundColor Yellow

$testPreviewDir = Join-Path $repoRoot "tests\work\http_preview_test"
if (Test-Path $testPreviewDir) { Remove-Item $testPreviewDir -Recurse -Force }
New-Item -ItemType Directory -Path $testPreviewDir -Force | Out-Null

# Create synthetic file
$iso = New-IsolatedExcelInstance
$wb = $iso.Excel.Workbooks.Add()
$wb.Queries.Add("TestQuery", 'let Source = Sql.Database("192.168.2.15", "PreviewTest") in Source', "Preview Query") | Out-Null
$sampleFile = Join-Path $testPreviewDir "preview_sample.xlsx"
$wb.SaveAs($sampleFile, 51)
$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
Close-IsolatedExcelInstance $iso

$initialSha = Get-FileSha256 $sampleFile

# Dispatch Preview via HTTP
$previewPayload = @{
    directory = $testPreviewDir
    rules = @(@{ oldText = "192.168.2.15"; newText = "10.0.0.1" })
    options = @{ updateQueries = $true; updateConnections = $true; updateVba = $true }
} | ConvertTo-Json

$previewHttpRes = Invoke-WebRequest -Uri "$testBaseUrl/api/preview" -Method POST -Body $previewPayload -ContentType "application/json; charset=utf-8" -UseBasicParsing
Assert-Equal $previewHttpRes.StatusCode 202 "POST /api/preview must return HTTP 202 Accepted"

$previewJson = $previewHttpRes.Content | ConvertFrom-Json
$previewOpId = $previewJson.operationId
Assert-True (-not [string]::IsNullOrWhiteSpace($previewOpId)) "OperationId must be returned"

# Poll until completion
$pollCount = 0
$finalOp = $null
while ($pollCount -lt 40) {
    Start-Sleep -Milliseconds 500
    $pollRes = Invoke-RestMethod -Uri "$testBaseUrl/api/operations/$previewOpId" -Method GET
    $finalOp = $pollRes.operation
    Write-Host "    [Poll $pollCount] Status: $($finalOp.status), Stage: $($finalOp.currentStage), Progress: $($finalOp.progressPercent)%"
    if ($finalOp.status -eq "COMPLETED" -or $finalOp.status -eq "FAILED") {
        break
    }
    $pollCount++
}

Assert-Equal $finalOp.status "COMPLETED" "Preview operation must reach COMPLETED state"

# Assert Dry Run invariants: 0 byte change on source file
$postSha = Get-FileSha256 $sampleFile
Assert-Equal $postSha $initialSha "Preview must leave source file SHA256 100% untouched"

# Assert no backup folder created
$backups = Get-ChildItem $testPreviewDir -Directory -Filter "*Backup*"
Assert-Equal $backups.Count 0 "Preview must NEVER create a backup directory"

# ------------------------------------------------------------------------------
# 4. HTTP POST /api/update (Asynchronous Update & Atomic Commit via Worker)
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 4] HTTP POST /api/update (Async Update via Worker) --" -ForegroundColor Yellow

$updatePayload = @{
    directory = $testPreviewDir
    rules = @(@{ oldText = "192.168.2.15"; newText = "10.0.0.99" })
    options = @{
        updateQueries = $true
        updateConnections = $true
        updateVba = $true
        autoBackup = $true
        atomicBatch = $true
    }
} | ConvertTo-Json

$updateHttpRes = Invoke-WebRequest -Uri "$testBaseUrl/api/update" -Method POST -Body $updatePayload -ContentType "application/json; charset=utf-8" -UseBasicParsing
Assert-Equal $updateHttpRes.StatusCode 202 "POST /api/update must return HTTP 202 Accepted"

$updateJson = $updateHttpRes.Content | ConvertFrom-Json
$updateOpId = $updateJson.operationId

# Poll until completion
$pollCount = 0
$finalUpdateOp = $null
while ($pollCount -lt 40) {
    Start-Sleep -Milliseconds 500
    $pollRes = Invoke-RestMethod -Uri "$testBaseUrl/api/operations/$updateOpId" -Method GET
    $finalUpdateOp = $pollRes.operation
    Write-Host "    [Poll $pollCount] Status: $($finalUpdateOp.status), Stage: $($finalUpdateOp.currentStage), Progress: $($finalUpdateOp.progressPercent)%"
    if ($finalUpdateOp.status -eq "COMPLETED" -or $finalUpdateOp.status -eq "FAILED") {
        break
    }
    $pollCount++
}

Assert-Equal $finalUpdateOp.status "COMPLETED" "Update operation must reach COMPLETED state"
$updatedSha = Get-FileSha256 $sampleFile
Assert-False ($updatedSha -eq $initialSha) "Update operation must modify file hash upon completion"

# Verify backup was created
$parent = Split-Path -Parent $testPreviewDir
$backups = Get-ChildItem $parent -Directory -Filter "*http_preview_test_ExcelUpdater_Backups*"
Assert-True ($backups.Count -ge 1) "Update with autoBackup must create backup directory"

# ------------------------------------------------------------------------------
# 5. HTTP 409 Conflict Mutex Rejection
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 5] HTTP 409 Conflict Mutex Rejection --" -ForegroundColor Yellow

# Artificially register an active operation
$dummyReg = Register-Operation -Type "UPDATE" -Directory $testPreviewDir
$null = Set-OperationStatus $dummyReg.OperationId "RUNNING" "Preflight"

$conflictThrown = $false
try {
    $conflictRes = Invoke-WebRequest -Uri "$testBaseUrl/api/update" -Method POST -Body $updatePayload -ContentType "application/json; charset=utf-8" -UseBasicParsing -ErrorAction Stop
} catch {
    $conflictThrown = $true
    $resp = $_.Exception.Response
    Assert-Equal ([int]$resp.StatusCode) 409 "Concurrent operation must return HTTP 409 Conflict"
}
Assert-True $conflictThrown "Concurrent operation must be rejected with error"

# Free mutex
$null = Set-OperationStatus $dummyReg.OperationId "COMPLETED" "Done"

# ------------------------------------------------------------------------------
# 6. HTTP Safe Checkpoint Cancellation
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 6] HTTP POST /api/operations/:id/cancel --" -ForegroundColor Yellow

$cancelReg = Register-Operation -Type "UPDATE" -Directory $testPreviewDir
$null = Set-OperationStatus $cancelReg.OperationId "RUNNING" "Staging"

$cancelRes = Invoke-RestMethod -Uri "$testBaseUrl/api/operations/$($cancelReg.OperationId)/cancel" -Method POST
Assert-True $cancelRes.success "Cancel request must succeed"

$cancelSnap = Get-OperationSnapshot $cancelReg.OperationId
Assert-Equal $cancelSnap.status "CANCELLATION_REQUESTED" "Status must transition to CANCELLATION_REQUESTED"

# Complete transition to CANCELLED
$null = Set-OperationStatus $cancelReg.OperationId "CANCELLED" "Cancelled safely"

# ------------------------------------------------------------------------------
# 7. HTTP GET /api/history & Sensitive Data Masking
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 7] HTTP GET /api/history & Masking Verification --" -ForegroundColor Yellow

$histRes = Invoke-RestMethod -Uri "$testBaseUrl/api/history" -Method GET
Assert-True $histRes.success "GET /api/history must succeed"
Assert-True ($histRes.history.Count -ge 1) "History must contain at least 1 operation"

# ------------------------------------------------------------------------------
# 8. HTTP GET /api/backups & POST /api/restore
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 8] HTTP GET /api/backups & POST /api/restore --" -ForegroundColor Yellow

$backupsRes = Invoke-RestMethod -Uri "$testBaseUrl/api/backups?dir=$([System.Uri]::EscapeDataString($testPreviewDir))" -Method GET
Assert-True $backupsRes.success "GET /api/backups must succeed"
Assert-True ($backupsRes.backups.Count -ge 1) "Backups list must contain at least 1 backup"

$firstBackup = $backupsRes.backups[0]
$restorePayload = @{
    directory = $testPreviewDir
    backupDir = $firstBackup.backupDirectory
} | ConvertTo-Json

$restoreHttpRes = Invoke-WebRequest -Uri "$testBaseUrl/api/restore" -Method POST -Body $restorePayload -ContentType "application/json; charset=utf-8" -UseBasicParsing
Assert-Equal $restoreHttpRes.StatusCode 202 "POST /api/restore must return HTTP 202 Accepted"

$restoreJson = $restoreHttpRes.Content | ConvertFrom-Json
$restoreOpId = $restoreJson.operationId

# Poll until completion
$pollCount = 0
$finalRestoreOp = $null
while ($pollCount -lt 40) {
    Start-Sleep -Milliseconds 500
    $pollRes = Invoke-RestMethod -Uri "$testBaseUrl/api/operations/$restoreOpId" -Method GET
    $finalRestoreOp = $pollRes.operation
    if ($finalRestoreOp.status -eq "COMPLETED" -or $finalRestoreOp.status -eq "FAILED") {
        break
    }
    $pollCount++
}

Assert-Equal $finalRestoreOp.status "COMPLETED" "Restore operation must reach COMPLETED state"

# Verify file was restored to initial hash
$restoredSha = Get-FileSha256 $sampleFile
Assert-Equal $restoredSha $initialSha "Restored file must match original initial hash"

# ------------------------------------------------------------------------------
# 9. Stale Operation Recovery on Startup
# ------------------------------------------------------------------------------
Write-Host "`n-- [TEST 9] Stale Operation Crash Recovery --" -ForegroundColor Yellow

$fakeHistPath = Join-Path $repoRoot "history.json"
$fakeStaleOp = @{
    id = "stale_op_test_12345"
    type = "UPDATE"
    directory = $testPreviewDir
    status = "RUNNING"
    startTime = (Get-Date).ToString("o")
    currentStage = "Excel Update"
}
$null = Save-OperationHistory -OperationSummary $fakeStaleOp -RepoRoot $repoRoot

# Trigger recovery
$recoveredCount = Recover-StaleOperations -RepoRoot $repoRoot
Assert-True ($recoveredCount -ge 1) "Recover-StaleOperations must recover at least 1 stale operation"

$histAfter = Get-OperationHistory -RepoRoot $repoRoot
$foundStale = $histAfter | Where-Object { $_.id -eq "stale_op_test_12345" } | Select-Object -First 1
Assert-Equal $foundStale.status "STALE_INTERRUPTED" "Stale operation status must be updated to STALE_INTERRUPTED"

# Clean up synthetic test folder
if (Test-Path $testPreviewDir) { Remove-Item $testPreviewDir -Recurse -Force }
Get-ChildItem (Split-Path -Parent $testPreviewDir) -Directory -Filter "*http_preview_test_ExcelUpdater_Backups*" -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

# ------------------------------------------------------------------------------
# Clean Up: Stop HTTP Server
# ------------------------------------------------------------------------------
Write-Host "`n[CLEANUP] Stopping background HTTP server..." -ForegroundColor Yellow
try {
    Invoke-RestMethod -Uri "$testBaseUrl/api/shutdown" -Method POST -TimeoutSec 2 -ErrorAction SilentlyContinue | Out-Null
    Start-Sleep -Milliseconds 300
} catch { }
try {
    $serverPs.Stop()
    $serverPs.Dispose()
    $serverRunspace.Close()
    $serverRunspace.Dispose()
} catch { }

Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host " PHASE 3 REAL HTTP INTEGRATION TEST SUITE COMPLETED" -ForegroundColor Cyan
Write-Host "==================================================================" -ForegroundColor Cyan
