# ==============================================================================
# SOAK TEST: 50 Consecutive Real COM Operations & Process Integrity
# ==============================================================================

$repoRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path))
if (-not (Test-Path (Join-Path $repoRoot "engine\excel_engine.ps1"))) {
    $repoRoot = (Get-Location).Path
}

. (Join-Path $repoRoot "engine\excel_engine.ps1")

$userPidsBefore = @(Get-Process excel -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
Write-Host "Initial Excel Processes (User Instances to Protect): $($userPidsBefore -join ', ')" -ForegroundColor Cyan

$soakTemp = Join-Path $env:TEMP "soak_test_work"
if (Test-Path $soakTemp) { Remove-Item $soakTemp -Recurse -Force -ErrorAction SilentlyContinue }
New-Item -ItemType Directory -Path $soakTemp -Force | Out-Null

$cycles = 50
$successfulCycles = 0
$orphanDetectedCount = 0
$userProcessKilled = $false

Write-Host "`nStarting Soak Test: $cycles iterations of Isolated Excel COM spawn & shutdown..." -ForegroundColor Yellow

$sw = [System.Diagnostics.Stopwatch]::StartNew()

for ($i = 1; $i -le $cycles; $i++) {
    try {
        $iso = New-IsolatedExcelInstance
        $spawnedPid = $iso.Pid
        
        # Perform real COM operations
        $wbs = $iso.Excel.Workbooks
        $wb = $wbs.Add()
        $ws = $wb.ActiveSheet
        
        $rA1 = $ws.Range("A1")
        $rA1.Value2 = "Soak Test Iteration $($i)"
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($rA1) | Out-Null
        $rA1 = $null

        $rB1 = $ws.Range("B1")
        $rB1.Value2 = (Get-Date).ToString("o")
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($rB1) | Out-Null
        $rB1 = $null

        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) | Out-Null
        $ws = $null
        
        $tempXlsx = Join-Path $soakTemp "cycle_$($i).xlsx"
        $wb.SaveAs($tempXlsx, 51) # xlOpenXMLWorkbook
        $wb.Close($false)
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
        $wb = $null
        
        # Verify snapshot on saved file
        $readWb = $wbs.Open($tempXlsx, 0, $true)
        $snap = Get-WorkbookSnapshot -Workbook $readWb
        $snap = $null
        $readWb.Close($false)
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($readWb) | Out-Null
        $readWb = $null

        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wbs) | Out-Null
        $wbs = $null

        Close-IsolatedExcelInstance $iso
        $iso = $null

        # Check user processes integrity
        foreach ($uPid in $userPidsBefore) {
            if (-not (Get-Process -Id $uPid -ErrorAction SilentlyContinue)) {
                $userProcessKilled = $true
                Write-Host "CRITICAL: User Excel process $uPid was KILLED during cycle $($i)!" -ForegroundColor Red
            }
        }

        # Check for orphan process (allow up to 2.5s for OS process table rundown under heavy load)
        $orphan = @()
        for ($chk = 0; $chk -lt 25; $chk++) {
            $procs = @(Get-Process excel -ErrorAction SilentlyContinue)
            $currentPids = @()
            foreach ($pr in $procs) {
                try {
                    if (-not $pr.HasExited) {
                        $currentPids += $pr.Id
                    }
                } catch { }
                finally {
                    $pr.Dispose()
                }
            }
            $orphan = @($currentPids | Where-Object { $_ -notin $userPidsBefore })
            if ($orphan.Count -eq 0) { break }
            Start-Sleep -Milliseconds 100
        }

        if ($orphan.Count -gt 0) {
            $orphanDetectedCount++
            Write-Host "WARNING: Orphan process detected after cycle $($i): $($orphan -join ', ')" -ForegroundColor Yellow
        }

        $successfulCycles++
        if ($i % 10 -eq 0 -or $i -eq $cycles) {
            Write-Host "  Completed $i / $cycles cycles (Elapsed: $([Math]::Round($sw.Elapsed.TotalSeconds, 1))s | Active User Procs: $($userPidsBefore.Count) | Orphans: $($orphan.Count))" -ForegroundColor Green
        }
    } catch {
        Write-Host "ERROR in cycle $($i): $_" -ForegroundColor Red
    }
}

$sw.Stop()

# Final check (allow up to 3s for final cycle rundown)
$finalOrphans = @()
for ($fChk = 0; $fChk -lt 30; $fChk++) {
    $procs = @(Get-Process excel -ErrorAction SilentlyContinue)
    $finalPids = @()
    foreach ($pr in $procs) {
        try {
            if (-not $pr.HasExited) {
                $finalPids += $pr.Id
            }
        } catch { }
        finally {
            $pr.Dispose()
        }
    }
    $finalOrphans = @($finalPids | Where-Object { $_ -notin $userPidsBefore })
    if ($finalOrphans.Count -eq 0) { break }
    Start-Sleep -Milliseconds 100
}

Write-Host "`n=================================================================="
Write-Host " SOAK TEST COMPLETED"
Write-Host " Total Cycles Planned: $cycles"
Write-Host " Successful Cycles:    $successfulCycles"
Write-Host " User Process Killed:  $userProcessKilled"
Write-Host " Final Orphan Count:   $($finalOrphans.Count)"
Write-Host " Total Time:           $([Math]::Round($sw.Elapsed.TotalSeconds, 1))s"
Write-Host "=================================================================="

# Cleanup soak temp
Remove-Item $soakTemp -Recurse -Force -ErrorAction SilentlyContinue

$passed = ($successfulCycles -eq $cycles -and -not $userProcessKilled -and $finalOrphans.Count -eq 0)

if ($passed) {
    Write-Host "SOAK TEST RESULT: PASS" -ForegroundColor Green
} else {
    Write-Host "SOAK TEST RESULT: FAIL" -ForegroundColor Red
}

return @{
    Passed = $passed
    Cycles = $cycles
    SuccessCount = $successfulCycles
    UserProcessKilled = $userProcessKilled
    FinalOrphanCount = $finalOrphans.Count
    ElapsedSeconds = [Math]::Round($sw.Elapsed.TotalSeconds, 1)
}
