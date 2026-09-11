# ==============================================================================
# Test Harness & Assertion Framework for Excel Updater
# ==============================================================================

$global:CorporateBaselineHashes = @{
    "PLAN v4.2.xlsm"              = "F068CDC18AA454946E5322B9C896D07ECF03E7AC0E2322010F63EE754E46A8B2"
    "REÇETE PATLATMA v4.xlsm"     = "667CB704F8DAA40137BB82255EB547E9DE7DA66F6B796ABB33458FA8ABD7657D"
    "SİPARİŞ İHTİYAÇ v3.xlsx"     = "2BA2BD3CDAEE576A0E25682C4C2381821F3B5B60F93D22792A48BF4FF148B90F"
    "ÖZET VE ÜRETİM DATA v1.xlsm" = "7F5E7ADB4EB63F827C10D31662AAA71214D3996934FBAEEB080B4540D29C16B5"
}

$global:TestResults = @{
    Passed = 0
    Failed = 0
    Tests = @()
}

function Assert-Equal ($actual, $expected, $testName) {
    if ($actual -eq $expected) {
        $global:TestResults.Passed++
        $global:TestResults.Tests += @{ Name = $testName; Status = "PASS"; Details = "" }
        Write-Host "  [PASS] $testName" -ForegroundColor Green
    } else {
        $global:TestResults.Failed++
        $msg = "Expected '$expected' but got '$actual'"
        $global:TestResults.Tests += @{ Name = $testName; Status = "FAIL"; Details = $msg }
        Write-Host "  [FAIL] $testName : $msg" -ForegroundColor Red
    }
}

function Assert-True ($condition, $testName) {
    if ($condition -eq $true) {
        $global:TestResults.Passed++
        $global:TestResults.Tests += @{ Name = $testName; Status = "PASS"; Details = "" }
        Write-Host "  [PASS] $testName" -ForegroundColor Green
    } else {
        $global:TestResults.Failed++
        $global:TestResults.Tests += @{ Name = $testName; Status = "FAIL"; Details = "Expected true but was false" }
        Write-Host "  [FAIL] $testName : Expected true" -ForegroundColor Red
    }
}

function Assert-False ($condition, $testName) {
    if ($condition -eq $false) {
        $global:TestResults.Passed++
        $global:TestResults.Tests += @{ Name = $testName; Status = "PASS"; Details = "" }
        Write-Host "  [PASS] $testName" -ForegroundColor Green
    } else {
        $global:TestResults.Failed++
        $global:TestResults.Tests += @{ Name = $testName; Status = "FAIL"; Details = "Expected false but was true" }
        Write-Host "  [FAIL] $testName : Expected false" -ForegroundColor Red
    }
}

function Assert-Throws ($scriptBlock, $testName) {
    $threw = $false
    try {
        & $scriptBlock
    } catch {
        $threw = $true
    }

    if ($threw) {
        $global:TestResults.Passed++
        $global:TestResults.Tests += @{ Name = $testName; Status = "PASS"; Details = "" }
        Write-Host "  [PASS] $testName" -ForegroundColor Green
    } else {
        $global:TestResults.Failed++
        $global:TestResults.Tests += @{ Name = $testName; Status = "FAIL"; Details = "Expected script to throw exception but it succeeded" }
        Write-Host "  [FAIL] $testName : Expected exception" -ForegroundColor Red
    }
}

function Assert-CorporateFilesUntouched ($repoRoot) {
    $allMatches = $true
    foreach ($k in $global:CorporateBaselineHashes.Keys) {
        $filePath = Join-Path $repoRoot $k
        if (Test-Path $filePath) {
            $currentHash = (Get-FileHash $filePath -Algorithm SHA256).Hash
            $expectedHash = $global:CorporateBaselineHashes[$k]
            if ($currentHash -ne $expectedHash) {
                $allMatches = $false
                Write-Host "  [CRITICAL ALARM] Corporate file modified: $k!" -ForegroundColor Red
                Write-Host "    Expected: $expectedHash" -ForegroundColor Red
                Write-Host "    Current : $currentHash" -ForegroundColor Red
            }
        }
    }
    return $allMatches
}
