# ==============================================================================
# Generate Synthetic Test Fixtures for Excel SQL Connect Pro
# (NEVER touches corporate files)
# ==============================================================================
$fixtureDir = Join-Path $PSScriptRoot "fixtures"
if (-not (Test-Path $fixtureDir)) { New-Item -ItemType Directory -Path $fixtureDir -Force | Out-Null }

$normFile = Join-Path $fixtureDir "test_norm.xlsx"
$macroFile = Join-Path $fixtureDir "test_macro.xlsm"
$binaryFile = Join-Path $fixtureDir "test_binary.xlsb"
$collisionFile = Join-Path $fixtureDir "test_collision.xlsx"

Write-Host "Checking / Generating synthetic test fixtures in $fixtureDir..."

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    # 1. test_norm.xlsx
    if (-not (Test-Path $normFile)) {
        $wb = $excel.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1)
        $ws.Name = "DataSheet"
        $ws.Cells.Item(1, 1).Value2 = "Server"
        $ws.Cells.Item(1, 2).Value2 = "192.168.1.50"
        $ws.Cells.Item(2, 1).Value2 = "Formula"
        $ws.Cells.Item(2, 2).Formula = '="Connected to " & B1'
        try {
            $wb.Queries.Add("TestQuery", 'let Source = Sql.Database("192.168.1.50", "TestDB") in Source', "Test SQL Query") | Out-Null
        } catch { }
        $wb.SaveAs($normFile, 51) # 51 = xlOpenXMLWorkbook (.xlsx)
        $wb.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
        Write-Host "  Generated test_norm.xlsx"
    }

    # 2. test_macro.xlsm
    if (-not (Test-Path $macroFile)) {
        $wb = $excel.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1)
        $ws.Cells.Item(1, 1).Value2 = "Macro Test"
        try {
            $comp = $wb.VBProject.VBComponents.Add(1) # 1 = vbext_ct_StdModule
            $comp.Name = "ModTest"
            $code = @"
Sub ConnectDB()
    Dim connStr As String
    connStr = "Server=192.168.1.1;Database=TestDB;"
End Sub
"@
            $comp.CodeModule.AddFromString($code)
        } catch {
            Write-Host "  Note: VBProject programmatic access not enabled or skipped: $_"
        }
        $wb.SaveAs($macroFile, 52) # 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
        $wb.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
        Write-Host "  Generated test_macro.xlsm"
    }

    # 3. test_binary.xlsb
    if (-not (Test-Path $binaryFile)) {
        $wb = $excel.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1)
        $ws.Cells.Item(1, 1).Value2 = "Binary Test"
        $ws.Cells.Item(1, 2).Value2 = "192.168.1.1"
        $wb.SaveAs($binaryFile, 50) # 50 = xlExcel12 (.xlsb)
        $wb.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
        Write-Host "  Generated test_binary.xlsb"
    }

    # 4. test_collision.xlsx
    if (-not (Test-Path $collisionFile)) {
        $wb = $excel.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1)
        $ws.Cells.Item(1, 1).Value2 = "Target"
        $ws.Cells.Item(1, 2).Value2 = "192.168.1.1"
        $ws.Cells.Item(2, 1).Value2 = "Collision1"
        $ws.Cells.Item(2, 2).Value2 = "192.168.1.15"
        $ws.Cells.Item(3, 1).Value2 = "Collision2"
        $ws.Cells.Item(3, 2).Value2 = "192.168.1.100"
        $wb.SaveAs($collisionFile, 51)
        $wb.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
        Write-Host "  Generated test_collision.xlsx"
    }
} catch {
    Write-Host "Fixture generation error: $_"
} finally {
    if ($excel) {
        try { $excel.Quit() } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { }
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
