# ==============================================================================
# Excel Automation Engine - Scanning, Backup & Bulk Updating via COM Interop
# ==============================================================================

$global:ProgressState = @{
    active = $false
    type = "idle"
    current = 0
    total = 0
    currentFile = ""
    percent = 0
}

function Set-ProgressState ($active, $type, $current, $total, $currentFile) {
    $percent = if ($total -gt 0) { [Math]::Min(100, [Math]::Round(($current / $total) * 100)) } else { 0 }
    $global:ProgressState = @{
        active = $active
        type = $type
        current = $current
        total = $total
        currentFile = $currentFile
        percent = $percent
    }
}

function New-OperationId () {
    return (Get-Date).ToString("yyyyMMdd_HHmmss") + "_" + [Guid]::NewGuid().ToString("N").Substring(0, 8)
}

function Get-FileSha256 ([string]$FilePath) {
    if (-not (Test-Path $FilePath)) { return $null }
    $stream = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $sha = [System.Security.Cryptography.SHA256]::Create()
        $hashBytes = $sha.ComputeHash($stream)
        return [System.BitConverter]::ToString($hashBytes).Replace("-", "").ToUpper()
    } finally {
        $stream.Close()
        $stream.Dispose()
        if ($sha) { $sha.Dispose() }
    }
}

function Test-FileWritable ([string]$FilePath) {
    if (-not (Test-Path $FilePath)) { return $false }
    try {
        $stream = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $stream.Close()
        $stream.Dispose()
        return $true
    } catch {
        return $false
    }
}

function Test-ReplacementRulesValid ($Rules) {
    if ($null -eq $Rules -or $Rules.Count -eq 0) {
        return @{ IsValid = $false; Error = "Hiçbir değişim kuralı belirtilmedi." ; ValidRules = @() }
    }

    $seenOld = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $validRules = @()

    foreach ($r in $Rules) {
        $oldText = if ($r.oldText) { $r.oldText.Trim() } else { "" }
        $newText = if ($r.newText) { $r.newText.Trim() } else { "" }

        if ([string]::IsNullOrWhiteSpace($oldText) -or [string]::IsNullOrWhiteSpace($newText)) {
            continue
        }

        if ($oldText -eq $newText) {
            return @{ IsValid = $false; Error = "Eski ve yeni metin aynı olamaz: '$oldText'"; ValidRules = @() }
        }

        if ($seenOld.Contains($oldText)) {
            return @{ IsValid = $false; Error = "Aynı arama metni için birden fazla kural tanımlanamaz (Çift kural): '$oldText'"; ValidRules = @() }
        }

        $seenOld.Add($oldText) | Out-Null
        $validRules += @{ oldText = $oldText; newText = $newText }
    }

    if ($validRules.Count -eq 0) {
        return @{ IsValid = $false; Error = "Geçerli bir arama/değiştirme kuralı bulunamadı."; ValidRules = @() }
    }

    return @{ IsValid = $true; Error = ""; ValidRules = $validRules }
}

function Invoke-SafeReplacement ([string]$InputText, [array]$Rules, [string]$Mode = "Auto") {
    if ([string]::IsNullOrEmpty($InputText) -or $null -eq $Rules -or $Rules.Count -eq 0) {
        return @{
            ResultText = $InputText
            Modified = $false
            ReplacementsCount = 0
            Details = @()
        }
    }

    $validation = Test-ReplacementRulesValid -Rules $Rules
    if (-not $validation.IsValid) {
        throw "Değiştirme kuralları geçersiz: $($validation.Error)"
    }
    $validRules = $validation.ValidRules

    $patterns = @()
    for ($i = 0; $i -lt $validRules.Count; $i++) {
        $old = $validRules[$i].oldText
        $escaped = [regex]::Escape($old)
        
        # Check if oldText is an IPv4 address
        $isIpv4 = $old -match '^(\d{1,3}\.){3}\d{1,3}$'
        if ($isIpv4) {
            # Strict boundary: cannot be preceded or followed by a digit or dot
            $pattern = "(?<r$i>(?<![\d\.])$escaped(?![\d\.]))"
        } else {
            # For hostnames, instances or words: boundary check
            $prefix = if ($old -match '^\w') { "(?<!\w)" } else { "" }
            $suffix = if ($old -match '\w$') { "(?!\w)" } else { "" }
            $pattern = "(?<r$i>$prefix$escaped$suffix)"
        }
        $patterns += $pattern
    }

    $combinedRegexStr = $patterns -join "|"
    $combinedRegex = [regex]::new($combinedRegexStr, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

    $replacementsCount = 0
    $details = @()

    $evaluator = [System.Text.RegularExpressions.MatchEvaluator]{
        param([System.Text.RegularExpressions.Match]$m)
        for ($k = 0; $k -lt $validRules.Count; $k++) {
            $grp = "r$k"
            if ($m.Groups[$grp].Success) {
                $script:curReplacementsCount++
                $matchedRule = $validRules[$k]
                $script:curDetails += "Eşleşme: '$($m.Value)' -> '$($matchedRule.newText)'"
                return $matchedRule.newText
            }
        }
        return $m.Value
    }

    $script:curReplacementsCount = 0
    $script:curDetails = @()

    $resultText = $combinedRegex.Replace($InputText, $evaluator)

    return @{
        ResultText = $resultText
        Modified = ($resultText -ne $InputText)
        ReplacementsCount = $script:curReplacementsCount
        Details = $script:curDetails
    }
}

function Cleanup-ExcelCOM ($excel) {
    if ($excel) {
        try { $excel.Quit() } catch { }
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { }
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

function Get-ExcelFiles ($DirectoryPath) {
    if (-not (Test-Path $DirectoryPath)) {
        return @()
    }
    return Get-ChildItem -Path $DirectoryPath -File | 
        Where-Object { 
            ($_.Extension -eq ".xlsx" -or $_.Extension -eq ".xlsm" -or $_.Extension -eq ".xlsb") -and
            -not $_.Name.StartsWith("~$") -and 
            -not $_.Name.StartsWith("backup_") -and
            -not $_.Name.Contains(".staging.") -and
            -not $_.Name.EndsWith(".old") -and
            -not $_.DirectoryName.Contains("_ExcelUpdater_Backups")
        }
}

function Copy-FileWithShare ($srcPath, $dstPath) {
    $srcStream = New-Object System.IO.FileStream($srcPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $dstStream = New-Object System.IO.FileStream($dstPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        try {
            $srcStream.CopyTo($dstStream)
        } finally {
            $dstStream.Close()
            $dstStream.Dispose()
        }
    } finally {
        $srcStream.Close()
        $srcStream.Dispose()
    }
}

function Scan-ExcelDirectory ($DirectoryPath) {
    $files = Get-ExcelFiles $DirectoryPath
    $fileList = @()
    $ipSummary = @{}
    $totalCount = $files.Count

    Set-ProgressState $true "scan" 0 $totalCount "Taramaya Başlanıyor..."

    if ($totalCount -eq 0) {
        Set-ProgressState $false "scan" 0 0 ""
        return @{
            success = $true
            message = "No Excel files found in directory."
            directory = $DirectoryPath
            totalFiles = 0
            files = @()
            detectedIPs = @()
        }
    }

    $excel = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false
    } catch {
        Set-ProgressState $false "scan" 0 0 ""
        return @{
            success = $false
            error = "Could not initialize Excel COM object: $_"
        }
    }

    $ipRegex = '\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b'
    $processedCount = 0

    foreach ($file in $files) {
        $processedCount++
        Set-ProgressState $true "scan" $processedCount $totalCount $file.Name

        $fileDetail = @{
            filePath = $file.FullName
            fileName = $file.Name
            extension = $file.Extension.ToLower()
            sizeBytes = $file.Length
            queries = @()
            connections = @()
            vbaMatches = @()
            foundIPs = @()
            hasVba = $false
            status = "Scanned"
        }

        try {
            $wb = $excel.Workbooks.Open($file.FullName, 0, $true) # Read-only open

            # 1. Check Power Queries
            try {
                foreach ($q in $wb.Queries) {
                    $formula = $q.Formula
                    $matches = [regex]::Matches($formula, $ipRegex)
                    $ips = @($matches | ForEach-Object { $_.Value })
                    
                    $qInfo = @{
                        name = $q.Name
                        formulaSnippet = if ($formula.Length -gt 150) { $formula.Substring(0, 150) + "..." } else { $formula }
                        fullFormula = $formula
                        ips = $ips
                    }
                    $fileDetail.queries += $qInfo

                    foreach ($ip in $ips) {
                        if (-not $fileDetail.foundIPs.Contains($ip)) { $fileDetail.foundIPs += $ip }
                        if (-not $ipSummary.ContainsKey($ip)) { $ipSummary[$ip] = 0 }
                        $ipSummary[$ip]++
                    }
                }
            } catch { }

            # 2. Check ALL Data Connections
            try {
                foreach ($conn in $wb.Connections) {
                    $connStr = ""
                    $cmdText = ""
                    try {
                        if ($conn.OLEDBConnection) {
                            $connStr = $conn.OLEDBConnection.ConnectionString
                            $cmdText = $conn.OLEDBConnection.CommandText
                        } elseif ($conn.ODBCConnection) {
                            $connStr = $conn.ODBCConnection.ConnectionString
                            $cmdText = $conn.ODBCConnection.CommandText
                        }
                    } catch { }

                    $combinedText = "$($conn.Name) $connStr $cmdText $($conn.Description)"
                    $matches = [regex]::Matches($combinedText, $ipRegex)
                    $ips = @($matches | ForEach-Object { $_.Value })

                    $cInfo = @{
                        name = $conn.Name
                        type = $conn.Type
                        connectionString = $connStr
                        commandText = $cmdText
                        ips = $ips
                    }
                    $fileDetail.connections += $cInfo

                    foreach ($ip in $ips) {
                        if (-not $fileDetail.foundIPs.Contains($ip)) { $fileDetail.foundIPs += $ip }
                        if (-not $ipSummary.ContainsKey($ip)) { $ipSummary[$ip] = 0 }
                        $ipSummary[$ip]++
                    }
                }
            } catch { }

            # 3. Check Worksheet QueryTables
            try {
                foreach ($ws in $wb.Worksheets) {
                    foreach ($qt in $ws.QueryTables) {
                        $qtConn = ""
                        try { $qtConn = $qt.Connection } catch { }
                        if (-not [string]::IsNullOrWhiteSpace($qtConn)) {
                            $matches = [regex]::Matches($qtConn, $ipRegex)
                            foreach ($m in $matches) {
                                $ip = $m.Value
                                if (-not $fileDetail.foundIPs.Contains($ip)) { $fileDetail.foundIPs += $ip }
                                if (-not $ipSummary.ContainsKey($ip)) { $ipSummary[$ip] = 0 }
                                $ipSummary[$ip]++
                            }
                        }
                    }
                }
            } catch { }

            # 4. Check VBA Macros
            if ($file.Extension.ToLower() -in @(".xlsm", ".xlsb")) {
                $fileDetail.hasVba = $true
                try {
                    foreach ($comp in $wb.VBProject.VBComponents) {
                        $cm = $comp.CodeModule
                        if ($cm.CountOfLines -gt 0) {
                            $code = $cm.Lines(1, $cm.CountOfLines)
                            $matches = [regex]::Matches($code, $ipRegex)
                            if ($matches.Count -gt 0) {
                                foreach ($m in $matches) {
                                    $ip = $m.Value
                                    $vMatch = @{
                                        module = $comp.Name
                                        ip = $ip
                                    }
                                    $fileDetail.vbaMatches += $vMatch
                                    if (-not $fileDetail.foundIPs.Contains($ip)) { $fileDetail.foundIPs += $ip }
                                    if (-not $ipSummary.ContainsKey($ip)) { $ipSummary[$ip] = 0 }
                                    $ipSummary[$ip]++
                                }
                            }
                        }
                    }
                } catch {
                    $fileDetail.vbaWarning = "VBA access restricted or protected."
                }
            }

            $wb.Close($false)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
        } catch {
            $fileDetail.status = "Error: $_"
        }

        $fileList += $fileDetail
    }

    Cleanup-ExcelCOM $excel
    Set-ProgressState $false "scan" $totalCount $totalCount "Tarama Tamamlandı!"

    $detectedIPsList = @()
    foreach ($k in $ipSummary.Keys) {
        $detectedIPsList += @{ ip = $k; count = $ipSummary[$k] }
    }

    return @{
        success = $true
        directory = $DirectoryPath
        totalFiles = $fileList.Count
        files = $fileList
        detectedIPs = $detectedIPsList
    }
}

function Create-ExcelBackup ($DirectoryPath) {
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    $files = Get-ExcelFiles $DirectoryPath
    if ($files.Count -eq 0) {
        return @{ success = $false; error = "No Excel files found to backup." }
    }

    $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $backupParent = Join-Path $env:TEMP "excel_backups"
    $backupDir = Join-Path $backupParent "backup_$timestamp"

    try {
        if (-not (Test-Path $backupDir)) {
            New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
        }
        $backedUpCount = 0

        foreach ($file in $files) {
            $destPath = Join-Path $backupDir $file.Name
            Copy-FileWithShare $file.FullName $destPath
            $backedUpCount++
        }

        return @{
            success = $true
            backupDirectory = $backupDir
            totalFiles = $backedUpCount
            timestamp = $timestamp
        }
    } catch {
        return @{ success = $false; error = "Backup failed: $_" }
    }
}

function Update-ExcelDirectory ($DirectoryPath, $Rules, $Options) {
    $files = Get-ExcelFiles $DirectoryPath
    $updateLog = @()
    $updatedFilesCount = 0
    $totalReplacements = 0
    $totalCount = $files.Count

    Set-ProgressState $true "update" 0 $totalCount "Güncellemeye Başlanıyor..."

    if ($totalCount -eq 0) {
        Set-ProgressState $false "update" 0 0 ""
        return @{ success = $false; error = "No Excel files found to update." }
    }

    # Validate rules
    $validRules = @()
    foreach ($r in $Rules) {
        if (-not [string]::IsNullOrWhiteSpace($r.oldText) -and -not [string]::IsNullOrWhiteSpace($r.newText)) {
            $validRules += @{ oldText = $r.oldText; newText = $r.newText }
        }
    }

    if ($validRules.Count -eq 0) {
        Set-ProgressState $false "update" 0 0 ""
        return @{ success = $false; error = "No valid search/replace rules provided." }
    }

    # Check auto-backup option
    if ($Options -and $Options.autoBackup -eq $true) {
        $backupRes = Create-ExcelBackup -DirectoryPath $DirectoryPath
        if (-not $backupRes.success) {
            Set-ProgressState $false "update" 0 0 ""
            return @{ success = $false; error = "Auto-backup failed prior to update: $($backupRes.error)" }
        }
    }

    $excel = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false
    } catch {
        Set-ProgressState $false "update" 0 0 ""
        return @{ success = $false; error = "Could not initialize Excel COM object: $_" }
    }

    $updatePowerQueries = if ($Options -and $Options.updateQueries -ne $null) { $Options.updateQueries } else { $true }
    $updateConnections  = if ($Options -and $Options.updateConnections -ne $null) { $Options.updateConnections } else { $true }
    $updateVba          = if ($Options -and $Options.updateVba -ne $null) { $Options.updateVba } else { $true }

    $processedCount = 0

    foreach ($file in $files) {
        $processedCount++
        Set-ProgressState $true "update" $processedCount $totalCount $file.Name

        $fileLog = @{
            fileName = $file.Name
            filePath = $file.FullName
            status = "Skipped"
            changesMade = 0
            details = @()
        }

        try {
            $wb = $excel.Workbooks.Open($file.FullName, 0, $false) # Open read-write
            $fileModified = $false

            # 1. Update Power Queries
            if ($updatePowerQueries) {
                try {
                    foreach ($q in $wb.Queries) {
                        $formula = $q.Formula
                        $newFormula = $formula
                        foreach ($rule in $validRules) {
                            if ($newFormula.Contains($rule.oldText)) {
                                $newFormula = $newFormula.Replace($rule.oldText, $rule.newText)
                                $fileLog.changesMade++
                                $totalReplacements++
                                $fileLog.details += "PowerQuery '$($q.Name)': '$($rule.oldText)' -> '$($rule.newText)'"
                            }
                        }
                        if ($newFormula -ne $formula) {
                            $q.Formula = $newFormula
                            $fileModified = $true
                        }
                    }
                } catch {
                    $fileLog.details += "PowerQuery error: $_"
                }
            }

            # 2. Update Data Connections
            if ($updateConnections) {
                try {
                    foreach ($conn in $wb.Connections) {
                        $oldName = $conn.Name
                        $newName = $oldName
                        foreach ($rule in $validRules) {
                            if ($newName.Contains($rule.oldText)) {
                                $newName = $newName.Replace($rule.oldText, $rule.newText)
                                $fileLog.changesMade++
                                $totalReplacements++
                                $fileLog.details += "Connection Name '$oldName': '$($rule.oldText)' -> '$($rule.newText)'"
                            }
                        }
                        if ($newName -ne $oldName) {
                            try { $conn.Name = $newName } catch { }
                            $fileModified = $true
                        }

                        try {
                            if ($conn.OLEDBConnection) {
                                $cStr = $conn.OLEDBConnection.ConnectionString
                                $newCStr = $cStr
                                foreach ($rule in $validRules) {
                                    if ($newCStr.Contains($rule.oldText)) {
                                        $newCStr = $newCStr.Replace($rule.oldText, $rule.newText)
                                        $fileLog.changesMade++
                                        $totalReplacements++
                                        $fileLog.details += "OLEDB ConnStr in '$($conn.Name)': '$($rule.oldText)' -> '$($rule.newText)'"
                                    }
                                }
                                if ($newCStr -ne $cStr) {
                                    $conn.OLEDBConnection.ConnectionString = $newCStr
                                    $fileModified = $true
                                }

                                $cmd = $conn.OLEDBConnection.CommandText
                                $newCmd = $cmd
                                foreach ($rule in $validRules) {
                                    if ($newCmd.Contains($rule.oldText)) {
                                        $newCmd = $newCmd.Replace($rule.oldText, $rule.newText)
                                        $fileLog.changesMade++
                                        $totalReplacements++
                                        $fileLog.details += "OLEDB CommandText in '$($conn.Name)': '$($rule.oldText)' -> '$($rule.newText)'"
                                    }
                                }
                                if ($newCmd -ne $cmd) {
                                    $conn.OLEDBConnection.CommandText = $newCmd
                                    $fileModified = $true
                                }
                            }
                        } catch { }

                        try {
                            if ($conn.ODBCConnection) {
                                $cStr = $conn.ODBCConnection.ConnectionString
                                $newCStr = $cStr
                                foreach ($rule in $validRules) {
                                    if ($newCStr.Contains($rule.oldText)) {
                                        $newCStr = $newCStr.Replace($rule.oldText, $rule.newText)
                                        $fileLog.changesMade++
                                        $totalReplacements++
                                        $fileLog.details += "ODBC ConnStr in '$($conn.Name)': '$($rule.oldText)' -> '$($rule.newText)'"
                                    }
                                }
                                if ($newCStr -ne $cStr) {
                                    $conn.ODBCConnection.ConnectionString = $newCStr
                                    $fileModified = $true
                                }

                                $cmd = $conn.ODBCConnection.CommandText
                                $newCmd = $cmd
                                foreach ($rule in $validRules) {
                                    if ($newCmd.Contains($rule.oldText)) {
                                        $newCmd = $newCmd.Replace($rule.oldText, $rule.newText)
                                        $fileLog.changesMade++
                                        $totalReplacements++
                                        $fileLog.details += "ODBC CommandText in '$($conn.Name)': '$($rule.oldText)' -> '$($rule.newText)'"
                                    }
                                }
                                if ($newCmd -ne $cmd) {
                                    $conn.ODBCConnection.CommandText = $newCmd
                                    $fileModified = $true
                                }
                            }
                        } catch { }
                    }
                } catch {
                    $fileLog.details += "Connection error: $_"
                }
            }

            # 3. Update QueryTables
            if ($updateConnections) {
                try {
                    foreach ($ws in $wb.Worksheets) {
                        foreach ($qt in $ws.QueryTables) {
                            try {
                                $qtConn = $qt.Connection
                                $newQtConn = $qtConn
                                foreach ($rule in $validRules) {
                                    if ($newQtConn.Contains($rule.oldText)) {
                                        $newQtConn = $newQtConn.Replace($rule.oldText, $rule.newText)
                                        $fileLog.changesMade++
                                        $totalReplacements++
                                        $fileLog.details += "QueryTable Connection on '$($ws.Name)': '$($rule.oldText)' -> '$($rule.newText)'"
                                    }
                                }
                                if ($newQtConn -ne $qtConn) {
                                    $qt.Connection = $newQtConn
                                    $fileModified = $true
                                }
                            } catch { }
                        }
                    }
                } catch { }
            }

            # 4. Update VBA Macros
            if ($updateVba -and ($file.Extension.ToLower() -in @(".xlsm", ".xlsb"))) {
                try {
                    foreach ($comp in $wb.VBProject.VBComponents) {
                        $cm = $comp.CodeModule
                        if ($cm.CountOfLines -gt 0) {
                            $lineCount = $cm.CountOfLines
                            for ($i = 1; $i -le $lineCount; $i++) {
                                $line = $cm.Lines($i, 1)
                                $newLine = $line
                                foreach ($rule in $validRules) {
                                    if ($newLine.Contains($rule.oldText)) {
                                        $newLine = $newLine.Replace($rule.oldText, $rule.newText)
                                        $fileLog.changesMade++
                                        $totalReplacements++
                                        $fileLog.details += "VBA Line $i in '$($comp.Name)': '$($rule.oldText)' -> '$($rule.newText)'"
                                    }
                                }
                                if ($newLine -ne $line) {
                                    $cm.ReplaceLine($i, $newLine)
                                    $fileModified = $true
                                }
                            }
                        }
                    }
                } catch {
                    $fileLog.details += "VBA access warning: $_"
                }
            }

            if ($fileModified) {
                $wb.Save()
                $fileLog.status = "Updated"
                $updatedFilesCount++
            } else {
                $fileLog.status = "No Changes"
            }

            $wb.Close($false)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
        } catch {
            $fileLog.status = "Error"
            $fileLog.details += "File error: $_"
        }

        $updateLog += $fileLog
    }

    Cleanup-ExcelCOM $excel
    Set-ProgressState $false "update" $totalCount $totalCount "Güncelleme Tamamlandı!"

    return @{
        success = $true
        directory = $DirectoryPath
        totalFilesProcessed = $files.Count
        updatedFilesCount = $updatedFilesCount
        totalReplacements = $totalReplacements
        logs = $updateLog
    }
}
