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
    try {
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
    } catch {
        return $null
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

    return @{ IsValid = $true; Error = ""; ValidRules = $validRules; Rules = $validRules }
}

function Invoke-SafeReplacement ([string]$InputText, [array]$Rules, [string]$Mode = "Auto") {
    if ([string]::IsNullOrEmpty($InputText) -or $null -eq $Rules -or $Rules.Count -eq 0) {
        return @{
            ResultText = $InputText
            Modified = $false
            ReplacementsCount = 0
            MatchCount = 0
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
        MatchCount = $script:curReplacementsCount
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
    if (-not (Test-Path -LiteralPath $DirectoryPath)) {
        return @()
    }
    $raw = Get-ChildItem -LiteralPath $DirectoryPath -File -ErrorAction SilentlyContinue | 
        Where-Object { 
            ($_.Extension -eq ".xlsx" -or $_.Extension -eq ".xlsm" -or $_.Extension -eq ".xlsb") -and
            -not $_.Name.StartsWith("~$") -and 
            -not $_.Name.StartsWith("backup_") -and
            -not $_.Name.Contains(".staging.") -and
            -not $_.Name.EndsWith(".old") -and
            -not $_.DirectoryName.Contains("_ExcelUpdater_Backups")
        }
    if ($raw) {
        return @($raw)
    }
    return @()
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

function Get-ComTextProperty {
    param(
        [Parameter(Mandatory=$true)]$ComObject,
        [Parameter(Mandatory=$true)][string[]]$PropertyNames
    )

    foreach ($propName in $PropertyNames) {
        try {
            $value = $ComObject.$propName
            if ($null -ne $value) {
                $text = if ($value -is [System.Array]) { ($value -join "`n") } else { "$value" }
                if (-not [string]::IsNullOrWhiteSpace($text)) {
                    return @{ Success = $true; PropertyName = $propName; Value = $text }
                }
            }
        } catch { }
    }

    return @{ Success = $false; PropertyName = ""; Value = "" }
}

function Set-ComTextProperty {
    param(
        [Parameter(Mandatory=$true)]$ComObject,
        [Parameter(Mandatory=$true)][string[]]$PropertyNames,
        [Parameter(Mandatory=$true)][string]$Value,
        [string]$PreferredProperty = ""
    )

    $orderedProps = @()
    if (-not [string]::IsNullOrWhiteSpace($PreferredProperty)) {
        $orderedProps += $PreferredProperty
    }
    foreach ($propName in $PropertyNames) {
        if ($orderedProps -notcontains $propName) {
            $orderedProps += $propName
        }
    }

    $lastError = ""
    foreach ($propName in $orderedProps) {
        try {
            $ComObject.$propName = $Value
            return @{ Success = $true; PropertyName = $propName; Error = "" }
        } catch {
            $lastError = "$_"
        }
    }

    return @{ Success = $false; PropertyName = ""; Error = $lastError }
}

function Get-ExcelConnectionText ($ConnectionObject) {
    # Excel WorkbookConnection.OLEDBConnection / ODBCConnection stores the live
    # connection string in .Connection. Some COM wrappers also expose
    # .ConnectionString, so read both for compatibility.
    return Get-ComTextProperty -ComObject $ConnectionObject -PropertyNames @("Connection", "ConnectionString")
}

function Set-ExcelConnectionText ($ConnectionObject, [string]$ConnectionText, [string]$PreferredProperty = "") {
    return Set-ComTextProperty -ComObject $ConnectionObject -PropertyNames @("Connection", "ConnectionString") -Value $ConnectionText -PreferredProperty $PreferredProperty
}

function Scan-ExcelDirectory ($DirectoryPath, [string]$OperationId = "") {
    if ($OperationId -and (Get-Command "Write-AuditLogEvent" -ErrorAction SilentlyContinue)) {
        Write-AuditLogEvent -OperationId $OperationId -Level "INFO" -EventName "SCAN_STARTED" -Stage "Preflight" -Message "Excel directory scan requested" -Data @{ directory = $DirectoryPath }
    }

    $files = Get-ExcelFiles $DirectoryPath
    $fileList = @()
    $ipSummary = @{}
    $totalCount = $files.Count

    if ($OperationId -and (Get-Command "Write-AuditLogEvent" -ErrorAction SilentlyContinue)) {
        Write-AuditLogEvent -OperationId $OperationId -Level "INFO" -EventName "SCAN_FILE_ENUMERATION_COMPLETED" -Stage "Preflight" -Message "Excel files enumerated" -Data @{ directory = $DirectoryPath; totalFiles = $totalCount }
    }

    Set-ProgressState $true "scan" 0 $totalCount "Taramaya Başlanıyor..."
    if ($OperationId -and (Get-Command "Update-OperationProgress" -ErrorAction SilentlyContinue)) {
        $null = Update-OperationProgress -OperationId $OperationId -ProcessedFiles 0 -TotalFiles $totalCount -CurrentFile "Taramaya Başlanıyor..." -CurrentStage "Scanning" -ProgressPercent 0
    }

    if ($totalCount -eq 0) {
        Set-ProgressState $false "scan" 0 0 ""
        if ($OperationId -and (Get-Command "Update-OperationProgress" -ErrorAction SilentlyContinue)) {
            $null = Update-OperationProgress -OperationId $OperationId -ProcessedFiles 0 -TotalFiles 0 -CurrentFile "Dosya Bulunamadı" -CurrentStage "Completed" -ProgressPercent 100
        }
        if ($OperationId -and (Get-Command "Write-AuditLogEvent" -ErrorAction SilentlyContinue)) {
            Write-AuditLogEvent -OperationId $OperationId -Level "INFO" -EventName "SCAN_COMPLETED_EMPTY" -Stage "Completed" -Message "No Excel files found in directory" -Data @{ directory = $DirectoryPath }
        }
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
        if ($OperationId -and (Get-Command "Write-AuditLogEvent" -ErrorAction SilentlyContinue)) {
            Write-AuditLogEvent -OperationId $OperationId -Level "INFO" -EventName "EXCEL_COM_INITIALIZING" -Stage "Excel COM" -Message "Creating Excel.Application COM object"
        }
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false
        $excel.EnableEvents = $false
        $excel.AskToUpdateLinks = $false
        try {
            $excel.AutomationSecurity = 3 # msoAutomationSecurityForceDisable
        } catch { }
        if ($OperationId -and (Get-Command "Write-AuditLogEvent" -ErrorAction SilentlyContinue)) {
            Write-AuditLogEvent -OperationId $OperationId -Level "INFO" -EventName "EXCEL_COM_READY" -Stage "Excel COM" -Message "Excel.Application COM object is ready" -Data @{ version = "$($excel.Version)" }
        }
    } catch {
        Set-ProgressState $false "scan" 0 0 ""
        if ($OperationId -and (Get-Command "Write-AuditLogEvent" -ErrorAction SilentlyContinue)) {
            Write-AuditLogEvent -OperationId $OperationId -Level "ERROR" -EventName "EXCEL_COM_INITIALIZE_FAILED" -Stage "Excel COM" -Message "Could not initialize Excel COM object: $_"
        }
        return @{
            success = $false
            error = "Could not initialize Excel COM object: $_"
        }
    }

    $ipRegex = '\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b'
    $processedCount = 0

    foreach ($file in $files) {
        # Cooperative cancellation checkpoint
        if ($OperationId -and (Get-Command "Test-OperationCancellationRequested" -ErrorAction SilentlyContinue) -and (Test-OperationCancellationRequested $OperationId)) {
            if (Get-Command "Set-OperationStatus" -ErrorAction SilentlyContinue) {
                $null = Set-OperationStatus $OperationId "CANCELLED" "Tarama kullanıcı tarafından iptal edildi"
            }
            break
        }

        $processedCount++
        $pct = if ($totalCount -gt 0) { [math]::Min(99, [math]::Max(1, [math]::Round((($processedCount - 1) / $totalCount) * 100))) } else { 0 }
        Set-ProgressState $true "scan" $processedCount $totalCount $file.Name
        if ($OperationId -and (Get-Command "Update-OperationProgress" -ErrorAction SilentlyContinue)) {
            $null = Update-OperationProgress -OperationId $OperationId -ProcessedFiles $processedCount -TotalFiles $totalCount -CurrentFile $file.Name -CurrentStage "Scanning" -ProgressPercent $pct
        }
        if ($OperationId -and (Get-Command "Write-AuditLogEvent" -ErrorAction SilentlyContinue)) {
            Write-AuditLogEvent -OperationId $OperationId -Level "INFO" -EventName "SCAN_FILE_STARTED" -File $file.FullName -Stage "WorkbookOpen" -Message "Opening workbook for read-only scan" -Data @{ index = $processedCount; totalFiles = $totalCount; sizeBytes = $file.Length; extension = $file.Extension }
        }

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
            $wb = $excel.Workbooks.Open($file.FullName, 0, $true, [Type]::Missing, [Type]::Missing, [Type]::Missing, $true) # Read-only open, IgnoreReadOnlyRecommended
            if ($OperationId -and (Get-Command "Write-AuditLogEvent" -ErrorAction SilentlyContinue)) {
                Write-AuditLogEvent -OperationId $OperationId -Level "INFO" -EventName "SCAN_WORKBOOK_OPENED" -File $file.FullName -Stage "WorkbookOpen" -Message "Workbook opened successfully"
            }

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
            } catch {
                if ($OperationId -and (Get-Command "Write-AuditLogEvent" -ErrorAction SilentlyContinue)) {
                    Write-AuditLogEvent -OperationId $OperationId -Level "WARNING" -EventName "SCAN_POWER_QUERY_READ_FAILED" -File $file.FullName -Stage "Power Query" -Message "$_"
                }
            }

            # 2. Check ALL Data Connections
            try {
                foreach ($conn in $wb.Connections) {
                    $connStr = ""
                    $cmdText = ""
                    try {
                        if ($conn.OLEDBConnection) {
                            $connText = Get-ExcelConnectionText -ConnectionObject $conn.OLEDBConnection
                            $connStr = $connText.Value
                            $cmdText = $conn.OLEDBConnection.CommandText
                        } elseif ($conn.ODBCConnection) {
                            $connText = Get-ExcelConnectionText -ConnectionObject $conn.ODBCConnection
                            $connStr = $connText.Value
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
            } catch {
                if ($OperationId -and (Get-Command "Write-AuditLogEvent" -ErrorAction SilentlyContinue)) {
                    Write-AuditLogEvent -OperationId $OperationId -Level "WARNING" -EventName "SCAN_CONNECTION_READ_FAILED" -File $file.FullName -Stage "Connections" -Message "$_"
                }
            }

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
            } catch {
                if ($OperationId -and (Get-Command "Write-AuditLogEvent" -ErrorAction SilentlyContinue)) {
                    Write-AuditLogEvent -OperationId $OperationId -Level "WARNING" -EventName "SCAN_QUERYTABLE_READ_FAILED" -File $file.FullName -Stage "QueryTables" -Message "$_"
                }
            }

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
                    if ($OperationId -and (Get-Command "Write-AuditLogEvent" -ErrorAction SilentlyContinue)) {
                        Write-AuditLogEvent -OperationId $OperationId -Level "WARNING" -EventName "SCAN_VBA_READ_FAILED" -File $file.FullName -Stage "VBA" -Message "$_"
                    }
                }
            }

            $wb.Close($false)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
            if ($OperationId -and (Get-Command "Write-AuditLogEvent" -ErrorAction SilentlyContinue)) {
                Write-AuditLogEvent -OperationId $OperationId -Level "INFO" -EventName "SCAN_FILE_COMPLETED" -File $file.FullName -Stage "Completed" -Message "Workbook scan completed" -Data @{ foundIpCount = @($fileDetail.foundIPs).Count; queryCount = @($fileDetail.queries).Count; connectionCount = @($fileDetail.connections).Count; vbaMatchCount = @($fileDetail.vbaMatches).Count }
            }
        } catch {
            $fileDetail.status = "Error: $_"
            if ($OperationId -and (Get-Command "Write-AuditLogEvent" -ErrorAction SilentlyContinue)) {
                Write-AuditLogEvent -OperationId $OperationId -Level "ERROR" -EventName "SCAN_FILE_FAILED" -File $file.FullName -Stage "File Scan" -Message "$_"
            }
        }

        $fileList += $fileDetail
        if ($OperationId -and (Get-Command "Append-OperationScannedFile" -ErrorAction SilentlyContinue)) {
            Append-OperationScannedFile -OperationId $OperationId -FileDetail $fileDetail -IpSummary $ipSummary
        }
    }

    Cleanup-ExcelCOM $excel
    Set-ProgressState $false "scan" $totalCount $totalCount "Tarama Tamamlandı!"

    if ($OperationId -and (Get-Command "Update-OperationProgress" -ErrorAction SilentlyContinue)) {
        $null = Update-OperationProgress -OperationId $OperationId -ProcessedFiles $processedCount -TotalFiles $totalCount -CurrentFile "Tamamlandı" -CurrentStage "Completed" -ProgressPercent 100
    }
    if ($OperationId -and (Get-Command "Write-AuditLogEvent" -ErrorAction SilentlyContinue)) {
        Write-AuditLogEvent -OperationId $OperationId -Level "INFO" -EventName "SCAN_COMPLETED" -Stage "Completed" -Message "Excel directory scan completed" -Data @{ directory = $DirectoryPath; processedFiles = $processedCount; totalFiles = $totalCount; detectedIpCount = $ipSummary.Count }
    }

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

function New-VerifiedBackup ([string]$SourcePath, [string]$BackupDir) {
    if (-not (Test-Path $SourcePath)) {
        return @{ Success = $false; Error = "Kaynak dosya bulunamadı: $SourcePath" }
    }

    try {
        if (-not (Test-Path $BackupDir)) {
            New-Item -ItemType Directory -Path $BackupDir -Force | Out-Null
        }

        $sourceFile = Get-Item $SourcePath
        $destPath = Join-Path $BackupDir $sourceFile.Name

        Copy-FileWithShare $SourcePath $destPath

        # Verification 1: File Exists
        if (-not (Test-Path $destPath)) {
            return @{ Success = $false; Error = "Yedek dosyası oluşturulamadı: $destPath" }
        }

        # Verification 2: Length
        $destFile = Get-Item $destPath
        if ($sourceFile.Length -ne $destFile.Length) {
            if (Test-Path $destPath) { Remove-Item $destPath -Force }
            return @{ Success = $false; Error = "Yedek dosya boyutu uyuşmuyor: Kaynak $($sourceFile.Length), Yedek $($destFile.Length)" }
        }

        # Verification 3: SHA256 cryptographic match
        $srcHash = Get-FileSha256 $SourcePath
        $dstHash = Get-FileSha256 $destPath
        if ($srcHash -ne $dstHash) {
            if (Test-Path $destPath) { Remove-Item $destPath -Force }
            return @{ Success = $false; Error = "Yedek kriptografik hash doğrulaması başarısız oldu!" }
        }

        return @{
            Success = $true
            BackupPath = $destPath
            Sha256 = $dstHash
            Length = $destFile.Length
        }
    } catch {
        return @{ Success = $false; Error = "Yedekleme hatası: $_" }
    }
}

function New-StagingFile ([string]$OriginalFilePath, [string]$OperationId = "") {
    if (-not (Test-Path $OriginalFilePath)) {
        return @{ Success = $false; Error = "Orijinal dosya bulunamadı: $OriginalFilePath" }
    }

    try {
        $opId = if (-not [string]::IsNullOrWhiteSpace($OperationId)) { $OperationId } else { New-OperationId }
        $file = Get-Item $OriginalFilePath
        $stagingName = "." + $file.BaseName + "." + $opId + ".staging" + $file.Extension
        $stagingPath = Join-Path $file.DirectoryName $stagingName

        if (Test-Path $stagingPath) {
            Remove-Item $stagingPath -Force -ErrorAction SilentlyContinue
        }

        # Safe byte copy to same directory
        [System.IO.File]::Copy($OriginalFilePath, $stagingPath, $true)

        # Verification 1: Exists
        if (-not (Test-Path $stagingPath)) {
            return @{ Success = $false; Error = "Staging kopyası oluşturulamadı: $stagingPath" }
        }

        # Verification 2: Length
        $stagingFile = Get-Item $stagingPath
        if ($file.Length -ne $stagingFile.Length) {
            Remove-Item $stagingPath -Force -ErrorAction SilentlyContinue
            return @{ Success = $false; Error = "Staging kopya boyutu uyuşmuyor: Orijinal $($file.Length), Staging $($stagingFile.Length)" }
        }

        # Verification 3: SHA256 integrity
        $origHash = Get-FileSha256 $OriginalFilePath
        $stageHash = Get-FileSha256 $stagingPath
        if ($origHash -ne $stageHash) {
            Remove-Item $stagingPath -Force -ErrorAction SilentlyContinue
            return @{ Success = $false; Error = "Staging kopyasının hash doğrulaması başarısız oldu!" }
        }

        # Set hidden attribute to avoid polluting user explorer view
        try {
            $stagingFile.Attributes = [System.IO.FileAttributes]::Hidden
        } catch { }

        return @{
            Success = $true
            OriginalPath = $OriginalFilePath
            StagingPath = $stagingPath
            InitialHash = $origHash
            OperationId = $opId
        }
    } catch {
        return @{ Success = $false; Error = "Staging oluşturma hatası: $_" }
    }
}

function Remove-StagingFile ([string]$StagingPath) {
    if (-not [string]::IsNullOrWhiteSpace($StagingPath) -and (Test-Path -LiteralPath $StagingPath)) {
        try {
            $f = Get-Item -LiteralPath $StagingPath -Force -ErrorAction SilentlyContinue
            if ($f) {
                $f.Attributes = [System.IO.FileAttributes]::Normal
            }
            Remove-Item -LiteralPath $StagingPath -Force -ErrorAction SilentlyContinue
        } catch { }
    }
}

function New-IsolatedExcelInstance () {
    # 1. Capture existing EXCEL processes before spawning (zero compiler dependency)
    $beforePids = [System.Collections.Generic.HashSet[int]]::new()
    Get-Process excel -ErrorAction SilentlyContinue | ForEach-Object { $beforePids.Add($_.Id) | Out-Null }
    $spawnTime = [System.DateTime]::UtcNow

    # 2. Instantiate isolated Excel COM Application
    $excel = New-Object -ComObject Excel.Application

    # 3. Deterministically detect newly spawned PID via process delta and disambiguation
    $excelPid = 0
    $isAmbiguous = $false
    try {
        $candidatePids = @()
        $nowProcesses = Get-Process excel -ErrorAction SilentlyContinue
        foreach ($p in $nowProcesses) {
            if (-not $beforePids.Contains($p.Id)) {
                $candidatePids += $p.Id
            }
        }

        if ($candidatePids.Count -eq 1) {
            $excelPid = $candidatePids[0]
        } else {
            # 0 or multiple candidates -> Ambiguous! Never kill an ambiguous PID.
            $excelPid = 0
            $isAmbiguous = $true
        }
    } catch {
        $excelPid = 0
        $isAmbiguous = $true
    }

    # 4. Capture original application properties to preserve them
    $origCalc = -4105 # xlCalculationAutomatic default
    try { $origCalc = $excel.Calculation } catch { }

    $origCalcBeforeSave = $true
    try { $origCalcBeforeSave = $excel.CalculateBeforeSave } catch { }

    $origSecurity = 1 # msoAutomationSecurityByUI default
    try { $origSecurity = $excel.AutomationSecurity } catch { }

    # 5. Apply strict automation isolation flags
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false
    try { $excel.AskToUpdateLinks = $false } catch { }
    try { $excel.AutomationSecurity = 3 } catch { } # 3 = msoAutomationSecurityForceDisable
    try { $excel.Calculation = -4135 } catch { }    # -4135 = xlCalculationManual during staging
    try { $excel.CalculateBeforeSave = $false } catch { }

    return @{
        App = $excel
        Excel = $excel
        Pid = $excelPid
        IsAmbiguousPid = $isAmbiguous
        OriginalCalculation = $origCalc
        OriginalCalculateBeforeSave = $origCalcBeforeSave
        OriginalAutomationSecurity = $origSecurity
    }
}

function Restore-ExcelIsolation ($excelContext) {
    if ($excelContext -and $excelContext.App) {
        $app = $excelContext.App
        try { $app.AutomationSecurity = $excelContext.OriginalAutomationSecurity } catch { }
        try { $app.Calculation = $excelContext.OriginalCalculation } catch { }
        try { $app.CalculateBeforeSave = $excelContext.OriginalCalculateBeforeSave } catch { }
        try { $app.EnableEvents = $true } catch { }
        try { $app.AskToUpdateLinks = $true } catch { }
    }
}

function Close-IsolatedExcelInstance ($excelContext) {
    # A) PRIMARY MECHANISM: Deterministic RCW cleanup and application quit
    if ($excelContext -and $excelContext.App) {
        $app = $excelContext.App
        try {
            Restore-ExcelIsolation $excelContext
            $wbs = $null
            try {
                $wbs = $app.Workbooks
                if ($wbs) {
                    for ($w = $wbs.Count; $w -ge 1; $w--) {
                        try {
                            $item = $wbs.Item($w)
                            $item.Close($false)
                            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($item) | Out-Null
                            $item = $null
                        } catch { }
                    }
                }
            } catch { }
            finally {
                if ($wbs) {
                    try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wbs) | Out-Null } catch { }
                    $wbs = $null
                }
            }
            $app.Quit()
        } catch { }
        try {
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null
        } catch { }
        $app = $null
        $excelContext.App = $null
        $excelContext.Excel = $null
    }

    [System.GC]::Collect(2, [System.GCCollectionMode]::Forced, $true)
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect(2, [System.GCCollectionMode]::Forced, $true)
    [System.GC]::WaitForPendingFinalizers()

    # B) Wait for graceful exit
    # Never kill if PID is 0 or ambiguous to protect user's open Excel documents
    if ($excelContext -and $excelContext.Pid -gt 0 -and (-not $excelContext.IsAmbiguousPid)) {
        $targetPid = $excelContext.Pid
        for ($i = 0; $i -lt 50; $i++) {
            $p = Get-Process -Id $targetPid -ErrorAction SilentlyContinue
            if (-not $p) { break }
            if ($p.HasExited) {
                $p.Dispose()
                # Wait until PID is completely purged from OS process table
                for ($v = 0; $v -lt 50; $v++) {
                    $check = Get-Process -Id $targetPid -ErrorAction SilentlyContinue
                    if (-not $check) { break }
                    $check.Dispose()
                    Start-Sleep -Milliseconds 100
                }
                break
            }
            $p.Dispose()
            Start-Sleep -Milliseconds 100
        }

        # Final verification: ensure PID is null
        for ($v = 0; $v -lt 30; $v++) {
            $check = Get-Process -Id $targetPid -ErrorAction SilentlyContinue
            if (-not $check) { break }
            $check.Dispose()
            Start-Sleep -Milliseconds 100
        }
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

function Get-WorkbookSnapshot ($Workbook, [bool]$HasVba = $false) {
    if ($null -eq $Workbook) { return $null }

    $sheetNames = @()
    try {
        $sheets = $Workbook.Worksheets
        if ($sheets) {
            $sheetCount = $sheets.Count
            for ($i = 1; $i -le $sheetCount; $i++) {
                $s = $sheets.Item($i)
                if ($s) {
                    $sheetNames += $s.Name
                    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($s) | Out-Null
                }
            }
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheets) | Out-Null
        }
    } catch { }

    $queryNames = @()
    try {
        $queries = $Workbook.Queries
        if ($queries) {
            $queryCount = $queries.Count
            for ($i = 1; $i -le $queryCount; $i++) {
                $q = $queries.Item($i)
                if ($q) {
                    $queryNames += $q.Name
                    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($q) | Out-Null
                }
            }
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($queries) | Out-Null
        }
    } catch { }

    $connNames = @()
    try {
        $conns = $Workbook.Connections
        if ($conns) {
            $connCount = $conns.Count
            for ($i = 1; $i -le $connCount; $i++) {
                $c = $conns.Item($i)
                if ($c) {
                    $connNames += $c.Name
                    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($c) | Out-Null
                }
            }
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($conns) | Out-Null
        }
    } catch { }

    $vbaModules = @()
    $isVbaSigned = $false
    $vbaAccessible = $true
    if ($HasVba) {
        try {
            $vbProj = $Workbook.VBProject
            if ($vbProj) {
                $comps = $vbProj.VBComponents
                if ($comps) {
                    $compCount = $comps.Count
                    for ($i = 1; $i -le $compCount; $i++) {
                        $comp = $comps.Item($i)
                        if ($comp) {
                            $vbaModules += $comp.Name
                            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($comp) | Out-Null
                        }
                    }
                    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($comps) | Out-Null
                }
                try {
                    if ($vbProj.Signature) {
                        $isVbaSigned = $true
                    }
                } catch { }
                [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($vbProj) | Out-Null
            }
        } catch {
            $vbaAccessible = $false
        }
    }

    $format = 0
    try { $format = $Workbook.FileFormat } catch { }

    $date1904 = $false
    try { $date1904 = $Workbook.Date1904 } catch { }

    $precision = $false
    try { $precision = $Workbook.PrecisionAsDisplayed } catch { }

    return @{
        FileFormat = $format
        WorksheetsCount = $sheetNames.Count
        WorksheetNames = $sheetNames
        QueriesCount = $queryNames.Count
        QueryNames = $queryNames
        ConnectionsCount = $connNames.Count
        ConnectionNames = $connNames
        HasVba = $HasVba
        VbaAccessible = $vbaAccessible
        VbaModulesCount = if ($vbaAccessible) { $vbaModules.Count } else { -1 }
        VbaModuleNames = $vbaModules
        IsVbaSigned = $isVbaSigned
        Date1904 = $date1904
        PrecisionAsDisplayed = $precision
    }
}

function Compare-WorkbookSnapshot ($PreSnapshot, $PostSnapshot) {
    if ($null -eq $PreSnapshot -or $null -eq $PostSnapshot) {
        return @{
            IsValid = $false
            ErrorCode = "NULL_SNAPSHOT"
            Message = "Karşılaştırma için geçerli snapshot bulunamadı."
            Differences = @("Snapshot verisi boş")
        }
    }

    $diffs = @()

    # 1. FileFormat
    if ($PreSnapshot.FileFormat -ne $PostSnapshot.FileFormat) {
        $diffs += "Dosya formatı değişti: Eski $($PreSnapshot.FileFormat), Yeni $($PostSnapshot.FileFormat)"
        return @{
            IsValid = $false
            ErrorCode = "VALIDATION_FAILED_FILE_FORMAT"
            Message = "Dosya format dönüşümüne uğradı!"
            Differences = $diffs
        }
    }

    # 2. Worksheet count and names
    if ($PreSnapshot.WorksheetsCount -ne $PostSnapshot.WorksheetsCount) {
        $diffs += "Çalışma sayfası sayısı değişti: Eski $($PreSnapshot.WorksheetsCount), Yeni $($PostSnapshot.WorksheetsCount)"
    } else {
        for ($i = 0; $i -lt $PreSnapshot.WorksheetsCount; $i++) {
            if ($PreSnapshot.WorksheetNames[$i] -ne $PostSnapshot.WorksheetNames[$i]) {
                $diffs += "Sayfa adı değişti: '$($PreSnapshot.WorksheetNames[$i])' -> '$($PostSnapshot.WorksheetNames[$i])'"
            }
        }
    }
    if ($diffs.Count -gt 0) {
        return @{
            IsValid = $false
            ErrorCode = "VALIDATION_FAILED_WORKSHEET_IDENTITY"
            Message = "Çalışma sayfalarında beklenmeyen yapısal değişim tespit edildi."
            Differences = $diffs
        }
    }

    # 3. Query count and names
    if ($PreSnapshot.QueriesCount -ne $PostSnapshot.QueriesCount) {
        $diffs += "Power Query sayısı değişti: Eski $($PreSnapshot.QueriesCount), Yeni $($PostSnapshot.QueriesCount)"
    } else {
        for ($i = 0; $i -lt $PreSnapshot.QueriesCount; $i++) {
            if ($PreSnapshot.QueryNames[$i] -ne $PostSnapshot.QueryNames[$i]) {
                $diffs += "Sorgu adı değişti: '$($PreSnapshot.QueryNames[$i])' -> '$($PostSnapshot.QueryNames[$i])'"
            }
        }
    }
    if ($diffs.Count -gt 0) {
        return @{
            IsValid = $false
            ErrorCode = "VALIDATION_FAILED_QUERY_IDENTITY"
            Message = "Power Query kimliklerinde beklenmeyen yapısal değişim tespit edildi."
            Differences = $diffs
        }
    }

    # 4. Connection count and names (MUST BE IMMUTABLE)
    if ($PreSnapshot.ConnectionsCount -ne $PostSnapshot.ConnectionsCount) {
        $diffs += "Bağlantı sayısı değişti: Eski $($PreSnapshot.ConnectionsCount), Yeni $($PostSnapshot.ConnectionsCount)"
    }
    if ($diffs.Count -gt 0) {
        return @{
            IsValid = $false
            ErrorCode = "VALIDATION_FAILED_CONNECTION_IDENTITY"
            Message = "Veri bağlantı kimliği/adı değiştirilemez! Orijinal bağlantı adları korunmalıdır."
            Differences = $diffs
        }
    }

    # 5. VBA structure
    if ($PreSnapshot.HasVba -and $PreSnapshot.VbaAccessible -and $PostSnapshot.VbaAccessible) {
        if ($PreSnapshot.VbaModulesCount -ne $PostSnapshot.VbaModulesCount) {
            $diffs += "VBA modül sayısı değişti: Eski $($PreSnapshot.VbaModulesCount), Yeni $($PostSnapshot.VbaModulesCount)"
            return @{
                IsValid = $false
                ErrorCode = "VALIDATION_FAILED_VBA_STRUCTURE"
                Message = "VBA modül yapısında bozulma tespit edildi."
                Differences = $diffs
            }
        }
    }

    return @{
        IsValid = $true
        ErrorCode = "NONE"
        Message = "Semantik snapshot doğrulaması başarılı."
        Differences = @()
    }
}

function Create-ExcelBackup ($DirectoryPath, $OperationId = "") {
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    $files = Get-ExcelFiles $DirectoryPath
    if ($files.Count -eq 0) {
        return @{ success = $false; error = "Yedeklenecek Excel dosyası bulunamadı." }
    }

    $opId = if (-not [string]::IsNullOrWhiteSpace($OperationId)) { $OperationId } else { New-OperationId }
    if ($OperationId -and (Get-Command "Update-OperationProgress" -ErrorAction SilentlyContinue)) {
        $null = Update-OperationProgress -OperationId $OperationId -ProcessedFiles 0 -TotalFiles $files.Count -CurrentFile "Yedek hazırlanıyor" -CurrentStage "Backup" -ProgressPercent 3
    }
    if ($OperationId -and (Get-Command "Write-AuditLogEvent" -ErrorAction SilentlyContinue)) {
        Write-AuditLogEvent -OperationId $OperationId -Level "INFO" -EventName "BACKUP_STARTED" -Stage "Backup" -Message "Backup preparation started" -Data @{ directory = $DirectoryPath; totalFiles = $files.Count }
    }

    $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    
    # Store backup adjacent to selected directory: <DirectoryPath>_ExcelUpdater_Backups_<opId>
    $parentPath = Split-Path -Parent $DirectoryPath
    $folderName = Split-Path -Leaf $DirectoryPath
    $backupDirName = "${folderName}_ExcelUpdater_Backups_${opId}"
    $backupDir = if ($parentPath) { Join-Path $parentPath $backupDirName } else { Join-Path $DirectoryPath $backupDirName }

    try {
        if (-not (Test-Path $backupDir)) {
            New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
        }
        
        # Test write permission on backup dir
        $testProbe = Join-Path $backupDir ".probe_$([Guid]::NewGuid().ToString('N')).tmp"
        try {
            [System.IO.File]::WriteAllText($testProbe, "probe")
            Remove-Item $testProbe -Force
        } catch {
            return @{ success = $false; error = "Yedekleme klasörüne yazma izni bulunmuyor: $backupDir" }
        }

        $backedUpCount = 0
        $verifiedBackups = @()

        foreach ($file in $files) {
            if ($OperationId -and (Get-Command "Update-OperationProgress" -ErrorAction SilentlyContinue)) {
                $backupPct = [math]::Min(9, [math]::Max(3, [math]::Floor(3 + (($backedUpCount / [double]$files.Count) * 6))))
                $null = Update-OperationProgress -OperationId $OperationId -ProcessedFiles $backedUpCount -TotalFiles $files.Count -CurrentFile "Yedekleniyor: $($file.Name)" -CurrentStage "Backup" -ProgressPercent $backupPct
            }
            $bRes = New-VerifiedBackup -SourcePath $file.FullName -BackupDir $backupDir
            if (-not $bRes.Success) {
                # Clean up corrupted/partial backup directory
                Remove-Item -Path $backupDir -Recurse -Force -ErrorAction SilentlyContinue
                return @{ success = $false; error = "Dosya yedeklenemedi ($($file.Name)): $($bRes.Error)" }
            }
            $verifiedBackups += @{
                OriginalPath = $file.FullName
                BackupPath = $bRes.BackupPath
                Sha256 = $bRes.Sha256
            }
            $backedUpCount++
        }

        if ($OperationId -and (Get-Command "Update-OperationProgress" -ErrorAction SilentlyContinue)) {
            $null = Update-OperationProgress -OperationId $OperationId -ProcessedFiles $files.Count -TotalFiles $files.Count -CurrentFile "Yedekleme tamamlandı" -CurrentStage "Backup Completed" -ProgressPercent 9
        }
        if ($OperationId -and (Get-Command "Write-AuditLogEvent" -ErrorAction SilentlyContinue)) {
            Write-AuditLogEvent -OperationId $OperationId -Level "INFO" -EventName "BACKUP_COMPLETED" -Stage "Backup" -Message "Backup preparation completed" -Data @{ directory = $DirectoryPath; totalFiles = $backedUpCount; backupDirectory = $backupDir }
        }

        # Write cryptographic backup manifest for integrity verification on restore
        $manifest = @{
            operationId = $opId
            timestamp = $timestamp
            directory = $DirectoryPath
            totalFiles = $backedUpCount
            files = $verifiedBackups
        }
        $manifestPath = Join-Path $backupDir "backup_manifest.json"
        try {
            $manifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
        } catch { }

        return @{
            success = $true
            backupDirectory = $backupDir
            totalFiles = $backedUpCount
            operationId = $opId
            timestamp = $timestamp
            backups = $verifiedBackups
        }
    } catch {
        return @{ success = $false; error = "Backup failed: $_" }
    }
}

function Commit-StagingWorkbook {
    param(
        [Parameter(Mandatory=$true)][string]$StagingPath,
        [Parameter(Mandatory=$true)][string]$OriginalPath,
        [Parameter(Mandatory=$true)][string]$InitialHash,
        [string]$OperationId = ""
    )

    if (-not (Test-Path -LiteralPath $StagingPath)) {
        return @{
            Success = $false
            ErrorCode = "STAGING_MISSING"
            Error = "Staging file does not exist: $StagingPath"
        }
    }

    $stagingItem = Get-Item -LiteralPath $StagingPath -Force
    if ($stagingItem.Length -le 0) {
        return @{
            Success = $false
            ErrorCode = "STAGING_EMPTY"
            Error = "Staging file is 0 bytes: $StagingPath"
        }
    }

    if (-not (Test-Path -LiteralPath $OriginalPath)) {
        return @{
            Success = $false
            ErrorCode = "ORIGINAL_MISSING"
            Error = "Original target file does not exist: $OriginalPath"
        }
    }

    $opTag = if ([string]::IsNullOrWhiteSpace($OperationId)) { [System.Guid]::NewGuid().ToString("N").Substring(0, 8) } else { $OperationId }
    $replaceBak = "$OriginalPath.$opTag.replace_bak"
    $renameBak = "$OriginalPath.$opTag.rename_bak"

    # 1. TOCTOU Concurrency Guard: verify original has not been modified since initial read
    $currentHash = Get-FileSha256 $OriginalPath
    if ($currentHash -ne $InitialHash) {
        # Check if rename_bak already exists and holds the initial file (double-fault / interrupted Phase 1 recovery)
        if ((Test-Path -LiteralPath $renameBak) -and ((Get-FileSha256 $renameBak) -eq $InitialHash)) {
            # Staging -> Original attempt
            $stagingMoveErr = $null
            try {
                Move-Item -LiteralPath $StagingPath -Destination $OriginalPath -Force -ErrorAction Stop
                $replaceBak = $renameBak
            } catch {
                $stagingMoveErr = $_
                # Staging move failed -> rollback original from renameBak!
                $rollbackSuccess = $false
                try {
                    Move-Item -LiteralPath $renameBak -Destination $OriginalPath -Force -ErrorAction Stop
                    $rollbackSuccess = (Test-Path -LiteralPath $OriginalPath)
                } catch {
                    $rollbackSuccess = $false
                }

                if (-not $rollbackSuccess) {
                    # DOUBLE FAULT!
                    return @{
                        Success = $false
                        ErrorCode = "CRITICAL_MANUAL_RECOVERY_REQUIRED"
                        Error = "CRITICAL: Two-phase commit double fault! Staging replacement failed ($stagingMoveErr) and rollback of original file from '$renameBak' to '$OriginalPath' also failed! Original file content remains at '$renameBak'. DO NOT DELETE."
                        recoveryFilePath = $renameBak
                        originalExpectedPath = $OriginalPath
                        stagingPath = $StagingPath
                        operationId = $opTag
                    }
                }

                return @{
                    Success = $false
                    ErrorCode = "COMMIT_RENAME_FAILED"
                    Error = "Two-phase rename fallback failed: $stagingMoveErr. Original file successfully rolled back from rename backup."
                }
            }
        } else {
            Remove-Item -LiteralPath $StagingPath -Force -ErrorAction SilentlyContinue
            return @{
                Success = $false
                ErrorCode = "CONCURRENT_MODIFICATION_DETECTED"
                Error = "Original file was modified by another process/user while update was in progress. Operation aborted for safety."
            }
        }
    }

    # 2. Capture staging hash and size
    $stagingHash = Get-FileSha256 $StagingPath
    $stagingLength = $stagingItem.Length

    # Clear hidden/system attributes on staging file so target file doesn't remain hidden
    try {
        [System.IO.File]::SetAttributes($StagingPath, [System.IO.FileAttributes]::Normal)
    } catch { }

    $methodUsed = "None"

    # 3. Primary: Try [System.IO.File]::Replace
    $primarySuccess = $false
    try {
        if (Test-Path -LiteralPath $replaceBak) {
            Remove-Item -LiteralPath $replaceBak -Force -ErrorAction SilentlyContinue
        }

        [System.IO.File]::Replace($StagingPath, $OriginalPath, $replaceBak, $true)
        $methodUsed = "File.Replace"
        $primarySuccess = $true
    } catch {
        # Fallback to Two-Phase Rename
        $methodUsed = "TwoPhaseRename"
        try {
            if (Test-Path -LiteralPath $renameBak) {
                Remove-Item -LiteralPath $renameBak -Force -ErrorAction SilentlyContinue
            }

            Move-Item -LiteralPath $OriginalPath -Destination $renameBak -Force -ErrorAction Stop
            try {
                Move-Item -LiteralPath $StagingPath -Destination $OriginalPath -Force -ErrorAction Stop
                $replaceBak = $renameBak
            } catch {
                $stagingMoveErr = $_
                # Staging move failed -> rollback original from renameBak!
                $rollbackSuccess = $false
                try {
                    if (Test-Path -LiteralPath $renameBak) {
                        Move-Item -LiteralPath $renameBak -Destination $OriginalPath -Force -ErrorAction Stop
                        $rollbackSuccess = (Test-Path -LiteralPath $OriginalPath)
                    }
                } catch {
                    $rollbackSuccess = $false
                }

                if (-not $rollbackSuccess) {
                    # DOUBLE FAULT!
                    return @{
                        Success = $false
                        ErrorCode = "CRITICAL_MANUAL_RECOVERY_REQUIRED"
                        Error = "CRITICAL: Two-phase commit double fault! Staging move failed ($stagingMoveErr) and rollback of original file from '$renameBak' to '$OriginalPath' also failed! Original file content remains at '$renameBak'. DO NOT DELETE."
                        recoveryFilePath = $renameBak
                        originalExpectedPath = $OriginalPath
                        stagingPath = $StagingPath
                        operationId = $opTag
                    }
                }

                return @{
                    Success = $false
                    ErrorCode = "COMMIT_RENAME_FAILED"
                    Error = "Two-phase rename fallback failed: $stagingMoveErr. Original file successfully rolled back from rename backup."
                }
            }
        } catch {
            return @{
                Success = $false
                ErrorCode = "COMMIT_BACKUP_FAILED"
                Error = "Failed to initiate rename fallback: $_"
            }
        }
    }

    # 4. Mandatory Post-Commit Verification
    if (-not (Test-Path -LiteralPath $OriginalPath)) {
        return @{
            Success = $false
            ErrorCode = "POST_COMMIT_MISSING"
            Error = "Target file missing after commit attempt!"
            recoveryFilePath = (if (Test-Path -LiteralPath $replaceBak) { $replaceBak } else { $null })
            originalExpectedPath = $OriginalPath
            stagingPath = $StagingPath
            operationId = $opTag
        }
    }

    $finalItem = Get-Item -LiteralPath $OriginalPath -Force
    if ($finalItem.Length -le 0 -or $finalItem.Length -ne $stagingLength) {
        return @{
            Success = $false
            ErrorCode = "POST_COMMIT_SIZE_MISMATCH"
            Error = "Target file length mismatch: expected $stagingLength bytes, found $($finalItem.Length) bytes."
            recoveryFilePath = (if (Test-Path -LiteralPath $replaceBak) { $replaceBak } else { $null })
            originalExpectedPath = $OriginalPath
            stagingPath = $StagingPath
            operationId = $opTag
        }
    }

    $finalHash = Get-FileSha256 $OriginalPath
    if ($finalHash -ne $stagingHash) {
        return @{
            Success = $false
            ErrorCode = "POST_COMMIT_HASH_MISMATCH"
            Error = "Target file hash mismatch after commit: expected $stagingHash, found $finalHash."
            recoveryFilePath = (if (Test-Path -LiteralPath $replaceBak) { $replaceBak } else { $null })
            originalExpectedPath = $OriginalPath
            stagingPath = $StagingPath
            operationId = $opTag
        }
    }

    # Post-commit verification passed: clean up temporary backup file
    if (Test-Path -LiteralPath $replaceBak) {
        Remove-Item -LiteralPath $replaceBak -Force -ErrorAction SilentlyContinue
    }

    return @{
        Success = $true
        ErrorCode = $null
        Method = $methodUsed
        FinalHash = $finalHash
        Length = $stagingLength
    }
}

function Restore-BatchBackup {
    param(
        [Parameter(Mandatory=$true)][string]$BackupDir,
        [Parameter(Mandatory=$true)][string]$TargetDir
    )

    if (-not (Test-Path -LiteralPath $BackupDir)) {
        return @{
            Success = $false
            Error = "Backup directory not found: $BackupDir"
            failedRestoreFiles = @()
            RestoredCount = 0
            Errors = @("Backup directory not found: $BackupDir")
        }
    }

    $manifestPath = Join-Path $BackupDir "backup_manifest.json"
    $manifestMap = @{}
    if (Test-Path -LiteralPath $manifestPath) {
        try {
            $manifestData = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($manifestData -and $manifestData.files) {
                foreach ($f in $manifestData.files) {
                    $leafName = Split-Path -Leaf $f.OriginalPath
                    $manifestMap[$leafName] = $f.Sha256
                }
            }
        } catch { }
    }

    $backupFiles = Get-ChildItem -LiteralPath $BackupDir -File | Where-Object { $_.Name -ne "backup_manifest.json" }
    $restoredCount = 0
    $failedRestoreFiles = @()
    $errors = @()

    foreach ($bf in $backupFiles) {
        $targetFile = Join-Path $TargetDir $bf.Name
        $expectedHash = if ($manifestMap.ContainsKey($bf.Name)) { $manifestMap[$bf.Name] } else { Get-FileSha256 $bf.FullName }
        $bfCurrentHash = Get-FileSha256 $bf.FullName

        # Pre-copy integrity check: verify backup file itself is not corrupted
        if ($manifestMap.ContainsKey($bf.Name) -and ($bfCurrentHash -ne $expectedHash)) {
            $failed = @{
                OriginalPath = $targetFile
                BackupPath = $bf.FullName
                ExpectedHash = $expectedHash
                CurrentHash = $bfCurrentHash
                Error = "Backup file itself is corrupted (checksum mismatch with backup manifest)"
            }
            $failedRestoreFiles += $failed
            $errors += "Backup file corrupted: $($bf.Name)"
            continue
        }

        try {
            Copy-Item -LiteralPath $bf.FullName -Destination $targetFile -Force
            if (-not (Test-Path -LiteralPath $targetFile)) {
                $failed = @{
                    OriginalPath = $targetFile
                    BackupPath = $bf.FullName
                    ExpectedHash = $expectedHash
                    CurrentHash = $null
                    Error = "Target file does not exist after restore copy"
                }
                $failedRestoreFiles += $failed
                $errors += "File missing after copy: $($bf.Name)"
                continue
            }

            $tfItem = Get-Item -LiteralPath $targetFile -Force
            if ($tfItem.Length -ne $bf.Length -or $tfItem.Length -le 0) {
                $failed = @{
                    OriginalPath = $targetFile
                    BackupPath = $bf.FullName
                    ExpectedHash = $expectedHash
                    CurrentHash = (Get-FileSha256 $targetFile)
                    Error = "Length mismatch after copy: expected $($bf.Length), got $($tfItem.Length)"
                }
                $failedRestoreFiles += $failed
                $errors += "Length mismatch after restoring $($bf.Name)"
                continue
            }

            $tfHash = Get-FileSha256 $targetFile
            if ($expectedHash -ne $tfHash) {
                $failed = @{
                    OriginalPath = $targetFile
                    BackupPath = $bf.FullName
                    ExpectedHash = $expectedHash
                    CurrentHash = $tfHash
                    Error = "Checksum mismatch: expected $expectedHash, got $tfHash"
                }
                $failedRestoreFiles += $failed
                $errors += "Checksum mismatch after restoring $($bf.Name)"
            } else {
                $restoredCount++
            }
        } catch {
            $failed = @{
                OriginalPath = $targetFile
                BackupPath = $bf.FullName
                ExpectedHash = $expectedHash
                CurrentHash = (Get-FileSha256 $targetFile)
                Error = "$_"
            }
            $failedRestoreFiles += $failed
            $errors += "Failed to restore $($bf.Name): $_"
        }
    }

    $isSuccess = ($failedRestoreFiles.Count -eq 0 -and $errors.Count -eq 0)

    return @{
        Success = $isSuccess
        RestoredCount = $restoredCount
        failedRestoreFiles = $failedRestoreFiles
        Errors = $errors
        Error = if (-not $isSuccess) { $errors -join "; " } else { $null }
    }
}

function Update-ExcelDirectory ($DirectoryPath, $Rules, $Options = @{}, [string]$OperationId = "") {
    $opId = if (-not [string]::IsNullOrWhiteSpace($OperationId)) { $OperationId } else { New-OperationId }
    if ($opId -and (Get-Command "Update-OperationProgress" -ErrorAction SilentlyContinue)) {
        $null = Update-OperationProgress -OperationId $opId -ProcessedFiles 0 -TotalFiles 0 -CurrentFile "Dosyalar listeleniyor" -CurrentStage "Preflight" -ProgressPercent 1
    }
    if ($opId -and (Get-Command "Write-AuditLogEvent" -ErrorAction SilentlyContinue)) {
        Write-AuditLogEvent -OperationId $opId -Level "INFO" -EventName "UPDATE_PREFLIGHT_STARTED" -Stage "Preflight" -Message "Update preflight started" -Data @{ directory = $DirectoryPath }
    }

    $files = Get-ExcelFiles $DirectoryPath
    $updateLog = @()
    $updatedFilesCount = 0
    $totalReplacements = 0
    $totalCount = $files.Count

    Set-ProgressState $true "update" 0 $totalCount "Güncellemeye Başlanıyor..."
    if ($opId -and (Get-Command "Update-OperationProgress" -ErrorAction SilentlyContinue)) {
        $null = Update-OperationProgress -OperationId $opId -ProcessedFiles 0 -TotalFiles $totalCount -CurrentFile "Ön kontrol hazırlanıyor" -CurrentStage "Preflight" -ProgressPercent 1
    }
    if ($opId -and (Get-Command "Write-AuditLogEvent" -ErrorAction SilentlyContinue)) {
        Write-AuditLogEvent -OperationId $opId -Level "INFO" -EventName "UPDATE_FILES_ENUMERATED" -Stage "Preflight" -Message "Excel files enumerated for update" -Data @{ directory = $DirectoryPath; totalFiles = $totalCount }
    }

    if ($totalCount -eq 0) {
        Set-ProgressState $false "update" 0 0 ""
        return @{ success = $false; error = "No Excel files found to update." }
    }

    # Validate rules
    $ruleValidation = Test-ReplacementRulesValid $Rules
    if (-not $ruleValidation.IsValid) {
        Set-ProgressState $false "update" 0 0 ""
        return @{ success = $false; error = "Rule validation failed: $($ruleValidation.Errors -join '; ')" }
    }
    $validRules = $ruleValidation.Rules
    if ($opId -and (Get-Command "Update-OperationProgress" -ErrorAction SilentlyContinue)) {
        $null = Update-OperationProgress -OperationId $opId -ProcessedFiles 0 -TotalFiles $totalCount -CurrentFile "Kurallar doğrulandı" -CurrentStage "Preflight" -ProgressPercent 2
    }

    # Pre-flight check: ensure all files are writable before doing anything
    $preflightCount = 0
    foreach ($f in $files) {
        if ($opId -and (Get-Command "Update-OperationProgress" -ErrorAction SilentlyContinue)) {
            $preflightPct = [math]::Min(3, [math]::Max(2, [math]::Floor(2 + (($preflightCount / [double]$totalCount) * 1))))
            $null = Update-OperationProgress -OperationId $opId -ProcessedFiles $preflightCount -TotalFiles $totalCount -CurrentFile "Kilit kontrolü: $($f.Name)" -CurrentStage "Preflight" -ProgressPercent $preflightPct
        }
        if (-not (Test-FileWritable $f.FullName)) {
            Set-ProgressState $false "update" 0 0 ""
            return @{
                success = $false
                error = "File is locked or read-only: $($f.Name). Update aborted before making any changes."
            }
        }
        $preflightCount++
    }
    if ($opId -and (Get-Command "Update-OperationProgress" -ErrorAction SilentlyContinue)) {
        $null = Update-OperationProgress -OperationId $opId -ProcessedFiles 0 -TotalFiles $totalCount -CurrentFile "Ön kontrol tamamlandı" -CurrentStage "Preflight Completed" -ProgressPercent 3
    }
    if ($opId -and (Get-Command "Write-AuditLogEvent" -ErrorAction SilentlyContinue)) {
        Write-AuditLogEvent -OperationId $opId -Level "INFO" -EventName "UPDATE_PREFLIGHT_COMPLETED" -Stage "Preflight" -Message "Update preflight completed" -Data @{ totalFiles = $totalCount }
    }

    # Cancellation Checkpoint 1 (After Preflight)
    if ($opId -and (Get-Command "Test-OperationCancellationRequested" -ErrorAction SilentlyContinue) -and (Test-OperationCancellationRequested $opId)) {
        if (Get-Command "Set-OperationStatus" -ErrorAction SilentlyContinue) { $null = Set-OperationStatus $opId "CANCELLED" "Cancelled by user" }
        return @{
            success = $false
            error = "İşlem kullanıcı tarafından iptal edildi."
            batchStatus = "CANCELLED"
            directory = $DirectoryPath
            totalFilesProcessed = 0
            updatedFilesCount = 0
            totalReplacements = 0
            logs = @()
            backupDirectory = $null
        }
    }

    # Backup handling:
    $backupDir = $null
    $autoBackupRequested = if ($Options -and $Options.autoBackup -ne $null) { $Options.autoBackup } else { $true }
    $atomicBatch = if ($Options -and $Options.atomicBatch -ne $null) { $Options.atomicBatch } else { $true }

    if ($autoBackupRequested -or $atomicBatch) {
        $backupRes = Create-ExcelBackup -DirectoryPath $DirectoryPath -OperationId $opId
        if (-not $backupRes.success) {
            Set-ProgressState $false "update" 0 0 ""
            return @{ success = $false; error = "Auto-backup failed prior to update: $($backupRes.error)" }
        }
        $backupDir = $backupRes.backupDirectory
    }

    # Cancellation Checkpoint 2 (After Backup)
    if ($opId -and (Get-Command "Test-OperationCancellationRequested" -ErrorAction SilentlyContinue) -and (Test-OperationCancellationRequested $opId)) {
        if (Get-Command "Set-OperationStatus" -ErrorAction SilentlyContinue) { $null = Set-OperationStatus $opId "CANCELLED" "Cancelled by user" }
        return @{
            success = $false
            error = "İşlem kullanıcı tarafından iptal edildi."
            batchStatus = "CANCELLED"
            directory = $DirectoryPath
            totalFilesProcessed = 0
            updatedFilesCount = 0
            totalReplacements = 0
            logs = @()
            backupDirectory = $backupDir
        }
    }

    $updatePowerQueries = if ($Options -and $Options.updateQueries -ne $null) { $Options.updateQueries } else { $true }
    $updateConnections  = if ($Options -and $Options.updateConnections -ne $null) { $Options.updateConnections } else { $true }
    $updateVba          = if ($Options -and $Options.updateVba -ne $null) { $Options.updateVba } else { $true }

    # Initialize hardened isolated Excel instance
    $iso = $null
    try {
        $iso = New-IsolatedExcelInstance
    } catch {
        Set-ProgressState $false "update" 0 0 ""
        return @{ success = $false; error = "Could not initialize isolated Excel instance: $_" }
    }

    $excel = $iso.Excel
    $processedCount = 0
    $committedFiles = @()
    $hasFatalFileError = $false

    try {
        $cancelledDuringBatch = $false
        foreach ($file in $files) {
            # Cancellation Checkpoint 3 (Before staging creation)
            if ($opId -and (Get-Command "Test-OperationCancellationRequested" -ErrorAction SilentlyContinue) -and (Test-OperationCancellationRequested $opId)) {
                $cancelledDuringBatch = $true
                break
            }

            $processedCount++
            Set-ProgressState $true "update" $processedCount $totalCount $file.Name
            if ($opId -and (Get-Command "Calculate-WeightedProgress" -ErrorAction SilentlyContinue)) {
                $pct = Calculate-WeightedProgress -TotalFiles $totalCount -CompletedFiles ($processedCount - 1) -CurrentFileStage "Excel Update"
                if (Get-Command "Update-OperationProgress" -ErrorAction SilentlyContinue) {
                    $null = Update-OperationProgress -OperationId $opId -ProcessedFiles $processedCount -TotalFiles $totalCount -CurrentFile $file.Name -CurrentStage "Excel Update" -ProgressPercent $pct
                }
            }

            $fileLog = @{
                fileName = $file.Name
                filePath = $file.FullName
                status = "Skipped"
                changesMade = 0
                details = @()
                initialHash = ""
                finalHash = ""
            }

            $stagingPath = $null
            $wb = $null

            try {
                # 1. Capture initial hash of original file
                $initialHash = Get-FileSha256 $file.FullName
                $fileLog.initialHash = $initialHash
                $fileLog.finalHash = $initialHash

                # 2. Create staging copy on same volume
                $stagingRes = New-StagingFile -OriginalFilePath $file.FullName -OperationId $opId
                if (-not $stagingRes.Success) {
                    throw "Failed to create staging file: $($stagingRes.Error)"
                }
                $stagingPath = $stagingRes.StagingPath

                # 3. Open staging file in isolated Excel instance (read-write)
                $wb = $excel.Workbooks.Open($stagingPath, 0, $false)
                $fileModified = $false

                # 4. Capture pre-update semantic snapshot
                $preSnapshot = Get-WorkbookSnapshot -Workbook $wb

                # 5. Apply Power Queries replacement
                if ($updatePowerQueries) {
                    try {
                        if ($wb.Queries) {
                            foreach ($q in $wb.Queries) {
                                $formula = $q.Formula
                                if ($formula) {
                                    $rep = Invoke-SafeReplacement -InputText $formula -Rules $validRules
                                    if ($rep.MatchCount -gt 0) {
                                        $q.Formula = $rep.ResultText
                                        $fileModified = $true
                                        $fileLog.changesMade += $rep.MatchCount
                                        $totalReplacements += $rep.MatchCount
                                        $fileLog.details += "PowerQuery '$($q.Name)': replaced $($rep.MatchCount) occurrence(s)"
                                    }
                                }
                            }
                        }
                    } catch {
                        $fileLog.details += "PowerQuery warning: $_"
                    }
                }

                # 6. Apply Data Connections replacement
                if ($updateConnections) {
                    try {
                        if ($wb.Connections) {
                            foreach ($conn in $wb.Connections) {
                                try {
                                    $connNameRep = Invoke-SafeReplacement -InputText "$($conn.Name)" -Rules $validRules
                                    if ($connNameRep.MatchCount -gt 0) {
                                        $conn.Name = $connNameRep.ResultText
                                        $fileModified = $true
                                        $fileLog.changesMade += $connNameRep.MatchCount
                                        $totalReplacements += $connNameRep.MatchCount
                                        $fileLog.details += "Connection name: replaced $($connNameRep.MatchCount) occurrence(s)"
                                    }
                                } catch {
                                    $fileLog.details += "Connection name warning: $_"
                                }

                                try {
                                    $connDesc = "$($conn.Description)"
                                    if (-not [string]::IsNullOrWhiteSpace($connDesc)) {
                                        $descRep = Invoke-SafeReplacement -InputText $connDesc -Rules $validRules
                                        if ($descRep.MatchCount -gt 0) {
                                            $conn.Description = $descRep.ResultText
                                            $fileModified = $true
                                            $fileLog.changesMade += $descRep.MatchCount
                                            $totalReplacements += $descRep.MatchCount
                                            $fileLog.details += "Connection description in '$($conn.Name)': replaced $($descRep.MatchCount) occurrence(s)"
                                        }
                                    }
                                } catch {
                                    $fileLog.details += "Connection description warning: $_"
                                }

                                if ($conn.OLEDBConnection) {
                                    try {
                                        $ole = $conn.OLEDBConnection
                                        $connText = Get-ExcelConnectionText -ConnectionObject $ole
                                        $cStr = $connText.Value
                                        if ($cStr) {
                                            $rep = Invoke-SafeReplacement -InputText $cStr -Rules $validRules
                                            if ($rep.MatchCount -gt 0) {
                                                $setRes = Set-ExcelConnectionText -ConnectionObject $ole -ConnectionText $rep.ResultText -PreferredProperty $connText.PropertyName
                                                if ($setRes.Success) {
                                                    $fileModified = $true
                                                    $fileLog.changesMade += $rep.MatchCount
                                                    $totalReplacements += $rep.MatchCount
                                                    $fileLog.details += "OLEDB $($setRes.PropertyName) in '$($conn.Name)': replaced $($rep.MatchCount) occurrence(s)"
                                                } else {
                                                    $fileLog.details += "OLEDB connection warning in '$($conn.Name)': could not write connection text ($($setRes.Error))"
                                                }
                                            }
                                        }
                                        $cmd = $ole.CommandText
                                        if ($cmd) {
                                            $rep = Invoke-SafeReplacement -InputText $cmd -Rules $validRules
                                            if ($rep.MatchCount -gt 0) {
                                                $ole.CommandText = $rep.ResultText
                                                $fileModified = $true
                                                $fileLog.changesMade += $rep.MatchCount
                                                $totalReplacements += $rep.MatchCount
                                                $fileLog.details += "OLEDB CommandText in '$($conn.Name)': replaced $($rep.MatchCount) occurrence(s)"
                                            }
                                        }
                                    } catch { }
                                }

                                if ($conn.ODBCConnection) {
                                    try {
                                        $odbc = $conn.ODBCConnection
                                        $connText = Get-ExcelConnectionText -ConnectionObject $odbc
                                        $cStr = $connText.Value
                                        if ($cStr) {
                                            $rep = Invoke-SafeReplacement -InputText $cStr -Rules $validRules
                                            if ($rep.MatchCount -gt 0) {
                                                $setRes = Set-ExcelConnectionText -ConnectionObject $odbc -ConnectionText $rep.ResultText -PreferredProperty $connText.PropertyName
                                                if ($setRes.Success) {
                                                    $fileModified = $true
                                                    $fileLog.changesMade += $rep.MatchCount
                                                    $totalReplacements += $rep.MatchCount
                                                    $fileLog.details += "ODBC $($setRes.PropertyName) in '$($conn.Name)': replaced $($rep.MatchCount) occurrence(s)"
                                                } else {
                                                    $fileLog.details += "ODBC connection warning in '$($conn.Name)': could not write connection text ($($setRes.Error))"
                                                }
                                            }
                                        }
                                        $cmd = $odbc.CommandText
                                        if ($cmd) {
                                            $rep = Invoke-SafeReplacement -InputText $cmd -Rules $validRules
                                            if ($rep.MatchCount -gt 0) {
                                                $odbc.CommandText = $rep.ResultText
                                                $fileModified = $true
                                                $fileLog.changesMade += $rep.MatchCount
                                                $totalReplacements += $rep.MatchCount
                                                $fileLog.details += "ODBC CommandText in '$($conn.Name)': replaced $($rep.MatchCount) occurrence(s)"
                                            }
                                        }
                                    } catch { }
                                }
                            }
                        }
                    } catch {
                        $fileLog.details += "Connection warning: $_"
                    }
                }

                # 7. Apply QueryTables replacement
                if ($updateConnections) {
                    try {
                        if ($wb.Worksheets) {
                            foreach ($ws in $wb.Worksheets) {
                                if ($ws.QueryTables) {
                                    foreach ($qt in $ws.QueryTables) {
                                        try {
                                            $qtConn = $qt.Connection
                                            if ($qtConn) {
                                                $rep = Invoke-SafeReplacement -InputText $qtConn -Rules $validRules
                                                if ($rep.MatchCount -gt 0) {
                                                    $qt.Connection = $rep.ResultText
                                                    $fileModified = $true
                                                    $fileLog.changesMade += $rep.MatchCount
                                                    $totalReplacements += $rep.MatchCount
                                                    $fileLog.details += "QueryTable in '$($ws.Name)': replaced $($rep.MatchCount) occurrence(s)"
                                                }
                                            }
                                        } catch { }
                                    }
                                }
                            }
                        }
                    } catch { }
                }

                # 8. Apply VBA Macros replacement
                if ($updateVba -and ($file.Extension.ToLower() -in @(".xlsm", ".xlsb"))) {
                    if ($preSnapshot.IsVbaSigned) {
                        $fileLog.details += "WARNING: VBA Project is digitally signed. Skipped VBA modification to prevent signature invalidation."
                    } else {
                        try {
                            if ($wb.VBProject.Protection -eq 1) {
                                $fileLog.details += "WARNING: VBA Project is password-protected/locked. Skipped VBA modification."
                            } else {
                                foreach ($comp in $wb.VBProject.VBComponents) {
                                    $cm = $comp.CodeModule
                                    if ($cm -and $cm.CountOfLines -gt 0) {
                                        $lineCount = $cm.CountOfLines
                                        for ($i = 1; $i -le $lineCount; $i++) {
                                            $line = $cm.Lines($i, 1)
                                            $rep = Invoke-SafeReplacement -InputText $line -Rules $validRules
                                            if ($rep.MatchCount -gt 0) {
                                                $cm.ReplaceLine($i, $rep.ResultText)
                                                $fileModified = $true
                                                $fileLog.changesMade += $rep.MatchCount
                                                $totalReplacements += $rep.MatchCount
                                                $fileLog.details += "VBA Line $i in '$($comp.Name)': replaced $($rep.MatchCount) occurrence(s)"
                                            }
                                        }
                                    }
                                }
                            }
                        } catch {
                            $fileLog.details += "VBA access warning: $_"
                        }
                    }
                }

                if (-not $fileModified) {
                    $wb.Close($false)
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
                    $wb = $null
                    Remove-StagingFile -StagingPath $stagingPath
                    $fileLog.status = "No Changes"
                } else {
                    # 9. Save staging workbook
                    $wb.Save()
                    $wb.Close($false)
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
                    $wb = $null

                    # 10. Reopen staging workbook ReadOnly to verify integrity and semantic snapshot
                    $verifyWb = $excel.Workbooks.Open($stagingPath, 0, $true)
                    $postSnapshot = Get-WorkbookSnapshot -Workbook $verifyWb
                    $verifyWb.Close($false)
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($verifyWb) | Out-Null
                    $verifyWb = $null

                    $cmp = Compare-WorkbookSnapshot -PreSnapshot $preSnapshot -PostSnapshot $postSnapshot
                    if (-not $cmp.IsValid) {
                        throw "Semantic validation failed on staged workbook: $($cmp.Errors -join '; ')"
                    }

                    # Cancellation Checkpoint 4 (Before commit)
                    if ($opId -and (Get-Command "Test-OperationCancellationRequested" -ErrorAction SilentlyContinue) -and (Test-OperationCancellationRequested $opId)) {
                        Remove-StagingFile -StagingPath $stagingPath
                        $cancelledDuringBatch = $true
                        break
                    }

                    # 11. Commit staging file using guarded Commit-StagingWorkbook (Rule 12: Critical section)
                    if ($opId -and (Get-Command "Set-OperationStatus" -ErrorAction SilentlyContinue)) {
                        $null = Set-OperationStatus $opId "COMMITTING" "Committing $($file.Name)"
                    }
                    $commitRes = Commit-StagingWorkbook -StagingPath $stagingPath -OriginalPath $file.FullName -InitialHash $initialHash -OperationId $opId
                    if ($opId -and (Get-Command "Set-OperationStatus" -ErrorAction SilentlyContinue)) {
                        $null = Set-OperationStatus $opId "RUNNING" "Committed $($file.Name)"
                    }
                    if (-not $commitRes.Success) {
                        if ($commitRes.ErrorCode -eq "CRITICAL_MANUAL_RECOVERY_REQUIRED") {
                            $fileLog.criticalRecovery = $commitRes
                        }
                        throw "Commit failed: $($commitRes.Error) (Code: $($commitRes.ErrorCode))"
                    }

                    $fileLog.status = "Updated"
                    $fileLog.finalHash = $commitRes.FinalHash
                    $fileLog.details += "Committed via $($commitRes.Method). Verified SHA256: $($commitRes.FinalHash)"
                    $updatedFilesCount++
                    $committedFiles += $file.FullName
                }
            } catch {
                $hasFatalFileError = $true
                $fileLog.status = "Error"
                $fileLog.details += "File error: $_"

                if ($commitRes -and $commitRes.ErrorCode -eq "CRITICAL_MANUAL_RECOVERY_REQUIRED") {
                    $fileLog.criticalRecovery = $commitRes
                    $fileLog.details += "CRITICAL: Manual recovery required. Recovery file: $($commitRes.recoveryFilePath)"
                } else {
                    if ($stagingPath -and (Test-Path -LiteralPath $stagingPath)) {
                        Remove-StagingFile -StagingPath $stagingPath
                    }
                }

                if ($wb) {
                    try { $wb.Close($false) } catch { }
                    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null } catch { }
                    $wb = $null
                }

                # If atomicBatch is enabled, stop immediately on first error
                if ($atomicBatch) {
                    $updateLog += $fileLog
                    break
                }
            }

            $updateLog += $fileLog
        }
    } finally {
        # Restore isolation settings and safely terminate isolated instance
        if ($iso) {
            Close-IsolatedExcelInstance -IsolationContext $iso
        }
    }

    # 11.5 Cooperative Cancellation Handler
    if ($cancelledDuringBatch) {
        if ($atomicBatch -and $committedFiles.Count -gt 0 -and $backupDir) {
            Set-ProgressState $false "update" $processedCount $totalCount "İşlem iptal edildi! Değişiklikler geri alınıyor (Rollback)..."
            if ($opId -and (Get-Command "Set-OperationStatus" -ErrorAction SilentlyContinue)) { $null = Set-OperationStatus $opId "ROLLING_BACK" "Rolling back due to cancellation" }
            $restoreRes = Restore-BatchBackup -BackupDir $backupDir -TargetDir $DirectoryPath
            if ($restoreRes.Success) {
                if ($opId -and (Get-Command "Set-OperationStatus" -ErrorAction SilentlyContinue)) { $null = Set-OperationStatus $opId "CANCELLED" "Cancelled and Rolled Back" }
                return @{
                    success = $false
                    error = "İşlem kullanıcı tarafından iptal edildi. Yapılan tüm değişiklikler başarıyla geri alındı (Rollback)."
                    batchStatus = "CANCELLED"
                    directory = $DirectoryPath
                    totalFilesProcessed = $processedCount
                    updatedFilesCount = 0
                    totalReplacements = 0
                    logs = $updateLog
                    backupDirectory = $backupDir
                }
            } else {
                if ($opId -and (Get-Command "Set-OperationStatus" -ErrorAction SilentlyContinue)) { $null = Set-OperationStatus $opId "ROLLBACK_PARTIAL_FAILURE" "Rollback partial failure on cancel" }
                return @{
                    success = $false
                    error = "İşlem iptal edildi fakat geri alma (Rollback) kısmen başarısız oldu! Acil müdahale gerekebilir: $($restoreRes.Errors -join '; ')"
                    batchStatus = "ROLLBACK_PARTIAL_FAILURE"
                    directory = $DirectoryPath
                    totalFilesProcessed = $processedCount
                    updatedFilesCount = 0
                    totalReplacements = 0
                    logs = $updateLog
                    backupDirectory = $backupDir
                    failedRestoreFiles = $restoreRes.failedRestoreFiles
                }
            }
        } else {
            if ($opId -and (Get-Command "Set-OperationStatus" -ErrorAction SilentlyContinue)) { $null = Set-OperationStatus $opId "CANCELLED" "Cancelled by user" }
            return @{
                success = $false
                error = "İşlem kullanıcı tarafından iptal edildi."
                batchStatus = "CANCELLED"
                directory = $DirectoryPath
                totalFilesProcessed = $processedCount
                updatedFilesCount = 0
                totalReplacements = 0
                logs = $updateLog
                backupDirectory = $backupDir
            }
        }
    }

    # 12. Transaction Rollback if atomicBatch is requested and an error occurred
    $batchRolledBack = $false
    if ($atomicBatch -and $hasFatalFileError -and $backupDir) {
        Set-ProgressState $false "update" $processedCount $totalCount "Hata tespit edildi! Toplu işlem geri alınıyor (Rollback)..."
        $restoreRes = Restore-BatchBackup -BackupDir $backupDir -TargetDir $DirectoryPath
        $batchRolledBack = $true

        if (-not $restoreRes.Success) {
            return @{
                success = $false
                error = "İşlem sırasında bir dosyada hata oluştu ve yedekten geri yükleme (Rollback) kısmen başarısız oldu! Acil manuel müdahale gerekebilir: $($restoreRes.Errors -join '; ')"
                batchStatus = "ROLLBACK_PARTIAL_FAILURE"
                directory = $DirectoryPath
                totalFilesProcessed = $processedCount
                updatedFilesCount = 0
                totalReplacements = 0
                logs = $updateLog
                backupDirectory = $backupDir
                failedRestoreFiles = $restoreRes.failedRestoreFiles
            }
        }

        return @{
            success = $false
            error = "İşlem sırasında bir dosyada hata oluştu. Veri güvenliği gereği tüm değişiklikler geri alındı (Batch Rollback)."
            batchStatus = "ROLLED_BACK"
            directory = $DirectoryPath
            totalFilesProcessed = $processedCount
            updatedFilesCount = 0
            totalReplacements = 0
            logs = $updateLog
            backupDirectory = $backupDir
            failedRestoreFiles = @()
        }
    }

    $overallSuccess = (-not $hasFatalFileError)
    $msg = if ($overallSuccess) { "Güncelleme Tamamlandı!" } else { "Güncelleme hatalarla tamamlandı." }
    Set-ProgressState $false "update" $totalCount $totalCount $msg

    return @{
        success = $overallSuccess
        batchStatus = if ($overallSuccess) { "COMMITTED" } else { "PARTIAL_ERROR" }
        directory = $DirectoryPath
        totalFilesProcessed = $files.Count
        updatedFilesCount = $updatedFilesCount
        totalReplacements = $totalReplacements
        logs = $updateLog
        backupDirectory = $backupDir
    }
}

function Preview-ExcelDirectory ($DirectoryPath, $Rules, $Options = @{}, [string]$OperationId = "") {
    $files = Get-ExcelFiles $DirectoryPath
    $totalCount = $files.Count
    
    if ($totalCount -eq 0) {
        return @{
            Success = $true
            TotalFiles = 0
            MatchingFiles = 0
            TotalReplacements = 0
            Files = @()
            Warnings = @()
        }
    }

    $ruleValidation = Test-ReplacementRulesValid $Rules
    if (-not $ruleValidation.IsValid) {
        return @{
            Success = $false
            Error = "Rule validation failed: $($ruleValidation.Error)"
            TotalFiles = $totalCount
            MatchingFiles = 0
            TotalReplacements = 0
            Files = @()
            Warnings = @()
        }
    }
    $validRules = $ruleValidation.Rules

    $iso = $null
    try {
        $iso = New-IsolatedExcelInstance
    } catch {
        return @{
            Success = $false
            Error = "Could not initialize isolated Excel instance: $_"
            TotalFiles = $totalCount
            MatchingFiles = 0
            TotalReplacements = 0
            Files = @()
            Warnings = @()
        }
    }

    $excel = $iso.Excel
    $previewFiles = @()
    $totalMatchingFiles = 0
    $overallReplacements = 0
    $globalWarnings = @()
    $processedCount = 0

    $checkQueries = if ($Options -and $Options.updateQueries -ne $null) { $Options.updateQueries } else { $true }
    $checkConnections = if ($Options -and $Options.updateConnections -ne $null) { $Options.updateConnections } else { $true }
    $checkVba = if ($Options -and $Options.updateVba -ne $null) { $Options.updateVba } else { $true }

    try {
        foreach ($file in $files) {
            $processedCount++
            if (-not [string]::IsNullOrWhiteSpace($OperationId)) {
                if (Get-Command "Calculate-WeightedProgress" -ErrorAction SilentlyContinue) {
                    $pct = Calculate-WeightedProgress -TotalFiles $totalCount -CompletedFiles ($processedCount - 1) -CurrentFileStage "Excel Update"
                    if (Get-Command "Update-OperationProgress" -ErrorAction SilentlyContinue) {
                        Update-OperationProgress -OperationId $OperationId -ProcessedFiles $processedCount -TotalFiles $totalCount -CurrentFile $file.Name -CurrentStage "Previewing" -ProgressPercent $pct
                    }
                }
            }

            $fSha = Get-FileSha256 $file.FullName
            $fWritable = Test-FileWritable $file.FullName

            $fPreview = @{
                fileName = $file.Name
                filePath = $file.FullName
                fileSha256 = $fSha
                isWritable = $fWritable
                extension = $file.Extension.ToLower()
                sizeBytes = $file.Length
                queryMatches = 0
                connectionMatches = 0
                vbaMatches = 0
                totalReplacements = 0
                details = @()
                warnings = @()
            }

            if (-not $fWritable) {
                $fPreview.warnings += "Dosya kilitli veya salt-okunur."
            }

            $wb = $null
            try {
                # Strictly READ-ONLY open with UpdateLinks = 0
                $wb = $excel.Workbooks.Open($file.FullName, 0, $true)

                # 1. Power Queries Simulation
                if ($checkQueries) {
                    try {
                        if ($wb.Queries) {
                            $queries = $wb.Queries
                            $qCount = $queries.Count
                            for ($qIdx = 1; $qIdx -le $qCount; $qIdx++) {
                                $q = $queries.Item($qIdx)
                                if ($q) {
                                    $formula = $q.Formula
                                    if ($formula) {
                                        $rep = Invoke-SafeReplacement -InputText $formula -Rules $validRules
                                        if ($rep.MatchCount -gt 0) {
                                            $fPreview.queryMatches += $rep.MatchCount
                                            $fPreview.totalReplacements += $rep.MatchCount
                                            $fPreview.details += "PowerQuery '$($q.Name)': $($rep.MatchCount) eşleşme"
                                        }
                                    }
                                    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($q) | Out-Null
                                }
                            }
                            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($queries) | Out-Null
                        }
                    } catch { }
                }

                # 2. Connections Simulation
                if ($checkConnections) {
                    try {
                        if ($wb.Connections) {
                            $conns = $wb.Connections
                            $cCount = $conns.Count
                            for ($cIdx = 1; $cIdx -le $cCount; $cIdx++) {
                                $conn = $conns.Item($cIdx)
                                if ($conn) {
                                    try {
                                        $connNameRep = Invoke-SafeReplacement -InputText "$($conn.Name)" -Rules $validRules
                                        if ($connNameRep.MatchCount -gt 0) {
                                            $fPreview.connectionMatches += $connNameRep.MatchCount
                                            $fPreview.totalReplacements += $connNameRep.MatchCount
                                            $fPreview.details += "Connection name '$($conn.Name)': $($connNameRep.MatchCount) eşleşme"
                                        }
                                    } catch { }

                                    try {
                                        $connDesc = "$($conn.Description)"
                                        if (-not [string]::IsNullOrWhiteSpace($connDesc)) {
                                            $descRep = Invoke-SafeReplacement -InputText $connDesc -Rules $validRules
                                            if ($descRep.MatchCount -gt 0) {
                                                $fPreview.connectionMatches += $descRep.MatchCount
                                                $fPreview.totalReplacements += $descRep.MatchCount
                                                $fPreview.details += "Connection description in '$($conn.Name)': $($descRep.MatchCount) eşleşme"
                                            }
                                        }
                                    } catch { }

                                    if ($conn.OLEDBConnection) {
                                        try {
                                            $connText = Get-ExcelConnectionText -ConnectionObject $conn.OLEDBConnection
                                            $cStr = $connText.Value
                                            if ($cStr) {
                                                $rep = Invoke-SafeReplacement -InputText $cStr -Rules $validRules
                                                if ($rep.MatchCount -gt 0) {
                                                    $fPreview.connectionMatches += $rep.MatchCount
                                                    $fPreview.totalReplacements += $rep.MatchCount
                                                    $fPreview.details += "OLEDB $($connText.PropertyName) in '$($conn.Name)': $($rep.MatchCount) eşleşme"
                                                }
                                            }
                                            $cmd = $conn.OLEDBConnection.CommandText
                                            if ($cmd) {
                                                $rep = Invoke-SafeReplacement -InputText $cmd -Rules $validRules
                                                if ($rep.MatchCount -gt 0) {
                                                    $fPreview.connectionMatches += $rep.MatchCount
                                                    $fPreview.totalReplacements += $rep.MatchCount
                                                    $fPreview.details += "OLEDB CommandText in '$($conn.Name)': $($rep.MatchCount) eşleşme"
                                                }
                                            }
                                        } catch { }
                                    }
                                    if ($conn.ODBCConnection) {
                                        try {
                                            $connText = Get-ExcelConnectionText -ConnectionObject $conn.ODBCConnection
                                            $cStr = $connText.Value
                                            if ($cStr) {
                                                $rep = Invoke-SafeReplacement -InputText $cStr -Rules $validRules
                                                if ($rep.MatchCount -gt 0) {
                                                    $fPreview.connectionMatches += $rep.MatchCount
                                                    $fPreview.totalReplacements += $rep.MatchCount
                                                    $fPreview.details += "ODBC $($connText.PropertyName) in '$($conn.Name)': $($rep.MatchCount) eşleşme"
                                                }
                                            }
                                            $cmd = $conn.ODBCConnection.CommandText
                                            if ($cmd) {
                                                $rep = Invoke-SafeReplacement -InputText $cmd -Rules $validRules
                                                if ($rep.MatchCount -gt 0) {
                                                    $fPreview.connectionMatches += $rep.MatchCount
                                                    $fPreview.totalReplacements += $rep.MatchCount
                                                    $fPreview.details += "ODBC CommandText in '$($conn.Name)': $($rep.MatchCount) eşleşme"
                                                }
                                            }
                                        } catch { }
                                    }
                                    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($conn) | Out-Null
                                }
                            }
                            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($conns) | Out-Null
                        }
                    } catch { }
                }

                # 3. VBA Macros Simulation
                if ($checkVba -and ($file.Extension.ToLower() -in @(".xlsm", ".xlsb"))) {
                    try {
                        $vbProj = $wb.VBProject
                        if ($vbProj) {
                            $isSigned = $false
                            try { if ($vbProj.Signature) { $isSigned = $true } } catch { }
                            if ($isSigned) {
                                $fPreview.warnings += "VBA Projesi dijital imzalı (imzanın bozulmaması için atlanacak)."
                            } elseif ($vbProj.Protection -eq 1) {
                                $fPreview.warnings += "VBA Projesi şifre korumalı / kilitli (atlanacak)."
                            } else {
                                $comps = $vbProj.VBComponents
                                if ($comps) {
                                    $cCount = $comps.Count
                                    for ($cIdx = 1; $cIdx -le $cCount; $cIdx++) {
                                        $comp = $comps.Item($cIdx)
                                        if ($comp) {
                                            $cm = $comp.CodeModule
                                            if ($cm -and $cm.CountOfLines -gt 0) {
                                                $lineCount = $cm.CountOfLines
                                                for ($l = 1; $l -le $lineCount; $l++) {
                                                    $line = $cm.Lines($l, 1)
                                                    $rep = Invoke-SafeReplacement -InputText $line -Rules $validRules
                                                    if ($rep.MatchCount -gt 0) {
                                                        $fPreview.vbaMatches += $rep.MatchCount
                                                        $fPreview.totalReplacements += $rep.MatchCount
                                                        $fPreview.details += "VBA Line $l in '$($comp.Name)': $($rep.MatchCount) eşleşme"
                                                    }
                                                }
                                            }
                                            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($comp) | Out-Null
                                        }
                                    }
                                    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($comps) | Out-Null
                                }
                            }
                            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($vbProj) | Out-Null
                        }
                    } catch {
                        $fPreview.warnings += "VBA erişim uyarısı: $_"
                    }
                }

                # Close strictly without saving!
                $wb.Close($false)
                [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
                $wb = $null

            } catch {
                $fPreview.warnings += "Dosya okuma hatası: $_"
                if ($wb) {
                    try { $wb.Close($false) } catch { }
                    try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null } catch { }
                    $wb = $null
                }
            }

            if ($fPreview.totalReplacements -gt 0) {
                $totalMatchingFiles++
                $overallReplacements += $fPreview.totalReplacements
            }

            $previewFiles += $fPreview
        }
    } finally {
        if ($iso) {
            Close-IsolatedExcelInstance $iso
            $iso = $null
        }
    }

    return @{
        Success = $true
        TotalFiles = $totalCount
        MatchingFiles = $totalMatchingFiles
        TotalReplacements = $overallReplacements
        Files = $previewFiles
        Warnings = $globalWarnings
    }
}

function Restore-VerifiedBackup ($BackupDirectory, $TargetDirectory, [string]$OperationId = "") {
    if ([string]::IsNullOrWhiteSpace($BackupDirectory) -or -not (Test-Path -LiteralPath $BackupDirectory -PathType Container)) {
        return @{ Success = $false; Error = "Yedek klasörü bulunamadı veya geçersiz: $BackupDirectory" }
    }
    if ([string]::IsNullOrWhiteSpace($TargetDirectory) -or -not (Test-Path -LiteralPath $TargetDirectory -PathType Container)) {
        return @{ Success = $false; Error = "Hedef klasör bulunamadı veya geçersiz: $TargetDirectory" }
    }

    $resolvedBackup = (Resolve-Path -LiteralPath $BackupDirectory).Path
    $resolvedTarget = (Resolve-Path -LiteralPath $TargetDirectory).Path

    $manifestPath = Join-Path $resolvedBackup "backup_manifest.json"
    if (-not (Test-Path -LiteralPath $manifestPath)) {
        return @{ Success = $false; Error = "Yedek klasöründe geçerli bir 'backup_manifest.json' bulunamadı: $resolvedBackup" }
    }

    # 1. Pre-restore Recovery Snapshot (Rule 21)
    if ($OperationId -and (Get-Command "Update-OperationProgress" -ErrorAction SilentlyContinue)) {
        $null = Update-OperationProgress -OperationId $OperationId -ProcessedFiles 0 -TotalFiles 1 -CurrentFile "Pre-Restore Snapshot" -CurrentStage "Pre-Restore Backup" -ProgressPercent 25
    }
    $preRecoveryRes = Create-ExcelBackup -DirectoryPath $resolvedTarget
    if (-not $preRecoveryRes.success) {
        return @{ Success = $false; Error = "Restore öncesi güvenlik yedeği alınamadı: $($preRecoveryRes.error)" }
    }

    # 2. Perform Verified Batch Restore
    if ($OperationId -and (Get-Command "Update-OperationProgress" -ErrorAction SilentlyContinue)) {
        $null = Update-OperationProgress -OperationId $OperationId -ProcessedFiles 0 -TotalFiles 1 -CurrentFile "Restoring files" -CurrentStage "Restoring" -ProgressPercent 65
    }
    $restoreRes = Restore-BatchBackup -BackupDir $resolvedBackup -TargetDir $resolvedTarget

    if (-not $restoreRes.Success) {
        return @{
            Success = $false
            Error = "Yedekten geri yükleme başarısız oldu: $($restoreRes.Errors -join '; ')"
            FailedFiles = $restoreRes.failedRestoreFiles
            PreRestoreBackup = $preRecoveryRes.backupDirectory
        }
    }

    if ($OperationId -and (Get-Command "Update-OperationProgress" -ErrorAction SilentlyContinue)) {
        $null = Update-OperationProgress -OperationId $OperationId -ProcessedFiles 1 -TotalFiles 1 -CurrentFile "Completed" -CurrentStage "Done" -ProgressPercent 100
    }

    return @{
        Success = $true
        RestoredFiles = $restoreRes.RestoredFiles
        PreRestoreBackup = $preRecoveryRes.backupDirectory
    }
}
