# ==============================================================================
# Excel SQL Connect Pro - Operation Manager & State Machine Engine
# ==============================================================================

# Ensure concurrent collection types are available
$null = [System.Collections.Concurrent.ConcurrentDictionary[string, object]]

# Cross-Runspace synchronization locks stored in AppDomain for process-wide sharing
$domainOpLock = [System.AppDomain]::CurrentDomain.GetData("OperationRegistryLock")
if ($null -eq $domainOpLock) {
    $domainOpLock = [object]::new()
    [System.AppDomain]::CurrentDomain.SetData("OperationRegistryLock", $domainOpLock)
}
$global:OperationRegistryLock = $domainOpLock

$domainHistLock = [System.AppDomain]::CurrentDomain.GetData("HistoryFileLock")
if ($null -eq $domainHistLock) {
    $domainHistLock = [object]::new()
    [System.AppDomain]::CurrentDomain.SetData("HistoryFileLock", $domainHistLock)
}
$global:HistoryFileLock = $domainHistLock

$domainAuditLock = [System.AppDomain]::CurrentDomain.GetData("AuditLogLock")
if ($null -eq $domainAuditLock) {
    $domainAuditLock = [object]::new()
    [System.AppDomain]::CurrentDomain.SetData("AuditLogLock", $domainAuditLock)
}
$global:AuditLogLock = $domainAuditLock

function Resolve-WritableLogDirectory {
    param([string]$PreferredDirectory = "")

    $candidates = @()
    if (-not [string]::IsNullOrWhiteSpace($PreferredDirectory)) {
        $candidates += $PreferredDirectory
    }
    if (-not [string]::IsNullOrWhiteSpace($env:TEMP)) {
        $candidates += (Join-Path $env:TEMP "ExcelSqlConnectPro\logs")
    }

    foreach ($candidate in $candidates) {
        try {
            [System.IO.Directory]::CreateDirectory($candidate) | Out-Null
            $testPath = Join-Path $candidate ".__audit_write_test_$([System.Guid]::NewGuid().ToString('N')).tmp"
            [System.IO.File]::WriteAllText($testPath, "test", [System.Text.Encoding]::UTF8)
            Remove-Item -Path $testPath -Force -ErrorAction SilentlyContinue
            return $candidate
        } catch { }
    }

    if (-not [string]::IsNullOrWhiteSpace($PreferredDirectory)) {
        return $PreferredDirectory
    }
    return (Join-Path (Get-Location) "logs")
}

$domainAuditDir = [System.AppDomain]::CurrentDomain.GetData("AuditLogDirectory")
if ([string]::IsNullOrWhiteSpace($domainAuditDir)) {
    $domainAuditDir = Resolve-WritableLogDirectory -PreferredDirectory (Join-Path (Get-Location) "logs")
    [System.AppDomain]::CurrentDomain.SetData("AuditLogDirectory", $domainAuditDir)
}
$global:AuditLogDirectory = "$domainAuditDir"

# Cross-Runspace in-memory registry for operation states stored in AppDomain
$domainReg = [System.AppDomain]::CurrentDomain.GetData("OperationsRegistry")
if ($null -eq $domainReg) {
    $domainReg = [System.Collections.Concurrent.ConcurrentDictionary[string, hashtable]]::new()
    [System.AppDomain]::CurrentDomain.SetData("OperationsRegistry", $domainReg)
}
$global:OperationsRegistry = $domainReg
$global:OPERATION_LOCK_TIMEOUT_MS = 1500
$global:LastActiveOperationId = $null

function Get-ActiveOperationId () {
    $lockTaken = $false
    try {
        [System.Threading.Monitor]::TryEnter($global:OperationRegistryLock, $global:OPERATION_LOCK_TIMEOUT_MS, [ref]$lockTaken) | Out-Null
        if (-not $lockTaken) {
            return $global:LastActiveOperationId
        }
        foreach ($key in $global:OperationsRegistry.Keys) {
            $existing = $global:OperationsRegistry[$key]
            if ($existing -and ($global:TERMINAL_OPERATION_STATES -notcontains $existing.status)) {
                $global:LastActiveOperationId = $existing.operationId
                return $existing.operationId
            }
        }
        return $null
    } finally {
        if ($lockTaken) {
            [System.Threading.Monitor]::Exit($global:OperationRegistryLock)
        }
    }
}

# ------------------------------------------------------------------------------
# 1. State Machine Catalog & Transition Validator
# ------------------------------------------------------------------------------
$global:VALID_OPERATION_STATES = @(
    "QUEUED",
    "RUNNING",
    "CANCELLATION_REQUESTED",
    "COMMITTING",
    "ROLLING_BACK",
    "COMPLETED",
    "FAILED",
    "CANCELLED",
    "ROLLBACK_PARTIAL_FAILURE",
    "CRITICAL_MANUAL_RECOVERY_REQUIRED",
    "STALE_INTERRUPTED"
)

$global:TERMINAL_OPERATION_STATES = @(
    "COMPLETED",
    "FAILED",
    "CANCELLED",
    "ROLLBACK_PARTIAL_FAILURE",
    "CRITICAL_MANUAL_RECOVERY_REQUIRED",
    "STALE_INTERRUPTED"
)

function Test-ValidStateTransition ([string]$FromState, [string]$ToState) {
    if ([string]::IsNullOrWhiteSpace($FromState) -or [string]::IsNullOrWhiteSpace($ToState)) {
        return $false
    }

    if ($FromState -eq $ToState) {
        return $true
    }

    # Terminal states can NEVER transition to any other state
    if ($global:TERMINAL_OPERATION_STATES -contains $FromState) {
        return $false
    }

    switch ($FromState) {
        "QUEUED" {
            return ($ToState -in @("RUNNING", "FAILED", "CANCELLED", "STALE_INTERRUPTED"))
        }
        "RUNNING" {
            return ($ToState -in @(
                "CANCELLATION_REQUESTED",
                "COMMITTING",
                "ROLLING_BACK",
                "COMPLETED",
                "FAILED",
                "CANCELLED",
                "CRITICAL_MANUAL_RECOVERY_REQUIRED",
                "STALE_INTERRUPTED"
            ))
        }
        "CANCELLATION_REQUESTED" {
            return ($ToState -in @(
                "COMMITTING", # If commit was already in progress
                "ROLLING_BACK",
                "CANCELLED",
                "FAILED",
                "ROLLBACK_PARTIAL_FAILURE",
                "STALE_INTERRUPTED"
            ))
        }
        "COMMITTING" {
            return ($ToState -in @(
                "RUNNING", # Next file in batch
                "COMPLETED",
                "ROLLING_BACK",
                "FAILED",
                "CRITICAL_MANUAL_RECOVERY_REQUIRED",
                "STALE_INTERRUPTED"
            ))
        }
        "ROLLING_BACK" {
            return ($ToState -in @(
                "CANCELLED",
                "FAILED",
                "ROLLBACK_PARTIAL_FAILURE",
                "CRITICAL_MANUAL_RECOVERY_REQUIRED",
                "STALE_INTERRUPTED"
            ))
        }
        default {
            return $false
        }
    }
}

function Assert-ValidStateTransition ([string]$FromState, [string]$ToState) {
    if (-not (Test-ValidStateTransition $FromState $ToState)) {
        throw "Invalid operation state transition from '$FromState' to '$ToState'."
    }
}

# ------------------------------------------------------------------------------
# 2. Sensitive Data & Credential Masking (Fail-Closed)
# ------------------------------------------------------------------------------
function Protect-SensitiveData ($InputData) {
    if ($null -eq $InputData) { return $null }

    # Masking regex rules
    $maskPatterns = @(
        @{ Pattern = '(?i)(Password|Pwd)\s*=\s*[^;\r\n"]+'; Replace = '$1=***' },
        @{ Pattern = '(?i)(User Id|Uid)\s*=\s*[^;\r\n"]+';  Replace = '$1=***' },
        @{ Pattern = '(?i)(Bearer\s+)[A-Za-z0-9_\-\.]+';     Replace = '$1***' },
        @{ Pattern = '(?i)(token=)[A-Za-z0-9_\-\.]+';       Replace = '$1***' }
    )

    if ($InputData -is [string]) {
        $result = $InputData
        foreach ($rule in $maskPatterns) {
            $result = [regex]::Replace($result, $rule.Pattern, $rule.Replace)
        }
        return $result
    }

    if ($InputData -is [System.Collections.IDictionary] -or $InputData -is [hashtable]) {
        $cleanHt = @{}
        foreach ($k in $InputData.Keys) {
            $val = $InputData[$k]
            $cleanHt[$k] = Protect-SensitiveData $val
        }
        return $cleanHt
    }

    if ($InputData -is [System.Collections.IEnumerable] -and -not ($InputData -is [string])) {
        $cleanList = @()
        foreach ($item in $InputData) {
            $cleanList += Protect-SensitiveData $item
        }
        return $cleanList
    }

    if ($InputData -is [System.Exception]) {
        return Protect-SensitiveData $InputData.Message
    }

    return $InputData
}

# ------------------------------------------------------------------------------
# 3. Thread-Safe Operation Registry
# ------------------------------------------------------------------------------
function Reset-OperationRegistry () {
    [System.Threading.Monitor]::Enter($global:OperationRegistryLock)
    try {
        $global:OperationsRegistry.Clear()
    } finally {
        [System.Threading.Monitor]::Exit($global:OperationRegistryLock)
    }
}

function New-OperationId () {
    $ts = (Get-Date).ToUniversalTime().ToString("yyyyMMdd-HHmmss")
    $rand = [System.Guid]::NewGuid().ToString("N").Substring(0, 6)
    return "op-$ts-$rand"
}

function Get-AuditLogFilePath {
    param(
        [Parameter(Mandatory=$true)][string]$OperationId,
        [string]$LogDirectory = ""
    )

    if ([string]::IsNullOrWhiteSpace($LogDirectory)) {
        $LogDirectory = if (-not [string]::IsNullOrWhiteSpace($global:AuditLogDirectory)) { $global:AuditLogDirectory } else { Join-Path (Get-Location) "logs" }
    }

    $monthDir = Join-Path $LogDirectory (Get-Date).ToUniversalTime().ToString("yyyy-MM")
    return (Join-Path $monthDir "$OperationId.jsonl")
}

function Write-RawAuditLogLine {
    param(
        [Parameter(Mandatory=$true)][string]$OperationId,
        [Parameter(Mandatory=$true)][string]$Level,
        [Parameter(Mandatory=$true)][string]$EventName,
        [string]$File = "",
        [string]$Stage = "",
        [string]$Message = "",
        [hashtable]$Data = @{},
        [string]$LogDirectory = ""
    )

    if ([string]::IsNullOrWhiteSpace($LogDirectory)) {
        $LogDirectory = if (-not [string]::IsNullOrWhiteSpace($global:AuditLogDirectory)) { $global:AuditLogDirectory } else { Join-Path (Get-Location) "logs" }
    }

    try {
        $logFilePath = Get-AuditLogFilePath -OperationId $OperationId -LogDirectory $LogDirectory
        $logDir = Split-Path -Parent $logFilePath
        [System.IO.Directory]::CreateDirectory($logDir) | Out-Null

        $logEntry = @{
            timestamp   = (Get-Date).ToUniversalTime().ToString("o")
            operationId = $OperationId
            level       = $Level.ToUpper()
            event       = $EventName.ToUpper()
            file        = $File
            stage       = $Stage
            message     = (Protect-SensitiveData $Message)
            data        = (Protect-SensitiveData $Data)
        }
        $jsonLine = $logEntry | ConvertTo-Json -Depth 6 -Compress
        [System.IO.File]::AppendAllText($logFilePath, $jsonLine + [System.Environment]::NewLine, [System.Text.Encoding]::UTF8)
    } catch {
        try {
            [System.IO.Directory]::CreateDirectory($LogDirectory) | Out-Null
            $fallbackPath = Join-Path $LogDirectory "audit-write-failures.log"
            $failureLine = "$(Get-Date -Format o) operation=$OperationId event=$EventName error=$($_.Exception.Message)"
            [System.IO.File]::AppendAllText($fallbackPath, $failureLine + [System.Environment]::NewLine, [System.Text.Encoding]::UTF8)
        } catch { }
    }
}

function Register-Operation {
    param(
        [Parameter(Mandatory=$true)][string]$Type, # SCAN | PREVIEW | UPDATE | RESTORE
        [Parameter(Mandatory=$true)][string]$Directory,
        $Options = @{},
        $Rules = @()
    )

    $optHt = @{}
    if ($Options -is [System.Collections.IDictionary] -or $Options -is [hashtable]) {
        foreach ($k in $Options.Keys) { $optHt[$k] = $Options[$k] }
    } elseif ($Options -is [System.Management.Automation.PSCustomObject]) {
        foreach ($prop in $Options.PSObject.Properties) { $optHt[$prop.Name] = $prop.Value }
    }

    $cleanRules = @()
    if ($Rules) {
        foreach ($r in $Rules) {
            if ($r -is [System.Management.Automation.PSCustomObject]) {
                $cleanRules += @{ oldText = "$($r.oldText)"; newText = "$($r.newText)" }
            } elseif ($r -is [hashtable] -or $r -is [System.Collections.IDictionary]) {
                $cleanRules += @{ oldText = "$($r.oldText)"; newText = "$($r.newText)" }
            } else {
                $cleanRules += $r
            }
        }
    }

    [System.Threading.Monitor]::Enter($global:OperationRegistryLock)
    try {
        # Check concurrency: Only ONE heavy COM operation can be active
        $heavyTypes = @("SCAN", "PREVIEW", "UPDATE", "RESTORE")
        if ($heavyTypes -contains $Type.ToUpper()) {
            foreach ($key in $global:OperationsRegistry.Keys) {
                $existing = $global:OperationsRegistry[$key]
                if ($existing -and ($global:TERMINAL_OPERATION_STATES -notcontains $existing.status)) {
                    return @{
                        Success = $false
                        ErrorCode = "OPERATION_ALREADY_RUNNING"
                        ActiveOperationId = $existing.operationId
                        ActiveType = $existing.type
                        Error = "An operation is already running: $($existing.operationId) ($($existing.type)). Please wait or cancel it."
                    }
                }
            }
        }

        $opId = New-OperationId
        $nowIso = (Get-Date).ToUniversalTime().ToString("o")

        $newOp = @{
            operationId = $opId
            type = $Type.ToUpper()
            status = "QUEUED"
            createdAt = $nowIso
            startedAt = $null
            finishedAt = $null
            directory = $Directory
            totalFiles = 0
            processedFiles = 0
            updatedFiles = 0
            skippedFiles = 0
            failedFiles = 0
            currentFile = ""
            currentStage = "Queued"
            progressPercent = 0
            warnings = @()
            errors = @()
            backupDirectory = ""
            batchStatus = "PENDING"
            cancelRequested = $false
            options = $optHt
            rules = $cleanRules
            logs = @()
            scannedFiles = @()
            detectedIPs = @()
            resultData = $null
            auditLogPath = Get-AuditLogFilePath -OperationId $opId
        }

        $global:OperationsRegistry[$opId] = $newOp
        $global:LastActiveOperationId = $opId
        Write-RawAuditLogLine -OperationId $opId -Level "INFO" -EventName "OPERATION_REGISTERED" -Stage "Queued" -Message "Operation registered" -Data @{ type = $Type.ToUpper(); directory = $Directory }

        return @{
            Success = $true
            OperationId = $opId
            Status = "QUEUED"
            ErrorCode = ""
            Error = ""
        }
    } finally {
        [System.Threading.Monitor]::Exit($global:OperationRegistryLock)
    }
}

function Set-OperationStatus {
    param(
        [Parameter(Mandatory=$true)][string]$OperationId,
        [Parameter(Mandatory=$true)][string]$NewStatus,
        [string]$CurrentStage = $null
    )

    [System.Threading.Monitor]::Enter($global:OperationRegistryLock)
    try {
        if (-not $global:OperationsRegistry.ContainsKey($OperationId)) {
            return $false
        }

        $op = $global:OperationsRegistry[$OperationId]
        Assert-ValidStateTransition $op.status $NewStatus

        $op.status = $NewStatus
        if (-not [string]::IsNullOrWhiteSpace($CurrentStage)) {
            $op.currentStage = $CurrentStage
        }

        if ($NewStatus -eq "RUNNING" -and [string]::IsNullOrWhiteSpace($op.startedAt)) {
            $op.startedAt = (Get-Date).ToUniversalTime().ToString("o")
        }

        if ($global:TERMINAL_OPERATION_STATES -contains $NewStatus) {
            $op.finishedAt = (Get-Date).ToUniversalTime().ToString("o")
            if ($NewStatus -eq "COMPLETED") {
                $op.progressPercent = 100
            }
        }

        return $true
    } finally {
        [System.Threading.Monitor]::Exit($global:OperationRegistryLock)
    }
}

function Update-OperationProgress {
    param(
        [Parameter(Mandatory=$true)][string]$OperationId,
        [int]$ProcessedFiles,
        [int]$TotalFiles,
        [string]$CurrentFile = "",
        [string]$CurrentStage = "",
        [int]$ProgressPercent = -1
    )

    [System.Threading.Monitor]::Enter($global:OperationRegistryLock)
    try {
        if (-not $global:OperationsRegistry.ContainsKey($OperationId)) { return $false }
        $op = $global:OperationsRegistry[$OperationId]

        if ($TotalFiles -gt 0) { $op.totalFiles = $TotalFiles }
        if ($ProcessedFiles -ge 0) { $op.processedFiles = $ProcessedFiles }
        if (-not [string]::IsNullOrWhiteSpace($CurrentFile)) { $op.currentFile = $CurrentFile }
        if (-not [string]::IsNullOrWhiteSpace($CurrentStage)) { $op.currentStage = $CurrentStage }

        if ($ProgressPercent -ge 0) {
            # Monotonic progress: never drop unless resetting
            if ($ProgressPercent -gt $op.progressPercent) {
                $op.progressPercent = [Math]::Min(99, $ProgressPercent)
            }
        }

        Write-RawAuditLogLine -OperationId $OperationId -Level "INFO" -EventName "OPERATION_PROGRESS" -Stage $op.currentStage -File $op.currentFile -Message "Operation progress updated" -Data @{ processedFiles = [int]$op.processedFiles; totalFiles = [int]$op.totalFiles; progressPercent = [int]$op.progressPercent; status = "$($op.status)" }

        return $true
    } finally {
        [System.Threading.Monitor]::Exit($global:OperationRegistryLock)
    }
}

function Add-OperationWarning ([string]$OperationId, [string]$WarningMessage) {
    [System.Threading.Monitor]::Enter($global:OperationRegistryLock)
    try {
        if (-not $global:OperationsRegistry.ContainsKey($OperationId)) { return }
        $op = $global:OperationsRegistry[$OperationId]
        $clean = Protect-SensitiveData $WarningMessage
        $op.warnings += $clean
    } finally {
        [System.Threading.Monitor]::Exit($global:OperationRegistryLock)
    }
}

function Add-OperationError ([string]$OperationId, [string]$ErrorMessage) {
    [System.Threading.Monitor]::Enter($global:OperationRegistryLock)
    try {
        if (-not $global:OperationsRegistry.ContainsKey($OperationId)) { return }
        $op = $global:OperationsRegistry[$OperationId]
        $clean = Protect-SensitiveData $ErrorMessage
        $op.errors += $clean
    } finally {
        [System.Threading.Monitor]::Exit($global:OperationRegistryLock)
    }
}

function Request-OperationCancellation ([string]$OperationId) {
    [System.Threading.Monitor]::Enter($global:OperationRegistryLock)
    try {
        if (-not $global:OperationsRegistry.ContainsKey($OperationId)) {
            return @{ Success = $false; Error = "Operation not found." }
        }

        $op = $global:OperationsRegistry[$OperationId]
        if ($global:TERMINAL_OPERATION_STATES -contains $op.status) {
            return @{ Success = $false; Error = "Operation is already in terminal state: $($op.status)" }
        }

        $op.cancelRequested = $true
        if ($op.status -in @("QUEUED", "RUNNING")) {
            $op.status = "CANCELLATION_REQUESTED"
            $op.currentStage = "Cancellation Requested"
        }

        return @{ Success = $true; Status = $op.status; Message = "Cancellation signal dispatched." }
    } finally {
        [System.Threading.Monitor]::Exit($global:OperationRegistryLock)
    }
}

function Test-OperationCancellationRequested ([string]$OperationId) {
    [System.Threading.Monitor]::Enter($global:OperationRegistryLock)
    try {
        if (-not $global:OperationsRegistry.ContainsKey($OperationId)) { return $false }
        return ($global:OperationsRegistry[$OperationId].cancelRequested -eq $true)
    } finally {
        [System.Threading.Monitor]::Exit($global:OperationRegistryLock)
    }
}

function Get-OperationSnapshot ([string]$OperationId) {
    $lockTaken = $false
    try {
        [System.Threading.Monitor]::TryEnter($global:OperationRegistryLock, $global:OPERATION_LOCK_TIMEOUT_MS, [ref]$lockTaken) | Out-Null
        if (-not $lockTaken) {
            return @{
                id              = "$OperationId"
                operationId     = "$OperationId"
                type            = "UNKNOWN"
                status          = "RUNNING"
                createdAt       = $null
                startedAt       = $null
                finishedAt      = $null
                directory       = ""
                totalFiles      = 0
                processedFiles  = 0
                updatedFiles    = 0
                skippedFiles    = 0
                failedFiles     = 0
                currentFile     = "Durum kilidi bekleniyor"
                currentStage    = "State Lock Busy"
                progressPercent = 0
                warnings        = @("Operasyon durum kilidi $($global:OPERATION_LOCK_TIMEOUT_MS) ms icinde alinamadi. Worker buyuk bir checkpoint yaziyor veya kilitli kaldi.")
                errors          = @()
                backupDirectory = ""
                batchStatus     = "PENDING"
                cancelRequested = $false
                options         = @{}
                rules           = @()
                logs            = @()
                scannedFiles    = @()
                detectedIPs     = @()
                resultData      = $null
                auditLogPath    = Get-AuditLogFilePath -OperationId $OperationId
            }
        }
        if (-not $global:OperationsRegistry.ContainsKey($OperationId)) { return $null }
        $src = $global:OperationsRegistry[$OperationId]

        # Produce a deep clone hashtable containing only serializable primitive objects
        $snapshot = @{
            id              = "$($src.operationId)"
            operationId     = "$($src.operationId)"
            type            = "$($src.type)"
            status          = "$($src.status)"
            createdAt       = "$($src.createdAt)"
            startedAt       = if ($src.startedAt) { "$($src.startedAt)" } else { $null }
            finishedAt      = if ($src.finishedAt) { "$($src.finishedAt)" } else { $null }
            directory       = "$($src.directory)"
            totalFiles      = [int]$src.totalFiles
            processedFiles  = [int]$src.processedFiles
            updatedFiles    = [int]$src.updatedFiles
            skippedFiles    = [int]$src.skippedFiles
            failedFiles     = [int]$src.failedFiles
            currentFile     = "$($src.currentFile)"
            currentStage    = "$($src.currentStage)"
            progressPercent = [int]$src.progressPercent
            warnings        = @($src.warnings)
            errors          = @($src.errors)
            backupDirectory = "$($src.backupDirectory)"
            batchStatus     = "$($src.batchStatus)"
            cancelRequested = [bool]$src.cancelRequested
            options         = Protect-SensitiveData $src.options
            rules           = Protect-SensitiveData $src.rules
            logs            = @($src.logs)
            scannedFiles    = if ($src.scannedFiles) { @($src.scannedFiles) } else { @() }
            detectedIPs     = if ($src.detectedIPs) { @($src.detectedIPs) } else { @() }
            resultData      = Protect-SensitiveData $src.resultData
            auditLogPath    = if ($src.auditLogPath) { "$($src.auditLogPath)" } else { Get-AuditLogFilePath -OperationId $OperationId }
        }
        return $snapshot
    } finally {
        if ($lockTaken) {
            [System.Threading.Monitor]::Exit($global:OperationRegistryLock)
        }
    }
}

function Append-OperationScannedFile {
    param(
        [Parameter(Mandatory=$true)][string]$OperationId,
        [Parameter(Mandatory=$true)][hashtable]$FileDetail,
        [hashtable]$IpSummary = @{}
    )

    [System.Threading.Monitor]::Enter($global:OperationRegistryLock)
    try {
        if (-not $global:OperationsRegistry.ContainsKey($OperationId)) { return }
        $op = $global:OperationsRegistry[$OperationId]
        if (-not $op.ContainsKey("scannedFiles") -or $null -eq $op["scannedFiles"]) {
            $op["scannedFiles"] = @()
        }
        $op["scannedFiles"] += $FileDetail

        if ($IpSummary -and $IpSummary.Count -gt 0) {
            $ipList = @()
            foreach ($k in $IpSummary.Keys) {
                $ipList += @{ ip = "$k"; count = [int]$IpSummary[$k] }
            }
            $op["detectedIPs"] = $ipList
        }
        Write-RawAuditLogLine -OperationId $OperationId -Level "INFO" -EventName "SCAN_FILE_RESULT_APPENDED" -Stage "File Result" -File "$($FileDetail.filePath)" -Message "Scanned file result appended" -Data @{ fileName = "$($FileDetail.fileName)"; status = "$($FileDetail.status)"; foundIpCount = @($FileDetail.foundIPs).Count }
    } finally {
        [System.Threading.Monitor]::Exit($global:OperationRegistryLock)
    }
}

function Get-ActiveOperation () {
    [System.Threading.Monitor]::Enter($global:OperationRegistryLock)
    try {
        foreach ($k in $global:OperationsRegistry.Keys) {
            $op = $global:OperationsRegistry[$k]
            if ($op -and ($global:TERMINAL_OPERATION_STATES -notcontains $op.status)) {
                return Get-OperationSnapshot $k
            }
        }
        return $null
    } finally {
        [System.Threading.Monitor]::Exit($global:OperationRegistryLock)
    }
}

# ------------------------------------------------------------------------------
# 4. Progress Model (Aşamalı Ağırlıklı İlerleme Hesabı)
# ------------------------------------------------------------------------------
function Calculate-WeightedProgress {
    param(
        [int]$TotalFiles,
        [int]$CompletedFiles,
        [string]$CurrentFileStage = "Preflight"
    )

    if ($TotalFiles -le 0) { return 0 }
    if ($CompletedFiles -ge $TotalFiles) { return 100 }

    # Overall weights
    $preflightWeight = 5.0
    $backupWeight = 5.0
    $cleanupWeight = 5.0
    $filesTotalWeight = 85.0

    # Per-file cumulative stage weights (strictly monotonic within a file)
    $stageWeights = @{
        "Preflight"    = 0.05
        "Staging"      = 0.15
        "Excel Update" = 0.65
        "Validation"   = 0.85
        "Commit"       = 0.95
        "Cleanup"      = 1.00
    }

    $stageFactor = if ($stageWeights.ContainsKey($CurrentFileStage)) { $stageWeights[$CurrentFileStage] } else { 0.10 }

    $perFileWeight = $filesTotalWeight / [double]$TotalFiles
    $completedWeight = [double]$CompletedFiles * $perFileWeight
    $currentFileProgress = $stageFactor * $perFileWeight

    $totalProgress = $preflightWeight + $backupWeight + $completedWeight + $currentFileProgress
    $clamped = [Math]::Floor([Math]::Max(0.0, [Math]::Min(99.0, $totalProgress)))

    return [int]$clamped
}

# ------------------------------------------------------------------------------
# 5. History Atomic Persistence & Corrupt Recovery
# ------------------------------------------------------------------------------
function Save-OperationHistory {
    param(
        [Parameter(Mandatory=$true)][hashtable]$OperationSummary,
        [string]$HistoryFilePath = "",
        [string]$RepoRoot = ""
    )

    if ([string]::IsNullOrWhiteSpace($HistoryFilePath)) {
        $base = if (-not [string]::IsNullOrWhiteSpace($RepoRoot)) { $RepoRoot } else { (Get-Location).Path }
        $HistoryFilePath = Join-Path $base "logs\history.json"
    }

    $histDir = Split-Path -Parent $HistoryFilePath
    if (-not (Test-Path $histDir)) {
        New-Item -ItemType Directory -Path $histDir -Force | Out-Null
    }

    [System.Threading.Monitor]::Enter($global:HistoryFileLock)
    try {
        $existing = @()
        if (Test-Path $HistoryFilePath) {
            try {
                $rawText = [System.IO.File]::ReadAllText($HistoryFilePath, [System.Text.Encoding]::UTF8)
                if (-not [string]::IsNullOrWhiteSpace($rawText)) {
                    $parsed = $rawText | ConvertFrom-Json
                    if ($parsed -is [System.Collections.IEnumerable]) {
                        $existing = @($parsed)
                    } else {
                        $existing = @($parsed)
                    }
                }
            } catch {
                # Corrupt history handling: preserve corrupted file with timestamp
                $ts = (Get-Date).ToString("yyyyMMdd_HHmmss")
                $corruptFile = Join-Path $histDir "history.corrupt.$ts.json"
                try { Move-Item -LiteralPath $HistoryFilePath -Destination $corruptFile -Force } catch { }
                $existing = @()
            }
        }

        # Mask sensitive credentials before persisting
        $safeSummary = Protect-SensitiveData $OperationSummary

        # Append or replace if exists
        $opId = $safeSummary.operationId
        if ([string]::IsNullOrWhiteSpace($opId)) { $opId = $safeSummary.id }
        $foundIdx = -1
        for ($i = 0; $i -lt $existing.Count; $i++) {
            $item = $existing[$i]
            $itemOpId = if ($item.operationId) { $item.operationId } else { $item.id }
            if ($itemOpId -eq $opId) {
                $foundIdx = $i
                break
            }
        }

        if ($foundIdx -ge 0) {
            $existing[$foundIdx] = $safeSummary
        } else {
            $existing += $safeSummary
        }

        # Atomic commit via .tmp
        $tmpFile = "$HistoryFilePath.tmp"
        $json = $existing | ConvertTo-Json -Depth 10
        if ($null -eq $json) {
            $json = "[]"
        }
        [System.IO.File]::WriteAllText($tmpFile, $json, [System.Text.Encoding]::UTF8)

        # Validate tmp file before commit
        $valTest = [System.IO.File]::ReadAllText($tmpFile, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
        if ($null -eq $valTest) {
            throw "Failed to validate serialized history JSON."
        }

        if (Test-Path $HistoryFilePath) {
            Remove-Item -LiteralPath $HistoryFilePath -Force
        }
        Move-Item -LiteralPath $tmpFile -Destination $HistoryFilePath -Force

        return $true
    } catch {
        return $false
    } finally {
        [System.Threading.Monitor]::Exit($global:HistoryFileLock)
    }
}

function Get-OperationHistory {
    param(
        [string]$HistoryFilePath = "",
        [string]$RepoRoot = ""
    )

    if ([string]::IsNullOrWhiteSpace($HistoryFilePath)) {
        $base = if (-not [string]::IsNullOrWhiteSpace($RepoRoot)) { $RepoRoot } else { (Get-Location).Path }
        $HistoryFilePath = Join-Path $base "logs\history.json"
    }

    if (-not (Test-Path $HistoryFilePath)) { return ,@() }

    [System.Threading.Monitor]::Enter($global:HistoryFileLock)
    try {
        $rawText = [System.IO.File]::ReadAllText($HistoryFilePath, [System.Text.Encoding]::UTF8)
        if ([string]::IsNullOrWhiteSpace($rawText)) { return ,@() }
        $parsed = $rawText | ConvertFrom-Json
        return ,@($parsed)
    } catch {
        $hDir = Split-Path -Parent $HistoryFilePath
        $ts = (Get-Date).ToString("yyyyMMdd_HHmmss")
        $corruptFile = Join-Path $hDir "history.corrupt.$ts.json"
        try { Move-Item -LiteralPath $HistoryFilePath -Destination $corruptFile -Force } catch { }
        return ,@()
    } finally {
        [System.Threading.Monitor]::Exit($global:HistoryFileLock)
    }
}

function Recover-StaleOperations {
    param(
        [string]$RepoRoot = "",
        [string]$HistoryFilePath = ""
    )

    if ([string]::IsNullOrWhiteSpace($HistoryFilePath)) {
        $base = if (-not [string]::IsNullOrWhiteSpace($RepoRoot)) { $RepoRoot } else { (Get-Location).Path }
        $HistoryFilePath = Join-Path $base "logs\history.json"
    }

    if (-not (Test-Path -LiteralPath $HistoryFilePath)) { return 0 }

    $recovered = 0
    [System.Threading.Monitor]::Enter($global:HistoryFileLock)
    try {
        $rawText = [System.IO.File]::ReadAllText($HistoryFilePath, [System.Text.Encoding]::UTF8)
        if ([string]::IsNullOrWhiteSpace($rawText)) { return 0 }
        $list = $rawText | ConvertFrom-Json
        $updatedList = @()
        $nowIso = (Get-Date).ToUniversalTime().ToString("o")

        foreach ($item in @($list)) {
            $ht = @{}
            foreach ($prop in $item.PSObject.Properties) {
                $ht[$prop.Name] = $prop.Value
            }

            $nonTerminal = @("RUNNING", "COMMITTING", "ROLLING_BACK", "CANCELLATION_REQUESTED", "QUEUED")
            if ($nonTerminal -contains $ht.status) {
                $ht.status = "STALE_INTERRUPTED"
                $ht.finishedAt = $nowIso
                $ht.currentStage = "Interrupted by Server Restart"
                $errs = if ($ht.errors) { @($ht.errors) } else { @() }
                $errs += "Sunucu beklenmedik şekilde kapandığı için işlem kesintiye uğradı (STALE_INTERRUPTED)."
                $ht.errors = $errs
                $recovered++
            }
            $updatedList += $ht
        }

        if ($recovered -gt 0) {
            $tmpFile = "$HistoryFilePath.tmp"
            $json = $updatedList | ConvertTo-Json -Depth 10
            [System.IO.File]::WriteAllText($tmpFile, $json, [System.Text.Encoding]::UTF8)
            Move-Item -LiteralPath $tmpFile -Destination $HistoryFilePath -Force
        }
    } catch { }
    finally {
        [System.Threading.Monitor]::Exit($global:HistoryFileLock)
    }

    return $recovered
}

# ------------------------------------------------------------------------------
# 6. Structured JSONL Audit Logging (Thread-Safe)
# ------------------------------------------------------------------------------
function Write-AuditLogEvent {
    param(
        [string]$LogDirectory = "",
        [Parameter(Mandatory=$true)][string]$OperationId,
        [Parameter(Mandatory=$true)][string]$Level,
        [Parameter(Mandatory=$true)][string]$EventName,
        [string]$File = "",
        [string]$Stage = "",
        [string]$Message = "",
        [hashtable]$Data = @{}
    )

    if ([string]::IsNullOrWhiteSpace($LogDirectory)) {
        $LogDirectory = if (-not [string]::IsNullOrWhiteSpace($global:AuditLogDirectory)) { $global:AuditLogDirectory } else { Join-Path (Get-Location) "logs" }
    }

    Write-RawAuditLogLine -OperationId $OperationId -Level $Level -EventName $EventName -File $File -Stage $Stage -Message $Message -Data $Data -LogDirectory $LogDirectory
}

# ------------------------------------------------------------------------------
# 7. Real PowerShell STA Runspace Management
# ------------------------------------------------------------------------------
function New-StaRunspace () {
    $runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $runspace.ApartmentState = "STA"
    $runspace.ThreadOptions = "ReuseThread"
    $runspace.Open()

    return @{
        Runspace = $runspace
    }
}

function Close-StaRunspace ($StaContext) {
    if ($StaContext -and $StaContext.Runspace) {
        try {
            $StaContext.Runspace.Close()
            $StaContext.Runspace.Dispose()
        } catch { }
        $StaContext.Runspace = $null
    }
}

# ------------------------------------------------------------------------------
# 8. Diagnostics System Health Check (Side-Effect Free)
# ------------------------------------------------------------------------------
function Invoke-SystemDiagnostics ([string]$RepoRoot = "") {
    if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
        $RepoRoot = (Get-Location).Path
    }

    $excelInstalled = $false
    $excelVer = "Not Detected"
    
    # 1. Non-intrusive Excel COM check
    try {
        $excelType = [System.Type]::GetTypeFromProgID("Excel.Application")
        if ($excelType) {
            $excelInstalled = $true
            # Read version cleanly without keeping any processes open
            $iso = New-IsolatedExcelInstance
            if ($iso -and $iso.Excel) {
                $excelVer = "$($iso.Excel.Version)"
                Close-IsolatedExcelInstance $iso
            }
        }
    } catch { }

    # 2. Directory writability checks
    $workingDirWritable = $false
    $testFile = Join-Path $RepoRoot ".__writetest_$([System.Guid]::NewGuid().ToString('N')).tmp"
    try {
        [System.IO.File]::WriteAllText($testFile, "test")
        $workingDirWritable = $true
        Remove-Item -LiteralPath $testFile -Force
    } catch { }

    $backupBase = Join-Path $RepoRoot "_ExcelUpdater_Backups"
    $backupDirWritable = $false
    try {
        if (-not (Test-Path $backupBase)) {
            New-Item -ItemType Directory -Path $backupBase -Force | Out-Null
        }
        $backupTest = Join-Path $backupBase ".__writetest_$([System.Guid]::NewGuid().ToString('N')).tmp"
        [System.IO.File]::WriteAllText($backupTest, "test")
        $backupDirWritable = $true
        Remove-Item -LiteralPath $backupTest -Force
    } catch { }

    # 3. Active Excel PIDs check
    $activePids = @()
    try {
        $procs = Get-Process excel -ErrorAction SilentlyContinue
        if ($procs) {
            $activePids = @($procs | Select-Object -ExpandProperty Id)
            $procs | ForEach-Object { $_.Dispose() }
        }
    } catch { }

    return @{
        Success = $true
        ExcelComAvailable = $excelInstalled
        ExcelVersion = $excelVer
        PowerShellVersion = "$($PSVersionTable.PSVersion)"
        CLRVersion = "$([System.Environment]::Version)"
        WorkingDirWritable = $workingDirWritable
        BackupDirWritable = $backupDirWritable
        ActiveExcelPids = $activePids
        ActiveOperationsCount = $global:OperationsRegistry.Count
    }
}

# ------------------------------------------------------------------------------
# 9. Asynchronous STA Operation Worker Dispatcher
# ------------------------------------------------------------------------------
function Start-AsyncOperationWorker {
    param(
        [Parameter(Mandatory=$true)][string]$OperationId,
        [Parameter(Mandatory=$true)][string]$RepoRoot
    )

    $op = Get-OperationSnapshot $OperationId
    if (-not $op) { return $false }

    $enginePath = Join-Path $RepoRoot "engine\excel_engine.ps1"
    $managerPath = Join-Path $RepoRoot "engine\operation_manager.ps1"

    $sta = New-StaRunspace
    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $sta.Runspace

    $scriptBlock = {
        param($opId, $repoRoot, $engineScript, $mgrScript)
        
        $ErrorActionPreference = "Stop"
        $result = $null
        try {
            Set-Location -LiteralPath $repoRoot

            # 1. Explicitly bootstrap dependencies in the new STA runspace (Rule 2)
            if (Test-Path $mgrScript) { . $mgrScript }
            if (Test-Path $engineScript) { . $engineScript }

            # 2. Transition state to RUNNING
            Set-OperationStatus $opId "RUNNING" "Initializing" | Out-Null
            try {
                Write-AuditLogEvent -OperationId $opId -Level "INFO" -EventName "OPERATION_STARTED" -Message "Worker started in STA runspace"
            } catch { }

            $type = ""
            $dir = ""
            $rules = @()
            $options = @{}

            [System.Threading.Monitor]::Enter($global:OperationRegistryLock)
            try {
                if ($global:OperationsRegistry.ContainsKey($opId)) {
                    $rawOp = $global:OperationsRegistry[$opId]
                    $type = "$($rawOp.type)"
                    $dir = "$($rawOp.directory)"
                    $rules = @($rawOp.rules)
                    $options = if ($rawOp.options) { @($rawOp.options)[0] } else { @{} }
                }
            } finally {
                [System.Threading.Monitor]::Exit($global:OperationRegistryLock)
            }
            switch ($type) {
                "UPDATE" {
                    Write-AuditLogEvent -OperationId $opId -Level "INFO" -EventName "WORKER_DISPATCH" -Stage "UPDATE" -Message "Dispatching update operation" -Data @{ directory = $dir }
                    $result = Update-ExcelDirectory -DirectoryPath $dir -Rules $rules -Options $options -OperationId $opId
                }
                "PREVIEW" {
                    Write-AuditLogEvent -OperationId $opId -Level "INFO" -EventName "WORKER_DISPATCH" -Stage "PREVIEW" -Message "Dispatching preview operation" -Data @{ directory = $dir }
                    $result = Preview-ExcelDirectory -DirectoryPath $dir -Rules $rules -Options $options -OperationId $opId
                }
                "SCAN" {
                    Write-AuditLogEvent -OperationId $opId -Level "INFO" -EventName "WORKER_DISPATCH" -Stage "SCAN" -Message "Dispatching scan operation" -Data @{ directory = $dir }
                    $result = Scan-ExcelDirectory -DirectoryPath $dir -OperationId $opId
                }
                "RESTORE" {
                    $backupDir = if ($options -and $options.backupDir) { $options.backupDir } else { "" }
                    Write-AuditLogEvent -OperationId $opId -Level "INFO" -EventName "WORKER_DISPATCH" -Stage "RESTORE" -Message "Dispatching restore operation" -Data @{ directory = $dir; backupDirectory = $backupDir }
                    $result = Restore-VerifiedBackup -BackupDirectory $backupDir -TargetDirectory $dir -OperationId $opId
                }
                default {
                    throw "Unsupported operation type: $type"
                }
            }

            # If not already terminal, update status from result
            $currentOp = Get-OperationSnapshot $opId
            if ($global:TERMINAL_OPERATION_STATES -notcontains $currentOp.status) {
                if ($result.success -or $result.Success) {
                    Set-OperationStatus $opId "COMPLETED" "Done" | Out-Null
                    $resultSummary = @{
                        success = $true
                        totalFiles = if ($result.totalFiles) { $result.totalFiles } else { $result.totalFilesProcessed }
                        updatedFilesCount = $result.updatedFilesCount
                        prospectiveUpdatedFilesCount = $result.prospectiveUpdatedFilesCount
                        totalProspectiveReplacements = $result.totalProspectiveReplacements
                        detectedIpCount = if ($result.detectedIPs) { @($result.detectedIPs).Count } else { 0 }
                    }
                    Write-AuditLogEvent -OperationId $opId -Level "INFO" -EventName "OPERATION_COMPLETED" -Message "Operation finished successfully" -Data $resultSummary
                } elseif ($result.batchStatus -eq "CANCELLED") {
                    Set-OperationStatus $opId "CANCELLED" "Cancelled by user" | Out-Null
                    Write-AuditLogEvent -OperationId $opId -Level "WARNING" -EventName "OPERATION_CANCELLED" -Message "Operation cancelled by user"
                } elseif ($result.batchStatus -eq "ROLLBACK_PARTIAL_FAILURE") {
                    Set-OperationStatus $opId "ROLLBACK_PARTIAL_FAILURE" "Rollback partial failure" | Out-Null
                    Write-AuditLogEvent -OperationId $opId -Level "CRITICAL" -EventName "ROLLBACK_PARTIAL_FAILURE" -Message "Rollback partially failed" -Data @{ error = $result.error }
                } else {
                    Set-OperationStatus $opId "FAILED" "Failed" | Out-Null
                    Add-OperationError $opId "$($result.error)"
                    Write-AuditLogEvent -OperationId $opId -Level "ERROR" -EventName "OPERATION_FAILED" -Message "Operation failed: $($result.error)" -Data @{ error = $result.error }
                }
            }

            # Store final result data (sanitized)
            [System.Threading.Monitor]::Enter($global:OperationRegistryLock)
            try {
                if ($global:OperationsRegistry.ContainsKey($opId)) {
                    $global:OperationsRegistry[$opId].resultData = Protect-SensitiveData $result
                    if ($result.logs) { $global:OperationsRegistry[$opId].logs = @($result.logs) }
                    if ($result.backupDirectory) { $global:OperationsRegistry[$opId].backupDirectory = "$($result.backupDirectory)" }
                    if ($result.updatedFilesCount) { $global:OperationsRegistry[$opId].updatedFiles = [int]$result.updatedFilesCount }
                    if ($result.totalFilesProcessed) { $global:OperationsRegistry[$opId].processedFiles = [int]$result.totalFilesProcessed }
                }
            } finally {
                [System.Threading.Monitor]::Exit($global:OperationRegistryLock)
            }

        } catch {
            Set-OperationStatus $opId "FAILED" "Error: $_" | Out-Null
            Add-OperationError $opId "$_"
            Write-AuditLogEvent -OperationId $opId -Level "ERROR" -EventName "OPERATION_EXCEPTION" -Message "$_"
        } finally {
            # Save to persistent history
            $finalSnap = Get-OperationSnapshot $opId
            if ($finalSnap) {
                Save-OperationHistory -OperationSummary $finalSnap | Out-Null
            }
        }
    }

    $ps.AddScript($scriptBlock) | Out-Null
    $ps.AddArgument($OperationId) | Out-Null
    $ps.AddArgument($RepoRoot) | Out-Null
    $ps.AddArgument($enginePath) | Out-Null
    $ps.AddArgument($managerPath) | Out-Null

    $asyncState = @{
        PowerShell = $ps
        RunspaceContext = $sta
        OperationId = $OperationId
    }

    $callback = [System.AsyncCallback]{
        param($ar)
        try {
            $state = $ar.AsyncState
            $p = $state.PowerShell
            $r = $state.RunspaceContext
            try { $p.EndInvoke($ar) | Out-Null } catch { }
            $p.Dispose()
            Close-StaRunspace $r
        } catch { }
    }

    try {
        Write-AuditLogEvent -OperationId $OperationId -Level "INFO" -EventName "WORKER_BEGIN_INVOKE" -Message "Starting asynchronous worker runspace" -Data @{ repoRoot = $RepoRoot }
        $inputCol = [System.Management.Automation.PSDataCollection[psobject]]::new()
        $null = $ps.BeginInvoke($inputCol, $null, $callback, $asyncState)
        return $true
    } catch {
        try {
            Set-OperationStatus $OperationId "FAILED" "Worker start failed" | Out-Null
            Add-OperationError $OperationId "Arka plan worker baslatilamadi: $_"
            Write-AuditLogEvent -OperationId $OperationId -Level "ERROR" -EventName "WORKER_BEGIN_INVOKE_FAILED" -Message "$_"
            Save-OperationHistory -OperationSummary (Get-OperationSnapshot $OperationId) -RepoRoot $RepoRoot | Out-Null
        } catch { }
        try { $ps.Dispose() } catch { }
        try { Close-StaRunspace $sta } catch { }
        return $false
    }
}
