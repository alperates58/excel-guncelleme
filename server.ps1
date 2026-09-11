# ==============================================================================
# Excel SQL Server Connection & Macro Bulk Updater - Server & API Engine
# ==============================================================================
param(
    [int]$Port = 3000
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Load Excel Engine functions
$engineScript = Join-Path $PSScriptRoot "engine\excel_engine.ps1"
if (Test-Path $engineScript) {
    . $engineScript
} else {
    Write-Error "Excel engine script not found at $engineScript"
    exit 1
}

# Load Operation Manager functions
$managerScript = Join-Path $PSScriptRoot "engine\operation_manager.ps1"
if (Test-Path $managerScript) {
    . $managerScript
    try {
        Recover-StaleOperations -RepoRoot $PSScriptRoot | Out-Null
        Write-AuditLogEvent -OperationId "SERVER_STARTUP" -Level "INFO" -EventName "SERVER_INITIALIZED" -Message "Excel Bulk Updater server started on port $Port"
    } catch { }
} else {
    Write-Error "Operation manager script not found at $managerScript"
    exit 1
}

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$Port/")
$listener.Prefixes.Add("http://127.0.0.1:$Port/")

try {
    $listener.Start()
} catch {
    Write-Host "Failed to start server on port $Port. Error: $_"
    exit 1
}

Write-Host "=================================================================" -ForegroundColor Green
Write-Host " Excel Bulk Updater Server is running on http://127.0.0.1:$Port/" -ForegroundColor Cyan
Write-Host " Open your browser to: http://127.0.0.1:$Port/" -ForegroundColor Yellow
Write-Host "=================================================================" -ForegroundColor Green

try {
    Start-Process "http://127.0.0.1:$Port/"
} catch { }

$publicDir = Join-Path $PSScriptRoot "public"

function Get-ContentType ($filePath) {
    $ext = [System.IO.Path]::GetExtension($filePath).ToLower()
    switch ($ext) {
        ".html" { return "text/html; charset=utf-8" }
        ".css"  { return "text/css; charset=utf-8" }
        ".js"   { return "application/javascript; charset=utf-8" }
        ".json" { return "application/json; charset=utf-8" }
        ".png"  { return "image/png" }
        ".jpg"  { return "image/jpeg" }
        ".svg"  { return "image/svg+xml" }
        ".ico"  { return "image/x-icon" }
        default { return "application/octet-stream" }
    }
}

function Send-JsonResponse ($response, $object, [int]$statusCode = 200) {
    $json = $object | ConvertTo-Json -Depth 10 -Compress
    $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
    $response.StatusCode = $statusCode
    $response.ContentType = "application/json; charset=utf-8"
    $response.ContentLength64 = $buffer.Length
    $response.OutputStream.Write($buffer, 0, $buffer.Length)
    $response.OutputStream.Close()
}

function Read-RequestBody ($request) {
    $reader = New-Object System.IO.StreamReader($request.InputStream, $request.ContentEncoding)
    $body = $reader.ReadToEnd()
    $reader.Close()
    if (-not [string]::IsNullOrWhiteSpace($body)) {
        return $body | ConvertFrom-Json
    }
    return $null
}

while ($listener.IsListening) {
    try {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response

        $path = $request.Url.AbsolutePath

        # Enforce local requests only
        if (-not $request.IsLocal) {
            $response.StatusCode = 403
            $response.OutputStream.Close()
            continue
        }

        # Restrict CORS to local origins only (eliminate wildcard *)
        $origin = $request.Headers["Origin"]
        $allowedOrigins = @("http://localhost:$Port", "http://127.0.0.1:$Port")
        if ($origin -and ($allowedOrigins -contains $origin)) {
            $response.AddHeader("Access-Control-Allow-Origin", $origin)
            $response.AddHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
            $response.AddHeader("Access-Control-Allow-Headers", "Content-Type")
        }

        if ($request.HttpMethod -eq "OPTIONS") {
            $response.StatusCode = 200
            $response.OutputStream.Close()
            continue
        }

        # ----------------------------------------------------------------------
        # API ENDPOINTS
        # ----------------------------------------------------------------------
        if ($path -eq "/api/status") {
            $status = @{
                status = "ok"
                port = $Port
                workspace = (Get-Location).Path
                userDesktop = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
                userDownloads = Join-Path $env:USERPROFILE "Downloads"
                excelInstalled = $true
            }
            Send-JsonResponse $response $status
            continue
        }

        if ($path -eq "/api/progress") {
            if ($global:ActiveOperationId) {
                $snap = Get-OperationSnapshot $global:ActiveOperationId
                if ($snap) {
                    $compatProgress = @{
                        active = $true
                        operationId = $snap.id
                        type = $snap.type.ToLower()
                        status = $snap.status
                        currentFile = $snap.currentFile
                        current = $snap.processedFiles
                        total = $snap.totalFiles
                        percent = $snap.progressPercent
                        stage = $snap.currentStage
                        canCancel = ($global:CANCELABLE_OPERATION_STATES -contains $snap.status)
                    }
                    Send-JsonResponse $response $compatProgress
                    continue
                }
            }
            Send-JsonResponse $response $global:ProgressState
            continue
        }

        # ----------------------------------------------------------------------
        # OPERATION MANAGEMENT & ASYNC WORKER ENDPOINTS
        # ----------------------------------------------------------------------
        if ($path -eq "/api/operations/active" -and $request.HttpMethod -eq "GET") {
            if ($global:ActiveOperationId) {
                $snap = Get-OperationSnapshot $global:ActiveOperationId
                Send-JsonResponse $response @{ active = $true; operation = $snap }
            } else {
                Send-JsonResponse $response @{ active = $false }
            }
            continue
        }

        if ($path -match "^/api/operations/([^/]+)/cancel$" -and $request.HttpMethod -eq "POST") {
            $targetOpId = $Matches[1]
            $cancelResult = Request-OperationCancellation -OperationId $targetOpId
            Send-JsonResponse $response $cancelResult
            continue
        }

        if ($path -match "^/api/operations/([^/]+)$" -and $request.HttpMethod -eq "GET") {
            $targetOpId = $Matches[1]
            $snap = Get-OperationSnapshot $targetOpId
            if ($snap) {
                Send-JsonResponse $response @{ success = $true; operation = $snap }
            } else {
                $histList = Get-OperationHistory -RepoRoot $PSScriptRoot
                $found = $histList | Where-Object { $_.id -eq $targetOpId } | Select-Object -First 1
                if ($found) {
                    Send-JsonResponse $response @{ success = $true; operation = $found }
                } else {
                    Send-JsonResponse $response @{ success = $false; error = "İşlem bulunamadı: $targetOpId" } 404
                }
            }
            continue
        }

        if ($path -eq "/api/history" -and $request.HttpMethod -eq "GET") {
            $histList = Get-OperationHistory -RepoRoot $PSScriptRoot
            Send-JsonResponse $response @{ success = $true; history = @($histList) }
            continue
        }

        if ($path -eq "/api/diagnostics" -and $request.HttpMethod -eq "GET") {
            $diag = Invoke-SystemDiagnostics -RepoRoot $PSScriptRoot
            Send-JsonResponse $response $diag
            continue
        }

        if ($path -eq "/api/backups" -and $request.HttpMethod -eq "GET") {
            $dirQuery = $request.QueryString["dir"]
            $targetDir = if (-not [string]::IsNullOrWhiteSpace($dirQuery)) { $dirQuery } else { (Get-Location).Path }

            $backupsList = @()
            if (Test-Path -LiteralPath $targetDir) {
                $backupRootDir = Join-Path $targetDir "_ExcelUpdater_Backups"
                if (Test-Path -LiteralPath $backupRootDir) {
                    $manifests = Get-ChildItem -LiteralPath $backupRootDir -Filter "manifest.json" -Recurse -File -ErrorAction SilentlyContinue
                    foreach ($m in $manifests) {
                        try {
                            $jsonContent = Get-Content -LiteralPath $m.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
                            $bFolder = $m.Directory
                            $backupsList += @{
                                timestamp = $jsonContent.timestamp
                                backupDirectory = $bFolder.FullName
                                folderName = $bFolder.Name
                                fileCount = if ($jsonContent.files) { @($jsonContent.files).Count } else { 0 }
                                operation = $jsonContent.operation
                            }
                        } catch { }
                    }
                }
            }
            $sorted = $backupsList | Sort-Object { $_.timestamp } -Descending
            Send-JsonResponse $response @{ success = $true; backups = @($sorted) }
            continue
        }

        if ($path -eq "/api/restore" -and $request.HttpMethod -eq "POST") {
            $data = Read-RequestBody $request
            $targetDir = if ($data -and $data.directory) { $data.directory } else { (Get-Location).Path }
            $backupDir = if ($data -and $data.backupDir) { $data.backupDir } else { "" }

            if ([string]::IsNullOrWhiteSpace($backupDir) -or -not (Test-Path -LiteralPath $backupDir)) {
                Send-JsonResponse $response @{ success = $false; error = "Geçersiz veya bulunamayan yedek klasörü: $backupDir" } 400
                continue
            }

            $manifestPath = Join-Path $backupDir "manifest.json"
            if (-not (Test-Path -LiteralPath $manifestPath)) {
                Send-JsonResponse $response @{ success = $false; error = "Yedek klasöründe manifest.json bulunamadı: $backupDir" } 400
                continue
            }

            $reg = Register-Operation -Type "RESTORE" -Directory $targetDir -Options @{ backupDir = $backupDir }
            if (-not $reg.success) {
                Send-JsonResponse $response $reg 409
                continue
            }

            Start-AsyncOperationWorker -OperationId $reg.operationId -RepoRoot $PSScriptRoot | Out-Null
            Send-JsonResponse $response @{
                success = $true
                operationId = $reg.operationId
                status = "QUEUED"
                message = "Restore operation registered and queued"
            } 202
            continue
        }

        if ($path -eq "/api/list-folders" -and $request.HttpMethod -eq "POST") {
            $data = Read-RequestBody $request
            $targetDir = if ($data -and $data.directory -and (Test-Path $data.directory)) { $data.directory } else { (Get-Location).Path }
            
            try {
                $dirObj = Get-Item $targetDir
                $parentDir = if ($dirObj.Parent) { $dirObj.Parent.FullName } else { "" }
                
                $subDirs = Get-ChildItem -Path $targetDir -Directory -ErrorAction SilentlyContinue | 
                    Select-Object Name, FullName | 
                    ForEach-Object { @{ name = $_.Name; path = $_.FullName } }

                $drives = @()
                if ([string]::IsNullOrWhiteSpace($parentDir)) {
                    $drives = Get-PSDrive -PSProvider FileSystem | ForEach-Object { @{ name = "$($_.Name):\"; path = "$($_.Root)" } }
                }

                $res = @{
                    success = $true
                    currentDir = $dirObj.FullName
                    parentDir = $parentDir
                    subFolders = if ($subDirs) { $subDirs } else { @() }
                    drives = $drives
                }
                Send-JsonResponse $response $res
            } catch {
                Send-JsonResponse $response @{ success = $false; error = "Klasor okunamadi: $_" } 400
            }
            continue
        }

        if ($path -eq "/api/open-folder" -and $request.HttpMethod -eq "POST") {
            $data = Read-RequestBody $request
            $rawDir = if ($data -and $data.directory) { $data.directory } else { (Get-Location).Path }
            
            if ([string]::IsNullOrWhiteSpace($rawDir) -or -not (Test-Path -LiteralPath $rawDir -PathType Container)) {
                Send-JsonResponse $response @{ success = $false; error = "Klasor bulunamadi veya gecersiz: $rawDir" } 400
                continue
            }

            $resolvedDir = (Resolve-Path -LiteralPath $rawDir).Path
            Send-JsonResponse $response @{ success = $true; message = "Opening folder: $resolvedDir" }
            
            [System.Threading.Tasks.Task]::Run([Action]{
                try {
                    $psi = New-Object System.Diagnostics.ProcessStartInfo
                    $psi.FileName = "explorer.exe"
                    $psi.Arguments = "`"$resolvedDir`""
                    $psi.UseShellExecute = $false
                    [System.Diagnostics.Process]::Start($psi) | Out-Null
                } catch { }
            }) | Out-Null
            continue
        }

        if ($path -eq "/api/scan" -and $request.HttpMethod -eq "POST") {
            $data = Read-RequestBody $request
            $targetDir = if ($data -and $data.directory) { $data.directory } else { (Get-Location).Path }

            if ($data -and $data.async -eq $true) {
                $reg = Register-Operation -Type "SCAN" -Directory $targetDir
                if (-not $reg.success) {
                    Send-JsonResponse $response $reg 409
                    continue
                }
                Start-AsyncOperationWorker -OperationId $reg.operationId -RepoRoot $PSScriptRoot | Out-Null
                Send-JsonResponse $response @{
                    success = $true
                    operationId = $reg.operationId
                    status = "QUEUED"
                    message = "Scan operation registered and queued"
                } 202
                continue
            }
            
            $results = Scan-ExcelDirectory -DirectoryPath $targetDir
            Send-JsonResponse $response $results
            continue
        }

        if ($path -eq "/api/backup" -and $request.HttpMethod -eq "POST") {
            $data = Read-RequestBody $request
            $targetDir = if ($data -and $data.directory) { $data.directory } else { (Get-Location).Path }
            
            $backupResult = Create-ExcelBackup -DirectoryPath $targetDir
            Send-JsonResponse $response $backupResult
            continue
        }

        if ($path -eq "/api/preview" -and $request.HttpMethod -eq "POST") {
            $data = Read-RequestBody $request
            $targetDir = if ($data -and $data.directory) { $data.directory } else { (Get-Location).Path }
            $rules = $data.rules
            $options = $data.options

            if ([string]::IsNullOrWhiteSpace($targetDir) -or -not (Test-Path -LiteralPath $targetDir -PathType Container)) {
                Send-JsonResponse $response @{ success = $false; error = "Hedef klasör bulunamadı: $targetDir" } 400
                continue
            }

            if (-not $rules -or @($rules).Count -eq 0) {
                Send-JsonResponse $response @{ success = $false; error = "En az bir kural belirtilmelidir." } 400
                continue
            }

            $reg = Register-Operation -Type "PREVIEW" -Directory $targetDir -Rules $rules -Options $options
            if (-not $reg.success) {
                Send-JsonResponse $response $reg 409
                continue
            }

            Start-AsyncOperationWorker -OperationId $reg.operationId -RepoRoot $PSScriptRoot | Out-Null
            Send-JsonResponse $response @{
                success = $true
                operationId = $reg.operationId
                status = "QUEUED"
                message = "Preview operation registered and queued"
            } 202
            continue
        }

        if ($path -eq "/api/update" -and $request.HttpMethod -eq "POST") {
            $data = Read-RequestBody $request
            $targetDir = if ($data -and $data.directory) { $data.directory } else { (Get-Location).Path }
            $rules = $data.rules
            $options = $data.options

            if ([string]::IsNullOrWhiteSpace($targetDir) -or -not (Test-Path -LiteralPath $targetDir -PathType Container)) {
                Send-JsonResponse $response @{ success = $false; error = "Hedef klasör bulunamadı: $targetDir" } 400
                continue
            }

            if (-not $rules -or @($rules).Count -eq 0) {
                Send-JsonResponse $response @{ success = $false; error = "En az bir kural belirtilmelidir." } 400
                continue
            }

            # Optional synchronous update if explicitly requested
            if ($request.QueryString["sync"] -eq "true" -or ($data -and $data.sync -eq $true)) {
                $updateResult = Update-ExcelDirectory -DirectoryPath $targetDir -Rules $rules -Options $options
                Send-JsonResponse $response $updateResult
                continue
            }

            $reg = Register-Operation -Type "UPDATE" -Directory $targetDir -Rules $rules -Options $options
            if (-not $reg.success) {
                Send-JsonResponse $response $reg 409
                continue
            }

            Start-AsyncOperationWorker -OperationId $reg.operationId -RepoRoot $PSScriptRoot | Out-Null
            Send-JsonResponse $response @{
                success = $true
                operationId = $reg.operationId
                status = "QUEUED"
                message = "Update operation registered and queued"
            } 202
            continue
        }

        # ----------------------------------------------------------------------
        # STATIC FILES
        # ----------------------------------------------------------------------
        if ($path -eq "/") {
            $path = "/index.html"
        }

        # UrlDecode requested path to prevent encoded traversal attacks
        $decodedPath = ""
        try {
            Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue
            $decodedPath = [System.Web.HttpUtility]::UrlDecode($path).TrimStart('/')
        } catch {
            $decodedPath = [System.Uri]::UnescapeDataString($path).TrimStart('/')
        }
        if ([string]::IsNullOrWhiteSpace($decodedPath)) {
            $decodedPath = [System.Uri]::UnescapeDataString($path).TrimStart('/')
        }

        # Resolve canonical public root directory with trailing separator
        $canonicalPublic = [System.IO.Path]::GetFullPath($publicDir)
        if (-not $canonicalPublic.EndsWith([System.IO.Path]::DirectorySeparatorChar.ToString())) {
            $canonicalPublic += [System.IO.Path]::DirectorySeparatorChar
        }

        # Resolve full canonical path of the requested item
        $canonicalTarget = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($canonicalPublic, $decodedPath))

        # Enforce canonical containment: candidate path must strictly start with canonical public root
        if ($canonicalTarget.StartsWith($canonicalPublic, [System.StringComparison]::OrdinalIgnoreCase) -and (Test-Path -LiteralPath $canonicalTarget -PathType Leaf)) {
            $buffer = [System.IO.File]::ReadAllBytes($canonicalTarget)
            $response.ContentType = Get-ContentType $canonicalTarget
            $response.ContentLength64 = $buffer.Length
            $response.OutputStream.Write($buffer, 0, $buffer.Length)
            $response.OutputStream.Close()
        } else {
            $response.StatusCode = 404
            $err = @{ error = "File not found: $path" }
            Send-JsonResponse $response $err 404
        }
    } catch {
        Write-Host "Request handling error: $_" -ForegroundColor Red
    }
}
