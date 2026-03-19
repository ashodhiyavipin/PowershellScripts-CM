<# 
.SYNOPSIS
    Script to download all UWP apps from a Windows system with structured logging functionality.

.DESCRIPTION
    This PowerShell script identifies and downloads all UWP apps from the Microsoft Store, 
    ensuring they are placed in a targeted user directory. It includes a robust preflight 
    validation, structured custom logging, retry logic with exponential backoff, and 
    idempotency checks to ensure safe and restartable execution.

.PARAMETER DownloadDirectory
    The destination directory where the UWP apps will be downloaded. Defaults to "$env:USERPROFILE\Downloads\UWPDownloads".

.PARAMETER LogFilePath
    The destination for the structured log file. Defaults to "$env:USERPROFILE\Downloads\UWPDownloads\DownloadUWPApps.log".

.PARAMETER MaxRetries
    The maximum number of retries for an app download failure (network errors, store throttling). Defaults to 3.

.NOTES
    Download-UWPApps.ps1
    Script History:
    Version 1.0 - Script inception
    Version 1.1 - Added Microsoft.WindowsAlarms and added logging function.
    Version 1.2 - Added folder renaming functionality.
    Version 2.0 - Complete rewrite: Added preflight checks, retry logic, structured logging, 
                  idempotency, strict mode, unified loop execution, and robust error handling.
    Version 2.1 - Removed msstore validation, optimized download directory targeting to prevent 
                  dependency clashing, and fixed StrictMode array count edge-cases.
#>
[CmdletBinding()]
param (
    [string]$DownloadDirectory = "$env:USERPROFILE\Downloads\UWPDownloads",
    [string]$LogFilePath = "$env:USERPROFILE\Downloads\UWPDownloads\DownloadUWPApps.log",
    [int]$MaxRetries = 3
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$WarningPreference = 'Continue'

# --- Initialization and Logging --- #
$script:stepIdCount = 1

function Write-StructuredLog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$LogLevel = 'INFO',
        
        [int]$ExitCode = 0,
        [double]$DurationSeconds = 0
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $durationStr = if ($DurationSeconds -gt 0) { "[Duration: $([math]::Round($DurationSeconds, 2))s]" } else { "" }
    $exitCodeStr = if ($LogLevel -ne 'INFO') { "[ExitCode: $ExitCode]" } else { "" }
    
    $stepPadded = $script:stepIdCount.ToString('000')
    $logEntry = "$timestamp | Step: $stepPadded | [$LogLevel] $Message $exitCodeStr $durationStr".Trim()
    
    # Send to appropriate streams
    switch ($LogLevel) {
        'ERROR' { Write-Error $logEntry -ErrorAction Continue }
        'WARN' { Write-Warning $logEntry }
        default { Write-Verbose $logEntry -Verbose }
    }
    
    # Write to local file
    try {
        Add-Content -Path $LogFilePath -Value $logEntry -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to write to log file: $_"
    }
    
    $script:stepIdCount++
}

# Ensure download directory exists cleanly
if (-not (Test-Path -Path $DownloadDirectory)) {
    try {
        New-Item -Path $DownloadDirectory -ItemType Directory -Force | Out-Null
    }
    catch {
        throw "Failed to create target download directory: $DownloadDirectory. Error: $_"
    }
}

# Ensure log directory parent exists cleanly
$logDir = Split-Path -Path $LogFilePath -Parent
if ($logDir -and -not (Test-Path -Path $logDir)) {
    try {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    }
    catch {
        throw "Failed to create log directory: $logDir. Error: $_"
    }
}

Write-StructuredLog -Message "Starting preflight checks..."

# --- Preflight Checks --- #

# 1. Validate Internet Connectivity (using standard Microsoft endpoint)
try {
    $null = Resolve-DnsName -Name "www.microsoft.com" -ErrorAction Stop
    Write-StructuredLog -Message "Internet connectivity validated."
}
catch {
    Write-StructuredLog -Message "Failed to validate internet connectivity: $_" -LogLevel 'ERROR'
    throw "No internet connectivity. Aborting script."
}

# 2. Validate winget availability
$wingetPath = Get-Command winget.exe -ErrorAction SilentlyContinue
if (-not $wingetPath) {
    Write-StructuredLog -Message "winget.exe not found in PATH." -LogLevel 'ERROR'
    throw "Winget is not installed or accessible. Aborting script."
}
Write-StructuredLog -Message "Winget found at $($wingetPath.Source)."

# 3. Validate winget version
try {
    $wingetVersion = & winget.exe --version
    Write-StructuredLog -Message "Winget version: $wingetVersion"
}
catch {
    Write-StructuredLog -Message "Failed to query winget version." -LogLevel 'WARN'
}


# --- App Manifest --- #
$appsToDownload = @{
    "9WZDNCRFJBMP" = @{ Name = "Microsoft.WindowsStore"; Platform = "Windows.Desktop" }
    "9N4D0MSMP0PT" = @{ Name = "Microsoft.VP9VideoExtensions"; Platform = "Windows.Universal" }
    "9MZ95KL8MR0L" = @{ Name = "Microsoft.ScreenSketch"; Platform = "Windows.Desktop" }
    "9WZDNCRFJ3PZ" = @{ Name = "Microsoft.CompanyPortal"; Platform = "Windows.Universal" }
    "9PMMSR1CGPWG" = @{ Name = "Microsoft.HEIFImageExtension"; Platform = "Windows.Universal" }
    "9PCFS5B6T72H" = @{ Name = "Microsoft.Paint"; Platform = "Windows.Desktop" }
    "9NCTDW2W1BH8" = @{ Name = "Microsoft.RawImageExtension"; Platform = "Windows.Universal" }
    "9N5TDP8VCMHS" = @{ Name = "Microsoft.WebMediaExtensions"; Platform = "Windows.Universal" }
    "9PG2DK419DRG" = @{ Name = "Microsoft.WebpImageExtension"; Platform = "Windows.Universal" }
    "9WZDNCRFJBH4" = @{ Name = "Microsoft.Windows.Photos"; Platform = "Windows.Desktop" }
    "9WZDNCRFHVN5" = @{ Name = "Microsoft.WindowsCalculator"; Platform = "Windows.Desktop" }
    "9WZDNCRFJBBG" = @{ Name = "Microsoft.WindowsCamera"; Platform = "Windows.Desktop" }
    "9MSMLRH6LZF3" = @{ Name = "Microsoft.WindowsNotepad"; Platform = "Windows.Desktop" }
    "9N0DX20HK701" = @{ Name = "Microsoft.WindowsTerminal"; Platform = "Windows.Desktop" }
    "9N1F85V9T8BN" = @{ Name = "MicrosoftCorporationII.Windows365"; Platform = "Windows.Desktop" }
    "9NBLGGH4QGHW" = @{ Name = "Microsoft.MicrosoftStickyNotes"; Platform = "Windows.Universal" }
    "9PGJGD53TN86" = @{ Name = "WinDbg"; Platform = "Windows.Desktop" }
    "9WZDNCRFJ3PR" = @{ Name = "Microsoft.WindowsAlarms"; Platform = "Windows.Desktop" }
}

# --- Main Unified Processing Loop --- #
Write-StructuredLog -Message "Starting Application Downloads..."

$summaryResults = @()

foreach ($id in $appsToDownload.Keys) {
    # Extract properties explicitly to avoid strict-mode hash lookup issues
    $appInfo = $appsToDownload[$id]
    $appName = $appInfo.Name
    $platform = $appInfo.Platform

    $resultObj = [PSCustomObject]@{
        Id             = $id
        AppName        = $appName
        DownloadStatus = 'Pending'
        RenameStatus   = 'Pending'
        ExitCode       = $null
    }

    $renamedPath = Join-Path -Path $DownloadDirectory -ChildPath $appName
    $idPath = Join-Path -Path $DownloadDirectory -ChildPath $id

    # Check Idempotency: Has the target app already been fully downloaded and renamed?
    if (Test-Path -Path $renamedPath) {
        Write-StructuredLog -Message "Application folder already exists: '$renamedPath'. Skipping download."
        $resultObj.DownloadStatus = 'Skipped'
        $resultObj.RenameStatus = 'Skipped (Already Finalized)'
        $summaryResults += $resultObj
        continue
    }

    # If the unrenamed ID directory exists, we still execute winget to guarantee 
    # the payload is complete (winget automatically validates/resumes existing files).
    $needsDownload = $true
    
    if ($needsDownload) {
        $retryCount = 0
        $downloadSuccess = $false
        $lastExit = -1

        while ($retryCount -lt $MaxRetries -and -not $downloadSuccess) {
            $sw = [Diagnostics.Stopwatch]::StartNew()
            Write-StructuredLog -Message "Initiating download for $($appName) ($id) - Attempt $($retryCount + 1)/$MaxRetries"
            
            try {
                # Executing natively allows Winget to handle MS Entra Enterprise auth prompts directly in host console
                $script:LASTEXITCODE = 0
                & winget.exe download --id $id --Platform $platform --architecture x64 --accept-source-agreements --accept-package-agreements -s msstore --download-directory $idPath
                $lastExit = $LASTEXITCODE
                
                if ($lastExit -eq 0) {
                    $downloadSuccess = $true
                    Write-StructuredLog -Message "Successfully downloaded $appName" -DurationSeconds $sw.Elapsed.TotalSeconds
                }
                else {
                    Write-StructuredLog -Message "Winget exited with code $lastExit for $appName" -LogLevel 'WARN' -ExitCode $lastExit -DurationSeconds $sw.Elapsed.TotalSeconds
                    $retryCount++
                }
            }
            catch {
                Write-StructuredLog -Message "Execution exception calling winget for ${appName}: $_" -LogLevel 'ERROR' -DurationSeconds $sw.Elapsed.TotalSeconds
                $retryCount++
                $lastExit = if ($LASTEXITCODE) { $LASTEXITCODE } else { 1 }
            }
            finally {
                if ($sw.IsRunning) { $sw.Stop() }
            }

            # Handle Exponential Backoff on Failure
            if (-not $downloadSuccess -and $retryCount -lt $MaxRetries) {
                $backoff = [Math]::Pow(2, $retryCount)
                Write-StructuredLog -Message "Executing exponential backoff for $backoff seconds before next retry..."
                Start-Sleep -Seconds $backoff
            }
        }

        $resultObj.ExitCode = $lastExit

        if (-not $downloadSuccess) {
            Write-StructuredLog -Message "Download permanently failed for $appName after $MaxRetries retries." -LogLevel 'ERROR' -ExitCode $lastExit
            $resultObj.DownloadStatus = 'Failed'
            $resultObj.RenameStatus = 'Not Attempted'
            $summaryResults += $resultObj
            continue
        }
        else {
            $resultObj.DownloadStatus = 'Success'
        }
    }
    else {
        $resultObj.DownloadStatus = 'Skipped (Found ID Folder)'
    }

    # --- Renaming Logistics --- #
    try {
        if (Test-Path -Path $idPath) {
            Write-StructuredLog -Message "Renaming standard output folder '$id' to '$appName'..."
            Rename-Item -Path $idPath -NewName $appName -ErrorAction Stop
            Write-StructuredLog -Message "Successfully renamed directory to '$appName'."
            $resultObj.RenameStatus = 'Success'
        }
        else {
            # Winget behavior fallback: Attempt a fuzzy match on created child directories
            $possibleFolders = Get-ChildItem -Path $DownloadDirectory -Directory -Filter "*$appName*" | Where-Object { $_.Name -ne $appName }
            
            if ($possibleFolders) {
                # Pick the most likely candidate logically
                $targetFolder = $possibleFolders[0]
                Write-StructuredLog -Message "Fallback Renaming: Modifying alternative folder '$($targetFolder.Name)' to '$appName'..."
                Rename-Item -Path $targetFolder.FullName -NewName $appName -ErrorAction Stop
                Write-StructuredLog -Message "Successfully renamed fallback directory to '$appName'."
                $resultObj.RenameStatus = 'Success (Fuzzy Match)'
            }
            else {
                Write-StructuredLog -Message "Severe logic error: Could not pinpoint source directory to rename for $appName. Expected path was '$idPath'." -LogLevel 'WARN'
                $resultObj.RenameStatus = 'Failed (Missing Artifacts)'
            }
        }
    }
    catch {
        Write-StructuredLog -Message "Exception thrown during Rename-Item for '$appName': $_" -LogLevel 'ERROR'
        $resultObj.RenameStatus = "Failed ($($_.Exception.Message))"
    }

    $summaryResults += $resultObj
}

# --- Final Output Formatting --- #
Write-StructuredLog -Message "All sequential operations completed. Preparing and outputting final manifest."

Write-Output ""
Write-Output "==============================================="
Write-Output "         Download UWP Apps Summary             "
Write-Output "==============================================="
$summaryResults | Format-Table -Property Id, AppName, DownloadStatus, RenameStatus, ExitCode -AutoSize | Out-String | Write-Output

$failedCount = @($summaryResults | Where-Object { $_.DownloadStatus -eq 'Failed' }).Count
Write-StructuredLog -Message "Script execution formally concluded. Processed Apps: $($appsToDownload.Count). Failed Downloads: $failedCount."