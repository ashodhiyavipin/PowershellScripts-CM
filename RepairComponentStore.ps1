#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Repairs Windows component store corruption using an offline WIM from SCCM.

.DESCRIPTION
    Uses DISM /RestoreHealth with a locally cached Windows 11 24H2 install.wim
    to repair component store corruption that prevents cumulative updates
    from installing.

    The WIM file contains a single image (Index 1: Windows 11 Pro).
    No index detection or edition matching is performed.

.NOTES
    RepairComponentStore.ps1
    Version 1.4 - 13/04/2026
    
    Changes from v1.3:
    - Removed entire WIM index detection and edition matching logic (STEP 2).
      The WIM contains a single image (Index 1: Windows 11 Pro) so index
      selection is hardcoded to 1.
    - DISM /Get-WimInfo is still executed and logged for audit/troubleshooting
      purposes but no parsing or matching is performed.
    - Significantly reduced script complexity and eliminated all regex parsing
      bugs that affected v1.0 through v1.2.
    
    Changes from v1.2:
    - Fixed $Matches capture group access using .Item(1) method syntax.
    - Added comprehensive before/after logging for every operation.
    - Added full raw DISM /Get-WimInfo output dump to log.
    - Added per-line diagnostic logging during WIM index parsing loop.
    
    Changes from v1.1:
    - Fixed critical bug: DISM /Get-WimInfo output not reliably split into
      individual lines.
    - Replaced deprecated Get-WmiObject with Get-CimInstance.
    - Fixed uint32 cast overflow on negative DISM exit codes.
    - Moved log directory creation before first Write-Log call.
    - Improved duration formatting for repairs exceeding 60 minutes.
    
    Changes from v1.0:
    - Fixed WIM index detection for non-English OS locales.
    - Added /English flag to DISM /Get-WimInfo.
    - Script now ABORTS if correct index cannot be determined.
    
    SCCM Package: Windows 11 24H2 x64 EN-US Rev: Nov 2025
    Package ID:   CAS04B31
    ISO Build:    10.0.26100.7171
    WIM Contents: Single image - Index 1: Windows 11 Pro
#>

# ============================================================
# Configuration
# ============================================================
$PackageID = "CAS04B31"
$ScratchPath = "C:\Scratch"
$logPath = "C:\Windows\fndr\logs"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$scriptLog = "$logPath\RepairComponentStore_$timestamp.log"
$dismRepairLog = "$logPath\DISM_RestoreHealth_$timestamp.log"
$dismExe = "$env:SystemRoot\System32\Dism.exe"

# WIM Index — hardcoded, single-image WIM (Index 1: Windows 11 Pro)
$wimIndex = 1

# Derived paths
$packagePath = Join-Path $ScratchPath $PackageID
$wimFile = Join-Path $packagePath "sources\install.wim"

# ============================================================
# Initialize Log Directory (once, before any Write-Log call)
# ============================================================
if (-not (Test-Path $logPath)) {
    New-Item -Path $logPath -ItemType Directory -Force | Out-Null
}

# ============================================================
# Logging
# ============================================================
function Write-Log {
    param (
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS", "DEBUG")]
        [string]$Level = "INFO"
    )
    $entry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Message"
    Write-Host $entry -ForegroundColor $(switch ($Level) {
            "ERROR" { "Red" } "WARN" { "Yellow" } "SUCCESS" { "Green" } "DEBUG" { "Cyan" } default { "White" }
        })
    Add-Content -Path $scriptLog -Value $entry
}

# ============================================================
# FUNCTION: Safe Hex Conversion
# ============================================================
function ConvertTo-HexString {
    param ([int]$ExitCode)

    if ($ExitCode -lt 0) {
        return "0x{0:X8}" -f ([uint32]([int64]$ExitCode + [int64]4294967296))
    }
    else {
        return "0x{0:X8}" -f ([uint32]$ExitCode)
    }
}

# ============================================================
# FUNCTION: Format Duration (handles >60 minutes)
# ============================================================
function Format-Duration {
    param ([TimeSpan]$Duration)

    $totalMinutes = [math]::Floor($Duration.TotalMinutes)
    $seconds = $Duration.Seconds

    if ($totalMinutes -ge 60) {
        $hours = [math]::Floor($totalMinutes / 60)
        $minutes = $totalMinutes % 60
        return "{0}h {1:D2}m {2:D2}s" -f $hours, $minutes, $seconds
    }
    else {
        return "{0}m {1:D2}s" -f $totalMinutes, $seconds
    }
}

# ============================================================
# STEP 1: Pre-Flight Checks
# ============================================================
Write-Log "================================================================"
Write-Log "Component Store Repair v1.4 - Starting"
Write-Log "================================================================"

# Log execution context
Write-Log "Logging execution context..."
Write-Log "Computer Name : $env:COMPUTERNAME"
Write-Log "Running User  : $env:USERNAME"
Write-Log "PS Version    : $($PSVersionTable.PSVersion)"
Write-Log "Script Log    : $scriptLog"
Write-Log "DISM Log      : $dismRepairLog"
Write-Log "Execution context logged."

# 64-bit check
Write-Log "Checking PowerShell architecture..."
if ($env:PROCESSOR_ARCHITEW6432 -eq "AMD64") {
    Write-Log "FATAL: Running in 32-bit PowerShell. Use 64-bit." -Level ERROR
    Exit 1
}
Write-Log "Architecture: $env:PROCESSOR_ARCHITECTURE" -Level SUCCESS
Write-Log "Architecture check passed."

# OS info
Write-Log "Reading OS information from registry..."
$os = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
Write-Log "OS: $($os.ProductName) $($os.DisplayVersion) Build $($os.CurrentBuildNumber).$($os.UBR)"
Write-Log "Edition: $($os.EditionID)"
Write-Log "OS information retrieved."

# Validate WIM exists
Write-Log "Checking for WIM file at: $wimFile"
if (-not (Test-Path $wimFile)) {
    Write-Log "WIM not found. Checking for ESD format..." -Level WARN
    $esdFile = Join-Path $packagePath "sources\install.esd"
    Write-Log "Checking for ESD file at: $esdFile"
    if (Test-Path $esdFile) {
        $wimFile = $esdFile
        Write-Log "Using install.esd format. Path updated to: $wimFile" -Level WARN
    }
    else {
        Write-Log "FATAL: Neither WIM nor ESD file found." -Level ERROR
        Write-Log "Expected WIM: $wimFile" -Level ERROR
        Write-Log "Expected ESD: $esdFile" -Level ERROR
        Write-Log "Ensure SCCM 'Download Package Content' step completed." -Level ERROR
        Exit 1
    }
}
$wimFileSize = [math]::Round((Get-Item $wimFile).Length / 1GB, 2)
Write-Log "WIM: $wimFile ($wimFileSize GB)" -Level SUCCESS
Write-Log "WIM file validation passed."

# Disk space
Write-Log "Checking available disk space on $env:SystemDrive..."
$disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$env:SystemDrive'"
$freeGB = [math]::Round($disk.FreeSpace / 1GB, 2)
$totalGB = [math]::Round($disk.Size / 1GB, 2)
Write-Log "Disk: $env:SystemDrive - Free: $freeGB GB / Total: $totalGB GB"
if ($freeGB -lt 5) {
    Write-Log "FATAL: Need at least 5GB free. Available: $freeGB GB" -Level ERROR
    Exit 1
}
Write-Log "Disk space check passed." -Level SUCCESS

# ============================================================
# STEP 2: Log WIM Contents (Audit Only — No Matching)
# ============================================================
Write-Log "================================================================"
Write-Log "STEP 2: WIM Image Audit (Index hardcoded to $wimIndex)"
Write-Log "================================================================"
Write-Log "Running DISM /Get-WimInfo for audit logging..."
Write-Log "Command: $dismExe /Get-WimInfo /WimFile:`"$wimFile`""

$wimInfoRaw = & $dismExe /Get-WimInfo /WimFile:"$wimFile" 2>&1

Write-Log "DISM /Get-WimInfo command completed."
Write-Log "----------------------------------------------------------------"
Write-Log "BEGIN: WIM Image Information"
Write-Log "----------------------------------------------------------------"

foreach ($line in $wimInfoRaw) {
    $lineStr = $line.ToString().Trim()
    if (-not [string]::IsNullOrWhiteSpace($lineStr)) {
        Write-Log "  $lineStr" -Level DEBUG
    }
}

Write-Log "----------------------------------------------------------------"
Write-Log "END: WIM Image Information"
Write-Log "----------------------------------------------------------------"
Write-Log "Using hardcoded WIM Index: $wimIndex" -Level SUCCESS
Write-Log "No index detection or edition matching performed."
Write-Log "WIM audit step completed."

# ============================================================
# STEP 3: Repair Component Store
# ============================================================
$repairSource = "WIM:${wimFile}:${wimIndex}"
Write-Log "================================================================"
Write-Log "STEP 3: DISM RestoreHealth"
Write-Log "================================================================"
Write-Log "Repair source: $repairSource"
Write-Log "DISM executable: $dismExe"
Write-Log "DISM repair log: $dismRepairLog"
Write-Log "This may take 15-45 minutes. Do not interrupt."
Write-Log "================================================================"

Write-Log "Starting DISM RestoreHealth..."
$startTime = Get-Date

$restoreProc = Start-Process -FilePath $dismExe `
    -ArgumentList "/Online /Cleanup-Image /RestoreHealth /Source:$repairSource /LimitAccess /LogPath:`"$dismRepairLog`"" `
    -Wait -PassThru -NoNewWindow

$duration = Format-Duration -Duration ((Get-Date) - $startTime)
$exitCode = $restoreProc.ExitCode

Write-Log "DISM RestoreHealth process ended."
Write-Log "Duration: $duration"
Write-Log "Exit code: $exitCode"

switch ($exitCode) {
    0 {
        Write-Log "Component store repair SUCCEEDED." -Level SUCCESS
    }
    3010 {
        Write-Log "Component store repair SUCCEEDED - reboot required." -Level SUCCESS
    }
    default {
        $hexCode = ConvertTo-HexString -ExitCode $exitCode
        Write-Log "RestoreHealth failed (exit code: $exitCode / $hexCode)." -Level WARN
        Write-Log "Attempting StartComponentCleanup before retry..." -Level WARN

        Write-Log "Starting StartComponentCleanup..."
        $cleanupStart = Get-Date

        Start-Process -FilePath $dismExe `
            -ArgumentList "/Online /Cleanup-Image /StartComponentCleanup" `
            -Wait -NoNewWindow

        $cleanupDuration = Format-Duration -Duration ((Get-Date) - $cleanupStart)
        Write-Log "StartComponentCleanup completed. Duration: $cleanupDuration"

        Write-Log "Starting RestoreHealth retry..."
        $retryStart = Get-Date

        $retryProc = Start-Process -FilePath $dismExe `
            -ArgumentList "/Online /Cleanup-Image /RestoreHealth /Source:$repairSource /LimitAccess /LogPath:`"$logPath\DISM_RestoreHealth_Retry_$timestamp.log`"" `
            -Wait -PassThru -NoNewWindow

        $retryDuration = Format-Duration -Duration ((Get-Date) - $retryStart)
        $exitCode = $retryProc.ExitCode

        Write-Log "RestoreHealth retry process ended."
        Write-Log "Retry duration: $retryDuration"
        Write-Log "Retry exit code: $exitCode"

        if ($exitCode -eq 0 -or $exitCode -eq 3010) {
            Write-Log "Component store repair SUCCEEDED on retry." -Level SUCCESS
        }
        else {
            $hexCode = ConvertTo-HexString -ExitCode $exitCode
            Write-Log "FATAL: Repair failed after retry. Exit code: $exitCode / $hexCode" -Level ERROR
            Write-Log "Review primary DISM log: $dismRepairLog" -Level ERROR
            Write-Log "Review retry DISM log: $logPath\DISM_RestoreHealth_Retry_$timestamp.log" -Level ERROR
        }
    }
}

# ============================================================
# STEP 4: SFC Scan
# ============================================================
Write-Log "================================================================"
Write-Log "STEP 4: SFC Scan"
Write-Log "================================================================"
Write-Log "Starting SFC /scannow..."
$sfcStart = Get-Date

$sfcProc = Start-Process -FilePath "$env:SystemRoot\System32\sfc.exe" `
    -ArgumentList "/scannow" -Wait -PassThru -NoNewWindow

$sfcDuration = Format-Duration -Duration ((Get-Date) - $sfcStart)
Write-Log "SFC process ended."
Write-Log "SFC duration: $sfcDuration"
Write-Log "SFC exit code: $($sfcProc.ExitCode)"

# ============================================================
# STEP 5: Cleanup Scratch Folder
# ============================================================
Write-Log "================================================================"
Write-Log "STEP 5: Cleanup"
Write-Log "================================================================"
Write-Log "Cleaning up package path: $packagePath"
try {
    if (Test-Path $packagePath) {
        Write-Log "Removing: $packagePath"
        Remove-Item -Path $packagePath -Recurse -Force -ErrorAction Stop
        Write-Log "Scratch folder removed." -Level SUCCESS
    }
    else {
        Write-Log "Package path does not exist (already cleaned): $packagePath" -Level WARN
    }

    if ((Test-Path $ScratchPath) -and
        (Get-ChildItem $ScratchPath -Force | Measure-Object).Count -eq 0) {
        Write-Log "Scratch root is empty. Removing: $ScratchPath"
        Remove-Item -Path $ScratchPath -Force
        Write-Log "Empty scratch root removed." -Level SUCCESS
    }
    elseif (Test-Path $ScratchPath) {
        $remainingItems = (Get-ChildItem $ScratchPath -Force | Measure-Object).Count
        Write-Log "Scratch root not empty ($remainingItems items remain). Keeping: $ScratchPath"
    }
}
catch {
    Write-Log "Cleanup warning: $_" -Level WARN
}

Write-Log "Cleanup step completed."

# ============================================================
# STEP 6: Final Report
# ============================================================
Write-Log "Reading post-repair OS build info..."
$postBuild = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"

Write-Log "================================================================"
Write-Log "STEP 6: FINAL REPORT"
Write-Log "================================================================"
Write-Log "Pre-repair build:  $($os.CurrentBuildNumber).$($os.UBR)"
Write-Log "Post-repair build: $($postBuild.CurrentBuildNumber).$($postBuild.UBR)"
Write-Log "RestoreHealth:     $(if($exitCode -eq 0 -or $exitCode -eq 3010){'PASSED'}else{'FAILED'})"
Write-Log "SFC:               $(if($sfcProc.ExitCode -eq 0){'PASSED'}else{'COMPLETED WITH FINDINGS'})"
Write-Log "Script log:        $scriptLog"
Write-Log "DISM log:          $dismRepairLog"
Write-Log "================================================================"

if ($exitCode -eq 0 -or $exitCode -eq 3010) {
    Write-Log "Machine is ready for cumulative updates after reboot." -Level SUCCESS
    Write-Log "The KB in Software Center should install successfully after restart." -Level SUCCESS
    Write-Log "Exiting with code 3010 (reboot required)."
    Exit 3010
}
else {
    Write-Log "Repair unsuccessful. Escalate with logs from: $logPath" -Level ERROR
    Write-Log "Exiting with code 1 (failure)."
    Exit 1
}