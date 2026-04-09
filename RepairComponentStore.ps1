<#
.SYNOPSIS
    Repairs Windows component store corruption using an offline WIM from SCCM.

.DESCRIPTION
    Uses DISM /RestoreHealth with a locally cached Windows 11 24H2 install.wim
    to repair component store corruption that prevents cumulative updates from installing.
    
    Resolves:
    - 0x800F0838 (CBS_E_MISSING_PREREQUISITE_BASELINES)
    - 0x800F0915 (CBS_E_SOURCE_NOT_IN_LIST)

.NOTES
    RepairComponentStore.ps1
    Version 1.0 - 08/04/2026
    
    SCCM Package: Windows 11 24H2 x64 EN-US Rev: Nov 2025
    Package ID:   CAS04B31
    ISO Build:    10.0.26100.7171
    
    SCCM Task Sequence:
      Step 1 - Download Package Content (CAS04B31) → C:\Scratch
      Step 2 - Run this script (64-bit PowerShell)
#>

#Requires -RunAsAdministrator

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

# Derived paths
$packagePath = Join-Path $ScratchPath $PackageID
$wimFile = Join-Path $packagePath "sources\install.wim"

# ============================================================
# Logging
# ============================================================
function Write-Log {
    param (
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    $entry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Message"
    Write-Host $entry -ForegroundColor $(switch ($Level) {
            "ERROR" { "Red" } "WARN" { "Yellow" } "SUCCESS" { "Green" } default { "White" }
        })
    if (-not (Test-Path $logPath)) { New-Item -Path $logPath -ItemType Directory -Force | Out-Null }
    Add-Content -Path $scriptLog -Value $entry
}

# ============================================================
# STEP 1: Pre-Flight Checks
# ============================================================
Write-Log "================================================================"
Write-Log "Component Store Repair — Starting"
Write-Log "================================================================"

# 64-bit check
if ($env:PROCESSOR_ARCHITEW6432 -eq "AMD64") {
    Write-Log "FATAL: Running in 32-bit PowerShell. Use 64-bit." -Level ERROR
    Exit 1
}
Write-Log "Architecture: $env:PROCESSOR_ARCHITECTURE" -Level SUCCESS

# OS info
$os = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
Write-Log "OS: $($os.ProductName) $($os.DisplayVersion) Build $($os.CurrentBuildNumber).$($os.UBR)"
Write-Log "Edition: $($os.EditionID)"

# Validate WIM exists
if (-not (Test-Path $wimFile)) {
    # Check for install.esd as fallback
    $esdFile = Join-Path $packagePath "sources\install.esd"
    if (Test-Path $esdFile) {
        $wimFile = $esdFile
        Write-Log "Using install.esd format." -Level WARN
    }
    else {
        Write-Log "FATAL: WIM file not found at: $wimFile" -Level ERROR
        Write-Log "Ensure SCCM 'Download Package Content' step completed successfully." -Level ERROR
        Exit 1
    }
}
Write-Log "WIM: $wimFile ($([math]::Round((Get-Item $wimFile).Length/1GB,2)) GB)" -Level SUCCESS

# Disk space (minimum 10GB)
$freeGB = [math]::Round((Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='$env:SystemDrive'").FreeSpace / 1GB, 2)
Write-Log "Free disk space: $freeGB GB"
if ($freeGB -lt 10) {
    Write-Log "FATAL: Need at least 10GB free. Available: $freeGB GB" -Level ERROR
    Exit 1
}

# ============================================================
# STEP 2: Detect Correct WIM Index
# ============================================================
Write-Log "Detecting WIM image index for edition: $($os.EditionID)"

$wimOutput = & $dismExe /Get-WimInfo /WimFile:"$wimFile" 2>&1
$selectedIndex = 0

# Parse indices and match to current edition
$editionMatch = switch ($os.EditionID) {
    "Professional" { "Pro" }
    "Enterprise" { "Enterprise" }
    "Education" { "Education" }
    "Core" { "Home" }
    default { $os.EditionID }
}

$currentIdx = $null
$wimOutput | ForEach-Object {
    if ($_ -match "^Index\s*:\s*(\d+)") { $currentIdx = [int]$Matches }
    if ($_ -match "^Name\s*:\s*(.+)$" -and $currentIdx) {
        $name = $Matches.Trim()
        Write-Log "  Index ${currentIdx}: $name"
        if ($selectedIndex -eq 0 -and $name -match $editionMatch) {
            $selectedIndex = $currentIdx
        }
    }
}

if ($selectedIndex -eq 0) {
    Write-Log "Could not match edition. Defaulting to Index 1." -Level WARN
    $selectedIndex = 1
}
Write-Log "Selected WIM Index: $selectedIndex" -Level SUCCESS

# ============================================================
# STEP 3: Repair Component Store
# ============================================================
$repairSource = "WIM:${wimFile}:${selectedIndex}"
Write-Log "================================================================"
Write-Log "Starting DISM RestoreHealth"
Write-Log "Source: $repairSource"
Write-Log "This may take 15-45 minutes. Do not interrupt."
Write-Log "================================================================"

$startTime = Get-Date

$restoreProc = Start-Process -FilePath $dismExe `
    -ArgumentList "/Online", "/Cleanup-Image", "/RestoreHealth", `
    "/Source:$repairSource", `
    "/LimitAccess", `
    "/LogPath:`"$dismRepairLog`"" `
    -Wait -PassThru -NoNewWindow

$duration = "{0:mm\:ss}" -f ((Get-Date) - $startTime)
$exitCode = $restoreProc.ExitCode

Write-Log "RestoreHealth completed in $duration"

# Handle result
switch ($exitCode) {
    0 {
        Write-Log "Component store repair SUCCEEDED." -Level SUCCESS
    }
    3010 {
        Write-Log "Component store repair SUCCEEDED — reboot required." -Level SUCCESS
    }
    default {
        Write-Log "RestoreHealth failed (exit code: $exitCode). Attempting cleanup + retry..." -Level WARN

        # Cleanup then retry
        Start-Process -FilePath $dismExe `
            -ArgumentList "/Online", "/Cleanup-Image", "/StartComponentCleanup" `
            -Wait -NoNewWindow

        Write-Log "Retrying RestoreHealth..."
        $retryProc = Start-Process -FilePath $dismExe `
            -ArgumentList "/Online", "/Cleanup-Image", "/RestoreHealth", `
            "/Source:$repairSource", "/LimitAccess", `
            "/LogPath:`"$logPath\DISM_RestoreHealth_Retry_$timestamp.log`"" `
            -Wait -PassThru -NoNewWindow

        $exitCode = $retryProc.ExitCode

        if ($exitCode -eq 0 -or $exitCode -eq 3010) {
            Write-Log "Component store repair SUCCEEDED on retry." -Level SUCCESS
        }
        else {
            Write-Log "FATAL: Repair failed after retry. Exit code: $exitCode" -Level ERROR
            Write-Log "Review: $dismRepairLog" -Level ERROR
        }
    }
}

# ============================================================
# STEP 4: SFC Scan
# ============================================================
Write-Log "Running SFC /scannow..."
$sfcProc = Start-Process -FilePath "$env:SystemRoot\System32\sfc.exe" `
    -ArgumentList "/scannow" -Wait -PassThru -NoNewWindow
Write-Log "SFC exit code: $($sfcProc.ExitCode)"

# ============================================================
# STEP 5: Cleanup Scratch Folder
# ============================================================
Write-Log "Cleaning up: $packagePath"
try {
    if (Test-Path $packagePath) {
        Remove-Item -Path $packagePath -Recurse -Force -ErrorAction Stop
        Write-Log "Scratch folder removed." -Level SUCCESS
    }
    # Remove parent if empty
    if ((Test-Path $ScratchPath) -and 
        (Get-ChildItem $ScratchPath -Force | Measure-Object).Count -eq 0) {
        Remove-Item -Path $ScratchPath -Force
        Write-Log "Empty scratch root removed." -Level SUCCESS
    }
}
catch {
    Write-Log "Cleanup warning: $_" -Level WARN
}

# ============================================================
# STEP 6: Final Report
# ============================================================
$postBuild = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"

Write-Log "================================================================"
Write-Log "REPAIR COMPLETE"
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
    Exit 3010
}
else {
    Write-Log "Repair unsuccessful. Escalate with logs from: $logPath" -Level ERROR
    Exit 1
}