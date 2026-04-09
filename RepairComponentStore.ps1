<#
.SYNOPSIS
    Repairs Windows component store corruption using an offline WIM from SCCM.

.DESCRIPTION
    Uses DISM /RestoreHealth with a locally cached Windows 11 24H2 install.wim
    to repair component store corruption that prevents cumulative updates
    from installing.

.NOTES
    RepairComponentStore.ps1 - EverGPT Turbo
    Version 1.1 - 08/04/2026
    
    Changes from v1.0:
    - Fixed WIM index detection for non-English OS locales
    - Added /English flag to DISM /Get-WimInfo
    - Improved edition matching with broader fallback
    - Script now ABORTS if correct index cannot be determined
    - Added all WIM indices to log for troubleshooting
    
    SCCM Package: Windows 11 24H2 x64 EN-US Rev: Nov 2025
    Package ID:   CAS04B31
    ISO Build:    10.0.26100.7171
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
Write-Log "Component Store Repair v1.1 — Starting"
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
    $esdFile = Join-Path $packagePath "sources\install.esd"
    if (Test-Path $esdFile) {
        $wimFile = $esdFile
        Write-Log "Using install.esd format." -Level WARN
    }
    else {
        Write-Log "FATAL: WIM file not found at: $wimFile" -Level ERROR
        Write-Log "Ensure SCCM 'Download Package Content' step completed." -Level ERROR
        Exit 1
    }
}
Write-Log "WIM: $wimFile ($([math]::Round((Get-Item $wimFile).Length/1GB,2)) GB)" -Level SUCCESS

# Disk space
$freeGB = [math]::Round((Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='$env:SystemDrive'").FreeSpace / 1GB, 2)
Write-Log "Free disk space: $freeGB GB"
if ($freeGB -lt 5) {
    Write-Log "FATAL: Need at least 5GB free. Available: $freeGB GB" -Level ERROR
    Exit 1
}

# ============================================================
# STEP 2: Detect Correct WIM Index (Locale-Independent)
# ============================================================
Write-Log "Detecting WIM image index for edition: $($os.EditionID)"

# Force English output regardless of OS locale
$wimInfoOutput = & $dismExe /Get-WimInfo /WimFile:"$wimFile" /English 2>&1

$selectedIndex = 0

# Edition matching map
$editionMatch = switch ($os.EditionID) {
    "Professional" { "Pro" }
    "Enterprise" { "Enterprise" }
    "Education" { "Education" }
    "Core" { "Home" }
    "ProfessionalN" { "Pro N" }
    "EnterpriseN" { "Enterprise N" }
    "ProfessionalEducation" { "Pro Education" }
    "ProfessionalWorkstation" { "Pro for Workstations" }
    default { $os.EditionID }
}

Write-Log "Looking for edition match: '$editionMatch'"

# Parse WIM info
$currentIdx = $null
$allIndices = @()

foreach ($line in $wimInfoOutput) {
    $lineStr = $line.ToString().Trim()
    
    if ($lineStr -match "^Index\s*:\s*(\d+)") {
        $currentIdx = [int]$Matches
    }
    if ($lineStr -match "^Name\s*:\s*(.+)$" -and $null -ne $currentIdx) {
        $name = $Matches.Trim()
        $allIndices += [PSCustomObject]@{ Index = $currentIdx; Name = $name }
        Write-Log "  Index ${currentIdx}: $name"
        
        # Exact match: "Windows 11 Pro" for Professional
        if ($selectedIndex -eq 0 -and $name -match "Windows 1\s+$editionMatch(\s|$)") {
            $selectedIndex = $currentIdx
        }
        $currentIdx = $null
    }
}

# Broader fallback matching
if ($selectedIndex -eq 0 -and $allIndices.Count -gt 0) {
    Write-Log "Exact match not found. Trying broader match..." -Level WARN
    
    foreach ($idx in $allIndices) {
        if ($idx.Name -match [regex]::Escape($editionMatch)) {
            # Avoid false positives: "Pro" shouldn't match "Pro Education"
            if ($editionMatch -eq "Pro" -and $idx.Name -match "Pro (Education|N|for Workstations)") {
                continue
            }
            $selectedIndex = $idx.Index
            Write-Log "Broad match: Index $($idx.Index) — $($idx.Name)" -Level SUCCESS
            break
        }
    }
}

# Abort if no match — do NOT default to Index 1
if ($selectedIndex -eq 0) {
    Write-Log "FATAL: Could not match edition '$($os.EditionID)' to any WIM image." -Level ERROR
    Write-Log "Available images in WIM:" -Level ERROR
    foreach ($idx in $allIndices) {
        Write-Log "  Index $($idx.Index): $($idx.Name)" -Level ERROR
    }
    if ($allIndices.Count -eq 0) {
        Write-Log "No images were parsed from WIM. Raw DISM output:" -Level ERROR
        foreach ($line in $wimInfoOutput) {
            Write-Log "  $line" -Level ERROR
        }
    }
    Write-Log "Cannot proceed without correct edition match. Aborting." -Level ERROR
    Exit 1
}

Write-Log "Selected WIM Index: $selectedIndex (matched '$editionMatch')" -Level SUCCESS

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

switch ($exitCode) {
    0 {
        Write-Log "Component store repair SUCCEEDED." -Level SUCCESS
    }
    3010 {
        Write-Log "Component store repair SUCCEEDED — reboot required." -Level SUCCESS
    }
    default {
        $hexCode = "0x{0:X8}" -f ([uint32]$exitCode)
        Write-Log "RestoreHealth failed (exit code: $exitCode / $hexCode). Attempting cleanup + retry..." -Level WARN

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
            $hexCode = "0x{0:X8}" -f ([uint32]$exitCode)
            Write-Log "FATAL: Repair failed after retry. Exit code: $exitCode / $hexCode" -Level ERROR
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