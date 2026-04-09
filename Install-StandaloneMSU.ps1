#Requires -RunAsAdministrator
#Requires -Version 5.1

<#
.SYNOPSIS
    Installs a single Microsoft Update Standalone Package (MSU) on an online Windows PC.

.DESCRIPTION
    This script automates the installation of a specified MSU file using DISM.
    Includes architecture validation, prerequisite checks, pending reboot detection,
    admin privilege verification, DISM timeout protection, and comprehensive error
    handling with mapped exit codes.

.NOTES
    Install-StandaloneMSU.ps1
    Script History:
    Version 1.0 - Script inception.
    Version 2.0 - Enhanced with arch check, reboot detection, prerequisite validation,
                   improved exit code handling, and separated logs.
    Version 3.0 - Fixed critical log file collision (script and DISM now use separate logs).
                   Fixed script path resolution inside Assert-64BitExecution function.
                   Fixed fragile pipeline exit code capture during 64-bit re-launch.
                   Fixed MSU discovery to use script directory instead of working directory.
                   Replaced deprecated Get-WmiObject with Get-CimInstance.
                   Added DISM execution timeout (60 minutes) to prevent infinite hangs.
                   Fixed Start-Process argument quoting for paths with spaces.
                   Added administrator privilege validation at script start.
                   Fixed uint32 cast overflow on negative DISM exit codes.
                   Moved log directory creation out of Write-Log to avoid per-call overhead and race conditions.
                   Added multi-MSU detection warning.
                   Added machine context logging (computer name, user, PS version, script path).
                   Added #Requires statements for prerequisites enforcement.
#>

#---------------------------------------------------------------------#
# CONFIGURATION
#---------------------------------------------------------------------#

$logFilePath = "C:\Windows\fndr\logs"
$scriptLogFile = "$logFilePath\Install-StandaloneMSU.log"
$dismLogFile = "$logFilePath\Install-StandaloneMSU-DISM.log"
$minDiskSpaceGB = 5
$dismTimeoutMs = 3600000  # 60 minutes

# Resolve the script's own full path at script scope (before any function call)
$script:ScriptFullPath = $MyInvocation.MyCommand.Definition
$script:ScriptDir = Split-Path -Parent $script:ScriptFullPath

#---------------------------------------------------------------------#
# INITIALIZE LOG DIRECTORY (once, before any Write-Log call)
#---------------------------------------------------------------------#
if (-not (Test-Path $logFilePath)) {
    New-Item -Path $logFilePath -ItemType Directory -Force | Out-Null
}

#---------------------------------------------------------------------#
# FUNCTION: Write-Log
#---------------------------------------------------------------------#
function Write-Log {
    param (
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR")]
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp [$Level] $Message"
    Write-Output $logMessage
    Add-Content -Path $scriptLogFile -Value $logMessage
}

#---------------------------------------------------------------------#
# FUNCTION: Ensure 64-bit Execution
#---------------------------------------------------------------------#
function Assert-64BitExecution {
    if ($env:PROCESSOR_ARCHITEW6432 -eq "AMD64") {
        Write-Log "Detected 32-bit (WOW64) PowerShell session. Re-launching in native 64-bit..." -Level WARN

        if (-not $script:ScriptFullPath -or -not (Test-Path $script:ScriptFullPath)) {
            Write-Log "Cannot determine script path for 64-bit re-launch. Aborting." -Level ERROR
            Exit 1
        }

        $relaunchExe = "$env:SystemRoot\SysNative\WindowsPowerShell\v1.0\powershell.exe"
        $relaunchArgs = "-ExecutionPolicy Bypass -NoProfile -File `"$($script:ScriptFullPath)`""

        Write-Log "Re-launching: $relaunchExe $relaunchArgs"

        $proc = Start-Process -FilePath $relaunchExe `
            -ArgumentList $relaunchArgs `
            -Wait -PassThru

        $exitCode = $proc.ExitCode
        Write-Log "64-bit process completed with exit code: $exitCode"
        Exit $exitCode
    }

    # Confirm we are running 64-bit
    if ([IntPtr]::Size -ne 8) {
        Write-Log "Failed to confirm 64-bit execution environment. Aborting." -Level ERROR
        Exit 1
    }

    Write-Log "Running in native 64-bit PowerShell. Architecture: $env:PROCESSOR_ARCHITECTURE"
}

#---------------------------------------------------------------------#
# FUNCTION: Assert Administrator Privileges
#---------------------------------------------------------------------#
function Assert-AdminPrivileges {
    $currentPrincipal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    $isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if (-not $isAdmin) {
        Write-Log "This script requires Administrator privileges. Current session is not elevated. Aborting." -Level ERROR
        Exit 1
    }

    Write-Log "Administrator privileges confirmed."
}

#---------------------------------------------------------------------#
# FUNCTION: Check Pending Reboot
#---------------------------------------------------------------------#
function Test-PendingReboot {
    Write-Log "Checking for pending reboot..."

    $rebootRequired = $false
    $reasons = @()

    # Check Component-Based Servicing
    $cbsKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending"
    if (Test-Path $cbsKey) {
        $rebootRequired = $true
        $reasons += "Component Based Servicing"
    }

    # Check Windows Update
    $wuKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
    if (Test-Path $wuKey) {
        $rebootRequired = $true
        $reasons += "Windows Update"
    }

    # Check Pending File Rename Operations
    $pfrValue = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" `
        -Name "PendingFileRenameOperations" -ErrorAction SilentlyContinue
    if ($pfrValue.PendingFileRenameOperations) {
        $rebootRequired = $true
        $reasons += "Pending File Rename Operations"
    }

    if ($rebootRequired) {
        Write-Log "PENDING REBOOT DETECTED. Reasons: $($reasons -join ', ')" -Level WARN
        return $true
    }

    Write-Log "No pending reboot detected."
    return $false
}

#---------------------------------------------------------------------#
# FUNCTION: Check Disk Space
#---------------------------------------------------------------------#
function Test-DiskSpace {
    param (
        [int]$MinimumGB = 5
    )

    $systemDrive = $env:SystemDrive
    $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$systemDrive'"
    $freeGB = [math]::Round($disk.FreeSpace / 1GB, 2)

    Write-Log "Free disk space on ${systemDrive}: $freeGB GB (minimum required: $MinimumGB GB)"

    if ($freeGB -lt $MinimumGB) {
        Write-Log "Insufficient disk space. Available: $freeGB GB, Required: $MinimumGB GB" -Level ERROR
        return $false
    }

    return $true
}

#---------------------------------------------------------------------#
# FUNCTION: Get OS Build Info
#---------------------------------------------------------------------#
function Get-OSBuildInfo {
    $ntCurrent = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
    $currentBuild = $ntCurrent.CurrentBuildNumber
    $ubr = $ntCurrent.UBR
    $fullBuild = "$currentBuild.$ubr"

    Write-Log "Current OS Build: $fullBuild (Build: $currentBuild, UBR: $ubr)"

    return @{
        Build     = $currentBuild
        UBR       = $ubr
        FullBuild = $fullBuild
    }
}

#---------------------------------------------------------------------#
# FUNCTION: Map DISM Exit Codes
#---------------------------------------------------------------------#
function Get-ExitCodeDescription {
    param ([int]$ExitCode)

    $exitCodeMap = @{
        0           = "Success - No reboot required"
        1           = "General failure"
        2           = "Update not applicable"
        3010        = "Success - Reboot required (ERROR_SUCCESS_REBOOT_REQUIRED)"
        1058        = "Service disabled (Windows Update service may be disabled)"
        1641        = "Success - Reboot initiated (ERROR_SUCCESS_REBOOT_INITIATED)"
        -2146498530 = "0x800F081E - CBS_E_NOT_APPLICABLE - Update not applicable"
        -2146498504 = "0x800F0838 - CBS_E_MISSING_PREREQUISITE_BASELINES - Missing prerequisite updates"
        -2146498512 = "0x800F0830 - CBS_E_ALREADY_EXISTS - Update already installed"
        -2149842967 = "0x80240009 - WU_E_OPERATIONINPROGRESS - Another update operation in progress"
        -2145124329 = "0x80243007 - WU_E_DM_NOTDOWNLOADED - Update not downloaded"
    }

    if ($exitCodeMap.ContainsKey($ExitCode)) {
        return $exitCodeMap[$ExitCode]
    }

    # Safe hex conversion handling negative values
    if ($ExitCode -lt 0) {
        $hex = "0x{0:X8}" -f ([uint32]([int64]$ExitCode + [int64]4294967296))
    }
    else {
        $hex = "0x{0:X8}" -f ([uint32]$ExitCode)
    }

    return "Unknown error code: $ExitCode ($hex)"
}

#---------------------------------------------------------------------#
# MAIN EXECUTION
#---------------------------------------------------------------------#

Write-Log "=========================================================="
Write-Log "Install-StandaloneMSU v3.0 - Starting"
Write-Log "=========================================================="

# Step 1: Log execution context
Write-Log "Computer Name : $env:COMPUTERNAME"
Write-Log "Running User  : $env:USERNAME"
Write-Log "PS Version    : $($PSVersionTable.PSVersion)"
Write-Log "Script Path   : $($script:ScriptFullPath)"
Write-Log "Script Dir    : $($script:ScriptDir)"

# Step 2: Ensure 64-bit execution
Assert-64BitExecution

# Step 3: Verify administrator privileges
Assert-AdminPrivileges

# Step 4: Log OS build information
$null = Get-OSBuildInfo

# Step 5: Check for pending reboot
if (Test-PendingReboot) {
    Write-Log "A pending reboot was detected. Installation may fail. Proceeding with caution..." -Level WARN
    # Uncomment the following to block installation when reboot is pending:
    # Write-Log "Aborting installation due to pending reboot." -Level ERROR
    # Exit 3010
}

# Step 6: Check disk space
if (-not (Test-DiskSpace -MinimumGB $minDiskSpaceGB)) {
    Write-Log "Aborting due to insufficient disk space." -Level ERROR
    Exit 1
}

# Step 7: Locate MSU file in script directory
$allMSUs = Get-ChildItem -Path $script:ScriptDir -Filter "*.msu" -ErrorAction SilentlyContinue

if (-not $allMSUs -or $allMSUs.Count -eq 0) {
    Write-Log "No MSU file found in script directory: $($script:ScriptDir)" -Level ERROR
    Exit 1
}

if ($allMSUs.Count -gt 1) {
    Write-Log "Multiple MSU files detected ($($allMSUs.Count)). Only the first will be installed." -Level WARN
    $allMSUs | ForEach-Object { Write-Log "  Found: $($_.Name)" }
}

$MSUFile = $allMSUs | Select-Object -First 1

Write-Log "Selected MSU file: $($MSUFile.FullName)"
Write-Log "MSU file size: $([math]::Round($MSUFile.Length / 1MB, 2)) MB"

# Step 8: Install via DISM
try {
    Write-Log "Starting DISM installation..."
    Write-Log "DISM log will be written to: $dismLogFile"

    $dismPath = "$env:SystemRoot\System32\Dism.exe"
    $arguments = "/Online /Add-Package /PackagePath:`"$($MSUFile.FullName)`" /Quiet /NoRestart /LogPath:`"$dismLogFile`""

    Write-Log "Executing: $dismPath $arguments"

    $process = Start-Process -FilePath $dismPath `
        -ArgumentList $arguments `
        -PassThru -NoNewWindow

    $completed = $process.WaitForExit($dismTimeoutMs)

    if (-not $completed) {
        Write-Log "DISM process timed out after $([math]::Round($dismTimeoutMs / 60000)) minutes. Terminating process." -Level ERROR
        try { $process.Kill() } catch { Write-Log "Failed to kill DISM process: $($_.Exception.Message)" -Level ERROR }
        Write-Log "=========================================================="
        Exit 1
    }

    $exitCode = $process.ExitCode
    $exitDescription = Get-ExitCodeDescription -ExitCode $exitCode

    Write-Log "DISM completed. Exit code: $exitCode - $exitDescription"

    # Evaluate result
    switch ($exitCode) {
        0 {
            Write-Log "MSU file installed successfully. No reboot required."
            Write-Log "=========================================================="
            Exit 0
        }
        3010 {
            Write-Log "MSU file installed successfully. A reboot is required to complete the installation." -Level WARN
            Write-Log "=========================================================="
            Exit 3010
        }
        1641 {
            Write-Log "MSU file installed successfully. Reboot was initiated." -Level WARN
            Write-Log "=========================================================="
            Exit 0
        }
        -2146498504 {
            # 0x800F0838 - Missing prerequisite baselines
            Write-Log "FAILURE: Missing prerequisite baseline updates." -Level ERROR
            Write-Log "The target update requires intermediate cumulative updates to be installed first." -Level ERROR
            Write-Log "Review the DISM log at $dismLogFile for 'CheckIfIntermediateBaselinesMissing' details." -Level ERROR
            Write-Log "=========================================================="
            Exit $exitCode
        }
        -2146498530 {
            # 0x800F081E - Not applicable
            Write-Log "The update is not applicable to this system (already installed or wrong OS version)." -Level WARN
            Write-Log "=========================================================="
            Exit 0  # Treat as success — update not needed
        }
        -2146498512 {
            # 0x800F0830 - Already installed
            Write-Log "The update is already installed on this system." -Level WARN
            Write-Log "=========================================================="
            Exit 0  # Treat as success
        }
        default {
            Write-Log "Installation failed with error code: $exitCode - $exitDescription" -Level ERROR
            Write-Log "Review the DISM log for details: $dismLogFile" -Level ERROR
            Write-Log "=========================================================="
            Exit $exitCode
        }
    }
}
catch {
    Write-Log "An exception occurred during the installation process." -Level ERROR
    Write-Log "Exception: $($_.Exception.Message)" -Level ERROR
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level ERROR
    Write-Log "=========================================================="
    Exit 1
}