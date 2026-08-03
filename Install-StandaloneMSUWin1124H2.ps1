#Requires -RunAsAdministrator
#Requires -Version 5.1

<#
.SYNOPSIS
    Installs two Microsoft Update Standalone Packages (MSU) in sequence on an online Windows PC.

.DESCRIPTION
    This script automates the sequential installation of two MSU files using DISM.
    The prerequisite update (KB5043080) is always installed first, followed by the
    second MSU found in the same directory.

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
                   Moved log directory creation out of Write-Log to avoid per-call overhead
                   and race conditions.
                   Added multi-MSU detection warning.
                   Added machine context logging (computer name, user, PS version, script path).
                   Added #Requires statements for prerequisites enforcement.
    Version 3.1 - Added sequential two-MSU installation with prerequisite ordering.
                   KB5043080 is always installed first (exact filename match), followed by
                   the second MSU in the script directory.
                   Extracted DISM installation into reusable Install-MSUPackage function.
                   Added strict validation requiring exactly two MSU files.
                   Added exit code aggregation across both installations.
                   Acceptable proceed codes from MSU 1: 0, 3010, 1641, already installed,
                   not applicable. All others halt execution.
                   3010 from either MSU propagates as final exit code.
#>

#---------------------------------------------------------------------#
# CONFIGURATION
#---------------------------------------------------------------------#

$logFilePath = "C:\Windows\fndr\logs"
$scriptLogFile = "$logFilePath\Install-StandaloneMSU.log"
$dismLogFile = "$logFilePath\Install-StandaloneMSU-DISM.log"
$minDiskSpaceGB = 5
$dismTimeoutMs = 21600000  # 360 minutes

# Prerequisite MSU — exact filename
$prerequisiteMSU = "windows11.0-kb5043080-x64_953449672073f8fb99badb4cc6d5d7849b9c83e8.msu"

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
# FUNCTION: Install a Single MSU Package via DISM
#---------------------------------------------------------------------#
function Install-MSUPackage {
    param (
        [Parameter(Mandatory)]
        [string]$MsuPath,

        [Parameter(Mandatory)]
        [string]$Label,

        [Parameter(Mandatory)]
        [string]$DismLogPath,

        [int]$TimeoutMs = 3600000
    )

    Write-Log "----------------------------------------------------------"
    Write-Log "[$Label] Starting installation..."
    Write-Log "[$Label] MSU: $MsuPath"
    Write-Log "[$Label] Size: $([math]::Round((Get-Item $MsuPath).Length / 1MB, 2)) MB"
    Write-Log "[$Label] DISM log: $DismLogPath"

    $dismPath = "$env:SystemRoot\System32\Dism.exe"
    $arguments = "/Online /Add-Package /PackagePath:`"$MsuPath`" /Quiet /NoRestart /LogPath:`"$DismLogPath`""

    Write-Log "[$Label] Executing: $dismPath $arguments"

    $process = Start-Process -FilePath $dismPath `
        -ArgumentList $arguments `
        -PassThru -NoNewWindow

    $completed = $process.WaitForExit($TimeoutMs)

    if (-not $completed) {
        Write-Log "[$Label] DISM process timed out after $([math]::Round($TimeoutMs / 60000)) minutes. Terminating process." -Level ERROR
        try { $process.Kill() } catch {
            Write-Log "[$Label] Failed to kill DISM process: $($_.Exception.Message)" -Level ERROR
        }
        return @{ ExitCode = 1; TimedOut = $true }
    }

    $exitCode = $process.ExitCode
    $exitDescription = Get-ExitCodeDescription -ExitCode $exitCode

    Write-Log "[$Label] DISM completed. Exit code: $exitCode - $exitDescription"

    return @{ ExitCode = $exitCode; TimedOut = $false }
}

#---------------------------------------------------------------------#
# FUNCTION: Evaluate if an Exit Code is Acceptable to Proceed
#---------------------------------------------------------------------#
function Test-AcceptableExitCode {
    param ([int]$ExitCode)

    $acceptableCodes = @(0, 3010, 1641, -2146498512, -2146498530)
    return ($acceptableCodes -contains $ExitCode)
}

#---------------------------------------------------------------------#
# MAIN EXECUTION
#---------------------------------------------------------------------#

Write-Log "=========================================================="
Write-Log "Install-StandaloneMSU v3.1 - Starting"
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
}

# Step 6: Check disk space
if (-not (Test-DiskSpace -MinimumGB $minDiskSpaceGB)) {
    Write-Log "Aborting due to insufficient disk space." -Level ERROR
    Exit 1
}

# Step 7: Locate and validate MSU files in script directory
Write-Log "Scanning for MSU files in: $($script:ScriptDir)"

$allMSUs = Get-ChildItem -Path $script:ScriptDir -Filter "*.msu" -ErrorAction SilentlyContinue

if (-not $allMSUs -or $allMSUs.Count -eq 0) {
    Write-Log "No MSU files found in script directory: $($script:ScriptDir)" -Level ERROR
    Exit 1
}

if ($allMSUs.Count -ne 2) {
    Write-Log "Expected exactly 2 MSU files but found $($allMSUs.Count):" -Level ERROR
    $allMSUs | ForEach-Object { Write-Log "  - $($_.Name)" -Level ERROR }
    Write-Log "Aborting. The script requires the prerequisite (KB5043080) and one target update." -Level ERROR
    Exit 1
}

# Identify prerequisite and target MSUs
$msuPrereq = $allMSUs | Where-Object { $_.Name -eq $prerequisiteMSU }
$msuTarget = $allMSUs | Where-Object { $_.Name -ne $prerequisiteMSU }

if (-not $msuPrereq) {
    Write-Log "Prerequisite MSU not found. Expected exact filename:" -Level ERROR
    Write-Log "  $prerequisiteMSU" -Level ERROR
    Write-Log "Files found:" -Level ERROR
    $allMSUs | ForEach-Object { Write-Log "  - $($_.Name)" -Level ERROR }
    Exit 1
}

if (-not $msuTarget) {
    Write-Log "Target MSU not found. Only the prerequisite MSU is present. Need a second MSU." -Level ERROR
    Exit 1
}

Write-Log "Prerequisite MSU : $($msuPrereq.Name)"
Write-Log "Target MSU       : $($msuTarget.Name)"
Write-Log "Installation order: Prerequisite first, then target."

# Step 8: Sequential installation
$rebootRequired = $false

try {

    # ---- MSU 1 of 2: Prerequisite (KB5043080) ----
    $result1 = Install-MSUPackage `
        -MsuPath $msuPrereq.FullName `
        -Label "MSU 1 of 2 - Prerequisite" `
        -DismLogPath $dismLogFile `
        -TimeoutMs $dismTimeoutMs

    if ($result1.TimedOut) {
        Write-Log "Prerequisite installation timed out. Aborting." -Level ERROR
        Write-Log "=========================================================="
        Exit 1
    }

    $exitCode1 = $result1.ExitCode

    # Track reboot requirement
    if ($exitCode1 -eq 3010 -or $exitCode1 -eq 1641) {
        $rebootRequired = $true
        Write-Log "[MSU 1 of 2] Reboot required flag set. Continuing to MSU 2..." -Level WARN
    }

    # Log friendly status for already-installed and not-applicable
    if ($exitCode1 -eq -2146498512) {
        Write-Log "[MSU 1 of 2] Prerequisite is already installed. Continuing to MSU 2..."
    }

    if ($exitCode1 -eq -2146498530) {
        Write-Log "[MSU 1 of 2] Prerequisite reported as not applicable. Continuing to MSU 2..."
    }

    # Check if we can proceed
    if (-not (Test-AcceptableExitCode -ExitCode $exitCode1)) {
        $desc = Get-ExitCodeDescription -ExitCode $exitCode1
        Write-Log "[MSU 1 of 2] Prerequisite failed with code: $exitCode1 - $desc" -Level ERROR
        Write-Log "Cannot proceed to target MSU. Aborting." -Level ERROR
        Write-Log "Review DISM log: $dismLogFile" -Level ERROR
        Write-Log "=========================================================="
        Exit $exitCode1
    }

    Write-Log "[MSU 1 of 2] Prerequisite completed successfully. Proceeding to target update..."

    # ---- MSU 2 of 2: Target Update ----
    $result2 = Install-MSUPackage `
        -MsuPath $msuTarget.FullName `
        -Label "MSU 2 of 2 - Target" `
        -DismLogPath $dismLogFile `
        -TimeoutMs $dismTimeoutMs

    if ($result2.TimedOut) {
        Write-Log "Target MSU installation timed out." -Level ERROR
        Write-Log "=========================================================="
        Exit 1
    }

    $exitCode2 = $result2.ExitCode

    # Track reboot requirement
    if ($exitCode2 -eq 3010 -or $exitCode2 -eq 1641) {
        $rebootRequired = $true
    }

    # Evaluate MSU 2 result
    switch ($exitCode2) {
        0 {
            Write-Log "[MSU 2 of 2] Target update installed successfully."
        }
        3010 {
            Write-Log "[MSU 2 of 2] Target update installed successfully. Reboot required." -Level WARN
        }
        1641 {
            Write-Log "[MSU 2 of 2] Target update installed successfully. Reboot initiated." -Level WARN
        }
        -2146498512 {
            Write-Log "[MSU 2 of 2] Target update is already installed." -Level WARN
        }
        -2146498530 {
            Write-Log "[MSU 2 of 2] Target update is not applicable to this system." -Level WARN
        }
        default {
            $desc = Get-ExitCodeDescription -ExitCode $exitCode2
            Write-Log "[MSU 2 of 2] Target update failed with code: $exitCode2 - $desc" -Level ERROR
            Write-Log "Review DISM log: $dismLogFile" -Level ERROR
            Write-Log "=========================================================="
            Exit $exitCode2
        }
    }

    # ---- Final Result ----
    Write-Log "=========================================================="
    Write-Log "Both MSU packages processed successfully."
    Write-Log "  MSU 1 (Prerequisite) exit code: $exitCode1"
    Write-Log "  MSU 2 (Target)       exit code: $exitCode2"

    if ($rebootRequired) {
        Write-Log "A reboot is required to complete the installation." -Level WARN
        Write-Log "=========================================================="
        Exit 3010
    }

    Write-Log "No reboot required. All updates applied."
    Write-Log "=========================================================="
    Exit 0

}
catch {
    Write-Log "An exception occurred during the installation process." -Level ERROR
    Write-Log "Exception: $($_.Exception.Message)" -Level ERROR
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level ERROR
    Write-Log "=========================================================="
    Exit 1
}