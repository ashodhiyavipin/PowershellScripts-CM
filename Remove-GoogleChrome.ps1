#Requires -Version 5.1
#Requires -RunAsAdministrator

#region COMMENT-BASED HELP
<#
.SYNOPSIS
    Completely removes Google Chrome and all associated artifacts from a Windows machine.

.DESCRIPTION
    This script performs an exhaustive removal of Google Chrome Enterprise including:
    - Graceful uninstall via registered uninstaller and MSI
    - Force termination of all Chrome-related processes
    - Removal of all Chrome Windows services
    - Removal of all Chrome scheduled tasks
    - Complete file system cleanup (all user profiles, temp, prefetch, shortcuts)
    - Complete registry cleanup (HKLM, all HKCU hives, offline NTUSER.DAT, Default User)
    - WMI / Windows Installer registration cleanup
    - Firewall rule removal
    - Post-removal validation with PASS/FAIL reporting
    
    Designed for deployment via SCCM Task Sequence (SYSTEM context) or manual execution.

.PARAMETER DryRun
    Audit-only mode. Logs all actions that would be taken without making any changes.

.PARAMETER LogPath
    Override the default log file directory. Default: C:\Windows\fndr\

.PARAMETER Force
    Suppresses confirmation prompts for silent SCCM execution.

.PARAMETER ExcludeUserData
    Skips per-user profile data cleanup. Useful if user data must be preserved.

.EXAMPLE
    .\Remove-GoogleChrome.ps1
    Runs full removal with verbose logging.

.EXAMPLE
    .\Remove-GoogleChrome.ps1 -DryRun
    Audit-only mode — logs what would be removed without making changes.

.EXAMPLE
    .\Remove-GoogleChrome.ps1 -Force
    Silent execution for SCCM task sequence deployment.
    Full removal (live execution as SYSTEM or Admin)
    .\Remove-GoogleChrome.ps1 -Force
    # Audit-only mode (see what WOULD be removed)
    .\Remove-GoogleChrome.ps1 -DryRun
    # Full removal with custom log path
    .\Remove-GoogleChrome.ps1 -Force -LogPath "D:\Logs\ChromeRemoval"
    # Full removal but preserve user data (bookmarks, etc.)
    .\Remove-GoogleChrome.ps1 -Force -ExcludeUserData

.NOTES
    Script Name    : Remove-GoogleChrome.ps1
    Version        : 1.1.0
    Author         : Vipin Anand Ashodhiya
    Created        : 06-04-2026
    Last Modified  : 14-04-2026
    PS Version     : 5.1+
    Context        : Runs as SYSTEM (SCCM) or local Administrator
    Log Location   : C:\Windows\fndr\Remove-GoogleChrome_<timestamp>.log
    Exit Codes     : 0 = Full Success | 3010 = Partial Success (Reboot Pending) | 2 = Critical Failure

    CHANGELOG:
    v1.1.0 (14-04-2026)
      - Phase Reordering: Graceful MSI/EXE Uninstall executed before Service and Task removal
      - Fix PowerShell 5.1 array-unwrapping bug in phase 1 detection preventing correct count display
      - Change partial success exit code to 3010 for valid SCCM soft-reboot queuing
      - Various logical, behavioral, and reliability improvements from code review
    v1.0.0 (06-04-2026)
      - Initial standalone release
#>

#endregion

#region PARAMETERS
# ============================================================================
# PARAMETERS
# ============================================================================

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, HelpMessage = "Audit-only mode. No changes are made.")]
    [switch]$DryRun,

    [Parameter(Mandatory = $false, HelpMessage = "Override default log directory path.")]
    [ValidateNotNullOrEmpty()]
    [string]$LogPath = "C:\Windows\fndr",

    [Parameter(Mandatory = $false, HelpMessage = "Suppress confirmation prompts for silent execution.")]
    [switch]$Force,

    [Parameter(Mandatory = $false, HelpMessage = "Skip per-user profile data cleanup.")]
    [switch]$ExcludeUserData
)

#endregion

#region CONFIGURATION
# ============================================================================
# CONFIGURATION — Constants and Script-Scoped Variables
# ============================================================================

# --- Script Identity ---
$SCRIPT_NAME = "Remove-GoogleChrome"
$SCRIPT_VERSION = "1.1.0"

# --- Timestamp ---
$SCRIPT_TIMESTAMP = Get-Date -Format "yyyyMMdd_HHmmss"
$SCRIPT_START_TIME = Get-Date

# --- Log Configuration ---
$LOG_DIRECTORY = $LogPath
$LOG_FILE_NAME = "${SCRIPT_NAME}_${SCRIPT_TIMESTAMP}.log"
$LOG_FILE_PATH = Join-Path -Path $LOG_DIRECTORY -ChildPath $LOG_FILE_NAME
$TRANSCRIPT_FILE_NAME = "${SCRIPT_NAME}_Transcript_${SCRIPT_TIMESTAMP}.log"
$TRANSCRIPT_FILE_PATH = Join-Path -Path $LOG_DIRECTORY -ChildPath $TRANSCRIPT_FILE_NAME

# --- Chrome Process Names ---
$CHROME_PROCESS_NAMES = @(
    "chrome"
    "GoogleUpdate"
    "GoogleCrashHandler"
    "GoogleCrashHandler64"
    "notification_helper"
    "elevation_service"
)

# --- Chrome Service Names ---
$CHROME_SERVICE_NAMES = @(
    "GoogleChromeElevationService"
    "gupdate"
    "gupdatem"
)

# --- Chrome Scheduled Task Patterns ---
$CHROME_TASK_PATTERNS = @(
    "GoogleUpdateTaskMachine*"
    "GoogleUpdateTaskUser*"
    "GoogleChrome*"
)

# --- Machine-Level File Paths to Remove ---
$CHROME_MACHINE_PATHS = @(
    "$env:ProgramFiles\Google\Chrome"
    "${env:ProgramFiles(x86)}\Google\Chrome"
    "$env:ProgramData\Google\Chrome"
    "${env:ProgramFiles(x86)}\Google\Update"
    "$env:ProgramFiles\Google\Update"
    "$env:ProgramData\Google\Update"
    "${env:ProgramFiles(x86)}\Google\CrashReports"
    "$env:ProgramFiles\Google\CrashReports"
    "$env:ProgramData\Google\CrashReports"
    "${env:ProgramFiles(x86)}\Google\Policies"
)

# --- Per-User Relative Paths (appended to each profile path) ---
$CHROME_USER_RELATIVE_PATHS = @(
    "AppData\Local\Google\Chrome"
    "AppData\Local\Google\Update"
    "AppData\Local\Google\CrashReports"
    "AppData\Local\Google\Software Reporter Tool"
)

# --- Temp File Patterns ---
$CHROME_TEMP_PATTERNS = @(
    "CR_*"
    "chrome_*"
    "scoped_dir*"
    "Google Chrome*"
)

# --- Prefetch File Patterns ---
$CHROME_PREFETCH_PATTERNS = @(
    "CHROME.EXE-*.pf"
    "GOOGLEUPDATE.EXE-*.pf"
    "GOOGLECRASHHANDLER.EXE-*.pf"
    "GOOGLECRASHHANDLER64.EXE-*.pf"
    # "SETUP.EXE-*.pf"  # Removed: overly broad — matches any setup.exe, not just Google's
)

# --- Shortcut Names to Search For ---
$CHROME_SHORTCUT_NAMES = @(
    "Google Chrome.lnk"
    "Google Chrome.url"
)

# --- HKLM Registry Keys to Remove (full keys) ---
$CHROME_HKLM_KEYS = @(
    # Uninstall entries
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome"
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome"
    # Google product keys
    "HKLM:\SOFTWARE\Google\Chrome"
    "HKLM:\SOFTWARE\WOW6432Node\Google\Chrome"
    "HKLM:\SOFTWARE\Google\Update"
    "HKLM:\SOFTWARE\WOW6432Node\Google\Update"
    "HKLM:\SOFTWARE\Google\No Reporting"
    # App registration
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe"
    "HKLM:\SOFTWARE\Clients\StartMenuInternet\Google Chrome"
    # File / protocol associations
    "HKLM:\SOFTWARE\Classes\ChromeHTML"
    "HKLM:\SOFTWARE\Classes\chrome"
    # Native Messaging
    "HKLM:\SOFTWARE\Google\Chrome\NativeMessagingHosts"
    # Policies
    "HKLM:\SOFTWARE\Policies\Google\Chrome"
    "HKLM:\SOFTWARE\Policies\Google\Update"
    # Residual service keys
    "HKLM:\SYSTEM\CurrentControlSet\Services\GoogleChromeElevationService"
    "HKLM:\SYSTEM\CurrentControlSet\Services\gupdate"
    "HKLM:\SYSTEM\CurrentControlSet\Services\gupdatem"
)

# --- HKLM Registry Values to Remove (specific values from shared keys) ---
# Format: @{ Path = "registry path"; Name = "value name" }
$CHROME_HKLM_VALUES = @(
    @{ Path = "HKLM:\SOFTWARE\RegisteredApplications"; Name = "Google Chrome" }
)

# --- Per-User Registry Keys (relative, under HKU:\<SID> or loaded hive root) ---
$CHROME_USER_REGISTRY_KEYS = @(
    "SOFTWARE\Google\Chrome"
    "SOFTWARE\Google\Update"
    "SOFTWARE\Google\No Reporting"
    "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome"
    "SOFTWARE\Classes\ChromeHTML"
    "SOFTWARE\Classes\chrome"
    "SOFTWARE\Policies\Google\Chrome"
    "SOFTWARE\Policies\Google\Update"
    "SOFTWARE\Google\Chrome\NativeMessagingHosts"
)

# --- Per-User Registry Run Values to Check ---
$CHROME_USER_RUN_VALUE_PATTERNS = @(
    "Google*"
    "Chrome*"
)

# --- File Extension OpenWithProgids to Clean ---
$CHROME_OPENWITH_EXTENSIONS = @(
    ".htm"
    ".html"
    ".pdf"
    ".svg"
    ".xht"
    ".xhtml"
    ".webp"
    ".shtml"
)

# --- Parent Folders to Clean if Empty ---
$GOOGLE_PARENT_FOLDERS = @(
    "$env:ProgramFiles\Google"
    "${env:ProgramFiles(x86)}\Google"
    "$env:ProgramData\Google"
)

# --- Parent Registry Keys to Clean if Empty ---
$GOOGLE_PARENT_REGISTRY_KEYS = @(
    "HKLM:\SOFTWARE\Google"
    "HKLM:\SOFTWARE\WOW6432Node\Google"
    "HKLM:\SOFTWARE\Policies\Google"
)

# --- Script State Variables ---
$script:errorCollection = [System.Collections.ArrayList]::new()
$script:phaseResults = [ordered]@{}
$script:totalItemsProcessed = 0
$script:totalItemsRemoved = 0
$script:totalItemsSkipped = 0
$script:totalItemsFailed = 0
$script:exitCode = 0
$script:logFileWriteFailed = $false
$script:installationType = "Unknown"
$script:chromeVersions = @()
$script:perMachineUninstaller = $null
$script:userProfiles = @()
$script:runningProcesses = @()
$script:chromeServices = @()
$script:chromeTasks = @()
$script:chromeFirewallRules = @()
$script:validationResults = [ordered]@{}

#endregion

#region LOGGING FRAMEWORK
# ============================================================================
# LOGGING FRAMEWORK
# ============================================================================

function Write-Log {
    <#
    .SYNOPSIS
        Writes a timestamped, categorized log entry to both file and console.
    .PARAMETER Message
        The log message text.
    .PARAMETER Category
        Log category: INFO, VERBOSE, WARNING, ERROR, SUCCESS, DRYRUN, SKIP, SECTION.
    .PARAMETER NoNewLine
        Suppresses the trailing newline (for progress-style output).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Message,

        [Parameter(Mandatory = $false, Position = 1)]
        [ValidateSet("INFO", "VERBOSE", "WARNING", "ERROR", "SUCCESS", "DRYRUN", "SKIP", "SECTION")]
        [string]$Category = "INFO",

        [Parameter(Mandatory = $false)]
        [switch]$NoNewLine
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Category] $Message"

    # Write to log file
    try {
        $logEntry | Out-File -FilePath $LOG_FILE_PATH -Append -Encoding UTF8 -ErrorAction SilentlyContinue
    }
    catch {
        # If log file write fails, warn once and continue with console output only
        if (-not $script:logFileWriteFailed) {
            $script:logFileWriteFailed = $true
            Write-Host "[WARNING] Log file write failed — subsequent entries will be console-only." -ForegroundColor Yellow
        }
    }

    # Console color mapping
    $consoleColor = switch ($Category) {
        "INFO" { "White" }
        "VERBOSE" { "Gray" }
        "WARNING" { "Yellow" }
        "ERROR" { "Red" }
        "SUCCESS" { "Green" }
        "DRYRUN" { "Cyan" }
        "SKIP" { "DarkGray" }
        "SECTION" { "Magenta" }
        default { "White" }
    }

    # Write to console
    if ($NoNewLine) {
        Write-Host $logEntry -ForegroundColor $consoleColor -NoNewline
    }
    else {
        Write-Host $logEntry -ForegroundColor $consoleColor
    }
}

function Write-LogSection {
    <#
    .SYNOPSIS
        Writes a formatted section header to the log for visual separation.
    .PARAMETER Title
        The section title.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title
    )

    $separator = "=" * 80
    Write-Log $separator "SECTION"
    Write-Log "  $Title" "SECTION"
    Write-Log $separator "SECTION"
}

function Write-LogSubSection {
    <#
    .SYNOPSIS
        Writes a formatted sub-section header to the log.
    .PARAMETER Title
        The sub-section title.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title
    )

    $separator = "-" * 60
    Write-Log $separator "INFO"
    Write-Log "  $Title" "INFO"
    Write-Log $separator "INFO"
}

function Initialize-Logging {
    <#
    .SYNOPSIS
        Creates log directory, initializes log file, and starts transcript.
    #>
    [CmdletBinding()]
    param()

    # Create log directory if it doesn't exist
    try {
        if (-not (Test-Path -Path $LOG_DIRECTORY -PathType Container)) {
            New-Item -Path $LOG_DIRECTORY -ItemType Directory -Force | Out-Null
        }
    }
    catch {
        Write-Host "[CRITICAL] Failed to create log directory: $LOG_DIRECTORY — $_" -ForegroundColor Red
        exit 2
    }

    # Initialize log file with header
    try {
        $header = @"
================================================================================
  $SCRIPT_NAME v$SCRIPT_VERSION
  Execution started: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
  Log File: $LOG_FILE_PATH
================================================================================
"@
        $header | Out-File -FilePath $LOG_FILE_PATH -Encoding UTF8 -Force
    }
    catch {
        Write-Host "[CRITICAL] Failed to initialize log file: $LOG_FILE_PATH — $_" -ForegroundColor Red
        exit 2
    }

    # Start transcript
    try {
        Start-Transcript -Path $TRANSCRIPT_FILE_PATH -Force -ErrorAction Stop | Out-Null
        Write-Log "Transcript started: $TRANSCRIPT_FILE_PATH" "INFO"
    }
    catch {
        Write-Log "Failed to start transcript: $_. Continuing without transcript." "WARNING"
    }

    Write-Log "Log file initialized: $LOG_FILE_PATH" "INFO"
}

#endregion

#region HELPER FUNCTIONS
# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

function Test-IsAdmin {
    <#
    .SYNOPSIS
        Tests if the current session is running with administrative privileges.
    .OUTPUTS
        [bool] True if running as admin, False otherwise.
    #>
    [CmdletBinding()]
    param()

    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Invoke-Action {
    <#
    .SYNOPSIS
        Wrapper function that gates destructive actions behind DryRun check.
    .PARAMETER Description
        Human-readable description of what the action does.
    .PARAMETER Action
        ScriptBlock containing the destructive action to execute.
    .PARAMETER Phase
        The phase number this action belongs to (for error tracking).
    .PARAMETER Item
        The specific item being acted upon (for error tracking).
    .OUTPUTS
        [bool] True if action succeeded or was skipped (DryRun), False if action failed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Description,

        [Parameter(Mandatory = $true)]
        [scriptblock]$Action,

        [Parameter(Mandatory = $false)]
        [string]$Phase = "Unknown",

        [Parameter(Mandatory = $false)]
        [string]$Item = ""
    )

    $script:totalItemsProcessed++

    if ($DryRun) {
        Write-Log "[WOULD EXECUTE] $Description" "DRYRUN"
        $script:totalItemsSkipped++
        return $true
    }

    try {
        Write-Log "Executing: $Description" "VERBOSE"
        $null = & $Action
        Write-Log "Completed: $Description" "SUCCESS"
        $script:totalItemsRemoved++
        return $true
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Log "FAILED: $Description — $errorMessage" "ERROR"
        $script:totalItemsFailed++
        $null = $script:errorCollection.Add([PSCustomObject]@{
                Timestamp   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Phase       = $Phase
                Item        = if ($Item) { $Item } else { $Description }
                Error       = $errorMessage
                ErrorRecord = $_
            })
        return $false
    }
}

function Test-PathSafe {
    <#
    .SYNOPSIS
        Safely tests if a path exists without throwing terminating errors.
    .PARAMETER Path
        The path to test.
    .OUTPUTS
        [bool] True if path exists, False otherwise.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    try {
        return (Test-Path -Path $Path -ErrorAction SilentlyContinue)
    }
    catch {
        return $false
    }
}

function Remove-DirectoryIfExists {
    <#
    .SYNOPSIS
        Removes a directory and all its contents if it exists. Handles locked files.
    .PARAMETER Path
        The directory path to remove.
    .PARAMETER Phase
        The phase number for error tracking.
    .OUTPUTS
        [bool] True if removed or didn't exist. False if removal failed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $false)]
        [string]$Phase = "Unknown"
    )

    if (-not (Test-PathSafe -Path $Path)) {
        Write-Log "Path not found (skip): $Path" "SKIP"
        return $true
    }

    # Get folder size for logging
    $folderSize = 0
    try {
        $folderSize = (Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue |
            Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
        $folderSizeMB = [math]::Round($folderSize / 1MB, 2)
        Write-Log "Found directory: $Path | Size: ${folderSizeMB} MB" "INFO"
    }
    catch {
        Write-Log "Found directory: $Path | Size: Unable to calculate" "INFO"
    }

    $result = Invoke-Action -Description "Remove directory: $Path" -Action {
        # First attempt: standard removal
        Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
    } -Phase $Phase -Item $Path

    # If first attempt failed, try taking ownership and retrying
    if (-not $result -and -not $DryRun) {
        if ([string]::IsNullOrWhiteSpace($Path)) {
            Write-Log "Cannot retry with ownership change: path is empty or whitespace" "ERROR"
            return $false
        }
        Write-Log "Retrying with ownership change: $Path" "WARNING"
        try {
            # Take ownership
            $null = takeown /F "$Path" /R /A /D Y 2>&1
            # Grant full control
            $null = icacls "$Path" /grant "SYSTEM:(F)" /T /C /Q 2>&1
            # Retry removal
            Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
            Write-Log "Successfully removed after ownership change: $Path" "SUCCESS"
            # Correct counters from the failed Invoke-Action
            if ($script:totalItemsFailed -gt 0) { $script:totalItemsFailed-- }
            $script:totalItemsRemoved++
            return $true
        }
        catch {
            Write-Log "Failed even after ownership change: $Path — $($_.Exception.Message)" "ERROR"
            return $false
        }
    }

    return $result
}

function Remove-FileIfExists {
    <#
    .SYNOPSIS
        Removes a single file if it exists.
    .PARAMETER Path
        The file path to remove.
    .PARAMETER Phase
        The phase number for error tracking.
    .OUTPUTS
        [bool] True if removed or didn't exist. False if removal failed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $false)]
        [string]$Phase = "Unknown"
    )

    if (-not (Test-PathSafe -Path $Path)) {
        Write-Log "File not found (skip): $Path" "SKIP"
        return $true
    }

    Write-Log "Found file: $Path" "INFO"
    return (Invoke-Action -Description "Remove file: $Path" -Action {
            Remove-Item -Path $Path -Force -ErrorAction Stop
        } -Phase $Phase -Item $Path)
}

function Remove-RegistryKeyIfExists {
    <#
    .SYNOPSIS
        Removes a registry key and all subkeys if it exists.
    .PARAMETER Path
        The full registry path (e.g., HKLM:\SOFTWARE\Google\Chrome).
    .PARAMETER Phase
        The phase number for error tracking.
    .OUTPUTS
        [bool] True if removed or didn't exist. False if removal failed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $false)]
        [string]$Phase = "Unknown"
    )

    if (-not (Test-PathSafe -Path $Path)) {
        Write-Log "Registry key not found (skip): $Path" "SKIP"
        return $true
    }

    Write-Log "Found registry key: $Path" "INFO"
    return (Invoke-Action -Description "Remove registry key: $Path" -Action {
            Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
        } -Phase $Phase -Item $Path)
}

function Remove-RegistryValueIfExists {
    <#
    .SYNOPSIS
        Removes a specific value from a registry key if it exists.
    .PARAMETER Path
        The registry key path containing the value.
    .PARAMETER Name
        The name of the value to remove.
    .PARAMETER Phase
        The phase number for error tracking.
    .OUTPUTS
        [bool] True if removed or didn't exist. False if removal failed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $false)]
        [string]$Phase = "Unknown"
    )

    if (-not (Test-PathSafe -Path $Path)) {
        Write-Log "Registry key not found for value removal (skip): $Path\$Name" "SKIP"
        return $true
    }

    try {
        $value = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
        if ($null -eq $value) {
            Write-Log "Registry value not found (skip): $Path\$Name" "SKIP"
            return $true
        }
    }
    catch {
        Write-Log "Registry value not found (skip): $Path\$Name" "SKIP"
        return $true
    }

    Write-Log "Found registry value: $Path\$Name" "INFO"
    return (Invoke-Action -Description "Remove registry value: $Path\$Name" -Action {
            Remove-ItemProperty -Path $Path -Name $Name -Force -ErrorAction Stop
        } -Phase $Phase -Item "$Path\$Name")
}

function Get-AllUserProfiles {
    <#
    .SYNOPSIS
        Enumerates all user profiles on the machine from the ProfileList registry.
        Excludes system profiles (systemprofile, LocalService, NetworkService).
    .OUTPUTS
        [array] Array of PSCustomObjects with SID, ProfilePath, IsLoaded, UserName properties.
    #>
    [CmdletBinding()]
    param()

    $profiles = @()
    $profileListPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"

    # System SIDs to exclude
    $excludeSIDs = @(
        "S-1-5-18"   # SYSTEM / LocalSystem
        "S-1-5-19"   # LocalService
        "S-1-5-20"   # NetworkService
    )

    try {
        $profileKeys = Get-ChildItem -Path $profileListPath -ErrorAction Stop

        foreach ($profileKey in $profileKeys) {
            $sid = $profileKey.PSChildName

            # Skip system accounts
            if ($sid -in $excludeSIDs) {
                Write-Log "Skipping system profile: $sid" "VERBOSE"
                continue
            }

            # Skip short SIDs (not real user profiles)
            if ($sid.Length -lt 8) {
                Write-Log "Skipping non-user SID: $sid" "VERBOSE"
                continue
            }

            try {
                $profilePath = (Get-ItemProperty -Path $profileKey.PSPath -Name "ProfileImagePath" -ErrorAction Stop).ProfileImagePath

                # Check if the profile directory actually exists
                $profileExists = Test-PathSafe -Path $profilePath

                # Check if this hive is currently loaded in HKU
                $isLoaded = Test-PathSafe -Path "Registry::HKEY_USERS\$sid"

                # Try to get username from the profile path
                $userName = Split-Path -Path $profilePath -Leaf

                # Check if NTUSER.DAT exists
                $ntuserPath = Join-Path -Path $profilePath -ChildPath "NTUSER.DAT"
                $hasNtuserDat = Test-PathSafe -Path $ntuserPath

                $profiles += [PSCustomObject]@{
                    SID           = $sid
                    ProfilePath   = $profilePath
                    UserName      = $userName
                    IsLoaded      = $isLoaded
                    ProfileExists = $profileExists
                    HasNtuserDat  = $hasNtuserDat
                    NtuserDatPath = $ntuserPath
                }

                Write-Log "Profile found: $userName (SID: $sid) | Path: $profilePath | Loaded: $isLoaded | Exists: $profileExists" "VERBOSE"
            }
            catch {
                Write-Log "Failed to read profile info for SID $sid — $($_.Exception.Message)" "WARNING"
            }
        }
    }
    catch {
        Write-Log "Failed to enumerate user profiles from registry — $($_.Exception.Message)" "ERROR"
    }

    return $profiles
}

function Mount-OfflineHive {
    <#
    .SYNOPSIS
        Loads an offline NTUSER.DAT registry hive into HKU for editing.
    .PARAMETER HivePath
        Full path to the NTUSER.DAT file.
    .PARAMETER MountName
        The name to mount under HKU (e.g., "TempHive_S-1-5-21-xxxxx").
    .OUTPUTS
        [bool] True if hive was loaded successfully, False otherwise.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$HivePath,

        [Parameter(Mandatory = $true)]
        [string]$MountName
    )

    if (-not (Test-PathSafe -Path $HivePath)) {
        Write-Log "NTUSER.DAT not found: $HivePath" "WARNING"
        return $false
    }

    try {
        Write-Log "Loading offline hive: $HivePath as HKU\$MountName" "VERBOSE"
        $regLoadResult = & reg.exe load "HKU\$MountName" "$HivePath" 2>&1
        $regLoadExitCode = $LASTEXITCODE

        if ($regLoadExitCode -eq 0) {
            Write-Log "Successfully loaded hive: HKU\$MountName" "SUCCESS"
            return $true
        }
        else {
            Write-Log "Failed to load hive (exit code: $regLoadExitCode): $regLoadResult" "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "Exception loading hive: $HivePath — $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Dismount-OfflineHive {
    <#
    .SYNOPSIS
        Unloads a previously mounted offline registry hive from HKU.
        Retries up to 3 times with garbage collection between attempts.
    .PARAMETER MountName
        The mount name used during loading (e.g., "TempHive_S-1-5-21-xxxxx").
    .OUTPUTS
        [bool] True if hive was unloaded successfully, False otherwise.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$MountName
    )

    $maxRetries = 3
    $retryDelay = 2

    for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
        try {
            # Force garbage collection to release any .NET handles on the hive
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            Start-Sleep -Milliseconds 500

            Write-Log "Unloading hive (attempt $attempt of $maxRetries): HKU\$MountName" "VERBOSE"
            $regUnloadResult = & reg.exe unload "HKU\$MountName" 2>&1
            $regUnloadExitCode = $LASTEXITCODE

            if ($regUnloadExitCode -eq 0) {
                Write-Log "Successfully unloaded hive: HKU\$MountName" "SUCCESS"
                return $true
            }
            else {
                Write-Log "Unload attempt $attempt failed (exit code: $regUnloadExitCode): $regUnloadResult" "WARNING"
                if ($attempt -lt $maxRetries) {
                    Write-Log "Waiting $retryDelay seconds before retry..." "VERBOSE"
                    Start-Sleep -Seconds $retryDelay
                }
            }
        }
        catch {
            Write-Log "Exception unloading hive (attempt $attempt): $($_.Exception.Message)" "WARNING"
            if ($attempt -lt $maxRetries) {
                Start-Sleep -Seconds $retryDelay
            }
        }
    }

    Write-Log "FAILED to unload hive after $maxRetries attempts: HKU\$MountName. A reboot may be required to release this hive." "ERROR"
    return $false
}

function Remove-RegistryKeyFromHive {
    <#
    .SYNOPSIS
        Removes a registry key from a mounted hive path (HKU:\SID\relative\path).
    .PARAMETER HiveRoot
        The HKU root for this user (e.g., "Registry::HKEY_USERS\S-1-5-21-xxx" or
        "Registry::HKEY_USERS\TempHive_S-1-5-21-xxx").
    .PARAMETER RelativePath
        The relative key path under the user hive (e.g., "SOFTWARE\Google\Chrome").
    .PARAMETER Phase
        The phase number for error tracking.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$HiveRoot,

        [Parameter(Mandatory = $true)]
        [string]$RelativePath,

        [Parameter(Mandatory = $false)]
        [string]$Phase = "Unknown"
    )

    $fullPath = Join-Path -Path $HiveRoot -ChildPath $RelativePath

    if (-not (Test-PathSafe -Path $fullPath)) {
        Write-Log "User registry key not found (skip): $fullPath" "SKIP"
        return $true
    }

    Write-Log "Found user registry key: $fullPath" "INFO"
    return (Invoke-Action -Description "Remove user registry key: $fullPath" -Action {
            Remove-Item -Path $fullPath -Recurse -Force -ErrorAction Stop
        } -Phase $Phase -Item $fullPath)
}

function Get-FolderIsEmpty {
    <#
    .SYNOPSIS
        Checks if a folder exists and is empty (no files or subfolders).
    .PARAMETER Path
        The folder path to check.
    .OUTPUTS
        [bool] True if folder exists and is empty. False otherwise.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-PathSafe -Path $Path)) {
        return $false
    }

    try {
        $items = Get-ChildItem -Path $Path -Force -ErrorAction SilentlyContinue
        return ($null -eq $items -or $items.Count -eq 0)
    }
    catch {
        return $false
    }
}

function Get-RegistryKeyIsEmpty {
    <#
    .SYNOPSIS
        Checks if a registry key exists and has no subkeys.
    .PARAMETER Path
        The registry key path to check.
    .OUTPUTS
        [bool] True if key exists and has no subkeys. False otherwise.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-PathSafe -Path $Path)) {
        return $false
    }

    try {
        $subkeys = Get-ChildItem -Path $Path -ErrorAction SilentlyContinue
        $hasSubkeys = ($null -ne $subkeys -and $subkeys.Count -gt 0)
        if ($hasSubkeys) { return $false }

        # Also check for registry values (excluding PS* metadata properties)
        try {
            $values = Get-ItemProperty -Path $Path -ErrorAction SilentlyContinue
            if ($null -ne $values) {
                $valueCount = @($values.PSObject.Properties | Where-Object { $_.Name -notmatch "^PS" }).Count
                if ($valueCount -gt 0) { return $false }
            }
        }
        catch {
            # If we can't read values, treat as empty
        }
        return $true
    }
    catch {
        return $false
    }
}

function Resolve-ShortcutTarget {
    <#
    .SYNOPSIS
        Resolves the target path of a .lnk shortcut file.
    .PARAMETER ShortcutPath
        The full path to the .lnk file.
    .OUTPUTS
        [string] The target path, or empty string if resolution fails.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ShortcutPath
    )

    try {
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($ShortcutPath)
        $target = $shortcut.TargetPath
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shortcut) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null
        return $target
    }
    catch {
        return ""
    }
}

#endregion

#region DETECTION FUNCTIONS
# ============================================================================
# DETECTION FUNCTIONS — Used by Phase 1
# ============================================================================

function Get-ChromeInstallation {
    <#
    .SYNOPSIS
        Detects Chrome installation type (per-machine, per-user, both, or none).
        Checks registry uninstall keys and file system for chrome.exe.
    .OUTPUTS
        [PSCustomObject] with InstallationType, PerMachine, PerUser details.
    #>
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        InstallationType      = "None"
        PerMachineFound       = $false
        PerMachineRegPath     = ""
        PerMachineInstallPath = ""
        PerMachineExeExists   = $false
        PerUserInstalls       = @()
    }

    Write-Log "Scanning for per-machine Chrome installation..." "VERBOSE"

    # --- Per-Machine: Check registry ---
    $perMachineRegPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome"
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome"
    )

    foreach ($regPath in $perMachineRegPaths) {
        if (Test-PathSafe -Path $regPath) {
            try {
                $regProperties = Get-ItemProperty -Path $regPath -ErrorAction Stop
                $installLocation = $regProperties.InstallLocation
                $displayName = $regProperties.DisplayName
                $displayVersion = $regProperties.DisplayVersion

                Write-Log "Per-machine registry entry found: $regPath" "INFO"
                Write-Log "  DisplayName: $displayName" "VERBOSE"
                Write-Log "  DisplayVersion: $displayVersion" "VERBOSE"
                Write-Log "  InstallLocation: $installLocation" "VERBOSE"

                $result.PerMachineFound = $true
                $result.PerMachineRegPath = $regPath

                if ($installLocation) {
                    $result.PerMachineInstallPath = $installLocation
                    $chromeExePath = Join-Path -Path $installLocation -ChildPath "chrome.exe"
                    if (Test-PathSafe -Path $chromeExePath) {
                        $result.PerMachineExeExists = $true
                        Write-Log "  chrome.exe confirmed at: $chromeExePath" "VERBOSE"
                    }
                    else {
                        Write-Log "  chrome.exe NOT found at: $chromeExePath (orphaned registry entry)" "WARNING"
                    }
                }

                # Found a valid entry, no need to check further
                break
            }
            catch {
                Write-Log "Failed to read registry properties at $regPath — $($_.Exception.Message)" "WARNING"
            }
        }
    }

    # --- Per-Machine: Also check standard file paths even if registry is missing ---
    if (-not $result.PerMachineFound) {
        $standardPaths = @(
            "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
            "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
        )

        foreach ($exePath in $standardPaths) {
            if (Test-PathSafe -Path $exePath) {
                Write-Log "Per-machine chrome.exe found on disk (no registry entry): $exePath" "WARNING"
                $result.PerMachineFound = $true
                $result.PerMachineInstallPath = Split-Path -Path (Split-Path -Path $exePath -Parent) -Parent
                $result.PerMachineExeExists = $true
                break
            }
        }
    }

    # --- Per-User: Check each user profile ---
    Write-Log "Scanning for per-user Chrome installations..." "VERBOSE"

    foreach ($profile in $script:userProfiles) {
        if (-not $profile.ProfileExists) {
            continue
        }

        $userChromePaths = @(
            Join-Path -Path $profile.ProfilePath -ChildPath "AppData\Local\Google\Chrome\Application\chrome.exe"
        )

        foreach ($userExePath in $userChromePaths) {
            if (Test-PathSafe -Path $userExePath) {
                Write-Log "Per-user Chrome found: $userExePath (User: $($profile.UserName))" "INFO"
                $result.PerUserInstalls += [PSCustomObject]@{
                    UserName    = $profile.UserName
                    SID         = $profile.SID
                    ProfilePath = $profile.ProfilePath
                    ChromeExe   = $userExePath
                }
            }
        }
    }

    # --- Determine installation type ---
    $hasPerMachine = $result.PerMachineFound
    $hasPerUser = ($result.PerUserInstalls.Count -gt 0)

    if ($hasPerMachine -and $hasPerUser) {
        $result.InstallationType = "Both"
    }
    elseif ($hasPerMachine) {
        $result.InstallationType = "PerMachine"
    }
    elseif ($hasPerUser) {
        $result.InstallationType = "PerUser"
    }
    else {
        $result.InstallationType = "None"
    }

    Write-Log "Installation type determined: $($result.InstallationType)" "INFO"
    return $result
}

function Get-ChromeVersion {
    <#
    .SYNOPSIS
        Retrieves Chrome version(s) from registry and file system.
    .OUTPUTS
        [array] Array of PSCustomObjects with Source, Version, Path.
    #>
    [CmdletBinding()]
    param()

    $versions = @()

    # --- From registry ---
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome"
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome"
    )

    foreach ($regPath in $regPaths) {
        if (Test-PathSafe -Path $regPath) {
            try {
                $displayVersion = (Get-ItemProperty -Path $regPath -ErrorAction Stop).DisplayVersion
                if ($displayVersion) {
                    $versions += [PSCustomObject]@{
                        Source  = "Registry"
                        Version = $displayVersion
                        Path    = $regPath
                    }
                    Write-Log "Chrome version from registry: $displayVersion ($regPath)" "VERBOSE"
                }
            }
            catch {
                Write-Log "Could not read version from: $regPath" "VERBOSE"
            }
        }
    }

    # --- From chrome.exe FileVersionInfo ---
    $exePaths = @(
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
    )

    foreach ($exePath in $exePaths) {
        if (Test-PathSafe -Path $exePath) {
            try {
                $fileVersion = (Get-Item -Path $exePath -ErrorAction Stop).VersionInfo.ProductVersion
                if ($fileVersion) {
                    $versions += [PSCustomObject]@{
                        Source  = "FileVersion"
                        Version = $fileVersion
                        Path    = $exePath
                    }
                    Write-Log "Chrome version from file: $fileVersion ($exePath)" "VERBOSE"
                }
            }
            catch {
                Write-Log "Could not read file version from: $exePath" "VERBOSE"
            }
        }
    }

    return $versions
}

function Get-ChromeUninstallString {
    <#
    .SYNOPSIS
        Retrieves Chrome uninstall string(s) from registry and validates if the binary exists.
    .OUTPUTS
        [PSCustomObject] with per-machine and per-user uninstaller details.
    #>
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        PerMachine     = $null
        PerUser        = @()
        MSIProductCode = $null
    }

    # --- Per-Machine Uninstall String ---
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome"
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome"
    )

    foreach ($regPath in $regPaths) {
        if (Test-PathSafe -Path $regPath) {
            try {
                $props = Get-ItemProperty -Path $regPath -ErrorAction Stop
                $uninstallString = $props.UninstallString

                if ($uninstallString) {
                    Write-Log "Per-machine UninstallString found: $uninstallString" "VERBOSE"

                    # Parse the executable path
                    $exePath = ""
                    $arguments = ""

                    if ($uninstallString -match '^"([^"]+)"(.*)$') {
                        $exePath = $Matches[1].Trim()
                        $arguments = $Matches[2].Trim()
                    }
                    elseif ($uninstallString -match '^(\S+)(.*)$') {
                        $exePath = $Matches[1].Trim()
                        $arguments = $Matches[2].Trim()
                    }

                    $binaryExists = Test-PathSafe -Path $exePath

                    # Check if this is an MSI-based uninstall
                    $isMSI = $false
                    if ($uninstallString -match 'MsiExec\.exe' -or $uninstallString -match 'msiexec') {
                        $isMSI = $true
                        # Extract product code
                        if ($uninstallString -match '\{([A-Fa-f0-9\-]+)\}') {
                            $result.MSIProductCode = "{$($Matches[1])}"
                            Write-Log "MSI Product Code detected: $($result.MSIProductCode)" "VERBOSE"
                        }
                    }

                    $result.PerMachine = [PSCustomObject]@{
                        RegPath         = $regPath
                        UninstallString = $uninstallString
                        ExePath         = $exePath
                        Arguments       = $arguments
                        BinaryExists    = $binaryExists
                        IsMSI           = $isMSI
                    }

                    Write-Log "  Executable: $exePath | Exists: $binaryExists | MSI: $isMSI" "VERBOSE"
                    break
                }
            }
            catch {
                Write-Log "Failed to read UninstallString from: $regPath — $($_.Exception.Message)" "WARNING"
            }
        }
    }

    # --- Per-Machine: Also check for MSI product code in Installer registry ---
    if ($null -eq $result.MSIProductCode) {
        try {
            $installerProductsPath = "HKLM:\SOFTWARE\Classes\Installer\Products"
            if (Test-PathSafe -Path $installerProductsPath) {
                $productKeys = Get-ChildItem -Path $installerProductsPath -ErrorAction SilentlyContinue
                foreach ($productKey in $productKeys) {
                    try {
                        $productName = (Get-ItemProperty -Path $productKey.PSPath -Name "ProductName" -ErrorAction SilentlyContinue).ProductName
                        if ($productName -match "Google Chrome") {
                            $packedGuid = $productKey.PSChildName
                            Write-Log "MSI Installer product entry found: $productName (Packed GUID: $packedGuid)" "VERBOSE"

                            # Convert packed GUID to standard GUID format
                            # Packed: reverse first 8, next 4, next 4, then swap pairs for remaining 16
                            if ($packedGuid.Length -eq 32) {
                                $block1 = -join ($packedGuid.Substring(0, 8).ToCharArray()[-1..-8])
                                $block2 = -join ($packedGuid.Substring(8, 4).ToCharArray()[-1..-4])
                                $block3 = -join ($packedGuid.Substring(12, 4).ToCharArray()[-1..-4])
                                $block4 = ""
                                for ($i = 16; $i -lt 32; $i += 2) {
                                    $block4 += $packedGuid[$i + 1]
                                    $block4 += $packedGuid[$i]
                                }
                                $standardGuid = "{$block1-$block2-$block3-$($block4.Substring(0,4))-$($block4.Substring(4,12))}"
                                $result.MSIProductCode = $standardGuid
                                Write-Log "Resolved standard GUID: $standardGuid" "VERBOSE"
                            }
                            break
                        }
                    }
                    catch {
                        continue
                    }
                }
            }
        }
        catch {
            Write-Log "Could not scan Installer\Products for Chrome MSI entries" "VERBOSE"
        }
    }

    # --- Per-User: Log only (we won't execute per-user uninstallers) ---
    foreach ($profile in $script:userProfiles) {
        if (-not $profile.ProfileExists) { continue }

        # Check loaded hives
        if ($profile.IsLoaded) {
            $userUninstallPath = "Registry::HKEY_USERS\$($profile.SID)\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome"
            if (Test-PathSafe -Path $userUninstallPath) {
                try {
                    $userUninstall = (Get-ItemProperty -Path $userUninstallPath -ErrorAction Stop).UninstallString
                    if ($userUninstall) {
                        $result.PerUser += [PSCustomObject]@{
                            UserName        = $profile.UserName
                            SID             = $profile.SID
                            UninstallString = $userUninstall
                            Note            = "Will NOT be executed (SYSTEM context — brute-force cleanup instead)"
                        }
                        Write-Log "Per-user UninstallString found for $($profile.UserName): $userUninstall" "VERBOSE"
                        Write-Log "  NOTE: Per-user uninstaller will be skipped (SYSTEM context)" "VERBOSE"
                    }
                }
                catch {
                    # No uninstall entry for this user
                }
            }
        }
    }

    return $result
}

function Get-GoogleUpdatePresence {
    <#
    .SYNOPSIS
        Detects Google Update presence on the machine.
    .OUTPUTS
        [PSCustomObject] with details about Google Update locations found.
    #>
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        Found         = $false
        RegistryPaths = @()
        FilePaths     = @()
    }

    # Registry locations
    $regPaths = @(
        "HKLM:\SOFTWARE\Google\Update"
        "HKLM:\SOFTWARE\WOW6432Node\Google\Update"
    )

    foreach ($regPath in $regPaths) {
        if (Test-PathSafe -Path $regPath) {
            $result.Found = $true
            $result.RegistryPaths += $regPath
            Write-Log "Google Update registry found: $regPath" "VERBOSE"
        }
    }

    # File locations
    $filePaths = @(
        "${env:ProgramFiles(x86)}\Google\Update\GoogleUpdate.exe"
        "$env:ProgramFiles\Google\Update\GoogleUpdate.exe"
    )

    foreach ($filePath in $filePaths) {
        if (Test-PathSafe -Path $filePath) {
            $result.Found = $true
            $result.FilePaths += $filePath
            Write-Log "Google Update binary found: $filePath" "VERBOSE"
        }
    }

    # Per-user locations
    foreach ($profile in $script:userProfiles) {
        if (-not $profile.ProfileExists) { continue }
        $userUpdatePath = Join-Path -Path $profile.ProfilePath -ChildPath "AppData\Local\Google\Update"
        if (Test-PathSafe -Path $userUpdatePath) {
            $result.Found = $true
            $result.FilePaths += $userUpdatePath
            Write-Log "Google Update per-user folder found: $userUpdatePath" "VERBOSE"
        }
    }

    Write-Log "Google Update detected: $($result.Found)" "INFO"
    return $result
}

function Get-ChromeProcesses {
    <#
    .SYNOPSIS
        Detects all running Chrome-related processes.
    .OUTPUTS
        [array] Array of process objects.
    #>
    [CmdletBinding()]
    param()

    $foundProcesses = @()

    foreach ($processName in $CHROME_PROCESS_NAMES) {
        try {
            $processes = Get-Process -Name $processName -ErrorAction SilentlyContinue
            if ($processes) {
                foreach ($proc in $processes) {
                    $foundProcesses += [PSCustomObject]@{
                        Name      = $proc.Name
                        Id        = $proc.Id
                        Path      = $proc.Path
                        StartTime = $proc.StartTime
                    }
                    Write-Log "Running process found: $($proc.Name) (PID: $($proc.Id)) | Path: $($proc.Path)" "VERBOSE"
                }
            }
        }
        catch {
            # Process not found, that's fine
        }
    }

    # Also check for setup.exe that belongs to Google
    try {
        $setupProcesses = Get-Process -Name "setup" -ErrorAction SilentlyContinue
        if ($setupProcesses) {
            foreach ($proc in $setupProcesses) {
                if ($proc.Path -and $proc.Path -match "Google") {
                    $foundProcesses += [PSCustomObject]@{
                        Name      = $proc.Name
                        Id        = $proc.Id
                        Path      = $proc.Path
                        StartTime = $proc.StartTime
                    }
                    Write-Log "Running Google setup.exe found (PID: $($proc.Id)) | Path: $($proc.Path)" "VERBOSE"
                }
            }
        }
    }
    catch {
        # No setup processes, that's fine
    }

    Write-Log "Total Chrome-related processes found: $($foundProcesses.Count)" "INFO"
    return $foundProcesses
}

function Get-ChromeServices {
    <#
    .SYNOPSIS
        Detects all Chrome-related Windows services.
    .OUTPUTS
        [array] Array of PSCustomObjects with service details.
    #>
    [CmdletBinding()]
    param()

    $foundServices = @()

    # Check known service names
    foreach ($serviceName in $CHROME_SERVICE_NAMES) {
        try {
            $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
            if ($service) {
                # Get binary path from WMI
                $cimService = Get-CimInstance -ClassName Win32_Service -Filter "Name='$serviceName'" -ErrorAction SilentlyContinue
                $binaryPath = if ($cimService) { $cimService.PathName } else { "Unknown" }

                $foundServices += [PSCustomObject]@{
                    Name        = $service.Name
                    DisplayName = $service.DisplayName
                    Status      = $service.Status
                    StartType   = $service.StartType
                    BinaryPath  = $binaryPath
                }
                Write-Log "Service found: $($service.Name) | Status: $($service.Status) | StartType: $($service.StartType) | Path: $binaryPath" "VERBOSE"
            }
        }
        catch {
            # Service not found, that's fine
        }
    }

    # Also scan for any service with binary path containing "Google\Chrome"
    try {
        $allGoogleServices = Get-CimInstance -ClassName Win32_Service -ErrorAction SilentlyContinue |
        Where-Object { $_.PathName -match "Google\\Chrome" -or $_.PathName -match "Google\\Update" }

        if ($allGoogleServices) {
            foreach ($svc in $allGoogleServices) {
                # Avoid duplicates
                if ($foundServices.Name -notcontains $svc.Name) {
                    $foundServices += [PSCustomObject]@{
                        Name        = $svc.Name
                        DisplayName = $svc.DisplayName
                        Status      = $svc.State
                        StartType   = $svc.StartMode
                        BinaryPath  = $svc.PathName
                    }
                    Write-Log "Additional Google service found via path scan: $($svc.Name) | Path: $($svc.PathName)" "VERBOSE"
                }
            }
        }
    }
    catch {
        Write-Log "Could not perform full service path scan: $($_.Exception.Message)" "WARNING"
    }

    Write-Log "Total Chrome-related services found: $($foundServices.Count)" "INFO"
    return $foundServices
}

function Get-ChromeScheduledTasks {
    <#
    .SYNOPSIS
        Detects all Chrome and Google Update scheduled tasks.
    .OUTPUTS
        [array] Array of PSCustomObjects with task details.
    #>
    [CmdletBinding()]
    param()

    $foundTasks = @()

    try {
        $allTasks = Get-ScheduledTask -ErrorAction SilentlyContinue

        if ($allTasks) {
            foreach ($pattern in $CHROME_TASK_PATTERNS) {
                $matchingTasks = $allTasks | Where-Object { $_.TaskName -like $pattern }
                foreach ($task in $matchingTasks) {
                    $foundTasks += [PSCustomObject]@{
                        TaskName = $task.TaskName
                        TaskPath = $task.TaskPath
                        State    = $task.State
                        URI      = $task.URI
                    }
                    Write-Log "Scheduled task found: $($task.TaskName) | Path: $($task.TaskPath) | State: $($task.State)" "VERBOSE"
                }
            }

            # Also check for any task in a Google path
            $googlePathTasks = $allTasks | Where-Object { $_.TaskPath -match "\\Google\\" }
            foreach ($task in $googlePathTasks) {
                if ($foundTasks.TaskName -notcontains $task.TaskName) {
                    $foundTasks += [PSCustomObject]@{
                        TaskName = $task.TaskName
                        TaskPath = $task.TaskPath
                        State    = $task.State
                        URI      = $task.URI
                    }
                    Write-Log "Additional Google task found via path scan: $($task.TaskName) | Path: $($task.TaskPath)" "VERBOSE"
                }
            }
        }
    }
    catch {
        Write-Log "Failed to enumerate scheduled tasks: $($_.Exception.Message)" "WARNING"
    }

    Write-Log "Total Chrome-related scheduled tasks found: $($foundTasks.Count)" "INFO"
    return $foundTasks
}

function Get-ChromeFirewallRules {
    <#
    .SYNOPSIS
        Detects Windows Firewall rules that reference Chrome or Google Update.
    .OUTPUTS
        [array] Array of PSCustomObjects with firewall rule details.
    #>
    [CmdletBinding()]
    param()

    $foundRules = @()

    try {
        # Get all firewall rules with their application filters
        $firewallRules = Get-NetFirewallRule -ErrorAction SilentlyContinue

        if ($firewallRules) {
            foreach ($rule in $firewallRules) {
                try {
                    $appFilter = $rule | Get-NetFirewallApplicationFilter -ErrorAction SilentlyContinue
                    if ($appFilter -and $appFilter.Program) {
                        if ($appFilter.Program -match "chrome\.exe" -or
                            $appFilter.Program -match "Google\\Chrome" -or
                            $appFilter.Program -match "GoogleUpdate\.exe" -or
                            $appFilter.Program -match "Google\\Update") {

                            $foundRules += [PSCustomObject]@{
                                Name        = $rule.Name
                                DisplayName = $rule.DisplayName
                                Direction   = $rule.Direction
                                Action      = $rule.Action
                                Enabled     = $rule.Enabled
                                Program     = $appFilter.Program
                            }
                            Write-Log "Firewall rule found: $($rule.DisplayName) | Direction: $($rule.Direction) | Program: $($appFilter.Program)" "VERBOSE"
                        }
                    }
                }
                catch {
                    # Skip rules we can't read
                    continue
                }
            }
        }
    }
    catch {
        Write-Log "Failed to enumerate firewall rules: $($_.Exception.Message)" "WARNING"
    }

    Write-Log "Total Chrome-related firewall rules found: $($foundRules.Count)" "INFO"
    return $foundRules
}

function Write-DetectionSummary {
    <#
    .SYNOPSIS
        Writes a formatted summary of all detection findings.
    #>
    [CmdletBinding()]
    param()

    Write-LogSubSection "DETECTION SUMMARY"

    Write-Log "Installation Type    : $($script:installationType)" "INFO"

    if ($script:chromeVersions.Count -gt 0) {
        foreach ($ver in $script:chromeVersions) {
            Write-Log "Chrome Version       : $($ver.Version) (Source: $($ver.Source))" "INFO"
        }
    }
    else {
        Write-Log "Chrome Version       : Not detected" "INFO"
    }

    if ($script:perMachineUninstaller) {
        if ($script:perMachineUninstaller.PerMachine) {
            Write-Log "Per-Machine Uninstall: $($script:perMachineUninstaller.PerMachine.UninstallString)" "INFO"
            Write-Log "  Binary Exists      : $($script:perMachineUninstaller.PerMachine.BinaryExists)" "INFO"
            Write-Log "  Is MSI             : $($script:perMachineUninstaller.PerMachine.IsMSI)" "INFO"
        }
        else {
            Write-Log "Per-Machine Uninstall: Not found" "INFO"
        }

        if ($script:perMachineUninstaller.PerUser.Count -gt 0) {
            Write-Log "Per-User Installs    : $($script:perMachineUninstaller.PerUser.Count) found (will use brute-force cleanup)" "INFO"
        }
        else {
            Write-Log "Per-User Installs    : None found" "INFO"
        }

        if ($script:perMachineUninstaller.MSIProductCode) {
            Write-Log "MSI Product Code     : $($script:perMachineUninstaller.MSIProductCode)" "INFO"
        }
    }

    Write-Log "User Profiles        : $($script:userProfiles.Count) total ($( ($script:userProfiles | Where-Object { $_.IsLoaded }).Count ) loaded)" "INFO"
    Write-Log "Running Processes    : $($script:runningProcesses.Count)" "INFO"
    Write-Log "Services             : $($script:chromeServices.Count)" "INFO"
    Write-Log "Scheduled Tasks      : $($script:chromeTasks.Count)" "INFO"
    Write-Log "Firewall Rules       : $($script:chromeFirewallRules.Count)" "INFO"
}

#endregion

#region PHASE 0 — PRE-FLIGHT CHECKS
# ============================================================================
# PHASE 0 — PRE-FLIGHT CHECKS
# ============================================================================

function Invoke-Phase0PreFlight {
    <#
    .SYNOPSIS
        Executes all pre-flight checks: admin rights, PS version, logging init, environment logging.
    .OUTPUTS
        [bool] True if all pre-flight checks pass. False if critical failure.
    #>
    [CmdletBinding()]
    param()

    # --- 0.1 Check Administrator Privileges ---
    if (-not (Test-IsAdmin)) {
        Write-Host "[CRITICAL] This script requires administrative privileges. Please run as Administrator or SYSTEM." -ForegroundColor Red
        return $false
    }

    # --- 0.3 Initialize Logging (must be before we log anything) ---
    Initialize-Logging

    # Now we can use Write-Log
    Write-LogSection "PHASE 0: PRE-FLIGHT CHECKS"

    # Log admin check result
    Write-Log "Administrative privileges confirmed" "SUCCESS"

    # --- 0.2 Check PowerShell Version ---
    $psVersion = $PSVersionTable.PSVersion
    $psVersionString = "$($psVersion.Major).$($psVersion.Minor)"

    if ($psVersion.Major -lt 5 -or ($psVersion.Major -eq 5 -and $psVersion.Minor -lt 1)) {
        Write-Log "PowerShell 5.1 or later is required. Current version: $psVersionString" "ERROR"
        return $false
    }

    Write-Log "PowerShell version: $psVersionString" "SUCCESS"

    # --- 0.4 Log Environment ---
    Write-LogSubSection "ENVIRONMENT DETAILS"

    $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
    $computerInfo = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue

    Write-Log "Script Name          : $SCRIPT_NAME" "INFO"
    Write-Log "Script Version       : $SCRIPT_VERSION" "INFO"
    Write-Log "Script Path          : $($MyInvocation.ScriptName)" "INFO"
    Write-Log "Hostname             : $env:COMPUTERNAME" "INFO"
    Write-Log "Domain               : $(if ($computerInfo) { $computerInfo.Domain } else { 'Unknown' })" "INFO"
    Write-Log "OS                   : $(if ($osInfo) { $osInfo.Caption } else { 'Unknown' })" "INFO"
    Write-Log "OS Version           : $(if ($osInfo) { $osInfo.Version } else { 'Unknown' })" "INFO"
    Write-Log "OS Build             : $(if ($osInfo) { $osInfo.BuildNumber } else { 'Unknown' })" "INFO"
    Write-Log "Architecture         : $env:PROCESSOR_ARCHITECTURE" "INFO"
    Write-Log "Executing User       : $env:USERNAME" "INFO"
    Write-Log "Executing Domain     : $env:USERDOMAIN" "INFO"
    Write-Log "System Directory     : $env:SystemRoot" "INFO"
    Write-Log "PowerShell Version   : $psVersionString" "INFO"
    Write-Log "PowerShell Edition   : $($PSVersionTable.PSEdition)" "INFO"

    # --- 0.5 Log DryRun Status ---
    if ($DryRun) {
        Write-Log "" "INFO"
        Write-Log "╔══════════════════════════════════════════════════════════════╗" "DRYRUN"
        Write-Log "║  DRY RUN MODE ACTIVE — No changes will be made to system   ║" "DRYRUN"
        Write-Log "║  All actions will be logged as [DRYRUN] for review         ║" "DRYRUN"
        Write-Log "╚══════════════════════════════════════════════════════════════╝" "DRYRUN"
        Write-Log "" "INFO"
    }
    else {
        Write-Log "DryRun Mode          : DISABLED (live execution)" "WARNING"
    }

    if ($Force) {
        Write-Log "Force Mode           : ENABLED (no confirmation prompts)" "INFO"
    }

    if ($ExcludeUserData) {
        Write-Log "ExcludeUserData      : ENABLED (per-user data will be skipped)" "INFO"
    }

    # --- 0.6 Initialize Error Collection ---
    Write-Log "Error collection initialized" "VERBOSE"
    Write-Log "Phase results tracker initialized" "VERBOSE"

    Write-Log "Pre-flight checks completed successfully" "SUCCESS"
    $script:phaseResults["Phase 0"] = "PASS"
    return $true
}

#endregion

#region PHASE 1 — DETECTION & INVENTORY
# ============================================================================
# PHASE 1 — DETECTION & INVENTORY
# ============================================================================

function Invoke-Phase1Detection {
    <#
    .SYNOPSIS
        Executes the full detection and inventory phase.
        Populates all script-scoped detection variables.
    .OUTPUTS
        [bool] True if detection completed (even if Chrome not found). False on critical error.
    #>
    [CmdletBinding()]
    param()

    Write-LogSection "PHASE 1: DETECTION & INVENTORY"

    try {
        # --- 1.5 Enumerate All User Profiles (must be done FIRST, other functions depend on it) ---
        Write-LogSubSection "1.5 — Enumerating User Profiles"
        $script:userProfiles = @(Get-AllUserProfiles)

        if ($script:userProfiles.Count -eq 0) {
            Write-Log "No user profiles found on this machine" "WARNING"
        }
        else {
            $loadedCount = ($script:userProfiles | Where-Object { $_.IsLoaded }).Count
            $totalCount = $script:userProfiles.Count
            Write-Log "User profiles enumerated: $totalCount total ($loadedCount currently loaded)" "SUCCESS"
        }

        # --- 1.1 Detect Installation Type ---
        Write-LogSubSection "1.1 — Detecting Chrome Installation Type"
        $installInfo = Get-ChromeInstallation
        $script:installationType = $installInfo.InstallationType

        # --- 1.2 Detect Chrome Version(s) ---
        Write-LogSubSection "1.2 — Detecting Chrome Version(s)"
        $script:chromeVersions = @(Get-ChromeVersion)

        if (-not $script:chromeVersions -or $script:chromeVersions.Count -eq 0) {
            Write-Log "No Chrome version information found" "INFO"
        }

        # --- 1.3 Detect Uninstall String(s) ---
        Write-LogSubSection "1.3 — Detecting Uninstall String(s)"
        $script:perMachineUninstaller = Get-ChromeUninstallString

        # --- 1.4 Detect Google Update ---
        Write-LogSubSection "1.4 — Detecting Google Update"
        $googleUpdateInfo = Get-GoogleUpdatePresence

        # --- 1.6 Detect Running Processes ---
        Write-LogSubSection "1.6 — Detecting Running Processes"
        $script:runningProcesses = @(Get-ChromeProcesses)

        # --- 1.7 Detect Services ---
        Write-LogSubSection "1.7 — Detecting Chrome Services"
        $script:chromeServices = @(Get-ChromeServices)

        # --- 1.8 Detect Scheduled Tasks ---
        Write-LogSubSection "1.8 — Detecting Chrome Scheduled Tasks"
        $script:chromeTasks = @(Get-ChromeScheduledTasks)

        # --- 1.9 Detect Firewall Rules ---
        Write-LogSubSection "1.9 — Detecting Chrome Firewall Rules"
        $script:chromeFirewallRules = @(Get-ChromeFirewallRules)

        # --- 1.10 Log Detection Summary ---
        Write-DetectionSummary

        Write-Log "Detection and inventory phase completed successfully" "SUCCESS"
        $script:phaseResults["Phase 1"] = "PASS"
        return $true
    }
    catch {
        Write-Log "Critical error during detection phase: $($_.Exception.Message)" "ERROR"
        $null = $script:errorCollection.Add([PSCustomObject]@{
                Timestamp   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Phase       = "Phase 1"
                Item        = "Detection & Inventory"
                Error       = $_.Exception.Message
                ErrorRecord = $_
            })
        $script:phaseResults["Phase 1"] = "FAIL"
        return $false
    }
}

#endregion

#region PHASE 2 — PROCESS TERMINATION
# ============================================================================
# PHASE 2 — PROCESS TERMINATION
# ============================================================================

function Stop-ChromeProcesses {
    <#
    .SYNOPSIS
        Force-kills all Chrome-related processes. Uses Stop-Process first,
        falls back to taskkill.exe if processes survive.
    .OUTPUTS
        [PSCustomObject] with counts of killed, skipped, and failed processes.
    #>
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        Killed  = 0
        Skipped = 0
        Failed  = 0
    }

    # --- 2.1 Kill each known Chrome process name ---
    foreach ($processName in $CHROME_PROCESS_NAMES) {
        try {
            $processes = Get-Process -Name $processName -ErrorAction SilentlyContinue
            if (-not $processes) {
                Write-Log "Process not running (skip): $processName" "SKIP"
                $result.Skipped++
                continue
            }

            $processCount = @($processes).Count
            Write-Log "Found $processCount instance(s) of: $processName" "INFO"

            foreach ($proc in $processes) {
                $procDescription = "$($proc.Name) (PID: $($proc.Id))"
                if ($proc.Path) {
                    $procDescription += " | Path: $($proc.Path)"
                }

                $killResult = Invoke-Action -Description "Force-kill process: $procDescription" -Action {
                    Stop-Process -Id $proc.Id -Force -ErrorAction Stop
                } -Phase "Phase 2" -Item $procDescription

                if ($killResult) {
                    $result.Killed++
                }
                else {
                    # Fallback to taskkill
                    Write-Log "Stop-Process failed for $procDescription — attempting taskkill.exe fallback" "WARNING"

                    $taskKillResult = Invoke-Action -Description "taskkill fallback: $procDescription" -Action {
                        $taskKillOutput = & taskkill.exe /F /PID $proc.Id /T 2>&1
                        if ($LASTEXITCODE -ne 0) {
                            throw "taskkill exited with code $LASTEXITCODE — $taskKillOutput"
                        }
                    } -Phase "Phase 2" -Item "taskkill: $procDescription"

                    if ($taskKillResult) {
                        $result.Killed++
                        # Remove the previous error since we recovered
                        $lastError = $script:errorCollection | Where-Object { $_.Item -eq $procDescription } | Select-Object -Last 1
                        if ($lastError) {
                            $script:errorCollection.Remove($lastError)
                            if ($script:totalItemsFailed -gt 0) { $script:totalItemsFailed-- }
                        }
                    }
                    else {
                        $result.Failed++
                    }
                }
            }
        }
        catch {
            Write-Log "Error processing $processName — $($_.Exception.Message)" "ERROR"
            $result.Failed++
        }
    }

    # --- Also kill setup.exe if it belongs to Google ---
    try {
        $setupProcesses = Get-Process -Name "setup" -ErrorAction SilentlyContinue
        if ($setupProcesses) {
            foreach ($proc in $setupProcesses) {
                if ($proc.Path -and $proc.Path -match "Google") {
                    $procDescription = "setup.exe (PID: $($proc.Id)) | Path: $($proc.Path)"
                    Write-Log "Found Google setup.exe: $procDescription" "INFO"

                    $killResult = Invoke-Action -Description "Force-kill process: $procDescription" -Action {
                        Stop-Process -Id $proc.Id -Force -ErrorAction Stop
                    } -Phase "Phase 2" -Item $procDescription

                    if ($killResult) {
                        $result.Killed++
                    }
                    else {
                        $result.Failed++
                    }
                }
            }
        }
    }
    catch {
        Write-Log "Error scanning for Google setup.exe — $($_.Exception.Message)" "WARNING"
    }

    return $result
}

function Invoke-Phase2ProcessTermination {
    <#
    .SYNOPSIS
        Orchestrates Phase 2: kills all Chrome-related processes and waits for handle release.
    .OUTPUTS
        [bool] True if phase completed. False on critical failure.
    #>
    [CmdletBinding()]
    param()

    Write-LogSection "PHASE 2: PROCESS TERMINATION"

    try {
        # --- Check if any processes need killing ---
        if ($script:runningProcesses.Count -eq 0) {
            Write-Log "No Chrome-related processes detected in Phase 1 — performing live recheck" "VERBOSE"
        }

        # --- Always do a live scan and kill (processes may have started since Phase 1) ---
        Write-LogSubSection "2.1 — Force-Killing Chrome Processes"
        $killResults = Stop-ChromeProcesses

        # --- 2.2 Wait for file handle release ---
        Write-LogSubSection "2.2 — Waiting for File Handle Release"

        if ($killResults.Killed -gt 0 -and -not $DryRun) {
            Write-Log "Waiting 3 seconds for file handles to release..." "INFO"
            Start-Sleep -Seconds 3
            Write-Log "Wait complete" "VERBOSE"
        }
        else {
            Write-Log "No processes were killed — skipping wait" "SKIP"
        }

        # --- 2.3 Verify all processes are stopped ---
        Write-LogSubSection "2.3 — Verifying Process Termination"

        $survivingProcesses = @()
        foreach ($processName in $CHROME_PROCESS_NAMES) {
            $remaining = Get-Process -Name $processName -ErrorAction SilentlyContinue
            if ($remaining) {
                foreach ($proc in $remaining) {
                    $survivingProcesses += "$($proc.Name) (PID: $($proc.Id))"
                }
            }
        }

        if ($survivingProcesses.Count -gt 0) {
            Write-Log "WARNING: $($survivingProcesses.Count) Chrome process(es) survived termination:" "WARNING"
            foreach ($survivor in $survivingProcesses) {
                Write-Log "  Still running: $survivor" "WARNING"
            }
            Write-Log "Continuing with cleanup — some files may be locked" "WARNING"
        }
        else {
            if (-not $DryRun) {
                Write-Log "All Chrome processes successfully terminated" "SUCCESS"
            }
        }

        # --- Phase Result ---
        Write-Log ("Phase 2 complete: {0} killed, {1} skipped (not running), {2} failed" -f
            $killResults.Killed, $killResults.Skipped, $killResults.Failed) "INFO"

        $script:phaseResults["Phase 2"] = if ($killResults.Failed -eq 0) { "PASS" } else { "PARTIAL" }
        return $true
    }
    catch {
        Write-Log "Critical error during process termination: $($_.Exception.Message)" "ERROR"
        $null = $script:errorCollection.Add([PSCustomObject]@{
                Timestamp   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Phase       = "Phase 2"
                Item        = "Process Termination Orchestration"
                Error       = $_.Exception.Message
                ErrorRecord = $_
            })
        $script:phaseResults["Phase 2"] = "FAIL"
        return $false
    }
}

#endregion

#region PHASE 4 — SERVICE REMOVAL
# ============================================================================
# PHASE 4 — SERVICE REMOVAL
# ============================================================================

function Remove-SingleService {
    <#
    .SYNOPSIS
        Stops and removes a single Windows service by name. Uses multiple methods
        for robust removal: Stop-Service, sc.exe stop, sc.exe delete, and
        direct registry removal as last resort.
    .PARAMETER ServiceName
        The name of the service to remove.
    .OUTPUTS
        [string] Result: "Removed", "NotFound", or "Failed"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ServiceName
    )

    # --- Detection ---
    $service = $null
    try {
        $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    }
    catch {
        # Service does not exist
    }

    if ($null -eq $service) {
        # Double check via sc.exe (sometimes Get-Service misses services in certain states)
        $scQueryResult = & sc.exe query $ServiceName 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Log "Service not found (skip): $ServiceName" "SKIP"
            return "NotFound"
        }
        else {
            Write-Log "Service found via sc.exe (not visible to Get-Service): $ServiceName" "VERBOSE"
        }
    }
    else {
        Write-Log "Service found: $ServiceName | Status: $($service.Status) | StartType: $($service.StartType)" "INFO"
    }

    # --- Stop the service if it's running ---
    if ($service -and $service.Status -eq 'Running') {
        $stopResult = Invoke-Action -Description "Stop service: $ServiceName" -Action {
            Stop-Service -Name $ServiceName -Force -ErrorAction Stop
            # Wait for the service to actually stop
            $timeout = 30
            $elapsed = 0
            while ($elapsed -lt $timeout) {
                $currentStatus = (Get-Service -Name $ServiceName -ErrorAction SilentlyContinue).Status
                if ($currentStatus -ne 'Running') {
                    break
                }
                Start-Sleep -Seconds 1
                $elapsed++
            }
            if ($elapsed -ge $timeout) {
                Write-Log "Service stop timed out after $timeout seconds: $ServiceName" "WARNING"
            }
        } -Phase "Phase 4" -Item "Stop: $ServiceName"

        if (-not $stopResult -and -not $DryRun) {
            # Fallback: sc.exe stop
            Write-Log "Attempting sc.exe stop fallback for: $ServiceName" "WARNING"
            try {
                $null = & sc.exe stop $ServiceName 2>&1
                Start-Sleep -Seconds 3
            }
            catch {
                Write-Log "sc.exe stop also failed: $($_.Exception.Message)" "WARNING"
            }
        }
    }
    elseif ($service) {
        Write-Log "Service is not running (status: $($service.Status)) — skipping stop" "VERBOSE"
    }

    # --- Disable the service before deletion (safety measure) ---
    if (-not $DryRun) {
        try {
            $null = & sc.exe config $ServiceName start= disabled 2>&1
            Write-Log "Service disabled: $ServiceName" "VERBOSE"
        }
        catch {
            Write-Log "Could not disable service (non-critical): $ServiceName" "VERBOSE"
        }
    }

    # --- Delete the service ---
    $deleteResult = Invoke-Action -Description "Delete service: $ServiceName" -Action {
        $scDeleteOutput = & sc.exe delete $ServiceName 2>&1
        $scDeleteExitCode = $LASTEXITCODE
        if ($scDeleteExitCode -ne 0) {
            throw "sc.exe delete failed with exit code $scDeleteExitCode — $scDeleteOutput"
        }
    } -Phase "Phase 3" -Item "Delete: $ServiceName"

    if ($deleteResult) {
        # Verify the service is gone
        Start-Sleep -Seconds 1
        $verifyService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        if ($null -eq $verifyService) {
            Write-Log "Service successfully removed and verified: $ServiceName" "SUCCESS"
            return "Removed"
        }
        else {
            Write-Log "Service still exists after sc.exe delete (may require reboot): $ServiceName" "WARNING"
        }
    }

    # --- Last resort: direct registry removal ---
    if (-not $DryRun) {
        $serviceRegPath = "HKLM:\SYSTEM\CurrentControlSet\Services\$ServiceName"
        if (Test-PathSafe -Path $serviceRegPath) {
            Write-Log "Attempting direct registry removal as last resort: $serviceRegPath" "WARNING"
            $regResult = Invoke-Action -Description "Registry delete service key: $serviceRegPath" -Action {
                Remove-Item -Path $serviceRegPath -Recurse -Force -ErrorAction Stop
            } -Phase "Phase 3" -Item "RegDelete: $ServiceName"

            if ($regResult) {
                Write-Log "Service registry key removed (reboot needed to complete): $ServiceName" "SUCCESS"
                return "Removed"
            }
        }
    }
    elseif ($DryRun) {
        # In DryRun, the Invoke-Action already logged what would happen
        return "Removed"
    }

    Write-Log "Failed to remove service: $ServiceName" "ERROR"
    return "Failed"
}

function Remove-ChromeServices {
    <#
    .SYNOPSIS
        Removes all Chrome-related Windows services (known names + dynamic discovery).
    .OUTPUTS
        [PSCustomObject] with counts of removed, not found, and failed services.
    #>
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        Removed  = 0
        NotFound = 0
        Failed   = 0
    }

    # --- 4.1 Remove known Chrome services ---
    foreach ($serviceName in $CHROME_SERVICE_NAMES) {
        $serviceResult = Remove-SingleService -ServiceName $serviceName

        switch ($serviceResult) {
            "Removed" { $result.Removed++ }
            "NotFound" { $result.NotFound++ }
            "Failed" { $result.Failed++ }
        }
    }

    # --- 4.2 Dynamic discovery: any other service with Google\Chrome or Google\Update in binary path ---
    Write-LogSubSection "4.2 — Scanning for Additional Google Services"

    try {
        $additionalServices = Get-CimInstance -ClassName Win32_Service -ErrorAction SilentlyContinue |
        Where-Object {
            ($_.PathName -match "Google\\Chrome" -or $_.PathName -match "Google\\Update") -and
            ($_.Name -notin $CHROME_SERVICE_NAMES)
        }

        if ($additionalServices) {
            foreach ($svc in $additionalServices) {
                Write-Log "Additional Google service discovered: $($svc.Name) | Path: $($svc.PathName)" "INFO"
                $serviceResult = Remove-SingleService -ServiceName $svc.Name

                switch ($serviceResult) {
                    "Removed" { $result.Removed++ }
                    "NotFound" { $result.NotFound++ }
                    "Failed" { $result.Failed++ }
                }
            }
        }
        else {
            Write-Log "No additional Google services found via path scan" "SKIP"
        }
    }
    catch {
        Write-Log "Error scanning for additional Google services: $($_.Exception.Message)" "WARNING"
    }

    return $result
}

function Invoke-Phase4ServiceRemoval {
    <#
    .SYNOPSIS
        Orchestrates Phase 4: stops and removes all Chrome-related Windows services.
    .OUTPUTS
        [bool] True if phase completed. False on critical failure.
    #>
    [CmdletBinding()]
    param()

    Write-LogSection "PHASE 4: SERVICE REMOVAL"

    try {
        # --- Pre-check: did Phase 1 find any services? ---
        if ($script:chromeServices.Count -eq 0) {
            Write-Log "No Chrome services were detected in Phase 1 — performing live scan anyway" "VERBOSE"
        }
        else {
            Write-Log "Phase 1 detected $($script:chromeServices.Count) Chrome-related service(s)" "INFO"
            foreach ($svc in $script:chromeServices) {
                Write-Log "  Queued for removal: $($svc.Name) ($($svc.DisplayName)) | Status: $($svc.Status)" "VERBOSE"
            }
        }

        # --- Execute service removal ---
        Write-LogSubSection "4.1 — Removing Known Chrome Services"
        $removalResults = Remove-ChromeServices

        # --- Phase Result ---
        Write-Log ("Phase 4 complete: {0} removed, {1} not found (skipped), {2} failed" -f
            $removalResults.Removed, $removalResults.NotFound, $removalResults.Failed) "INFO"

        if ($removalResults.Failed -eq 0) {
            $script:phaseResults["Phase 4"] = "PASS"
        }
        elseif ($removalResults.Removed -gt 0) {
            $script:phaseResults["Phase 4"] = "PARTIAL"
        }
        else {
            $script:phaseResults["Phase 4"] = "FAIL"
        }

        return $true
    }
    catch {
        Write-Log "Critical error during service removal: $($_.Exception.Message)" "ERROR"
        $null = $script:errorCollection.Add([PSCustomObject]@{
                Timestamp   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Phase       = "Phase 4"
                Item        = "Service Removal Orchestration"
                Error       = $_.Exception.Message
                ErrorRecord = $_
            })
        $script:phaseResults["Phase 4"] = "FAIL"
        return $false
    }
}

#endregion

#region PHASE 5 — SCHEDULED TASK REMOVAL
# ============================================================================
# PHASE 5 — SCHEDULED TASK REMOVAL
# ============================================================================

function Remove-SingleScheduledTask {
    <#
    .SYNOPSIS
        Removes a single scheduled task by name. Uses Unregister-ScheduledTask
        first, falls back to schtasks.exe if needed.
    .PARAMETER TaskName
        The name of the scheduled task to remove.
    .PARAMETER TaskPath
        The path of the scheduled task (e.g., "\Google\" or "\").
    .OUTPUTS
        [string] Result: "Removed", "NotFound", or "Failed"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TaskName,

        [Parameter(Mandatory = $false)]
        [string]$TaskPath = "\"
    )

    # --- Detection ---
    $task = $null
    try {
        $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    }
    catch {
        # Task does not exist
    }

    if ($null -eq $task) {
        Write-Log "Scheduled task not found (skip): $TaskName" "SKIP"
        return "NotFound"
    }

    $taskFullPath = "$($task.TaskPath)$($task.TaskName)"
    Write-Log "Scheduled task found: $taskFullPath | State: $($task.State)" "INFO"

    # --- If task is running, stop it first ---
    if ($task.State -eq 'Running') {
        $stopResult = Invoke-Action -Description "Stop running task: $taskFullPath" -Action {
            Stop-ScheduledTask -TaskName $TaskName -ErrorAction Stop
            Start-Sleep -Seconds 2
        } -Phase "Phase 5" -Item "Stop: $taskFullPath"

        if (-not $stopResult) {
            Write-Log "Could not stop running task — will attempt removal anyway: $taskFullPath" "WARNING"
        }
    }

    # --- Disable the task before removal (safety measure) ---
    if (-not $DryRun) {
        try {
            Disable-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue | Out-Null
            Write-Log "Task disabled: $taskFullPath" "VERBOSE"
        }
        catch {
            Write-Log "Could not disable task (non-critical): $taskFullPath" "VERBOSE"
        }
    }

    # --- Primary removal: Unregister-ScheduledTask ---
    $removeResult = Invoke-Action -Description "Unregister scheduled task: $taskFullPath" -Action {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction Stop
    } -Phase "Phase 5" -Item $taskFullPath

    if ($removeResult) {
        # Verify removal
        $verifyTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        if ($null -eq $verifyTask) {
            Write-Log "Scheduled task successfully removed and verified: $taskFullPath" "SUCCESS"
            return "Removed"
        }
        else {
            Write-Log "Task still exists after Unregister-ScheduledTask: $taskFullPath" "WARNING"
        }
    }

    # --- Fallback: schtasks.exe ---
    if (-not $DryRun) {
        Write-Log "Attempting schtasks.exe fallback for: $taskFullPath" "WARNING"

        $schtasksResult = Invoke-Action -Description "schtasks.exe delete: $taskFullPath" -Action {
            $schtasksOutput = & schtasks.exe /Delete /TN "$taskFullPath" /F 2>&1
            $schtasksExitCode = $LASTEXITCODE
            if ($schtasksExitCode -ne 0) {
                throw "schtasks.exe delete failed with exit code $schtasksExitCode — $schtasksOutput"
            }
        } -Phase "Phase 4" -Item "schtasks: $taskFullPath"

        if ($schtasksResult) {
            Write-Log "Scheduled task removed via schtasks.exe fallback: $taskFullPath" "SUCCESS"
            # Remove the previous error since we recovered
            $previousErrors = @($script:errorCollection | Where-Object { $_.Item -eq $taskFullPath })
            foreach ($prevErr in $previousErrors) {
                $script:errorCollection.Remove($prevErr)
                if ($script:totalItemsFailed -gt 0) { $script:totalItemsFailed-- }
            }
            return "Removed"
        }
    }
    elseif ($DryRun) {
        return "Removed"
    }

    Write-Log "Failed to remove scheduled task: $taskFullPath" "ERROR"
    return "Failed"
}

function Remove-ChromeScheduledTasks {
    <#
    .SYNOPSIS
        Removes all Chrome and Google Update scheduled tasks.
        Uses pattern matching to discover tasks dynamically.
    .OUTPUTS
        [PSCustomObject] with counts of removed, not found, and failed tasks.
    #>
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        Removed  = 0
        NotFound = 0
        Failed   = 0
    }

    # --- Build a deduplicated list of tasks to remove ---
    $tasksToRemove = [System.Collections.ArrayList]::new()
    $taskNamesProcessed = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    # --- Method 1: Pattern-based discovery ---
    Write-Log "Scanning for Chrome-related scheduled tasks by pattern..." "VERBOSE"

    try {
        $allTasks = Get-ScheduledTask -ErrorAction SilentlyContinue

        if ($allTasks) {
            # Match by task name patterns
            foreach ($pattern in $CHROME_TASK_PATTERNS) {
                $matchingTasks = $allTasks | Where-Object { $_.TaskName -like $pattern }
                foreach ($task in $matchingTasks) {
                    if ($taskNamesProcessed.Add($task.TaskName)) {
                        $null = $tasksToRemove.Add([PSCustomObject]@{
                                TaskName = $task.TaskName
                                TaskPath = $task.TaskPath
                                State    = $task.State
                                Source   = "PatternMatch"
                            })
                    }
                }
            }

            # Match by task path containing Google
            $googlePathTasks = $allTasks | Where-Object { $_.TaskPath -match "\\Google\\" }
            foreach ($task in $googlePathTasks) {
                if ($taskNamesProcessed.Add($task.TaskName)) {
                    $null = $tasksToRemove.Add([PSCustomObject]@{
                            TaskName = $task.TaskName
                            TaskPath = $task.TaskPath
                            State    = $task.State
                            Source   = "PathMatch"
                        })
                }
            }

            # Match by task name containing Google Chrome (catch-all)
            $chromeNameTasks = $allTasks | Where-Object {
                $_.TaskName -match "Google" -and $_.TaskName -match "Chrome"
            }
            foreach ($task in $chromeNameTasks) {
                if ($taskNamesProcessed.Add($task.TaskName)) {
                    $null = $tasksToRemove.Add([PSCustomObject]@{
                            TaskName = $task.TaskName
                            TaskPath = $task.TaskPath
                            State    = $task.State
                            Source   = "NameMatch"
                        })
                }
            }
        }
    }
    catch {
        Write-Log "Error enumerating scheduled tasks: $($_.Exception.Message)" "WARNING"
    }

    # --- Method 2: Direct check for known task names (in case Get-ScheduledTask missed them) ---
    $knownTaskNames = @(
        "GoogleUpdateTaskMachineCore"
        "GoogleUpdateTaskMachineUA"
    )

    foreach ($knownName in $knownTaskNames) {
        if ($taskNamesProcessed.Add($knownName)) {
            # Only add if not already found by pattern scan
            $null = $tasksToRemove.Add([PSCustomObject]@{
                    TaskName = $knownName
                    TaskPath = "\"
                    State    = "Unknown"
                    Source   = "KnownName"
                })
        }
    }

    # --- Method 3: Per-user Google Update tasks (SID-based) ---
    foreach ($profile in $script:userProfiles) {
        $userTaskNames = @(
            "GoogleUpdateTaskUserS-$($profile.SID)Core"
            "GoogleUpdateTaskUserS-$($profile.SID)UA"
        )
        foreach ($userTaskName in $userTaskNames) {
            if ($taskNamesProcessed.Add($userTaskName)) {
                $null = $tasksToRemove.Add([PSCustomObject]@{
                        TaskName = $userTaskName
                        TaskPath = "\"
                        State    = "Unknown"
                        Source   = "PerUserSID"
                    })
            }
        }
    }

    # --- Log what we found ---
    if ($tasksToRemove.Count -eq 0) {
        Write-Log "No Chrome-related scheduled tasks found by any discovery method" "SKIP"
        $result.NotFound++
        return $result
    }

    Write-Log "Discovered $($tasksToRemove.Count) scheduled task(s) to process:" "INFO"
    foreach ($task in $tasksToRemove) {
        Write-Log "  Task: $($task.TaskName) | Path: $($task.TaskPath) | State: $($task.State) | Discovery: $($task.Source)" "VERBOSE"
    }

    # --- Remove each task ---
    foreach ($task in $tasksToRemove) {
        $taskResult = Remove-SingleScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath

        switch ($taskResult) {
            "Removed" { $result.Removed++ }
            "NotFound" { $result.NotFound++ }
            "Failed" { $result.Failed++ }
        }
    }

    # --- Cleanup: Remove Google task folder if it exists and is empty ---
    try {
        $googleTaskFolderPath = Join-Path -Path "$env:SystemRoot\System32\Tasks" -ChildPath "Google"
        if (Test-PathSafe -Path $googleTaskFolderPath) {
            if (Get-FolderIsEmpty -Path $googleTaskFolderPath) {
                Invoke-Action -Description "Remove empty Google task folder: $googleTaskFolderPath" -Action {
                    Remove-Item -Path $googleTaskFolderPath -Force -ErrorAction Stop
                } -Phase "Phase 4" -Item $googleTaskFolderPath | Out-Null
            }
            else {
                Write-Log "Google task folder is not empty — leaving intact: $googleTaskFolderPath" "VERBOSE"
            }
        }
    }
    catch {
        Write-Log "Could not check/remove Google task folder: $($_.Exception.Message)" "VERBOSE"
    }

    return $result
}

function Invoke-Phase5ScheduledTaskRemoval {
    <#
    .SYNOPSIS
        Orchestrates Phase 5: removes all Chrome-related scheduled tasks.
    .OUTPUTS
        [bool] True if phase completed. False on critical failure.
    #>
    [CmdletBinding()]
    param()

    Write-LogSection "PHASE 5: SCHEDULED TASK REMOVAL"

    try {
        # --- Pre-check: did Phase 1 find any tasks? ---
        if ($script:chromeTasks.Count -eq 0) {
            Write-Log "No Chrome scheduled tasks were detected in Phase 1 — performing comprehensive live scan" "VERBOSE"
        }
        else {
            Write-Log "Phase 1 detected $($script:chromeTasks.Count) Chrome-related scheduled task(s)" "INFO"
            foreach ($task in $script:chromeTasks) {
                Write-Log "  Queued: $($task.TaskName) | Path: $($task.TaskPath) | State: $($task.State)" "VERBOSE"
            }
        }

        # --- Execute task removal (includes its own comprehensive discovery) ---
        Write-LogSubSection "5.1 — Discovering and Removing Chrome Scheduled Tasks"
        $removalResults = Remove-ChromeScheduledTasks

        # --- Phase Result ---
        Write-Log ("Phase 5 complete: {0} removed, {1} not found (skipped), {2} failed" -f
            $removalResults.Removed, $removalResults.NotFound, $removalResults.Failed) "INFO"

        if ($removalResults.Failed -eq 0) {
            $script:phaseResults["Phase 5"] = "PASS"
        }
        elseif ($removalResults.Removed -gt 0) {
            $script:phaseResults["Phase 5"] = "PARTIAL"
        }
        else {
            $script:phaseResults["Phase 5"] = "FAIL"
        }

        return $true
    }
    catch {
        Write-Log "Critical error during scheduled task removal: $($_.Exception.Message)" "ERROR"
        $null = $script:errorCollection.Add([PSCustomObject]@{
                Timestamp   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Phase       = "Phase 5"
                Item        = "Scheduled Task Removal Orchestration"
                Error       = $_.Exception.Message
                ErrorRecord = $_
            })
        $script:phaseResults["Phase 5"] = "FAIL"
        return $false
    }
}

#endregion

#region PHASE 3 — GRACEFUL UNINSTALL
# ============================================================================
# PHASE 3 — GRACEFUL UNINSTALL
# ============================================================================

function Invoke-ChromeEXEUninstall {
    <#
    .SYNOPSIS
        Runs Chrome's setup.exe-based uninstaller with force flags.
    .PARAMETER ExePath
        Path to the setup.exe uninstaller binary.
    .PARAMETER Arguments
        Original arguments from the UninstallString.
    .OUTPUTS
        [PSCustomObject] with ExitCode, Success, and Message.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExePath,

        [Parameter(Mandatory = $false)]
        [string]$Arguments = ""
    )

    $result = [PSCustomObject]@{
        ExitCode = -1
        Success  = $false
        Message  = ""
    }

    # Build comprehensive uninstall arguments
    # Ensure we have all necessary flags for silent, forced removal
    $uninstallArgs = $Arguments

    # Add force-uninstall if not already present
    if ($uninstallArgs -notmatch '--force-uninstall') {
        $uninstallArgs = "$uninstallArgs --force-uninstall"
    }

    # Add system-level if not already present (per-machine install)
    if ($uninstallArgs -notmatch '--system-level') {
        $uninstallArgs = "$uninstallArgs --system-level"
    }

    # Ensure uninstall flag is present
    if ($uninstallArgs -notmatch '--uninstall') {
        $uninstallArgs = "--uninstall $uninstallArgs"
    }

    $uninstallArgs = $uninstallArgs.Trim()

    Write-Log "Uninstaller path: $ExePath" "VERBOSE"
    Write-Log "Uninstaller arguments: $uninstallArgs" "VERBOSE"

    $invokeResult = Invoke-Action -Description "Run Chrome EXE uninstaller: `"$ExePath`" $uninstallArgs" -Action {
        $process = Start-Process -FilePath $ExePath -ArgumentList $uninstallArgs -PassThru -WindowStyle Hidden -ErrorAction Stop
        $waitResult = $process.WaitForExit(300000)  # 5-minute timeout
        if (-not $waitResult) {
            Write-Log "Uninstaller timed out after 5 minutes — force-killing" "WARNING"
            $process | Stop-Process -Force -ErrorAction SilentlyContinue
            throw "Uninstaller process timed out after 300 seconds"
        }

        # Set result properties in the outer scope
        $result.ExitCode = $process.ExitCode
    } -Phase "Phase 3" -Item "EXE Uninstaller"

    if ($invokeResult) {
        # Interpret exit codes
        switch ($result.ExitCode) {
            0 {
                $result.Success = $true
                $result.Message = "Uninstaller completed successfully (exit code 0)"
                Write-Log $result.Message "SUCCESS"
            }
            19 {
                # Chrome-specific: uninstall successful but requires reboot
                $result.Success = $true
                $result.Message = "Uninstaller completed — reboot required (exit code 19)"
                Write-Log $result.Message "SUCCESS"
            }
            default {
                $result.Success = $false
                $result.Message = "Uninstaller returned non-zero exit code: $($result.ExitCode)"
                Write-Log $result.Message "WARNING"
                Write-Log "Non-zero exit code is not necessarily a failure — brute-force cleanup will handle remaining artifacts" "VERBOSE"
            }
        }
    }
    else {
        $result.Message = "Failed to execute uninstaller"
        Write-Log $result.Message "ERROR"
    }

    return $result
}

function Invoke-ChromeMSIUninstall {
    <#
    .SYNOPSIS
        Runs MSI-based uninstall using msiexec.exe with the Chrome product code.
    .PARAMETER ProductCode
        The MSI product code GUID (e.g., "{xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}").
    .OUTPUTS
        [PSCustomObject] with ExitCode, Success, and Message.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProductCode
    )

    $result = [PSCustomObject]@{
        ExitCode = -1
        Success  = $false
        Message  = ""
    }

    $msiArgs = "/x $ProductCode /qn /norestart REBOOT=ReallySuppress"

    Write-Log "MSI Product Code: $ProductCode" "VERBOSE"
    Write-Log "MSI Arguments: msiexec.exe $msiArgs" "VERBOSE"

    $invokeResult = Invoke-Action -Description "Run MSI uninstall: msiexec.exe $msiArgs" -Action {
        $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $msiArgs -PassThru -WindowStyle Hidden -ErrorAction Stop
        $waitResult = $process.WaitForExit(300000)  # 5-minute timeout
        if (-not $waitResult) {
            Write-Log "MSI uninstall timed out after 5 minutes — force-killing" "WARNING"
            $process | Stop-Process -Force -ErrorAction SilentlyContinue
            throw "MSI uninstall process timed out after 300 seconds"
        }

        $result.ExitCode = $process.ExitCode
    } -Phase "Phase 3" -Item "MSI Uninstaller"

    if ($invokeResult) {
        # Interpret MSI exit codes
        switch ($result.ExitCode) {
            0 {
                $result.Success = $true
                $result.Message = "MSI uninstall completed successfully (exit code 0)"
                Write-Log $result.Message "SUCCESS"
            }
            1605 {
                $result.Success = $true
                $result.Message = "MSI product not installed / already removed (exit code 1605)"
                Write-Log $result.Message "INFO"
            }
            1614 {
                $result.Success = $true
                $result.Message = "MSI product unregistered / not found (exit code 1614)"
                Write-Log $result.Message "INFO"
            }
            3010 {
                $result.Success = $true
                $result.Message = "MSI uninstall completed — reboot required (exit code 3010)"
                Write-Log $result.Message "SUCCESS"
            }
            1618 {
                $result.Success = $false
                $result.Message = "Another MSI installation is in progress (exit code 1618) — retry may be needed"
                Write-Log $result.Message "WARNING"
            }
            1603 {
                $result.Success = $false
                $result.Message = "MSI uninstall fatal error (exit code 1603)"
                Write-Log $result.Message "WARNING"
                Write-Log "Fatal MSI error — brute-force cleanup will handle remaining artifacts" "VERBOSE"
            }
            default {
                $result.Success = $false
                $result.Message = "MSI uninstall returned exit code: $($result.ExitCode)"
                Write-Log $result.Message "WARNING"
                Write-Log "Brute-force cleanup will handle remaining artifacts" "VERBOSE"
            }
        }
    }
    else {
        $result.Message = "Failed to execute MSI uninstall"
        Write-Log $result.Message "ERROR"
    }

    return $result
}

function Invoke-WMIProductUninstall {
    <#
    .SYNOPSIS
        Attempts to uninstall Chrome via WMI Win32_Product (fallback method).
        This is intentionally used only as a last resort because Win32_Product
        queries can trigger MSI reconfiguration of other products.
    .OUTPUTS
        [PSCustomObject] with ExitCode, Success, and Message.
    #>
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        ExitCode = -1
        Success  = $false
        Message  = ""
    }

    Write-Log "Attempting WMI-based product detection (Win32_Product)..." "VERBOSE"
    Write-Log "NOTE: Win32_Product queries can be slow and may trigger MSI reconfiguration" "VERBOSE"

    try {
        # Use CIM for better performance where possible
        $chromeProduct = Get-CimInstance -ClassName Win32_Product -Filter "Name LIKE '%Google Chrome%'" -ErrorAction SilentlyContinue

        if ($null -eq $chromeProduct) {
            $result.Message = "Chrome not found in Win32_Product — may not have been MSI-installed"
            Write-Log $result.Message "VERBOSE"
            return $result
        }

        Write-Log "Found Chrome in Win32_Product: $($chromeProduct.Name) v$($chromeProduct.Version)" "INFO"
        Write-Log "WMI IdentifyingNumber: $($chromeProduct.IdentifyingNumber)" "VERBOSE"

        $invokeResult = Invoke-Action -Description "WMI uninstall: $($chromeProduct.Name)" -Action {
            $uninstallResult = $chromeProduct | Invoke-CimMethod -MethodName Uninstall -ErrorAction Stop
            $result.ExitCode = $uninstallResult.ReturnValue
        } -Phase "Phase 3" -Item "WMI Uninstall"

        if ($invokeResult) {
            if ($result.ExitCode -eq 0) {
                $result.Success = $true
                $result.Message = "WMI uninstall completed successfully (return value 0)"
                Write-Log $result.Message "SUCCESS"
            }
            else {
                $result.Message = "WMI uninstall returned non-zero: $($result.ExitCode)"
                Write-Log $result.Message "WARNING"
            }
        }
    }
    catch {
        $result.Message = "WMI product query/uninstall failed: $($_.Exception.Message)"
        Write-Log $result.Message "WARNING"
    }

    return $result
}

function Invoke-Phase3GracefulUninstall {
    <#
    .SYNOPSIS
        Orchestrates Phase 3: attempts graceful uninstall through multiple methods
        in priority order: EXE uninstaller → MSI uninstaller → WMI uninstall.
        Failure is non-blocking — brute-force cleanup continues regardless.
    .OUTPUTS
        [bool] True if phase completed (even if uninstall failed). False on critical failure.
    #>
    [CmdletBinding()]
    param()

    Write-LogSection "PHASE 3: GRACEFUL UNINSTALL"

    $uninstallAttempted = $false
    $uninstallSucceeded = $false

    try {
        # --- 3.1 Per-Machine EXE Uninstaller ---
        Write-LogSubSection "3.1 — Per-Machine EXE Uninstaller"

        if ($null -ne $script:perMachineUninstaller -and $null -ne $script:perMachineUninstaller.PerMachine) {
            $pmUninstaller = $script:perMachineUninstaller.PerMachine

            if (-not $pmUninstaller.IsMSI) {
                # This is a setup.exe-based uninstaller
                if ($pmUninstaller.BinaryExists) {
                    Write-Log "Per-machine EXE uninstaller available and binary exists" "INFO"
                    $uninstallAttempted = $true

                    $exeResult = Invoke-ChromeEXEUninstall -ExePath $pmUninstaller.ExePath -Arguments $pmUninstaller.Arguments

                    if ($exeResult.Success) {
                        $uninstallSucceeded = $true
                        Write-Log "Graceful EXE uninstall succeeded" "SUCCESS"

                        # Wait for uninstaller cleanup to finish
                        if (-not $DryRun) {
                            Write-Log "Waiting 5 seconds for uninstaller cleanup to complete..." "VERBOSE"
                            Start-Sleep -Seconds 5
                        }
                    }
                    else {
                        Write-Log "Graceful EXE uninstall did not fully succeed — continuing to brute-force phases" "WARNING"
                    }
                }
                else {
                    Write-Log "Per-machine UninstallString references non-existent binary: $($pmUninstaller.ExePath)" "WARNING"
                    Write-Log "Skipping graceful EXE uninstall — brute-force cleanup in subsequent phases will handle removal" "WARNING"
                }
            }
            else {
                Write-Log "Per-machine uninstaller is MSI-based — will be handled in step 3.3" "VERBOSE"
            }
        }
        else {
            Write-Log "No per-machine EXE uninstaller registered" "SKIP"
        }

        # --- 3.2 Per-User Uninstaller (Skip — SYSTEM context) ---
        Write-LogSubSection "3.2 — Per-User Uninstaller"

        if ($null -ne $script:perMachineUninstaller -and $script:perMachineUninstaller.PerUser.Count -gt 0) {
            Write-Log "Per-user Chrome installation(s) detected: $($script:perMachineUninstaller.PerUser.Count)" "INFO"
            foreach ($perUserInstall in $script:perMachineUninstaller.PerUser) {
                Write-Log "  User: $($perUserInstall.UserName) (SID: $($perUserInstall.SID))" "VERBOSE"
                Write-Log "  UninstallString: $($perUserInstall.UninstallString)" "VERBOSE"
            }
            Write-Log "SKIPPING per-user uninstaller(s) — running under SYSTEM context makes per-user uninstallers unreliable" "WARNING"
            Write-Log "Per-user installations will be cleaned by brute-force file and registry removal in Phases 6-7" "WARNING"
        }
        else {
            Write-Log "No per-user installations detected" "SKIP"
        }

        # --- 3.3 MSI-Based Uninstall ---
        Write-LogSubSection "3.3 — MSI-Based Uninstall"

        if (-not $uninstallSucceeded) {
            # Only attempt MSI if EXE uninstall didn't already succeed

            $msiProductCode = $null

            # Check if we have a product code from detection phase
            if ($null -ne $script:perMachineUninstaller -and $script:perMachineUninstaller.MSIProductCode) {
                $msiProductCode = $script:perMachineUninstaller.MSIProductCode
                Write-Log "MSI product code available from detection: $msiProductCode" "INFO"
            }

            # Also check if the UninstallString itself is MSI-based
            if ($null -eq $msiProductCode -and
                $null -ne $script:perMachineUninstaller -and
                $null -ne $script:perMachineUninstaller.PerMachine -and
                $script:perMachineUninstaller.PerMachine.IsMSI) {

                # Extract product code from UninstallString
                $uninstallStr = $script:perMachineUninstaller.PerMachine.UninstallString
                if ($uninstallStr -match '\{([A-Fa-f0-9\-]+)\}') {
                    $msiProductCode = "{$($Matches[1])}"
                    Write-Log "MSI product code extracted from UninstallString: $msiProductCode" "INFO"
                }
            }

            if ($msiProductCode) {
                $uninstallAttempted = $true
                $msiResult = Invoke-ChromeMSIUninstall -ProductCode $msiProductCode

                if ($msiResult.Success) {
                    $uninstallSucceeded = $true
                    Write-Log "Graceful MSI uninstall succeeded" "SUCCESS"

                    # Wait for MSI cleanup
                    if (-not $DryRun) {
                        Write-Log "Waiting 5 seconds for MSI cleanup to complete..." "VERBOSE"
                        Start-Sleep -Seconds 5
                    }
                }
                else {
                    Write-Log "MSI uninstall did not fully succeed — continuing to brute-force phases" "WARNING"
                }
            }
            else {
                Write-Log "No MSI product code found — MSI uninstall not applicable" "SKIP"
            }
        }
        else {
            Write-Log "EXE uninstall already succeeded — skipping MSI attempt" "VERBOSE"
        }

        # --- 5.4 WMI Fallback (only if nothing else worked) ---
        Write-LogSubSection "5.4 — WMI Uninstall Fallback"

        if (-not $uninstallSucceeded -and -not $uninstallAttempted) {
            Write-Log "No standard uninstaller succeeded or was available — attempting WMI fallback" "INFO"
            $wmiResult = Invoke-WMIProductUninstall

            if ($wmiResult.Success) {
                $uninstallSucceeded = $true
                $uninstallAttempted = $true
                Write-Log "WMI uninstall fallback succeeded" "SUCCESS"

                if (-not $DryRun) {
                    Write-Log "Waiting 5 seconds for WMI uninstall cleanup..." "VERBOSE"
                    Start-Sleep -Seconds 5
                }
            }
            elseif ($wmiResult.ExitCode -eq -1) {
                # Product not found in WMI — not an error, just not applicable
                Write-Log "Chrome not found via WMI — WMI uninstall not applicable" "SKIP"
            }
            else {
                Write-Log "WMI uninstall fallback did not succeed" "WARNING"
                $uninstallAttempted = $true
            }
        }
        elseif ($uninstallSucceeded) {
            Write-Log "Graceful uninstall already succeeded — skipping WMI fallback" "VERBOSE"
        }
        else {
            Write-Log "Standard uninstall was attempted but did not succeed — skipping WMI to avoid slowdown" "VERBOSE"
            Write-Log "Brute-force cleanup in Phases 6-9 will handle complete removal" "INFO"
        }

        # --- Phase Summary ---
        Write-LogSubSection "Phase 3 Summary"

        if ($uninstallSucceeded) {
            Write-Log "Graceful uninstall: SUCCEEDED" "SUCCESS"
            Write-Log "Proceeding to brute-force cleanup phases to remove any remaining artifacts" "INFO"
            $script:phaseResults["Phase 3"] = "PASS"
        }
        elseif ($uninstallAttempted) {
            Write-Log "Graceful uninstall: ATTEMPTED but did not fully succeed" "WARNING"
            Write-Log "Brute-force cleanup in Phases 6-9 will handle complete removal" "INFO"
            $script:phaseResults["Phase 3"] = "PARTIAL"
        }
        else {
            if ($script:installationType -eq "None") {
                Write-Log "Graceful uninstall: NOT APPLICABLE (no installation detected)" "INFO"
                Write-Log "Proceeding with artifact cleanup only" "INFO"
                $script:phaseResults["Phase 3"] = "PASS"
            }
            else {
                Write-Log "Graceful uninstall: SKIPPED (no usable uninstaller found)" "WARNING"
                Write-Log "Brute-force cleanup in Phases 6-9 will handle removal" "INFO"
                $script:phaseResults["Phase 3"] = "PARTIAL"
            }
        }

        Write-Log "Proceeding to brute-force cleanup regardless of graceful uninstall outcome" "INFO"
        return $true
    }
    catch {
        Write-Log "Critical error during graceful uninstall: $($_.Exception.Message)" "ERROR"
        $null = $script:errorCollection.Add([PSCustomObject]@{
                Timestamp   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Phase       = "Phase 3"
                Item        = "Graceful Uninstall Orchestration"
                Error       = $_.Exception.Message
                ErrorRecord = $_
            })
        Write-Log "Graceful uninstall phase failed — brute-force cleanup will continue in subsequent phases" "WARNING"
        $script:phaseResults["Phase 3"] = "FAIL"
        return $true  # Return true because this is non-blocking — we continue anyway
    }
}

#endregion

#region PHASE 6 — FILE SYSTEM CLEANUP
# ============================================================================
# PHASE 6 — FILE SYSTEM CLEANUP
# ============================================================================

function Remove-ChromeMachineFiles {
    <#
    .SYNOPSIS
        Removes Chrome-related directories from machine-level paths
        (Program Files, ProgramData, etc.).
    .OUTPUTS
        [PSCustomObject] with counts of removed, skipped, and failed paths.
    #>
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        Removed = 0
        Skipped = 0
        Failed  = 0
    }

    Write-Log "Processing $($CHROME_MACHINE_PATHS.Count) machine-level directory targets..." "VERBOSE"

    foreach ($targetPath in $CHROME_MACHINE_PATHS) {
        # Resolve environment variables that may not have expanded
        $resolvedPath = [System.Environment]::ExpandEnvironmentVariables($targetPath)

        if (-not (Test-PathSafe -Path $resolvedPath)) {
            Write-Log "Machine path not found (skip): $resolvedPath" "SKIP"
            $result.Skipped++
            continue
        }

        $removeSuccess = Remove-DirectoryIfExists -Path $resolvedPath -Phase "Phase 6"

        if ($removeSuccess) {
            # Verify it's actually gone (Remove-DirectoryIfExists returns true for "not found" too)
            if (-not (Test-PathSafe -Path $resolvedPath)) {
                $result.Removed++
            }
            else {
                $result.Skipped++
            }
        }
        else {
            $result.Failed++
        }
    }

    return $result
}

function Remove-ChromeUserFiles {
    <#
    .SYNOPSIS
        Removes Chrome-related directories from ALL user profiles.
        Iterates each profile and removes Chrome, Update, CrashReports,
        and Software Reporter Tool folders.
    .OUTPUTS
        [PSCustomObject] with counts of removed, skipped, and failed paths.
    #>
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        Removed = 0
        Skipped = 0
        Failed  = 0
    }

    if ($ExcludeUserData) {
        Write-Log "ExcludeUserData flag is set — skipping all per-user file cleanup" "SKIP"
        return $result
    }

    if ($script:userProfiles.Count -eq 0) {
        Write-Log "No user profiles to process" "SKIP"
        return $result
    }

    Write-Log "Processing $($script:userProfiles.Count) user profile(s)..." "INFO"

    foreach ($profile in $script:userProfiles) {
        if (-not $profile.ProfileExists) {
            Write-Log "Profile directory does not exist (skip): $($profile.ProfilePath) ($($profile.UserName))" "SKIP"
            continue
        }

        Write-Log "Processing user profile: $($profile.UserName) ($($profile.SID))" "VERBOSE"

        $userPathsProcessed = 0
        $userPathsRemoved = 0

        foreach ($relativePath in $CHROME_USER_RELATIVE_PATHS) {
            $fullPath = Join-Path -Path $profile.ProfilePath -ChildPath $relativePath
            $userPathsProcessed++

            if (-not (Test-PathSafe -Path $fullPath)) {
                Write-Log "  User path not found (skip): $fullPath" "SKIP"
                $result.Skipped++
                continue
            }

            $removeSuccess = Remove-DirectoryIfExists -Path $fullPath -Phase "Phase 6"

            if ($removeSuccess -and -not (Test-PathSafe -Path $fullPath)) {
                $result.Removed++
                $userPathsRemoved++
            }
            elseif ($removeSuccess) {
                $result.Skipped++
            }
            else {
                $result.Failed++
            }
        }

        Write-Log "  User $($profile.UserName): $userPathsRemoved of $userPathsProcessed paths removed" "VERBOSE"
    }

    return $result
}

function Remove-ChromeTempFiles {
    <#
    .SYNOPSIS
        Removes Chrome-related temporary files from system and user temp directories.
    .OUTPUTS
        [PSCustomObject] with counts of removed, skipped, and failed items.
    #>
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        Removed = 0
        Skipped = 0
        Failed  = 0
    }

    # --- Build list of temp directories to scan ---
    $tempDirectories = [System.Collections.ArrayList]::new()

    # System temp
    $systemTemp = [System.Environment]::GetEnvironmentVariable("TEMP", "Machine")
    if (-not $systemTemp) { $systemTemp = "C:\Windows\Temp" }
    $null = $tempDirectories.Add($systemTemp)

    # Windows Temp (explicit)
    $windowsTemp = Join-Path -Path $env:SystemRoot -ChildPath "Temp"
    if ($windowsTemp -ne $systemTemp) {
        $null = $tempDirectories.Add($windowsTemp)
    }

    # Per-user temp directories
    foreach ($profile in $script:userProfiles) {
        if (-not $profile.ProfileExists) { continue }

        $userTempPaths = @(
            Join-Path -Path $profile.ProfilePath -ChildPath "AppData\Local\Temp"
        )

        foreach ($userTempPath in $userTempPaths) {
            if (Test-PathSafe -Path $userTempPath) {
                $null = $tempDirectories.Add($userTempPath)
            }
        }
    }

    Write-Log "Scanning $($tempDirectories.Count) temp director(ies) for Chrome artifacts..." "VERBOSE"

    # --- Scan each temp directory for matching patterns ---
    foreach ($tempDir in $tempDirectories) {
        if (-not (Test-PathSafe -Path $tempDir)) {
            Write-Log "Temp directory not found (skip): $tempDir" "SKIP"
            continue
        }

        Write-Log "Scanning temp directory: $tempDir" "VERBOSE"

        foreach ($pattern in $CHROME_TEMP_PATTERNS) {
            try {
                $matchingItems = Get-ChildItem -Path $tempDir -Filter $pattern -Force -ErrorAction SilentlyContinue

                if (-not $matchingItems) { continue }

                $matchCount = @($matchingItems).Count
                Write-Log "  Found $matchCount item(s) matching pattern '$pattern' in: $tempDir" "INFO"

                foreach ($item in $matchingItems) {
                    if ($item.PSIsContainer) {
                        # It's a directory
                        $removeSuccess = Remove-DirectoryIfExists -Path $item.FullName -Phase "Phase 6"
                    }
                    else {
                        # It's a file
                        $removeSuccess = Remove-FileIfExists -Path $item.FullName -Phase "Phase 6"
                    }

                    if ($removeSuccess -and -not (Test-PathSafe -Path $item.FullName)) {
                        $result.Removed++
                    }
                    elseif ($removeSuccess) {
                        $result.Skipped++
                    }
                    else {
                        $result.Failed++
                    }
                }
            }
            catch {
                Write-Log "  Error scanning $tempDir for pattern '$pattern': $($_.Exception.Message)" "WARNING"
            }
        }
    }

    return $result
}

function Remove-ChromeInstallerCache {
    <#
    .SYNOPSIS
        Removes Chrome-related cached MSI/MSP files from C:\Windows\Installer.
        Identifies Chrome MSI files by reading MSI database summary information.
    .OUTPUTS
        [PSCustomObject] with counts of removed, skipped, and failed items.
    #>
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        Removed = 0
        Skipped = 0
        Failed  = 0
    }

    $installerPath = Join-Path -Path $env:SystemRoot -ChildPath "Installer"

    if (-not (Test-PathSafe -Path $installerPath)) {
        Write-Log "Windows Installer cache not found: $installerPath" "SKIP"
        return $result
    }

    Write-Log "Scanning Windows Installer cache for Chrome MSI files..." "VERBOSE"

    # --- Method 1: Use known product code if available ---
    if ($null -ne $script:perMachineUninstaller -and $script:perMachineUninstaller.MSIProductCode) {
        $productCode = $script:perMachineUninstaller.MSIProductCode
        Write-Log "Searching for cached MSI with product code: $productCode" "VERBOSE"

        # Check the installer registration for the cached MSI path
        $installerRegPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products"
            "HKLM:\SOFTWARE\Classes\Installer\Products"
        )

        foreach ($regBasePath in $installerRegPaths) {
            if (-not (Test-PathSafe -Path $regBasePath)) { continue }

            try {
                $productSubKeys = Get-ChildItem -Path $regBasePath -ErrorAction SilentlyContinue
                foreach ($subKey in $productSubKeys) {
                    try {
                        $productName = (Get-ItemProperty -Path $subKey.PSPath -Name "ProductName" -ErrorAction SilentlyContinue).ProductName
                        if ($productName -and $productName -match "Google Chrome") {
                            Write-Log "Found Chrome installer registry entry: $($subKey.PSPath)" "VERBOSE"
                            Write-Log "  ProductName: $productName" "VERBOSE"

                            # Look for LocalPackage value
                            $installProperties = Get-ChildItem -Path $subKey.PSPath -ErrorAction SilentlyContinue |
                            Where-Object { $_.PSChildName -eq "InstallProperties" }

                            if ($installProperties) {
                                $localPackage = (Get-ItemProperty -Path $installProperties.PSPath -Name "LocalPackage" -ErrorAction SilentlyContinue).LocalPackage
                                if ($localPackage -and (Test-PathSafe -Path $localPackage)) {
                                    Write-Log "  Cached MSI found: $localPackage" "INFO"
                                    $removeSuccess = Remove-FileIfExists -Path $localPackage -Phase "Phase 6"
                                    if ($removeSuccess -and -not (Test-PathSafe -Path $localPackage)) {
                                        $result.Removed++
                                    }
                                    else {
                                        $result.Failed++
                                    }
                                }
                            }
                        }
                    }
                    catch {
                        continue
                    }
                }
            }
            catch {
                Write-Log "Error scanning installer registry at $regBasePath : $($_.Exception.Message)" "WARNING"
            }
        }
    }

    # --- Method 2: Scan MSI files by reading their summary info ---
    try {
        $msiFiles = Get-ChildItem -Path $installerPath -Filter "*.msi" -Force -ErrorAction SilentlyContinue

        if ($msiFiles) {
            Write-Log "Found $(@($msiFiles).Count) MSI file(s) in Installer cache — scanning for Chrome..." "VERBOSE"

            $windowsInstaller = $null
            try {
                $windowsInstaller = New-Object -ComObject WindowsInstaller.Installer -ErrorAction Stop
            }
            catch {
                Write-Log "Could not create WindowsInstaller.Installer COM object — skipping MSI content scan" "WARNING"
            }

            if ($null -ne $windowsInstaller) {
                foreach ($msiFile in $msiFiles) {
                    try {
                        $database = $windowsInstaller.GetType().InvokeMember(
                            "OpenDatabase",
                            [System.Reflection.BindingFlags]::InvokeMethod,
                            $null,
                            $windowsInstaller,
                            @($msiFile.FullName, 0)
                        )

                        $view = $database.GetType().InvokeMember(
                            "OpenView",
                            [System.Reflection.BindingFlags]::InvokeMethod,
                            $null,
                            $database,
                            @("SELECT `Value` FROM `Property` WHERE `Property` = 'ProductName'")
                        )

                        $view.GetType().InvokeMember("Execute", [System.Reflection.BindingFlags]::InvokeMethod, $null, $view, $null)
                        $record = $view.GetType().InvokeMember("Fetch", [System.Reflection.BindingFlags]::InvokeMethod, $null, $view, $null)

                        if ($null -ne $record) {
                            $msiProductName = $record.GetType().InvokeMember(
                                "StringData",
                                [System.Reflection.BindingFlags]::GetProperty,
                                $null,
                                $record,
                                @(1)
                            )

                            if ($msiProductName -match "Google Chrome") {
                                Write-Log "Chrome MSI found in cache: $($msiFile.FullName) | Product: $msiProductName" "INFO"
                                $removeSuccess = Remove-FileIfExists -Path $msiFile.FullName -Phase "Phase 6"
                                if ($removeSuccess -and -not (Test-PathSafe -Path $msiFile.FullName)) {
                                    $result.Removed++
                                }
                                else {
                                    $result.Failed++
                                }
                            }

                            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($record) | Out-Null
                        }

                        $view.GetType().InvokeMember("Close", [System.Reflection.BindingFlags]::InvokeMethod, $null, $view, $null)
                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($view) | Out-Null
                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($database) | Out-Null
                    }
                    catch {
                        # Could not read this MSI — skip it (might be corrupt or locked)
                        continue
                    }
                }

                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($windowsInstaller) | Out-Null
            }
        }
        else {
            Write-Log "No MSI files found in Installer cache" "SKIP"
        }
    }
    catch {
        Write-Log "Error scanning Windows Installer cache: $($_.Exception.Message)" "WARNING"
    }

    # --- Also check for MSP (patch) files ---
    try {
        $mspFiles = Get-ChildItem -Path $installerPath -Filter "*.msp" -Force -ErrorAction SilentlyContinue

        if ($mspFiles) {
            Write-Log "Found $(@($mspFiles).Count) MSP file(s) — checking for Chrome patches is not reliable; skipping MSP scan" "VERBOSE"
            # Note: MSP files are harder to identify without extracting patch metadata.
            # Removing wrong MSP files could break other products.
            # The MSI removal above is sufficient.
        }
    }
    catch {
        # Non-critical
    }

    return $result
}

function Remove-ChromePrefetchFiles {
    <#
    .SYNOPSIS
        Removes Chrome-related Windows Prefetch files.
    .OUTPUTS
        [PSCustomObject] with counts of removed, skipped, and failed items.
    #>
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        Removed = 0
        Skipped = 0
        Failed  = 0
    }

    $prefetchPath = Join-Path -Path $env:SystemRoot -ChildPath "Prefetch"

    if (-not (Test-PathSafe -Path $prefetchPath)) {
        Write-Log "Prefetch directory not found: $prefetchPath" "SKIP"
        return $result
    }

    Write-Log "Scanning Prefetch directory for Chrome-related files..." "VERBOSE"

    foreach ($pattern in $CHROME_PREFETCH_PATTERNS) {
        try {
            $matchingFiles = Get-ChildItem -Path $prefetchPath -Filter $pattern -Force -ErrorAction SilentlyContinue

            if (-not $matchingFiles) {
                Write-Log "  No prefetch files matching: $pattern" "SKIP"
                $result.Skipped++
                continue
            }

            $matchCount = @($matchingFiles).Count
            Write-Log "  Found $matchCount prefetch file(s) matching: $pattern" "INFO"

            foreach ($file in $matchingFiles) {
                $removeSuccess = Remove-FileIfExists -Path $file.FullName -Phase "Phase 6"
                if ($removeSuccess -and -not (Test-PathSafe -Path $file.FullName)) {
                    $result.Removed++
                }
                elseif ($removeSuccess) {
                    $result.Skipped++
                }
                else {
                    $result.Failed++
                }
            }
        }
        catch {
            Write-Log "  Error scanning prefetch for pattern '$pattern': $($_.Exception.Message)" "WARNING"
        }
    }

    return $result
}

function Remove-ChromeShortcuts {
    <#
    .SYNOPSIS
        Removes all Chrome shortcuts from Desktop, Start Menu, Taskbar, Quick Launch,
        and Startup folders for all users and public/shared locations.
    .OUTPUTS
        [PSCustomObject] with counts of removed, skipped, and failed items.
    #>
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        Removed = 0
        Skipped = 0
        Failed  = 0
    }

    # --- Build list of all shortcut directories to scan ---
    $shortcutLocations = [System.Collections.ArrayList]::new()

    # --- Public / All Users locations ---
    $publicDesktop = [System.Environment]::GetFolderPath("CommonDesktopDirectory")
    if (-not $publicDesktop) { $publicDesktop = "C:\Users\Public\Desktop" }

    $publicStartMenu = [System.Environment]::GetFolderPath("CommonStartMenu")
    if (-not $publicStartMenu) {
        $publicStartMenu = "$env:ProgramData\Microsoft\Windows\Start Menu"
    }

    $publicStartMenuPrograms = [System.Environment]::GetFolderPath("CommonPrograms")
    if (-not $publicStartMenuPrograms) {
        $publicStartMenuPrograms = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs"
    }

    $publicStartup = [System.Environment]::GetFolderPath("CommonStartup")
    if (-not $publicStartup) {
        $publicStartup = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Startup"
    }

    $null = $shortcutLocations.Add(@{ Path = $publicDesktop; Description = "Public Desktop" })
    $null = $shortcutLocations.Add(@{ Path = $publicStartMenu; Description = "Public Start Menu" })
    $null = $shortcutLocations.Add(@{ Path = $publicStartMenuPrograms; Description = "Public Start Menu Programs" })
    $null = $shortcutLocations.Add(@{ Path = $publicStartup; Description = "Public Startup" })

    # --- Per-User locations ---
    foreach ($profile in $script:userProfiles) {
        if (-not $profile.ProfileExists) { continue }

        $userDesktop = Join-Path -Path $profile.ProfilePath -ChildPath "Desktop"
        $userStartMenu = Join-Path -Path $profile.ProfilePath -ChildPath "AppData\Roaming\Microsoft\Windows\Start Menu"
        $userStartMenuPrograms = Join-Path -Path $profile.ProfilePath -ChildPath "AppData\Roaming\Microsoft\Windows\Start Menu\Programs"
        $userStartup = Join-Path -Path $profile.ProfilePath -ChildPath "AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
        $userQuickLaunch = Join-Path -Path $profile.ProfilePath -ChildPath "AppData\Roaming\Microsoft\Internet Explorer\Quick Launch"
        $userTaskbar = Join-Path -Path $profile.ProfilePath -ChildPath "AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"

        $null = $shortcutLocations.Add(@{ Path = $userDesktop; Description = "$($profile.UserName) Desktop" })
        $null = $shortcutLocations.Add(@{ Path = $userStartMenu; Description = "$($profile.UserName) Start Menu" })
        $null = $shortcutLocations.Add(@{ Path = $userStartMenuPrograms; Description = "$($profile.UserName) Start Menu Programs" })
        $null = $shortcutLocations.Add(@{ Path = $userStartup; Description = "$($profile.UserName) Startup" })
        $null = $shortcutLocations.Add(@{ Path = $userQuickLaunch; Description = "$($profile.UserName) Quick Launch" })
        $null = $shortcutLocations.Add(@{ Path = $userTaskbar; Description = "$($profile.UserName) Taskbar" })
    }

    Write-Log "Scanning $($shortcutLocations.Count) shortcut location(s)..." "VERBOSE"

    # --- Scan each location ---
    foreach ($location in $shortcutLocations) {
        $locationPath = $location.Path
        $locationDesc = $location.Description

        if (-not (Test-PathSafe -Path $locationPath)) {
            Write-Log "  Shortcut location not found (skip): $locationDesc ($locationPath)" "SKIP"
            continue
        }

        # --- Method 1: Check for known shortcut names ---
        foreach ($shortcutName in $CHROME_SHORTCUT_NAMES) {
            $shortcutFullPath = Join-Path -Path $locationPath -ChildPath $shortcutName

            if (Test-PathSafe -Path $shortcutFullPath) {
                Write-Log "  Chrome shortcut found: $shortcutFullPath ($locationDesc)" "INFO"
                $removeSuccess = Remove-FileIfExists -Path $shortcutFullPath -Phase "Phase 6"
                if ($removeSuccess -and -not (Test-PathSafe -Path $shortcutFullPath)) {
                    $result.Removed++
                }
                else {
                    $result.Failed++
                }
            }
        }

        # --- Method 2: Scan all .lnk files and resolve targets ---
        try {
            $lnkFiles = Get-ChildItem -Path $locationPath -Filter "*.lnk" -Force -ErrorAction SilentlyContinue

            if ($lnkFiles) {
                foreach ($lnkFile in $lnkFiles) {
                    # Skip if this was already handled by known name check
                    if ($lnkFile.Name -in $CHROME_SHORTCUT_NAMES) { continue }

                    $target = Resolve-ShortcutTarget -ShortcutPath $lnkFile.FullName
                    if ($target -and ($target -match "chrome\.exe" -or $target -match "Google\\Chrome")) {
                        Write-Log "  Chrome shortcut found by target resolution: $($lnkFile.FullName) → $target ($locationDesc)" "INFO"
                        $removeSuccess = Remove-FileIfExists -Path $lnkFile.FullName -Phase "Phase 6"
                        if ($removeSuccess -and -not (Test-PathSafe -Path $lnkFile.FullName)) {
                            $result.Removed++
                        }
                        else {
                            $result.Failed++
                        }
                    }
                }
            }
        }
        catch {
            Write-Log "  Error scanning .lnk files in $locationPath : $($_.Exception.Message)" "WARNING"
        }
    }

    # --- Also check for Chrome folders in Start Menu Programs ---
    $startMenuChromeFolders = @(
        "$publicStartMenuPrograms\Google Chrome"
        "$publicStartMenuPrograms\Google"
    )

    foreach ($profile in $script:userProfiles) {
        if (-not $profile.ProfileExists) { continue }
        $startMenuChromeFolders += Join-Path -Path $profile.ProfilePath -ChildPath "AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Google Chrome"
    }

    foreach ($chromeFolder in $startMenuChromeFolders) {
        if (Test-PathSafe -Path $chromeFolder) {
            Write-Log "  Chrome Start Menu folder found: $chromeFolder" "INFO"
            $removeSuccess = Remove-DirectoryIfExists -Path $chromeFolder -Phase "Phase 6"
            if ($removeSuccess -and -not (Test-PathSafe -Path $chromeFolder)) {
                $result.Removed++
            }
            elseif (-not $removeSuccess) {
                $result.Failed++
            }
        }
    }

    return $result
}

function Remove-EmptyGoogleFolders {
    <#
    .SYNOPSIS
        Removes empty parent "Google" folders after Chrome subfolders have been removed.
        Checks machine-level and per-user locations. Only removes if truly empty.
    .OUTPUTS
        [PSCustomObject] with counts of removed and skipped folders.
    #>
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        Removed = 0
        Skipped = 0
    }

    # --- Machine-level parent folders ---
    Write-Log "Checking machine-level Google parent folders for cleanup..." "VERBOSE"

    foreach ($parentFolder in $GOOGLE_PARENT_FOLDERS) {
        $resolvedPath = [System.Environment]::ExpandEnvironmentVariables($parentFolder)

        if (-not (Test-PathSafe -Path $resolvedPath)) {
            Write-Log "  Parent folder not found (skip): $resolvedPath" "SKIP"
            $result.Skipped++
            continue
        }

        if (Get-FolderIsEmpty -Path $resolvedPath) {
            Write-Log "  Parent Google folder is empty — removing: $resolvedPath" "INFO"
            $removeSuccess = Invoke-Action -Description "Remove empty Google folder: $resolvedPath" -Action {
                Remove-Item -Path $resolvedPath -Force -ErrorAction Stop
            } -Phase "Phase 6" -Item $resolvedPath

            if ($removeSuccess) {
                $result.Removed++
            }
        }
        else {
            Write-Log "  Parent Google folder is NOT empty (other Google products may exist) — leaving intact: $resolvedPath" "INFO"
            $result.Skipped++

            # Log what remains
            try {
                $remainingItems = Get-ChildItem -Path $resolvedPath -Force -ErrorAction SilentlyContinue
                foreach ($item in $remainingItems) {
                    Write-Log "    Remaining: $($item.Name) ($($item.GetType().Name))" "VERBOSE"
                }
            }
            catch {
                # Non-critical
            }
        }
    }

    # --- Per-User parent folders ---
    Write-Log "Checking per-user Google parent folders for cleanup..." "VERBOSE"

    foreach ($profile in $script:userProfiles) {
        if (-not $profile.ProfileExists) { continue }

        $userGoogleFolder = Join-Path -Path $profile.ProfilePath -ChildPath "AppData\Local\Google"

        if (-not (Test-PathSafe -Path $userGoogleFolder)) {
            continue
        }

        if (Get-FolderIsEmpty -Path $userGoogleFolder) {
            Write-Log "  User Google folder is empty — removing: $userGoogleFolder ($($profile.UserName))" "INFO"
            $removeSuccess = Invoke-Action -Description "Remove empty user Google folder: $userGoogleFolder" -Action {
                Remove-Item -Path $userGoogleFolder -Force -ErrorAction Stop
            } -Phase "Phase 6" -Item $userGoogleFolder

            if ($removeSuccess) {
                $result.Removed++
            }
        }
        else {
            Write-Log "  User Google folder is NOT empty — leaving intact: $userGoogleFolder ($($profile.UserName))" "VERBOSE"
            $result.Skipped++
        }
    }

    return $result
}

function Invoke-Phase6FileSystemCleanup {
    <#
    .SYNOPSIS
        Orchestrates Phase 6: complete file system cleanup across all locations.
    .OUTPUTS
        [bool] True if phase completed. False on critical failure.
    #>
    [CmdletBinding()]
    param()

    Write-LogSection "PHASE 6: FILE SYSTEM CLEANUP"

    $phaseRemoved = 0
    $phaseSkipped = 0
    $phaseFailed = 0

    try {
        # --- 6.1 Machine-Level Directories ---
        Write-LogSubSection "6.1 — Machine-Level Directories"
        $machineResult = Remove-ChromeMachineFiles
        $phaseRemoved += $machineResult.Removed
        $phaseSkipped += $machineResult.Skipped
        $phaseFailed += $machineResult.Failed
        Write-Log ("Machine files: {0} removed, {1} skipped, {2} failed" -f
            $machineResult.Removed, $machineResult.Skipped, $machineResult.Failed) "INFO"

        # --- 6.2 Per-User Directories ---
        Write-LogSubSection "6.2 — Per-User Directories (All Profiles)"
        $userResult = Remove-ChromeUserFiles
        $phaseRemoved += $userResult.Removed
        $phaseSkipped += $userResult.Skipped
        $phaseFailed += $userResult.Failed
        Write-Log ("User files: {0} removed, {1} skipped, {2} failed" -f
            $userResult.Removed, $userResult.Skipped, $userResult.Failed) "INFO"

        # --- 6.3 Temp Files ---
        Write-LogSubSection "6.3 — Temporary Files"
        $tempResult = Remove-ChromeTempFiles
        $phaseRemoved += $tempResult.Removed
        $phaseSkipped += $tempResult.Skipped
        $phaseFailed += $tempResult.Failed
        Write-Log ("Temp files: {0} removed, {1} skipped, {2} failed" -f
            $tempResult.Removed, $tempResult.Skipped, $tempResult.Failed) "INFO"

        # --- 6.4 Windows Installer Cache ---
        Write-LogSubSection "6.4 — Windows Installer Cache"
        $installerResult = Remove-ChromeInstallerCache
        $phaseRemoved += $installerResult.Removed
        $phaseSkipped += $installerResult.Skipped
        $phaseFailed += $installerResult.Failed
        Write-Log ("Installer cache: {0} removed, {1} skipped, {2} failed" -f
            $installerResult.Removed, $installerResult.Skipped, $installerResult.Failed) "INFO"

        # --- 6.5 Prefetch Files ---
        Write-LogSubSection "6.5 — Prefetch Files"
        $prefetchResult = Remove-ChromePrefetchFiles
        $phaseRemoved += $prefetchResult.Removed
        $phaseSkipped += $prefetchResult.Skipped
        $phaseFailed += $prefetchResult.Failed
        Write-Log ("Prefetch files: {0} removed, {1} skipped, {2} failed" -f
            $prefetchResult.Removed, $prefetchResult.Skipped, $prefetchResult.Failed) "INFO"

        # --- 6.6 Shortcuts ---
        Write-LogSubSection "6.6 — Shortcuts (Desktop, Start Menu, Taskbar, Quick Launch, Startup)"
        $shortcutResult = Remove-ChromeShortcuts
        $phaseRemoved += $shortcutResult.Removed
        $phaseSkipped += $shortcutResult.Skipped
        $phaseFailed += $shortcutResult.Failed
        Write-Log ("Shortcuts: {0} removed, {1} skipped, {2} failed" -f
            $shortcutResult.Removed, $shortcutResult.Skipped, $shortcutResult.Failed) "INFO"

        # --- 6.7 Cleanup Empty Google Folders ---
        Write-LogSubSection "6.7 — Cleanup Empty Google Parent Folders"
        $emptyFolderResult = Remove-EmptyGoogleFolders
        $phaseRemoved += $emptyFolderResult.Removed
        $phaseSkipped += $emptyFolderResult.Skipped
        Write-Log ("Empty Google folders: {0} removed, {1} not empty (preserved)" -f
            $emptyFolderResult.Removed, $emptyFolderResult.Skipped) "INFO"

        # --- Phase Result ---
        Write-LogSubSection "Phase 6 Summary"
        Write-Log ("Phase 6 complete: {0} items removed, {1} items skipped, {2} items failed" -f
            $phaseRemoved, $phaseSkipped, $phaseFailed) "INFO"

        if ($phaseFailed -eq 0) {
            $script:phaseResults["Phase 6"] = "PASS"
        }
        elseif ($phaseRemoved -gt 0) {
            $script:phaseResults["Phase 6"] = "PARTIAL"
        }
        else {
            $script:phaseResults["Phase 6"] = "FAIL"
        }

        return $true
    }
    catch {
        Write-Log "Critical error during file system cleanup: $($_.Exception.Message)" "ERROR"
        $null = $script:errorCollection.Add([PSCustomObject]@{
                Timestamp   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Phase       = "Phase 6"
                Item        = "File System Cleanup Orchestration"
                Error       = $_.Exception.Message
                ErrorRecord = $_
            })
        $script:phaseResults["Phase 6"] = "FAIL"
        return $false
    }
}

#endregion

#region PHASE 7 — REGISTRY CLEANUP
# ============================================================================
# PHASE 7 — REGISTRY CLEANUP
# ============================================================================

function Remove-ChromeHKLMRegistry {
    <#
    .SYNOPSIS
        Removes all Chrome-related registry keys from HKLM (machine-wide).
        Handles both full key removal and specific value removal from shared keys.
    .OUTPUTS
        [PSCustomObject] with counts of removed, skipped, and failed items.
    #>
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        Removed = 0
        Skipped = 0
        Failed  = 0
    }

    # --- Remove full registry keys ---
    Write-Log "Processing $($CHROME_HKLM_KEYS.Count) HKLM registry key targets..." "VERBOSE"

    foreach ($regKey in $CHROME_HKLM_KEYS) {
        if (Test-PathSafe -Path $regKey) {
            $removeSuccess = Remove-RegistryKeyIfExists -Path $regKey -Phase "Phase 7"
            if ($removeSuccess -and -not (Test-PathSafe -Path $regKey)) {
                $result.Removed++
            }
            elseif ($removeSuccess) {
                $result.Skipped++
            }
            else {
                $result.Failed++
            }
        }
        else {
            Write-Log "HKLM key not found (skip): $regKey" "SKIP"
            $result.Skipped++
        }
    }

    # --- Remove specific values from shared keys ---
    Write-Log "Processing $($CHROME_HKLM_VALUES.Count) HKLM registry value target(s)..." "VERBOSE"

    foreach ($valueTarget in $CHROME_HKLM_VALUES) {
        $removeSuccess = Remove-RegistryValueIfExists -Path $valueTarget.Path -Name $valueTarget.Name -Phase "Phase 7"
        if ($removeSuccess) {
            # Check if the value is actually gone
            try {
                $checkValue = Get-ItemProperty -Path $valueTarget.Path -Name $valueTarget.Name -ErrorAction SilentlyContinue
                if ($null -eq $checkValue -or $null -eq $checkValue."$($valueTarget.Name)") {
                    $result.Removed++
                }
                else {
                    $result.Skipped++
                }
            }
            catch {
                $result.Removed++
            }
        }
        else {
            $result.Failed++
        }
    }

    # --- Dynamically scan and remove versioned ChromeHTML keys ---
    Write-Log "Scanning for versioned ChromeHTML class registrations..." "VERBOSE"

    $classRoots = @(
        "HKLM:\SOFTWARE\Classes"
        "HKLM:\SOFTWARE\WOW6432Node\Classes"
    )

    foreach ($classRoot in $classRoots) {
        if (-not (Test-PathSafe -Path $classRoot)) { continue }

        try {
            $chromeHTMLKeys = Get-ChildItem -Path $classRoot -ErrorAction SilentlyContinue |
            Where-Object { $_.PSChildName -match "^ChromeHTML" }

            if ($chromeHTMLKeys) {
                foreach ($key in $chromeHTMLKeys) {
                    Write-Log "Found versioned ChromeHTML key: $($key.PSPath)" "INFO"
                    $removeSuccess = Remove-RegistryKeyIfExists -Path $key.PSPath -Phase "Phase 7"
                    if ($removeSuccess -and -not (Test-PathSafe -Path $key.PSPath)) {
                        $result.Removed++
                    }
                    else {
                        $result.Failed++
                    }
                }
            }
            else {
                Write-Log "No versioned ChromeHTML keys found in: $classRoot" "SKIP"
            }
        }
        catch {
            Write-Log "Error scanning for ChromeHTML keys in $classRoot : $($_.Exception.Message)" "WARNING"
        }
    }

    return $result
}

function Remove-ChromeCOMRegistrations {
    <#
    .SYNOPSIS
        Dynamically discovers and removes Chrome-related COM/CLSID, AppID,
        and TypeLib registry entries. Scans by examining default values
        and InprocServer32/LocalServer32 paths for Google/Chrome references.
    .OUTPUTS
        [PSCustomObject] with counts of removed, skipped, and failed items.
    #>
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        Removed = 0
        Skipped = 0
        Failed  = 0
    }

    # --- CLSID Scan ---
    Write-Log "Scanning CLSID entries for Chrome COM registrations..." "VERBOSE"

    $clsidRoots = @(
        "HKLM:\SOFTWARE\Classes\CLSID"
        "HKLM:\SOFTWARE\Classes\WOW6432Node\CLSID"
        "HKLM:\SOFTWARE\WOW6432Node\Classes\CLSID"
    )

    foreach ($clsidRoot in $clsidRoots) {
        if (-not (Test-PathSafe -Path $clsidRoot)) { continue }

        try {
            $clsidKeys = Get-ChildItem -Path $clsidRoot -ErrorAction SilentlyContinue

            if (-not $clsidKeys) { continue }

            foreach ($clsidKey in $clsidKeys) {
                $isChromeRelated = $false

                try {
                    # Check default value of CLSID key
                    $defaultValue = (Get-ItemProperty -Path $clsidKey.PSPath -Name "(Default)" -ErrorAction SilentlyContinue)."(Default)"
                    if ($defaultValue -and ($defaultValue -match "Google Chrome" -or $defaultValue -match "Google Update" -or $defaultValue -match "GoogleCrash")) {
                        $isChromeRelated = $true
                    }

                    # Check InprocServer32 path
                    if (-not $isChromeRelated) {
                        $inprocPath = Join-Path -Path $clsidKey.PSPath -ChildPath "InprocServer32"
                        if (Test-PathSafe -Path $inprocPath) {
                            $inprocValue = (Get-ItemProperty -Path $inprocPath -Name "(Default)" -ErrorAction SilentlyContinue)."(Default)"
                            if ($inprocValue -and ($inprocValue -match "Google\\Chrome" -or $inprocValue -match "Google\\Update" -or $inprocValue -match "GoogleCrashHandler")) {
                                $isChromeRelated = $true
                            }
                        }
                    }

                    # Check LocalServer32 path
                    if (-not $isChromeRelated) {
                        $localServerPath = Join-Path -Path $clsidKey.PSPath -ChildPath "LocalServer32"
                        if (Test-PathSafe -Path $localServerPath) {
                            $localServerValue = (Get-ItemProperty -Path $localServerPath -Name "(Default)" -ErrorAction SilentlyContinue)."(Default)"
                            if ($localServerValue -and ($localServerValue -match "Google\\Chrome" -or $localServerValue -match "Google\\Update" -or $localServerValue -match "GoogleCrashHandler")) {
                                $isChromeRelated = $true
                            }
                        }
                    }

                    if ($isChromeRelated) {
                        Write-Log "Chrome COM CLSID found: $($clsidKey.PSChildName) in $clsidRoot" "INFO"
                        if ($defaultValue) { Write-Log "  Description: $defaultValue" "VERBOSE" }

                        $removeSuccess = Remove-RegistryKeyIfExists -Path $clsidKey.PSPath -Phase "Phase 7"
                        if ($removeSuccess -and -not (Test-PathSafe -Path $clsidKey.PSPath)) {
                            $result.Removed++
                        }
                        else {
                            $result.Failed++
                        }
                    }
                }
                catch {
                    # Skip unreadable CLSID entries
                    continue
                }
            }
        }
        catch {
            Write-Log "Error scanning CLSID in $clsidRoot : $($_.Exception.Message)" "WARNING"
        }
    }

    # --- AppID Scan ---
    Write-Log "Scanning AppID entries for Chrome registrations..." "VERBOSE"

    $appIdRoots = @(
        "HKLM:\SOFTWARE\Classes\AppID"
        "HKLM:\SOFTWARE\WOW6432Node\Classes\AppID"
    )

    foreach ($appIdRoot in $appIdRoots) {
        if (-not (Test-PathSafe -Path $appIdRoot)) { continue }

        try {
            $appIdKeys = Get-ChildItem -Path $appIdRoot -ErrorAction SilentlyContinue

            if (-not $appIdKeys) { continue }

            foreach ($appIdKey in $appIdKeys) {
                try {
                    $defaultValue = (Get-ItemProperty -Path $appIdKey.PSPath -Name "(Default)" -ErrorAction SilentlyContinue)."(Default)"
                    if ($defaultValue -and ($defaultValue -match "Google Chrome" -or $defaultValue -match "Google Update" -or $defaultValue -match "GoogleCrash" -or $defaultValue -match "chrome")) {
                        Write-Log "Chrome AppID found: $($appIdKey.PSChildName) — $defaultValue" "INFO"
                        $removeSuccess = Remove-RegistryKeyIfExists -Path $appIdKey.PSPath -Phase "Phase 7"
                        if ($removeSuccess -and -not (Test-PathSafe -Path $appIdKey.PSPath)) {
                            $result.Removed++
                        }
                        else {
                            $result.Failed++
                        }
                    }
                }
                catch {
                    continue
                }
            }
        }
        catch {
            Write-Log "Error scanning AppID in $appIdRoot : $($_.Exception.Message)" "WARNING"
        }
    }

    # --- TypeLib Scan ---
    Write-Log "Scanning TypeLib entries for Chrome registrations..." "VERBOSE"

    $typeLibRoots = @(
        "HKLM:\SOFTWARE\Classes\TypeLib"
        "HKLM:\SOFTWARE\WOW6432Node\Classes\TypeLib"
    )

    foreach ($typeLibRoot in $typeLibRoots) {
        if (-not (Test-PathSafe -Path $typeLibRoot)) { continue }

        try {
            $typeLibKeys = Get-ChildItem -Path $typeLibRoot -ErrorAction SilentlyContinue

            if (-not $typeLibKeys) { continue }

            foreach ($typeLibKey in $typeLibKeys) {
                try {
                    # TypeLibs have version subkeys, check each version
                    $versionKeys = Get-ChildItem -Path $typeLibKey.PSPath -ErrorAction SilentlyContinue
                    $isChromeTypeLib = $false

                    foreach ($versionKey in $versionKeys) {
                        $defaultValue = (Get-ItemProperty -Path $versionKey.PSPath -Name "(Default)" -ErrorAction SilentlyContinue)."(Default)"
                        if ($defaultValue -and ($defaultValue -match "Google Chrome" -or $defaultValue -match "Google Update" -or $defaultValue -match "GoogleCrash")) {
                            $isChromeTypeLib = $true
                            break
                        }

                        # Also check win32 path subkeys
                        $win32Paths = Get-ChildItem -Path $versionKey.PSPath -Recurse -ErrorAction SilentlyContinue
                        foreach ($win32Path in $win32Paths) {
                            $pathDefault = (Get-ItemProperty -Path $win32Path.PSPath -Name "(Default)" -ErrorAction SilentlyContinue)."(Default)"
                            if ($pathDefault -and ($pathDefault -match "Google\\Chrome" -or $pathDefault -match "Google\\Update" -or $pathDefault -match "GoogleCrashHandler")) {
                                $isChromeTypeLib = $true
                                break
                            }
                        }

                        if ($isChromeTypeLib) { break }
                    }

                    if ($isChromeTypeLib) {
                        Write-Log "Chrome TypeLib found: $($typeLibKey.PSChildName)" "INFO"
                        $removeSuccess = Remove-RegistryKeyIfExists -Path $typeLibKey.PSPath -Phase "Phase 7"
                        if ($removeSuccess -and -not (Test-PathSafe -Path $typeLibKey.PSPath)) {
                            $result.Removed++
                        }
                        else {
                            $result.Failed++
                        }
                    }
                }
                catch {
                    continue
                }
            }
        }
        catch {
            Write-Log "Error scanning TypeLib in $typeLibRoot : $($_.Exception.Message)" "WARNING"
        }
    }

    return $result
}

function Remove-ChromePolicyKeys {
    <#
    .SYNOPSIS
        Removes Chrome and Google Update policy registry keys from HKLM.
        These are typically set by Group Policy or enterprise management tools.
    .OUTPUTS
        [PSCustomObject] with counts of removed, skipped, and failed items.
    #>
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        Removed = 0
        Skipped = 0
        Failed  = 0
    }

    $policyPaths = @(
        "HKLM:\SOFTWARE\Policies\Google\Chrome"
        "HKLM:\SOFTWARE\Policies\Google\Update"
        "HKLM:\SOFTWARE\WOW6432Node\Policies\Google\Chrome"
        "HKLM:\SOFTWARE\WOW6432Node\Policies\Google\Update"
    )

    foreach ($policyPath in $policyPaths) {
        if (Test-PathSafe -Path $policyPath) {
            Write-Log "Chrome policy key found: $policyPath" "INFO"

            # Log policy contents before removal (useful for audit trail)
            try {
                $policyValues = Get-ItemProperty -Path $policyPath -ErrorAction SilentlyContinue
                if ($policyValues) {
                    $properties = $policyValues.PSObject.Properties |
                    Where-Object { $_.Name -notmatch "^PS" }
                    $policyCount = @($properties).Count
                    Write-Log "  Policy contains $policyCount setting(s)" "VERBOSE"
                    foreach ($prop in $properties) {
                        Write-Log "    Policy: $($prop.Name) = $($prop.Value)" "VERBOSE"
                    }
                }
            }
            catch {
                Write-Log "  Could not enumerate policy values" "VERBOSE"
            }

            $removeSuccess = Remove-RegistryKeyIfExists -Path $policyPath -Phase "Phase 7"
            if ($removeSuccess -and -not (Test-PathSafe -Path $policyPath)) {
                $result.Removed++
            }
            elseif ($removeSuccess) {
                $result.Skipped++
            }
            else {
                $result.Failed++
            }
        }
        else {
            Write-Log "Policy key not found (skip): $policyPath" "SKIP"
            $result.Skipped++
        }
    }

    return $result
}

function Remove-ChromeRunValues {
    <#
    .SYNOPSIS
        Removes Chrome and Google Update values from Run/RunOnce keys in a given
        registry root. Works for both HKLM and per-user hive roots.
    .PARAMETER RegistryRoot
        The registry root to scan (e.g., "HKLM:" or "Registry::HKEY_USERS\S-1-5-21-xxx").
    .OUTPUTS
        [PSCustomObject] with counts of removed, skipped, and failed items.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RegistryRoot
    )

    $result = [PSCustomObject]@{
        Removed = 0
        Skipped = 0
        Failed  = 0
    }

    $runPaths = @(
        "$RegistryRoot\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
        "$RegistryRoot\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
    )

    foreach ($runPath in $runPaths) {
        # Normalize path separators
        $runPath = $runPath -replace "\\\\", "\"

        if (-not (Test-PathSafe -Path $runPath)) {
            $result.Skipped++
            continue
        }

        try {
            $runProperties = Get-ItemProperty -Path $runPath -ErrorAction SilentlyContinue
            if ($null -eq $runProperties) { continue }

            $properties = $runProperties.PSObject.Properties |
            Where-Object { $_.Name -notmatch "^PS" }

            foreach ($prop in $properties) {
                $isGoogleRelated = $false

                # Check value name
                if ($prop.Name -match "Google" -or $prop.Name -match "Chrome") {
                    $isGoogleRelated = $true
                }

                # Check value data
                if (-not $isGoogleRelated -and $prop.Value) {
                    if ($prop.Value -match "Google" -or $prop.Value -match "Chrome" -or
                        $prop.Value -match "chrome\.exe" -or $prop.Value -match "GoogleUpdate") {
                        $isGoogleRelated = $true
                    }
                }

                if ($isGoogleRelated) {
                    Write-Log "Chrome/Google Run value found: $runPath\$($prop.Name) = $($prop.Value)" "INFO"
                    $removeSuccess = Remove-RegistryValueIfExists -Path $runPath -Name $prop.Name -Phase "Phase 7"
                    if ($removeSuccess) {
                        $result.Removed++
                    }
                    else {
                        $result.Failed++
                    }
                }
            }
        }
        catch {
            Write-Log "Error scanning Run key $runPath : $($_.Exception.Message)" "WARNING"
        }
    }

    return $result
}

function Remove-ChromeOpenWithProgids {
    <#
    .SYNOPSIS
        Removes ChromeHTML references from file extension OpenWithProgids keys
        in a given registry root.
    .PARAMETER RegistryRoot
        The registry root to scan (e.g., "HKLM:" or "Registry::HKEY_USERS\S-1-5-21-xxx").
    .OUTPUTS
        [PSCustomObject] with counts of removed, skipped, and failed items.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RegistryRoot
    )

    $result = [PSCustomObject]@{
        Removed = 0
        Skipped = 0
        Failed  = 0
    }

    foreach ($extension in $CHROME_OPENWITH_EXTENSIONS) {
        $openWithPath = "$RegistryRoot\SOFTWARE\Classes\$extension\OpenWithProgids"
        $openWithPath = $openWithPath -replace "\\\\", "\"

        if (-not (Test-PathSafe -Path $openWithPath)) {
            continue
        }

        try {
            $progIds = Get-ItemProperty -Path $openWithPath -ErrorAction SilentlyContinue
            if ($null -eq $progIds) { continue }

            $properties = $progIds.PSObject.Properties |
            Where-Object { $_.Name -notmatch "^PS" -and $_.Name -match "ChromeHTML" }

            foreach ($prop in $properties) {
                Write-Log "ChromeHTML OpenWithProgid found: $openWithPath\$($prop.Name)" "INFO"
                $removeSuccess = Remove-RegistryValueIfExists -Path $openWithPath -Name $prop.Name -Phase "Phase 7"
                if ($removeSuccess) {
                    $result.Removed++
                }
                else {
                    $result.Failed++
                }
            }
        }
        catch {
            Write-Log "Error scanning OpenWithProgids for $extension : $($_.Exception.Message)" "VERBOSE"
        }
    }

    # --- Also check user-specific FileExts (HKCU-level only) ---
    if ($RegistryRoot -ne "HKLM:") {
        foreach ($extension in $CHROME_OPENWITH_EXTENSIONS) {
            $fileExtPath = "$RegistryRoot\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$extension\OpenWithProgids"
            $fileExtPath = $fileExtPath -replace "\\\\", "\"

            if (-not (Test-PathSafe -Path $fileExtPath)) {
                continue
            }

            try {
                $progIds = Get-ItemProperty -Path $fileExtPath -ErrorAction SilentlyContinue
                if ($null -eq $progIds) { continue }

                $properties = $progIds.PSObject.Properties |
                Where-Object { $_.Name -notmatch "^PS" -and $_.Name -match "ChromeHTML" }

                foreach ($prop in $properties) {
                    Write-Log "ChromeHTML FileExts OpenWithProgid found: $fileExtPath\$($prop.Name)" "INFO"
                    $removeSuccess = Remove-RegistryValueIfExists -Path $fileExtPath -Name $prop.Name -Phase "Phase 7"
                    if ($removeSuccess) {
                        $result.Removed++
                    }
                    else {
                        $result.Failed++
                    }
                }
            }
            catch {
                Write-Log "Error scanning FileExts for $extension : $($_.Exception.Message)" "VERBOSE"
            }
        }
    }

    return $result
}

function Remove-ChromeFirewallRegistryValues {
    <#
    .SYNOPSIS
        Removes Chrome-related values from the firewall rules registry key.
        This complements Phase 9 (netsh/API-level removal) by cleaning residual
        registry entries.
    .OUTPUTS
        [PSCustomObject] with counts of removed, skipped, and failed items.
    #>
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        Removed = 0
        Skipped = 0
        Failed  = 0
    }

    $firewallRegPath = "HKLM:\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\FirewallRules"

    if (-not (Test-PathSafe -Path $firewallRegPath)) {
        Write-Log "Firewall rules registry key not found" "SKIP"
        return $result
    }

    try {
        $firewallProps = Get-ItemProperty -Path $firewallRegPath -ErrorAction SilentlyContinue
        if ($null -eq $firewallProps) { return $result }

        $properties = $firewallProps.PSObject.Properties |
        Where-Object { $_.Name -notmatch "^PS" }

        foreach ($prop in $properties) {
            if ($prop.Value -and ($prop.Value -match "chrome\.exe" -or
                    $prop.Value -match "Google\\Chrome" -or
                    $prop.Value -match "GoogleUpdate\.exe" -or
                    $prop.Value -match "Google\\Update")) {

                Write-Log "Chrome firewall rule registry value found: $($prop.Name)" "INFO"
                Write-Log "  Value: $($prop.Value)" "VERBOSE"
                $removeSuccess = Remove-RegistryValueIfExists -Path $firewallRegPath -Name $prop.Name -Phase "Phase 7"
                if ($removeSuccess) {
                    $result.Removed++
                }
                else {
                    $result.Failed++
                }
            }
        }
    }
    catch {
        Write-Log "Error scanning firewall registry: $($_.Exception.Message)" "WARNING"
    }

    return $result
}

function Remove-ChromeAppCompatFlags {
    <#
    .SYNOPSIS
        Removes Chrome-related entries from AppCompatFlags\Layers registry.
    .OUTPUTS
        [PSCustomObject] with counts of removed, skipped, and failed items.
    #>
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        Removed = 0
        Skipped = 0
        Failed  = 0
    }

    $appCompatPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers"

    if (-not (Test-PathSafe -Path $appCompatPath)) {
        Write-Log "AppCompatFlags\Layers key not found" "SKIP"
        return $result
    }

    try {
        $appCompatProps = Get-ItemProperty -Path $appCompatPath -ErrorAction SilentlyContinue
        if ($null -eq $appCompatProps) { return $result }

        $properties = $appCompatProps.PSObject.Properties |
        Where-Object { $_.Name -notmatch "^PS" }

        foreach ($prop in $properties) {
            if ($prop.Name -match "chrome\.exe" -or $prop.Name -match "Google\\Chrome" -or
                $prop.Name -match "GoogleUpdate" -or $prop.Name -match "Google\\Update") {

                Write-Log "Chrome AppCompat entry found: $($prop.Name) = $($prop.Value)" "INFO"
                $removeSuccess = Remove-RegistryValueIfExists -Path $appCompatPath -Name $prop.Name -Phase "Phase 7"
                if ($removeSuccess) {
                    $result.Removed++
                }
                else {
                    $result.Failed++
                }
            }
        }
    }
    catch {
        Write-Log "Error scanning AppCompatFlags: $($_.Exception.Message)" "WARNING"
    }

    return $result
}

function Remove-ChromeHKLMUninstallGUIDs {
    <#
    .SYNOPSIS
        Scans HKLM Uninstall keys for any GUID-based entries that are Chrome-related.
        This catches entries that don't use "Google Chrome" as the key name.
    .OUTPUTS
        [PSCustomObject] with counts of removed, skipped, and failed items.
    #>
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        Removed = 0
        Skipped = 0
        Failed  = 0
    }

    $uninstallRoots = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )

    foreach ($uninstallRoot in $uninstallRoots) {
        if (-not (Test-PathSafe -Path $uninstallRoot)) { continue }

        try {
            $subKeys = Get-ChildItem -Path $uninstallRoot -ErrorAction SilentlyContinue

            foreach ($subKey in $subKeys) {
                # Skip the named key we already handled
                if ($subKey.PSChildName -eq "Google Chrome") { continue }

                try {
                    $props = Get-ItemProperty -Path $subKey.PSPath -ErrorAction SilentlyContinue
                    $displayName = $props.DisplayName
                    $publisher = $props.Publisher
                    $installLocation = $props.InstallLocation

                    $isChromeRelated = $false

                    if ($displayName -and $displayName -match "Google Chrome") {
                        $isChromeRelated = $true
                    }
                    elseif ($publisher -and $publisher -match "Google" -and
                        $installLocation -and $installLocation -match "Chrome") {
                        $isChromeRelated = $true
                    }

                    if ($isChromeRelated) {
                        Write-Log "Chrome GUID uninstall entry found: $($subKey.PSChildName) — $displayName" "INFO"
                        $removeSuccess = Remove-RegistryKeyIfExists -Path $subKey.PSPath -Phase "Phase 7"
                        if ($removeSuccess -and -not (Test-PathSafe -Path $subKey.PSPath)) {
                            $result.Removed++
                        }
                        else {
                            $result.Failed++
                        }
                    }
                }
                catch {
                    continue
                }
            }
        }
        catch {
            Write-Log "Error scanning uninstall GUIDs in $uninstallRoot : $($_.Exception.Message)" "WARNING"
        }
    }

    return $result
}

function Remove-ChromeUserRegistryFromHive {
    <#
    .SYNOPSIS
        Removes all Chrome-related registry keys and values from a single user
        registry hive. Works with both loaded hives (HKU:\SID) and mounted
        offline hives (HKU:\TempHive_SID).
    .PARAMETER HiveRoot
        The full registry root for this user hive
        (e.g., "Registry::HKEY_USERS\S-1-5-21-xxx").
    .PARAMETER UserName
        The user name for logging purposes.
    .PARAMETER SID
        The user SID for logging purposes.
    .OUTPUTS
        [PSCustomObject] with counts of removed, skipped, and failed items.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$HiveRoot,

        [Parameter(Mandatory = $false)]
        [string]$UserName = "Unknown",

        [Parameter(Mandatory = $false)]
        [string]$SID = "Unknown"
    )

    $result = [PSCustomObject]@{
        Removed = 0
        Skipped = 0
        Failed  = 0
    }

    Write-Log "Cleaning user registry hive: $UserName ($SID)" "VERBOSE"
    Write-Log "  Hive root: $HiveRoot" "VERBOSE"

    # --- Remove Chrome registry keys ---
    foreach ($relativeKey in $CHROME_USER_REGISTRY_KEYS) {
        $fullPath = Join-Path -Path $HiveRoot -ChildPath $relativeKey

        if (Test-PathSafe -Path $fullPath) {
            Write-Log "  User Chrome key found: $relativeKey" "INFO"
            $removeSuccess = Remove-RegistryKeyIfExists -Path $fullPath -Phase "Phase 7"
            if ($removeSuccess -and -not (Test-PathSafe -Path $fullPath)) {
                $result.Removed++
            }
            elseif ($removeSuccess) {
                $result.Skipped++
            }
            else {
                $result.Failed++
            }
        }
        else {
            $result.Skipped++
        }
    }

    # --- Remove Chrome uninstall GUID entries ---
    $userUninstallRoot = Join-Path -Path $HiveRoot -ChildPath "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    if (Test-PathSafe -Path $userUninstallRoot) {
        try {
            $subKeys = Get-ChildItem -Path $userUninstallRoot -ErrorAction SilentlyContinue
            foreach ($subKey in $subKeys) {
                if ($subKey.PSChildName -eq "Google Chrome") { continue }  # Already handled above

                try {
                    $displayName = (Get-ItemProperty -Path $subKey.PSPath -ErrorAction SilentlyContinue).DisplayName
                    if ($displayName -and $displayName -match "Google Chrome") {
                        Write-Log "  User Chrome GUID uninstall entry: $($subKey.PSChildName)" "INFO"
                        $removeSuccess = Remove-RegistryKeyIfExists -Path $subKey.PSPath -Phase "Phase 7"
                        if ($removeSuccess -and -not (Test-PathSafe -Path $subKey.PSPath)) {
                            $result.Removed++
                        }
                        else {
                            $result.Failed++
                        }
                    }
                }
                catch { continue }
            }
        }
        catch {
            Write-Log "  Error scanning user uninstall keys: $($_.Exception.Message)" "VERBOSE"
        }
    }

    # --- Remove versioned ChromeHTML class keys ---
    $userClassesRoot = Join-Path -Path $HiveRoot -ChildPath "SOFTWARE\Classes"
    if (Test-PathSafe -Path $userClassesRoot) {
        try {
            $chromeHTMLKeys = Get-ChildItem -Path $userClassesRoot -ErrorAction SilentlyContinue |
            Where-Object { $_.PSChildName -match "^ChromeHTML" }

            foreach ($key in $chromeHTMLKeys) {
                Write-Log "  User versioned ChromeHTML key: $($key.PSChildName)" "INFO"
                $removeSuccess = Remove-RegistryKeyIfExists -Path $key.PSPath -Phase "Phase 7"
                if ($removeSuccess -and -not (Test-PathSafe -Path $key.PSPath)) {
                    $result.Removed++
                }
                else {
                    $result.Failed++
                }
            }
        }
        catch {
            Write-Log "  Error scanning user Classes for ChromeHTML: $($_.Exception.Message)" "VERBOSE"
        }
    }

    # --- Remove Run/RunOnce values ---
    $runResult = Remove-ChromeRunValues -RegistryRoot $HiveRoot
    $result.Removed += $runResult.Removed
    $result.Skipped += $runResult.Skipped
    $result.Failed += $runResult.Failed

    # --- Remove OpenWithProgids references ---
    $openWithResult = Remove-ChromeOpenWithProgids -RegistryRoot $HiveRoot
    $result.Removed += $openWithResult.Removed
    $result.Skipped += $openWithResult.Skipped
    $result.Failed += $openWithResult.Failed

    # --- Clean empty parent Google keys ---
    $userGoogleKeys = @(
        Join-Path -Path $HiveRoot -ChildPath "SOFTWARE\Google"
        Join-Path -Path $HiveRoot -ChildPath "SOFTWARE\Policies\Google"
    )

    foreach ($googleKey in $userGoogleKeys) {
        if (Test-PathSafe -Path $googleKey) {
            if (Get-RegistryKeyIsEmpty -Path $googleKey) {
                Write-Log "  Empty user Google parent key — removing: $googleKey" "INFO"
                $removeSuccess = Invoke-Action -Description "Remove empty user Google key: $googleKey" -Action {
                    Remove-Item -Path $googleKey -Force -ErrorAction Stop
                } -Phase "Phase 7" -Item $googleKey

                if ($removeSuccess) {
                    $result.Removed++
                }
            }
            else {
                Write-Log "  User Google parent key not empty — leaving intact: $googleKey" "VERBOSE"
            }
        }
    }

    Write-Log "  User $UserName registry cleanup: $($result.Removed) removed, $($result.Skipped) skipped, $($result.Failed) failed" "VERBOSE"
    return $result
}

function Remove-ChromeLoadedUserRegistry {
    <#
    .SYNOPSIS
        Cleans Chrome registry keys from all currently loaded user hives
        (users who are logged in or whose hive is loaded).
    .OUTPUTS
        [PSCustomObject] with counts of removed, skipped, and failed items.
    #>
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        Removed = 0
        Skipped = 0
        Failed  = 0
    }

    $loadedProfiles = $script:userProfiles | Where-Object { $_.IsLoaded -eq $true }

    if ($loadedProfiles.Count -eq 0) {
        Write-Log "No loaded user hives to process" "SKIP"
        return $result
    }

    Write-Log "Processing $($loadedProfiles.Count) loaded user hive(s)..." "INFO"

    foreach ($profile in $loadedProfiles) {
        $hiveRoot = "Registry::HKEY_USERS\$($profile.SID)"

        if (-not (Test-PathSafe -Path $hiveRoot)) {
            Write-Log "Loaded hive not accessible: $hiveRoot ($($profile.UserName))" "WARNING"
            continue
        }

        $userResult = Remove-ChromeUserRegistryFromHive -HiveRoot $hiveRoot -UserName $profile.UserName -SID $profile.SID
        $result.Removed += $userResult.Removed
        $result.Skipped += $userResult.Skipped
        $result.Failed += $userResult.Failed
    }

    return $result
}

function Remove-ChromeOfflineUserRegistry {
    <#
    .SYNOPSIS
        Cleans Chrome registry keys from all offline user hives (users NOT logged in).
        Loads NTUSER.DAT, cleans, then unloads.
    .OUTPUTS
        [PSCustomObject] with counts of removed, skipped, and failed items.
    #>
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        Removed = 0
        Skipped = 0
        Failed  = 0
    }

    $offlineProfiles = $script:userProfiles | Where-Object { $_.IsLoaded -eq $false -and $_.HasNtuserDat -eq $true }

    if ($offlineProfiles.Count -eq 0) {
        Write-Log "No offline user hives to process" "SKIP"
        return $result
    }

    Write-Log "Processing $($offlineProfiles.Count) offline user hive(s)..." "INFO"

    foreach ($profile in $offlineProfiles) {
        $mountName = "TempHive_$($profile.SID)"
        $hiveRoot = "Registry::HKEY_USERS\$mountName"

        Write-Log "Processing offline hive for: $($profile.UserName) ($($profile.SID))" "INFO"

        # --- Load the hive ---
        if ($DryRun) {
            Write-Log "[WOULD LOAD] NTUSER.DAT: $($profile.NtuserDatPath) as HKU\$mountName" "DRYRUN"
            Write-Log "[WOULD SCAN] Offline hive for Chrome registry keys: $($profile.UserName)" "DRYRUN"
            Write-Log "  Skipping offline hive mount in DryRun mode to avoid file lock mutations" "DRYRUN"
            continue
        }

        # --- Normal (non-DryRun) execution ---
        $loadSuccess = Mount-OfflineHive -HivePath $profile.NtuserDatPath -MountName $mountName

        if (-not $loadSuccess) {
            Write-Log "Failed to load offline hive for: $($profile.UserName) — skipping this profile" "ERROR"
            $null = $script:errorCollection.Add([PSCustomObject]@{
                    Timestamp   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    Phase       = "Phase 7"
                    Item        = "Load hive: $($profile.UserName)"
                    Error       = "Failed to load NTUSER.DAT from $($profile.NtuserDatPath)"
                    ErrorRecord = $null
                })
            continue
        }

        # --- Clean the hive ---
        try {
            $userResult = Remove-ChromeUserRegistryFromHive -HiveRoot $hiveRoot -UserName $profile.UserName -SID $profile.SID
            $result.Removed += $userResult.Removed
            $result.Skipped += $userResult.Skipped
            $result.Failed += $userResult.Failed
        }
        catch {
            Write-Log "Error cleaning offline hive for $($profile.UserName): $($_.Exception.Message)" "ERROR"
            $result.Failed++
        }
        finally {
            # --- ALWAYS unload the hive, even if cleanup failed ---
            # Clear any .NET references first
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()

            $unloadSuccess = Dismount-OfflineHive -MountName $mountName

            if (-not $unloadSuccess) {
                Write-Log "WARNING: Failed to unload hive for $($profile.UserName). A reboot may be required." "WARNING"
                $null = $script:errorCollection.Add([PSCustomObject]@{
                        Timestamp   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                        Phase       = "Phase 7"
                        Item        = "Unload hive: $($profile.UserName)"
                        Error       = "Failed to unload NTUSER.DAT hive HKU\$mountName"
                        ErrorRecord = $null
                    })
            }
        }
    }

    return $result
}

function Remove-ChromeDefaultUserRegistry {
    <#
    .SYNOPSIS
        Cleans Chrome registry keys from the Default User profile NTUSER.DAT.
        This prevents new user profiles from inheriting Chrome artifacts.
    .OUTPUTS
        [PSCustomObject] with counts of removed, skipped, and failed items.
    #>
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        Removed = 0
        Skipped = 0
        Failed  = 0
    }

    $defaultUserPath = Join-Path -Path (Split-Path -Path $env:USERPROFILE -Parent) -ChildPath "Default"
    $defaultNtuserDat = Join-Path -Path $defaultUserPath -ChildPath "NTUSER.DAT"

    # Fallback path
    if (-not (Test-PathSafe -Path $defaultNtuserDat)) {
        $defaultNtuserDat = "C:\Users\Default\NTUSER.DAT"
    }

    if (-not (Test-PathSafe -Path $defaultNtuserDat)) {
        Write-Log "Default User NTUSER.DAT not found — skipping" "SKIP"
        return $result
    }

    Write-Log "Processing Default User hive: $defaultNtuserDat" "INFO"

    $mountName = "TempHive_DefaultUser"
    $hiveRoot = "Registry::HKEY_USERS\$mountName"

    $loadSuccess = Mount-OfflineHive -HivePath $defaultNtuserDat -MountName $mountName

    if (-not $loadSuccess) {
        Write-Log "Failed to load Default User hive" "WARNING"
        return $result
    }

    try {
        $userResult = Remove-ChromeUserRegistryFromHive -HiveRoot $hiveRoot -UserName "Default User" -SID "Default"
        $result.Removed += $userResult.Removed
        $result.Skipped += $userResult.Skipped
        $result.Failed += $userResult.Failed
    }
    catch {
        Write-Log "Error cleaning Default User hive: $($_.Exception.Message)" "ERROR"
        $result.Failed++
    }
    finally {
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()

        $unloadSuccess = Dismount-OfflineHive -MountName $mountName
        if (-not $unloadSuccess) {
            Write-Log "WARNING: Failed to unload Default User hive. A reboot may be required." "WARNING"
        }
    }

    return $result
}

function Remove-EmptyGoogleRegistryKeys {
    <#
    .SYNOPSIS
        Removes empty parent Google registry keys from HKLM after
        Chrome subkeys have been removed.
    .OUTPUTS
        [PSCustomObject] with counts of removed and skipped keys.
    #>
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        Removed = 0
        Skipped = 0
    }

    foreach ($parentKey in $GOOGLE_PARENT_REGISTRY_KEYS) {
        if (-not (Test-PathSafe -Path $parentKey)) {
            $result.Skipped++
            continue
        }

        if (Get-RegistryKeyIsEmpty -Path $parentKey) {
            Write-Log "Empty Google parent registry key — removing: $parentKey" "INFO"
            $removeSuccess = Invoke-Action -Description "Remove empty Google registry key: $parentKey" -Action {
                Remove-Item -Path $parentKey -Force -ErrorAction Stop
            } -Phase "Phase 7" -Item $parentKey

            if ($removeSuccess) {
                $result.Removed++
            }
        }
        else {
            Write-Log "Google parent registry key not empty — leaving intact: $parentKey" "INFO"
            $result.Skipped++

            # Log remaining subkeys
            try {
                $remainingKeys = Get-ChildItem -Path $parentKey -ErrorAction SilentlyContinue
                foreach ($key in $remainingKeys) {
                    Write-Log "  Remaining subkey: $($key.PSChildName)" "VERBOSE"
                }
            }
            catch {
                # Non-critical
            }
        }
    }

    return $result
}

function Invoke-Phase7RegistryCleanup {
    <#
    .SYNOPSIS
        Orchestrates Phase 7: complete registry cleanup across HKLM,
        loaded user hives, offline hives, and Default User.
    .OUTPUTS
        [bool] True if phase completed. False on critical failure.
    #>
    [CmdletBinding()]
    param()

    Write-LogSection "PHASE 7: REGISTRY CLEANUP"

    $phaseRemoved = 0
    $phaseSkipped = 0
    $phaseFailed = 0

    try {
        # --- 7.1 HKLM Machine-Wide Keys ---
        Write-LogSubSection "7.1 — HKLM Machine-Wide Registry Keys"
        $hklmResult = Remove-ChromeHKLMRegistry
        $phaseRemoved += $hklmResult.Removed
        $phaseSkipped += $hklmResult.Skipped
        $phaseFailed += $hklmResult.Failed
        Write-Log ("HKLM keys: {0} removed, {1} skipped, {2} failed" -f
            $hklmResult.Removed, $hklmResult.Skipped, $hklmResult.Failed) "INFO"

        # --- 7.2 HKLM GUID-Based Uninstall Entries ---
        Write-LogSubSection "7.2 — HKLM GUID-Based Uninstall Entries"
        $guidResult = Remove-ChromeHKLMUninstallGUIDs
        $phaseRemoved += $guidResult.Removed
        $phaseSkipped += $guidResult.Skipped
        $phaseFailed += $guidResult.Failed
        Write-Log ("GUID uninstall entries: {0} removed, {1} skipped, {2} failed" -f
            $guidResult.Removed, $guidResult.Skipped, $guidResult.Failed) "INFO"

        # --- 7.3 COM/CLSID/AppID/TypeLib Registrations ---
        Write-LogSubSection "7.3 — COM/CLSID/AppID/TypeLib Registrations"
        $comResult = Remove-ChromeCOMRegistrations
        $phaseRemoved += $comResult.Removed
        $phaseSkipped += $comResult.Skipped
        $phaseFailed += $comResult.Failed
        Write-Log ("COM registrations: {0} removed, {1} skipped, {2} failed" -f
            $comResult.Removed, $comResult.Skipped, $comResult.Failed) "INFO"

        # --- 7.4 Policy Keys ---
        Write-LogSubSection "7.4 — Chrome Policy Registry Keys"
        $policyResult = Remove-ChromePolicyKeys
        $phaseRemoved += $policyResult.Removed
        $phaseSkipped += $policyResult.Skipped
        $phaseFailed += $policyResult.Failed
        Write-Log ("Policy keys: {0} removed, {1} skipped, {2} failed" -f
            $policyResult.Removed, $policyResult.Skipped, $policyResult.Failed) "INFO"

        # --- 7.5 HKLM Run/RunOnce Values ---
        Write-LogSubSection "7.5 — HKLM Run/RunOnce Values"
        $runResult = Remove-ChromeRunValues -RegistryRoot "HKLM:"
        $phaseRemoved += $runResult.Removed
        $phaseSkipped += $runResult.Skipped
        $phaseFailed += $runResult.Failed
        Write-Log ("HKLM Run values: {0} removed, {1} skipped, {2} failed" -f
            $runResult.Removed, $runResult.Skipped, $runResult.Failed) "INFO"

        # --- 7.6 HKLM OpenWithProgids ---
        Write-LogSubSection "7.6 — HKLM OpenWithProgids"
        $openWithResult = Remove-ChromeOpenWithProgids -RegistryRoot "HKLM:"
        $phaseRemoved += $openWithResult.Removed
        $phaseSkipped += $openWithResult.Skipped
        $phaseFailed += $openWithResult.Failed
        Write-Log ("HKLM OpenWithProgids: {0} removed, {1} skipped, {2} failed" -f
            $openWithResult.Removed, $openWithResult.Skipped, $openWithResult.Failed) "INFO"

        # --- 7.7 Firewall Registry Values ---
        Write-LogSubSection "7.7 — Firewall Registry Values"
        $fwRegResult = Remove-ChromeFirewallRegistryValues
        $phaseRemoved += $fwRegResult.Removed
        $phaseSkipped += $fwRegResult.Skipped
        $phaseFailed += $fwRegResult.Failed
        Write-Log ("Firewall registry values: {0} removed, {1} skipped, {2} failed" -f
            $fwRegResult.Removed, $fwRegResult.Skipped, $fwRegResult.Failed) "INFO"

        # --- 7.8 AppCompatFlags ---
        Write-LogSubSection "7.8 — AppCompatFlags"
        $appCompatResult = Remove-ChromeAppCompatFlags
        $phaseRemoved += $appCompatResult.Removed
        $phaseSkipped += $appCompatResult.Skipped
        $phaseFailed += $appCompatResult.Failed
        Write-Log ("AppCompat entries: {0} removed, {1} skipped, {2} failed" -f
            $appCompatResult.Removed, $appCompatResult.Skipped, $appCompatResult.Failed) "INFO"

        # --- 7.9 Loaded User Hives ---
        Write-LogSubSection "7.9 — Loaded User Hives (Currently Logged-In Users)"
        $loadedUserResult = Remove-ChromeLoadedUserRegistry
        $phaseRemoved += $loadedUserResult.Removed
        $phaseSkipped += $loadedUserResult.Skipped
        $phaseFailed += $loadedUserResult.Failed
        Write-Log ("Loaded user hives: {0} removed, {1} skipped, {2} failed" -f
            $loadedUserResult.Removed, $loadedUserResult.Skipped, $loadedUserResult.Failed) "INFO"

        # --- 7.10 Offline User Hives ---
        Write-LogSubSection "7.10 — Offline User Hives (Not Currently Logged In)"
        $offlineUserResult = Remove-ChromeOfflineUserRegistry
        $phaseRemoved += $offlineUserResult.Removed
        $phaseSkipped += $offlineUserResult.Skipped
        $phaseFailed += $offlineUserResult.Failed
        Write-Log ("Offline user hives: {0} removed, {1} skipped, {2} failed" -f
            $offlineUserResult.Removed, $offlineUserResult.Skipped, $offlineUserResult.Failed) "INFO"

        # --- 7.11 Default User Hive ---
        Write-LogSubSection "7.11 — Default User Profile Hive"
        $defaultUserResult = Remove-ChromeDefaultUserRegistry
        $phaseRemoved += $defaultUserResult.Removed
        $phaseSkipped += $defaultUserResult.Skipped
        $phaseFailed += $defaultUserResult.Failed
        Write-Log ("Default User hive: {0} removed, {1} skipped, {2} failed" -f
            $defaultUserResult.Removed, $defaultUserResult.Skipped, $defaultUserResult.Failed) "INFO"

        # --- 7.12 Cleanup Empty Parent Google Keys ---
        Write-LogSubSection "7.12 — Cleanup Empty Google Parent Registry Keys"
        $emptyKeyResult = Remove-EmptyGoogleRegistryKeys
        $phaseRemoved += $emptyKeyResult.Removed
        $phaseSkipped += $emptyKeyResult.Skipped
        Write-Log ("Empty Google parent keys: {0} removed, {1} not empty (preserved)" -f
            $emptyKeyResult.Removed, $emptyKeyResult.Skipped) "INFO"

        # --- Phase Result ---
        Write-LogSubSection "Phase 7 Summary"
        Write-Log ("Phase 7 complete: {0} items removed, {1} items skipped, {2} items failed" -f
            $phaseRemoved, $phaseSkipped, $phaseFailed) "INFO"

        if ($phaseFailed -eq 0) {
            $script:phaseResults["Phase 7"] = "PASS"
        }
        elseif ($phaseRemoved -gt 0) {
            $script:phaseResults["Phase 7"] = "PARTIAL"
        }
        else {
            $script:phaseResults["Phase 7"] = "FAIL"
        }

        return $true
    }
    catch {
        Write-Log "Critical error during registry cleanup: $($_.Exception.Message)" "ERROR"
        $null = $script:errorCollection.Add([PSCustomObject]@{
                Timestamp   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Phase       = "Phase 7"
                Item        = "Registry Cleanup Orchestration"
                Error       = $_.Exception.Message
                ErrorRecord = $_
            })
        $script:phaseResults["Phase 7"] = "FAIL"
        return $false
    }
}

#endregion

#region PHASE 8 — WMI CLEANUP
# ============================================================================
# PHASE 8 — WMI CLEANUP
# ============================================================================

function Remove-ChromeInstallerReg {
    <#
    .SYNOPSIS
        Removes Chrome entries from the Windows Installer registration in the registry.
        Covers Products, Features, UpgradeCodes, and UserData installer keys.
    .OUTPUTS
        [PSCustomObject] with counts of removed, skipped, and failed items.
    #>
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        Removed = 0
        Skipped = 0
        Failed  = 0
    }

    # --- Installer Products ---
    $installerRoots = @(
        "HKLM:\SOFTWARE\Classes\Installer\Products"
        "HKLM:\SOFTWARE\Classes\Installer\Features"
        "HKLM:\SOFTWARE\Classes\Installer\UpgradeCodes"
    )

    foreach ($installerRoot in $installerRoots) {
        if (-not (Test-PathSafe -Path $installerRoot)) {
            Write-Log "Installer registry root not found (skip): $installerRoot" "SKIP"
            $result.Skipped++
            continue
        }

        Write-Log "Scanning installer registry: $installerRoot" "VERBOSE"

        try {
            $subKeys = Get-ChildItem -Path $installerRoot -ErrorAction SilentlyContinue

            if (-not $subKeys) {
                $result.Skipped++
                continue
            }

            foreach ($subKey in $subKeys) {
                try {
                    $isChromeRelated = $false

                    # Check ProductName value
                    $productName = (Get-ItemProperty -Path $subKey.PSPath -Name "ProductName" -ErrorAction SilentlyContinue).ProductName
                    if ($productName -and $productName -match "Google Chrome") {
                        $isChromeRelated = $true
                    }

                    # For Features/UpgradeCodes, check subkeys for Chrome references
                    if (-not $isChromeRelated) {
                        $defaultValue = (Get-ItemProperty -Path $subKey.PSPath -Name "(Default)" -ErrorAction SilentlyContinue)."(Default)"
                        if ($defaultValue -and $defaultValue -match "Google Chrome") {
                            $isChromeRelated = $true
                        }
                    }

                    # For UpgradeCodes, the subkey values are packed product GUIDs — check each
                    if (-not $isChromeRelated -and $installerRoot -match "UpgradeCodes") {
                        $upgradeValues = Get-ItemProperty -Path $subKey.PSPath -ErrorAction SilentlyContinue
                        if ($upgradeValues) {
                            $props = $upgradeValues.PSObject.Properties | Where-Object { $_.Name -notmatch "^PS" }
                            foreach ($prop in $props) {
                                # Each value name is a packed product GUID — check if it maps to Chrome
                                # We'll check the Products key with this GUID
                                $packedGuid = $prop.Name
                                $productKeyPath = "HKLM:\SOFTWARE\Classes\Installer\Products\$packedGuid"
                                if (Test-PathSafe -Path $productKeyPath) {
                                    $prodName = (Get-ItemProperty -Path $productKeyPath -Name "ProductName" -ErrorAction SilentlyContinue).ProductName
                                    if ($prodName -and $prodName -match "Google Chrome") {
                                        $isChromeRelated = $true
                                        break
                                    }
                                }
                            }
                        }
                    }

                    if ($isChromeRelated) {
                        Write-Log "Chrome installer entry found: $($subKey.PSPath)" "INFO"
                        if ($productName) { Write-Log "  ProductName: $productName" "VERBOSE" }

                        $removeSuccess = Remove-RegistryKeyIfExists -Path $subKey.PSPath -Phase "Phase 8"
                        if ($removeSuccess -and -not (Test-PathSafe -Path $subKey.PSPath)) {
                            $result.Removed++
                        }
                        else {
                            $result.Failed++
                        }
                    }
                }
                catch {
                    continue
                }
            }
        }
        catch {
            Write-Log "Error scanning $installerRoot : $($_.Exception.Message)" "WARNING"
        }
    }

    # --- UserData installer entries ---
    Write-Log "Scanning UserData installer entries..." "VERBOSE"

    $userDataRoots = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData"
    )

    foreach ($userDataRoot in $userDataRoots) {
        if (-not (Test-PathSafe -Path $userDataRoot)) { continue }

        try {
            # Each SID has a Products subkey
            $sidKeys = Get-ChildItem -Path $userDataRoot -ErrorAction SilentlyContinue
            foreach ($sidKey in $sidKeys) {
                $productsPath = Join-Path -Path $sidKey.PSPath -ChildPath "Products"
                if (-not (Test-PathSafe -Path $productsPath)) { continue }

                $productSubKeys = Get-ChildItem -Path $productsPath -ErrorAction SilentlyContinue
                foreach ($productKey in $productSubKeys) {
                    try {
                        $installPropPath = Join-Path -Path $productKey.PSPath -ChildPath "InstallProperties"
                        if (Test-PathSafe -Path $installPropPath) {
                            $displayName = (Get-ItemProperty -Path $installPropPath -Name "DisplayName" -ErrorAction SilentlyContinue).DisplayName
                            if ($displayName -and $displayName -match "Google Chrome") {
                                Write-Log "Chrome UserData installer entry found: $($productKey.PSPath)" "INFO"
                                Write-Log "  DisplayName: $displayName" "VERBOSE"

                                $removeSuccess = Remove-RegistryKeyIfExists -Path $productKey.PSPath -Phase "Phase 8"
                                if ($removeSuccess -and -not (Test-PathSafe -Path $productKey.PSPath)) {
                                    $result.Removed++
                                }
                                else {
                                    $result.Failed++
                                }
                            }
                        }
                    }
                    catch {
                        continue
                    }
                }
            }
        }
        catch {
            Write-Log "Error scanning UserData at $userDataRoot : $($_.Exception.Message)" "WARNING"
        }
    }

    return $result
}

function Remove-ChromeWMIEntries {
    <#
    .SYNOPSIS
        Removes Chrome-related WMI event subscriptions (filters, consumers, bindings).
        Does NOT use Win32_Product.Uninstall (that was handled in Phase 3).
    .OUTPUTS
        [PSCustomObject] with counts of removed, skipped, and failed items.
    #>
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        Removed = 0
        Skipped = 0
        Failed  = 0
    }

    # --- WMI Event Subscriptions ---
    $wmiSubscriptionClasses = @(
        "__EventFilter"
        "CommandLineEventConsumer"
        "ActiveScriptEventConsumer"
        "__FilterToConsumerBinding"
    )

    Write-Log "Scanning WMI event subscriptions for Chrome references..." "VERBOSE"

    foreach ($className in $wmiSubscriptionClasses) {
        try {
            $instances = Get-CimInstance -Namespace "root\subscription" -ClassName $className -ErrorAction SilentlyContinue

            if (-not $instances) { continue }

            foreach ($instance in $instances) {
                $isChromeRelated = $false
                $instanceDescription = "$className : $($instance.Name)"

                # Check all string properties for Chrome/Google references
                $instance.CimInstanceProperties | ForEach-Object {
                    if ($_.Value -is [string] -and
                        ($_.Value -match "Chrome" -or $_.Value -match "Google" -or
                        $_.Value -match "chrome\.exe" -or $_.Value -match "GoogleUpdate")) {
                        $isChromeRelated = $true
                    }
                }

                if ($isChromeRelated) {
                    Write-Log "Chrome WMI subscription found: $instanceDescription" "INFO"

                    $removeSuccess = Invoke-Action -Description "Remove WMI instance: $instanceDescription" -Action {
                        $instance | Remove-CimInstance -ErrorAction Stop
                    } -Phase "Phase 8" -Item $instanceDescription

                    if ($removeSuccess) {
                        $result.Removed++
                    }
                    else {
                        $result.Failed++
                    }
                }
            }
        }
        catch {
            Write-Log "Error scanning WMI class $className : $($_.Exception.Message)" "VERBOSE"
        }
    }

    if ($result.Removed -eq 0 -and $result.Failed -eq 0) {
        Write-Log "No Chrome-related WMI event subscriptions found" "SKIP"
        $result.Skipped++
    }

    return $result
}

function Invoke-Phase8WMICleanup {
    <#
    .SYNOPSIS
        Orchestrates Phase 8: WMI and Windows Installer registration cleanup.
    .OUTPUTS
        [bool] True if phase completed. False on critical failure.
    #>
    [CmdletBinding()]
    param()

    Write-LogSection "PHASE 8: WMI & INSTALLER REGISTRATION CLEANUP"

    $phaseRemoved = 0
    $phaseSkipped = 0
    $phaseFailed = 0

    try {
        # --- 8.1 Windows Installer Registry Entries ---
        Write-LogSubSection "8.1 — Windows Installer Registry Entries"
        $installerResult = Remove-ChromeInstallerReg
        $phaseRemoved += $installerResult.Removed
        $phaseSkipped += $installerResult.Skipped
        $phaseFailed += $installerResult.Failed
        Write-Log ("Installer registry entries: {0} removed, {1} skipped, {2} failed" -f
            $installerResult.Removed, $installerResult.Skipped, $installerResult.Failed) "INFO"

        # --- 8.2 WMI Event Subscriptions ---
        Write-LogSubSection "8.2 — WMI Event Subscriptions"
        $wmiResult = Remove-ChromeWMIEntries
        $phaseRemoved += $wmiResult.Removed
        $phaseSkipped += $wmiResult.Skipped
        $phaseFailed += $wmiResult.Failed
        Write-Log ("WMI entries: {0} removed, {1} skipped, {2} failed" -f
            $wmiResult.Removed, $wmiResult.Skipped, $wmiResult.Failed) "INFO"

        # --- Phase Result ---
        Write-Log ("Phase 8 complete: {0} items removed, {1} items skipped, {2} items failed" -f
            $phaseRemoved, $phaseSkipped, $phaseFailed) "INFO"

        if ($phaseFailed -eq 0) {
            $script:phaseResults["Phase 8"] = "PASS"
        }
        elseif ($phaseRemoved -gt 0) {
            $script:phaseResults["Phase 8"] = "PARTIAL"
        }
        else {
            $script:phaseResults["Phase 8"] = "FAIL"
        }

        return $true
    }
    catch {
        Write-Log "Critical error during WMI cleanup: $($_.Exception.Message)" "ERROR"
        $null = $script:errorCollection.Add([PSCustomObject]@{
                Timestamp   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Phase       = "Phase 8"
                Item        = "WMI Cleanup Orchestration"
                Error       = $_.Exception.Message
                ErrorRecord = $_
            })
        $script:phaseResults["Phase 8"] = "FAIL"
        return $false
    }
}

#endregion

#region PHASE 9 — FIREWALL RULE CLEANUP
# ============================================================================
# PHASE 9 — FIREWALL RULE CLEANUP
# ============================================================================

function Remove-ChromeFirewallRules {
    <#
    .SYNOPSIS
        Removes Windows Firewall rules that reference Chrome or Google Update
        using the NetSecurity PowerShell module. Falls back to netsh if needed.
    .OUTPUTS
        [PSCustomObject] with counts of removed, skipped, and failed rules.
    #>
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        Removed = 0
        Skipped = 0
        Failed  = 0
    }

    # --- Method 1: PowerShell NetSecurity cmdlets ---
    Write-Log "Scanning firewall rules via NetSecurity module..." "VERBOSE"

    try {
        $allRules = Get-NetFirewallRule -ErrorAction SilentlyContinue

        if (-not $allRules) {
            Write-Log "No firewall rules found or unable to enumerate" "SKIP"
            $result.Skipped++
            return $result
        }

        $chromeRules = @()

        foreach ($rule in $allRules) {
            try {
                $appFilter = $rule | Get-NetFirewallApplicationFilter -ErrorAction SilentlyContinue
                if ($appFilter -and $appFilter.Program) {
                    if ($appFilter.Program -match "chrome\.exe" -or
                        $appFilter.Program -match "Google\\Chrome" -or
                        $appFilter.Program -match "GoogleUpdate\.exe" -or
                        $appFilter.Program -match "Google\\Update" -or
                        $appFilter.Program -match "GoogleCrashHandler") {
                        $chromeRules += [PSCustomObject]@{
                            Rule    = $rule
                            Program = $appFilter.Program
                        }
                    }
                }
            }
            catch {
                continue
            }
        }

        if ($chromeRules.Count -eq 0) {
            Write-Log "No Chrome-related firewall rules found" "SKIP"
            $result.Skipped++
            return $result
        }

        Write-Log "Found $($chromeRules.Count) Chrome-related firewall rule(s)" "INFO"

        foreach ($chromeRule in $chromeRules) {
            $ruleName = $chromeRule.Rule.Name
            $ruleDisplayName = $chromeRule.Rule.DisplayName
            $ruleDirection = $chromeRule.Rule.Direction
            $ruleProgram = $chromeRule.Program

            $ruleDescription = "$ruleDisplayName | Direction: $ruleDirection | Program: $ruleProgram"
            Write-Log "Firewall rule: $ruleDescription" "INFO"

            $removeSuccess = Invoke-Action -Description "Remove firewall rule: $ruleDescription" -Action {
                Remove-NetFirewallRule -Name $ruleName -ErrorAction Stop
            } -Phase "Phase 9" -Item $ruleDescription

            if ($removeSuccess) {
                $result.Removed++
            }
            else {
                $result.Failed++
            }
        }
    }
    catch {
        Write-Log "Error using NetSecurity module: $($_.Exception.Message)" "WARNING"
        Write-Log "Falling back to netsh.exe..." "WARNING"

        # --- Method 2: netsh fallback ---
        try {
            $netshOutput = & netsh.exe advfirewall firewall show rule name=all verbose 2>&1

            if ($netshOutput) {
                $netshText = $netshOutput -join "`n"
                # Parse for Chrome-related rules
                $ruleBlocks = $netshText -split "(?=Rule Name:)"

                foreach ($block in $ruleBlocks) {
                    if ($block -match "chrome\.exe" -or $block -match "Google\\Chrome" -or
                        $block -match "GoogleUpdate" -or $block -match "Google\\Update") {

                        # Extract rule name
                        if ($block -match "Rule Name:\s+(.+)") {
                            $netshRuleName = $Matches[1].Trim()
                            Write-Log "Chrome firewall rule found via netsh: $netshRuleName" "INFO"

                            $removeSuccess = Invoke-Action -Description "netsh delete firewall rule: $netshRuleName" -Action {
                                $deleteOutput = & netsh.exe advfirewall firewall delete rule name="$netshRuleName" 2>&1
                                if ($LASTEXITCODE -ne 0) {
                                    throw "netsh delete failed: $deleteOutput"
                                }
                            } -Phase "Phase 9" -Item "netsh: $netshRuleName"

                            if ($removeSuccess) {
                                $result.Removed++
                            }
                            else {
                                $result.Failed++
                            }
                        }
                    }
                }
            }
        }
        catch {
            Write-Log "netsh fallback also failed: $($_.Exception.Message)" "ERROR"
            $result.Failed++
        }
    }

    return $result
}

function Invoke-Phase9FirewallCleanup {
    <#
    .SYNOPSIS
        Orchestrates Phase 9: removes all Chrome-related Windows Firewall rules.
    .OUTPUTS
        [bool] True if phase completed. False on critical failure.
    #>
    [CmdletBinding()]
    param()

    Write-LogSection "PHASE 9: FIREWALL RULE CLEANUP"

    try {
        # --- Pre-check from Phase 1 ---
        if ($script:chromeFirewallRules.Count -eq 0) {
            Write-Log "No Chrome firewall rules were detected in Phase 1 — performing live scan" "VERBOSE"
        }
        else {
            Write-Log "Phase 1 detected $($script:chromeFirewallRules.Count) Chrome-related firewall rule(s)" "INFO"
        }

        # --- Execute firewall rule removal ---
        Write-LogSubSection "9.1 — Removing Chrome Firewall Rules"
        $removalResult = Remove-ChromeFirewallRules

        # --- Phase Result ---
        Write-Log ("Phase 9 complete: {0} rules removed, {1} skipped, {2} failed" -f
            $removalResult.Removed, $removalResult.Skipped, $removalResult.Failed) "INFO"

        if ($removalResult.Failed -eq 0) {
            $script:phaseResults["Phase 9"] = "PASS"
        }
        elseif ($removalResult.Removed -gt 0) {
            $script:phaseResults["Phase 9"] = "PARTIAL"
        }
        else {
            $script:phaseResults["Phase 9"] = "FAIL"
        }

        return $true
    }
    catch {
        Write-Log "Critical error during firewall cleanup: $($_.Exception.Message)" "ERROR"
        $null = $script:errorCollection.Add([PSCustomObject]@{
                Timestamp   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Phase       = "Phase 9"
                Item        = "Firewall Cleanup Orchestration"
                Error       = $_.Exception.Message
                ErrorRecord = $_
            })
        $script:phaseResults["Phase 9"] = "FAIL"
        return $false
    }
}

#endregion

#region PHASE 10 — POST-REMOVAL VALIDATION
# ============================================================================
# PHASE 10 — POST-REMOVAL VALIDATION
# ============================================================================

function Confirm-ChromeRemoval {
    <#
    .SYNOPSIS
        Performs comprehensive post-removal validation checks.
        Each check produces a PASS/FAIL result.
    .OUTPUTS
        [ordered hashtable] of validation check names and their PASS/FAIL results.
    #>
    [CmdletBinding()]
    param()

    $validations = [ordered]@{}

    # --- 10.1 Process Check ---
    Write-Log "Validation: Checking for running Chrome processes..." "VERBOSE"
    $chromeProcs = @()
    foreach ($processName in $CHROME_PROCESS_NAMES) {
        $procs = Get-Process -Name $processName -ErrorAction SilentlyContinue
        if ($procs) { $chromeProcs += $procs }
    }

    if ($chromeProcs.Count -eq 0) {
        $validations["Processes"] = "PASS"
        Write-Log "  Processes: PASS — No Chrome processes running" "SUCCESS"
    }
    else {
        $validations["Processes"] = "FAIL"
        Write-Log "  Processes: FAIL — $($chromeProcs.Count) Chrome process(es) still running" "ERROR"
        foreach ($proc in $chromeProcs) {
            Write-Log "    Still running: $($proc.Name) (PID: $($proc.Id))" "ERROR"
        }
    }

    # --- 10.2 Service Check ---
    Write-Log "Validation: Checking for Chrome services..." "VERBOSE"
    $remainingServices = @()
    foreach ($serviceName in $CHROME_SERVICE_NAMES) {
        $svc = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
        if ($svc) { $remainingServices += $svc }
    }

    if ($remainingServices.Count -eq 0) {
        $validations["Services"] = "PASS"
        Write-Log "  Services: PASS — No Chrome services found" "SUCCESS"
    }
    else {
        $validations["Services"] = "FAIL"
        Write-Log "  Services: FAIL — $($remainingServices.Count) Chrome service(s) still exist" "ERROR"
        foreach ($svc in $remainingServices) {
            Write-Log "    Still exists: $($svc.Name) ($($svc.Status))" "ERROR"
        }
    }

    # --- 10.3 Scheduled Task Check ---
    Write-Log "Validation: Checking for Chrome scheduled tasks..." "VERBOSE"
    $remainingTasks = @()
    try {
        $allTasks = Get-ScheduledTask -ErrorAction SilentlyContinue
        if ($allTasks) {
            foreach ($pattern in $CHROME_TASK_PATTERNS) {
                $matching = $allTasks | Where-Object { $_.TaskName -like $pattern }
                if ($matching) { $remainingTasks += $matching }
            }
        }
    }
    catch {
        Write-Log "  Could not enumerate scheduled tasks for validation" "WARNING"
    }

    if ($remainingTasks.Count -eq 0) {
        $validations["Scheduled Tasks"] = "PASS"
        Write-Log "  Scheduled Tasks: PASS — No Chrome tasks found" "SUCCESS"
    }
    else {
        $validations["Scheduled Tasks"] = "FAIL"
        Write-Log "  Scheduled Tasks: FAIL — $($remainingTasks.Count) Chrome task(s) still exist" "ERROR"
        foreach ($task in $remainingTasks) {
            Write-Log "    Still exists: $($task.TaskName)" "ERROR"
        }
    }

    # --- 10.4 File System Check ---
    Write-Log "Validation: Checking key file system paths..." "VERBOSE"
    $criticalPaths = @(
        "$env:ProgramFiles\Google\Chrome"
        "${env:ProgramFiles(x86)}\Google\Chrome"
        "$env:ProgramData\Google\Chrome"
    )

    $remainingPaths = @()
    foreach ($path in $criticalPaths) {
        if (Test-PathSafe -Path $path) {
            $remainingPaths += $path
        }
    }

    # Spot-check 2 user profiles
    $spotCheckProfiles = $script:userProfiles | Select-Object -First 2
    foreach ($profile in $spotCheckProfiles) {
        if (-not $profile.ProfileExists) { continue }
        $userChromePath = Join-Path -Path $profile.ProfilePath -ChildPath "AppData\Local\Google\Chrome"
        if (Test-PathSafe -Path $userChromePath) {
            $remainingPaths += $userChromePath
        }
    }

    if ($remainingPaths.Count -eq 0) {
        $validations["File System"] = "PASS"
        Write-Log "  File System: PASS — No Chrome directories found in key locations" "SUCCESS"
    }
    else {
        $validations["File System"] = "FAIL"
        Write-Log "  File System: FAIL — $($remainingPaths.Count) Chrome path(s) still exist" "ERROR"
        foreach ($path in $remainingPaths) {
            Write-Log "    Still exists: $path" "ERROR"
        }
    }

    # --- 10.5 Registry Check ---
    Write-Log "Validation: Checking key registry paths..." "VERBOSE"
    $criticalRegKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome"
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome"
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe"
        "HKLM:\SOFTWARE\Classes\ChromeHTML"
        "HKLM:\SOFTWARE\Clients\StartMenuInternet\Google Chrome"
        "HKLM:\SOFTWARE\Google\Chrome"
    )

    $remainingRegKeys = @()
    foreach ($regKey in $criticalRegKeys) {
        if (Test-PathSafe -Path $regKey) {
            $remainingRegKeys += $regKey
        }
    }

    if ($remainingRegKeys.Count -eq 0) {
        $validations["Registry"] = "PASS"
        Write-Log "  Registry: PASS — No Chrome registry keys found in critical locations" "SUCCESS"
    }
    else {
        $validations["Registry"] = "FAIL"
        Write-Log "  Registry: FAIL — $($remainingRegKeys.Count) Chrome registry key(s) still exist" "ERROR"
        foreach ($regKey in $remainingRegKeys) {
            Write-Log "    Still exists: $regKey" "ERROR"
        }
    }

    # --- 10.6 Firewall Check ---
    Write-Log "Validation: Checking for Chrome firewall rules..." "VERBOSE"
    $remainingFWRules = @()
    try {
        $fwRules = Get-NetFirewallRule -ErrorAction SilentlyContinue
        if ($fwRules) {
            foreach ($rule in $fwRules) {
                try {
                    $appFilter = $rule | Get-NetFirewallApplicationFilter -ErrorAction SilentlyContinue
                    if ($appFilter -and $appFilter.Program -and
                        ($appFilter.Program -match "chrome\.exe" -or $appFilter.Program -match "Google\\Chrome")) {
                        $remainingFWRules += $rule
                    }
                }
                catch { continue }
            }
        }
    }
    catch {
        Write-Log "  Could not enumerate firewall rules for validation" "WARNING"
    }

    if ($remainingFWRules.Count -eq 0) {
        $validations["Firewall Rules"] = "PASS"
        Write-Log "  Firewall Rules: PASS — No Chrome firewall rules found" "SUCCESS"
    }
    else {
        $validations["Firewall Rules"] = "FAIL"
        Write-Log "  Firewall Rules: FAIL — $($remainingFWRules.Count) Chrome rule(s) still exist" "ERROR"
        foreach ($rule in $remainingFWRules) {
            Write-Log "    Still exists: $($rule.DisplayName)" "ERROR"
        }
    }

    # --- 10.7 Chrome Executable Check (definitive test) ---
    Write-Log "Validation: Checking for chrome.exe binary..." "VERBOSE"
    $chromeExePaths = @(
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
    )

    $exeFound = $false
    foreach ($exePath in $chromeExePaths) {
        if (Test-PathSafe -Path $exePath) {
            $exeFound = $true
            Write-Log "    Still exists: $exePath" "ERROR"
        }
    }

    if (-not $exeFound) {
        $validations["Chrome Binary"] = "PASS"
        Write-Log "  Chrome Binary: PASS — chrome.exe not found" "SUCCESS"
    }
    else {
        $validations["Chrome Binary"] = "FAIL"
        Write-Log "  Chrome Binary: FAIL — chrome.exe still exists on disk" "ERROR"
    }

    return $validations
}

function Write-ValidationSummary {
    <#
    .SYNOPSIS
        Writes a formatted validation summary table to the log.
    .PARAMETER ValidationResults
        Ordered hashtable of check name → PASS/FAIL.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.Specialized.OrderedDictionary]$ValidationResults
    )

    Write-LogSubSection "VALIDATION SUMMARY TABLE"

    $passCount = 0
    $failCount = 0

    # Table header
    Write-Log ("  {0,-25} {1,-10}" -f "Validation Check", "Result") "INFO"
    Write-Log ("  {0,-25} {1,-10}" -f ("-" * 25), ("-" * 10)) "INFO"

    foreach ($check in $ValidationResults.GetEnumerator()) {
        $checkName = $check.Key
        $checkResult = $check.Value

        if ($checkResult -eq "PASS") {
            Write-Log ("  {0,-25} {1,-10}" -f $checkName, $checkResult) "SUCCESS"
            $passCount++
        }
        else {
            Write-Log ("  {0,-25} {1,-10}" -f $checkName, $checkResult) "ERROR"
            $failCount++
        }
    }

    Write-Log ("  {0,-25} {1,-10}" -f ("-" * 25), ("-" * 10)) "INFO"
    Write-Log ("  {0,-25} {1} PASS, {2} FAIL" -f "TOTAL", $passCount, $failCount) "INFO"

    return @{ PassCount = $passCount; FailCount = $failCount }
}

function Invoke-Phase10Validation {
    <#
    .SYNOPSIS
        Orchestrates Phase 10: post-removal validation.
    .OUTPUTS
        [bool] True if phase completed.
    #>
    [CmdletBinding()]
    param()

    Write-LogSection "PHASE 10: POST-REMOVAL VALIDATION"

    try {
        if ($DryRun) {
            Write-Log "DryRun mode — validation will reflect pre-removal state (no changes were made)" "DRYRUN"
        }

        # --- Run all validation checks ---
        Write-LogSubSection "10.1–10.7 — Running Validation Checks"
        $script:validationResults = Confirm-ChromeRemoval

        # --- Write summary table ---
        $summary = Write-ValidationSummary -ValidationResults $script:validationResults

        # --- Determine exit code ---
        Write-LogSubSection "10.8 — Determining Exit Code"

        if ($DryRun) {
            Write-Log "DryRun mode — exit code set to 0 (no changes were attempted)" "DRYRUN"
            $script:exitCode = 0
            $script:phaseResults["Phase 10"] = "PASS"
        }
        elseif ($summary.FailCount -eq 0) {
            Write-Log "All validation checks PASSED — exit code: 0 (Full Success)" "SUCCESS"
            $script:exitCode = 0
            $script:phaseResults["Phase 10"] = "PASS"
        }
        elseif ($summary.PassCount -gt 0) {
            # Check if Chrome binary is still present (critical failure)
            if ($script:validationResults["Chrome Binary"] -eq "FAIL") {
                Write-Log "Chrome binary still exists — exit code: 2 (Critical Failure)" "ERROR"
                $script:exitCode = 2
                $script:phaseResults["Phase 10"] = "FAIL"
            }
            else {
                Write-Log "Some validation checks failed but Chrome binary is removed — exit code: 3010 (Partial Success / Reboot Pending)" "WARNING"
                $script:exitCode = 3010
                $script:phaseResults["Phase 10"] = "PARTIAL"
            }
        }
        else {
            Write-Log "All validation checks FAILED — exit code: 2 (Critical Failure)" "ERROR"
            $script:exitCode = 2
            $script:phaseResults["Phase 10"] = "FAIL"
        }

        return $true
    }
    catch {
        Write-Log "Error during validation: $($_.Exception.Message)" "ERROR"
        $script:exitCode = 3010
        $script:phaseResults["Phase 10"] = "FAIL"
        return $false
    }
}

#endregion

#region PHASE 11 — SUMMARY & EXIT
# ============================================================================
# PHASE 11 — SUMMARY & EXIT
# ============================================================================

function Write-ExecutionSummary {
    <#
    .SYNOPSIS
        Writes a comprehensive execution summary including phase results,
        item counts, timing, and recommendations.
    #>
    [CmdletBinding()]
    param()

    $executionEndTime = Get-Date
    $executionDuration = $executionEndTime - $SCRIPT_START_TIME
    $durationFormatted = "{0:D2}h {1:D2}m {2:D2}s" -f
    $executionDuration.Hours, $executionDuration.Minutes, $executionDuration.Seconds

    Write-LogSection "EXECUTION SUMMARY"

    # --- Execution Details ---
    Write-Log "" "INFO"
    Write-Log "  Script                : $SCRIPT_NAME v$SCRIPT_VERSION" "INFO"
    Write-Log "  Hostname              : $env:COMPUTERNAME" "INFO"
    Write-Log "  Start Time            : $($SCRIPT_START_TIME.ToString('yyyy-MM-dd HH:mm:ss'))" "INFO"
    Write-Log "  End Time              : $($executionEndTime.ToString('yyyy-MM-dd HH:mm:ss'))" "INFO"
    Write-Log "  Duration              : $durationFormatted" "INFO"
    Write-Log "  DryRun Mode           : $DryRun" "INFO"
    Write-Log "  Installation Detected : $($script:installationType)" "INFO"
    Write-Log "" "INFO"

    # --- Item Counts ---
    Write-LogSubSection "Item Counts"
    Write-Log "  Total Items Processed : $($script:totalItemsProcessed)" "INFO"
    Write-Log "  Successfully Removed  : $($script:totalItemsRemoved)" "SUCCESS"
    Write-Log "  Skipped (Not Found)   : $($script:totalItemsSkipped)" "INFO"
    Write-Log "  Failed                : $($script:totalItemsFailed)" $(if ($script:totalItemsFailed -gt 0) { "ERROR" } else { "INFO" })
    Write-Log "" "INFO"

    # --- Phase Results ---
    Write-LogSubSection "Phase Results"
    Write-Log ("  {0,-20} {1,-10}" -f "Phase", "Result") "INFO"
    Write-Log ("  {0,-20} {1,-10}" -f ("-" * 20), ("-" * 10)) "INFO"

    foreach ($phase in $script:phaseResults.GetEnumerator()) {
        $phaseCategory = switch ($phase.Value) {
            "PASS" { "SUCCESS" }
            "PARTIAL" { "WARNING" }
            "FAIL" { "ERROR" }
            default { "INFO" }
        }
        Write-Log ("  {0,-20} {1,-10}" -f $phase.Key, $phase.Value) $phaseCategory
    }

    Write-Log "" "INFO"

    # --- Exit Code ---
    $exitDescription = switch ($script:exitCode) {
        0 { "Full Success — All Chrome artifacts removed and validated" }
        1 { "Partial Success — Some artifacts could not be removed (see errors below)" }
        2 { "Critical Failure — Chrome may still be functional on this machine" }
    }

    $exitCategory = switch ($script:exitCode) {
        0 { "SUCCESS" }
        1 { "WARNING" }
        2 { "ERROR" }
    }

    Write-Log "  Exit Code             : $($script:exitCode) — $exitDescription" $exitCategory
    Write-Log "" "INFO"
}

function Write-ErrorSummary {
    <#
    .SYNOPSIS
        Writes a formatted error summary table if any errors occurred.
    #>
    [CmdletBinding()]
    param()

    if ($script:errorCollection.Count -eq 0) {
        Write-Log "No errors were recorded during execution" "SUCCESS"
        return
    }

    Write-LogSubSection "ERROR DETAILS"

    Write-Log "Total errors: $($script:errorCollection.Count)" "ERROR"
    Write-Log "" "INFO"

    Write-Log ("  {0,-5} {1,-12} {2,-40} {3}" -f "#", "Phase", "Item", "Error") "INFO"
    Write-Log ("  {0,-5} {1,-12} {2,-40} {3}" -f ("-" * 5), ("-" * 12), ("-" * 40), ("-" * 40)) "INFO"

    $errorIndex = 1
    foreach ($err in $script:errorCollection) {
        # Truncate long strings for table formatting
        $itemTruncated = if ($err.Item.Length -gt 38) { $err.Item.Substring(0, 38) + ".." } else { $err.Item }
        $errorTruncated = if ($err.Error.Length -gt 38) { $err.Error.Substring(0, 38) + ".." } else { $err.Error }

        Write-Log ("  {0,-5} {1,-12} {2,-40} {3}" -f
            $errorIndex, $err.Phase, $itemTruncated, $errorTruncated) "ERROR"

        $errorIndex++
    }

    Write-Log "" "INFO"
    Write-Log "Full error details are available in the log file: $LOG_FILE_PATH" "INFO"

    # Also write full error details (untruncated) to log file only
    Write-Log "" "VERBOSE"
    Write-Log "=== FULL ERROR DETAILS (VERBOSE) ===" "VERBOSE"
    $errorIndex = 1
    foreach ($err in $script:errorCollection) {
        Write-Log "Error #$errorIndex :" "VERBOSE"
        Write-Log "  Timestamp : $($err.Timestamp)" "VERBOSE"
        Write-Log "  Phase     : $($err.Phase)" "VERBOSE"
        Write-Log "  Item      : $($err.Item)" "VERBOSE"
        Write-Log "  Error     : $($err.Error)" "VERBOSE"
        if ($err.ErrorRecord) {
            Write-Log "  Exception : $($err.ErrorRecord.Exception.GetType().FullName)" "VERBOSE"
            Write-Log "  StackTrace: $($err.ErrorRecord.ScriptStackTrace)" "VERBOSE"
        }
        Write-Log "" "VERBOSE"
        $errorIndex++
    }
}

function Invoke-Phase11Summary {
    <#
    .SYNOPSIS
        Orchestrates Phase 11: final summary, error report, reboot recommendation, and exit.
    #>
    [CmdletBinding()]
    param()

    Write-LogSection "PHASE 11: SUMMARY & EXIT"

    # --- 11.1 Execution Summary ---
    Write-ExecutionSummary

    # --- 11.2 Error Summary ---
    Write-ErrorSummary

    # --- 11.3 Reboot Recommendation ---
    Write-LogSubSection "REBOOT RECOMMENDATION"

    Write-Log "" "INFO"
    Write-Log "  ╔══════════════════════════════════════════════════════════════════╗" "WARNING"
    Write-Log "  ║                                                                  ║" "WARNING"
    Write-Log "  ║   A system reboot is RECOMMENDED to complete the cleanup.        ║" "WARNING"
    Write-Log "  ║                                                                  ║" "WARNING"
    Write-Log "  ║   Some file locks, service registrations, and registry entries   ║" "WARNING"
    Write-Log "  ║   may only be fully released after a reboot.                     ║" "WARNING"
    Write-Log "  ║                                                                  ║" "WARNING"
    Write-Log "  ║   This is especially important if:                               ║" "WARNING"
    Write-Log "  ║   - Any offline registry hives failed to unload                  ║" "WARNING"
    Write-Log "  ║   - Any services were marked for deletion                        ║" "WARNING"
    Write-Log "  ║   - Any files were locked during removal                         ║" "WARNING"
    Write-Log "  ║                                                                  ║" "WARNING"
    Write-Log "  ╚══════════════════════════════════════════════════════════════════╝" "WARNING"
    Write-Log "" "INFO"

    # --- SCCM-Specific Messaging ---
    Write-Log "SCCM Task Sequence Note: This script returns exit code $($script:exitCode)." "INFO"
    Write-Log "  0 = Full Success | 3010 = Partial Success (Reboot Pending) | 2 = Critical Failure" "INFO"
    Write-Log "  Exit code 3010 signals SCCM that a reboot is required to finalize cleanup." "INFO"
    Write-Log "  Configure your task sequence step to treat exit code 3010 as success." "INFO"
    Write-Log "" "INFO"

    # --- Log file location reminder ---
    Write-Log "Log files:" "INFO"
    Write-Log "  Main log    : $LOG_FILE_PATH" "INFO"
    Write-Log "  Transcript  : $TRANSCRIPT_FILE_PATH" "INFO"
    Write-Log "" "INFO"

    $script:phaseResults["Phase 11"] = "PASS"
}

#endregion

#region MAIN EXECUTION
# ============================================================================
# MAIN EXECUTION — Orchestrates all phases in sequence
# ============================================================================

# ============================================================
# PHASE 0: PRE-FLIGHT CHECKS
# ============================================================
$preFlightResult = Invoke-Phase0PreFlight

if (-not $preFlightResult) {
    # Pre-flight failed — cannot continue
    Write-Host "[CRITICAL] Pre-flight checks failed. Script cannot continue." -ForegroundColor Red

    # Attempt to stop transcript if it was started
    try { Stop-Transcript -ErrorAction SilentlyContinue } catch { }

    exit 2
}

# --- Confirmation Prompt (unless -Force is specified) ---
if (-not $Force -and -not $DryRun -and [Environment]::UserInteractive) {
    Write-Log "Awaiting user confirmation (use -Force to suppress this prompt)..." "INFO"
    $confirmation = Read-Host "This will PERMANENTLY remove Google Chrome and all artifacts. Type 'YES' to continue"
    if ($confirmation -ne 'YES') {
        Write-Log "User cancelled execution" "WARNING"
        try { Stop-Transcript -ErrorAction SilentlyContinue } catch { }
        exit 0
    }
    Write-Log "User confirmed — proceeding with removal" "SUCCESS"
}

# ============================================================
# PHASE 1: DETECTION & INVENTORY
# ============================================================
$detectionResult = Invoke-Phase1Detection

if (-not $detectionResult) {
    Write-Log "Detection phase encountered a critical error — continuing with best effort cleanup" "WARNING"
}

# ============================================================
# PHASE 2: PROCESS TERMINATION
# ============================================================
$processResult = Invoke-Phase2ProcessTermination

if (-not $processResult) {
    Write-Log "Process termination encountered errors — continuing with cleanup" "WARNING"
}

# ============================================================
# PHASE 3: GRACEFUL UNINSTALL
# ============================================================
$uninstallResult = Invoke-Phase3GracefulUninstall

if (-not $uninstallResult) {
    Write-Log "Graceful uninstall encountered errors — continuing with cleanup" "WARNING"
}

# ============================================================
# PHASE 4: SERVICE REMOVAL
# ============================================================
$serviceResult = Invoke-Phase4ServiceRemoval

if (-not $serviceResult) {
    Write-Log "Service removal encountered errors — continuing with cleanup" "WARNING"
}

# ============================================================
# PHASE 5: SCHEDULED TASK REMOVAL
# ============================================================
$taskResult = Invoke-Phase5ScheduledTaskRemoval

if (-not $taskResult) {
    Write-Log "Scheduled task removal encountered errors — continuing with brute-force cleanup" "WARNING"
}

# ============================================================
# PHASE 6: FILE SYSTEM CLEANUP
# ============================================================
$fileResult = Invoke-Phase6FileSystemCleanup

if (-not $fileResult) {
    Write-Log "File system cleanup encountered errors — continuing with registry cleanup" "WARNING"
}

# ============================================================
# PHASE 7: REGISTRY CLEANUP
# ============================================================
$registryResult = Invoke-Phase7RegistryCleanup

if (-not $registryResult) {
    Write-Log "Registry cleanup encountered errors — continuing with WMI cleanup" "WARNING"
}

# ============================================================
# PHASE 8: WMI & INSTALLER CLEANUP
# ============================================================
$wmiResult = Invoke-Phase8WMICleanup

if (-not $wmiResult) {
    Write-Log "WMI cleanup encountered errors — continuing with firewall cleanup" "WARNING"
}

# ============================================================
# PHASE 9: FIREWALL RULE CLEANUP
# ============================================================
$firewallResult = Invoke-Phase9FirewallCleanup

if (-not $firewallResult) {
    Write-Log "Firewall cleanup encountered errors — continuing with validation" "WARNING"
}

# ============================================================
# PHASE 10: POST-REMOVAL VALIDATION
# ============================================================
$validationResult = Invoke-Phase10Validation

# ============================================================
# PHASE 11: SUMMARY & EXIT
# ============================================================
Invoke-Phase11Summary

# ============================================================
# STOP TRANSCRIPT & EXIT
# ============================================================
Write-Log "Script execution complete. Exiting with code: $($script:exitCode)" "INFO"

try {
    Stop-Transcript -ErrorAction SilentlyContinue
}
catch {
    # Transcript may not have been started successfully
}

exit $script:exitCode

#endregion
