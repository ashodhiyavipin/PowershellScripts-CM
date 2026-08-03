<#
.SYNOPSIS
    Core module for modular application uninstaller system.

.DESCRIPTION
    This module provides the central dispatcher, logging, and shared utilities for 
    the modular application uninstaller system. It contains:
    - CMTrace-compatible logging function
    - Central pattern-to-function mapping table
    - Uninstall-Application dispatcher
    - Uninstall-RegistryApp fallback function

.NOTES
    File Name:    Uninstall.Core.psm1
    Author:       V.Ashodhiya
    Version:      1.0
    Date:         2026-01-27
#>

#region Global Variables

# Global log file path and app name for logging
$Global:CurrentAppName = $null
$Global:CurrentLogFile = $null
$Global:UninstallStartTime = $null

# Central mapping table: regex pattern (case-insensitive) -> function name
$Global:AppMapping = @{
    '^keepass'            = 'Uninstall-KeePass'
    'adobe.*reader'       = 'Uninstall-AdobeReader'
    'adobe.*acrobat'      = 'Uninstall-AdobeReader'
    'dell.*supportassist' = 'Uninstall-DellSupportAssist'
    '^PlantronicsHub'     = 'Uninstall-PlantronicsHub'
    '^wireshark'          = 'Uninstall-Wireshark'
    # Add more patterns as needed
}

#endregion

#region Logging Functions

<#
.SYNOPSIS
    Writes a log entry in CMTrace-compatible format.

.DESCRIPTION
    Creates log entries that can be read by CMTrace.exe. The log format includes:
    - Message text
    - Timestamp with milliseconds
    - Component name
    - Severity level (1=Info, 2=Warning, 3=Error)
    - Thread ID
    - Source file name

.PARAMETER Message
    The message to log.

.PARAMETER Component
    The component or function name generating the log entry.

.PARAMETER Severity
    The severity level: 1=Information, 2=Warning, 3=Error. Default is 1.

.EXAMPLE
    Write-CMLog -Message "Starting uninstall" -Component "Uninstall-Application" -Severity 1
#>
function Write-CMLog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [string]$Component = "Uninstaller",
        
        [Parameter(Mandatory = $false)]
        [ValidateSet(1, 2, 3)]
        [int]$Severity = 1
    )
    
    # Ensure log directory exists
    if ($Global:CurrentLogFile) {
        $logDir = Split-Path -Parent $Global:CurrentLogFile
        if (-not (Test-Path $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }
        
        # Get current time in CMTrace format
        $time = Get-Date -Format "HH:mm:ss.fff"
        $date = Get-Date -Format "MM-dd-yyyy"
        
        # Get thread ID and construct file name
        $thread = [System.Threading.Thread]::CurrentThread.ManagedThreadId
        $file = Split-Path -Leaf $PSCommandPath
        if (-not $file) { $file = "PowerShell" }
        
        # Build CMTrace log line
        # Format: <![LOG[message]LOG]!><time="HH:mm:ss.fff+000" date="MM-dd-yyyy" component="Component" context="" type="1" thread="1234" file="file">
        $cmtraceLine = "<![LOG[$Message]LOG]!><time=`"$time+000`" date=`"$date`" component=`"$Component`" context=`"`" type=`"$Severity`" thread=`"$thread`" file=`"$file`">"
        
        # Write to log file (thread-safe)
        try {
            Add-Content -Path $Global:CurrentLogFile -Value $cmtraceLine -ErrorAction Stop
        }
        catch {
            Write-Warning "Failed to write to log file: $_"
        }
    }
    
    # Also write to console for real-time feedback
    switch ($Severity) {
        1 { Write-Host $Message }
        2 { Write-Warning $Message }
        3 { Write-Error $Message }
    }
}

#endregion

#region Registry-Based Uninstall

<#
.SYNOPSIS
    Generic registry-based application uninstaller.

.DESCRIPTION
    Searches the Windows registry for applications matching the specified name and 
    attempts to uninstall them using their registered uninstall commands. This function
    searches all three common registry paths (HKLM x64, HKLM x86, HKCU).

.PARAMETER AppName
    The name (or partial name) of the application to uninstall. Case-insensitive.

.OUTPUTS
    Returns a hashtable with:
    - Success: Boolean indicating overall success
    - ExitCode: 0 for success, 1 for failure
    - Found: Number of matching applications found
    - Uninstalled: Number successfully uninstalled
    - Failed: Number that failed to uninstall
    - Versions: Array of version strings that were handled

.EXAMPLE
    Uninstall-RegistryApp -AppName "Adobe Reader"
#>
function Uninstall-RegistryApp {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$AppName
    )
    
    Write-CMLog -Message "=== Starting registry-based uninstall for: $AppName ===" -Component "Uninstall-RegistryApp" -Severity 1
    
    $registryPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    $result = @{
        Success     = $false
        ExitCode    = 1
        Found       = 0
        Uninstalled = 0
        Failed      = 0
        Versions    = @()
    }
    
    foreach ($regPath in $registryPaths) {
        try {
            Write-CMLog -Message "Searching registry path: $regPath" -Component "Uninstall-RegistryApp" -Severity 1
            
            $apps = Get-ItemProperty $regPath -ErrorAction SilentlyContinue | 
            Where-Object { $_.DisplayName -like "*$AppName*" }
            
            foreach ($app in $apps) {
                $result.Found++
                
                $displayName = $app.DisplayName
                $displayVersion = $app.DisplayVersion
                $versionString = "$displayName ($displayVersion)"
                $result.Versions += $versionString
                
                Write-CMLog -Message "Found: $versionString" -Component "Uninstall-RegistryApp" -Severity 1
                
                # Determine uninstall command
                $uninstallCmd = $null
                if ($app.QuietUninstallString) {
                    $uninstallCmd = $app.QuietUninstallString
                    Write-CMLog -Message "Using QuietUninstallString: $uninstallCmd" -Component "Uninstall-RegistryApp" -Severity 1
                }
                elseif ($app.UninstallString) {
                    $uninstallCmd = $app.UninstallString
                    Write-CMLog -Message "Using UninstallString: $uninstallCmd" -Component "Uninstall-RegistryApp" -Severity 1
                }
                
                if ($uninstallCmd) {
                    # Check if msiexec.exe has /I and replace it with /X
                    if ($uninstallCmd -match "^MsiExec\.exe\s*/I\s*\{[A-Z0-9-]+\}") {
                        Write-CMLog -Message "Converting /I to /X for MSI uninstall" -Component "Uninstall-RegistryApp" -Severity 1
                        $uninstallCmd = $uninstallCmd.Replace("/I", "/X")
                    }
                    
                    # Add silent switches for MSI uninstalls
                    if ($uninstallCmd -like "MsiExec.exe*") {
                        $uninstallCmd += " /qn /norestart"
                        Write-CMLog -Message "Added silent switches: $uninstallCmd" -Component "Uninstall-RegistryApp" -Severity 1
                    }
                    
                    # Execute uninstall command
                    try {
                        Write-CMLog -Message "Executing: $uninstallCmd" -Component "Uninstall-RegistryApp" -Severity 1
                        
                        $process = Start-Process -FilePath "cmd.exe" `
                            -ArgumentList "/c", $uninstallCmd `
                            -Wait `
                            -WindowStyle Hidden `
                            -PassThru `
                            -ErrorAction Stop
                        
                        if ($process.ExitCode -eq 0) {
                            Write-CMLog -Message "Successfully uninstalled: $versionString" -Component "Uninstall-RegistryApp" -Severity 1
                            $result.Uninstalled++
                        }
                        else {
                            Write-CMLog -Message "Uninstall failed with exit code $($process.ExitCode): $versionString" -Component "Uninstall-RegistryApp" -Severity 3
                            $result.Failed++
                        }
                    }
                    catch {
                        Write-CMLog -Message "Exception during uninstall: $_" -Component "Uninstall-RegistryApp" -Severity 3
                        Write-CMLog -Message "Command: $uninstallCmd" -Component "Uninstall-RegistryApp" -Severity 3
                        $result.Failed++
                    }
                }
                else {
                    Write-CMLog -Message "No uninstall command found for: $versionString" -Component "Uninstall-RegistryApp" -Severity 2
                    $result.Failed++
                }
            }
        }
        catch {
            Write-CMLog -Message "Error reading registry path $regPath : $_" -Component "Uninstall-RegistryApp" -Severity 3
        }
    }
    
    # Determine overall success
    if ($result.Found -eq 0) {
        Write-CMLog -Message "No applications matching '$AppName' found in registry" -Component "Uninstall-RegistryApp" -Severity 2
        $result.Success = $false
        $result.ExitCode = 1
    }
    elseif ($result.Uninstalled -gt 0 -and $result.Failed -eq 0) {
        Write-CMLog -Message "All found applications uninstalled successfully" -Component "Uninstall-RegistryApp" -Severity 1
        $result.Success = $true
        $result.ExitCode = 0
    }
    elseif ($result.Uninstalled -gt 0 -and $result.Failed -gt 0) {
        Write-CMLog -Message "Partial success: $($result.Uninstalled) succeeded, $($result.Failed) failed" -Component "Uninstall-RegistryApp" -Severity 2
        $result.Success = $true  # Partial success still counts
        $result.ExitCode = 0
    }
    else {
        Write-CMLog -Message "All uninstall attempts failed" -Component "Uninstall-RegistryApp" -Severity 3
        $result.Success = $false
        $result.ExitCode = 1
    }
    
    Write-CMLog -Message "=== Registry uninstall complete ===" -Component "Uninstall-RegistryApp" -Severity 1
    
    return $result
}

#endregion

#region Central Dispatcher

<#
.SYNOPSIS
    Central dispatcher for application uninstallation.

.DESCRIPTION
    This is the main entry point for uninstalling applications. It performs the following:
    1. Normalizes input and initializes logging
    2. Matches the app name against the central pattern mapping table
    3. Calls the appropriate app-specific function or falls back to registry-based uninstall
    4. Handles errors and fallback logic
    5. Outputs a structured summary of the operation

.PARAMETER AppName
    The name of the application to uninstall. Can be a partial name or pattern.

.PARAMETER LogPath
    Optional. The directory path where logs should be written. 
    Default: C:\Windows\fndr\logs

.OUTPUTS
    Returns a structured hashtable with:
    - AppNameInput: Original input
    - MatchedPattern: The regex pattern that matched (or "None")
    - SelectedFunction: Name of the function called
    - VersionsHandled: Array of version strings
    - PrimaryMethod: 'AppSpecific' or 'RegistryFallback'
    - FallbackUsed: Boolean
    - FinalExitCode: 0=success, non-zero=failure
    - Status: 'Success' or 'Failure'

.EXAMPLE
    Uninstall-Application -AppName "KeePass"

.EXAMPLE
    Uninstall-Application -AppName "Adobe Reader" -LogPath "C:\Logs"
#>
function Uninstall-Application {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$AppName,
        
        [Parameter(Mandatory = $false)]
        [string]$LogPath = "C:\Windows\fndr\logs"
    )
    
    # Normalize input
    $AppName = $AppName.Trim()
    
    # Initialize logging
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $sanitizedAppName = $AppName -replace '[^\w\s-]', '_'
    $Global:CurrentAppName = $AppName
    $Global:CurrentLogFile = Join-Path $LogPath "ModularUninstall_$($sanitizedAppName)_$timestamp.log"
    $Global:UninstallStartTime = Get-Date
    
    Write-CMLog -Message "========================================" -Component "Uninstall-Application" -Severity 1
    Write-CMLog -Message "Modular Application Uninstaller v1.0" -Component "Uninstall-Application" -Severity 1
    Write-CMLog -Message "========================================" -Component "Uninstall-Application" -Severity 1
    Write-CMLog -Message "Start Time: $($Global:UninstallStartTime.ToString('yyyy-MM-dd HH:mm:ss'))" -Component "Uninstall-Application" -Severity 1
    Write-CMLog -Message "Requested Application: $AppName" -Component "Uninstall-Application" -Severity 1
    Write-CMLog -Message "Log File: $Global:CurrentLogFile" -Component "Uninstall-Application" -Severity 1
    
    # Initialize result structure
    $summary = @{
        AppNameInput     = $AppName
        MatchedPattern   = "None"
        SelectedFunction = $null
        VersionsHandled  = @()
        PrimaryMethod    = $null
        FallbackUsed     = $false
        FinalExitCode    = 1
        Status           = "Failure"
    }
    
    try {
        # Pattern matching
        $matchedFunction = $null
        $matchedPattern = $null
        
        Write-CMLog -Message "Searching pattern mapping table..." -Component "Uninstall-Application" -Severity 1
        
        foreach ($pattern in $Global:AppMapping.Keys) {
            if ($AppName -imatch $pattern) {
                $matchedPattern = $pattern
                $matchedFunction = $Global:AppMapping[$pattern]
                $summary.MatchedPattern = $pattern
                $summary.SelectedFunction = $matchedFunction
                
                Write-CMLog -Message "Pattern matched: '$pattern' -> $matchedFunction" -Component "Uninstall-Application" -Severity 1
                break
            }
        }
        
        # Execute app-specific function or fallback
        if ($matchedFunction) {
            Write-CMLog -Message "Calling app-specific function: $matchedFunction" -Component "Uninstall-Application" -Severity 1
            $summary.PrimaryMethod = "AppSpecific"
            
            try {
                # Check if function exists
                if (Get-Command $matchedFunction -ErrorAction SilentlyContinue) {
                    # Call the app-specific function
                    $appResult = & $matchedFunction
                    
                    # Check result
                    if ($appResult -and $appResult.ExitCode -eq 0) {
                        Write-CMLog -Message "App-specific function succeeded" -Component "Uninstall-Application" -Severity 1
                        $summary.VersionsHandled = $appResult.Versions
                        $summary.FinalExitCode = 0
                        $summary.Status = "Success"
                    }
                    else {
                        Write-CMLog -Message "App-specific function failed or returned non-zero exit code" -Component "Uninstall-Application" -Severity 2
                        Write-CMLog -Message "Attempting fallback to registry-based uninstall..." -Component "Uninstall-Application" -Severity 2
                        
                        $summary.FallbackUsed = $true
                        $fallbackResult = Uninstall-RegistryApp -AppName $AppName
                        
                        if ($fallbackResult.Success) {
                            Write-CMLog -Message "Fallback succeeded" -Component "Uninstall-Application" -Severity 1
                            $summary.VersionsHandled = $fallbackResult.Versions
                            $summary.FinalExitCode = 0
                            $summary.Status = "Success"
                        }
                        else {
                            Write-CMLog -Message "Fallback also failed" -Component "Uninstall-Application" -Severity 3
                            $summary.FinalExitCode = 1
                            $summary.Status = "Failure"
                        }
                    }
                }
                else {
                    Write-CMLog -Message "Function '$matchedFunction' not found! Falling back to registry uninstall." -Component "Uninstall-Application" -Severity 2
                    
                    $summary.FallbackUsed = $true
                    $summary.PrimaryMethod = "RegistryFallback"
                    $fallbackResult = Uninstall-RegistryApp -AppName $AppName
                    
                    if ($fallbackResult.Success) {
                        $summary.VersionsHandled = $fallbackResult.Versions
                        $summary.FinalExitCode = 0
                        $summary.Status = "Success"
                    }
                    else {
                        $summary.FinalExitCode = 1
                        $summary.Status = "Failure"
                    }
                }
            }
            catch {
                Write-CMLog -Message "Exception calling app-specific function: $_" -Component "Uninstall-Application" -Severity 3
                Write-CMLog -Message "Stack Trace: $($_.ScriptStackTrace)" -Component "Uninstall-Application" -Severity 3
                Write-CMLog -Message "Attempting fallback to registry-based uninstall..." -Component "Uninstall-Application" -Severity 2
                
                $summary.FallbackUsed = $true
                $fallbackResult = Uninstall-RegistryApp -AppName $AppName
                
                if ($fallbackResult.Success) {
                    Write-CMLog -Message "Fallback succeeded" -Component "Uninstall-Application" -Severity 1
                    $summary.VersionsHandled = $fallbackResult.Versions
                    $summary.FinalExitCode = 0
                    $summary.Status = "Success"
                }
                else {
                    Write-CMLog -Message "Fallback also failed" -Component "Uninstall-Application" -Severity 3
                    $summary.FinalExitCode = 1
                    $summary.Status = "Failure"
                }
            }
        }
        else {
            # No pattern match - go straight to registry fallback
            Write-CMLog -Message "No pattern match found. Using registry-based uninstall." -Component "Uninstall-Application" -Severity 1
            
            $summary.SelectedFunction = "Uninstall-RegistryApp"
            $summary.PrimaryMethod = "RegistryFallback"
            $summary.FallbackUsed = $false  # This IS the primary method
            
            $registryResult = Uninstall-RegistryApp -AppName $AppName
            
            if ($registryResult.Success) {
                $summary.VersionsHandled = $registryResult.Versions
                $summary.FinalExitCode = 0
                $summary.Status = "Success"
            }
            else {
                $summary.FinalExitCode = 1
                $summary.Status = "Failure"
            }
        }
        
    }
    catch {
        Write-CMLog -Message "Fatal error in Uninstall-Application: $_" -Component "Uninstall-Application" -Severity 3
        Write-CMLog -Message "Stack Trace: $($_.ScriptStackTrace)" -Component "Uninstall-Application" -Severity 3
        $summary.FinalExitCode = 1
        $summary.Status = "Failure"
    }
    
    # Final summary
    $endTime = Get-Date
    $duration = $endTime - $Global:UninstallStartTime
    
    Write-CMLog -Message "========================================" -Component "Uninstall-Application" -Severity 1
    Write-CMLog -Message "OPERATION SUMMARY" -Component "Uninstall-Application" -Severity 1
    Write-CMLog -Message "========================================" -Component "Uninstall-Application" -Severity 1
    Write-CMLog -Message "AppNameInput     : $($summary.AppNameInput)" -Component "Uninstall-Application" -Severity 1
    Write-CMLog -Message "MatchedPattern   : $($summary.MatchedPattern)" -Component "Uninstall-Application" -Severity 1
    Write-CMLog -Message "SelectedFunction : $($summary.SelectedFunction)" -Component "Uninstall-Application" -Severity 1
    Write-CMLog -Message "PrimaryMethod    : $($summary.PrimaryMethod)" -Component "Uninstall-Application" -Severity 1
    Write-CMLog -Message "FallbackUsed     : $($summary.FallbackUsed)" -Component "Uninstall-Application" -Severity 1
    Write-CMLog -Message "VersionsHandled  : $($summary.VersionsHandled.Count) version(s)" -Component "Uninstall-Application" -Severity 1
    
    foreach ($version in $summary.VersionsHandled) {
        Write-CMLog -Message "  - $version" -Component "Uninstall-Application" -Severity 1
    }
    
    Write-CMLog -Message "FinalExitCode    : $($summary.FinalExitCode)" -Component "Uninstall-Application" -Severity 1
    Write-CMLog -Message "Status           : $($summary.Status)" -Component "Uninstall-Application" -Severity 1
    Write-CMLog -Message "Duration         : $($duration.ToString('mm\:ss'))" -Component "Uninstall-Application" -Severity 1
    Write-CMLog -Message "End Time         : $($endTime.ToString('yyyy-MM-dd HH:mm:ss'))" -Component "Uninstall-Application" -Severity 1
    Write-CMLog -Message "========================================" -Component "Uninstall-Application" -Severity 1
    
    return $summary
}

#endregion

# Export module members
Export-ModuleMember -Function @(
    'Write-CMLog',
    'Uninstall-RegistryApp',
    'Uninstall-Application'
)
