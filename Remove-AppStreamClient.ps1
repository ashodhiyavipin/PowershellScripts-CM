#Requires -Version 5.1
<#
.SYNOPSIS
    Removes residual Amazon AppStream 2.0 client folders from all user profiles.

.DESCRIPTION
    Tenable Plugin 226455 detects leftover AppStreamClient installations in user
    profile AppData\Local directories. This script enumerates all user profiles
    and removes the AppStreamClient folder where found.

    Scope: ONLY <UserProfile>\AppData\Local\AppStreamClient
    No registry changes. No shortcut removal. No changes outside reported paths.

.NOTES
    Deployment  : SCCM Package/Program (run as SYSTEM, hidden)
    Command     : powershell.exe -ExecutionPolicy Bypass -NoProfile -File Remove-AppStreamClient.ps1
    Max Runtime : 15 minutes
    Reboot      : Not required
    Author      : Vipin Ashodhiya
    Version     : 1.2
    Date        : 2026-01-01

.CHANGELOG
    1.0 - Initial version
    1.1 - Applied code review fixes:
           - Fixed null Sum on empty folders causing terminating error
           - Fixed inaccurate .SI property logged as "User" (now "Session")
           - Replaced hardcoded C:\Users with Win32_UserProfile CIM query
           - Attribute reset now targets IsReadOnly instead of overwriting all attributes
           - Attribute reset now includes the parent AppStreamClient directory itself
    1.2 - Fixed counter logic:
           - Replaced fragile FoldersFailed++/-- pattern with $removed boolean flag
           - Counters now update exactly once per folder based on final outcome
           - Eliminates any possibility of negative counter values
#>

# ============================================================================
# CONFIGURATION
# ============================================================================

$TargetFolder = "AppData\Local\AppStreamClient"
$ProcessName = "AppStreamClient"
$LogPath = "C:\Windows\fndr\logs\AppStreamCleanup.log"

# ============================================================================
# FUNCTIONS
# ============================================================================

$logDir = Split-Path -Path $LogPath -Parent
if (-not (Test-Path -Path $logDir)) {
    New-Item -Path $logDir -ItemType Directory -Force | Out-Null
}

function Write-Log {
    <#
    .SYNOPSIS
        Writes a timestamped entry to the log file and console.
    #>
    param (
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$timestamp] [$Level] $Message"

    Add-Content -Path $LogPath -Value $entry -Encoding UTF8 -ErrorAction SilentlyContinue

    switch ($Level) {
        'ERROR' { Write-Host $entry -ForegroundColor Red }
        'WARN' { Write-Host $entry -ForegroundColor Yellow }
        'SUCCESS' { Write-Host $entry -ForegroundColor Green }
        default { Write-Host $entry }
    }
}

function Get-UserProfilePaths {
    <#
    .SYNOPSIS
        Returns an array of real user profile local paths by querying
        Win32_UserProfile. Filters out special/system profiles automatically.
    #>

    try {
        $profiles = Get-CimInstance -ClassName Win32_UserProfile -ErrorAction Stop |
        Where-Object { -not $_.Special } |
        Select-Object -ExpandProperty LocalPath

        if (-not $profiles) {
            Write-Log "Win32_UserProfile returned no non-special profiles." -Level 'WARN'
            return @()
        }

        # Filter to only profiles that actually exist on disk
        $validProfiles = @()
        foreach ($path in $profiles) {
            if (Test-Path -Path $path -PathType Container) {
                $validProfiles += $path
            }
            else {
                Write-Log "Profile path registered but not found on disk: $path" -Level 'WARN'
            }
        }

        return $validProfiles
    }
    catch {
        Write-Log "Failed to query Win32_UserProfile: $($_.Exception.Message)" -Level 'ERROR'
        return @()
    }
}

function Stop-AppStreamProcesses {
    <#
    .SYNOPSIS
        Terminates any running AppStreamClient processes as a safety measure.
    #>
    $processes = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue

    if ($processes) {
        foreach ($proc in $processes) {
            try {
                $proc | Stop-Process -Force -ErrorAction Stop
                Write-Log "Terminated process: $($proc.Name) (PID: $($proc.Id), Session: $($proc.SI))" -Level 'WARN'
            }
            catch {
                Write-Log "Failed to terminate process PID $($proc.Id): $($_.Exception.Message)" -Level 'ERROR'
            }
        }
        # Brief pause to release file locks
        Start-Sleep -Seconds 2
    }
    else {
        Write-Log "No running $ProcessName processes found (expected)." -Level 'INFO'
    }
}

function Remove-AppStreamFromProfiles {
    <#
    .SYNOPSIS
        Iterates through all user profiles and removes the AppStreamClient folder.
    .OUTPUTS
        Returns a hashtable with counters for SCCM exit code logic.
    #>

    $results = @{
        ProfilesScanned = 0
        FoldersFound    = 0
        FoldersRemoved  = 0
        FoldersFailed   = 0
        ProfilesClean   = 0
    }

    # Get real user profile paths via CIM
    $profilePaths = Get-UserProfilePaths

    if ($profilePaths.Count -eq 0) {
        Write-Log "No valid user profiles found. Nothing to do." -Level 'WARN'
        return $results
    }

    Write-Log "Found $($profilePaths.Count) user profile(s) to scan." -Level 'INFO'

    foreach ($profilePath in $profilePaths) {

        $results.ProfilesScanned++
        $profileName = Split-Path -Path $profilePath -Leaf
        $appStreamPath = Join-Path -Path $profilePath -ChildPath $TargetFolder

        if (Test-Path -Path $appStreamPath -PathType Container) {

            $results.FoldersFound++
            Write-Log "FOUND: $appStreamPath" -Level 'WARN'

            # Capture folder size for reporting (safe against empty folders)
            try {
                $folderSize = (Get-ChildItem -Path $appStreamPath -Recurse -Force -ErrorAction SilentlyContinue |
                    Measure-Object -Property Length -Sum).Sum

                if ($null -eq $folderSize) { $folderSize = 0 }

                $folderSizeMB = [math]::Round($folderSize / 1MB, 2)
                Write-Log "  Profile: $profileName | Folder size: $folderSizeMB MB" -Level 'INFO'
            }
            catch {
                Write-Log "  Profile: $profileName | Could not calculate folder size." -Level 'WARN'
            }

            # ============================================================
            # REMOVAL WITH RETRY - Boolean flag tracks final outcome
            # ============================================================
            $removed = $false

            # First attempt: straight removal
            try {
                Remove-Item -Path $appStreamPath -Recurse -Force -ErrorAction Stop
                $removed = $true
                Write-Log "  REMOVED successfully: $appStreamPath" -Level 'SUCCESS'
            }
            catch {
                Write-Log "  FAILED to remove: $appStreamPath - $($_.Exception.Message)" -Level 'ERROR'

                # Retry: clear ReadOnly on the folder itself AND its contents
                Write-Log "  Retrying with ReadOnly attribute reset..." -Level 'INFO'
                try {
                    # Reset the parent directory itself
                    $parentItem = Get-Item -Path $appStreamPath -Force -ErrorAction Stop
                    if ($parentItem.Attributes -band [System.IO.FileAttributes]::ReadOnly) {
                        $parentItem.IsReadOnly = $false
                        Write-Log "  Cleared ReadOnly on parent: $appStreamPath" -Level 'INFO'
                    }

                    # Reset all child items
                    Get-ChildItem -Path $appStreamPath -Recurse -Force -ErrorAction SilentlyContinue |
                    ForEach-Object {
                        if ($_ -is [System.IO.FileInfo] -and $_.IsReadOnly) {
                            $_.IsReadOnly = $false
                        }
                        elseif ($_ -is [System.IO.DirectoryInfo]) {
                            if ($_.Attributes -band [System.IO.FileAttributes]::ReadOnly) {
                                $_.Attributes = $_.Attributes -band (-bnot [System.IO.FileAttributes]::ReadOnly)
                            }
                        }
                    }

                    Remove-Item -Path $appStreamPath -Recurse -Force -ErrorAction Stop
                    $removed = $true
                    Write-Log "  REMOVED on retry: $appStreamPath" -Level 'SUCCESS'
                }
                catch {
                    Write-Log "  RETRY FAILED: $appStreamPath - $($_.Exception.Message)" -Level 'ERROR'
                }
            }

            # Update counters exactly once based on final outcome
            if ($removed) {
                $results.FoldersRemoved++
            }
            else {
                $results.FoldersFailed++
            }
        }
        else {
            $results.ProfilesClean++
        }
    }

    return $results
}

# ============================================================================
# MAIN EXECUTION
# ============================================================================

# Initialize log
$separator = "=" * 70
Set-Content -Path $LogPath -Value $separator -Encoding UTF8 -Force
Write-Log "AppStream 2.0 Residual Cleanup - Script Started"
Write-Log "Computer Name  : $env:COMPUTERNAME"
Write-Log "Executed By    : $env:USERNAME"
Write-Log "Script Version : 1.2"
Write-Log "Target         : <UserProfile>\$TargetFolder"
Write-Log $separator

# Step 1: Kill any running AppStream processes (safety measure)
Write-Log "--- Step 1: Checking for running AppStream processes ---"
Stop-AppStreamProcesses

# Step 2: Scan profiles and remove AppStreamClient folders
Write-Log "--- Step 2: Scanning user profiles and removing residual folders ---"
$results = Remove-AppStreamFromProfiles

# Step 3: Summary
Write-Log $separator
Write-Log "--- SUMMARY ---"
Write-Log "Profiles scanned       : $($results.ProfilesScanned)"
Write-Log "Profiles already clean  : $($results.ProfilesClean)"
Write-Log "Folders found           : $($results.FoldersFound)"
Write-Log "Folders removed         : $($results.FoldersRemoved)"
Write-Log "Folders failed          : $($results.FoldersFailed)"
Write-Log $separator

# Step 4: Determine exit code for SCCM
if ($results.FoldersFailed -gt 0 -and $results.FoldersRemoved -eq 0) {
    Write-Log "EXIT CODE: 1 - All removals failed." -Level 'ERROR'
    Write-Log "AppStream 2.0 Residual Cleanup - Script Ended"
    exit 1
}
elseif ($results.FoldersFailed -gt 0 -and $results.FoldersRemoved -gt 0) {
    Write-Log "EXIT CODE: 2 - Partial success. Some folders could not be removed." -Level 'WARN'
    Write-Log "AppStream 2.0 Residual Cleanup - Script Ended"
    exit 2
}
else {
    Write-Log "EXIT CODE: 0 - Completed successfully." -Level 'SUCCESS'
    Write-Log "AppStream 2.0 Residual Cleanup - Script Ended"
    exit 0
}
