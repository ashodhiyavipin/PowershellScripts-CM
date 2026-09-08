#=============================================================================
# Script:  Detect-ComponentStoreCorruption.ps1
# Purpose: Runs DISM health checks, parses CBS.log for corruption summary,
#          writes results to registry for SCCM Configuration Baseline pickup.
# Author:  Vipin / Global Software Delivery Engineering
# Version: 2.2
# Date:    2026-09-07
#
# Changelog from v1.0 (code review fixes):
#  - Fixed Extract-Value returning the whole $Matches hashtable instead of
#     the captured group (this broke every downstream value).
#  - CBS.log parsing no longer depends on a literal "Staged Packages:"
#     anchor; it now captures the summary block up to the next blank line,
#     which is far more resilient across OS builds/log formats.
#  - Parse failures no longer fall back to "HEALTHY" or trust ScanHealth's
#     exit code as a corruption signal (it doesn't reliably indicate that).
#     A failed parse now reports UNKNOWN and exits non-zero so the baseline
#     flags it for review instead of silently passing.
#  - RestoreHealth is now OFF by default (-RunRestoreHealth to enable) since
#     it can hang for a long time against a blocked CDN, fleet-wide, on every
#     baseline evaluation.
#  - Added a manual timeout around each DISM invocation so a hung process
#     doesn't hang the whole baseline remediation cycle.
#  - CBS.log is now tailed instead of fully loaded with -Raw, to avoid
#     loading huge logs into memory on exactly the machines most likely to
#     have large ones.
#  - Added try/catch around DISM calls, registry writes, and log parsing.
#  - Removed the non-functional Send-SCCMStatusMessage stub (it created a
#     COM object and never called anything on it) in favor of a documented,
#     registry-only status surface. See note at bottom if you want a real
#     status-message implementation.
#  - Renamed functions to use approved PowerShell verbs.
#  - Added simple size-based log rotation.
#  - Timestamp is now captured after the scan completes, not before.
#=============================================================================

[CmdletBinding()]
param(
    # RestoreHealth is expensive and, if the CDN is blocked as expected in
    # this environment, will spend time retrying before failing. Off by
    # default - turn on only where you specifically want repair attempted.
    [switch]$RunRestoreHealth,

    # Per-DISM-call timeout. A hung DISM process (component store lock, AV
    # interference, etc.) will be killed after this many minutes rather than
    # hanging the whole baseline evaluation indefinitely.
    [int]$DismTimeoutMinutes = 15,

    # How many lines from the end of CBS.log to search. CBS.log can be huge;
    # this avoids loading the whole file for every run. Increase if your
    # summary blocks aren't being found within this window.
    [int]$CBSTailLines = 8000,

    # Rotate our own log once it passes this size.
    [int]$LogRotateSizeMB = 5
)

#region --- Configuration ---
$RegistryPath = "HKLM:\SOFTWARE\Foundever\ComponentStoreHealth"
$LogFolder    = "$env:SystemRoot\fndr\logs"
$LogFile      = Join-Path $LogFolder "ComponentStoreHealthCheck.log"
$CBSLogPath   = "$env:SystemRoot\Logs\CBS\CBS.log"
#endregion

#region --- Functions ---
function Write-Log {
    <#
    .SYNOPSIS
        Writes a timestamped line to the log file and echoes it to the host.

    .NOTES
        Deliberately does NOT use Write-Output / does NOT return a value.
        Every caller of this function (Invoke-DISMCommand, Get-CBSLogSummary,
        etc.) calls Write-Log many times as a bare statement. If Write-Log
        emitted anything to the success/output stream, each of those calls
        would silently become part of the CALLING function's own return
        value - e.g. Invoke-DISMCommand's "return $exitCode" would actually
        come back as @(logline1, logline2, ..., $exitCode), an array, not an
        int. That is exactly what caused:
            "Cannot convert the System.Object[] value ... to type Int32"
        Write-Host (host stream) and file writes below are both safe: they
        cannot be accidentally captured by a caller's variable assignment.

        Add-Content is wrapped with a short retry: log viewers like CMTrace
        can briefly hold the file in a share mode that blocks writers, and
        that should degrade to "one log line delayed/dropped," never to a
        wall of repeated terminal errors or a script failure.
    #>
    param([string]$Message)

    $Entry = "[$( Get-Date -Format 'yyyy-MM-dd HH:mm:ss' )] $Message"
    Write-Host $Entry

    if (-not (Test-Path $LogFolder)) {
        try { New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null } catch { }
    }

    $attempts = 0
    $maxAttempts = 3
    while ($attempts -lt $maxAttempts) {
        try {
            Add-Content -Path $LogFile -Value $Entry -Force -ErrorAction Stop
            break
        }
        catch {
            $attempts++
            if ($attempts -ge $maxAttempts) {
                # Give up on this line rather than crash the run - e.g. the
                # file is open exclusively in CMTrace or a similar viewer.
                # The host output above still shows the message.
                break
            }
            Start-Sleep -Milliseconds 200
        }
    }
}

function Invoke-LogRotation {
    if (Test-Path $LogFile) {
        $sizeMB = (Get-Item $LogFile).Length / 1MB
        if ($sizeMB -ge $LogRotateSizeMB) {
            $archiveName = "ComponentStoreHealthCheck_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss")
            $archivePath = Join-Path $LogFolder $archiveName
            try {
                Move-Item -Path $LogFile -Destination $archivePath -Force
            }
            catch {
                # Non-fatal - worst case the log keeps growing this run.
            }
        }
    }
}

function Invoke-DISMCommand {
    <#
    .SYNOPSIS
        Runs a DISM command with a hard timeout so a hung process can't
        block the baseline evaluation cycle indefinitely.

    .NOTES
        Touches $process.Handle immediately after Start-Process. This is a
        documented .NET/PowerShell quirk: if you never access .Handle before
        the process exits, Windows can invalidate the process handle the
        instant it terminates, and .ExitCode then comes back empty even
        though WaitForExit() correctly reported the process as finished.
        That was silently happening here - $exitCode was blank, and it only
        avoided throwing downstream because [int]$null casts to 0, which
        would have recorded every DISM run as "succeeded" regardless of what
        actually happened. Touching .Handle forces .NET to retain the handle
        so ExitCode is reliably populated afterward.
    #>
    param(
        [string]$ArgumentString,
        [string]$StepName,
        [int]$TimeoutMinutes = 15
    )

    Write-Log "Running: DISM $ArgumentString (timeout: $TimeoutMinutes min)"

    try {
        $process = Start-Process -FilePath "DISM.exe" `
                                 -ArgumentList $ArgumentString `
                                 -NoNewWindow -PassThru

        # Force .NET to retain the process handle - see .NOTES above.
        $null = $process.Handle

        $timedOut = -not $process.WaitForExit($TimeoutMinutes * 60 * 1000)

        if ($timedOut) {
            Write-Log "ERROR: $StepName exceeded $TimeoutMinutes minute timeout - killing process."
            try { $process.Kill() } catch { }
            return -1
        }

        $exitCode = $process.ExitCode

        if ($null -eq $exitCode -or $exitCode -eq "") {
            # Belt-and-suspenders: if ExitCode still comes back empty despite
            # the Handle fix, do NOT let it silently become 0 (success) via
            # an [int]$null cast later. Surface it as -1 (unknown) instead,
            # so a real DISM failure can never be recorded as a pass.
            Write-Log "WARNING: $StepName process exited but ExitCode could not be read - recording as -1 (unknown), not assuming success."
            return -1
        }

        Write-Log "$StepName completed with exit code: $exitCode"
        return $exitCode
    }
    catch {
        Write-Log "ERROR: Failed to run DISM for $StepName - $($_.Exception.Message)"
        return -1
    }
}

function Get-CBSLogSummary {
    <#
    .SYNOPSIS
        Parses the CBS.log Summary block produced by DISM ScanHealth.
        Looks for the LAST summary block in the log (most recent scan).
        Returns $null if no summary block could be found - callers must
        NOT treat that as "healthy".
    #>
    param(
        # Explicit params instead of relying on the parent scope: this
        # function resolves fine today via PowerShell's scope chain, but an
        # implicit dependency on script-scoped variables breaks silently
        # (falls back to $null / full-file reads) the moment this function
        # is copied into a module or called from another script.
        [Parameter(Mandatory)]
        [string]$LogPath,

        [Parameter(Mandatory)]
        [int]$TailLines
    )

    Write-Log "Parsing CBS.log at: $LogPath (tailing last $TailLines lines)"

    if (-not (Test-Path $LogPath)) {
        Write-Log "ERROR: CBS.log not found at $LogPath"
        return $null
    }

    try {
        # Tail instead of -Raw: avoids loading a multi-hundred-MB log into
        # memory, which is common on exactly the machines this runs against.
        $cbsLines = Get-Content -Path $LogPath -Tail $TailLines -ErrorAction Stop
    }
    catch {
        Write-Log "ERROR: Could not read CBS.log - $($_.Exception.Message)"
        return $null
    }

    $cbsContent = $cbsLines -join "`n"

    # Capture from "Summary:" up to the next blank line, rather than
    # anchoring on a specific literal like "Staged Packages:" which is not
    # guaranteed to immediately follow across builds/log variants.
    $summaryPattern = '(?ms)Info\s+CBS\s+Summary:(?<body>.*?)(?:\r?\n\s*\r?\n|\z)'
    $summaryMatches = [regex]::Matches($cbsContent, $summaryPattern)

    if ($summaryMatches.Count -eq 0) {
        Write-Log "WARNING: No Summary block found in the tailed portion of CBS.log"
        return $null
    }

    $lastSummary = $summaryMatches[$summaryMatches.Count - 1].Groups['body'].Value
    Write-Log "Found Summary block (using last occurrence of $($summaryMatches.Count) found in tail window)"

    $result = @{
        TotalDetectedCorruption = Get-RegexValue -Text $lastSummary -Pattern 'Total Detected Corruption:\s*(\d+)'
        CBSManifestCorruption   = Get-RegexValue -Text $lastSummary -Pattern 'CBS Manifest Corruption:\s*(\d+)'
        CBSMetadataCorruption   = Get-RegexValue -Text $lastSummary -Pattern 'CBS Metadata Corruption:\s*(\d+)'
        CSIManifestCorruption   = Get-RegexValue -Text $lastSummary -Pattern 'CSI Manifest Corruption:\s*(\d+)'
        CSIMetadataCorruption   = Get-RegexValue -Text $lastSummary -Pattern 'CSI Metadata Corruption:\s*(\d+)'
        CSIPayloadCorruption    = Get-RegexValue -Text $lastSummary -Pattern 'CSI Payload Corruption:\s*(\d+)'
        CSIFileFlagsCorrupt     = Get-RegexValue -Text $lastSummary -Pattern 'CSI FileFlags Corrupt:\s*(\d+)'
        TotalRepairedCorruption = Get-RegexValue -Text $lastSummary -Pattern 'Total Repaired Corruption:\s*(\d+)'
        OperationResult         = Get-RegexValue -Text $lastSummary -Pattern 'Operation result:\s*(0x[0-9a-fA-F]+)'
    }

    foreach ($key in $result.Keys) {
        Write-Log "  $key = $($result[$key])"
    }

    return $result
}

function Get-RegexValue {
    <#
    .SYNOPSIS
        Returns the first captured group from a regex match, or "N/A".
        (v1 bug: this used to return the whole $Matches hashtable.)
    #>
    param(
        [string]$Text,
        [string]$Pattern
    )
    if ($Text -match $Pattern) {
        return $Matches[1]
    }
    return "N/A"
}

function Set-RegistryResults {
    param(
        [hashtable]$Results,
        [string]$Status,
        [string]$Timestamp
    )

    Write-Log "Writing results to registry: $RegistryPath"

    try {
        if (-not (Test-Path $RegistryPath)) {
            New-Item -Path $RegistryPath -Force | Out-Null
        }

        Set-ItemProperty -Path $RegistryPath -Name "Status"      -Value $Status   -Type String -Force
        Set-ItemProperty -Path $RegistryPath -Name "LastScanDate" -Value $Timestamp -Type String -Force

        # DISM exit codes - coalesce nulls to -1 so -Type DWord never fails.
        Set-ItemProperty -Path $RegistryPath -Name "ScanHealthExitCode"   -Value ([int]($Results.ScanHealthExitCode))   -Type DWord -Force
        Set-ItemProperty -Path $RegistryPath -Name "CheckHealthExitCode"  -Value ([int]($Results.CheckHealthExitCode))  -Type DWord -Force
        Set-ItemProperty -Path $RegistryPath -Name "RestoreHealthExitCode" -Value ([int]($Results.RestoreHealthExitCode)) -Type DWord -Force

        # CBS parsed values (string - may legitimately be "N/A")
        Set-ItemProperty -Path $RegistryPath -Name "TotalDetectedCorruption" -Value $Results.TotalDetectedCorruption -Type String -Force
        Set-ItemProperty -Path $RegistryPath -Name "CSIPayloadCorruption"   -Value $Results.CSIPayloadCorruption   -Type String -Force
        Set-ItemProperty -Path $RegistryPath -Name "CBSManifestCorruption"  -Value $Results.CBSManifestCorruption  -Type String -Force
        Set-ItemProperty -Path $RegistryPath -Name "CSIManifestCorruption"  -Value $Results.CSIManifestCorruption  -Type String -Force
        Set-ItemProperty -Path $RegistryPath -Name "TotalRepairedCorruption" -Value $Results.TotalRepairedCorruption -Type String -Force
        Set-ItemProperty -Path $RegistryPath -Name "OperationResult"        -Value $Results.OperationResult        -Type String -Force

        Write-Log "Registry updated - Status: $Status"
    }
    catch {
        Write-Log "ERROR: Failed to write registry results - $($_.Exception.Message)"
        throw
    }
}
#endregion

#region --- Main Execution ---
Invoke-LogRotation

Write-Log "=========================================="
Write-Log "Component Store Health Check - START"
Write-Log "Computer: $env:COMPUTERNAME"
Write-Log "=========================================="

$checkHealthExit   = Invoke-DISMCommand -ArgumentString "/Online /Cleanup-Image /CheckHealth" -StepName "CheckHealth" -TimeoutMinutes $DismTimeoutMinutes
$scanHealthExit    = Invoke-DISMCommand -ArgumentString "/Online /Cleanup-Image /ScanHealth" -StepName "ScanHealth" -TimeoutMinutes $DismTimeoutMinutes

$restoreHealthExit = -1
if ($RunRestoreHealth) {
    $restoreHealthExit = Invoke-DISMCommand -ArgumentString "/Online /Cleanup-Image /RestoreHealth" -StepName "RestoreHealth" -TimeoutMinutes $DismTimeoutMinutes
}
else {
    Write-Log "RestoreHealth skipped (use -RunRestoreHealth to enable). Detection only."
}

$cbsResults = Get-CBSLogSummary -LogPath $CBSLogPath -TailLines $CBSTailLines

# --- Status determination ---
# Three real outcomes, not two:
#   HEALTHY - parse succeeded, zero corruption detected
#   CORRUPT - parse succeeded, corruption detected
#   UNKNOWN - parse failed; we do NOT guess HEALTHY here. ScanHealth's exit
#              code does not reliably indicate corruption, so it is not used
#              as a corruption signal - only as a run-failure signal.
if ($cbsResults -and $cbsResults.TotalDetectedCorruption -ne "N/A") {
    $totalCorruption = [int]$cbsResults.TotalDetectedCorruption
    if ($totalCorruption -gt 0) {
        $status = "CORRUPT"
        Write-Log "*** CORRUPTION DETECTED - Total: $totalCorruption ***"
    }
    else {
        $status = "HEALTHY"
        Write-Log "Machine is HEALTHY - no corruption detected."
    }
}
else {
    $totalCorruption = "N/A"
    $status = "UNKNOWN"
    Write-Log "*** COULD NOT DETERMINE STATUS - CBS.log summary not found/parseable. Flagging for review. ***"
    if ($scanHealthExit -eq -1) {
        Write-Log "  (ScanHealth itself failed/timed out - exit code -1)"
    }
}

$Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

$allResults = @{
    ScanHealthExitCode      = $scanHealthExit
    CheckHealthExitCode     = $checkHealthExit
    RestoreHealthExitCode   = if ($RunRestoreHealth) { $restoreHealthExit } else { -1 }
    TotalDetectedCorruption = if ($cbsResults) { $cbsResults.TotalDetectedCorruption } else { "N/A" }
    CSIPayloadCorruption    = if ($cbsResults) { $cbsResults.CSIPayloadCorruption } else { "N/A" }
    CBSManifestCorruption   = if ($cbsResults) { $cbsResults.CBSManifestCorruption } else { "N/A" }
    CSIManifestCorruption   = if ($cbsResults) { $cbsResults.CSIManifestCorruption } else { "N/A" }
    TotalRepairedCorruption = if ($cbsResults) { $cbsResults.TotalRepairedCorruption } else { "N/A" }
    OperationResult         = if ($cbsResults) { $cbsResults.OperationResult } else { "N/A" }
}

try {
    Set-RegistryResults -Results $allResults -Status $status -Timestamp $Timestamp
}
catch {
    Write-Log "FATAL: Registry write failed - baseline will not see updated status."
    Write-Log "=========================================="
    exit 3
}

Write-Log "=========================================="
Write-Log "Component Store Health Check - END"
Write-Log "Status: $status"
Write-Log "=========================================="

# Exit codes for the Configuration Baseline / CI script evaluation:
#   0 = compliant (HEALTHY)
#   1 = non-compliant (CORRUPT) - remediation should run
#   2 = indeterminate (UNKNOWN) - treat as non-compliant so it surfaces for
#       manual review rather than silently passing
#   3 = script-level failure (registry write failed, see log)
#   4 = internal error - $status held an unrecognized value (should never
#       happen; indicates a code defect if seen, see log)
switch ($status) {
    "HEALTHY" { exit 0 }
    "CORRUPT" { exit 1 }
    "UNKNOWN" { exit 2 }
    default   {
        Write-Log "ERROR: Unrecognized status value '$status' - exiting non-zero rather than falling through to an implicit 0."
        exit 4
    }
}
#endregion

#=============================================================================
# NOTE on SCCM status messages:
# v1 had a Send-SCCMStatusMessage function that created a
# Microsoft.SMS.Client COM object but never called anything on it - it only
# wrote to this script's own log, so it was effectively dead code. It's
# removed here. The registry values under $RegistryPath are what the
# Configuration Baseline should read for compliance state.
#
# If you actually want a status message visible in SCCM Monitoring (not just
# CB compliance state), that requires either:
#  - MsiProvider/WriteStatusMessage via the client SDK with a real
#     component/message ID registered in your site's .MOF, or
#  - Writing to Software Center Endpoint via LogPropertyBag /
#     the SMS_TASK API.
# That's a separate piece of work from this detection script and worth
# scoping on its own - happy to help design it if you want status messages
# rather than just baseline compliance state.
#=============================================================================
