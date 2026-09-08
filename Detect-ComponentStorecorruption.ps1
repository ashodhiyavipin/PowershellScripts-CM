#=============================================================================
# Script:  Detect-ComponentStoreCorruption.ps1
# Purpose: Runs DISM health checks, parses CBS.log for corruption summary,
#          writes results to registry for SCCM Configuration Baseline pickup.
# Author:  Vipin / Global Software Delivery Engineering
# Version: 1.0
# Date:    2026-09-07
#=============================================================================

#region --- Configuration ---
$RegistryPath    = "HKLM:\SOFTWARE\Foundever\ComponentStoreHealth"
$LogFolder       = "$env:SystemRoot\Logs\Foundever"
$LogFile         = Join-Path $LogFolder "ComponentStoreHealthCheck.log"
$CBSLogPath      = "$env:SystemRoot\Logs\CBS\CBS.log"
$Timestamp       = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
#endregion

#region --- Functions ---
function Write-Log {
    param([string]$Message)
    $Entry = "[$( Get-Date -Format 'yyyy-MM-dd HH:mm:ss' )] $Message"
    if (-not (Test-Path $LogFolder)) { New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null }
    Add-Content -Path $LogFile -Value $Entry -Force
    Write-Output $Entry
}

function Run-DISMCommand {
    param(
        [string]$ArgumentString,
        [string]$StepName
    )
    Write-Log "Running: DISM $ArgumentString"
    $process = Start-Process -FilePath "DISM.exe" `
                             -ArgumentList $ArgumentString `
                             -Wait -NoNewWindow -PassThru
    $exitCode = $process.ExitCode
    Write-Log "$StepName completed with exit code: $exitCode"
    return $exitCode
}

function Parse-CBSLogSummary {
    <#
    .SYNOPSIS
        Parses the CBS.log Summary block produced by DISM ScanHealth.
        Looks for the LAST summary block in the log (most recent scan).
    #>
    Write-Log "Parsing CBS.log at: $CBSLogPath"

    if (-not (Test-Path $CBSLogPath)) {
        Write-Log "ERROR: CBS.log not found at $CBSLogPath"
        return $null
    }

    $cbsContent = Get-Content -Path $CBSLogPath -Raw

    $summaryPattern = '(?s)Info\s+CBS\s+Summary:.*?(?=\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2},\s+Info\s+CBS\s+Staged Packages:)'
    $regexMatches = [regex]::Matches($cbsContent, $summaryPattern)

    if ($regexMatches.Count -eq 0) {
        Write-Log "WARNING: No Summary block found in CBS.log"
        return $null
    }

    $lastSummary = $regexMatches[$regexMatches.Count - 1].Value
    Write-Log "Found Summary block (using last occurrence of $($regexMatches.Count) found)"

    $result = @{
        TotalDetectedCorruption  = Extract-Value $lastSummary 'Total Detected Corruption:\s*(\d+)'
        CBSManifestCorruption    = Extract-Value $lastSummary 'CBS Manifest Corruption:\s*(\d+)'
        CBSMetadataCorruption    = Extract-Value $lastSummary 'CBS Metadata Corruption:\s*(\d+)'
        CSIManifestCorruption    = Extract-Value $lastSummary 'CSI Manifest Corruption:\s*(\d+)'
        CSIMetadataCorruption    = Extract-Value $lastSummary 'CSI Metadata Corruption:\s*(\d+)'
        CSIPayloadCorruption     = Extract-Value $lastSummary 'CSI Payload Corruption:\s*(\d+)'
        CSIFileFlagsCorrupt      = Extract-Value $lastSummary 'CSI FileFlags Corrupt:\s*(\d+)'
        TotalRepairedCorruption  = Extract-Value $lastSummary 'Total Repaired Corruption:\s*(\d+)'
        OperationResult          = Extract-Value $lastSummary 'Operation result:\s*(0x[0-9a-fA-F]+)'
    }

    foreach ($key in $result.Keys) {
        Write-Log "  $key = $($result[$key])"
    }

    return $result
}

function Extract-Value {
    param(
        [string]$Text,
        [string]$Pattern
    )
    if ($Text -match $Pattern) {
        return $Matches
    }
    return "N/A"
}

function Write-RegistryResults {
    param([hashtable]$Results, [string]$Status)

    Write-Log "Writing results to registry: $RegistryPath"

    if (-not (Test-Path $RegistryPath)) {
        New-Item -Path $RegistryPath -Force | Out-Null
    }

    Set-ItemProperty -Path $RegistryPath -Name "Status" -Value $Status -Type String -Force
    Set-ItemProperty -Path $RegistryPath -Name "LastScanDate" -Value $Timestamp -Type String -Force

    Set-ItemProperty -Path $RegistryPath -Name "ScanHealthExitCode" -Value $Results.ScanHealthExitCode -Type DWord -Force
    Set-ItemProperty -Path $RegistryPath -Name "CheckHealthExitCode" -Value $Results.CheckHealthExitCode -Type DWord -Force
    Set-ItemProperty -Path $RegistryPath -Name "RestoreHealthExitCode" -Value $Results.RestoreHealthExitCode -Type DWord -Force

    Set-ItemProperty -Path $RegistryPath -Name "TotalDetectedCorruption" -Value $Results.TotalDetectedCorruption -Type String -Force
    Set-ItemProperty -Path $RegistryPath -Name "CSIPayloadCorruption" -Value $Results.CSIPayloadCorruption -Type String -Force
    Set-ItemProperty -Path $RegistryPath -Name "CBSManifestCorruption" -Value $Results.CBSManifestCorruption -Type String -Force
    Set-ItemProperty -Path $RegistryPath -Name "CSIManifestCorruption" -Value $Results.CSIManifestCorruption -Type String -Force
    Set-ItemProperty -Path $RegistryPath -Name "TotalRepairedCorruption" -Value $Results.TotalRepairedCorruption -Type String -Force
    Set-ItemProperty -Path $RegistryPath -Name "OperationResult" -Value $Results.OperationResult -Type String -Force

    Write-Log "Registry updated - Status: $Status"
}

function Send-SCCMStatusMessage {
    param(
        [string]$Status,
        [int]$CorruptionCount
    )
    try {
        $SCCMClient = New-Object -ComObject Microsoft.SMS.Client
        $messageText = "ComponentStoreHealth: $Status | TotalCorruption: $CorruptionCount | Machine: $env:COMPUTERNAME"

        if ($Status -eq "CORRUPT") {
            Write-Log "Sending SCCM status message: CORRUPT (ID: 40001)"
        }
        else {
            Write-Log "Sending SCCM status message: HEALTHY (ID: 40002)"
        }
        Write-Log "Status message content: $messageText"
    }
    catch {
        Write-Log "WARNING: Could not send SCCM status message - $($_.Exception.Message)"
    }
}
#endregion

#region --- Main Execution ---
Write-Log "=========================================="
Write-Log "Component Store Health Check - START"
Write-Log "Computer: $env:COMPUTERNAME"
Write-Log "=========================================="

# Step 1: Run CheckHealth
$checkHealthExit = Run-DISMCommand "/Online /Cleanup-Image /CheckHealth" "CheckHealth"

# Step 2: Run ScanHealth (full deep scan)
$scanHealthExit = Run-DISMCommand "/Online /Cleanup-Image /ScanHealth" "ScanHealth"

# Step 3: Run RestoreHealth (expected to fail - confirms CDN block)
$restoreHealthExit = Run-DISMCommand "/Online /Cleanup-Image /RestoreHealth" "RestoreHealth"

# Step 4: Parse CBS.log Summary
$cbsResults = Parse-CBSLogSummary

# Step 5: Determine status
$totalCorruption = 0
if ($cbsResults -and $cbsResults.TotalDetectedCorruption -ne "N/A") {
    $totalCorruption = [int]$cbsResults.TotalDetectedCorruption
}

if ($totalCorruption -gt 0) {
    $status = "CORRUPT"
    Write-Log "*** CORRUPTION DETECTED - Total: $totalCorruption ***"
}
elseif ($scanHealthExit -ne 0) {
    $status = "CORRUPT"
    Write-Log "*** CORRUPTION DETECTED (via exit code) - ScanHealth exit: $scanHealthExit ***"
}
else {
    $status = "HEALTHY"
    Write-Log "Machine is HEALTHY - no corruption detected."
}

# Step 6: Build results hashtable
$allResults = @{
    ScanHealthExitCode       = $scanHealthExit
    CheckHealthExitCode      = $checkHealthExit
    RestoreHealthExitCode    = $restoreHealthExit
    TotalDetectedCorruption  = if ($cbsResults) { $cbsResults.TotalDetectedCorruption } else { "N/A" }
    CSIPayloadCorruption     = if ($cbsResults) { $cbsResults.CSIPayloadCorruption } else { "N/A" }
    CBSManifestCorruption    = if ($cbsResults) { $cbsResults.CBSManifestCorruption } else { "N/A" }
    CSIManifestCorruption    = if ($cbsResults) { $cbsResults.CSIManifestCorruption } else { "N/A" }
    TotalRepairedCorruption  = if ($cbsResults) { $cbsResults.TotalRepairedCorruption } else { "N/A" }
    OperationResult          = if ($cbsResults) { $cbsResults.OperationResult } else { "N/A" }
}

# Step 7: Write to registry
Write-RegistryResults -Results $allResults -Status $status

# Step 8: Send SCCM status message
Send-SCCMStatusMessage -Status $status -CorruptionCount $totalCorruption

Write-Log "=========================================="
Write-Log "Component Store Health Check - END"
Write-Log "Status: $status"
Write-Log "=========================================="

# Exit code for CI/CB evaluation
# Exit 0 = compliant (HEALTHY), Exit 1 = non-compliant (CORRUPT)
if ($status -eq "CORRUPT") {
    exit 1
}
else {
    exit 0
}
#endregion