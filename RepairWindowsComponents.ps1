<#
.SYNOPSIS
    Performs CBS component repair to fix all types of Windows Update Agent corruptions using DISM and system file checks.
.DESCRIPTION
    This script automates the execution of wsreset, DISM, and SFC commands to repair the component store and system integrity in Windows.
    It waits for each process to finish and logs all activities, including any errors encountered during the execution process, to a log file for troubleshooting purposes.
.NOTES
    RepairWindowsComponents.ps1 - V.Ashodhiya.
    Date: 19-02-2025
    Script History:
    Version 1.0 - Script inception
#>

# Define the path for the log file
$logFilePath = "C:\Windows\fndr\logs"
$logFileName = "$logFilePath\RepairWindowsComponents.log"
# Function to write logs
function Write-Log{
    param (
        [string]$message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $message"
    Write-Logut $logMessage
 
    # Ensure log file path exists
    if (-not (Test-Path $logFilePath)) {
        New-Item -Path $logFilePath -ItemType Directory | Out-Null
    }
 
    # Write Log Write-Logmessage to log file
    Add-Content -Path $logFileName -Value $logMessage
}

# Start the logging
Write-Log "Script execution started."

# 1. Run wsreset.exe and wait for the process to finish
try {
    Write-Log "Executing wsreset.exe..."
    Start-Process -FilePath "wsreset.exe" -Wait -PassThru
    Write-Log "wsreset.exe completed successfully."
} catch {
    Write-Log "wsreset.exe failed with error: $_"
}

# 2. Run DISM /online /cleanup-image /restorehealth and wait for the process to finish
try {
    Write-Log "Executing DISM command..."
    Start-Process -FilePath "DISM.exe" -ArgumentList "/online", "/cleanup-image", "/restorehealth" -Wait -PassThru
    Write-Log "DISM command completed successfully."
} catch {
    Write-Log "DISM command failed with error: $_"
}

# 3. Run sfc /scannow and wait for the process to finish
try {
    Write-Log "Executing sfc /scannow..."
    Start-Process -FilePath "sfc.exe" -ArgumentList "/scannow" -Wait -PassThru
    Write-Log "sfc /scannow completed successfully."
} catch {
    Write-Log "sfc /scannow failed with error: $_"
}

# End the logging
Write-Log "Script execution completed."