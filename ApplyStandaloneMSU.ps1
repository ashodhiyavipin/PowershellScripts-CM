<#
.SYNOPSIS
    Installs a single Microsoft Update Standalone Package (CAB) on an online Windows PC.

.DESCRIPTION
    This script automates the installation of a specified CAB file using the Deployment Image Servicing and Management (DISM) tool.
    It logs all activities, including any errors encountered during the installation process, to a log file for troubleshooting purposes.

.NOTES
    ApplyStandaloneCAB.ps1 - A.Fletcher - 25/10/2024
    Script History:
    Version 1.0 - Script inception
#>
#---------------------------------------------------------------------#

# Define the path for the log file
$logFilePath = "C:\Windows\fndr\logs"
$logFileName = "$logFilePath\ApplyStandaloneCAB.log"

# Function to write logs
function Write-Log {
    param (
        [string]$message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $message"
    Write-Output $logMessage

    # Ensure log file path exists
    if (-not (Test-Path $logFilePath)) {
        New-Item -Path $logFilePath -ItemType Directory | Out-Null
    }

    # Write log message to log file
    Add-Content -Path $logFileName -Value $logMessage
}

# SCCM package's directory is our current working directory so we can directly search for the CAB file in the current directory
$CABFile = Get-ChildItem -Path (Get-Location) -Filter "*.cab" | Select-Object -First 1

if (-not $CABFile) {
    Write-Log "No CAB file found in the working directory: $(Get-Location)"
    Exit 1
}

# Log the CAB file that will be installed
Write-Log "Found CAB file: $($CABFile.FullName)"

# Run the DISM command to install the CAB file, we like using DISM as sometime the old ways are the best :-)
try {
    Write-Log "Starting installation of CAB file: $($CABFile.FullName)"
    $process = Start-Process dism -ArgumentList "/Online", "/Add-Package", "/PackagePath:$($CABFile.FullName)", "/LogPath:$logFileName", "/Quiet", "/NoRestart" -Wait -PassThru

    # Check the exit code of the DISM process
    if ($process.ExitCode -ne 0) {
        Write-Log "Installation failed with error code: $($process.ExitCode)"
        Write-Log "Check the log file for more details: $logFileName"
        Exit $process.ExitCode
    } else {
        Write-Log "CAB file installed successfully!"
        Write-Log "Log file available at: $logFileName"
    }
} catch {
    Write-Log "An error occurred during the installation process. Error: $_"
    Exit 1
}