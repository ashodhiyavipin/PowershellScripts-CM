<#
.SYNOPSIS
    Installs a single Microsoft Update Standalone Package (MSU) on an online Windows PC.

.DESCRIPTION
    This script automates the installation of a specified MSU file using the Deployment Image Servicing and Management (DISM) tool.
    It logs all activities, including any errors encountered during the installation process, to a log file for troubleshooting purposes.

.NOTES
    ApplyStandaloneMSU.ps1 - A.Fletcher - 29/08/2024
    Script History:
    Version 1.0 - Script inception
#>
#---------------------------------------------------------------------#

# Define the path for the log file
$logFilePath = "C:\Windows\fndr\logs"
$logFileName = "$logFilePath\ApplyStandaloneMSU.log"

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

# SCCM package's directory is our current working directory so we can directly search for the MSU file in the current directory
$MSUFile = Get-ChildItem -Path (Get-Location) -Filter "*.msu" | Select-Object -First 1

if (-not $MSUFile) {
    Write-Log "No MSU file found in the working directory: $(Get-Location)"
    Exit 1
}

# Log the MSU file that will be installed
Write-Log "Found MSU file: $($MSUFile.FullName)"

# Run the DISM command to install the MSU file, we like using DISM as sometime the old ways are the best :-)
try {
    Write-Log "Starting installation of MSU file: $($MSUFile.FullName)"
    $process = Start-Process wusa $($MSUFile.FullName), "/quiet", "/norestart", "/LogPath:$logFileName" -Wait -PassThru

    # Check the exit code of the DISM process
    if ($process.ExitCode -ne 0) {
        Write-Log "Installation failed with error code: $($process.ExitCode)"
        Write-Log "Check the log file for more details: $logFileName"
        Exit $process.ExitCode
    } else {
        Write-Log "MSU file installed successfully!"
        Write-Log "Log file available at: $logFileName"
    }
} catch {
    Write-Log "An error occurred during the installation process. Error: $_"
    Exit 1
}