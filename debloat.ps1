<# 
.SYNOPSIS
    Script to remove specific provisioned Appx packages from a Windows system with logging functionality.

.DESCRIPTION
    This PowerShell script identifies and removes specified provisioned Appx packages from the system, 
    ensuring they are uninstalled for all users. It includes a logging mechanism that records each action 
    to a log file for auditing and troubleshooting purposes. The log is stored in the specified directory 
    and captures the timestamp and details of each package removal attempt.

.PARAMETERS
    None. The script targets predefined Appx packages listed within the script body.

.OUTPUTS
    A log file located at 'C:\Windows\fndr\logs\AppX.log' that contains a timestamped record of all actions taken.

.NOTES
    AppX-Debloat.ps1
    Script History:
    Version 1.1 - Enhanced Logging
    Version 1.0 - Script inception
#>
#---------------------------------------------------------------------#
$logFilePath = "C:\Windows\fndr\logs"
$logFileName = "$logFilePath\AppX-Debloat.log"

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

# Function to remove Appx provisioned and installed packages with error handling
function Remove-AppxPackageAndProvisionedPackage {
    param (
        [string]$packageName
    )

    # Log that we're checking for the package's existence
    Write-Log "Checking if package '$packageName' is provisioned or installed for any users."

    # Check if the provisioned package exists
    $provisionedPackage = Get-AppxProvisionedPackage -Online | Where-Object DisplayName -Like "*$packageName*"
    $installedPackage = Get-AppxPackage -AllUsers | Where-Object Name -Like "*$packageName*"

    if ($provisionedPackage) {
        Write-Log "Provisioned package '$packageName' exists. Attempting to remove it..."
        try {
            $provisionedPackage | Remove-AppxProvisionedPackage -Online -ErrorAction Stop
            Write-Log "Successfully removed provisioned package '$packageName'."
        }
        catch {
            Write-Log "Failed to remove provisioned package '$packageName'. Error: $_"
        }
    } else {
        Write-Log "Provisioned package '$packageName' does not exist. Skipping removal."
    }

    if ($installedPackage) {
        Write-Log "Installed package '$packageName' exists for one or more users. Attempting to remove it..."
        try {
            $installedPackage | Remove-AppxPackage -AllUsers -ErrorAction Stop
            Write-Log "Successfully removed installed package '$packageName'."
        }
        catch {
            Write-Log "Failed to remove installed package '$packageName'. Error: $_"
        }
    } else {
        Write-Log "Installed package '$packageName' does not exist for any users. Skipping removal."
    }
}

# Remove the specified Appx Packages
# Microsoft Apps
Remove-AppxPackageAndProvisionedPackage "Microsoft.Microsoft3DViewer"
Remove-AppxPackageAndProvisionedPackage "Microsoft.MSPaint"