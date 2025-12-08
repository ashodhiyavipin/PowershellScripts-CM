<# 
.SYNOPSIS
    Script to remove all old versions of a specific provisioned Appx package from a Windows system with logging functionality.

.DESCRIPTION
    This PowerShell script identifies and removes all old versions of a specific provisioned Appx package from the system, 
    ensuring they are uninstalled for all users. It includes a logging mechanism that records each action 
    to a log file for auditing and troubleshooting purposes. The log is stored in the specified directory 
    and captures the timestamp and details of each package removal attempt.

.PARAMETERS
    targetApp
        The name of the Appx package to target for removal (e.g., "Microsoft.YourApp").

.OUTPUTS
    A log file located at 'C:\Windows\fndr\logs\RemoveOldAppxVersions.log' that contains a timestamped record of all actions taken.

.NOTES
    RemoveOldAppxVersions.ps1
    Script History:
    Version 1.0 - Script inception
#>
#---------------------------------------------------------------------#
param(
    [Parameter(Mandatory = $true)]
    [string]$targetApp
)

$logFilePath = "C:\Windows\fndr\logs"
$logFileName = "$logFilePath\RemoveOldAppxVersions.log"

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
function OldPackages {
    param (
        [string]$packageName
    )

    Write-Log "Searching for all versions of package '$packageName'..."

    # Get installed and provisioned packages
    $installedPackages = Get-AppxPackage -AllUsers | Where-Object { $_.Name -like "*$packageName*" }
    $provisionedPackages = Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -like "*$packageName*" }

    # Combine both sets (normalize to objects with Name + Version)
    $allPackages = @()
    if ($installedPackages) {
        $allPackages += $installedPackages | Select-Object Name, Version, PackageFullName
    }
    if ($provisionedPackages) {
        $allPackages += $provisionedPackages | Select-Object DisplayName, Version, PackageName
    }

    if (-not $allPackages) {
        Write-Log "No packages found for '$packageName'."
        return @()
    }

    # Sort by version and identify latest
    $sorted = $allPackages | Sort-Object Version -Descending
    $latest = $sorted[0]
    $older = $sorted | Select-Object -Skip 1

    Write-Log "Latest version detected: $($latest.Version)"
    if ($older) {
        Write-Log "Found $($older.Count) older versions to remove."
    }
    else {
        Write-Log "No older versions found."
    }

    return $older
}

# Function to remove Appx provisioned and installed packages with error handling
function Remove-AppxPackageAndProvisionedPackage {
    param ( [string]$packageName )

    Write-Log "Checking if package '$packageName' is provisioned or installed for any users."

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
    }
    else {
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
    }
    else {
        Write-Log "Installed package '$packageName' does not exist for any users. Skipping removal."
    }
}

# ------------------ Main Logic ------------------ #


# Get all old versions
$oldVersions = OldPackages -packageName $targetApp

# Remove each old version
foreach ($pkg in $oldVersions) {
    $pkgName = $pkg.Name
    Write-Log "Preparing to remove old version: $pkgName ($($pkg.Version))"
    Remove-AppxPackageAndProvisionedPackage -packageName $pkgName
}

Write-Log "Script completed."