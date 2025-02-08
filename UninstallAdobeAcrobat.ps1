<# 
.SYNOPSIS 
The provided PowerShell script is designed to uninstall Adobe Acrobat Reader from a Windows operating system. It checks for installed versions of Adobe Reader, excluding specific editions and a particular version, and performs an uninstallation using the MSIExec command. The script includes logging functionality to track both successful and failed operations, and it ensures that it is run with administrative privileges.
 
.DESCRIPTION  
The script fetches a list of installed programs from the Windows registry for both 32-bit and 64-bit architectures.
It filters the retrieved installed programs to identify entries for Adobe Acrobat Reader
If eligible Adobe Reader installations are found, the script attempts to uninstall each one using their product codes via the msiexec command. Success or failure of the uninstallation process is logged.
The script exits with a status code indicating success (0) or failure (1) based on the outcome of the uninstallation attempts.


.NOTES
Script Version: 1.0
Last Updated: 29 Jan 2025
Author: Vipin Anand Ashodhiya
Change Log:
1.0: Initial creation of the script; added logging functionality and filtering for Adobe Acrobat Reader installations.


#>
#Logging Function
$logFilePath = "C:\Windows\fndr\logs"
$logFileName = "$logFilePath\AdobeReaderUninstall.log"
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

# Check if the script is running with administrative privileges
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Please run this script as an administrator."
    Pause
    exit
}

# Get installed programs for both 32-bit and 64-bit architectures
$paths = @('HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\','HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\')

$installedPrograms = foreach ($registryPath in $paths) {
    try {
        Get-ChildItem -LiteralPath $registryPath | Get-ItemProperty | Where-Object { $_.PSChildName -ne $null }
    } catch {
        Write-Log ("Failed to access registry path: $registryPath. Error: $_")
        return @()
    }
}

# Filter programs with Adobe Acrobat Reader in their display name, excluding Standard and Professional and version 24.002.20965
# Change Line 66 if the verison is updates
$adobeReaderEntries = $installedPrograms | Where-Object {
    $_.DisplayName -like '*Adobe Acrobat*'
# and $_.DisplayName -notlike '*Standard*' -and
#    $_.DisplayName -notlike '*Professional*' -and
#    $_.DisplayVersion -notlike '*24.002.20965*'
}

if ($adobeReaderEntries.Count -eq 0) {
    Write-Log "No Adobe Acrobat Reader installations found to uninstall."
    Write-Host "No Adobe Reader to Uninstall"
    pause
    exit 1
}

# Try to uninstall Adobe Acrobat Reader for each matching entry
foreach ($entry in $adobeReaderEntries) {
    $productCode = $entry.PSChildName

    try {
        # Use the MSIExec command to uninstall the product
        Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $productCode /qn" -Wait -PassThru

        Write-Log ("Adobe Acrobat Reader has been successfully uninstalled using product code: $productCode")
        Write-host ("Adobe Acrobat Reader has been successfully uninstalled using product code: $productCode")
        Exit 0
    } catch {
        Write-Log ("Failed to uninstall Adobe Acrobat Reader with product code $productCode. Error: $_")
        Exit 1
    }
}