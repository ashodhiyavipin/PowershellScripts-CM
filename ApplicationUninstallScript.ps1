<#
.SYNOPSIS
Uninstalls applications via WMI and Registry automatically.
.DESCRIPTION
This script automates the removal of applications from a machine completely and silently with no user input and no reboot. 
.NOTES
ApplicationUninstallScript.ps1 - V.Ashodhiya - 13-02-2025
Script History:
Version 1.0 - Script inception
Version 1.1 - Added Logging function.
#>
#---------------------------------------------------------------------#
# Define the path for the log file
$logFilePath = "C:\Windows\fndr\logs"
$logFileName = "$logFilePath\ApplicationUninstall.log"
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
# Function to uninstall applications via Win32_Product (WMI)

function Uninstall-WmiApp {
    param (
        [string]$appName
    )
    Write-Log "Checking for $appName via WMI..."
    Write-Host "Checking for $appName via WMI..."
    $installedApps = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE Name LIKE '%$appName%'"
    if ($installedApps.Count -eq 0) {
        Write-Log "$appName not found using WMI."
        Write-Host "$appName not found using WMI."
        return $false
    } else {
        foreach ($app in $installedApps) {
            Write-Log "Uninstalling $($app.Name) via WMI..."
            Write-Host "Uninstalling $($app.Name) via WMI..."
            $app.Uninstall() | Out-Null
            Write-Log "$($app.Name) has been uninstalled using WMI."
            Write-Host "$($app.Name) has been uninstalled using WMI."
        }
        return $true
    }
}

 # Function to uninstall applications via registry

 function Uninstall-RegistryApp {
    param (
        [string]$appName
    )
    $registryPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $appFound = $false
    foreach ($regPath in $registryPaths) {
        $apps = Get-ItemProperty $regPath | Where-Object { $_.DisplayName -like "*$appName*" }
        foreach ($app in $apps) {
            $appFound = $true
            Write-Log "Uninstalling $($app.DisplayName) from the registry..."
            Write-Host "Uninstalling $($app.DisplayName) from the registry..." 
            if ($app.UninstallString) {
                # Remove leading and trailing double quotes from the UninstallString if present
                $uninstallCmd = $app.UninstallString.Trim('"')
                & "$uninstallCmd" /S
                Write-Log "$($app.DisplayName) has been uninstalled using the registry."
                Write-Host "$($app.DisplayName) has been uninstalled using the registry."
            } else {
                Write-Log "Uninstall string not found for $($app.DisplayName)."
            }
        }
    }

    if (-not $appFound) {
        Write-Log "$appName not found in the registry."
        Write-Host "$appName not found in the registry."
    }
    return $appFound
 }
 
 # Main uninstall function combining both methods
 function Uninstall-Application {
    param (
        [string]$appName
    )
    $uninstalled = $false
    # Try to uninstall via WMI first
    $uninstalled = Uninstall-WmiApp $appName
    # If not found via WMI, try via the registry
    if (-not $uninstalled) {
        Uninstall-RegistryApp $appName
    }
 }
 
 #Uninstall ApplicationName
 Uninstall-Application "Avaya Agent" 