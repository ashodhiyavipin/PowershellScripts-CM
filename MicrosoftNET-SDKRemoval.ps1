<#
.SYNOPSIS
Uninstalls .NET SDK all versions. 
.DESCRIPTION
This script automates the removal of .NET SDK from a machine completely and silently with no user input and no reboot. 
.NOTES
MicrosoftNET-SDKRemoval.ps1 - V.Ashodhiya - 04-12-2024
Script History:
Version 1.0 - Script inception
#>
#---------------------------------------------------------------------#
# Define the path for the log file
$logFilePath = "C:\Windows\fndr\logs"
$logFileName = "$logFilePath\Net-sdkuninstall.log"
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
# Function to uninstall applications via Win32_Product (WMI)

function Uninstall-WmiApp {
    param (
        [string]$appName
    )
 
    Write-Log "Checking for $appName via WMI..." -Verbose
 
    $installedApps = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE Name LIKE '%$appName%'"
 
    if ($installedApps.Count -eq 0) {
        Write-Log "$appName not found using WMI." -Verbose
        return $false
    } else {
        foreach ($app in $installedApps) {
            try {
                Write-Log "Uninstalling $($app.Name) via WMI..." -Verbose
                $app.Uninstall() | Out-Null
                Write-Log "$($app.Name) has been uninstalled using WMI." -Verbose
            } catch {
                Write-Log "Failed to uninstall $($app.Name) via WMI. Error: $_" -ForegroundColor Red
            }
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
            Write-Log "Uninstalling $($app.DisplayName) from the registry..." -Verbose
            
            # Try to read QuietUninstallString first
            $uninstallCmd = if ($app.QuietUninstallString) {
                $app.QuietUninstallString
            } else {
                $app.UninstallString
            }

            if ($uninstallCmd) {
                # Use regex to extract text within double quotes
                $uninstallCmd = $uninstallCmd -replace '^.*?"(.*?)".*$', '$1'
                
                Write-Log "Using '$uninstallCmd' to remove Application" -Verbose
                try {
                    # Execute the uninstall command
                    Start-Process -FilePath $uninstallCmd -ArgumentList "/uninstall", "/quiet" -Wait -NoNewWindow
                    Write-Log "$($app.DisplayName) has been uninstalled using the registry." -Verbose
                } catch {
                    Write-Log "Failed to uninstall $($app.DisplayName). Error: $_" -ForegroundColor Red
                }
            } else {
                Write-Log "Uninstall string not found for $($app.DisplayName)." -ForegroundColor Yellow
            }
        }
    }

    if (-not $appFound) {
        Write-Log "$appName not found in the registry." -Verbose
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
    $uninstalled = Uninstall-RegistryApp $appName

    # If not found via WMI, try via the registry
    if (-not $uninstalled) {
        Uninstall-WmiApp $appName
    }
}

# Uninstall Microsoft .NET SDK 5.0.202 (x64)
Uninstall-Application "Microsoft .NET SDK 6.0.136 (x64)"

# Uninstall Microsoft .NET SDK 5.0.401 (x64)
Uninstall-Application "Microsoft .NET SDK 6.0.136 (x86)"

Uninstall-Application "Microsoft .NET Runtime - 6.0.36 (x86)"

Uninstall-Application "Microsoft .NET Runtime - 6.0.36 (x64)"

Uninstall-Application "Microsoft Windows Desktop Runtime - 6.0.36 (x86)"

Uninstall-Application "Microsoft Windows Desktop Runtime - 6.0.36 (x64)"

Uninstall-Application "Microsoft ASP.NET Core 6.0.36 - Shared Framework (x86)"

Uninstall-Application "Microsoft ASP.NET Core 6.0.36 - Shared Framework (x64)"