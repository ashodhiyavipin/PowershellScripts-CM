# Function to uninstall applications via Win32_Product (WMI)

function Uninstall-WmiApp {
    param (
        [string]$appName
    )
 
    Write-Host "Checking for $appName via WMI..."
 
    $installedApps = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE Name LIKE '%$appName%'"
 
    if ($installedApps.Count -eq 0) {
 
        Write-Host "$appName not found using WMI."
 
        return $false
 
    } else {
 
        foreach ($app in $installedApps) {
 
            Write-Host "Uninstalling $($app.Name) via WMI..."
 
            $app.Uninstall() | Out-Null
 
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
 
            Write-Host "Uninstalling $($app.DisplayName) from the registry..."
 
            if ($app.UninstallString) {
 
                # Remove leading and trailing double quotes from the UninstallString if present
 
                $uninstallCmd = $app.UninstallString.Trim('"')
 
                & "$uninstallCmd" /S
 
                Write-Host "$($app.DisplayName) has been uninstalled using the registry."
 
            } else {
 
                Write-Host "Uninstall string not found for $($app.DisplayName)."
            }
        }
    }
 
    if (-not $appFound) {
 
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