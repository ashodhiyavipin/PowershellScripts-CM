<#
.SYNOPSIS
    Silently uninstalls applications by querying the registry and executing their uninstall commands, with detailed logging.

.DESCRIPTION
    This script automates the silent removal of applications from Windows systems by searching common uninstall registry paths and executing any available QuietUninstallString or UninstallString entries. 
    It runs without user interaction, suppresses popup windows, and avoids triggering MSI self-healing (no use of Win32_Product). 

    The script logs all actions, including uninstall success/failure, to a persistent log file. A summary report is also included at the end of each execution. It supports both interactive use and parameterized automation.

.PARAMETER AppName
    The name of the application to search for and uninstall. If not supplied, the user will be prompted.

.EXAMPLE
    .\ApplicationUninstallScript.ps1 -AppName "Avaya Agent"

.NOTES
    File Name:    ApplicationUninstallScript.ps1
    Author:       V.Ashodhiya | Daniel Fletcher / Updated by A.Fletcher
    Log Path:     C:\Windows\fndr\logs\ApplicationUninstall.log

    Script History:
    Version 1.0 - Initial script for basic app removal via registry
    Version 1.1 - Added logging function and log file creation
    Version 1.2 - Removed WMI dependency, improved error handling, added verbose logs
    Version 1.3 - Summary reporting, support for param/prompt input, improved path handling, suppressed CMD popup window, and enhanced exception diagnostics, updated header section.
    Version 1.4 - Added /qn switch for silent uninstall of applications. 
    Version 1.5 - Added logic to change /I to /X where some applications have it in the UninstallString Registry Key.
    Version 1.6 - Added logic to suppress automatic reboots during application uninstallation via msiexec.
#>

param (
    [string]$AppName = $null
)

#---------------------------------------------------------------------#
$logFilePath = "C:\Windows\fndr\logs"
$logFileName = "$logFilePath\ApplicationUninstall.log"

# Summary tracking
$Summary = @{
    Found = 0
    Success = 0
    Failed = 0
    NotFound = $true
}

function Write-Log {
    param ([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $message"
    Write-Output $logMessage

    if (-not (Test-Path $logFilePath)) {
        New-Item -Path $logFilePath -ItemType Directory | Out-Null
    }

    Add-Content -Path $logFileName -Value $logMessage
}

function Uninstall-RegistryApp {
    param ([string]$appName)

    $registryPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $appFound = $false

    foreach ($regPath in $registryPaths) {
        try {
            $apps = Get-ItemProperty $regPath -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*$appName*" }
            foreach ($app in $apps) {
                $appFound = $true
                $Summary.Found++
                $Summary.NotFound = $false

                Write-Log "Uninstalling $($app.DisplayName) from registry..."

                $uninstallCmd = $null
                if ($app.QuietUninstallString) {
                    $uninstallCmd = $app.QuietUninstallString
                    Write-Log "Using QuietUninstallString."
                } elseif ($app.UninstallString) {
                    $uninstallCmd = $app.UninstallString
                    Write-Log "Using UninstallString."
                }

                if ($uninstallCmd) {
                    #Check if mseiexec.exe has /I and replace it with /X
                    if ($uninstallCmd -match "^MsiExec\.exe\s*/I\s*\{[A-Z0-9-]+\}") {
                        # Replace /I with /X and keep the product code
                        Write-Log "Registry uninstall string has /I changing it to /X"
                        $uninstallCmd = $uninstallCmd.Replace("/I", "/X")
                    }
                    # Check if the uninstall command is using MsiExec and append /qn if it is
                    if ($uninstallCmd -like "MsiExec.exe*") {
                        $uninstallCmd += " /qn /norestart"
                        Write-Log "Modified uninstall command for silent uninstallation: $uninstallCmd"
                    }

                    Write-Log "Attempting to run uninstall command as-is: $uninstallCmd"

                    try {
                        Start-Process -FilePath "cmd.exe" `
                                      -ArgumentList "/c", $uninstallCmd `
                                      -Wait `
                                      -WindowStyle Hidden `
                                      -ErrorAction Stop

                        Write-Log "$($app.DisplayName) successfully uninstalled."
                        $Summary.Success++
                    } catch {
                        Write-Log "Error running uninstall command:"
                        Write-Log "  Command       : $uninstallCmd"
                        Write-Log "  Exception     : $_"
                        if ($_.Exception.InnerException) {
                            Write-Log "  InnerException: $($_.Exception.InnerException.Message)"
                        }
                        $Summary.Failed++
                    }
                } else {
                    Write-Log "No uninstall command found for $($app.DisplayName)."
                    $Summary.Failed++
                }
            }
        } catch {
            Write-Log "Error reading registry path $regPath`: $_"
        }
    }

    return $appFound
}

function Uninstall-Application {
    param ([string]$appName)
    Write-Log "----- Starting uninstall for $appName -----"
    $result = Uninstall-RegistryApp $appName

    if (-not $result) {
        Write-Log "$appName not found in registry."
    }

    Write-Log "----- Finished uninstall for $appName -----"
    Write-Log "Summary:"
    Write-Log "Apps Found: $($Summary.Found)"
    Write-Log "Uninstalled Successfully: $($Summary.Success)"
    Write-Log "Failed Uninstalls: $($Summary.Failed)"
    if ($Summary.NotFound) {
        Write-Log "No matching applications were found."
    }
}

# Main Script Execution
Write-Log "Script Version 1.4"
if (-not $AppName) {
    $AppName = Read-Host "Enter the name of the application to uninstall"
}

try {
    Uninstall-Application $AppName
} catch {
    Write-Log "Fatal error during uninstall: $_"
}