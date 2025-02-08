<# 
.SYNOPSIS 
This PowerShell script automates the uninstallation process of Microsoft Office 16, including the removal of installation folders, scheduled tasks, running processes, associated services, registry entries, Start menu shortcuts, and Click-to-Run components.
 
.DESCRIPTION  
The script performs a comprehensive and systematic uninstallation of Microsoft Office 16 from a Windows system. It follows a series of defined steps to ensure that all components of Office 16 are completely removed, including files, folders, services, scheduled tasks, and registry entries.
Key Features:
1. Remove Installation Folders: The script locates and deletes the primary installation directory of Microsoft Office 16.
2. Delete Scheduled Tasks: It removes any scheduled tasks associated with Microsoft Office to prevent future automated processes from starting.
3. Terminate Running Processes: The script forcefully stops any active Click-to-Run processes related to Office to ensure no files are in use during the uninstallation.
4.Delete Office Service: It removes the ClickToRun service from the system, which is responsible for managing Office installations.
5.Comprehensive File Deletio: The script deletes all leftover files and folders associated with Microsoft Office from the Program Files, ProgramData, and Common Program Files directories.
6.Registry Cleanu: It removes relevant registry entries to ensure no remnants of Office 16 exist in the Windows registry, reducing the chance of conflicts with future software installations.
7.Remove Start Menu Shortcut: The script cleans up Start Menu items related to Microsoft Office, removing shortcuts for Office applications and tools.
8.Uninstall Click-to-Run Component: It executes the necessary MSI commands to uninstall Click-to-Run components, depending on the system's architecture (x86 or x64).

.NOTES 
OfficeUninstallation.ps1 - V.Ashodhiya - 21-01-2025
Script History:
Version 1.0 - Script inception
 
#>
#Logging Function
$logFilePath = "C:\Windows\fndr\logs"
$logFileName = "$logFilePath\OfficeUninstall.log"
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
# Define the description for the restore point
$restorePointDescription = "Office Uninstall RP"

# Create the restore point
Checkpoint-Computer -Description $restorePointDescription -RestorePointType "MODIFY_SETTINGS"

# Output a message indicating the restore point has been created
Write-Log "System restore point '$restorePointDescription' created successfully."

Function Stop-OfficeProcess {
    Write-Log "Stopping running Office applications ..."
    $OfficeProcessesArray = "lync", "winword", "excel", "msaccess", "mstore", "infopath", "setlang", "msouc", "ois", "onenote", "outlook", "powerpnt", "mspub", "groove", "visio", "winproj", "graph", "teams"
    foreach ($ProcessName in $OfficeProcessesArray) {
        if (get-process -Name $ProcessName -ErrorAction SilentlyContinue) {
            if (Stop-Process -Name $ProcessName -Force -ErrorAction SilentlyContinue) {
                Write-Log "Process $ProcessName was stopped."
            }
            else {
                Write-Log "Process $ProcessName could not be stopped."
            }
        } 
    }
}

Stop-OfficeProcess

# Step 0: Uninstall the Office 16 Click-To-Run Licensing Component, Extensibility Component and Localization Component

try {
    # Check for the installed version of Microsoft Office
    $officeKeyPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"

    # Check if Office is installed
    if (Test-Path $officeKeyPath) {
        $officeBitness = (Get-ItemProperty -Path $officeKeyPath -Name "Platform").Platform
        Write-Host "Installed Microsoft Office is: $officeBitness"

        if ($officeBitness -eq "x64") {
            # 64-bit Office installed on a 64-bit OS
            $msiCommands = @(
                "MsiExec.exe /qn /X{90160000-007E-0000-1000-0000000FF1CE}",
                "MsiExec.exe /qn /X{90160000-008C-0000-1000-0000000FF1CE}",
                "MsiExec.exe /qn /X{90160000-008C-0409-1000-0000000FF1CE}"
            )
        } elseif ($officeBitness -eq "x86") {
            # 32-bit Office installed on a 64-bit OS
            $msiCommands = @(
                "MsiExec.exe /qn /X{90160000-008F-0000-1000-0000000FF1CE}",
                "MsiExec.exe /qn /X{90160000-008C-0000-0000-0000000FF1CE}",
                "MsiExec.exe /qn /X{90160000-008C-0409-0000-0000000FF1CE}"
            )
        } elseif ($officeBitness -eq "x86" -and [Environment]::Is64BitOperatingSystem -eq $false) {
            # 32-bit Office installed on a 32-bit OS
            $msiCommands = @(
                "MsiExec.exe /qn /X{90160000-007E-0000-0000-0000000FF1CE}",
                "MsiExec.exe /qn /X{90160000-008C-0000-0000-0000000FF1CE}",
                "MsiExec.exe /qn /X{90160000-008C-0409-0000-0000000FF1CE}"
            )
        } else {
            Write-Host "Unknown Office installation bitness."
            return
        }

        # Execute the MSI commands
        foreach ($command in $msiCommands) {
            Start-Process -FilePath cmd.exe -ArgumentList "/c $command" -Wait
        }

        Write-Host "Office removal commands executed successfully."
    } else {
        Write-Host "Microsoft Office is not installed or the registry path is different."
    }
} catch {
    Write-Host "An error occurred: $_"
}

# Step 1: Remove the Windows Installer packages
$officeFolderPath = "C:\Program Files\Microsoft Office\Office16"
Write-Log "Removing Windows Installer Packages from $officeFolderPath" 
if (Test-Path $officeFolderPath) {
    Write-Log "Found $officeFolderPath"
    Remove-Item -Path $officeFolderPath -Recurse -Force
    Write-Log "Removed Windows Installer Packages from $officeFolderPath"
} else {
    Write-Log "Unable to remove Windows Installed Packages or Windows Installed Packages not found."
}


# Step 2: Remove the Office scheduled tasks
$tasks = @(
    "\Microsoft\Office\Office Automatic Updates",
    "\Microsoft\Office\Office Subscription Maintenance",
    "\Microsoft\Office\Office ClickToRun Service Monitor",
    "\Microsoft\Office\OfficeTelemetryAgentLogOn2016",
    "\Microsoft\Office\OfficeTelemetryAgentFallBack2016"
)
Write-Log "Starting Removal of Office Scheduled Tasks"

try {
    foreach ($task in $tasks) {
        Write-Log "Found Scheduled Tasks now removing them"
        schtasks.exe /delete /tn $task /f
        Write-Log "Scheduled tasks removed."
    }
} catch {
    Write-Log "An error occurred while removing scheduled tasks: $_"
}

# Step 3: Use Task Manager to end the Click-to-Run tasks
$processes = @("OfficeClickToRun.exe", "OfficeC2RClient.exe", "AppVShNotify.exe", "setup*.exe")
Write-Log "Stopping all click-to-run tasks currently running."
try {
    foreach ($process in $processes) {
        Get-Process -Name $process -ErrorAction SilentlyContinue | Stop-Process -Force
        Write-Log "All Click-to-run tasks have been stopped"
    }
} catch {
    Write-Log "An error occurred while stopping processes: $_"
}

# Step 4: Delete the Office service
try {
    # Attempt to delete the ClickToRunSvc service
    sc.exe delete ClickToRunSvc
    Write-Log "Service 'ClickToRunSvc' deleted successfully."
} catch {
    # Handle any errors that occur during the execution
    Write-Log "An error occurred: $_"
}

# Step 5: Delete the Office files
$foldersToDelete = @(
    "C:\Program Files\Microsoft Office",
    "C:\Program Files (x86)\Microsoft Office",
    "C:\Program Files\Common Files\Microsoft Shared\ClickToRun",
    "C:\ProgramData\Microsoft\ClickToRun",
    "C:\ProgramData\Microsoft\Office"
)
Write-Log "Starting Removal of Office Files and Folder from Program Files and ProgramData"
try {
    foreach ($folder in $foldersToDelete) {
        if (Test-Path $folder) {
            Write-Log "Removing Items $folder"
            Remove-Item -Path $folder -Recurse -Force
            Write-Log "$folder Removed"
        }
    }
} catch {
    Write-Log "An error occurred while deleting folders: $_"
}

# Step 6: Delete the Office registry subkeys
$registryKeys = @(
    "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun",
    "HKLM:\SOFTWARE\Microsoft\AppVISV",
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Office *",
    "HKCU:\Software\Microsoft\Office"
)
Write-Log "Starting removal of Residual Office registry keys"
try {
    foreach ($key in $registryKeys) {
        Write-Log "Removing registry key $key"
        Remove-Item -Path $key -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log "$key removed"
    }
    
}
catch {
    Write-Log "An error occurred while Registry Keys: $_"
    
}

# Step 7: Delete the Start menu shortcuts
$startMenuPath = "$env:ALLUSERSPROFILE\Microsoft\Windows\Start Menu\Programs"
$shortcutsToDelete = @("Microsoft Office 2016 Tools", "* 2016")
Write-Log "Starting removal of start menu shortcuts"
try {
    foreach ($item in $shortcutsToDelete) {
        Write-Log "Removing $item"
        $fullPath = Join-Path -Path $startMenuPath -ChildPath $item
        if (Test-Path $fullPath) {
            Remove-Item -Path $fullPath -Recurse -Force
        }
        Write-Log "$item Removed"
    }
    
}
catch {
    Write-Log "An error occurred while deleting : $_"
}
Write-Log "Microsoft Office 16 has been uninstalled successfully."