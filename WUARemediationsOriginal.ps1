<#
.SYNOPSIS
    Performs CBS component repair to fix all types of Windows Update Agent corruptions using dism
.DESCRIPTION
    This script automates the corruption detection and remediations of a CBS and Windows Update Agent Component Store using the Deployment Image Servicing and Management (DISM) tool.
    It logs all activities, including any errors encountered during the remediations process, to a log file for troubleshooting purposes.
.NOTES
    WindowsUpdateAgentRemediation.ps1 - V.Ashodhiya - 07/11/2024
    Script History:
    Version 1.0 - Script inception
#>
#---------------------------------------------------------------------#
# Define the path for the log file
$logFilePath = "C:\Windows\fndr\logs"
$logFileName = "$logFilePath\WUARemediations.log"
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
 
# Define the function to check free disk space
function Get-FreeDiskSpace {
    param (
        [int]$thresholdGB = 5  # Default threshold set to 5 GB
    )
    # Get the volume information for the C: drive
    $volume = Get-Volume -DriveLetter C
    $global:freeDiskSpace = 1  # Initialize variable to 1 (condition not met)
    # Check if the free space on C: is greater than the threshold
    if ($volume.SizeRemaining -gt ($thresholdGB * 1GB)) {
        $global:freeDiskSpace = 0  # Condition met (more than threshold GB of free space)
    }
}
# Define the cleanup function (leave blank for user to populate)
function Start-Cleanup {
    # Add your cleanup code here
    Write-Log "Performing cleanup"
    Write-Log "Stopping Windows Update Services"
    Stop-Service -Name BITS | Out-Null
    Stop-Service -Name wuauserv | Out-Null
    Stop-Service -Name appidsvc | Out-Null
    Stop-Service -Name cryptsvc | Out-Null
    Write-Log "Remove QMGR Data file"
    Remove-Item "$env:allusersprofile\Application Data\Microsoft\Network\Downloader\qmgr*.dat" -ErrorAction SilentlyContinue 
    Write-Log "Removing the Software Distribution and CatRoot Folder"
    Remove-Item $env:systemroot\SoftwareDistribution\DataStore -ErrorAction SilentlyContinue
    Remove-Item $env:systemroot\SoftwareDistribution\Download -ErrorAction SilentlyContinue 
    Remove-Item $env:systemroot\System32\Catroot2 -ErrorAction SilentlyContinue 
    Write-Log "Removing old Windows Update log"
    Remove-Item $env:systemroot\WindowsUpdate.log -ErrorAction SilentlyContinue 
    Write-Log "Resetting the Windows Update Services to default settings"
    Start-Process -FilePath "$env:systemroot\system32\sc.exe" -ArgumentList "sdset bits D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)" -Wait
    Start-Process -FilePath "$env:systemroot\system32\sc.exe" -ArgumentList "sdset wuauserv D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)" -Wait

    Write-Log "Removing Windows Temp File"
    Get-ChildItem -Path C:\windows\Temp -File | Remove-Item -Verbose -Force
 
    Set-Location $env:systemroot\system32 
    Write-Log "Registering some DLLs"
    regsvr32.exe /s atl.dll 
    regsvr32.exe /s urlmon.dll 
    regsvr32.exe /s mshtml.dll 
    regsvr32.exe /s shdocvw.dll 
    regsvr32.exe /s browseui.dll 
    regsvr32.exe /s jscript.dll 
    regsvr32.exe /s vbscript.dll 
    regsvr32.exe /s scrrun.dll 
    regsvr32.exe /s msxml.dll 
    regsvr32.exe /s msxml3.dll 
    regsvr32.exe /s msxml6.dll 
    regsvr32.exe /s actxprxy.dll 
    regsvr32.exe /s softpub.dll 
    regsvr32.exe /s wintrust.dll 
    regsvr32.exe /s dssenh.dll 
    regsvr32.exe /s rsaenh.dll 
    regsvr32.exe /s gpkcsp.dll 
    regsvr32.exe /s sccbase.dll 
    regsvr32.exe /s slbcsp.dll 
    regsvr32.exe /s cryptdlg.dll 
    regsvr32.exe /s oleaut32.dll 
    regsvr32.exe /s ole32.dll 
    regsvr32.exe /s shell32.dll 
    regsvr32.exe /s initpki.dll 
    regsvr32.exe /s wuapi.dll 
    regsvr32.exe /s wuaueng.dll 
    regsvr32.exe /s wuaueng1.dll 
    regsvr32.exe /s wucltui.dll 
    regsvr32.exe /s wups.dll 
    regsvr32.exe /s wups2.dll 
    regsvr32.exe /s wuweb.dll 
    regsvr32.exe /s qmgr.dll 
    regsvr32.exe /s qmgrprxy.dll 
    regsvr32.exe /s wucltux.dll 
    regsvr32.exe /s muweb.dll 
    regsvr32.exe /s wuwebv.dll 
    Write-Log "Resetting the WinSock"
    netsh winsock reset | Out-Null
    netsh winhttp reset proxy  | Out-Null
    Write-Log "Delete all BITS jobs"
    Get-BitsTransfer | Remove-BitsTransfer 
    Write-Log "Starting Windows Update Services"
    Start-Service -Name BITS | Out-Null
    Start-Service -Name wuauserv | Out-Null
    Start-Service -Name appidsvc | Out-Null
    Start-Service -Name cryptsvc | Out-Null
}
# Main script
Get-FreeDiskSpace  # Initial check
# If free space is below the threshold, attempt cleanup
if ($global:freeDiskSpace -eq 1) {
    Write-Log "Disk Space is less than 5GB Performing Cleanup."
    Start-Cleanup # Execute the cleanup function
    Write-Log "Checking if disk space requirements are now met."
    Get-FreeDiskSpace  # Recheck the free disk space after cleanup
    # If free space is still below threshold, throw an error and stop execution
    if ($global:freeDiskSpace -eq 1) {
        Write-Log "Cleanup did not free up enough space stopping execution." -ErrorAction Stop
    } else {
        Write-Log "Cleanup successful. Sufficient free disk space available."
    }
} else {
    Write-Log "Sufficient free disk space available. No cleanup needed."
}
 
# Step 1: Get the current working directory.

$sourcePath = Get-Location

# Step 2: Check for any .MSU files in current directory.
$MSUFile = Get-ChildItem -Path $sourcePath -Filter "*.msu" | Select-Object -First 1

if ($MSUFile) {
    Write-Log "Found .MSU file: $($MSUFile.Name)"
    # Step 3: Run the Extract-MSUAndCAB.ps1 script with -filePath and -destinationPath parameters
    $scriptPath = "$sourcePath\Extract-MSUAndCABOriginal.ps1"
    $extractedPath = "C:\Temp\Sources"
    # Ensure the Sources folder exists
    if (!(Test-Path -Path $extractedPath)) {
        New-Item -ItemType Directory -Path $extractedPath | Out-Null
    }

    Write-Log "Running Extract-MSUAndCAB.ps1 script to extract $($MSUFile.Name)"
    & $scriptPath -filePath $MSUFile.FullName -destinationPath $extractedPath
    
    # Step 4: Wait until the extraction script has finished executing
    do {
        Start-Sleep -Seconds 1
    } while (!(Get-Process | Where-Object { $_.Path -eq $scriptPath }) -eq $null)

    # Step 5: Check return code from Extract-MSUAndCAB.ps1 for success or failure. WIP keep working from here
        Write-Log "Extraction complete."
        Write-Log "Running DISM command to fix the corruption of WUA components and store"
        Dism.exe /Online /Cleanup-Image /RestoreHealth /Source:$extractedPath /LimitAccess
            if ($?){
                Write-Log "Successfully Completed DISM based repair of WUA Components and Store"
            }
            else {
                Write-Log "DISM based repair of WUA Components failed." -ErrorAction Stop
            }
    # Step 7: Run DISM /ScanHealth command to scan for issues
    Write-Log "Running DISM command to verify component store health post corruption fix applied."
    dism /Online /Cleanup-Image /ScanHealth
    Write-Log "DISM ScanHealth operation completed."
} else {
    Write-Log "No .MSU files found in $destinationPath."
}
# Step 7: Cleanup - Remove all files and folders inside C:\Temp
Write-Log "Cleaning up C:\Temp to free up disk space..."
Remove-Item -Path "C:\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue
Write-Log "Cleanup of C:\Temp completed."