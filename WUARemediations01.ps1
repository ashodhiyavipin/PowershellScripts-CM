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
    Version 2.0 - Modified Script to become function based.
#>
#---------------------------------------------------------------------#
# Define the path for the log file
$logFilePath = "C:\Windows\fndr\logs"
$logFileName = "$logFilePath\WUARemediations.log"
#1. Function to write logs
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
# 2. Function to check free disk space
function Get-FreeDiskSpace {
    param (
        [int]$thresholdGB = 5  # Default threshold set to 5 GB
    )
    # Get the volume information for the C: drive
    Write-Log "Checking the free space on C drive of this machine."
    $volume = Get-Volume -DriveLetter C
    $global:freeDiskSpace = 1  # Initialize variable to 1 (condition not met)
    # Check if the free space on C: is greater than the threshold
    if ($volume.SizeRemaining -gt ($thresholdGB * 1GB)) {
        $global:freeDiskSpace = 0  # Condition met (more than threshold GB of free space)
    }
}

# 3. Function to perform cleanup
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

# 4. Function to get the latest cumulative update number
function Get-LatestCumulativeUpdateNumber {
    Write-Log "Getting the last installed KB article installed on this machine."
    $baseKBNumber = (Get-WmiObject -Query "SELECT * FROM Win32_QuickFixEngineering" | Sort-Object InstalledOn -Descending | Select-Object -First 1).HotFixID
    # Search for the .msu file with the KB number
    Write-Log "Last KB article installed on the PC is $baseKBNumber"
    $msuFile = Get-ChildItem -Path (Get-Location) -Filter "*$baseKBNumber*.msu" -ErrorAction SilentlyContinue
    if ($msuFile) {
        $global:msuFilePath = $msuFile.FullName
        return $global:msuFilePath
    } else {
        Write-Log "No MSU file found for KB: $baseKBNumber"
        return $null
    }
}

# 5. Function to extract MSU and CAB file
function Start-ExtractMSUandCABFile {
    $scriptPath = ".\Extract-MSUAndCABOriginal.ps1"
    $global:extractedPath = "C:\Temp\Sources"
    Write-Log "Starting the Extraction of $global:msuFilePath to $global:extractedPath"
    param (
        [string]$global:msuFilePath,
        [string]$global:extractedPath
    )

    & $scriptPath -filePath $global:msuFilePath -destinationPath $global:extractedPath
    Wait-Process -Name "Extract-MSUAndCABOriginal" -ErrorAction SilentlyContinue
}

# 6. Function to perform DISM repair
function Start-DismRepair {
    Write-Log "Starting DISM command to fix corruptions of WUA components."
    Dism.exe /Online /Cleanup-Image /RestoreHealth /Source:$global:extractedPath /LimitAccess
    if ($?) {
        Write-Log "DISM repair was successful."
    } else {
        Write-Log "DISM repair failed, check dism.log"
    }
}

# 7. Function to run DISM health check
function Start-DismHealthCheck {
    Write-Log "Starting DISM command health scan to verify if the corruptions are resolved."
    Dism /Online /Cleanup-Image /ScanHealth
}

# 8. Cleanup function
function Start-TempCleanup {
    Write-Log "Removing temp folder."
    Remove-Item "C:\Temp\*" -Recurse -Force
    Remove-Item "C:\Temp" -Recurse -Force
}

# 9. Check free disk space
$freeDiskSpace = Get-FreeDiskSpace

# 10. Check and perform cleanup if necessary
if ($freeDiskSpace -eq 1) {
    Write-Log "Starting Cleanup since disk space is less than 5GB"
    Start-Cleanup
    $freeDiskSpace = Get-FreeDiskSpace
    if ($freeDiskSpace -eq 1) {
        throw "Error: Insufficient disk space, execution stopped."
    } else {
        Write-Log "Cleanup successful, sufficient disk space is available."
    }
} else {
    Write-Log "Sufficient disk space is available, no cleanup needed."
}

# 11. Get latest cumulative update number
Get-LatestCumulativeUpdateNumber

# 12. Extract MSU and CAB file
Start-ExtractMSUandCABFile -filePath $global:msuFilePath -extractedPath $global:extractedPath

# 13. Start DISM repair
Start-DismRepair

# 14. Start DISM health check
Start-DismHealthCheck

# 15. Perform Temp folder cleanup.
Start-TempCleanup