<#
.SYNOPSIS
Performs CBS component repair to fix all types of Windows Update Agent corruptions using DISM.

.DESCRIPTION
    This script automates the corruption detection and remediations of a CBS and Windows Update Agent Component Store using the Deployment Image Servicing and Management (DISM) tool.
    It logs all activities, including any errors encountered during the remediations process, to a log file for troubleshooting purposes.

.NOTES
    WindowsUpdateAgentRemediation.ps1 - V.Ashodhiya - 07/11/2024
    Script History:
    Version 1.0 - Script inception.
    Version 1.1 - .MSU and .CAB extraction functions added.
    Version 1.2 - Error handling and return codes added.
#>
# Define the path for the log file
$logFilePath = "C:\Windows\fndr\logs"
$logFileName = "$logFilePath\WUARemediations.log"
$sourcePath = Get-Location

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
    $freeSpace = Get-Volume -DriveLetter C | Select-Object -Property SizeRemaining
    Write-Log = "Current Free Space in C drive is $freeSpace"
    $global:freeDiskSpace = 1  # Initialize variable to 1 (condition not met)
 
    # Check if the free space on C: is greater than the threshold
    if ($volume.SizeRemaining -gt ($thresholdGB * 1GB)) {
        $global:freeDiskSpace = 0  # Condition met (more than threshold GB of free space)
    }
}
 
# Define the cleanup function (leave blank for user to populate)
function Start-Cleanup {
    Write-Log "Free Space is less than 5 GB hence Performing cleanup"
    
    Write-Log "Stopping Windows Update Services"
    Stop-Service -Name BITS | Out-Null
    Stop-Service -Name wuauserv | Out-Null
    Stop-Service -Name appidsvc | Out-Null
    Stop-Service -Name cryptsvc | Out-Null
 
    Write-Log "Remove QMGR Data file"
    Remove-Item "$env:allusersprofile\Application Data\Microsoft\Network\Downloader\qmgr*.dat" -ErrorAction SilentlyContinue 
 
    Write-Log "Removing the Software Distribution and CatRoot Folder"
    Remove-Item $env:systemroot\SoftwareDistribution\DataStore -Recurse -ErrorAction SilentlyContinue
    Remove-Item $env:systemroot\SoftwareDistribution\Download -Recurse -ErrorAction SilentlyContinue 
    Remove-Item $env:systemroot\System32\Catroot2 -Recurse -ErrorAction SilentlyContinue 
 
    Write-Log "Removing old Windows Update log"
    Remove-Item $env:systemroot\WindowsUpdate.log -ErrorAction SilentlyContinue 
 
    Write-Log "Resetting the Windows Update Services to default settings"
    Start-Process -FilePath "$env:systemroot\system32\sc.exe" -ArgumentList "sdset bits D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)" -Wait
    Start-Process -FilePath "$env:systemroot\system32\sc.exe" -ArgumentList "sdset wuauserv D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)" -Wait

    Write-Log "Removing Windows Temp File"
    Get-ChildItem -Path C:\windows\Temp -File | Remove-Item -Verbose -Force
   
    Write-Log "Resetting the WinSock"
    netsh winsock reset | Out-Null
    netsh winhttp reset proxy | Out-Null
 
    Write-Log "Delete all BITS jobs"
    Get-BitsTransfer | Remove-BitsTransfer
 
    Write-Log "Starting Windows Update Services"
    Start-Service -Name BITS | Out-Null
    Start-Service -Name wuauserv | Out-Null
    Start-Service -Name appidsvc | Out-Null
    Start-Service -Name cryptsvc | Out-Null
}
 
# Main script flow
Get-FreeDiskSpace  # Initial check

# If free space is below the threshold, attempt cleanup
if ($global:freeDiskSpace -eq 1) {
    Write-Log "Disk Space Requirements not met. Starting Disk Space Cleanup."
    Start-Cleanup  # Execute the cleanup function
    Get-FreeDiskSpace  # Recheck the free disk space after cleanup

    # If free space is still below threshold, throw an error and stop execution
    if ($global:freeDiskSpace -eq 1) {
        Write-Log "Cleanup did not free up enough space to continue Stopping Execution." -ErrorAction Stop
    }
    else {
        Write-Log "Cleanup successful. Sufficient free disk space available."
    }
}
else {
    Write-Log "Sufficient free disk space available. No cleanup needed."
}

# Ensure C:\Temp exists, if not create it
if (!(Test-Path -Path $extractionPath)) {
    New-Item -ItemType Directory -Path $extractionPath | Out-Null
}

$MSUFile = Get-ChildItem -Path (Get-Location) -Filter "*.msu" | Select-Object -First 1
if ($MSUFile) {
    Write-Log "Found .MSU file: $($MSUFile.Name)"
    # Remove quotes if present
    $filePath = $MSUFile -replace '"', '' #Dubious not sure if it works run and verify
    $destinationPath = "C:\Temp\" -replace '"', '' #Dubious not sure if it works run and verify

    # Trim trailing backslash if present
    $destinationPath = $destinationPath.TrimEnd('\')

    if (-not (Test-Path $filePath -PathType Leaf)) {
        Write-Log "The specified file does not exist: $filePath"
        return
    }
    if (-not (Test-Path $destinationPath -PathType Container)) {
        Write-Log "Creating destination directory: $destinationPath"
        New-Item -Path $destinationPath -ItemType Directory | Out-Null
    }

    $processedFiles = @{}
    function Read-File ($file, $destination) {
        Write-Log "Extracting $file to $destination"
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c expand.exe `"$file`" -f:* `"$destination`" > nul 2>&1" -Wait -WindowStyle Hidden | Out-Null
        $processedFiles[$file] = $true
        Write-Log "Extraction completed for $file"
    }
    function Rename-File ($file) {
        if (Test-Path -Path $file) {
            $newName = [System.IO.Path]::GetFileNameWithoutExtension($file) + "_" + [System.Guid]::NewGuid().ToString("N") + [System.IO.Path]::GetExtension($file)
            $newPath = Join-Path -Path ([System.IO.Path]::GetDirectoryName($file)) -ChildPath $newName
            Write-Log "Renaming $file to $newPath"
            Rename-Item -Path $file -NewName $newPath
            Write-Log "Renamed $file to $newPath"
            return $newPath
        }
        Write-Log "File $file does not exist for renaming"
        return $null
    }
    function Read-CabFiles ($directory) {
            
        while ($true) {
            $cabFiles = Get-ChildItem -Path $directory -Filter "*.cab" -File | Where-Object { -not $processedFiles[$_.FullName] -and $_.Name -ne "wsusscan.cab" }

            if ($cabFiles.Count -eq 0) {
                Write-Log "No more CAB files found in $directory"
                break
            }

            foreach ($cabFile in $cabFiles) {
                Write-Log "Processing CAB file $($cabFile.FullName)"
                $cabFilePath = Rename-File -file $cabFile.FullName

                if ($cabFilePath -ne $null) {
                    Read-File -file $cabFilePath -destination $directory
                    Read-CabFiles -directory $directory
                }
            }
        }
    }

    try {
        # Initial extraction
        if ($filePath.EndsWith(".msu")) {
            Write-Log "Extracting .msu file to: $destinationPath"
            Read-File -file $filePath -destination $destinationPath
        }
        elseif ($filePath.EndsWith(".cab")) {
            Write-Log "Extracting .cab file to: $destinationPath"
            Read-File -file $filePath -destination $destinationPath
        }
        else {
            Write-Log "Could not find any .msu or .cab files to extract from: $filePath"
            return
        }

        # Process all .cab files recursively
        Write-Log "Starting to process CAB files in $destinationPath"
        Read-CabFiles -directory $destinationPath
    }
    catch {
        Write-Log "An error occurred while extracting the file. Error: $_" -ErrorAction Stop
        return
    }
    Write-Log "Extraction completed. Files are located in $destinationPath"
    return $destinationPath

    Write-Log "Starting DISM command to fix corruptions of WUA components and store."
    Dism.exe /Online /Cleanup-Image /RestoreHealth /Source:$destinationPath /LimitAccess
        if ($?) {
            Write-Log "Successfully completed DISM based repair of WUA Components and Store."
        }
        else {
            Write-Log "DISM based repair of WUA Components failed please check dism.log for further details" -ErrorAction Stop
        }
    
    # Step 7: Run DISM /ScanHealth command to scan for issues
    Write-Log "Running DISM command to verify component store health post corruption fix applied."
    Dism.exe /Online /Cleanup-Image /ScanHealth
    if ($?) {
        Write-Log "Dism Health is checked and no issues were found."
    } else {
        Write-Log "Dism Health check failed please check dism.log for further details."
    }
}
else {
    Write-Log "No .MSU files found in $sourcePath."
}

# Step 7: Cleanup - Remove all files and folders inside C:\Temp
Write-Log "Cleaning up C:\Temp to free up disk space..."
Remove-Item -Path "$extractionPath\*" -Recurse -Force -ErrorAction SilentlyContinue
Write-Log "Cleanup of C:\Temp completed."