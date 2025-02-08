<#
.SYNOPSIS
Extracts the .MSU and .CAB files for fixing corruptions in the Windows Update Agent and Component Store.

.DESCRIPTION
    This script automates the extraction of .MSU and .CAB files for fixing corruptions in the Windows Update Agent and Component Store.
    It logs all activities, including any errors encountered during the remediations process, to a log file for troubleshooting purposes.

.NOTES
    Extract-MSUAndCAB.ps1 - V.Ashodhiya - 08/11/2024
    Script History:
    Version 1.0 - Script inception.
#>
#---------------------------------------------------------------------#
# Define the path for the log file
param (
    [Parameter(Mandatory = $true)]
    [string]$filePath,
    [Parameter(Mandatory = $true)]
    [string]$destinationPath
)
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

# Display the note to the user
Write-Log "==========================="
Write-Log
Write-Log -ForegroundColor Yellow "Note: Do not close any Windows opened by this script until it is completed."
Write-Log
Write-Log "==========================="
Write-Log


# Remove quotes if present
$filePath = $filePath -replace '"', ''
$destinationPath = $destinationPath -replace '"', ''

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

function Extract-File ($file, $destination) {
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

function Process-CabFiles ($directory) {
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
                Extract-File -file $cabFilePath -destination $directory
                Process-CabFiles -directory $directory
            }
        }
    }
}

try {
    # Initial extraction
    if ($filePath.EndsWith(".msu")) {
        Write-Log "Extracting .msu file to: $destinationPath"
        Extract-File -file $filePath -destination $destinationPath
    } elseif ($filePath.EndsWith(".cab")) {
        Write-Log "Extracting .cab file to: $destinationPath"
        Extract-File -file $filePath -destination $destinationPath
    } else {
        Write-Log "The specified file is not a .msu or .cab file: $filePath"
        return
    }

    # Process all .cab files recursively
    Write-Log "Starting to process CAB files in $destinationPath"
    Process-CabFiles -directory $destinationPath
}
catch {
    Write-Log "An error occurred while extracting the file. Error: $_"
    return $extractionReturnCode = 1
}
    Write-Log "Extraction completed. Files are located in $destinationPath"
    return $destinationPath
    return $extractionReturnCode = 0
