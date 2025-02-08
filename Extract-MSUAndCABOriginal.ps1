param (
    [Parameter(Mandatory = $true)]
    [string]$filePath,
    [Parameter(Mandatory = $true)]
    [string]$destinationPath
)

# Display the note to the user
Write-Host "==========================="
Write-Host
Write-Host -ForegroundColor Yellow "Note: Do not close any Windows opened by this script until it is completed."
Write-Host
Write-Host "==========================="
Write-Host


# Remove quotes if present
$filePath = $filePath -replace '"', ''
$destinationPath = $destinationPath -replace '"', ''

# Trim trailing backslash if present
$destinationPath = $destinationPath.TrimEnd('\')

if (-not (Test-Path $filePath -PathType Leaf)) {
    Write-Host "The specified file does not exist: $filePath"
    return
}

if (-not (Test-Path $destinationPath -PathType Container)) {
    Write-Host "Creating destination directory: $destinationPath"
    New-Item -Path $destinationPath -ItemType Directory | Out-Null
}

$processedFiles = @{}

function Extract-File ($file, $destination) {
    Write-Host "Extracting $file to $destination"
    Start-Process -FilePath "cmd.exe" -ArgumentList "/c expand.exe `"$file`" -f:* `"$destination`" > nul 2>&1" -Wait -WindowStyle Hidden | Out-Null
    $processedFiles[$file] = $true
    Write-Host "Extraction completed for $file"
}

function Rename-File ($file) {
    if (Test-Path -Path $file) {
        $newName = [System.IO.Path]::GetFileNameWithoutExtension($file) + "_" + [System.Guid]::NewGuid().ToString("N") + [System.IO.Path]::GetExtension($file)
        $newPath = Join-Path -Path ([System.IO.Path]::GetDirectoryName($file)) -ChildPath $newName
        Write-Host "Renaming $file to $newPath"
        Rename-Item -Path $file -NewName $newPath
        Write-Host "Renamed $file to $newPath"
        return $newPath
    }
    Write-Host "File $file does not exist for renaming"
    return $null
}

function Process-CabFiles ($directory) {
    while ($true) {
        $cabFiles = Get-ChildItem -Path $directory -Filter "*.cab" -File | Where-Object { -not $processedFiles[$_.FullName] -and $_.Name -ne "wsusscan.cab" }

        if ($cabFiles.Count -eq 0) {
            Write-Host "No more CAB files found in $directory"
            break
        }

        foreach ($cabFile in $cabFiles) {
            Write-Host "Processing CAB file $($cabFile.FullName)"
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
        Write-Host "Extracting .msu file to: $destinationPath"
        Extract-File -file $filePath -destination $destinationPath
    } elseif ($filePath.EndsWith(".cab")) {
        Write-Host "Extracting .cab file to: $destinationPath"
        Extract-File -file $filePath -destination $destinationPath
    } else {
        Write-Host "The specified file is not a .msu or .cab file: $filePath"
        return
    }

    # Process all .cab files recursively
    Write-Host "Starting to process CAB files in $destinationPath"
    Process-CabFiles -directory $destinationPath
}
catch {
    Write-Host "An error occurred while extracting the file. Error: $_"
    return
}

Write-Host "Extraction completed. Files are located in $destinationPath"
return $destinationPath