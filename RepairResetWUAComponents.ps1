<# 
.SYNOPSIS 
RepairResetWUAComponents.ps1 - Resets the Windows Update components
.DESCRIPTION  
This script will reset all of the Windows Updates components to DEFAULT SETTINGS. 
.OUTPUTS 
Results are printed to the console.
.NOTES 
https://docs.microsoft.com/en-us/windows/deployment/update/windows-update-resources

    RepairResetWUAComponents.ps1 - V.Ashodhiya - 25/10/2024
    Script History:
    Version 1.0 - Script inception
    Version 1.1 - Added Logging Function.
    Version 1.2 - Fixed Logging function typo. 
                - Added logic to fix system files using dism and sfc.
    Version 1.3 - Added logic to remove pending.xml which could block update. Removed dism and sfc since it causes delays during script running.
    Version 1.4 - Corrected logfile name.
    Version 1.5 - Added better control over Services.
#>
# Define the path for the log file
$logFilePath = "C:\Windows\fndr\logs"
$logFileName = "$logFilePath\RepairWUAComponents.log"
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
 
    # Write Log Write-Logmessage to log file
    Add-Content -Path $logFileName -Value $logMessage
}
# -------------------------
#   Stopping Services
# -------------------------
Write-Log "Stopping Services..."
# Services you want to control
$services = 'BITS', 'wuauserv', 'appidsvc', 'cryptsvc'

# Store original startup types
$original = foreach ($svc in $services) {
    $obj = Get-Service -Name $svc -ErrorAction SilentlyContinue
    if ($obj) {
        [PSCustomObject]@{
            Name      = $svc
            StartType = (Get-WmiObject -Class Win32_Service -Filter "Name='$svc'").StartMode
        }
    }
}

# Disable + Stop
foreach ($svc in $services) {
    Set-Service -Name $svc -StartupType Disabled -ErrorAction SilentlyContinue
    Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
}

Write-Log "Services stopped and locked down."

# -------------------------
#   YOUR MAIN SCRIPT HERE
# -------------------------

# --- Step 1: Remove QMGR Data file ---
Write-Log "Step 1. Removing QMGR data file..."
$qmgrPath = "$env:allusersprofile\Application Data\Microsoft\Network\Downloader\qmgr*.dat"
$qmgrFiles = Get-ChildItem -Path $qmgrPath -ErrorAction SilentlyContinue

if ($qmgrFiles) {
    Write-Log "Step 1. QMGR data file(s) found: $($qmgrFiles.FullName -join ', '). Attempting removal..."
    try {
        Remove-Item $qmgrPath -Force -ErrorAction Stop
        Write-Log "Step 1. QMGR data file removal command executed."
    }
    catch {
        Write-Log "Step 1. ERROR - Failed to remove QMGR data file(s): $_"
    }

    # Post-action verification
    $qmgrFilesPostCheck = Get-ChildItem -Path $qmgrPath -ErrorAction SilentlyContinue
    if (-not $qmgrFilesPostCheck) {
        Write-Log "Step 1. VERIFIED - QMGR data file(s) successfully removed."
    }
    else {
        Write-Log "Step 1. WARNING - QMGR data file(s) still present after removal attempt: $($qmgrFilesPostCheck.FullName -join ', '). Continuing with script."
    }
}
else {
    Write-Log "Step 1. QMGR data file(s) not found at path. No removal needed. Continuing."
}   

 
# --- Step 2: Remove Software Distribution and CatRoot Folders ---
Write-Log "Step 2. Removing the Software Distribution and CatRoot Folders..."

$step2Folders = @(
    "$env:systemroot\SoftwareDistribution\DataStore",
    "$env:systemroot\SoftwareDistribution\Download",
    "$env:systemroot\System32\Catroot2"
)

foreach ($folder in $step2Folders) {
    $folderName = Split-Path $folder -Leaf

    if (Test-Path $folder) {
        Write-Log "Step 2. [$folderName] folder found at '$folder'. Attempting removal..."
        try {
            Remove-Item $folder -Recurse -Force -ErrorAction Stop
            Write-Log "Step 2. [$folderName] removal command executed."
        }
        catch {
            Write-Log "Step 2. ERROR - Failed to remove [$folderName]: $_"
        }

        # Post-action verification
        if (-not (Test-Path $folder)) {
            Write-Log "Step 2. VERIFIED - [$folderName] successfully removed."
        }
        else {
            Write-Log "Step 2. WARNING - [$folderName] still present after removal attempt. Continuing with script."
        }
    }
    else {
        Write-Log "Step 2. [$folderName] not found at '$folder'. No removal needed. Continuing."
    }
}

# --- Step 3: Remove old Windows Update log ---
Write-Log "Step 3. Removing old Windows Update log..."
$wuLogPath = "$env:systemroot\WindowsUpdate.log"

if (Test-Path $wuLogPath) {
    Write-Log "Step 3. Windows Update log found at '$wuLogPath'. Attempting removal..."
    try {
        Remove-Item $wuLogPath -Force -ErrorAction Stop
        Write-Log "Step 3. Windows Update log removal command executed."
    }
    catch {
        Write-Log "Step 3. ERROR - Failed to remove Windows Update log: $_"
    }

    # Post-action verification
    if (-not (Test-Path $wuLogPath)) {
        Write-Log "Step 3. VERIFIED - Windows Update log successfully removed."
    }
    else {
        Write-Log "Step 3. WARNING - Windows Update log still present after removal attempt. Continuing with script."
    }
}
else {
    Write-Log "Step 3. Windows Update log not found at '$wuLogPath'. No removal needed. Continuing."
}

# --- Step 4: Remove Windows Temp Files ---
Write-Log "Step 4. Removing Windows Temp Files..."
$tempPath = "C:\windows\Temp"
$tempFiles = Get-ChildItem -Path $tempPath -File -ErrorAction SilentlyContinue

if ($tempFiles) {
    $fileCount = $tempFiles.Count
    Write-Log "Step 4. Found $fileCount temp file(s) in '$tempPath'. Attempting removal..."
    try {
        $tempFiles | Remove-Item -Force -ErrorAction Stop
        Write-Log "Step 4. Temp file removal command executed."
    }
    catch {
        Write-Log "Step 4. ERROR - Failed to remove some temp file(s): $_"
    }

    # Post-action verification
    $remainingFiles = Get-ChildItem -Path $tempPath -File -ErrorAction SilentlyContinue
    if (-not $remainingFiles) {
        Write-Log "Step 4. VERIFIED - All $fileCount temp file(s) successfully removed."
    }
    else {
        $remainingCount = $remainingFiles.Count
        Write-Log "Step 4. WARNING - $remainingCount temp file(s) still present after removal attempt. Continuing with script."
    }
}
else {
    Write-Log "Step 4. No temp files found in '$tempPath'. No removal needed. Continuing."
}
 
Set-Location $env:systemroot\system32 
 
# --- Step 5: Register DLLs ---
Write-Log "Step 5. Registering DLLs..."

$dlls = @(
    'atl.dll', 'urlmon.dll', 'mshtml.dll', 'shdocvw.dll', 'browseui.dll',
    'jscript.dll', 'vbscript.dll', 'scrrun.dll', 'msxml.dll', 'msxml3.dll',
    'msxml6.dll', 'actxprxy.dll', 'softpub.dll', 'wintrust.dll', 'dssenh.dll',
    'rsaenh.dll', 'gpkcsp.dll', 'sccbase.dll', 'slbcsp.dll', 'cryptdlg.dll',
    'oleaut32.dll', 'ole32.dll', 'shell32.dll', 'initpki.dll', 'wuapi.dll',
    'wuaueng.dll', 'wuaueng1.dll', 'wucltui.dll', 'wups.dll', 'wups2.dll',
    'wuweb.dll', 'qmgr.dll', 'qmgrprxy.dll', 'wucltux.dll', 'muweb.dll',
    'wuwebv.dll'
)

$successCount = 0
$failCount = 0

foreach ($dll in $dlls) {
    Write-Log "Step 5. Registering $dll..."
    try {
        $process = Start-Process -FilePath "regsvr32.exe" -ArgumentList "/s $dll" -Wait -PassThru -ErrorAction Stop
        if ($process.ExitCode -eq 0) {
            Write-Log "Step 5. VERIFIED - $dll registered successfully."
            $successCount++
        }
        else {
            Write-Log "Step 5. WARNING - $dll registration returned exit code $($process.ExitCode). Continuing with script."
            $failCount++
        }
    }
    catch {
        Write-Log "Step 5. ERROR - Failed to register $dll`: $_"
        $failCount++
    }
}

Write-Log "Step 5. DLL registration complete. Success: $successCount, Failed: $failCount out of $($dlls.Count) total."
 
# --- Step 6: Reset WinSock and WinHTTP Proxy ---
Write-Log "Step 6. Resetting the WinSock and WinHTTP Proxy..."

# WinSock reset
Write-Log "Step 6. Attempting WinSock reset..."
try {
    netsh winsock reset 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
        Write-Log "Step 6. VERIFIED - WinSock reset completed successfully."
    }
    else {
        Write-Log "Step 6. WARNING - WinSock reset returned exit code $LASTEXITCODE. Continuing with script."
    }
}
catch {
    Write-Log "Step 6. ERROR - WinSock reset failed: $_"
}

# WinHTTP proxy reset
Write-Log "Step 6. Attempting WinHTTP proxy reset..."
try {
    netsh winhttp reset proxy 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
        Write-Log "Step 6. VERIFIED - WinHTTP proxy reset completed successfully."
    }
    else {
        Write-Log "Step 6. WARNING - WinHTTP proxy reset returned exit code $LASTEXITCODE. Continuing with script."
    }
}
catch {
    Write-Log "Step 6. ERROR - WinHTTP proxy reset failed: $_"
}
 
# --- Step 7: Delete all BITS jobs ---
Write-Log "Step 7. Deleting all BITS jobs..."
$bitsJobs = Get-BitsTransfer -ErrorAction SilentlyContinue

if ($bitsJobs) {
    $jobCount = @($bitsJobs).Count
    Write-Log "Step 7. Found $jobCount BITS job(s). Attempting removal..."
    try {
        $bitsJobs | Remove-BitsTransfer -ErrorAction Stop
        Write-Log "Step 7. BITS job removal command executed."
    }
    catch {
        Write-Log "Step 7. ERROR - Failed to remove BITS job(s): $_"
    }

    # Post-action verification
    $remainingJobs = Get-BitsTransfer -ErrorAction SilentlyContinue
    if (-not $remainingJobs) {
        Write-Log "Step 7. VERIFIED - All BITS job(s) successfully removed."
    }
    else {
        $remainingCount = @($remainingJobs).Count
        Write-Log "Step 7. WARNING - $remainingCount BITS job(s) still present after removal attempt. Continuing with script."
    }
}
else {
    Write-Log "Step 7. No BITS jobs found. No removal needed. Continuing."
}

# --- Step 8: Restore original startup types ---
foreach ($item in $original) {
    Set-Service -Name $item.Name -StartupType $item.StartType -ErrorAction SilentlyContinue
}

Write-Log "Step 8. Startup types restored."

# --- Step 9: Force discovery ---
Write-Log "Step 9. Forcing discovery..." 
wuauclt /resetauthorization /detectnow 

# --- Step 10: Clearing SCCM Cache ---  
Write-Log "Step 10. Clearing SCCM Cache..."
## Initialize the CCM resource manager com object
[__comobject]$CCMComObject = New-Object -ComObject 'UIResource.UIResourceMgr'
## Get the CacheElementIDs to delete
$CacheInfo = $CCMComObject.GetCacheInfo().GetCacheElements()
## Remove cache items
ForEach ($CacheItem in $CacheInfo) {
    $null = $CCMComObject.GetCacheInfo().DeleteCacheElement([string]$($CacheItem.CacheElementID))
}

Write-Log "Step 11. Remove pending.xml if present..."
# Remove pending.xml if present
$pendingXml = "$env:systemroot\WinSxS\pending.xml"
if (Test-Path $pendingXml) {
    Remove-Item $pendingXml -Force
    Write-Log "Removed pending.xml"
}