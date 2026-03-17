#Requires -RunAsAdministrator
<# 
.SYNOPSIS 
RepairResetWUAComponents-New.ps1 - Resets the Windows Update components
.DESCRIPTION  
This script will reset all of the Windows Updates components to DEFAULT SETTINGS. 
.OUTPUTS 
Results are printed to the console and logged to C:\Windows\fndr\logs\RepairWUAComponents.log
.NOTES 
https://docs.microsoft.com/en-us/windows/deployment/update/windows-update-resources

    RepairResetWUAComponents-New.ps1 - V.Ashodhiya - 25/10/2024
    Script History:
    Version 1.0 - Script inception
    Version 1.1 - Added Logging Function.
    Version 1.2 - Fixed Logging function typo. 
                - Added logic to fix system files using dism and sfc.
    Version 1.3 - Added logic to remove pending.xml which could block update. Removed dism and sfc since it causes delays during script running.
    Version 1.4 - Corrected logfile name.
    Version 1.5 - Added better control over Services.
    Version 1.6 - Added error handling and extended logging for all steps.
    Version 1.7 - Fixed netsh/regsvr32 error handling, added post-action verification,
                  made script idempotent, improved SCCM compatibility, restored full DLL list,
                  added CIM StartMode mapping, added admin check.
    Version 1.8 - Removed DLLs not present on modern Windows or not supporting self-registration.
                  Moved WinSock/WinHTTP reset to after service restore for reliability.
                  Fixed SCCM cache empty collection detection.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

# Define the path for the log file
$logFilePath = "C:\Windows\fndr\logs"
$logFileName = "$logFilePath\RepairWUAComponents.log"

# One-time log directory initialization
if (-not (Test-Path $logFilePath)) {
    try {
        New-Item -Path $logFilePath -ItemType Directory -Force | Out-Null
    }
    catch {
        Write-Host "FATAL - Unable to create log directory '$logFilePath': $_" -ForegroundColor Red
        exit 1
    }
}

# Function to write logs
function Write-Log {
    param ([string]$message)

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $message"
    Write-Host $logMessage
    Add-Content -Path $logFileName -Value $logMessage -ErrorAction SilentlyContinue
}

# Map CIM Win32_Service.StartMode values to Set-Service -StartupType values
# CIM returns "Auto" but Set-Service expects "Automatic"
$startTypeMap = @{
    'Auto'     = 'Automatic'
    'Manual'   = 'Manual'
    'Disabled' = 'Disabled'
    'Boot'     = 'Boot'
    'System'   = 'System'
}

# -------------------------
#   Stopping Services
# -------------------------
Write-Log "Stopping Services..."

$services = 'BITS', 'wuauserv', 'appidsvc', 'cryptsvc'

# Store original startup types (modern CIM)
$original = foreach ($svc in $services) {
    $obj = Get-Service -Name $svc -ErrorAction SilentlyContinue
    if ($obj) {
        $svcInfo = Get-CimInstance -ClassName Win32_Service -Filter "Name='$svc'" -ErrorAction SilentlyContinue
        if ($svcInfo) {
            [PSCustomObject]@{
                Name      = $svc
                StartType = $svcInfo.StartMode
            }
        }
    }
}

if (-not $original) {
    Write-Log "WARNING - None of the target services were found on this system."
}

# Disable + Stop
foreach ($svc in $services) {
    Set-Service -Name $svc -StartupType Disabled -ErrorAction SilentlyContinue
    Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
}

Write-Log "Services stopped and locked down."

# -------------------------
#   Step 1: Remove QMGR Data
# -------------------------
Write-Log "Step 1. Removing QMGR data file..."

$qmgrPath = "$env:ProgramData\Microsoft\Network\Downloader"
$qmgrFiles = Get-ChildItem -Path $qmgrPath -Filter "qmgr*.dat" -ErrorAction SilentlyContinue

if ($qmgrFiles) {
    Write-Log "Step 1. QMGR data file(s) found: $($qmgrFiles.FullName -join ', '). Attempting removal..."
    foreach ($file in $qmgrFiles) {
        try {
            Remove-Item $file.FullName -Force -ErrorAction Stop
            Write-Log "Step 1. Removed $($file.Name)."
        }
        catch {
            Write-Log "Step 1. ERROR - Failed to remove $($file.FullName): $_"
        }
    }

    # Post-action verification
    $qmgrFilesPostCheck = Get-ChildItem -Path $qmgrPath -Filter "qmgr*.dat" -ErrorAction SilentlyContinue
    if (-not $qmgrFilesPostCheck) {
        Write-Log "Step 1. VERIFIED - All QMGR data file(s) successfully removed."
    }
    else {
        Write-Log "Step 1. WARNING - QMGR data file(s) still present: $($qmgrFilesPostCheck.FullName -join ', '). Continuing with script."
    }
}
else {
    Write-Log "Step 1. No QMGR files found. No removal needed. Continuing."
}

# -------------------------
# Step 2: Rename SoftwareDistribution & Catroot2
# -------------------------
# Design note: We rename rather than delete to preserve data for post-mortem debugging.
# If a previous .old folder exists, a timestamped suffix is used for idempotency.
Write-Log "Step 2. Resetting SoftwareDistribution and Catroot2..."

$sdFolder = "$env:SystemRoot\SoftwareDistribution"
$catroot2 = "$env:SystemRoot\System32\Catroot2"
$renameTimestamp = Get-Date -Format "yyyyMMddHHmmss"

# SoftwareDistribution
if (Test-Path $sdFolder) {
    $newName = "SoftwareDistribution.old"
    if (Test-Path "$env:SystemRoot\$newName") {
        $newName = "SoftwareDistribution.old.$renameTimestamp"
    }
    Write-Log "Step 2. SoftwareDistribution found. Renaming to '$newName'..."
    try {
        Rename-Item -Path $sdFolder -NewName $newName -ErrorAction Stop
        Write-Log "Step 2. Rename command executed."
    }
    catch {
        Write-Log "Step 2. ERROR renaming SoftwareDistribution: $_"
    }

    # Post-action verification
    if (-not (Test-Path $sdFolder)) {
        Write-Log "Step 2. VERIFIED - SoftwareDistribution renamed successfully."
    }
    else {
        Write-Log "Step 2. WARNING - SoftwareDistribution still present after rename attempt. Continuing with script."
    }
}
else {
    Write-Log "Step 2. SoftwareDistribution not found at '$sdFolder'. No action needed. Continuing."
}

# Catroot2
if (Test-Path $catroot2) {
    $newCatName = "Catroot2.old"
    if (Test-Path "$env:SystemRoot\System32\$newCatName") {
        $newCatName = "Catroot2.old.$renameTimestamp"
    }
    Write-Log "Step 2. Catroot2 found. Renaming to '$newCatName'..."
    try {
        Rename-Item -Path $catroot2 -NewName $newCatName -ErrorAction Stop
        Write-Log "Step 2. Rename command executed."
    }
    catch {
        Write-Log "Step 2. ERROR renaming Catroot2: $_"
    }

    # Post-action verification
    if (-not (Test-Path $catroot2)) {
        Write-Log "Step 2. VERIFIED - Catroot2 renamed successfully."
    }
    else {
        Write-Log "Step 2. WARNING - Catroot2 still present after rename attempt. Continuing with script."
    }
}
else {
    Write-Log "Step 2. Catroot2 not found at '$catroot2'. No action needed. Continuing."
}

# -------------------------
# Step 3: Remove Windows Temp Files
# -------------------------
Write-Log "Step 3. Removing Windows Temp Files..."

$tempPath = "C:\Windows\Temp"

if (Test-Path "$tempPath\*") {
    Write-Log "Step 3. Temp files found in '$tempPath'. Attempting removal..."
    try {
        Remove-Item -Path "$tempPath\*" -Recurse -Force -ErrorAction Stop
        Write-Log "Step 3. Temp file removal command executed."
    }
    catch {
        Write-Log "Step 3. ERROR removing temp files: $_"
    }

    # Post-action verification
    if (-not (Test-Path "$tempPath\*")) {
        Write-Log "Step 3. VERIFIED - All temp files successfully removed."
    }
    else {
        $remainingCount = @(Get-ChildItem -Path $tempPath -Recurse -ErrorAction SilentlyContinue).Count
        Write-Log "Step 3. WARNING - $remainingCount item(s) still present in temp folder (may be locked). Continuing with script."
    }
}
else {
    Write-Log "Step 3. No temp files found in '$tempPath'. No removal needed. Continuing."
}

# Save current directory
$originalLocation = Get-Location
Set-Location $env:SystemRoot\System32

# -------------------------
# Step 4: Register DLLs
# -------------------------
# Curated DLL list: Only DLLs confirmed present on modern Windows (10/11) that support
# self-registration (regsvr32 exit code 0). Removed legacy/missing DLLs and those that
# do not export DllRegisterServer (exit code 4):
#   Removed (missing):   msxml.dll, gpkcsp.dll, sccbase.dll, slbcsp.dll, initpki.dll,
#                        wuaueng1.dll, wucltui.dll, wuweb.dll, qmgrprxy.dll, wucltux.dll,
#                        muweb.dll, wuwebv.dll
#   Removed (exit 4):    mshtml.dll, shdocvw.dll, browseui.dll, wuaueng.dll, qmgr.dll
Write-Log "Step 4. Registering DLLs..."

$dlls = @(
    'atl.dll', 'urlmon.dll', 'jscript.dll', 'vbscript.dll', 'scrrun.dll',
    'msxml3.dll', 'msxml6.dll', 'actxprxy.dll', 'softpub.dll', 'wintrust.dll',
    'dssenh.dll', 'rsaenh.dll', 'cryptdlg.dll', 'oleaut32.dll', 'ole32.dll',
    'shell32.dll', 'wuapi.dll', 'wups.dll', 'wups2.dll'
)

$successCount = 0
$failCount = 0

foreach ($dll in $dlls) {
    Write-Log "Step 4. Registering $dll..."
    $process = Start-Process -FilePath "$env:SystemRoot\System32\regsvr32.exe" `
        -ArgumentList "/s $dll" -Wait -PassThru -ErrorAction SilentlyContinue
    if ($process -and $process.ExitCode -eq 0) {
        Write-Log "Step 4. VERIFIED - $dll registered successfully."
        $successCount++
    }
    else {
        $exitCode = if ($process) { $process.ExitCode } else { 'N/A' }
        Write-Log "Step 4. WARNING - $dll registration returned exit code $exitCode. Continuing with script."
        $failCount++
    }
}

Write-Log "Step 4. DLL registration complete. Success: $successCount, Failed: $failCount out of $($dlls.Count) total."

# -------------------------
# Step 5: Delete BITS jobs
# -------------------------
# Note: BITS service must be running to query/remove jobs.
# We temporarily enable and start it here, then stop it again before restore.
Write-Log "Step 5. Deleting BITS jobs..."

Set-Service -Name BITS -StartupType Manual -ErrorAction SilentlyContinue
Start-Service -Name BITS -ErrorAction SilentlyContinue

$bitsJobs = Get-BitsTransfer -AllUsers -ErrorAction SilentlyContinue

if ($bitsJobs) {
    $jobCount = @($bitsJobs).Count
    Write-Log "Step 5. Found $jobCount BITS job(s). Attempting removal..."
    foreach ($job in $bitsJobs) {
        try {
            Remove-BitsTransfer -BitsJob $job -ErrorAction Stop
            Write-Log "Step 5. Removed BITS job: $($job.JobId)"
        }
        catch {
            Write-Log "Step 5. ERROR removing BITS job $($job.JobId): $_"
        }
    }

    # Post-action verification
    $remainingJobs = Get-BitsTransfer -AllUsers -ErrorAction SilentlyContinue
    if (-not $remainingJobs) {
        Write-Log "Step 5. VERIFIED - All BITS job(s) successfully removed."
    }
    else {
        $remainingCount = @($remainingJobs).Count
        Write-Log "Step 5. WARNING - $remainingCount BITS job(s) still present. Continuing with script."
    }
}
else {
    Write-Log "Step 5. No BITS jobs found. No removal needed. Continuing."
}

# Stop BITS again before the restore step to maintain consistent state
Stop-Service -Name BITS -Force -ErrorAction SilentlyContinue
Set-Service -Name BITS -StartupType Disabled -ErrorAction SilentlyContinue

# -------------------------
# Step 6: Restore service startup types + start services
# -------------------------
Write-Log "Step 6. Restoring service startup types..."

if ($original) {
    $restoreSuccess = 0
    $restoreFail = 0

    foreach ($item in $original) {
        # Map CIM StartMode (e.g. "Auto") to Set-Service StartupType (e.g. "Automatic")
        $mappedStartType = $startTypeMap[$item.StartType]
        if (-not $mappedStartType) { $mappedStartType = $item.StartType }

        Write-Log "Step 6. Restoring '$($item.Name)' to startup type '$mappedStartType' (CIM: '$($item.StartType)')..."
        try {
            Set-Service -Name $item.Name -StartupType $mappedStartType -ErrorAction Stop
            Start-Service -Name $item.Name -ErrorAction SilentlyContinue
            Write-Log "Step 6. Restore command executed for '$($item.Name)'."
        }
        catch {
            Write-Log "Step 6. ERROR restoring '$($item.Name)': $_"
            $restoreFail++
            continue
        }

        # Post-action verification (CIM returns "Auto", so compare against original CIM value)
        $currentStartType = (Get-CimInstance -ClassName Win32_Service -Filter "Name='$($item.Name)'" -ErrorAction SilentlyContinue).StartMode
        if ($currentStartType -eq $item.StartType) {
            Write-Log "Step 6. VERIFIED - '$($item.Name)' startup type restored to '$currentStartType'."
            $restoreSuccess++
        }
        else {
            Write-Log "Step 6. WARNING - '$($item.Name)' startup type is '$currentStartType', expected '$($item.StartType)'. Continuing with script."
            $restoreFail++
        }
    }

    Write-Log "Step 6. Service restore complete. Success: $restoreSuccess, Failed: $restoreFail out of $(@($original).Count) total."
}
else {
    Write-Log "Step 6. No original service states were captured. Nothing to restore."
}

# -------------------------
# Step 7: Reset WinSock & WinHTTP
# -------------------------
# Note: Runs after service restore to ensure dependent services are available.
Write-Log "Step 7. Resetting WinSock and WinHTTP..."

# WinSock reset
Write-Log "Step 7. Attempting WinSock reset..."
netsh winsock reset 2>&1 | Out-Null
if ($LASTEXITCODE -eq 0) {
    Write-Log "Step 7. VERIFIED - WinSock reset completed successfully."
}
else {
    Write-Log "Step 7. WARNING - WinSock reset returned exit code $LASTEXITCODE. Continuing with script."
}

# WinHTTP proxy reset
Write-Log "Step 7. Attempting WinHTTP proxy reset..."
netsh winhttp reset proxy 2>&1 | Out-Null
if ($LASTEXITCODE -eq 0) {
    Write-Log "Step 7. VERIFIED - WinHTTP proxy reset completed successfully."
}
else {
    Write-Log "Step 7. WARNING - WinHTTP proxy reset returned exit code $LASTEXITCODE. Continuing with script."
}

# -------------------------
# Step 8: SCCM Cache Clear
# -------------------------
Write-Log "Step 8. Clearing SCCM Cache..."

try {
    $CCMComObject = New-Object -ComObject 'UIResource.UIResourceMgr' -ErrorAction Stop
    $cacheElements = $CCMComObject.GetCacheInfo().GetCacheElements()
    $cacheCount = @($cacheElements).Count

    if ($cacheCount -gt 0) {
        Write-Log "Step 8. Found $cacheCount cache element(s). Attempting removal..."

        $clearSuccess = 0
        $clearFail = 0

        foreach ($CacheItem in $cacheElements) {
            try {
                # Use DeleteCacheElement (compatible with all SCCM client versions)
                $null = $CCMComObject.GetCacheInfo().DeleteCacheElement([string]$($CacheItem.CacheElementID))
                $clearSuccess++
            }
            catch {
                Write-Log "Step 8. ERROR - Failed to remove cache element '$($CacheItem.CacheElementID)': $_"
                $clearFail++
            }
        }

        # Post-action verification
        $remainingCount = @($CCMComObject.GetCacheInfo().GetCacheElements()).Count
        if ($remainingCount -eq 0) {
            Write-Log "Step 8. VERIFIED - All $cacheCount cache element(s) successfully removed."
        }
        else {
            Write-Log "Step 8. WARNING - $remainingCount cache element(s) still present after removal. Continuing with script."
        }

        Write-Log "Step 8. SCCM Cache clear complete. Success: $clearSuccess, Failed: $clearFail out of $cacheCount total."
    }
    else {
        Write-Log "Step 8. No SCCM cache elements found. No removal needed. Continuing."
    }
}
catch {
    Write-Log "Step 8. ERROR - Unable to access SCCM Cache (SCCM client may not be installed): $_. Continuing with script."
}

# -------------------------
# Step 9: Remove pending.xml
# -------------------------
# Note: pending.xml in WinSxS is owned by TrustedInstaller. Even as Administrator,
# removal may fail due to ACL restrictions. This is a known Windows limitation.
Write-Log "Step 9. Checking pending.xml..."

$pendingXml = "$env:SystemRoot\WinSxS\pending.xml"

if (Test-Path $pendingXml) {
    Write-Log "Step 9. pending.xml found at '$pendingXml'. Attempting removal..."
    try {
        Remove-Item $pendingXml -Force -ErrorAction Stop
        Write-Log "Step 9. Removal command executed."
    }
    catch {
        Write-Log "Step 9. ERROR - Failed to remove pending.xml (may require TrustedInstaller ownership): $_"
    }

    # Post-action verification
    if (-not (Test-Path $pendingXml)) {
        Write-Log "Step 9. VERIFIED - pending.xml successfully removed."
    }
    else {
        Write-Log "Step 9. WARNING - pending.xml still present after removal attempt. Continuing with script."
    }
}
else {
    Write-Log "Step 9. pending.xml not found at '$pendingXml'. No removal needed. Continuing."
}

# Restore original working directory
Set-Location $originalLocation

Write-Log "Script completed. All 9 steps have been executed."
