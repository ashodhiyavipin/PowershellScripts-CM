<#
.SYNOPSIS
    Detects legacy Microsoft printer drivers that block Windows 11 upgrades,
    logs full driver metadata, and attempts automated removal of the affected drivers.

.DESCRIPTION
    This script analyzes the Windows Setup ScanResult.xml file to identify printer
    drivers that are flagged as blocking OS migration. It specifically targets the
    legacy Microsoft printer drivers from 2006 (prnms001.inf and prnms009.inf),
    which are known to block Windows 11 24H2 upgrades.

    Once detected, the script:
      • Logs detailed metadata for each blocked driver
      • Enumerates DriverStore contents
      • Captures pnputil driver information
      • Checks for Version-3 registry entries
      • Attempts removal using pnputil /delete-driver /uninstall /force

    All actions are logged to both console (when -Verbose is used) and a persistent
    log file at C:\Temp\DriverBlock.log.

.NOTES
    Author:        Vipin  
    Purpose:       Detection and remediation of legacy printer drivers  
    Logging:       Full verbose commentary written to C:\Temp\DriverBlock.log  
    Requirements:  Must be run with administrative privileges  
    Tested On:     Windows 10, Windows 11, Windows 11 24H2 upgrade scenarios  

.VERSION
    1.0.0   Initial combined script with detection + metadata reporting
    1.1.0   Added full DriverStore enumeration and pnputil metadata capture
    1.2.0   Added automated removal phase for blocked drivers
    1.3.0   Finalized verbose logging and structured reporting format
    1.4.0   Added automated removal phase for blocked drivers using removal of optional windows features.
#>

param(
    [switch]$Verbose
)

# Make verbose the default when the script is run without -Verbose
if (-not $PSBoundParameters.ContainsKey('Verbose')) {
    $Verbose = $true
}

# Log file setup
$LogFolder = 'C:\Temp'
$LogFile = Join-Path $LogFolder 'DriverBlock.log'

if (-not (Test-Path $LogFolder)) {
    New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )

    $time = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
    $line = "$time [$Level] $Message"

    # Write to file
    try {
        Add-Content -Path $LogFile -Value $line -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to write to log file $LogFile : $_"
    }

    # Write to console when verbose
    if ($Verbose) {
        switch ($Level) {
            'INFO' { Write-Host "[INFO]  $Message" -ForegroundColor Green }
            'WARN' { Write-Warning $Message }
            'ERROR' { Write-Error $Message }
        }
    }
}

# Path to ScanResult.xml
$ScanPath = 'C:\$WINDOWS.~BT\Sources\Panther\ScanResult.xml'
$result = 0
$BlockedInfs = @()

Write-Log "Starting driver block detection..." 'INFO'
Write-Log "Looking for ScanResult.xml at: $ScanPath" 'INFO'

if (-not (Test-Path $ScanPath)) {
    Write-Log "ScanResult.xml not found. Returning 0." 'WARN'
    Write-Output $result
    exit
}

Write-Log "ScanResult.xml found. Loading XML..." 'INFO'

[xml]$xml = Get-Content $ScanPath

# Register namespace
$ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$ns.AddNamespace("c", "http://www.microsoft.com/ApplicationExperience/UpgradeAdvisor/01012009")

Write-Log "Selecting DriverPackage nodes using namespace..." 'INFO'

$drivers = $xml.SelectNodes("//c:DriverPackage", $ns)

if (-not $drivers) {
    Write-Log "No DriverPackage nodes found. Namespace may be wrong. Returning 0." 'WARN'
    Write-Output $result
    exit
}

Write-Log "Found $($drivers.Count) driver entries in XML." 'INFO'

# Identify blocked drivers
$blocked = $drivers | Where-Object {
    $_.BlockMigration -eq "True" -and $_.HasSignedBinaries -eq "False"
}

Write-Log "Blocked drivers detected: $($blocked.Count)" 'INFO'

foreach ($drv in $blocked) {

    $inf = $drv.Inf
    Write-Log "Processing blocked INF: $inf" 'INFO'

    # Match INF reliably
    $driverInfo = Get-WindowsDriver -Online |
    Where-Object { $_.Driver -like "*$inf" }

    if (-not $driverInfo) {
        Write-Log "No matching installed driver found for $inf" 'WARN'
        continue
    }

    foreach ($d in $driverInfo) {
        Write-Log "Matched driver: $($d.Driver)" 'INFO'
        Write-Log "Provider: $($d.ProviderName); Class: $($d.ClassName); Date: $($d.Date); OriginalFile: $($d.OriginalFileName)" 'INFO'

        $isLegacyPrinter =
        ($d.ProviderName -eq "Microsoft") -and
        ($d.ClassName -eq "Printer") -and
        ($d.Date -like "*2006*") -and
        ($d.OriginalFileName -match "prnms001|prnms009")

        if ($isLegacyPrinter) {
            Write-Log "Legacy 2006 Microsoft printer driver detected for INF $inf. Adding to blocked list." 'WARN'
            $BlockedInfs += $inf
            $result = 1
        }
        else {
            Write-Log "Driver for INF $inf does NOT match legacy printer criteria." 'INFO'
        }
    }
}

Write-Log "Final detection result: $result" 'INFO'
Write-Log "Blocked INF list: $($BlockedInfs -join ', ')" 'INFO'

# ============================================================
# STEP 4 — FULL DRIVER METADATA REPORT FOR EACH BLOCKED INF
# ============================================================

foreach ($inf in $BlockedInfs) {

    Write-Log "======================================================" 'INFO'
    Write-Log " FULL DRIVER REPORT FOR: $inf" 'INFO'
    Write-Log "======================================================" 'INFO'

    # 1. Get-WindowsDriver metadata
    $driver = Get-WindowsDriver -Online | Where-Object { $_.Driver -like "*$inf" }

    if (-not $driver) {
        Write-Log "No driver metadata found for $inf" 'WARN'
        continue
    }

    Write-Log "Provider: $($driver.ProviderName)" 'INFO'
    Write-Log "Class: $($driver.ClassName)" 'INFO'
    Write-Log "Date: $($driver.Date)" 'INFO'
    Write-Log "Original INF Path: $($driver.OriginalFileName)" 'INFO'
    Write-Log "Catalog File: $($driver.CatalogFile)" 'INFO'
    Write-Log "Signature: $($driver.DriverSignature)" 'INFO'
    Write-Log "Version: $($driver.Version)" 'INFO'
    Write-Log "Class GUID: $($driver.ClassGuid)" 'INFO'
    Write-Log "Inbox: $($driver.Inbox)" 'INFO'

    # 2. Extract DriverStore folder
    $storeFolder = Split-Path $driver.OriginalFileName -Parent
    Write-Log "DriverStore Path: $storeFolder" 'INFO'

    # 3. List all files in DriverStore
    if (Test-Path $storeFolder) {
        Write-Log "Listing DriverStore files..." 'INFO'
        Get-ChildItem $storeFolder -Recurse | ForEach-Object {
            Write-Log "File: $($_.FullName)" 'INFO'
        }
    }
    else {
        Write-Log "DriverStore folder not found." 'WARN'
    }

    # 4. pnputil metadata
    Write-Log "pnputil metadata:" 'INFO'
    pnputil /enum-drivers | Select-String -Pattern $inf -Context 0, 8 | ForEach-Object {
        Write-Log $_.Line 'INFO'
    }

    # 5. Registry Version-3 entries
    Write-Log "Registry Version-3 entries:" 'INFO'
    $reg = Get-ChildItem "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments\Windows x64\Drivers\Version-3" `
        -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -match ($inf.Replace('.inf', '')) }

    if ($reg) {
        foreach ($key in $reg) {
            Write-Log "Registry Key: $($key.Name)" 'INFO'
            Get-ItemProperty $key.PsPath | ForEach-Object {
                Write-Log "  $_" 'INFO'
            }
        }
    }
    else {
        Write-Log "No Version-3 registry entries found." 'INFO'
    }
}

Write-Log "Driver block detection + full metadata reporting completed." 'INFO'

# ============================================================
# STEP 5 — REMOVE DEVICES, DISABLE FEATURES, AND REMOVE DRIVERS
# ============================================================

Write-Log "Stopping Print Spooler service for safe driver removal..." 'INFO'
try {
    Stop-Service -Name spooler -Force -ErrorAction Stop
    Write-Log "Print Spooler service stopped successfully." 'INFO'
}
catch {
    Write-Log "Failed to stop Print Spooler service: $_" 'ERROR'
}

# Try to import PnPDevice module (for Get-PnpDevice / Remove-PnpDevice)
try {
    Import-Module PnpDevice -ErrorAction Stop
    $pnpAvailable = $true
    Write-Log "PnpDevice module loaded successfully." 'INFO'
}
catch {
    $pnpAvailable = $false
    Write-Log "PnpDevice module not available. PnP device cleanup will be skipped." 'WARN'
}

foreach ($inf in $BlockedInfs) {

    Write-Log "======================================================" 'INFO'
    Write-Log " REMEDIATION FOR BLOCKED DRIVER: ${inf}" 'INFO'
    Write-Log "======================================================" 'INFO'

    # Re-resolve driver metadata to map to printer drivers
    $driver = Get-WindowsDriver -Online | Where-Object { $_.Driver -like "*$inf" }

    if (-not $driver) {
        Write-Log "No driver metadata found for ${inf} during remediation. Skipping." 'WARN'
        continue
    }

    # 1. Remove all printers using printer drivers that reference this INF
    try {
        $printerDrivers = Get-PrinterDriver -ErrorAction SilentlyContinue |
        Where-Object { $_.InfPath -and ($_.InfPath -like "*$inf*") }

        if ($printerDrivers) {
            foreach ($pd in $printerDrivers) {
                Write-Log "Found printer driver using ${inf}: $($pd.Name) (InfPath: $($pd.InfPath))" 'INFO'

                $printers = Get-Printer -ErrorAction SilentlyContinue |
                Where-Object { $_.DriverName -eq $pd.Name }

                if ($printers) {
                    foreach ($prn in $printers) {
                        Write-Log "Removing printer: Name='$($prn.Name)' Driver='$($prn.DriverName)'" 'INFO'
                        try {
                            Remove-Printer -Name $prn.Name -ErrorAction Stop
                            Write-Log "Successfully removed printer '$($prn.Name)'." 'INFO'
                        }
                        catch {
                            Write-Log "Failed to remove printer '$($prn.Name)': $_" 'ERROR'
                        }
                    }
                }
                else {
                    Write-Log "No printers found using driver '$($pd.Name)'." 'INFO'
                }
            }
        }
        else {
            Write-Log "No printer drivers found that reference INF ${inf}." 'INFO'
        }
    }
    catch {
        Write-Log "Error while enumerating/removing printers for INF ${inf}: $_" 'ERROR'
    }

    # 2. Remove PnP devices bound to this INF (if module available)
    if ($pnpAvailable) {
        try {
            $pnpDevices = Get-PnpDevice -ErrorAction SilentlyContinue |
            Where-Object { $_.Driver -and ($_.Driver -like "*$inf*") }

            if ($pnpDevices) {
                foreach ($dev in $pnpDevices) {
                    Write-Log "Found PnP device using ${inf}: InstanceId='$($dev.InstanceId)' Name='$($dev.FriendlyName)'" 'INFO'
                    try {
                        Disable-PnpDevice -InstanceId $dev.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
                        Write-Log "Disabled PnP device '$($dev.InstanceId)'." 'INFO'
                    }
                    catch {
                        Write-Log "Failed to disable PnP device '$($dev.InstanceId)': $_" 'WARN'
                    }

                    try {
                        Remove-PnpDevice -InstanceId $dev.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
                        Write-Log "Removed PnP device '$($dev.InstanceId)'." 'INFO'
                    }
                    catch {
                        Write-Log "Failed to remove PnP device '$($dev.InstanceId)': $_" 'WARN'
                    }
                }
            }
            else {
                Write-Log "No PnP devices found using INF ${inf}." 'INFO'
            }
        }
        catch {
            Write-Log "Error while enumerating/removing PnP devices for INF ${inf}: $_" 'ERROR'
        }
    }

    # 3. Disable Windows Optional Features that hold these drivers open
    Write-Log "Searching for Windows Optional Features related to Print Workflow..." 'INFO'

    try {
        $workflowFeatures = Get-WindowsOptionalFeature -Online |
        Where-Object { $_.FeatureName -match "Print" -and $_.FeatureName -match "Work" }

        if ($workflowFeatures) {
            foreach ($wf in $workflowFeatures) {
                Write-Log "Disabling workflow feature: $($wf.FeatureName)" 'INFO'
                try {
                    Disable-WindowsOptionalFeature -Online -FeatureName $wf.FeatureName -NoRestart -ErrorAction SilentlyContinue | Out-Null
                    Write-Log "Feature disabled (or already disabled): $($wf.FeatureName)" 'INFO'
                }
                catch {
                    Write-Log "Failed to disable workflow feature $($wf.FeatureName): $_" 'WARN'
                }
            }
        }
        else {
            Write-Log "No Print Workflow features found on this system." 'INFO'
        }
    }
    catch {
        Write-Log "Error while searching for Print Workflow features: $_" 'ERROR'
    }

    # Disable XPS and PrintToPDF (these names are consistent across builds)
    $staticFeatures = @(
        "Printing-XPSServices-Features",
        "Printing-PrintToPDFServices-Features"
    )

    foreach ($feature in $staticFeatures) {
        try {
            Write-Log "Disabling feature: ${feature}" 'INFO'
            Disable-WindowsOptionalFeature -Online -FeatureName $feature -NoRestart -ErrorAction SilentlyContinue | Out-Null
            Write-Log "Feature disabled (or already disabled): ${feature}" 'INFO'
        }
        catch {
            Write-Log "Failed to disable feature ${feature}: $_" 'WARN'
        }
    }


    # 4. Attempt driver removal via pnputil
    Write-Log "Attempting final driver removal via pnputil for ${inf}..." 'INFO'
    $cmd = "pnputil /delete-driver ${inf} /uninstall /force"
    Write-Log "Executing: $cmd" 'INFO'

    try {
        $output = pnputil /delete-driver $inf /uninstall /force 2>&1
        foreach ($line in $output) {
            Write-Log "Removal output: $line" 'INFO'
        }
    }
    catch {
        Write-Log "Error executing pnputil for ${inf} : $_" 'ERROR'
    }
}

Write-Log "Starting Print Spooler service..." 'INFO'
try {
    Start-Service -Name spooler -ErrorAction Stop
    Write-Log "Print Spooler service started successfully." 'INFO'
}
catch {
    Write-Log "Failed to start Print Spooler service: $_" 'ERROR'
}

# Output ONLY the number
Write-Output $result