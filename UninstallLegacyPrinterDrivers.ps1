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
# STEP 5 — ATTEMPT DRIVER REMOVAL FOR EACH BLOCKED INF
# ============================================================

foreach ($inf in $BlockedInfs) {

    Write-Log "======================================================" 'INFO'
    Write-Log " ATTEMPTING REMOVAL FOR: $inf" 'INFO'
    Write-Log "======================================================" 'INFO'

    $cmd = "pnputil /delete-driver $inf /uninstall /force"
    Write-Log "Executing: $cmd" 'INFO'

    try {
        $output = pnputil /delete-driver $inf /uninstall /force 2>&1
        foreach ($line in $output) {
            Write-Log "Removal output: $line" 'INFO'
        }
    }
    catch {
        Write-Log "Error executing pnputil for $inf : $_" 'ERROR'
    }
}

# Output ONLY the number
Write-Output $result