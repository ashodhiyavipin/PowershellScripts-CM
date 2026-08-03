param(
    [switch]$Verbose
)

function Write-Log {
    param([string]$Message)
    if ($Verbose) {
        Write-Host "[INFO] $Message"
    }
}

# Path to ScanResult.xml
$ScanPath = 'C:\$WINDOWS.~BT\Sources\Panther\ScanResult.xml'
$result = 0

Write-Log "Starting driver block detection..."
Write-Log "Looking for ScanResult.xml at: $ScanPath"

if (-not (Test-Path $ScanPath)) {
    Write-Log "ScanResult.xml not found. Returning 0."
    Write-Output $result
    exit
}

Write-Log "ScanResult.xml found. Loading XML..."

[xml]$xml = Get-Content $ScanPath

# Register namespace
$ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$ns.AddNamespace("c", "http://www.microsoft.com/ApplicationExperience/UpgradeAdvisor/01012009")

Write-Log "Selecting DriverPackage nodes using namespace..."

$drivers = $xml.SelectNodes("//c:DriverPackage", $ns)

if (-not $drivers) {
    Write-Log "No DriverPackage nodes found. Namespace may be wrong. Returning 0."
    Write-Output $result
    exit
}

Write-Log "Found $($drivers.Count) driver entries in XML."

# Identify blocked drivers
$blocked = $drivers | Where-Object {
    $_.BlockMigration -eq "True" -and $_.HasSignedBinaries -eq "False"
}

Write-Log "Blocked drivers detected: $($blocked.Count)"

foreach ($drv in $blocked) {

    $inf = $drv.Inf
    Write-Log "Processing blocked INF: $inf"

    # Match INF reliably
    $driverInfo = Get-WindowsDriver -Online |
        Where-Object { $_.Driver -like "*$inf" }

    if (-not $driverInfo) {
        Write-Log "No matching installed driver found for $inf"
        continue
    }

    foreach ($d in $driverInfo) {
        Write-Log "Matched driver: $($d.Driver)"
        Write-Log "Provider: $($d.ProviderName), Class: $($d.ClassName), Date: $($d.Date), OriginalFile: $($d.OriginalFileName)"

        $isLegacyPrinter =
            ($d.ProviderName -eq "Microsoft") -and
            ($d.ClassName -eq "Printer") -and
            ($d.Date -like "*2006*") -and
            ($d.OriginalFileName -match "prnms001|prnms009")

        if ($isLegacyPrinter) {
            Write-Log "Legacy 2006 Microsoft printer driver detected. Setting result = 1"
            $result = 1
        }
        else {
            Write-Log "Driver does NOT match legacy printer criteria."
        }
    }
}

Write-Log "Final result: $result"
Write-Output $result
