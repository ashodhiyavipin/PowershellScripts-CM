<# 
.SYNOPSIS
    Script to download all UWP apps from a Windows system with logging functionality.

.DESCRIPTION
    This PowerShell script identifies and downloads all UWP apps from the system, 
    ensuring they are downloaded for all users. It includes a logging mechanism that records each action 
    to a log file for auditing and troubleshooting purposes. The log is stored in the specified directory 
    and captures the timestamp and details of each package download attempt.

.PARAMETERS
    None

.OUTPUTS
    A log file located at 'C:\Windows\fndr\logs\DownloadUWPApps.log' that contains a timestamped record of all actions taken.

.NOTES
    DownloadUWPApps.ps1
    Script History:
    Version 1.0 - Script inception
#>
$appsToDownload = @{
    "9WZDNCRFJBMP" = @{ Name = "Microsoft.WindowsStore"; Platform = "Windows.Desktop" }
    "9N4D0MSMP0PT" = @{ Name = "Microsoft.VP9VideoExtensions"; Platform = "Windows.Universal" }
    "9MZ95KL8MR0L" = @{ Name = "Microsoft.ScreenSketch"; Platform = "Windows.Universal" }
    "9WZDNCRFJ3PZ" = @{ Name = "Microsoft.CompanyPortal"; Platform = "Windows.Universal" }
    "9PMMSR1CGPWG" = @{ Name = "Microsoft.HEIFImageExtension"; Platform = "Windows.Universal" }
    "9PCFS5B6T72H" = @{ Name = "Microsoft.Paint"; Platform = "Windows.Desktop" }
    "9NCTDW2W1BH8" = @{ Name = "Microsoft.RawImageExtension"; Platform = "Windows.Universal" }
    "9N5TDP8VCMHS" = @{ Name = "Microsoft.WebMediaExtensions"; Platform = "Windows.Universal" }
    "9PG2DK419DRG" = @{ Name = "Microsoft.WebpImageExtension"; Platform = "Windows.Universal" }
    "9WZDNCRFJBH4" = @{ Name = "Microsoft.Windows.Photos"; Platform = "Windows.Desktop" }
    "9WZDNCRFHVN5" = @{ Name = "Microsoft.WindowsCalculator"; Platform = "Windows.Universal" }
    "9WZDNCRFJBBG" = @{ Name = "Microsoft.WindowsCamera"; Platform = "Windows.Desktop" }
    "9MSMLRH6LZF3" = @{ Name = "Microsoft.WindowsNotepad"; Platform = "Windows.Desktop" }
    "9N0DX20HK701" = @{ Name = "Microsoft.WindowsTerminal"; Platform = "Windows.Desktop" }
    "9N1F85V9T8BN" = @{ Name = "MicrosoftCorporationII.Windows365"; Platform = "Windows.Desktop" }
    "9NBLGGH4QGHW" = @{ Name = "Microsoft.MicrosoftStickyNotes"; Platform = "Windows.Desktop" }
    "9PGJGD53TN86" = @{ Name = "WinDbg"; Platform = "Windows.Desktop" }
}
foreach ($id in $appsToDownload.Keys) {
    $app = $appsToDownload[$id]
    $platform = $app.Platform

    if (-not $platform) { $platform = "Windows.Desktop" } # Default fallback

    Write-Host "Downloading $($app.Name) ($id) for platform: $platform..."

    winget download --id $id --Platform $platform --architecture x64 --accept-source-agreements --accept-package-agreements -s msstore
}