<#
.SYNOPSIS
    Entry point script for the modular application uninstaller system.

.DESCRIPTION
    This is the main entry point for uninstalling applications using the modular uninstaller system.
    It loads the core module and all app-specific uninstaller modules, then dispatches to the 
    appropriate uninstaller based on pattern matching.
    
    The script supports:
    - Loose, case-insensitive matching of application names
    - Centralized pattern-to-function mapping
    - App-specific uninstall functions with fallback to registry-based uninstall
    - CMTrace-compatible logging for easy troubleshooting
    - Detailed structured summary output

.PARAMETER AppName
    The name (or partial name) of the application to uninstall.
    Case-insensitive. Examples: "KeePass", "Adobe Reader", "Chrome"

.PARAMETER LogPath
    Optional. The directory path where log files should be written.
    Default: C:\Windows\fndr\logs
    
    Log files are named: ModularUninstall_<AppName>_<timestamp>.log

.EXAMPLE
    .\MainUninstallScript.ps1 -AppName "KeePass"
    
    Uninstalls all versions of KeePass using the app-specific KeePass uninstaller.

.EXAMPLE
    .\MainUninstallScript.ps1 -AppName "Adobe Reader"
    
    Uninstalls all versions of Adobe Reader using the app-specific Adobe uninstaller.

.EXAMPLE
    .\MainUninstallScript.ps1 -AppName "SomeApp" -LogPath "C:\Logs"
    
    Uninstalls SomeApp and writes logs to C:\Logs directory.

.NOTES
    File Name:      MainUninstallScript.ps1
    Author:         V.Ashodhiya
    Version:        1.0
    Date:           2026-01-27
    Log Path:       C:\Windows\fndr\logs\ModularUninstall_<AppName>_<timestamp>.log
    Log Format:     CMTrace-compatible (can be viewed with CMTrace.exe)
    
    Version History:
    1.0 - Initial release with modular architecture
        - Core module with dispatcher and CMTrace logging
        - App-specific modules for KeePass and Adobe Reader
        - Registry-based fallback for unknown applications
        - Structured summary output
    
    Directory Structure:
    - Application-Uninstall\
      - Core\Uninstall.Core.psm1             (Core module)
      - Apps\Uninstall.KeePass.ps1           (KeePass uninstaller)
      - Apps\Uninstall.AdobeReader.ps1       (Adobe Reader uninstaller)
      - Apps\Uninstall.Template.ps1          (Template for new apps)
      - MainUninstallScript.ps1              (This file)
    
    Adding New Applications:
    1. Copy Apps\Uninstall.Template.ps1 to Apps\Uninstall.<AppName>.ps1
    2. Rename function to Uninstall-<AppName>
    3. Implement detection and uninstall logic
    4. Add pattern to $AppMapping in Core\Uninstall.Core.psm1
    5. Test with: .\MainUninstallScript.ps1 -AppName "YourApp"

.LINK
    CMTrace.exe - https://learn.microsoft.com/en-us/mem/configmgr/core/support/cmtrace
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, Position = 0, HelpMessage = "Name of the application to uninstall")]
    [ValidateNotNullOrEmpty()]
    [string]$AppName,
    
    [Parameter(Mandatory = $false, Position = 1, HelpMessage = "Directory path for log files")]
    [ValidateNotNullOrEmpty()]
    [string]$LogPath = "C:\Windows\fndr\logs"
)

#region Script Initialization

# Set error action preference
$ErrorActionPreference = "Stop"

# Get script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$coreModulePath = Join-Path $scriptDir "Core\Uninstall.Core.psm1"
$appsDir = Join-Path $scriptDir "Apps"

# Display banner
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Modular Application Uninstaller v1.0" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Application: $AppName" -ForegroundColor Yellow
Write-Host "Log Path:    $LogPath" -ForegroundColor Yellow
Write-Host ""

#endregion

#region Module Loading

try {
    # Validate core module exists
    if (-not (Test-Path $coreModulePath)) {
        Write-Error "Core module not found at: $coreModulePath"
        Write-Error "Please ensure the directory structure is intact."
        exit 1
    }
    
    # Import core module
    Write-Host "[1/2] Loading core module..." -ForegroundColor Cyan
    Import-Module $coreModulePath -Force -ErrorAction Stop
    Write-Host "      Core module loaded successfully" -ForegroundColor Green
    
    # Load all app-specific modules
    Write-Host "[2/2] Loading app-specific modules..." -ForegroundColor Cyan
    
    if (Test-Path $appsDir) {
        $appScripts = Get-ChildItem -Path $appsDir -Filter "Uninstall.*.ps1" -ErrorAction SilentlyContinue
        
        if ($appScripts) {
            $loadedCount = 0
            foreach ($script in $appScripts) {
                # Skip the template
                if ($script.Name -eq "Uninstall.Template.ps1") {
                    continue
                }
                
                try {
                    # Dot-source the script to load the function
                    . $script.FullName
                    Write-Host "      Loaded: $($script.Name)" -ForegroundColor Gray
                    $loadedCount++
                }
                catch {
                    Write-Warning "Failed to load $($script.Name): $_"
                }
            }
            Write-Host "      $loadedCount app module(s) loaded successfully" -ForegroundColor Green
        }
        else {
            Write-Host "      No app modules found (will use registry fallback)" -ForegroundColor Yellow
        }
    }
    else {
        Write-Warning "Apps directory not found: $appsDir"
        Write-Host "      Will use registry fallback for all applications" -ForegroundColor Yellow
    }
    
    Write-Host ""
    
}
catch {
    Write-Error "Failed to load modules: $_"
    Write-Error "Stack Trace: $($_.ScriptStackTrace)"
    exit 1
}

#endregion

#region Execute Uninstall

try {
    Write-Host "Starting uninstall operation..." -ForegroundColor Cyan
    Write-Host ""
    
    # Call the central dispatcher
    $result = Uninstall-Application -AppName $AppName -LogPath $LogPath
    
    # Display summary
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  OPERATION SUMMARY" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Application:        $($result.AppNameInput)" -ForegroundColor White
    Write-Host "Matched Pattern:    $($result.MatchedPattern)" -ForegroundColor White
    Write-Host "Selected Function:  $($result.SelectedFunction)" -ForegroundColor White
    Write-Host "Primary Method:     $($result.PrimaryMethod)" -ForegroundColor White
    Write-Host "Fallback Used:      $($result.FallbackUsed)" -ForegroundColor White
    Write-Host "Versions Handled:   $($result.VersionsHandled.Count)" -ForegroundColor White
    
    if ($result.VersionsHandled.Count -gt 0) {
        foreach ($version in $result.VersionsHandled) {
            Write-Host "  - $version" -ForegroundColor Gray
        }
    }
    
    Write-Host "Exit Code:          $($result.FinalExitCode)" -ForegroundColor White
    
    if ($result.Status -eq "Success") {
        Write-Host "Status:             $($result.Status)" -ForegroundColor Green
    }
    else {
        Write-Host "Status:             $($result.Status)" -ForegroundColor Red
    }
    
    Write-Host ""
    Write-Host "Log file: $Global:CurrentLogFile" -ForegroundColor Gray
    Write-Host "Tip: View log with CMTrace.exe for best experience" -ForegroundColor Gray
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    
    # Exit with appropriate code
    exit $result.FinalExitCode
    
}
catch {
    Write-Error "Fatal error during uninstall operation: $_"
    Write-Error "Stack Trace: $($_.ScriptStackTrace)"
    exit 1
}

#endregion
