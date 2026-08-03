<#
.SYNOPSIS
    Template for creating new app-specific uninstallers.

.DESCRIPTION
    This is a reusable template for creating new application-specific uninstaller functions.
    Follow the instructions below to create a new uninstaller:
    
    INSTRUCTIONS:
    1. Copy this file to: Uninstall.<AppName>.ps1
       Example: Uninstall.7Zip.ps1, Uninstall.Chrome.ps1
    
    2. Rename the function below from "Uninstall-Template" to "Uninstall-<AppName>"
       Example: Uninstall-7Zip, Uninstall-Chrome
    
    3. Update the .SYNOPSIS and .DESCRIPTION sections above
    
    4. Implement the detection logic:
       - Modify the Where-Object filter to match your specific application
       - Add any exclusions for specific versions or editions if needed
    
    5. Implement the uninstall logic:
       - Most Windows apps use MSI (msiexec.exe)
       - Some use custom uninstallers (check UninstallString)
       - Add any app-specific switches or parameters
    
    6. Update the central mapping table in Uninstall.Core.psm1:
       - Add a new entry: 'pattern' = 'Uninstall-<AppName>'
       - Use a regex pattern that matches your app name (case-insensitive)
       Example: '7-?zip' = 'Uninstall-7Zip'
    
    7. Test your new uninstaller:
       - Run: .\MainUninstallScript.ps1 -AppName "YourApp"
       - Check the log file in C:\Windows\fndr\logs\
       - Verify with CMTrace.exe

.NOTES
    File Name:    Uninstall.Template.ps1
    Author:       Your Name
    Version:      1.0
    Date:         YYYY-MM-DD
    
.EXAMPLE
    Uninstall-Template
#>

function Uninstall-Template {
    [CmdletBinding()]
    param ()
    
    # TODO: Update component name to match your app
    Write-CMLog -Message "=== Starting [AppName]-specific uninstall ===" -Component "Uninstall-Template" -Severity 1
    
    # Standard registry paths for installed applications
    $registryPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    # Initialize result structure (DO NOT MODIFY)
    $result = @{
        Success     = $false
        ExitCode    = 1
        Found       = 0
        Uninstalled = 0
        Failed      = 0
        Versions    = @()
    }
    
    # Search for all installations
    foreach ($regPath in $registryPaths) {
        try {
            Write-CMLog -Message "Searching: $regPath" -Component "Uninstall-Template" -Severity 1
            
            # TODO: Modify the filter to match your application
            # Examples:
            # - $_.DisplayName -like '*7-Zip*'
            # - $_.DisplayName -like '*Google Chrome*'
            # - $_.DisplayName -like '*VLC*'
            $apps = Get-ItemProperty $regPath -ErrorAction SilentlyContinue | 
            Where-Object { 
                $_.DisplayName -like '*YourAppName*'
                # Optional: Add exclusions
                # -and $_.DisplayName -notlike '*Exclude*'
                # -and $_.DisplayVersion -notlike '*1.0.0*'
            }
            
            foreach ($app in $apps) {
                $result.Found++
                
                # Gather information
                $displayName = $app.DisplayName
                $displayVersion = $app.DisplayVersion
                $productCode = $app.PSChildName
                $versionString = "$displayName ($displayVersion)"
                $result.Versions += $versionString
                
                Write-CMLog -Message "Found: $versionString" -Component "Uninstall-Template" -Severity 1
                Write-CMLog -Message "Product Code: $productCode" -Component "Uninstall-Template" -Severity 1
                
                # Determine uninstall method
                $uninstallCmd = $null
                
                # Method 1: QuietUninstallString (preferred - already silent)
                if ($app.QuietUninstallString) {
                    $uninstallCmd = $app.QuietUninstallString
                    Write-CMLog -Message "Using QuietUninstallString: $uninstallCmd" -Component "Uninstall-Template" -Severity 1
                }
                # Method 2: UninstallString (may need silent switches)
                elseif ($app.UninstallString) {
                    $uninstallCmd = $app.UninstallString
                    Write-CMLog -Message "Using UninstallString: $uninstallCmd" -Component "Uninstall-Template" -Severity 1
                    
                    # If it's MSI, make it silent
                    if ($uninstallCmd -imatch "msiexec") {
                        # Convert /I to /X if present
                        if ($uninstallCmd -imatch "/I") {
                            $uninstallCmd = $uninstallCmd -ireplace "/I", "/X"
                            Write-CMLog -Message "Converted /I to /X" -Component "Uninstall-Template" -Severity 1
                        }
                        
                        # Add silent switches if not present
                        if ($uninstallCmd -notmatch "/qn") {
                            $uninstallCmd += " /qn /norestart"
                            Write-CMLog -Message "Added silent switches" -Component "Uninstall-Template" -Severity 1
                        }
                    }
                    
                    # TODO: Add app-specific silent switches here
                    # Examples:
                    # - NSIS installers: add "/S" 
                    # - InnoSetup: add "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART"
                    # - Custom installers: check documentation
                }
                # Method 3: Use product code directly (MSI only)
                elseif ($productCode -match "^\{[A-F0-9-]+\}$") {
                    $uninstallCmd = "msiexec.exe /x $productCode /qn /norestart"
                    Write-CMLog -Message "Using product code for MSI uninstall: $uninstallCmd" -Component "Uninstall-Template" -Severity 1
                }
                
                # Execute uninstall
                if ($uninstallCmd) {
                    try {
                        Write-CMLog -Message "Executing uninstall: $uninstallCmd" -Component "Uninstall-Template" -Severity 1
                        
                        # TODO: Choose execution method based on your app
                        
                        # Option A: Run through cmd.exe (most compatible)
                        $process = Start-Process -FilePath "cmd.exe" `
                            -ArgumentList "/c", $uninstallCmd `
                            -Wait `
                            -WindowStyle Hidden `
                            -PassThru `
                            -ErrorAction Stop
                        
                        # Option B: Run directly (for msiexec, setup.exe, etc.)
                        # Extract executable and arguments first, then use Start-Process
                        
                        # Check exit code
                        # Note: Some installers use different success codes
                        # Common codes: 0 = success, 3010 = success with reboot required
                        if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
                            Write-CMLog -Message "Successfully uninstalled: $versionString (Exit Code: $($process.ExitCode))" -Component "Uninstall-Template" -Severity 1
                            $result.Uninstalled++
                        }
                        else {
                            Write-CMLog -Message "Uninstall failed with exit code $($process.ExitCode): $versionString" -Component "Uninstall-Template" -Severity 3
                            $result.Failed++
                        }
                    }
                    catch {
                        Write-CMLog -Message "Exception during uninstall: $_" -Component "Uninstall-Template" -Severity 3
                        Write-CMLog -Message "Command: $uninstallCmd" -Component "Uninstall-Template" -Severity 3
                        $result.Failed++
                    }
                }
                else {
                    Write-CMLog -Message "No uninstall method found for: $versionString" -Component "Uninstall-Template" -Severity 2
                    $result.Failed++
                }
            }
        }
        catch {
            Write-CMLog -Message "Error searching registry path $regPath : $_" -Component "Uninstall-Template" -Severity 3
        }
    }
    
    # Determine overall result (DO NOT MODIFY THIS SECTION)
    if ($result.Found -eq 0) {
        Write-CMLog -Message "No installations found" -Component "Uninstall-Template" -Severity 2
        $result.Success = $false
        $result.ExitCode = 1
    }
    elseif ($result.Uninstalled -gt 0 -and $result.Failed -eq 0) {
        Write-CMLog -Message "All versions uninstalled successfully ($($result.Uninstalled) version(s))" -Component "Uninstall-Template" -Severity 1
        $result.Success = $true
        $result.ExitCode = 0
    }
    elseif ($result.Uninstalled -gt 0) {
        Write-CMLog -Message "Partial success: $($result.Uninstalled) succeeded, $($result.Failed) failed" -Component "Uninstall-Template" -Severity 2
        $result.Success = $true  # Partial success
        $result.ExitCode = 0
    }
    else {
        Write-CMLog -Message "All uninstall attempts failed" -Component "Uninstall-Template" -Severity 3
        $result.Success = $false
        $result.ExitCode = 1
    }
    
    Write-CMLog -Message "=== Uninstall complete ===" -Component "Uninstall-Template" -Severity 1
    
    return $result
}

# QUICK REFERENCE: Common Silent Switches
# ----------------------------------------
# MSI (Windows Installer):
#   msiexec.exe /x {ProductCode} /qn /norestart
#
# NSIS Installer:
#   uninstall.exe /S
#
# InnoSetup:
#   unins000.exe /VERYSILENT /SUPPRESSMSGBOXES /NORESTART
#
# InstallShield:
#   setup.exe /s /v"/qn"
#
# Wise Installer:
#   uninstall.exe /s
#
# Custom Executables:
#   Check vendor documentation for silent switches
#   Common: /silent, /quiet, /q, -silent, --silent
