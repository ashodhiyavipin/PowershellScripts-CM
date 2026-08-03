<#
.SYNOPSIS
    Uninstalls Plantronics Hub Software.

.DESCRIPTION
    This script uninstalls all versions of Plantronics Hub Software found on the system.
    Uses the MSI uninstall method with silent, non-interactive switches and prevents reboot.

.NOTES
    File Name:    Uninstall.PlantronicsHub.ps1
    Author:       Auto-generated
    Version:      1.0
    Date:         2026-01-29
    
.EXAMPLE
    Uninstall-PlantronicsHub
#>

function Uninstall-PlantronicsHub {
    [CmdletBinding()]
    param ()
    
    Write-CMLog -Message "=== Starting Plantronics Hub Software uninstall ===" -Component "Uninstall-PlantronicsHub" -Severity 1
    
    # Standard registry paths for installed applications
    $registryPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    # Initialize result structure
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
            Write-CMLog -Message "Searching: $regPath" -Component "Uninstall-PlantronicsHub" -Severity 1
            
            # Filter for Plantronics Hub Software
            $apps = Get-ItemProperty $regPath -ErrorAction SilentlyContinue | 
            Where-Object { 
                $_.DisplayName -like '*Plantronics Hub*'
            }
            
            foreach ($app in $apps) {
                $result.Found++
                
                # Gather information
                $displayName = $app.DisplayName
                $displayVersion = $app.DisplayVersion
                $productCode = $app.PSChildName
                $versionString = "$displayName ($displayVersion)"
                $result.Versions += $versionString
                
                Write-CMLog -Message "Found: $versionString" -Component "Uninstall-PlantronicsHub" -Severity 1
                Write-CMLog -Message "Product Code: $productCode" -Component "Uninstall-PlantronicsHub" -Severity 1
                
                # Determine uninstall method
                $uninstallCmd = $null
                
                # Method 1: QuietUninstallString (preferred - already silent)
                if ($app.QuietUninstallString) {
                    $uninstallCmd = $app.QuietUninstallString
                    Write-CMLog -Message "Using QuietUninstallString: $uninstallCmd" -Component "Uninstall-PlantronicsHub" -Severity 1
                    
                    # Ensure /norestart is added
                    if ($uninstallCmd -notmatch "/norestart") {
                        $uninstallCmd += " /norestart"
                        Write-CMLog -Message "Added /norestart switch" -Component "Uninstall-PlantronicsHub" -Severity 1
                    }
                }
                # Method 2: UninstallString (may need silent switches)
                elseif ($app.UninstallString) {
                    $uninstallCmd = $app.UninstallString
                    Write-CMLog -Message "Using UninstallString: $uninstallCmd" -Component "Uninstall-PlantronicsHub" -Severity 1
                    
                    # If it's MSI, make it silent
                    if ($uninstallCmd -imatch "msiexec") {
                        # Convert /I to /X if present
                        if ($uninstallCmd -imatch "/I") {
                            $uninstallCmd = $uninstallCmd -ireplace "/I", "/X"
                            Write-CMLog -Message "Converted /I to /X" -Component "Uninstall-PlantronicsHub" -Severity 1
                        }
                        
                        # Add silent switches if not present
                        if ($uninstallCmd -notmatch "/qn") {
                            $uninstallCmd += " /qn /norestart"
                            Write-CMLog -Message "Added silent switches (/qn /norestart)" -Component "Uninstall-PlantronicsHub" -Severity 1
                        }
                        elseif ($uninstallCmd -notmatch "/norestart") {
                            $uninstallCmd += " /norestart"
                            Write-CMLog -Message "Added /norestart switch" -Component "Uninstall-PlantronicsHub" -Severity 1
                        }
                    }
                }
                # Method 3: Use product code directly (MSI only)
                elseif ($productCode -match "^\{[A-F0-9-]+\}$") {
                    $uninstallCmd = "msiexec.exe /x $productCode /qn /norestart"
                    Write-CMLog -Message "Using product code for MSI uninstall: $uninstallCmd" -Component "Uninstall-PlantronicsHub" -Severity 1
                }
                
                # Execute uninstall
                if ($uninstallCmd) {
                    try {
                        Write-CMLog -Message "Executing uninstall: $uninstallCmd" -Component "Uninstall-PlantronicsHub" -Severity 1
                        
                        # Run through cmd.exe for maximum compatibility
                        $process = Start-Process -FilePath "cmd.exe" `
                            -ArgumentList "/c", $uninstallCmd `
                            -Wait `
                            -WindowStyle Hidden `
                            -PassThru `
                            -ErrorAction Stop
                        
                        # Check exit code
                        # 0 = success, 3010 = success with reboot required
                        if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
                            Write-CMLog -Message "Successfully uninstalled: $versionString (Exit Code: $($process.ExitCode))" -Component "Uninstall-PlantronicsHub" -Severity 1
                            $result.Uninstalled++
                        }
                        else {
                            Write-CMLog -Message "Uninstall failed with exit code $($process.ExitCode): $versionString" -Component "Uninstall-PlantronicsHub" -Severity 3
                            $result.Failed++
                        }
                    }
                    catch {
                        Write-CMLog -Message "Exception during uninstall: $_" -Component "Uninstall-PlantronicsHub" -Severity 3
                        Write-CMLog -Message "Command: $uninstallCmd" -Component "Uninstall-PlantronicsHub" -Severity 3
                        $result.Failed++
                    }
                }
                else {
                    Write-CMLog -Message "No uninstall method found for: $versionString" -Component "Uninstall-PlantronicsHub" -Severity 2
                    $result.Failed++
                }
            }
        }
        catch {
            Write-CMLog -Message "Error searching registry path $regPath : $_" -Component "Uninstall-PlantronicsHub" -Severity 3
        }
    }
    
    # Determine overall result
    if ($result.Found -eq 0) {
        Write-CMLog -Message "No Plantronics Hub installations found" -Component "Uninstall-PlantronicsHub" -Severity 2
        $result.Success = $false
        $result.ExitCode = 1
    }
    elseif ($result.Uninstalled -gt 0 -and $result.Failed -eq 0) {
        Write-CMLog -Message "All versions uninstalled successfully ($($result.Uninstalled) version(s))" -Component "Uninstall-PlantronicsHub" -Severity 1
        $result.Success = $true
        $result.ExitCode = 0
    }
    elseif ($result.Uninstalled -gt 0) {
        Write-CMLog -Message "Partial success: $($result.Uninstalled) succeeded, $($result.Failed) failed" -Component "Uninstall-PlantronicsHub" -Severity 2
        $result.Success = $true  # Partial success
        $result.ExitCode = 0
    }
    else {
        Write-CMLog -Message "All uninstall attempts failed" -Component "Uninstall-PlantronicsHub" -Severity 3
        $result.Success = $false
        $result.ExitCode = 1
    }
    
    Write-CMLog -Message "=== Plantronics Hub uninstall complete ===" -Component "Uninstall-PlantronicsHub" -Severity 1
    
    return $result
}
