<#
.SYNOPSIS
    Uninstalls Dell SupportAssist and related components.

.DESCRIPTION
    This script removes Dell SupportAssist and Dell SupportAssist OS Recovery Plugin for Dell Update
    from the system using the registry uninstall information.
    
    Target Applications:
    - Dell SupportAssist (Version: 4.10.6.48716)
    - Dell SupportAssist OS Recovery Plugin for Dell Update (Version: 5.5.13.0)

.NOTES
    File Name:    Uninstall.DellSupportAssist.ps1
    Author:       Vipin Ashodhiya
    Version:      1.0
    Date:         2026-02-03
    
.EXAMPLE
    Uninstall-DellSupportAssist
#>

function Uninstall-DellSupportAssist {
    [CmdletBinding()]
    param ()
    
    Write-CMLog -Message "=== Starting Dell SupportAssist-specific uninstall ===" -Component "Uninstall-DellSupportAssist" -Severity 1
    
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
            Write-CMLog -Message "Searching: $regPath" -Component "Uninstall-DellSupportAssist" -Severity 1
            
            # Filter to match Dell SupportAssist applications
            # Matches:
            # - Dell SupportAssist
            # - Dell SupportAssist OS Recovery Plugin for Dell Update
            $apps = Get-ItemProperty $regPath -ErrorAction SilentlyContinue | 
            Where-Object { 
                $_.DisplayName -like '*Dell SupportAssist*' -and
                $_.Publisher -like '*Dell*'
            }
            
            foreach ($app in $apps) {
                $result.Found++
                
                # Gather information
                $displayName = $app.DisplayName
                $displayVersion = $app.DisplayVersion
                $productCode = $app.PSChildName
                $versionString = "$displayName ($displayVersion)"
                $result.Versions += $versionString
                
                Write-CMLog -Message "Found: $versionString" -Component "Uninstall-DellSupportAssist" -Severity 1
                Write-CMLog -Message "Product Code: $productCode" -Component "Uninstall-DellSupportAssist" -Severity 1
                
                # Determine uninstall method
                $uninstallCmd = $null
                
                # Method 1: QuietUninstallString (preferred - already silent)
                if ($app.QuietUninstallString) {
                    $uninstallCmd = $app.QuietUninstallString
                    Write-CMLog -Message "Using QuietUninstallString: $uninstallCmd" -Component "Uninstall-DellSupportAssist" -Severity 1
                }
                # Method 2: UninstallString (may need silent switches)
                elseif ($app.UninstallString) {
                    $uninstallCmd = $app.UninstallString
                    Write-CMLog -Message "Using UninstallString: $uninstallCmd" -Component "Uninstall-DellSupportAssist" -Severity 1
                    
                    # Dell SupportAssist uses MSI, ensure proper uninstall command
                    if ($uninstallCmd -imatch "msiexec") {
                        # Convert /I to /X if present (some registry entries have /I instead of /X)
                        if ($uninstallCmd -imatch "/I") {
                            $uninstallCmd = $uninstallCmd -ireplace "/I", "/X"
                            Write-CMLog -Message "Converted /I to /X for proper uninstall" -Component "Uninstall-DellSupportAssist" -Severity 1
                        }
                        
                        # Add silent switches if not present
                        if ($uninstallCmd -notmatch "/qn") {
                            $uninstallCmd += " /qn /norestart"
                            Write-CMLog -Message "Added silent switches" -Component "Uninstall-DellSupportAssist" -Severity 1
                        }
                    }
                }
                # Method 3: Use product code directly (MSI only)
                elseif ($productCode -match "^\{[A-F0-9-]+\}$") {
                    $uninstallCmd = "msiexec.exe /x $productCode /qn /norestart"
                    Write-CMLog -Message "Using product code for MSI uninstall: $uninstallCmd" -Component "Uninstall-DellSupportAssist" -Severity 1
                }
                
                # Execute uninstall
                if ($uninstallCmd) {
                    try {
                        Write-CMLog -Message "Executing uninstall: $uninstallCmd" -Component "Uninstall-DellSupportAssist" -Severity 1
                        
                        # Run through cmd.exe for MSI compatibility
                        $process = Start-Process -FilePath "cmd.exe" `
                            -ArgumentList "/c", $uninstallCmd `
                            -Wait `
                            -WindowStyle Hidden `
                            -PassThru `
                            -ErrorAction Stop
                        
                        # Check exit code
                        # Common codes: 0 = success, 3010 = success with reboot required
                        if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
                            Write-CMLog -Message "Successfully uninstalled: $versionString (Exit Code: $($process.ExitCode))" -Component "Uninstall-DellSupportAssist" -Severity 1
                            $result.Uninstalled++
                        }
                        else {
                            Write-CMLog -Message "Uninstall failed with exit code $($process.ExitCode): $versionString" -Component "Uninstall-DellSupportAssist" -Severity 3
                            $result.Failed++
                        }
                    }
                    catch {
                        Write-CMLog -Message "Exception during uninstall: $_" -Component "Uninstall-DellSupportAssist" -Severity 3
                        Write-CMLog -Message "Command: $uninstallCmd" -Component "Uninstall-DellSupportAssist" -Severity 3
                        $result.Failed++
                    }
                }
                else {
                    Write-CMLog -Message "No uninstall method found for: $versionString" -Component "Uninstall-DellSupportAssist" -Severity 2
                    $result.Failed++
                }
            }
        }
        catch {
            Write-CMLog -Message "Error searching registry path $regPath : $_" -Component "Uninstall-DellSupportAssist" -Severity 3
        }
    }
    
    # Determine overall result (DO NOT MODIFY THIS SECTION)
    if ($result.Found -eq 0) {
        Write-CMLog -Message "No Dell SupportAssist installations found" -Component "Uninstall-DellSupportAssist" -Severity 2
        $result.Success = $false
        $result.ExitCode = 1
    }
    elseif ($result.Uninstalled -gt 0 -and $result.Failed -eq 0) {
        Write-CMLog -Message "All Dell SupportAssist versions uninstalled successfully ($($result.Uninstalled) version(s))" -Component "Uninstall-DellSupportAssist" -Severity 1
        $result.Success = $true
        $result.ExitCode = 0
    }
    elseif ($result.Uninstalled -gt 0) {
        Write-CMLog -Message "Partial success: $($result.Uninstalled) succeeded, $($result.Failed) failed" -Component "Uninstall-DellSupportAssist" -Severity 2
        $result.Success = $true  # Partial success
        $result.ExitCode = 0
    }
    else {
        Write-CMLog -Message "All Dell SupportAssist uninstall attempts failed" -Component "Uninstall-DellSupportAssist" -Severity 3
        $result.Success = $false
        $result.ExitCode = 1
    }
    
    Write-CMLog -Message "=== Dell SupportAssist Uninstall complete ===" -Component "Uninstall-DellSupportAssist" -Severity 1
    
    return $result
}

# Known Product Codes for Dell SupportAssist:
# -------------------------------------------
# Dell SupportAssist (4.10.6.48716):
#   Product Code: {00B9D238-5A72-4BD9-B2DC-0144D8DBD5C8}
#   Uninstall: MsiExec.exe /X{00B9D238-5A72-4BD9-B2DC-0144D8DBD5C8}
#
# Dell SupportAssist OS Recovery Plugin for Dell Update (5.5.13.0):
#   Product Code: {F5391400-4596-46A6-9D3C-9D7647230679}
#   Uninstall: MsiExec.exe /X{F5391400-4596-46A6-9D3C-9D7647230679}
#   Note: Registry may show /I but script converts to /X for proper uninstall
