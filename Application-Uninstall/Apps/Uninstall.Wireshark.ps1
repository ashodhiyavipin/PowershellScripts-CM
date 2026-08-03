<#
.SYNOPSIS
    Wireshark-specific uninstaller function.

.DESCRIPTION
    Uninstalls all versions of Wireshark network protocol analyzer.
    Wireshark uses NSIS installer, so the silent switch is /S.
    This function handles both 32-bit and 64-bit versions.

.NOTES
    File Name:    Uninstall.Wireshark.ps1
    Author:       V.Ashodhiya
    Version:      1.0
    Date:         2026-02-06
    
.EXAMPLE
    Uninstall-Wireshark
#>

function Uninstall-Wireshark {
    [CmdletBinding()]
    param ()
    
    Write-CMLog -Message "=== Starting Wireshark-specific uninstall ===" -Component "Uninstall-Wireshark" -Severity 1
    
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
    
    # Search for all Wireshark installations
    foreach ($regPath in $registryPaths) {
        try {
            Write-CMLog -Message "Searching: $regPath" -Component "Uninstall-Wireshark" -Severity 1
            
            # Match Wireshark applications (main app and components)
            # Wireshark may appear as "Wireshark X.X.X 64-bit" or "Wireshark X.X.X"
            $apps = Get-ItemProperty $regPath -ErrorAction SilentlyContinue | 
            Where-Object { 
                $_.DisplayName -like '*Wireshark*'
            }
            
            foreach ($app in $apps) {
                $result.Found++
                
                # Gather information
                $displayName = $app.DisplayName
                $displayVersion = $app.DisplayVersion
                $productCode = $app.PSChildName
                $versionString = "$displayName ($displayVersion)"
                $result.Versions += $versionString
                
                Write-CMLog -Message "Found: $versionString" -Component "Uninstall-Wireshark" -Severity 1
                Write-CMLog -Message "Product Code: $productCode" -Component "Uninstall-Wireshark" -Severity 1
                
                # Determine uninstall method
                $uninstallCmd = $null
                
                # Method 1: QuietUninstallString (preferred - already silent)
                if ($app.QuietUninstallString) {
                    $uninstallCmd = $app.QuietUninstallString
                    Write-CMLog -Message "Using QuietUninstallString: $uninstallCmd" -Component "Uninstall-Wireshark" -Severity 1
                }
                # Method 2: UninstallString (needs silent switches for NSIS)
                elseif ($app.UninstallString) {
                    $rawUninstallString = $app.UninstallString
                    Write-CMLog -Message "Using UninstallString: $rawUninstallString" -Component "Uninstall-Wireshark" -Severity 1
                    
                    # Wireshark uses NSIS installer - /S for silent mode
                    # Handle quoted paths properly
                    if ($rawUninstallString -match '^"(.+)"(.*)$') {
                        # Quoted path: extract path and existing args
                        $uninstallExe = $matches[1]
                        $existingArgs = $matches[2].Trim()
                        
                        # Build the command - add /S if not already present
                        if ($existingArgs -notmatch '/S') {
                            $uninstallCmd = "`"$uninstallExe`" /S $existingArgs"
                        }
                        else {
                            $uninstallCmd = "`"$uninstallExe`" $existingArgs"
                        }
                    }
                    elseif ($rawUninstallString -match '^(.+\.exe)\s*(.*)$') {
                        # Unquoted path
                        $uninstallExe = $matches[1]
                        $existingArgs = $matches[2].Trim()
                        
                        # Build the command - add /S if not already present
                        if ($existingArgs -notmatch '/S') {
                            $uninstallCmd = "`"$uninstallExe`" /S $existingArgs"
                        }
                        else {
                            $uninstallCmd = "`"$uninstallExe`" $existingArgs"
                        }
                    }
                    else {
                        # Fallback if pattern doesn't match
                        $uninstallCmd = "$rawUninstallString /S"
                    }
                    
                    Write-CMLog -Message "Built silent command: $uninstallCmd" -Component "Uninstall-Wireshark" -Severity 1
                }
                # Method 3: Use product code directly (MSI only)
                elseif ($productCode -match "^\{[A-F0-9-]+\}$") {
                    $uninstallCmd = "msiexec.exe /x $productCode /qn /norestart"
                    Write-CMLog -Message "Using product code for MSI uninstall: $uninstallCmd" -Component "Uninstall-Wireshark" -Severity 1
                }
                
                # Execute uninstall
                if ($uninstallCmd) {
                    try {
                        Write-CMLog -Message "Executing uninstall: $uninstallCmd" -Component "Uninstall-Wireshark" -Severity 1
                        
                        # Run through cmd.exe for compatibility
                        $process = Start-Process -FilePath "cmd.exe" `
                            -ArgumentList "/c", $uninstallCmd `
                            -Wait `
                            -WindowStyle Hidden `
                            -PassThru `
                            -ErrorAction Stop
                        
                        # Check exit code
                        # NSIS returns 0 on success
                        if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
                            Write-CMLog -Message "Successfully uninstalled: $versionString (Exit Code: $($process.ExitCode))" -Component "Uninstall-Wireshark" -Severity 1
                            $result.Uninstalled++
                        }
                        else {
                            Write-CMLog -Message "Uninstall failed with exit code $($process.ExitCode): $versionString" -Component "Uninstall-Wireshark" -Severity 3
                            $result.Failed++
                        }
                    }
                    catch {
                        Write-CMLog -Message "Exception during uninstall: $_" -Component "Uninstall-Wireshark" -Severity 3
                        Write-CMLog -Message "Command: $uninstallCmd" -Component "Uninstall-Wireshark" -Severity 3
                        $result.Failed++
                    }
                }
                else {
                    Write-CMLog -Message "No uninstall method found for: $versionString" -Component "Uninstall-Wireshark" -Severity 2
                    $result.Failed++
                }
            }
        }
        catch {
            Write-CMLog -Message "Error searching registry path $regPath : $_" -Component "Uninstall-Wireshark" -Severity 3
        }
    }
    
    # Determine overall result
    if ($result.Found -eq 0) {
        Write-CMLog -Message "No Wireshark installations found" -Component "Uninstall-Wireshark" -Severity 2
        $result.Success = $false
        $result.ExitCode = 1
    }
    elseif ($result.Uninstalled -gt 0 -and $result.Failed -eq 0) {
        Write-CMLog -Message "All Wireshark versions uninstalled successfully ($($result.Uninstalled) version(s))" -Component "Uninstall-Wireshark" -Severity 1
        $result.Success = $true
        $result.ExitCode = 0
    }
    elseif ($result.Uninstalled -gt 0) {
        Write-CMLog -Message "Partial success: $($result.Uninstalled) succeeded, $($result.Failed) failed" -Component "Uninstall-Wireshark" -Severity 2
        $result.Success = $true  # Partial success
        $result.ExitCode = 0
    }
    else {
        Write-CMLog -Message "All uninstall attempts failed" -Component "Uninstall-Wireshark" -Severity 3
        $result.Success = $false
        $result.ExitCode = 1
    }
    
    Write-CMLog -Message "=== Wireshark uninstall complete ===" -Component "Uninstall-Wireshark" -Severity 1
    
    return $result
}
