<#
.SYNOPSIS
    App-specific uninstaller for KeePass Password Safe.

.DESCRIPTION
    This script provides a specialized uninstaller for all versions of KeePass.
    It searches the registry for all installed KeePass versions and uninstalls them.

.NOTES
    File Name:    Uninstall.KeePass.ps1
    Author:       V.Ashodhiya
    Version:      1.0
    Date:         2026-01-27
    
.EXAMPLE
    Uninstall-KeePass
#>

function Uninstall-KeePass {
    [CmdletBinding()]
    param ()
    
    Write-CMLog -Message "=== Starting KeePass-specific uninstall ===" -Component "Uninstall-KeePass" -Severity 1
    
    $registryPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    $result = @{
        Success     = $false
        ExitCode    = 1
        Found       = 0
        Uninstalled = 0
        Failed      = 0
        Versions    = @()
    }
    
    # Search for all KeePass installations
    foreach ($regPath in $registryPaths) {
        try {
            Write-CMLog -Message "Searching: $regPath" -Component "Uninstall-KeePass" -Severity 1
            
            $apps = Get-ItemProperty $regPath -ErrorAction SilentlyContinue | 
            Where-Object { 
                $_.DisplayName -like '*KeePass*' 
            }
            
            foreach ($app in $apps) {
                $result.Found++
                
                $displayName = $app.DisplayName
                $displayVersion = $app.DisplayVersion
                $productCode = $app.PSChildName
                $versionString = "$displayName ($displayVersion)"
                $result.Versions += $versionString
                
                Write-CMLog -Message "Found KeePass: $versionString" -Component "Uninstall-KeePass" -Severity 1
                Write-CMLog -Message "Product Code: $productCode" -Component "Uninstall-KeePass" -Severity 1
                
                # Determine uninstall method
                $uninstallCmd = $null
                
                # Prefer QuietUninstallString
                if ($app.QuietUninstallString) {
                    $uninstallCmd = $app.QuietUninstallString
                    Write-CMLog -Message "Using QuietUninstallString: $uninstallCmd" -Component "Uninstall-KeePass" -Severity 1
                }
                # Try UninstallString
                elseif ($app.UninstallString) {
                    $uninstallCmd = $app.UninstallString
                    Write-CMLog -Message "Using UninstallString: $uninstallCmd" -Component "Uninstall-KeePass" -Severity 1
                    
                    # If it's an MSI, make it silent
                    if ($uninstallCmd -match "msiexec", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase) {
                        # Convert /I to /X if present
                        if ($uninstallCmd -imatch "/I") {
                            $uninstallCmd = $uninstallCmd -ireplace "/I", "/X"
                            Write-CMLog -Message "Converted /I to /X" -Component "Uninstall-KeePass" -Severity 1
                        }
                        
                        # Add silent switches
                        if ($uninstallCmd -notmatch "/qn") {
                            $uninstallCmd += " /qn /norestart"
                            Write-CMLog -Message "Added silent switches" -Component "Uninstall-KeePass" -Severity 1
                        }
                    }
                }
                # Try using product code directly
                elseif ($productCode -match "^\{[A-F0-9-]+\}$") {
                    $uninstallCmd = "msiexec.exe /x $productCode /qn /norestart"
                    Write-CMLog -Message "Using product code for MSI uninstall: $uninstallCmd" -Component "Uninstall-KeePass" -Severity 1
                }
                
                if ($uninstallCmd) {
                    try {
                        Write-CMLog -Message "Executing uninstall: $uninstallCmd" -Component "Uninstall-KeePass" -Severity 1
                        
                        $process = Start-Process -FilePath "cmd.exe" `
                            -ArgumentList "/c", $uninstallCmd `
                            -Wait `
                            -WindowStyle Hidden `
                            -PassThru `
                            -ErrorAction Stop
                        
                        if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
                            # 0 = success, 3010 = success with reboot required
                            Write-CMLog -Message "Successfully uninstalled: $versionString (Exit Code: $($process.ExitCode))" -Component "Uninstall-KeePass" -Severity 1
                            $result.Uninstalled++
                        }
                        else {
                            Write-CMLog -Message "Uninstall failed with exit code $($process.ExitCode): $versionString" -Component "Uninstall-KeePass" -Severity 3
                            $result.Failed++
                        }
                    }
                    catch {
                        Write-CMLog -Message "Exception during uninstall: $_" -Component "Uninstall-KeePass" -Severity 3
                        Write-CMLog -Message "Command: $uninstallCmd" -Component "Uninstall-KeePass" -Severity 3
                        $result.Failed++
                    }
                }
                else {
                    Write-CMLog -Message "No uninstall method found for: $versionString" -Component "Uninstall-KeePass" -Severity 2
                    $result.Failed++
                }
            }
        }
        catch {
            Write-CMLog -Message "Error searching registry path $regPath : $_" -Component "Uninstall-KeePass" -Severity 3
        }
    }
    
    # Determine overall result
    if ($result.Found -eq 0) {
        Write-CMLog -Message "No KeePass installations found" -Component "Uninstall-KeePass" -Severity 2
        $result.Success = $false
        $result.ExitCode = 1
    }
    elseif ($result.Uninstalled -gt 0 -and $result.Failed -eq 0) {
        Write-CMLog -Message "All KeePass versions uninstalled successfully ($($result.Uninstalled) version(s))" -Component "Uninstall-KeePass" -Severity 1
        $result.Success = $true
        $result.ExitCode = 0
    }
    elseif ($result.Uninstalled -gt 0) {
        Write-CMLog -Message "Partial success: $($result.Uninstalled) succeeded, $($result.Failed) failed" -Component "Uninstall-KeePass" -Severity 2
        $result.Success = $true  # Partial success
        $result.ExitCode = 0
    }
    else {
        Write-CMLog -Message "All KeePass uninstall attempts failed" -Component "Uninstall-KeePass" -Severity 3
        $result.Success = $false
        $result.ExitCode = 1
    }
    
    Write-CMLog -Message "=== KeePass uninstall complete ===" -Component "Uninstall-KeePass" -Severity 1
    
    return $result
}
