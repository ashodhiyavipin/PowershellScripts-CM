<#
.SYNOPSIS
    App-specific uninstaller for Adobe Acrobat Reader.

.DESCRIPTION
    This script provides a specialized uninstaller for all versions of Adobe Acrobat Reader.
    Based on the existing UninstallAdobeAcrobat.ps1 script, it searches the registry for 
    all installed Adobe Acrobat versions and uninstalls them.
    
    You can optionally exclude specific editions (Standard, Professional) or versions
    by modifying the filter logic in this function.

.NOTES
    File Name:    Uninstall.AdobeReader.ps1
    Author:       V.Ashodhiya
    Version:      1.0
    Date:         2026-01-27
    Based On:     UninstallAdobeAcrobat.ps1
    
.EXAMPLE
    Uninstall-AdobeReader
#>

function Uninstall-AdobeReader {
    [CmdletBinding()]
    param ()
    
    Write-CMLog -Message "=== Starting Adobe Reader-specific uninstall ===" -Component "Uninstall-AdobeReader" -Severity 1
    
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
    
    # Search for all Adobe Acrobat/Reader installations
    foreach ($regPath in $registryPaths) {
        try {
            Write-CMLog -Message "Searching: $regPath" -Component "Uninstall-AdobeReader" -Severity 1
            
            $apps = Get-ItemProperty $regPath -ErrorAction SilentlyContinue | 
            Where-Object { 
                $_.DisplayName -like '*Adobe Acrobat*' 
                # Optionally exclude specific editions or versions:
                # -and $_.DisplayName -notlike '*Standard*' 
                # -and $_.DisplayName -notlike '*Professional*'
                # -and $_.DisplayVersion -notlike '*24.002.20965*'
            }
            
            foreach ($app in $apps) {
                $result.Found++
                
                $displayName = $app.DisplayName
                $displayVersion = $app.DisplayVersion
                $productCode = $app.PSChildName
                $versionString = "$displayName ($displayVersion)"
                $result.Versions += $versionString
                
                Write-CMLog -Message "Found Adobe: $versionString" -Component "Uninstall-AdobeReader" -Severity 1
                Write-CMLog -Message "Product Code: $productCode" -Component "Uninstall-AdobeReader" -Severity 1
                
                # Adobe Reader is typically installed via MSI
                # Use product code for uninstall
                if ($productCode -match "^\{[A-F0-9-]+\}$") {
                    $uninstallCmd = "msiexec.exe /x $productCode /qn /norestart"
                    Write-CMLog -Message "Using MSI uninstall: $uninstallCmd" -Component "Uninstall-AdobeReader" -Severity 1
                    
                    try {
                        Write-CMLog -Message "Executing uninstall..." -Component "Uninstall-AdobeReader" -Severity 1
                        
                        $process = Start-Process -FilePath "msiexec.exe" `
                            -ArgumentList "/x", $productCode, "/qn", "/norestart" `
                            -Wait `
                            -WindowStyle Hidden `
                            -PassThru `
                            -ErrorAction Stop
                        
                        if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
                            # 0 = success, 3010 = success with reboot required
                            Write-CMLog -Message "Successfully uninstalled: $versionString (Exit Code: $($process.ExitCode))" -Component "Uninstall-AdobeReader" -Severity 1
                            $result.Uninstalled++
                        }
                        else {
                            Write-CMLog -Message "Uninstall failed with exit code $($process.ExitCode): $versionString" -Component "Uninstall-AdobeReader" -Severity 3
                            $result.Failed++
                        }
                    }
                    catch {
                        Write-CMLog -Message "Exception during uninstall: $_" -Component "Uninstall-AdobeReader" -Severity 3
                        Write-CMLog -Message "Product Code: $productCode" -Component "Uninstall-AdobeReader" -Severity 3
                        $result.Failed++
                    }
                }
                else {
                    # Try using UninstallString if product code doesn't match GUID pattern
                    $uninstallCmd = $null
                    
                    if ($app.QuietUninstallString) {
                        $uninstallCmd = $app.QuietUninstallString
                        Write-CMLog -Message "Using QuietUninstallString: $uninstallCmd" -Component "Uninstall-AdobeReader" -Severity 1
                    }
                    elseif ($app.UninstallString) {
                        $uninstallCmd = $app.UninstallString
                        Write-CMLog -Message "Using UninstallString: $uninstallCmd" -Component "Uninstall-AdobeReader" -Severity 1
                        
                        # Add silent switches if MSI
                        if ($uninstallCmd -imatch "msiexec") {
                            if ($uninstallCmd -imatch "/I") {
                                $uninstallCmd = $uninstallCmd -ireplace "/I", "/X"
                            }
                            if ($uninstallCmd -notmatch "/qn") {
                                $uninstallCmd += " /qn /norestart"
                            }
                        }
                    }
                    
                    if ($uninstallCmd) {
                        try {
                            Write-CMLog -Message "Executing: $uninstallCmd" -Component "Uninstall-AdobeReader" -Severity 1
                            
                            $process = Start-Process -FilePath "cmd.exe" `
                                -ArgumentList "/c", $uninstallCmd `
                                -Wait `
                                -WindowStyle Hidden `
                                -PassThru `
                                -ErrorAction Stop
                            
                            if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
                                Write-CMLog -Message "Successfully uninstalled: $versionString (Exit Code: $($process.ExitCode))" -Component "Uninstall-AdobeReader" -Severity 1
                                $result.Uninstalled++
                            }
                            else {
                                Write-CMLog -Message "Uninstall failed with exit code $($process.ExitCode): $versionString" -Component "Uninstall-AdobeReader" -Severity 3
                                $result.Failed++
                            }
                        }
                        catch {
                            Write-CMLog -Message "Exception during uninstall: $_" -Component "Uninstall-AdobeReader" -Severity 3
                            $result.Failed++
                        }
                    }
                    else {
                        Write-CMLog -Message "No uninstall method found for: $versionString" -Component "Uninstall-AdobeReader" -Severity 2
                        $result.Failed++
                    }
                }
            }
        }
        catch {
            Write-CMLog -Message "Error searching registry path $regPath : $_" -Component "Uninstall-AdobeReader" -Severity 3
        }
    }
    
    # Determine overall result
    if ($result.Found -eq 0) {
        Write-CMLog -Message "No Adobe Acrobat/Reader installations found" -Component "Uninstall-AdobeReader" -Severity 2
        $result.Success = $false
        $result.ExitCode = 1
    }
    elseif ($result.Uninstalled -gt 0 -and $result.Failed -eq 0) {
        Write-CMLog -Message "All Adobe versions uninstalled successfully ($($result.Uninstalled) version(s))" -Component "Uninstall-AdobeReader" -Severity 1
        $result.Success = $true
        $result.ExitCode = 0
    }
    elseif ($result.Uninstalled -gt 0) {
        Write-CMLog -Message "Partial success: $($result.Uninstalled) succeeded, $($result.Failed) failed" -Component "Uninstall-AdobeReader" -Severity 2
        $result.Success = $true  # Partial success
        $result.ExitCode = 0
    }
    else {
        Write-CMLog -Message "All Adobe uninstall attempts failed" -Component "Uninstall-AdobeReader" -Severity 3
        $result.Success = $false
        $result.ExitCode = 1
    }
    
    Write-CMLog -Message "=== Adobe Reader uninstall complete ===" -Component "Uninstall-AdobeReader" -Severity 1
    
    return $result
}
