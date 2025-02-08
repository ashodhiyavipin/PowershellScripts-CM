<# 
.SYNOPSIS
Automates the cleanup of a specified Software Update Group (SUG) and triggers the Auto Deployment Rule (ADR) for it.
 
.DESCRIPTION
This script:
1. Retrieves the updates in a specified Software Update Group (SUG).
2. Removes all updates from the SUG.
3. Verifies that the SUG is empty.
4. Triggers an ADR for the SUG to repopulate updates.
 
.NOTES
Author: Vipin A.
Date: 20/11/2024
Version: 1.0
#>

#---------------------- Code to Connect to CAS SMS Provider ---------------------------------
#
# Press 'F5' to run this script. Running this script will load the ConfigurationManager
# module for Windows PowerShell and will connect to the site.
#
# This script was auto-generated at '11/20/2024 1:58:50 AM'.

# Site configuration
$SiteCode = "CAS" # Site code 
$ProviderMachineName = "us145sccmcas.nac.sitel-world.net" # SMS Provider machine name

# Customizations
$initParams = @{}
$initParams.Add("Verbose", $true) # Uncomment this line to enable verbose logging
$initParams.Add("ErrorAction", "Stop") # Uncomment this line to stop the script on any errors

# Do not change anything below this line

# Import the ConfigurationManager.psd1 module 
if((Get-Module ConfigurationManager) -eq $null) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
}

# Connect to the site's drive if it is not already present
if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
}

# Set the current location to be the site code.
Set-Location "$($SiteCode):\" @initParams
#----------------------------- Code to Connect to CAS SMS Provider ---------------------------------

# ----------------------------------- Script Starts Here -------------------------------------------
# Define the Software Update Group (SUG) name
$SUG = "Your_SUG_Name"  # Replace with the actual SUG name
 
# Function to write logs
function Write-Log {
    param (
        [string]$message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $message"
    Write-Output $logMessage
 
    # Ensure log file path exists
    if (-not (Test-Path $logFilePath)) {
        New-Item -Path $logFilePath -ItemType Directory | Out-Null
    }
 
    # Write log message to log file
    Add-Content -Path $logFileName -Value $logMessage
}
 
try {
    # Step 1: Get Software Update Group Information
    Write-Log "Retrieving information for Software Update Group: $SUG"
    $SUGInfo = Get-CMSoftwareUpdateGroup -Name $SUG
 
    if (-not $SUGInfo) {
        Write-Log "Software Update Group '$SUG' not found." -ErrorAction Stop
    }
 
    # Extract Software Update IDs
    Write-Log "Extracting Software Update IDs from the SUG..."
    $SoftwareUpdateIDs = $SUGInfo.Updates
    $softwareUpdateGroupIDs = $SUGInfo.CI_ID
    Write-Log "Software Update Group ID is $softwareUpdateGroupIDs"

    if (-not $SoftwareUpdateIDs) {
        Write-Log "No updates found in Software Update Group '$SUG'."
    } else {
        Write-Log "Found updates to remove: $($SoftwareUpdateIDs -join ', ')"
    }

    # Step 2: Remove updates from the SUG
    Write-Log "Removing updates from Software Update Group: $SUG"
    foreach ($SoftwareUpdateID in $SoftwareUpdateIDs) {
        Remove-CMSoftwareUpdateFromGroup -SoftwareUpdateGroupID $softwareUpdateGroupIDs -SoftwareUpdateID $SoftwareUpdateID -ErrorAction Stop
    }
    Write-Log "All updates removed successfully from SUG: $SUG."

    # Step 3: Verify the SUG is empty
    Write-Log "Verifying that the SUG is empty..."
    $SUGInfo = Get-CMSoftwareUpdateGroup -Name $SUG
    $RemainingUpdates = $SUGInfo.Updates.Count

    if ($RemainingUpdates -eq 0) {
        Write-Log "Verification successful: No updates remain in SUG '$SUG'."
    } else {
        Write-Log "Verification failed: $RemainingUpdates updates still remain in SUG '$SUG'." -ErrorAction Stop
    }
 
    # Step 4: Trigger ADR for the SUG
    Write-Log "Triggering Auto Deployment Rule (ADR) for SUG: $SUG"
    Invoke-CMSoftwareUpdateAutoDeploymentRule -Name $SUG -ErrorAction Stop
    Write-Log "ADR triggered successfully for SUG: $SUG."
 
    # Return success
    Write-Log "Script completed successfully."
    exit 0
 
} catch {
    # Log and return failure
    Write-Log "Error: $_"
    exit 1
}