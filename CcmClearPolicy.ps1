<# 
.Synopsis
This PowerShell script is designed to reset and update Group Policy settings on a Windows machine.

.Description
Removes the Group Policy registry file to clear existing policies.
Pauses execution for 500 milliseconds to ensure the file removal is processed.
Resets the policy using the Invoke-CIMMethod cmdlet.
Forces a Group Policy update using gpupdate.exe.
Triggers machine policy assignments and evaluations using Invoke-CIMMethod with specific schedule IDs.
Sends any unsent state messages to ensure the system's state is up-to-date.
Description
The script is useful for administrators who need to reset and reapply Group Policy settings on a machine. It ensures that any changes made to Group Policy are enforced immediately and that the system's policy state is consistent with the desired configuration.

.Notes
Version 1.0: Initial version of the script.
Date: March 27, 2025
Author: Vipin A.
Changes: Created the script to automate the process of resetting and updating Group Policy settings. #>
#---------------------------------------------------------------------#
# Define the path for the log file
$logFilePath = "C:\Windows\fndr\logs"
$logFileName = "$logFilePath\CcmClearPolicy.log"
# Function to write logs
function Write-Log{
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
 
    # Write Log Write-Logmessage to log file
    Add-Content -Path $logFileName -Value $logMessage
}

# ---------------------------------------Script Main Body ------------------------------------------#
Write-Log "Removing the registry.pol file"
Remove-Item -Path 'C:\Windows\System32\GroupPolicy\Machine\Registry.pol' -Force
[System.Threading.Thread]::Sleep(500)

# Reset Policy
Write-Log "Resetting ConfigMgr Policies"
Invoke-CIMMethod -Namespace root\ccm -ClassName SMS_CLIENT -MethodName "ResetPolicy" -Arguments @{ uFlags = [uint32]1}

# Group Policy Update
Write-Log "Fetching fresh group policy and applying them"
gpupdate.exe -Force

# Machine Policy Assignments Request
Write-Log "Fetching fresh ConfigMgr policies and applying them"
Invoke-CIMMethod -Namespace root\ccm -ClassName SMS_CLIENT -MethodName "TriggerSchedule" -Arguments @{ sScheduleID = '{00000000-0000-0000-0000-000000000021}'}

# Machine Policy Evaluation
Invoke-CIMMethod -Namespace root\ccm -ClassName SMS_CLIENT -MethodName "TriggerSchedule" -Arguments @{ sScheduleID = '{00000000-0000-0000-0000-000000000022}'}

# Send Unsent State Message
Invoke-CIMMethod -Namespace root\ccm -ClassName SMS_CLIENT -MethodName "TriggerSchedule" -Arguments @{ sScheduleID = '{00000000-0000-0000-0000-000000000111}'}