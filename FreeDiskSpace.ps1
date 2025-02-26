<# 
.SYNOPSIS 
FreeDiskSpace.ps1 - Frees up disk space by cleaning temp folders. 
 
.DESCRIPTION  
This script will free up disk space by cleaning temp folders.

.NOTES 
    FreeDiskSpace.ps1 - V.Ashodhiya - 16-01-2025
    Script History:
    Version 1.0 - Script inception
 
#>

# Define the path for the log file
$logFilePath = "C:\Windows\fndr\logs"
$logFileName = "$logFilePath\FreeDiskSpace.log"
# Function to write logs
function Write-Log{
    param (
        [string]$message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $message"
    Write-Logut $logMessage
 
    # Ensure log file path exists
    if (-not (Test-Path $logFilePath)) {
        New-Item -Path $logFilePath -ItemType Directory | Out-Null
    }
 
    # Write Log Write-Logmessage to log file
    Add-Content -Path $logFileName -Value $logMessage
}

Write-Log "1. Stopping Windows Update Services..."
Write-Host "1. Stopping Windows Update Services..."
Stop-Service -Name BITS | Out-Null
Stop-Service -Name wuauserv | Out-Null
Stop-Service -Name appidsvc | Out-Null
Stop-Service -Name cryptsvc | Out-Null

Write-Log "2. Remove QMGR Data file..." 
Write-Host "2. Remove QMGR Data file..." 
Remove-Item "$env:allusersprofile\Application Data\Microsoft\Network\Downloader\qmgr*.dat" -ErrorAction SilentlyContinue 
 
Write-Log "3. Removing the Software Distribution and CatRoot Folder..." 
Write-Host "3. Removing the Software Distribution and CatRoot Folder..." 
Remove-Item $env:systemroot\SoftwareDistribution\DataStore -ErrorAction SilentlyContinue
Remove-Item $env:systemroot\SoftwareDistribution\Download -ErrorAction SilentlyContinue 
Remove-Item $env:systemroot\System32\Catroot2 -ErrorAction SilentlyContinue 
 
Write-Log "4. Removing old Windows Update log..." 
Write-Host "4. Removing old Windows Update log..." 
Remove-Item $env:systemroot\WindowsUpdate.log -ErrorAction SilentlyContinue 
 
Write-Log "5. Resetting the Windows Update Services to default settings..." 
Write-Host "5. Resetting the Windows Update Services to default settings..." 
Start-Process -FilePath "$env:systemroot\system32\sc.exe" -ArgumentList "sdset bits D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)" -Wait
Start-Process -FilePath "$env:systemroot\system32\sc.exe" -ArgumentList "sdset wuauserv D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)" -Wait

Write-Log "5. Removing Windows Temp File..."
Write-Host "5. Removing Windows Temp File..."
Get-ChildItem -Path C:\windows\Temp -File | Remove-Item -Verbose -Force

Write-Log "6. Delete all BITS jobs..." 
Write-Host "6. Delete all BITS jobs..." 
Get-BitsTransfer | Remove-BitsTransfer 
 
 
Write-Log "7. Starting Windows Update Services..." 
Write-Host "7. Starting Windows Update Services..." 
Start-Service -Name BITS | Out-Null
Start-Service -Name wuauserv | Out-Null
Start-Service -Name appidsvc | Out-Null
Start-Service -Name cryptsvc | Out-Null
 
Write-Log "8. Forcing Inventory and Discovery..." 
Write-Host "8. Forcing Inventory and Discovery..." 
wuauclt /resetauthorization /detectnow

$comp = ”localhost”

$HardwareInventoryID = ‘{00000000-0000-0000-0000-000000000001}’

$HeartbeatID = ‘{00000000-0000-0000-0000-000000000003}’

Get-WmiObject -ComputerName $comp -Namespace ‘Root\CCM\INVAGT’ -Class ‘InventoryActionStatus’ -Filter “InventoryActionID=’$HardwareInventoryID'” | Remove-WmiObject

Get-WmiObject -ComputerName $comp -Namespace ‘Root\CCM\INVAGT’ -Class ‘InventoryActionStatus’ -Filter “InventoryActionID=’$HeartbeatID'” | Remove-WmiObject

Start-Sleep -s 5
Write-Log "Running Heartbeat Discovery"
Write-Host "Running Heartbeat Discovery"
Invoke-WmiMethod -computername $comp -Namespace root\CCM -Class SMS_Client -Name TriggerSchedule -ArgumentList $HeartbeatID
Write-Log "Running Hardware Inventory"
Write-Host "Running Heartbeat Discovery"
Invoke-WmiMethod -computername $comp -Namespace root\CCM -Class SMS_Client -Name TriggerSchedule -ArgumentList $HardwareInventoryID