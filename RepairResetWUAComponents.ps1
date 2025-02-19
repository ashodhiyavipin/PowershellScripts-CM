<# 
.SYNOPSIS 
RepairResetWUAComponents.ps1 - Resets the Windows Update components
 
.DESCRIPTION  
This script will reset all of the Windows Updates components to DEFAULT SETTINGS. 
 
.OUTPUTS 
Results are printed to the console. Future releases will support outputting to a log file.  
 
.NOTES 
https://docs.microsoft.com/en-us/windows/deployment/update/windows-update-resources
    
    RepairResetWUAComponents.ps1 - V.Ashodhiya - 25/10/2024
    Script History:
    Version 1.0 - Script inception
 
#>
 
Write-Output "1. Stopping Windows Update Services..." 
Stop-Service -Name BITS | Out-Null
Stop-Service -Name wuauserv | Out-Null
Stop-Service -Name appidsvc | Out-Null
Stop-Service -Name cryptsvc | Out-Null
 
Write-Output "2. Remove QMGR Data file..." 
Remove-Item "$env:allusersprofile\Application Data\Microsoft\Network\Downloader\qmgr*.dat" -ErrorAction SilentlyContinue 
 
Write-Output "3. Removing the Software Distribution and CatRoot Folder..." 
Remove-Item $env:systemroot\SoftwareDistribution\DataStore -ErrorAction SilentlyContinue
Remove-Item $env:systemroot\SoftwareDistribution\Download -ErrorAction SilentlyContinue 
Remove-Item $env:systemroot\System32\Catroot2 -ErrorAction SilentlyContinue 
 
Write-Output "4. Removing old Windows Update log..." 
Remove-Item $env:systemroot\WindowsUpdate.log -ErrorAction SilentlyContinue 
 
Write-Output "5. Resetting the Windows Update Services to default settings..." 
Start-Process -FilePath "$env:systemroot\system32\sc.exe" -ArgumentList "sdset bits D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)" -Wait
Start-Process -FilePath "$env:systemroot\system32\sc.exe" -ArgumentList "sdset wuauserv D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)" -Wait

Write-Output "5. Removing Windows Temp File..."
Get-ChildItem -Path C:\windows\Temp -File | Remove-Item -Verbose -Force

 
Set-Location $env:systemroot\system32 
 
Write-Output "6. Registering some DLLs..." 
regsvr32.exe /s atl.dll 
regsvr32.exe /s urlmon.dll 
regsvr32.exe /s mshtml.dll 
regsvr32.exe /s shdocvw.dll 
regsvr32.exe /s browseui.dll 
regsvr32.exe /s jscript.dll 
regsvr32.exe /s vbscript.dll 
regsvr32.exe /s scrrun.dll 
regsvr32.exe /s msxml.dll 
regsvr32.exe /s msxml3.dll 
regsvr32.exe /s msxml6.dll 
regsvr32.exe /s actxprxy.dll 
regsvr32.exe /s softpub.dll 
regsvr32.exe /s wintrust.dll 
regsvr32.exe /s dssenh.dll 
regsvr32.exe /s rsaenh.dll 
regsvr32.exe /s gpkcsp.dll 
regsvr32.exe /s sccbase.dll 
regsvr32.exe /s slbcsp.dll 
regsvr32.exe /s cryptdlg.dll 
regsvr32.exe /s oleaut32.dll 
regsvr32.exe /s ole32.dll 
regsvr32.exe /s shell32.dll 
regsvr32.exe /s initpki.dll 
regsvr32.exe /s wuapi.dll 
regsvr32.exe /s wuaueng.dll 
regsvr32.exe /s wuaueng1.dll 
regsvr32.exe /s wucltui.dll 
regsvr32.exe /s wups.dll 
regsvr32.exe /s wups2.dll 
regsvr32.exe /s wuweb.dll 
regsvr32.exe /s qmgr.dll 
regsvr32.exe /s qmgrprxy.dll 
regsvr32.exe /s wucltux.dll 
regsvr32.exe /s muweb.dll 
regsvr32.exe /s wuwebv.dll 
 
Write-Output "7. Resetting the WinSock..." 
netsh winsock reset | Out-Null
netsh winhttp reset proxy  | Out-Null
 
Write-Output "8. Delete all BITS jobs..." 
Get-BitsTransfer | Remove-BitsTransfer 
 
 
Write-Output "9. Starting Windows Update Services..." 
Start-Service -Name BITS | Out-Null
Start-Service -Name wuauserv | Out-Null
Start-Service -Name appidsvc | Out-Null
Start-Service -Name cryptsvc | Out-Null
 
Write-Output "10. Forcing discovery..." 
wuauclt /resetauthorization /detectnow 

Write-Output "11. Clearing SCCM Cache..."
## Initialize the CCM resource manager com object
[__comobject]$CCMComObject = New-Object -ComObject 'UIResource.UIResourceMgr'
## Get the CacheElementIDs to delete
$CacheInfo = $CCMComObject.GetCacheInfo().GetCacheElements()
## Remove cache items
ForEach ($CacheItem in $CacheInfo) {
    $null = $CCMComObject.GetCacheInfo().DeleteCacheElement([string]$($CacheItem.CacheElementID))
}