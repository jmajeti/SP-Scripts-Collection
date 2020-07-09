$farm = Get-SPFarm  
$disabledTimers = $farm.TimerService.Instances | where {$_.Status -ne "Online"}  
if ($disabledTimers -ne $null)  
{  
foreach ($timer in $disabledTimers)  
{  
Write-Host "Timer service instance on server " $timer.Server.Name " is not Online. Current status:" $timer.Status 
Write-Host "Attempting to set the status of the service instance to online"  
$timer.Status = [Microsoft.SharePoint.Administration.SPObjectStatus]::Online  
$timer.Update()  
}  
}  
else  
{  
Write-Host "All Timer Service Instances in the farm are online! No problems found" 
}



#Make sure the Timer service instances are online
#and check the values for AllowServiceJobs and AllowContentDatabaseJobs
Add-PSSnapin microsoft.sharepoint.powershell -ErrorAction SilentlyContinue
$farm = Get-SPFarm
$FarmTimers = $farm.TimerService.Instances
foreach ($FT in $FarmTimers){write-host "Server: " $FT.Server.Name.ToString(); write-host "Status: " $FT.status; write-host "Allow Service Jobs: " $FT.AllowServiceJobs; write-host "Allow Content DB Jobs: " $FT.AllowContentDatabaseJobs;"`n"}
$disabledTimers = $farm.TimerService.Instances | where {$_.Status -ne "Online"}
if ($disabledTimers -ne $null)
{foreach ($timer in $disabledTimers)
{Write-Host -ForegroundColor Red "Timer service instance on server " $timer.Server.Name " is NOT Online. Current status:" $timer.Status
Write-Host -ForegroundColor Green "Attempting to set the status of the service instance to online..."
$timer.Status = [Microsoft.SharePoint.Administration.SPObjectStatus]::Online
$timer.Update()
write-host -ForegroundColor Red "You MUST now go restart the SharePoint timer service on server " $timer.Server.Name}}
else{Write-Host -ForegroundColor Green  "All Timer Service Instances in the farm are online. No problems found!"}  