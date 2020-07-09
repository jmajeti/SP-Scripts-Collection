
[system.reflection.assembly]::LoadWithPartialName("System.Configuration") | Out-Null

# intentionally leave off the trailing ".config" as OpenExeConfiguration will auto-append that
$configFilePath = "$env:ProgramFiles\AppFabric 1.1 for Windows Server\DistributedCacheService.exe"
$appFabricConfig = [System.Configuration.ConfigurationManager]::OpenExeConfiguration($configFilePath)

# if backgroundGC setting does not exist add it, else check if value is "false" and change to "true"
if($appFabricConfig.AppSettings.Settings.AllKeys -notcontains "backgroundGC")
{
    $appFabricConfig.AppSettings.Settings.Add("backgroundGC", "true")
}
elseif ($appFabricConfig.AppSettings.Settings["backgroundGC"].Value -eq "false")
{
    $appFabricConfig.AppSettings.Settings["backgroundGC"].Value = "true"
}

# save changes to config file
$appFabricConfig.Save()

