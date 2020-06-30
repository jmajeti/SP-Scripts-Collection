Add-PSSnapin microsoft.sharepoint.powershell

$servers = Get-SPServer | where { $_.Role -eq "Application"}

foreach($server in $servers)
{

$webApps = Get-SPWebApplication

$backupDir = 'D$\WebConfigBackup'

$serverPath = "\\"+ $server.Name + "\" + $backupDir

foreach ($webApp in $webApps)
{

$zone = $webApp.AlternateUrls[0].UrlZone

$iisSettings = $webApp.IISSettings[$zone]

$path = $iisSettings.Path.ToString() + "\web.config"


copy-item $path -destination $serverPath\$webapp

Write-Host "Backup for " $webapp "on" $server -foreground Green

}

}