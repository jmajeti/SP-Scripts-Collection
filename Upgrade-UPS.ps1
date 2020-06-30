function CreateUPServiceApplication{

    $upService = Get-SPServiceInstance | where {$_.TypeName -eq "User Profile Service"}
    if ($upService.Status -ne "Online") {
      Write-Host "Starting the User Profile Service..."
      $upService | Start-SPServiceInstance | Out-Null
    }

    $upServiceApplicationName = "User Profile Service Application"
    $upServiceApplication = Get-SPServiceApplication | where {$_.Name -eq $upServiceApplicationName}
    if($upServiceApplication -eq $null) {
        Write-Host "Creating the User Profile Service Application..."
        $upServiceApplication = New-SPProfileServiceApplication `
                                   -Name $upServiceApplicationName `
                                   -ProfileDBName "SharePoint_Service_User_Profile_Data" `
                                   -ProfileDBServer $sqlserver `
                                   -SocialDBName "SharePoint_Service_User_Profile_Social" `
                                   -SocialDBServer $sqlserver `
                                   -ProfileSyncDBName "SharePoint_Service_User_Profile_Sync" `
                                   -ProfileSyncDBServer $sqlserver `
                                   -ApplicationPool $serviceAppPoolName `
                                   -MySiteHostLocation $mySiteHostUrl `
                                   -MySiteManagedPath $mySiteManagedPath `
                                   -SiteNamingConflictResolution None `
                                   -DeferUpgradeActions:$false
                                   
                                   

    }

    if($upServiceApplication -ne $null) {
        $upServiceApplicationProxyName = "User Profile Service Application Proxy"
        $upServiceApplicationProxy = Get-SPServiceApplicationProxy | where { $_.Name -eq $upServiceApplicationProxyName}
        if ($upServiceApplicationProxy -eq $null) {
            Write-Host "Creating the User Profile Service Application Proxy..."
            $upServiceApplicationProxy = New-SPProfileServiceApplicationProxy `
                                                    -ServiceApplication $upServiceApplication `
                                                    -Name $upServiceApplicationProxyName `
                                                    -DefaultProxyGroup

        }
    } else {
        Write-Host "User Profile Proxy Creaion Skipped" -ForegroundColor Yellow
    }
}



function Grant-ServiceApplicationPermission($app, $user, $permission, $admin){
    
    $sec = $app | Get-SPServiceApplicationSecurity -Admin:$admin
    $claim = New-SPClaimsPrincipal -Identity $user -IdentityType WindowsSamAccountName
    $sec | Grant-SPObjectSecurity -Principal $claim -Rights $permission
    $app | Set-SPServiceApplicationSecurity -ObjectSecurity $sec -Admin:$admin

}

function CreateDefaultServiceApplicationPool{

  $serviceAppPool = Get-SPServiceApplicationPool -Identity $serviceAppPoolName -ErrorAction SilentlyContinue
  if($serviceAppPool -eq $null){
    Write-Host "Creating default application pool for service applications..."
    $serviceAppPoolNameIdentity = Get-SPManagedAccount -Identity "dht\srv_sp_services"
    New-SPServiceApplicationPool -Name $serviceAppPoolName -Account $serviceAppPoolNameIdentity
    Write-Host
  }
}



# load in SharePoint snap-in
Add-PSSnapin Microsoft.SharePoint.PowerShell

#+++++++++++++++++++++++++++++++++++++++++++++++++
# create variables used in this script
$server = "SHEPHERD"
$sqlserver = "SPSQL"

$serviceAppPoolName = "SharePoint Services App Pool"

#+++++++++++++++++++++++++++++++++++++++++++++++++

CreateDefaultServiceApplicationPool

$mySiteHostUrl = "http://my.doghousetoys.com"
$mySiteManagedPath = "personal"

$mp = Get-SPManagedPath -WebApplication $mySiteHostUrl -Identity $mySiteManagedPath -ErrorAction SilentlyContinue
if ($mp -eq $null) {
    Write-Host "Creating managed path $mySiteManagedPath"
    Get-SPWebApplication -Identity $mySiteHostUrl | New-SPManagedPath -RelativeURL $mySiteManagedPath  | Out-Null
}

CreateUPServiceApplication
