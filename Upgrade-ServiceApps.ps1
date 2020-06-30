function CreateManagedMetadataService{

    $service = Get-SPServiceInstance | where {$_.TypeName -eq "Managed Metadata Web Service"}
    if ($service.Status -ne "Online") {
        Write-Host "Starting Managed Metadata Service..."
        $service | Start-SPServiceInstance | Out-Null
    }

    $serviceApplicationName = "Managed Metadata Service Application"
    $serviceApplication = Get-SPServiceApplication | where {$_.Name -eq $serviceApplicationName}

    if($serviceApplication -eq $null) {
        Write-Host "Creating the Managed Metadata Service Application..."
        $serviceApplication = New-SPMetadataServiceApplication `
                                  -Name $serviceApplicationName `
                                  -ApplicationPool $serviceAppPoolName `
                                  -DatabaseName "SharePoint_Service_Managed_Metadata"
    
        $serviceApplicationProxyName = "Managed Metadata Service Application Proxy"
        Write-Host "Creating the Managed Metadata Service Application Proxy..."
        $serviceApplicationProxy = New-SPMetadataServiceApplicationProxy `
                                       -Name $serviceApplicationProxyName `
                                       -ServiceApplication $serviceApplication `
                                       -DefaultProxyGroup

        Grant-ServiceApplicationPermission $serviceApplication "dht\Administrator" "Full Control" $true
    }

}

function CreateSecureStoreServiceApplication{

    $secureStoreService = Get-SPServiceInstance | where {$_.TypeName -eq "Secure Store Service"}
    if ($secureStoreService.Status -ne "Online") {
      Write-Host "Starting the Secure Store Service..."
      $secureStoreService | Start-SPServiceInstance | Out-Null
    }

    $secureStoreServiceApplicationName = "Secure Store Services Service Application"
    $secureStoreServiceApplication = Get-SPServiceApplication | where {$_.Name -eq $secureStoreServiceApplicationName}
    if($secureStoreServiceApplication -eq $null) {
        Write-Host "Creating the Secure Store Service Application..."
        $secureStoreServiceApplication = New-SPSecureStoreServiceApplication `
                                            -Name $secureStoreServiceApplicationName `
                                            -DatabaseName "SharePoint_Service_Secure_Store" `
                                            -DatabaseServer $sqlserver `
                                            -ApplicationPool $serviceAppPoolName `
                                            -AuditingEnabled:$false
    }

    $secureStoreServiceApplicationProxyName = "Secure Store Services Service Application Proxy"
    $secureStoreServiceApplicationProxy = Get-SPServiceApplicationProxy | where { $_.Name -eq $secureStoreServiceApplicationProxyName}
    if ($secureStoreServiceApplicationProxy -eq $null) {
        Write-Host "Creating the Secure Store Service Application Proxy..."
        $secureStoreServiceApplicationProxy = New-SPSecureStoreServiceApplicationProxy `
                                                -ServiceApplication $secureStoreServiceApplication `
                                                -Name $secureStoreServiceApplicationProxyName `
                                                -DefaultProxyGroup
    }
  
    # update and synchronize passphrase
    $secureStoreServiceApplicationPassphrase = "pass@word1"

    # take a pause to ensure proxy has been created and initialized
    Start-Sleep -Seconds 5

    Write-Host "Updating passphrase (e.g. master key) for Secure Store Service Application..."
    Update-SPSecureStoreMasterKey -ServiceApplicationProxy $secureStoreServiceApplicationProxy `
                                  -Passphrase $secureStoreServiceApplicationPassphrase
    
    Write-Host "Synchronizing passphrase for Secure Store Service Application..."
    while ($true) {
        # keep trying until Update-SPSecureStoreApplicationServerKey completes successfully
        try {
            Start-Sleep -Seconds 5
            Update-SPSecureStoreApplicationServerKey `
                -ServiceApplicationProxy $secureStoreServiceApplicationProxy `
                -Passphrase $secureStoreServiceApplicationPassphrase
            break
        }
        catch { }
    }
}


function CreateBCSApplication{

    $service = Get-SPServiceInstance | where {$_.TypeName -eq "Business Data Connectivity Service"}
    if ($service.Status -ne "Online") {
        Write-Host "Starting Business Data Connectivity Service..."
        $service | Start-SPServiceInstance | Out-Null
    }

    $serviceApplicationName = "Business Connectivity Service Application"
    $serviceApplication = Get-SPServiceApplication | where {$_.Name -eq $serviceApplicationName}

    if($serviceApplication -eq $null) {
        Write-Host "Creating the Business Connectivity Service Application..."
        $serviceApplicationDB = "SharePoint_Service_BCS"
        $serviceApplication = New-SPBusinessDataCatalogServiceApplication `
                                  -Name $serviceApplicationName `
                                  -DatabaseServer $sqlserver `
                                  -DatabaseName $serviceApplicationDB `
                                  -ApplicationPool $serviceAppPoolName
 
    }
}

function CreateManagedMetadataService{

    $service = Get-SPServiceInstance | where {$_.TypeName -eq "Managed Metadata Web Service"}
    if ($service.Status -ne "Online") {
        Write-Host "Starting Managed Metadata Service..."
        $service | Start-SPServiceInstance | Out-Null
    }

    $serviceApplicationName = "Managed Metadata Service Application"
    $serviceApplication = Get-SPServiceApplication | where {$_.Name -eq $serviceApplicationName}

    if($serviceApplication -eq $null) {
        Write-Host "Creating the Managed Metadata Service Application..."
        $serviceApplication = New-SPMetadataServiceApplication `
                                  -Name $serviceApplicationName `
                                  -ApplicationPool $serviceAppPoolName `
                                  -DatabaseName "SharePoint_Service_Managed_Metadata"
    
        $serviceApplicationProxyName = "Managed Metadata Service Application Proxy"
        Write-Host "Creating the Managed Metadata Service Application Proxy..."
        $serviceApplicationProxy = New-SPMetadataServiceApplicationProxy `
                                       -Name $serviceApplicationProxyName `
                                       -ServiceApplication $serviceApplication `
                                       -DefaultProxyGroup

        Grant-ServiceApplicationPermission $serviceApplication "dht\Administrator" "Full Control" $true
    }

}

function CreateAppManagementServiceApplication{

    $service = Get-SPServiceInstance | where {$_.TypeName -eq "App Management Service"}
    if ($service.Status -ne "Online") {
        Write-Host "Starting App Management Service..."
        $service | Start-SPServiceInstance | Out-Null
    }

    $serviceApplicationName = "App Management Service Application"
    $serviceApplication = Get-SPServiceApplication | where {$_.Name -eq $serviceApplicationName}

    if($serviceApplication -eq $null) {
        Write-Host "Creating the App Management Service Application..."
        $serviceApplication = New-SPAppManagementServiceApplication `
                                  -Name $serviceApplicationName `
                                  -ApplicationPool $serviceAppPoolName `
                                  -DatabaseName "SharePoint_Service_App_Management"
    
        $serviceApplicationProxyName = "App Management Service Application Proxy"
        Write-Host "Creating the App Management Service Application Proxy..."
        $serviceApplicationProxy = New-SPAppManagementServiceApplicationProxy `
                                       -Name $serviceApplicationProxyName `
                                       -ServiceApplication $serviceApplication `
                                       -UseDefaultProxyGroup 
    }

}

function CreateSubscriptionSettingsService{

    # assign root domain name to configure URL used to access app webs
    Set-SPAppDomain "app.doghousetoys.com" –confirm:$false 

    $subscriptionSettingsService = Get-SPServiceInstance | where {$_.TypeName -like "Microsoft SharePoint Foundation Subscription Settings Service"}

    if($subscriptionSettingsService.Status -ne "Online") { 
        Write-Host "Starting Subscription Settings Service" 
        Start-SPServiceInstance $subscriptionSettingsService | Out-Null
    } 

    # wait for subscription service to start" 
    while ($service.Status -ne "Online") {
      # delay 5 seconds then check to see if service has started   sleep 5
      $service = Get-SPServiceInstance | where {$_.TypeName -like "Microsoft SharePoint Foundation Subscription Settings Service"}
    } 

    $subscriptionSettingsServiceApplicationName = "Subscription Settings Service Application"
    $subscriptionSettingsServiceApplication = Get-SPServiceApplication | where {$_.Name -eq $subscriptionSettingsServiceApplicationName} 

    # create an instance Subscription Service Application and proxy if they do not exist 
    if($subscriptionSettingsServiceApplication -eq $null) { 
      Write-Host "Creating Subscription Settings Service Application..." 
      
      $subscriptionSettingsServiceDB= "SharePoint_Service_Subscription_Settings"
      $subscriptionSettingsServiceApplication = New-SPSubscriptionSettingsServiceApplication `
                                                    -ApplicationPool $serviceAppPoolName `
                                                    -Name $subscriptionSettingsServiceApplicationName `
                                                    -DatabaseName $subscriptionSettingsServiceDB 

      Write-Host "Creating Subscription Settings Service Application Proxy..." 
      $subscriptionSettingsServicApplicationProxy = New-SPSubscriptionSettingsServiceApplicationProxy `
                                                      -ServiceApplication $subscriptionSettingsServiceApplication

    }

    # assign name to default tenant to configure URL used to access web apps 
    Set-SPAppSiteSubscriptionName -Name "app" -Confirm:$false
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

function Grant-ServiceApplicationPermission($app, $user, $permission, $admin){
    
    $sec = $app | Get-SPServiceApplicationSecurity -Admin:$admin
    $claim = New-SPClaimsPrincipal -Identity $user -IdentityType WindowsSamAccountName
    $sec | Grant-SPObjectSecurity -Principal $claim -Rights $permission
    $app | Set-SPServiceApplicationSecurity -ObjectSecurity $sec -Admin:$admin

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

CreateManagedMetadataService

CreateSecureStoreServiceApplication

CreateBCSApplication

CreateAppManagementServiceApplication

CreateSubscriptionSettingsService

