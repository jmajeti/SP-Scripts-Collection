#Restore the Search Service Application
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
 
function RestoreSearchServiceApplication{
    Clear-Host
    $service = Get-SPServiceInstance | where {$_.TypeName -eq "SharePoint Server Search"}
    if ($service.Status -ne "Online") {
        Write-Host "Starting SharePoint Server Search Service..."
        $service | Start-SPServiceInstance | Out-Null
        Write-Host "Waiting for Search Service to Start..." -NoNewline -ForegroundColor Yellow
        while ($service.Status -ne "Online"){
            Start-Sleep -Seconds 3
            $service = Get-SPServiceInstance | where {$_.TypeName -eq "SharePoint Server Search"}
            Write-Host "." -NoNewline -ForegroundColor Yellow
        }
        Write-Host
        Write-Host "Search Service to Started" -ForegroundColor Green
    }

    $searchInstance = Get-SPEnterpriseSearchServiceInstance -Local
    $searchDBStub = "SharePoint_Service_Search"

    $serviceApplicationName = "Enterprise Search Service Application"
    $serviceApplication = Get-SPServiceApplication | where {$_.Name -eq $serviceApplicationName}

    if($serviceApplication -eq $null) {
        Write-Host "Creating the $($serviceApplicationName)..."
        $serviceApplication = Restore-SPEnterpriseSearchServiceApplication `
                                  -Name $searchDBStub `
                                  -ApplicationPool $serviceAppPoolName `
                                  -DatabaseServer $sqlserver `
                                  -DatabaseName  "$($searchDBStub)_Admin"`
                                  -AdminSearchServiceInstance $searchInstance

        if ($serviceApplication -eq $null){
            Write-Host "Error Provisioning Search" -ForegroundColor Red
            Exit
        }
        #Rename the Service Application
        Write-Host "Renaming Search Service to $($serviceApplicationName)" -ForegroundColor Green
        $serviceApplication = Get-SPServiceApplication | where {$_.Name -eq $searchDBStub}
        $serviceApplication.Name = $serviceApplicationName
        $serviceApplication.Update() 

        $serviceApplicationProxyName = "$($serviceApplicationName) Proxy"
        Write-Host "Creating the $($serviceApplicationProxyName)..."
        $serviceApplicationProxy = New-SPEnterpriseSearchServiceApplicationProxy `
                                       -Name $serviceApplicationProxyName `
                                       -SearchApplication $serviceApplication

    }
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

function SetSearchServiceApplicationProperties{
	Write-Host "Setting search service properties..."
    while ($true) {
    	Get-SPEnterpriseSearchService | Set-SPEnterpriseSearchService `
    		-ContactEmail "sharepoint@doghousetoys.com" `
    		-ErrorAction SilentlyContinue -ErrorVariable err
    	if ($err) {
            if ($err[0].Exception.Message -like "*update conflict*") { Write-Warning "An update conflict occured"; Start-Sleep 2; continue; }
    		throw $err
    	}
        break
	}
}

$server = "SHEPHERD"
$sqlserver = "SPSQL"

$serviceAppPoolName = "SharePoint Services App Pool"

#+++++++++++++++++++++++++++++++++++++++++++++++++

CreateDefaultServiceApplicationPool

RestoreSearchServiceApplication

SetSearchServiceApplicationProperties

#Do you have a custom thesarus?
$thesaurusFile = "\\dobie\Library\2013-2016 Upgrade\Search\thesarus.csv"
$serviceApplication = Get-SPEnterpriseSearchServiceApplication
Import-SPEnterpriseSearchThesaurus -SearchApplication $serviceApplication -FileName $thesaurusFile