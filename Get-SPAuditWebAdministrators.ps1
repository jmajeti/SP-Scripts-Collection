#####################################################################################################################################
# Filename: Get-SPAuditWebAdministrators.ps1
# Version : 1.0
# Description : This script gets list of users in site administrators for each site within sharepoint farm and saves them to csv file
#               for audit purposes. This csv can be opened in Microsoft Excel for better viewing.
# Written by  : Mohit Goyal
#####################################################################################################################################

#Set file location for saving information. We'll create a tab separated file.
$FileLocation = "D:\WebAdminsInfo.csv"

#Load SharePoint snap-in
Add-PSSnapin Microsoft.SharePoint.PowerShell

#Fetches webapplications in the farm
$WebApplications = Get-SPWebApplication
Write-Output "Site URL `t Site ID `t Site Description `t Site Administrators" | Out-file $FileLocation

foreach($WebApplication in $WebApplications){
    #Fetches site collections list within sharepoint webapplication
    Write-Output ""
    Write-Output "Fetching Site Collections list from $WebApplication"
    $Sites = Get-SPSite -WebApplication $WebApplication -Limit All    

    foreach($Site in $Sites){
        #Fetches list of sites within site collection
        Write-Output "Fetching Site list from $Site"    
        $Webs = Get-SPWeb -Site $Site -Limit All

        foreach($Web in $Webs){            
            #Fetches information for each  site
            $WebSiteUrl = $Web.Url
            $WebSiteDescription = $Web.Description
            $WebSiteId = $Web.ID.Guid
            $WebSiteAdministrators = ""
            foreach($Item in $Web.SiteAdministrators.DisplayName){ 
                $WebSiteAdministrators += $Item; 
                $WebSiteAdministrators += ";"
            }
        
            Write-Output "$WebSiteUrl `t $WebSiteId `t $WebSiteDescription `t $WebSiteAdministrators" | Out-File $FileLocation -Append
            $Web.Dispose()
        }
    }
}

#Unload SharePoint snap-in
Remove-PSSnapin Microsoft.SharePoint.PowerShell

Write-Output ""
Write-Output "Script Execution finished"
    
#########################################################################################################################################
## End of Script
#########################################################################################################################################