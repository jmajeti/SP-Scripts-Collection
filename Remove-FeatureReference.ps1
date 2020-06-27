<# 
This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.
THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,
INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of the Sample Code, provided that You agree:
(i) to not use Our name, logo, or trademarks to market Your software product in which the Sample Code is embedded
(ii) to include a valid copyright notice on Your software product in which the Sample Code is embedded; and 
(iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits, including attorneys’ fees, that arise or result from the use or distribution of the Sample Code.
#>

$session = Get-PSSession -Name SP2016Search

Invoke-Command -Session $session -ScriptBlock {
    Add-PSSnapin Microsoft.SharePoint.Powershell

    $featureID = "ed37484a-c496-455b-b083-3fc157b1603c"
    $siteID = "108d28e9-dac1-4eea-9566-6591394e6d40"
    

    #Display site information
    $site = Get-SPSite | ?{$_.id -eq $siteID}    
    Write-Host "Checking Site:" $site.Url

    #Remove the feature from all subsites
    ForEach ($web in $Site.AllWebs)
        {
            If($web.Features[$featureID])
                {
                    Write-Host "`nFound Feature $featureID in web:"$Web.Url"`nRemoving feature"
                    $web.Features.Remove($featureID, $true)
                }
                else
                {
                    Write-Host "`nDid not find feature $featureID in web:" $Web.Url
                }   
        }

    #Remove the feature from the site collection
    If ($Site.Features[$featureID])
        {
            Write-Host "`nFound feature $featureID in site:"$site.Url"`nRemoving Feature"
            $site.Features.Remove($featureID, $true)
        }
        else
        {
            Write-Host "Did not find feature $featureID in site:" $site.Url
        }
}




