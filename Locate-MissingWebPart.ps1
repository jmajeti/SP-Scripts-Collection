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
    Add-PSSnapin Microsoft.SharePoint.PowerShell

    $siteID = "108D28E9-DAC1-4EEA-9566-6591394E6D40"
    $webID = "4E068646-2C87-4868-924E-850C31F607DF"
    $dirName = "sites/lab5-4/SitePages"
    $leafName = "Home.aspx"
    $webPartID = "g_bb56b03e_b830_4e37_ba16_62250601ac26"

    ### Locate web parts maintenance page to remove web part ###
    #Get Web
    $web = Get-SPweb -Identity $webID -Site $siteID

    #Build page url
    $pageURL = "{0}{1}/{2}" -f ($site.WebApplication).url, $dirName, $leafName
    $webPartsMaintenancePage = "{0}?contents=1" -f $pageURL
    Write-Host "You can remove the web part by navigating to:`n"$webPartsMaintenancePage
    $web.dispose()
}