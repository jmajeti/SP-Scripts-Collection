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
    
    #Variables
    $hostType = "0"
    $siteID = "108D28E9-DAC1-4EEA-9566-6591394E6D40"
    $webID = "00000000-0000-0000-0000-000000000000"
    $hostID = "108D28E9-DAC1-4EEA-9566-6591394E6D40"
    $assemblyID = "80A9D912-768A-4289-91AB-9BF368922F8F"


    #Perform cleanup

    if ($hostType -eq "0")
        {
            $site = GEt-SPSite -limit all -Identity $siteID
            ($site.EventReceivers | ?{$_.id -eq $assemblyID}).delete()
            $site.dispose()
        }
        elseif ($hostype -eq "1")
        {
            $web = Get-SPWeb -Identity $webID -Site $siteID
            ($web.EventReceivers | ?{$_.id -eq $assemblyID}).delete()
            $web.dispose()
        }
        elseif ($hostype -eq "2")
        {
            $web = Get-SPWeb -Identity $webID -Site $siteID
            $list = $web.lists | ?{$_.id -eq $hostID}
            ($list.EventReceivers | ?{$_.id -eq $assemblyID}).delete()
            $web.dispose()
        }



}



