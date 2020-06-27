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

    #File Information
    $setupFileID = "07462F03-A4C6-455C-B383-947DDE85DF36" 
    $siteID = "108D28E9-DAC1-4EEA-9566-6591394E6D40"
    $WebID = "4E068646-2C87-4868-924E-850C31F607DF"

    #Get file
    $site = Get-SPSite -limit all | ?{$_.id -eq $siteID}
    $web = Get-SPWeb -Identity $WebID -Site $siteID
    $file = $web.GetFile([GUID]$setupFileID)

    #Report on location
    $filelocation = "{0}{1}" -f ($site.WebApplication.Url).TrimEnd("/"), $file.ServerRelativeUrl
    Write-Host "Found file location:" $filelocation

    #Delete the file, the Delete() method bypasses the recycle bin
    $file.Delete()

    $web.dispose()
    $site.dispose()
}