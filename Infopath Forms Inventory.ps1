Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
 
#Configuration parameters
$WebAppURL="https://intranet.crescent.com"
$ReportOutput="C:\InfoPath-ListForms.csv"
#Array to Hold Result - PSObjects
$ResultColl = @()
 
#Get All Webs of the Web Application
$WebsColl = Get-SPWebApplication $WebAppURL | Get-SPSite -Limit All | Get-SPWeb -Limit All
 
#Iterate through each web
Foreach($Web in $WebsColl)
{
 #Get All Lists with InfoPath List Forms in use
 Foreach ($List in $web.Lists | Where { $_.ContentTypes[0].ResourceFolder.Properties["_ipfs_infopathenabled"]})
    {
            Write-Host "Found an InfoPath Form at: $($Web.URL), $($List.Title)"
            $Result = new-object PSObject
            $Result | add-member -membertype NoteProperty -name "Site URL" -Value $web.Url
            $Result | add-member -membertype NoteProperty -name "List Name" -Value $List.Title
            $Result | add-member -membertype NoteProperty -name "List URL" -Value "$($Web.Url)/$($List.RootFolder.Url)"
            $Result | add-member -membertype NoteProperty -name "Template" -Value $list.ContentTypes[0].ResourceFolder.Properties["_ipfs_solutionName"]
            $ResultColl += $Result
    }
}
#Export Results to a CSV File
$ResultColl | Export-csv $ReportOutput -notypeinformation
Write-Host "InfoPath Lists Forms Report has been Generated!" -f Green


