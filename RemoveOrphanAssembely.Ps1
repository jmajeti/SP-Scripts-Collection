Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
 
Function Run-SQLScript($SQLServer, $SQLDatabase, $SQLQuery)
{
    $ConnectionString = "Server =" + $SQLServer + "; Database =" + $SQLDatabase + "; Integrated Security = True"
    $Connection = new-object system.data.SqlClient.SQLConnection($ConnectionString)
    $Command = new-object system.data.sqlclient.sqlcommand($SQLQuery,$Connection)
    $Connection.Open()
    $Adapter = New-Object System.Data.sqlclient.sqlDataAdapter $Command
    $Dataset = New-Object System.Data.DataSet
    $Adapter.Fill($Dataset)
    $Connection.Close()
    $Dataset.Tables[0]
}
 
#Define configuration parameters
$Server="Abj-SQL-001"
$Database="WSS_Content_KM"
$AssemblyInfo="DealEventHandler, Version=1.2.6.0, Culture=neutral, PublicKeyToken=d256f51c842671b3"          
 
#Query SQL Server content Database to get information about the MissingAssembly
$Query = "SELECT distinct Id, SiteId, WebId, HostId, HostType from EventReceivers where Assembly = '"+$AssemblyInfo+"'"
$QueryResults = Run-SQLScript -SQLServer $Server -SQLDatabase $Database -SQLQuery $Query # | select Id, Name, SiteId, WebId, HostId, HostType
 
#Iterate through results
foreach ($Result in $QueryResults)
{
    if($Result.id -ne $Null)
    {
        #Get the location where the event receiver is referred
        if ($Result.HostType -eq 0) #Site Event Receiver
        {
            $Site = Get-SPSite -Limit all | where {$_.Id -eq $Result.SiteId}
            $EventReceiver = $Site.EventReceivers | where {$_.Id -eq $Result.Id}
            #To Delete the Event Receiver
            #$EventReceiver.Delete()
            write-host "Site Event Receivers Found at $($Site.URL)" -foregroundcolor green
        }
        if ($Result.HostType -eq 1) #Web Event Receiver
        {
            $Site = Get-SPSite -Limit all | where {$_.Id -eq $Result.SiteId}
            $Web = $Site | Get-SPWeb -Limit all | where { $_.Id -eq $Result.WebId }
            $EventReceiver = $Web.EventReceivers | where {$_.Id -eq $Result.Id}
            #To Delete the Event Receiver
            #$EventReceiver.Delete()
            write-host "Web Event Receivers Found at $($web.URL)" -foregroundcolor green
        }
        if ($Result.HostType -eq 2) #List Event Receiver
        {
            $Site = Get-SPSite -Limit all | where {$_.Id -eq $Result.SiteId}
            $Web = $Site | Get-SPWeb -Limit all | where { $_.Id -eq $Result.WebId }
            $List = $web.Lists | where {$_.Id -eq $Result.HostId}
            $EventReceiver = $List.EventReceivers | where {$_.Id -eq $Result.Id}
            #To Delete the Event Receiver
            #$EventReceiver.Delete()
            write-host "List Event Receivers Found at $($web.url)/$($List.RootFolder.Url)" -foregroundcolor green
        }
    }
}


#Read more: https://www.sharepointdiary.com/2016/01/fix-missingassembly-error-test-spcontentdatabase-in-sharepoint-migration.html#ixzz6Qbkp7SUY