#SharePoint PowerSehll
Add-PSSnapin Microsoft.SharePoint.PowerShell
Clear-Host
#Test-SPContentDatabase

$databases = @{"SharePoint_Content_Apps" = "http://shepherd"; `
                "SharePoint_Content_Suspect_Sites" = "http://shepherd"
                "SharePoint_Content_Deleted_Sites" = "http://intranet.doghousetoys.com"; `
				"SharePoint_Content_Extranet" = "https://extranet.doghousetoys.com/"; `
				"SharePoint_Content_Intranet" = "http://intranet.doghousetoys.com"; `
				"SharePoint_Content_MySite_Host" = "http://my.doghousetoys.com"; `
				"SharePoint_Content_MySite_Personal_01" = "http://my.doghousetoys.com"; `
				"SharePoint_Content_MySite_Personal_02" = "http://my.doghousetoys.com"; `
				"SharePoint_Content_Operations_Web" = "http://operations.doghousetoys.com/"; `
				"SharePoint_Content_Orphan_Sites" = "http://shepherd"; `
				"SharePoint_Content_Partners" = "http://partners.doghousetoys.com"; `
				"SharePoint_Content_Sales" = "http://sales.doghousetoys.com"; `
				"SharePoint_Content_Search_Center" = "http://search.doghousetoys.com"; `
				"SharePoint_Content_WWW" = "http://www.doghousetoys.com"; }

$databases.GetEnumerator() | foreach {
    Write-Host "Testing $($_.Key) against $($_.Value)"
    Test-SPContentDatabase -Name $_.Key -WebApplication $_.Value
} | Out-File C:\Users\administrator.DHT\Desktop\Test-SPContentDatabase.txt

#My Site Host
#Dismount-SPContentDatabase -Identity ZZ_SharePoint_Content_MySite 

