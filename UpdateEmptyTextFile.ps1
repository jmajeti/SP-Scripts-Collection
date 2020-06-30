Add-PSSnapin microsoft.sharepoint.powershell

$url =Read-Host "Enter url"

$listName= "Enter List name"
$site = New-Object Microsoft.SharePoint.SPSite($url)

$web = $site.OpenWeb()

$list = $web.Lists[$listName]

$listitem = $list.Items

$f=$listitem.File

$f.length

foreach ($item in $f)
{

if ($item.Length.Equals(0)) # if length of file is zero which is in case of an empty .TXT file

{


$item.OpenBinary() # open the file
$encode = New-Object System.Text.ASCIIEncoding

$content=$encode.GetBytes("a") # for writing “a” character into the file

$item.SaveBinary($content) # save the file

$item.Update()

}


Else 
{
Write-Host "not empty"
}
}