## AUTHOR : RIPON KUNDU
## RIPONKUNDU@OUTLOOK.COM
## Version 1.0
#### FIND LARGE LISTS IN SHAREPOINT 2010/2013 FARM ####

if ( (Get-PSSnapin -Name microsoft.sharepoint.powershell -ErrorAction SilentlyContinue) -eq $null )
{
    Add-pssnapin microsoft.sharepoint.powershell
}
$SPWebApps = Get-SPWebApplication
$Exceldata = New-Object PSObject
foreach($SPWebApp in $SPWebApps)
{
  $Threshold = $SPWebApp.MaxItemsPerThrottledOperation
  $Warning = $SPWebApp.MaxItemsPerThrottledOperation * (50 / 100)
  $Critical = $SPWebApp.MaxItemsPerThrottledOperation * (75 / 100)
  foreach($SPsite in $SPWebApp.Sites)
  {
    foreach($SPweb in $SPsite.AllWebs)
    {
      foreach($SPlist in $SPweb.Lists)
      {
        if($SPlist.ItemCount -gt $Threshold)
        {
          $obj = New-Object PSObject
          $obj | Add-Member NoteProperty Title($SPlist.Title)
          $obj | Add-Member NoteProperty WebURL($SPweb.URL)
          $obj | Add-Member NoteProperty Count($SPlist.ItemCount)
          $obj | Add-Member NoteProperty Threshold($Threshold)
          $obj | Add-Member NoteProperty Level("Exceeded Limit")
          $obj
        }
        elseif($SPlist.ItemCount -gt $Critical)
        {
          $obj = New-Object PSObject
          $obj | Add-Member NoteProperty Title($SPlist.Title)
          $obj | Add-Member NoteProperty WebURL($SPweb.URL)
          $obj | Add-Member NoteProperty Count($SPlist.ItemCount)
          $obj | Add-Member NoteProperty Threshold($Threshold)
          $obj | Add-Member NoteProperty Level("Above 75%")
          $obj
        }
        elseif($SPlist.ItemCount -gt $Warning)
        {
          $obj = New-Object PSObject
          $obj | Add-Member NoteProperty Title($SPlist.Title)
          $obj | Add-Member NoteProperty WebURL($SPweb.URL)
          $obj | Add-Member NoteProperty Count($SPlist.ItemCount)
          $obj | Add-Member NoteProperty Threshold($Threshold)
          $obj | Add-Member NoteProperty Level("Above 50%")
          $obj
        }
      }
      $SPweb.Dispose()
    }
    $SPsite.Dispose()
  }
}