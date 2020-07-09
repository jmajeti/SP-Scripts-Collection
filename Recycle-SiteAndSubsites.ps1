function Recycle-SiteAndSubsites
{
  Param
    (
     [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][String]$StartWeb,
     [Boolean]$IncludeStartWeb = $true

    )

  Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

  #Get site and subsites starting at $StartWeb 
  $subsites = ((Get-SPWeb $StartWeb).Site).allwebs | ?{$_.url -like "$StartWeb*"}

  #Reverse the order of the site array
  [array]::Reverse($subsites)

  #Traverse the array and delete each site
  foreach($subsite in $subsites)
  {
    write-host "Recycling Site " $subsite.url
    Remove-SPWeb $subsite.url -Confirm:$false -Recycle:$true
  }
}

#Recycle-SiteAndSubsites -StartWeb "http://mysite.domain.com/alpha"