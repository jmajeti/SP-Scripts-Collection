#-----------------------------------------------------------#
#  Description: Script used to Unlock the locked files      #
#                                                           #
#                                                           #
#  Author: Sutha Thavaratnarajah, 	     				    #
#  Date: July 2017                                          #
#-----------------------------------------------------------#
# Change History                                            #
# Author    Date     Description                            #
#-----------------------------------------------------------#
# <every time this file is changed, add change log here>    #
# Usage: Release-FileLock.ps1 [Web URL] [file URL]"         #
#-----------------------------------------------------------#
param( 
    [string] $webUrl = $(throw "No Web URL! Usage: Release-FileLock.ps1 [Web URL] [file URL]"),
	[string] $fileURL = $(throw "file URL! Usage: Release-FileLock.ps1 [Web URL] [file URL]")
)

Function ReleaseLock
{
[CmdletBinding()]
param(
		[Parameter(Position=0, Mandatory=$true)]
		[ValidateNotNullOrEmpty()]
		[System.String]
		$webUrl,
		[Parameter(Position=2, Mandatory=$true)]
		[ValidateNotNullOrEmpty()]
		[System.String]
		$fileURL
	)
	
	$web = get-spweb $webUrl

	$File = $web.GetFile($fileURL)

	if ($File.LockId -ne $null)
	 {
		$userId = $File.LockedByUser.ID
		Write-host "The file has locked by:" $File.LockedByUser.LoginName ",Lock Expires on: "$file.LockExpires

		 #impersonation to release the lock
		$user = $web.AllUsers.GetById($userId)
		$impersonateSite = New-Object Microsoft.SharePoint.SPSite($web.Url, $user.UserToken);
		$impersonateWeb = $impSite.OpenWeb();
		$impersonateItem = $impWeb.GetFile($fileURL);
		$impersonateItem.File.ReleaseLock($impItem.File.LockId)
		Write-host "lock has been released!"
	 }
	Else {
	Write-host "File is not Loked " $File.Name
	}
	$web.Dispose()
}

ReleaseLock $webUrl $FileURL