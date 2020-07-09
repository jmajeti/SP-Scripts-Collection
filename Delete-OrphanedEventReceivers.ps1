$ContentDbS = Get-SPDatabase | where {($_.name -eq 'SP_Content_01')}
$Assembly = 'WebProvisionedEvent, Version=1.0.0.0, Culture=neutral, PublicKeyToken=a505ed585a006cfc'

foreach($ContentDb in $ContentDbS)
{
	$SiteCollections = Get-SPSite -ContentDatabase $ContentDb -Limit:All;
	foreach($SiteCollection in $SiteCollections)
	{
		Write-Host '--------------------------------';
		Write-Host $SiteCollection.URL;
		$SiteWasReadOnly=$False;
		try
		{
			if($SiteCollection.ReadOnly)
			{
				Write-Host $SiteCollection.URL 'is read-only. Changing to read-write.';
				$SiteCollection.ReadOnly=$False;
				$SiteWasReadOnly=$True;
				Write-Host $SiteCollection.URL 'is now read-write.';
			}
			if(($SiteCollection.EventReceivers | ?{$_.Assembly -eq $Assembly}) -ne $null)
			{
				Write-Host 'Site with ' $SiteCollection.URL ' has assembly' -ForegroundColor Yellow;
				$er = $SiteCollection.EventReceivers | ?{$_.Assembly -eq $Assembly}
				$er.Delete()
				Write-Host 'Assembly ' $Assembly ' deleted!' -ForegroundColor Green
			}
			foreach($web in $SiteCollection.AllWebs)
			{
				if(($web.EventReceivers | ?{$_.Assembly -eq $Assembly}) -ne $null)
				{
					Write-Host 'Site with ' $web.URL ' has assembly' -ForegroundColor Yellow;
					$webEventreceiver = $web.EventReceivers | ?{$_.Assembly -eq $Assembly}
					$webEventreceiver.Delete()
					Write-Host 'Assembly ' $Assembly ' deleted!' -ForegroundColor Green
				}
				foreach ($list in $web.Lists)
				{
					if(($list.EventReceivers | ?{$_.Assembly -eq $Assembly}) -ne $null)
					{
						Write-Host 'Site with ' $list.Title ' has assembly' -ForegroundColor Yellow;
						$listEventreceiver = $list.EventReceivers | ?{$_.Assembly -eq $Assembly}
						$listEventreceiver.Delete()
						Write-Host 'Assembly ' $Assembly ' deleted!' -ForegroundColor Green
					}
				}
			}
			
		}
		catch [system.exception]
		{
			Write-Host 'Cannot access' $SPWeb.URL -ForegroundColor Red;
		}
		if($SiteWasReadOnly)
		{
			Write-Host $SiteCollection.URL 'was read-only and now is read-write. Changing back to read-only.';
			$SiteCollection.ReadOnly=$True;
			Write-Host $SiteCollection.URL 'is read-only.';
		}
	}
}