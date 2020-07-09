####################################################################################################
#
#  Author.......: David Shvartsman
#  Date.........: 05/11/2016
#  Description..: Output a list of all 2013 Workflows in the SharePoint 2013 Farm to a CSV file
#
####################################################################################################
if ((Get-PSSnapin 'Microsoft.SharePoint.PowerShell' -ErrorAction SilentlyContinue) -eq $null) {
  Add-PSSnapin 'Microsoft.SharePoint.PowerShell'
}
CLS
$spAssignment = Start-SPAssignment
$outputFile = 'D:\Temp\2013Workflows.csv'
$output = '';
$wfResults = @();
$i = 0;
Write-Host 'Searching 2013 Workflows ....' -NoNewline;
 
# Get All Web Applications
$WebApps = Get-SPWebApplication
foreach($webApp in $WebApps) {
  # Get All Site Collection
  foreach ($spSite in $webApp.Sites)    {
    # get the collection of webs
    foreach($spWeb in $spSite.AllWebs) {
      $wfm = New-object Microsoft.SharePoint.WorkflowServices.WorkflowServicesManager($spWeb)
      $wfsService = $wfm.GetWorkflowSubscriptionService()
      foreach ($spList in $spWeb.Lists) {
        $subscriptions = $wfsService.EnumerateSubscriptionsByList($spList.ID)
        foreach ($subscription in $subscriptions) {
          #$subscriptions.name
          #$subscriptions.PropertyDefinitions#._UIVersionString #_IsCurrentVersion
          $i++
          #excluding multiple version of the same workflow
          if (($spWeb.Url + $spList.Title + $subscriptions.Name) -ne $output) {
            $output = $spWeb.Url + $spList.Title + $subscription.Name    
            $wfID = $subscription.PropertyDefinitions["SharePointWorkflowContext.ActivationProperties.WebId"]        
            $wfResult = New-Object PSObject;
            $wfResult | Add-Member -type NoteProperty -name 'URL' -value ($spWeb.URL);
            $wfResult | Add-Member -type NoteProperty -name 'ListName' -value ($spList.Title);
            $wfResult | Add-Member -type NoteProperty -name 'wfName' -value ($subscription.Name);
            $wfResult | Add-Member -type NoteProperty -name 'wfID' -value ($wfID);
            $wfResults += $wfResult;
          }
          if ($i -eq 10) {Write-Host '.' -NoNewline; $i = 0;}
        }
      }
    }
  }
}
$wfResults | Export-CSV $outputFile -Force -NoTypeInformation
Write-Host
Write-Host 'Script Completed'
Stop-SPAssignment $spAssignment