<#PSScriptInfo

.VERSION 1.1

.GUID 4EC4B3A4-D13C-467B-A252-FFEFD4069F07

.COMPANYNAME Microsoft

#>

<#
.SYNOPSIS 
    Gathers troubleshooting information for On-Demand Assessments in Log Analytics (OMS)

.DESCRIPTION
    This script collects information from your registry, collected data and assessment logs and bundles them altogether on your desktop into a zip file. Once you run this script please attach the zip file that gets created on the desktop and send an email to serviceshubteam@ppas.uservoice.com. Please review the contents of the folder before uploading them or passing them on as they may contain some of your environment configuration details that we would need to further investigate.
 
#>
Set-StrictMode -Version Latest;

# Global variablkes
$idealTargetCount = 5;
$script:LogText = @();
$ErrorActionPreference = 'SilentlyContinue'
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$script:OutPutDirectories = New-Object System.Collections.ArrayList
$includeTracking;


$rootHealthServiceRegistryKey;
# healthServiceExePathFromRegistry contains the path of the HealthService.exe file with " around it.
$script:healthServiceExePathFromRegistry;
[string] $script:hsPath; 

$OutputFolderName = get-date -format "MM.dd.yyyy hh.mm.ss"
$OutputFolder = New-Item -ItemType directory -Path "$DesktopPath\$OutputFolderName" -Force

$script:issueNumber = 0

Write-Host "This script collects information from your registry, collected data and assessment logs and bundles them altogether on your desktop into a zip file. Once you run this script please attach the zip file that gets created on the desktop and send an email to serviceshubteam@ppas.uservoice.com. Please review the contents of the folder before uploading them or passing them on as they may contain some of your environment configuration details that we would need to further investigate."


function CreateIssue ([string] $problemDescription, [string] $fixAction, [string] $links)
{
    $props = @{
        "#" = ++$script:issueNumber
        "Problem Description" = $problemDescription
        "Fix Action" = $fixAction
        "Links" = $links
    }
    new-object psobject -Property $props;
}

# This function makes sure that whether current PS module dll can be used or not. Not for checking the .Net version on the computer
function Test-PowerShellCLRVersion ()
{

    $local:psVersion = $PSVersionTable;

    if ($PSVersionTable.CLRVersion -lt [System.Version]::Parse("4.0.0.0"))
    {
        CreateIssue "Minimum CLRVersion required in Powershell is 4.0.0.0 as the Powershell modules are coded against that.", "Upgrade the version of the Powershell to the latest version.", "https://www.microsoft.com/en-us/download/details.aspx?id=54616";
    }

     $script:LogText += "PowerShell Version $($PSVersionTable.PSVersion) and CLR Version $($PSVersionTable.CLRVersion) `n"
}

function Test-Administrator  
{  
    $user = [Security.Principal.WindowsIdentity]::GetCurrent();
    if (-not (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator))
    {
        CreateIssue "Script needs to run as administrator." "Right click and open Powershell as Administrator and run this script." "No links";
    }
}

function Test-HealthServiceRegistry
{
    $rootHealthServiceRegistryKey = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\HealthService" -ErrorAction SilentlyContinue;
    if (-not $rootHealthServiceRegistryKey)
    {
        CreateIssue "Unable to find registry key Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\HealthService to locate HealthService install folder.", "Install the MMA Agent after downloading it from the Azure portal-> Log Analytics -> Advanced.", "http://portal.azure.com";
        return;
    }

    $script:healthServiceExePathFromRegistry = $rootHealthServiceRegistryKey.ImagePath.Replace('"', "");
    $script:hsPath = [System.IO.Path]::GetDirectoryName($healthServiceExePathFromRegistry);
}

# Get the version of the HealthService
function Test-HealthServiceExe
{
    $HealthServiceFile = gci $healthServiceExePathFromRegistry -ErrorAction SilentlyContinue;
    if (-not $HealthServiceFile)
    {
       CreateIssue "Unable to find the HealthService.exe file in $healthServiceExePathFromRegistry." "1) Is the MMA Agent/HealthService installed ? 2) Can you find HealthService in the Services.msc or Get-service HealthService?" "No Links";
       return   
    }
    
    Get-HealthServiceVersion

    # We can use WMI to further locate the HealthService exe but have not noticed this issue.
    # https://stackoverflow.com/questions/24449113/how-can-i-extract-path-to-executable-of-all-services-with-powershell
}

function Get-HealthServiceVersion
{
    $HealthServiceFile = gci $healthServiceExePathFromRegistry -ErrorAction SilentlyContinue;
    # Not checking for Null as Test-HealthServiceExe is expected to be called before this.
    $healthServiceFileVersionInfo = (Get-Command $HealthServiceFile).FileVersionInfo
    $script:LogText += "HealthService.exe located at $HealthServiceFile has FileVersion: $($healthServiceFileVersionInfo.FileVersion) and ProductVersion: $($healthServiceFileVersionInfo.ProductVersion) `n";
    
	# Checking Health Service (Microsoft Monitoring Agent) status
    $HealthService = Get-Service "HealthService";
    if ($HealthService -and $HealthService.Status -ne "Running")
    {
       CreateIssue "Health Service (Microsoft Monitoring Agent) is NOT in running state" "Go to Services.msc and Start the Microsoft Monitoring Agent service." "No Links";
    }

    $script:LogText += "Health Service (Microsoft Monitoring Agent) is in $($HealthService.Status) state."
}

# The path where the OMS Powershell module is located should be included in $env:PsModulePath for automatic import/tabbing to work.
function Test-PowershellModulePath
{
    # Let us see if the Powershell Module file is present where we expect it to.
    # In Windows 2008 R2 or some installations we have noticed that this missing. 
    
    $rootDirectoryName = [System.IO.Path]::GetDirectoryName($healthServiceExePathFromRegistry);

    $OMSModuleDllLocation = [System.IO.Path]::Combine($rootDirectoryName, "PowerShell\Microsoft.PowerShell.Oms.Assessments\Microsoft.PowerShell.Oms.Assessments.dll");
    $OMSModuleDirectory =  [System.IO.Path]::Combine($rootDirectoryName, "PowerShell\Microsoft.PowerShell.Oms.Assessments");
    $AADApplicationManagerLocation = [System.IO.Path]::Combine($rootDirectoryName, "PowerShell\Microsoft.Assessments.AADApplicationManager\Microsoft.Assessments.AADApplicationManager.psd1");


    $PowershellModulePath = @();

    # Sometimes the .msi fails to add the OMS Powershell Module path to the env:PSModulePath
    # In my machine, I have this path 'D:\Program Files\Microsoft Monitoring Agent\Agent\PowerShell' in $env:PsModulePath

    foreach ($aModulePath in $env:PSModulePath.Split(";"))
    {
        if ($aModulePath -imatch  "Agent\\PowerShell\\")
	    {
		    $PowershellModulePath += "$aModulePath, ";
	    }
    }

    if ($PowershellModulePath.Count -gt 1)
    { 
        CreateIssue "`$env:PSModulePath contains folders from which Powershell modules are imported. Multiple \Agent\Powershell found in `$env:PSModulePath and this will cause issues.You can check this by executing `$env:PsModulePath in Powershell." "Copy the `$env:PSModulePath content using `$env:PSModulePath | clip and clean up duplicate PowerShell modules and assign the updated paths to `$env:PSModulePath using `$env:PSModulePath = << Updated paths>>" "No Links";
    }

    if(Test-Path -path $OMSModuleDllLocation)
    {
        $script:LogText += "Microsoft.PowerShell.Oms.Assessments.dll located at $OMSModuleDllLocation with version $($($(Get-Command $OMSModuleDllLocation).FileVersionInfo).FileVersion)`n"
    }
    else
    {
        CreateIssue "Powershell Module is expected to be in $OMSModuleDirectory. That path - $OMSModuleDirectory is expected to be in `$env:PsModulePath so that import-module does not need to be done." "Import-module $OMSModuleDllLocation will fix the issue in the session or restart the computer and check again. $OMSModuleDirectory should be added to $env:PsModulePath" "No Links";
    }

    if(Test-Path -path $AADApplicationManagerLocation)
    {
        foreach($line in [System.IO.File]::ReadLines($AADApplicationManagerLocation))
        {
            if($line -match "ModuleVersion")
            {
                $script:LogText += "Microsoft.Assessments.AADApplicationManager located at $AADApplicationManagerLocation with $line `n"
                break
            }
        }
    }
    else
    {
        CreateIssue "AAD Application Manager Module is expected to be in $AADApplicationManagerLocation." "Login to Services Hub and add Assessments in the Log Analytics Workspace. After this step, restart Microsoft Monitoring Agent, wait for 10 minutes and check AAD Application Manager availability." "No Links";
    }

    $script:LogText += "Located Agent --> PowerShell module $PowershellModulePath in `$env:PsModulePath `n";
}

# The path where the Powershell module is located should be included in $env:Path for PowerShell to work
function Test-WindowsEnvironmentVariablesPath
{
    $PowershellModulePath = @();

    foreach ($aModulePath in $env:Path.Split(";"))
    {
        if ($aModulePath -imatch  "\\WindowsPowerShell\\")
	    {
		    $PowershellModulePath += "$aModulePath, ";
	    }
    }

    if ($PowershellModulePath.Count -eq 0)
    { 
        CreateIssue "Windows Powershell Module Path is expected to be in Windows System Environment Variables Path which can be checked using `$env:Path in PowerShell." "Add C:\Windows\System32\WindowsPowerShell\v1.0 to Environment Variables Path variable." "https://www.architectryan.com/2018/03/17/add-to-the-path-on-windows-10/";
    }

    $script:LogText += "Located Windows PowerShell module $PowershellModulePath in `$env:Path `n";
}

function FindCountOfTargets ([string] $fileName)
{
    $commandLine = Get-content $fileName;
    $token = $commandLine.Split(" ", [System.StringSplitOptions]::RemoveEmptyEntries);
    for ($i = 0;$i -lt $token.Count; $i++ )
    {
        switch ($token[$i])
        {
            # locate the string after the CASE statement and split it by ; 
            # That gives the number of servers configured as target.
            "-ServerName"
            {
                if ($i+1 -lt $token.Count-1)
                { 
                    $targetServers = $token[$i+1];
                    $targetServers.Split(";", [System.StringSplitOptions]::RemoveEmptyEntries).Count;
                }
            }
            "-SQLServerName"
            {
                if ($i+1 -lt $token.Count-1)
                { 
                    $targetServers = $token[$i+1];
                    $targetServers.Split(";", [System.StringSplitOptions]::RemoveEmptyEntries).Count;
                }
            }
            "-TargetNames"
            {
                if ($i+1 -lt $token.Count-1)
                { 
                    $targetServers = $token[$i+1];
                    $targetServers.Split(";", [System.StringSplitOptions]::RemoveEmptyEntries).Count;
                }
            }
        }
    }

    return -1; # something went wrong.
}


function Get-TaskInfo ([string] $taskName, [string] $AssessmentName)
{
	$tasks = Get-ScheduledTask -TaskName $taskName;
    if (-not $tasks)
    {
        return;
    }
	
	$script:LogText += "Task Scheduler Information for $AssessmentName :"
    foreach($task in $tasks)
	{
		$script:LogText += "TaskName = $($task.TaskName)"
		$script:LogText += "Description = $($task.Description)"
		$script:LogText += "TaskPath = $($task.TaskPath)"
		$script:LogText += "TaskTriggerInterval = $($task.Triggers.Repetition.Interval | Select-Object -First 1)"

        # Dump the contents of the tasks.
        if ($task.Actions.Count -eq 1)
        {
            $taskCommand = GC $($task.Actions[0].Execute)
            $script:LogText += ".cmdLocation $($task.Actions[0].Execute) contains $taskCommand";
            $targetCount = FindCountOfTargets $taskCommand;
                
            if ($targetCount -gt $idealTargetCount)
            {
                CreateIssue "For $($tasks.TaskName), the target server count is $targetCount. Large upload files are produced because of this which cause upload issue" "Create multiple tasks with 5 or fewer target servers to solve/prevent upload issues" "No Links";
            }
        }

		$taskinfo = Get-ScheduledTaskInfo -InputObject $task
		$script:LogText += "LastRunResult = $($taskinfo.LastTaskResult)"
		$script:LogText += "LastRun = $($taskinfo.LastRunTime)"
		$script:LogText += "NextRun = $($taskinfo.NextRunTime)"
        
        # Check the Task Run as user details and administrative privileges
        Check-TaskRunAsPrivileges $task.TaskPath $task.TaskName

	}

    $script:LogText +="`n"
}
 
# Function to check the Task Run as user details and administrative privileges
function Check-TaskRunAsPrivileges([string] $taskPath, [string] $taskName)
{
    $taskProperties = schtasks /query /tn "$taskPath$taskName" /v /fo CSV | ConvertFrom-Csv | ? {$_.Status -ne "Status"}
    $runAsUser = $taskProperties."Run As User"
    
    if (-not [bool](Get-LocalGroupMember -Group "Administrators" | Where-Object {$_.Name.Split("\")[1] -eq $runAsUser}))
    {
        CreateIssue "$taskName task needs to run as administrator." "Get the Task Run as User details by checking the properties of the task in Task scheduler and add the user to computer administrators group. We didn't check the group membership. If the task run as user is part of some group which is part of computer Administrators group, please ignore this error." "No links";
    }
}


function Get-InformationAboutTasks ()
{
    Get-TaskInfo "ADAssessment*" "AD Assessment";
    Get-TaskInfo "ADSecurityAssessment*" "AD Security Assessment";
    Get-TaskInfo "ExchangeAssessment*" "Exchange Assessment";
    Get-TaskInfo "SCCMAssessment*" "SCCM Assessment";
    Get-TaskInfo "SCOMAssessment*" "SCOM Assessment";
    Get-TaskInfo "SfBAssessment*" "SFB Assessment";
    Get-TaskInfo "SharepointAssessment*" "Sharepoint Assessment";
    Get-TaskInfo "SQLAssessment*" "SQL Assessment";
    Get-TaskInfo "WindowsClientAssessment*" "Windows Client Assessment";
    Get-TaskInfo "WindowsServerAssessment*" "Windows Server Assessment";
    
    Get-TaskInfo "ExchangeOnlineAssessmentTask*" "Exchange Online Assessment";
    Get-TaskInfo "SharePointOnlineAssessmentTask*" "SharePoint Online Assessment";
    Get-TaskInfo "SfBOnlineAssessmentTask*" "Sfb Online Assessment";
}

function Get-OutputDirectory ()
{
    $RegistryPath = "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CURRENTCONTROLSET\services\healthservice\parameters\Management Groups\"
    $ManagementGroupDir = Get-ChildItem -Path $RegistryPath
    $ManagementGroups = $ManagementGroupDir.PSChildName

    foreach($ManagementGroup in $ManagementGroups)
    {
        $script:LogText += "Registry Information for Management Group $ManagementGroup :"

        $solutions = Get-ChildItem "$RegistryPath\$ManagementGroup\Solutions" -Recurse
        foreach($solution in $solutions)
        {
            $RegPath = $solution.Name
            
            $OutputDirectoryObject = Get-ItemProperty -Path $solution.PSPath
            if(-not $OutputDirectoryObject)
            {
                $script:LogText += "Output directory is null or empty for $ManagementGroup --> $solution. This could be the root HivePath for a multi-target assessment. `n";
                continue;        
            }

            # To handle SQL type assessments where there is no output directory and scheduled task, but MMA tries to add registry key always.
            if(-not [bool]($OutputDirectoryObject.PSObject.Properties.name -match "DisplayName") -and -not [bool]($OutputDirectoryObject.PSObject.Properties.name -match "OutputDirectoryPath"))
            {
                continue
            }

            $solutionName = $OutputDirectoryObject.DisplayName
            # Display name in registry and package name in Agent folder are different for Share Point Assessment.
			if($solutionName -eq "SharePointAssessment")
            {
                $solutionName = "SPAssessment"
            }

            $script:LogText += "Registry Information for $solutionName :"
            $OutputDirectory = $OutputDirectoryObject.OutputDirectoryPath
            $script:OutPutDirectories.Add($OutputDirectory) | Out-Null
			
			$execPkg = Get-ChildItem -Recurse -Filter "*$solutionName*.execpkg" -Path $script:hsPath;
            $mgmtPkg =  Get-ChildItem -Recurse -Filter "*$solutionName*.xml" -Path $script:hsPath;
            $xdoc = New-Object System.Xml.XmlDocument

            if($execPkg -and $mgmtPkg) 
            {
                $xdoc.load($mgmtPkg.FullName)
                $mgmtPkgVersion = $xdoc.ManagementPack.Manifest.Identity.Version
                $script:LogText += "Located $($execPkg.Name) in $($execPkg.FullName) with version $mgmtPkgVersion.";
                $Checksum = Get-FileHash -Path $execPkg.FullName
            }
            else
            {
                $script:LogText += "$solutionName execpkg not downloaded"
                CreateIssue "ScheduledTask created for Assessment $solutionName which has not been enabled for this LA workspace and this causes the execPkg file - $solutionName.execpkg to be not download in MMA Agent. ScheduledTask will fail" "Add the Assessment $solutionName to the workspace in Services Hub and wait for sometime or restart the HealthService." "No Links";
            }

            $script:LogText += "OutPutDirectory = $OutputDirectory ManagementGroup = $ManagementGroup Checksum = $($Checksum.Hash) `n"
        }
    }
}

# Get the number of rows in a file
function Get-LineCountInFile ([string] $fileName)
{
    if (-not (Get-ChildItem $fileName))
    {
        return -1;
    }
    
    [int] $count = 0;

    [System.IO.File]::ReadLines($fileName) | ForEach-Object { $count++; }

    $count;
}

# Copy Log files to desktop to zip it up
function Copy-LogFiles ()
{    
    foreach($directory in $script:OutPutDirectories)
    {
        if($directory)
        {
			$contents = dir -File -Path $directory
            foreach($file in $contents)
            {
                $lineCount = Get-LineCountInFile $file.FullName;
                $script:LogText += "$($file.FullName)    $($file.Length/1024) KB      $($file.LastWriteTime) LineCount:$lineCount";
            }
           
            $SironaLog = Get-ChildItem -Path $directory -Recurse -Filter "SironaLog*" | sort LastWriteTime | select -last 1
            $DiscoveryLog = Get-ChildItem -Path $directory -Recurse -Filter "DiscoveryTrace*" | sort LastWriteTime | select -last 1
            $TraceLog = Get-ChildItem -Path $directory -Recurse -Filter "*.trace.*" | sort LastWriteTime | select -last 1
            $PreReqsLog = Get-ChildItem -Path $directory -Recurse -Filter "*.prerequisites.*" | sort LastWriteTime | select -last 1
            $TrackingFolder = Get-ChildItem -Path $directory -Recurse -Filter "Tracking" | sort LastWriteTime | select -last 1
            $LogsFolder = [System.IO.Path]::Combine($directory, "Logs");
         
            $name = $directory
            [IO.Path]::GetinvalidFileNameChars() | ForEach-Object {$name = $name.Replace($_," ")}
            $name = $name.trim()
            if($name.ToCharArray().count -gt 150)
            {
                $name = $directory.Split('\')[$directory.Split('\').Count - 2]
            }
            $AssessmentFolder = New-Item -ItemType directory -Path "$DesktopPath\$OutputFolderName\$name" -Force
            $DestinationLogsFolder = New-Item -ItemType directory -Path "$DesktopPath\$OutputFolderName\$name\Logs" -Force

            Copy-Item -Path $SironaLog.FullName -Destination $AssessmentFolder -Force
            Copy-Item -Path $DiscoveryLog.FullName -Destination $AssessmentFolder -Force
            Copy-Item -Path $PreReqsLog.FullName -Destination $AssessmentFolder -Force
            Copy-Item -Path $TraceLog.FullName -Destination $AssessmentFolder -Force
            Copy-Item -Path $LogsFolder -Destination $DestinationLogsFolder -Recurse;

            if($includeTracking)
            {
                Copy-Item -Path $TrackingFolder.FullName -Destination $AssessmentFolder -Container -Recurse -Force
            }    
        }

        $script:LogText += "`n"
    }

    $EventLogExport = "$OutputFolder\EventLogs.evtx"
    $AppLogExport = "$OutputFolder\ApplicationLogs.evtx"
    $TaskLogExport = "$OutputFolder\TaskScheduler.evtx"
    $PrerequisitesLogExport = "$OutputFolder\PrerequisitesLogs.evtx"

    wevtutil epl 'Operations Manager' $EventLogExport

    wevtutil epl 'Application' $AppLogExport

    wevtutil epl 'Microsoft-Windows-TaskScheduler/Operational' $TaskLogExport

    wevtutil epl 'Microsoft-Assessments-Prerequisites/Operational' $PrerequisitesLogExport

    $script:LogText | Out-File -FilePath "$OutputFolder\TraceLogs.txt" -Force

}

function Check-EventLogError ([string] $eventLogName, [string] $eventID)
{
    get-winevent -FilterHashTable @{LogName=$eventLogName; ID=$eventID; StartTime=(get-date).AddDays(-7)}
}

# Check for various event log problems
function Check-EventLog ()
{
    [string] $opsManager = "Operations Manager";
    $result = Check-EventLogError $opsManager 4502 | Where { $_.Message -match ".*Powershell module even after retries.*" }
    if ($result)
    {
    CreateIssue "Powershell windows with Microsoft.PowerShell.Oms.Assessments.dll module was kept open when HealthService tried to update the Module file." "Close all Powershell Window (including this) and restart the HealthService." "No Links";
    }

    $result = Check-EventLogError $opsManager 4502 | Where { $_.Message -match ".*The remote server returned an error: (403) Forbidden*" }
    if ($result)
    {
    CreateIssue "Azure Subscription may be in Disabled state. " "Visit the Azure Portal and make sure the subscription is in Active state in the Subscriptions page (https://ms.portal.azure.com/#blade/Microsoft_Azure_Billing/SubscriptionsBlade). Work with Azure Subscription owner if you don't have access to this page." "No Links";
    }

    $result = Check-EventLogError $opsManager 4501 | Where { $_.Message -match "*reported an error 87L*" }
    if ($result)
    {
    CreateIssue "Some of the recommendations in the recommendations file are very large. " "Create a support ticket and mention that 87L error reported in event viewer. Support team can help you on ingesting data to Log Analytics." "No Links";
    }
}

# Checks the CloudConnection problems
function CheckCloudConnection ()
{
    $script:LogText += "Cloud connection test results:"
   
    if(-not (Test-Path -path "$script:hsPath\TestCloudConnection.exe"))
    {
    $script:LogText +=  "TestCloudConnection.exe is not present in Microsoft Monitoring Agent folder. This may be due to SCOM server setup."
    return
    }

    $cloudConnectionResults = . "$script:hsPath\TestCloudConnection.exe"  
    $script:LogText +=  $cloudConnectionResults 
    $noofWorkspaces =  ($cloudConnectionResults| Select-String "Starting connectivity test for workspace id" -AllMatches).matches.count
    
    if ($cloudConnectionResults -Match "failed") 
    {
    CreateIssue "Cloud connection failed." "Check internet connection and Firewall rules. Below are the URLs with Port 443, ByPass HTTPS inspection and Outbound direction to be allowed to have successful cloud connection. a. *.azure-automation.net
		b. *.blob.core.windows.net
		c. *.ods.opinsights.azure.com
		d. *.oms.opinsights.azure.com" "No Links";   
    }   

    if ($noofWorkspaces -gt 1) 
    {
    CreateIssue "Found multiple workspaces in Microsoft Monitoring Agent Control Panel --> Azure Log Analytics tab" "Go to Microsoft Monitoring Agent Control Panel --> Azure Log Analytics tab and delete the redundant workspace. Restart Microsoft Monitoring Agent service after cleanup. " "No Links";   
    }
}

# Checks the unload data at user logoff registry setting
function Registry_Logoff ()
{
    $script:LogText += "Registry Logoff setting details:"
    $DisableForceUnload = Get-ItemProperty -Path "Registry::HKLM\Software\Policies\Microsoft\Windows\System" -Name "DisableForceUnload" -ErrorAction SilentlyContinue
    if(-not $DisableForceUnload)
    {
    $script:LogText += "Do not forcefully unload the users registry at user logoff is not configured. `n"
    CreateIssue "User logoff is not configured." "Do not forcefully unload the users registry at user logoff is not configured." "https://docs.microsoft.com/en-us/services-hub/health/assessments-troubleshooting#verify-the-user-account-group-policies";
    }
    else
    {
    $script:LogText += "Do not forcefully unload the users registry at user logoff = $($DisableForceUnload.DisableForceUnload) `n"
    }
}

# Checks the FIPS Policy details
function CheckFIPSPolicy ()
{
    $FIPSPolicy = Get-ItemProperty -Path "Registry::HKLM\SYSTEM\CurrentControlSet\Control\Lsa\FipsAlgorithmPolicy" -ErrorAction SilentlyContinue
    if(-not $FIPSPolicy)
    {
    $script:LogText += "Could not find FIPS Policy."
    CreateIssue "Could not find FIPS Policy." "Try adding FIPS Policy." "https://docs.microsoft.com/en-us/services-hub/health/assessments-troubleshooting#disable-the-fips-policy";
    return
    }
    
    $script:LogText += "FIPSPolicy details: `n $FIPSPolicy `n"
}

# Checks for minimum required .Net Framework version by querying the registry in PowerShell
function CheckDotNetVersion ()
{
    $releaseVersionNumberOfdotNet = 394802; 
    #394802 is the relese version of .Net 4.6.2. Reference: https://docs.microsoft.com/en-us/dotnet/framework/migration-guide/how-to-determine-which-versions-are-installed?view=netframework-4.7.2
    $version = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\' | Get-ItemPropertyValue -Name Release | Foreach-Object { $_ -ge $releaseVersionNumberOfdotNet } 
    if(-not $version)
    {
    CreateIssue "In the future, a minimum .Net version of 4.6.2 is required on this computer. Please upgrade." "Upgrade .Net version on this computer to 4.6.2 or higher" "Suggested Link: https://dotnet.microsoft.com/download/dotnet-framework-runtime"
    }
}

# Checks the specified issue in the specified log file
function CheckLogFile ([string] $logFile, [string] $issuetoCheck, [string] $problemDescription, [string] $fixAction, [string] $links)
{
    $anyIssueFoundRows = 0
    foreach($directory in $script:OutPutDirectories)
    {
        if($directory)
        {
            $issueFoundLines =@()
            $logFiletoCheck = Get-ChildItem -Path $directory -Recurse -Filter "$logFile*" | sort LastWriteTime | select -last 1
            $logFilePath = $logFiletoCheck.FullName
            $lineNumber = 0
            foreach($line in [System.IO.File]::ReadLines($logFilePath))
            {
                $lineNumber++
                if($line -match $issuetoCheck)
                {
                     $issueFoundLines += "$line at line number $lineNumber";
                }
            }

            if(@($issueFoundLines).Count -gt 0)
            {
             $anyIssueFoundRows = 1
             $script:LogText += "Issue - $issuetoCheck record(s) in $logFilePath file:"          
             $script:LogText += $issueFoundLines -join "`n"
             $script:LogText += "`n"
            }
        }
    }
    
    if($anyIssueFoundRows)
    {
        CreateIssue $problemDescription $fixAction $links
    }
}


$style = @'
<style type="text/css">
table {
  border: 1px solid #1C6EA4;
  background-color: #EEEEEE;
  width: 100%;
  text-align: left;
  border-collapse: collapse;
  font-family: Segoe UI
}
table td, table th {
  border: 1px solid #AAAAAA;
  padding: 5px;
}
table tbody td {
  font-size: 13px;
}
table tr:nth-child(even) {
  background: #D0E4F5;
}
table thead {
  background: #1C6EA4;
  background: -moz-linear-gradient(top, #5592bb 0%, #327cad 66%, #1C6EA4 100%);
  background: -webkit-linear-gradient(top, #5592bb 0%, #327cad 66%, #1C6EA4 100%);
  background: linear-gradient(to bottom, #5592bb 0%, #327cad 66%, #1C6EA4 100%);
  border-bottom: 2px solid #444444;
}
table thead th {
  font-size: 15px;
  font-weight: bold;
  color: #FFFFFF;
  border-left: 2px solid #D0E4F5;
}
table thead th:first-child {
  border-left: none;
}

table tfoot {
  font-size: 14px;
  font-weight: bold;
  color: #FFFFFF;
  background: #D0E4F5;
  background: -moz-linear-gradient(top, #dcebf7 0%, #d4e6f6 66%, #D0E4F5 100%);
  background: -webkit-linear-gradient(top, #dcebf7 0%, #d4e6f6 66%, #D0E4F5 100%);
  background: linear-gradient(to bottom, #dcebf7 0%, #d4e6f6 66%, #D0E4F5 100%);
  border-top: 2px solid #444444;
}
table tfoot td {
  font-size: 14px;
}
table tfoot .links {
  text-align: right;
}
table tfoot .links a{
  display: inline-block;
  background: #1C6EA4;
  color: #FFFFFF;
  padding: 2px 8px;
  border-radius: 5px;
}
</style>
'@;


function Main ()
{
    Write-Host "------------------------------------------------------"
    Write-Host "Started collecting data"

    $adminResult = Test-Administrator;
    if ($adminResult)
    {
        Write-Host $adminResult;
        exit -1;
    }   

    $failedResults = @();
    
    $versionResult = Test-PowerShellCLRVersion;
    if ($versionResult)
    {
        $failedResults += $versionResult;
    }

    $healthServiceRegistry = Test-HealthServiceRegistry;
    if ($healthServiceRegistry)
    {
        $failedResults += $healthServiceRegistry;
    }
    
    $healthServiceExePresence = Test-HealthServiceExe;
    if ($healthServiceExePresence)
    {
        $failedResults += healthServiceExePresence;
    }

    $psModulePathTest = Test-PowershellModulePath;
    if ($psModulePathTest)
    {
        $failedResults += $psModulePathTest;
    }

    $psModuleWindowsPathTest = Test-WindowsEnvironmentVariablesPath;
    if ($psModuleWindowsPathTest)
    {
        $failedResults += $psModuleWindowsPathTest;
    }
    
    $taskResults =  Get-InformationAboutTasks;
    if ($taskResults)
    {
        $failedResults += $taskResults;
    }

    $outputDirectoryResults = Get-OutputDirectory;
    if ($outputDirectoryResults)
    {
        $failedResults += $outputDirectoryResults;
    }

    $fipsPolicyResults = CheckFIPSPolicy;
    if ($fipsPolicyResults)
    {
        $failedResults += $fipsPolicyResults;
    }

    $registrysignoffResults = Registry_Logoff;
    if ($registrysignoffResults)
    {
        $failedResults += $registrysignoffResults;
    }

    $cloudConnectionResults = CheckCloudConnection; 
    if ($cloudConnectionResults) 
    {
        $failedResults += $cloudConnectionResults;
    }

    $eventLogResults = Check-EventLog;
    if($eventLogResults)
    {
        $failedResults += $eventLogResults;
    }

    $dotNetVersion = CheckDotNetVersion;
    if($dotNetVersion)
    {
        $failedResults += $dotNetVersion;
    }

    #This is to check UserName is NULL issue for Online Assessments in Discovery trace file
    $userNameNULLTestResults = CheckLogFile "DiscoveryTrace" "UserName=<null>" "Office 365 assessment unable to find credentials in Windows Credential Manager. This usually happens when scheduled task is configured to run as a different user than the one who is running the Add-* cmdlet." "Login as the user who will run the scheduled task and run the add-* cmdlet to install the Office 365 assessments." "No Links"
    if($userNameNULLTestResults)
    {
        $failedResults += $userNameNULLTestResults;
    }

    #This is to check 2907 Copying files error in Assessments_*_Commandlet file
    $copyingFileTestResults = CheckLogFile "*Commandlet*" "[2907]" "Copying files from Microsoft Monitoring Agent folder to working directiry is not successful. One of the following has happened:
											1. A Powershell window is open with Microsoft.Powershell.OMS.Assessments.dll module loaded and the file is being used in that process.Because of this, the Microsoft.Powershell.OMS.Assessments.dll download failed and this dll is trying to copy an older version of the binaries (which do not exist as the newer version was downloaded) needed to run the assessment. Please close all the Powershell windows and restart the Microsoft Health Service
											2. The Monitoring Agent is linked to more than one workspace and these workspaces have different scopes with different required dlls/versions needed to run the assessment. Please remove all but one workspace using the AgentControlPanel.exe.
											3. $env:PSModulePath contains a directory that has an older version of Microsoft.Powershell.OMS.Assessments.dll but the latest version of this dll is in a different directory. Please remove from $env:PSModulePath the directory containing the old version of Microsoft.Powershell.OMS.Assessments.dll" "If everything looks good and ddidn't find the issue, stop the Microsoft Monitoring Agent service, rename Health Service State folder to -old under Agent folder and start the service. It will create Health Service State folder again. Go inside this folder and check all Management Packs are downloaded. After this, run the Task again and monitor the results." "No Links"
    if($copyingFileTestResults)
    {
        $failedResults += $copyingFileTestResults;
    }

    Copy-LogFiles;

    $LogText | Out-File -FilePath "$OutputFolder\TraceLogs.txt" -Force
    $failedResults | ConvertTo-Html -As table -Title "Issues Summary:" -Head $style -Property "#", "Problem Description", "Fix Action", "Links" | Out-File "$OutputFolder\Report.html"

    Write-Host "Completed collecting data"
    Write-Host "------------------------------------------------------" 
    
    Write-Host "Creating zip folder"
    Compress-Archive -Path $OutputFolder -DestinationPath "$DesktopPath\$OutputFolderName.zip" -Force
    Write-Host "Completed creating zip folder. Zip folder is available in Desktop."
   
}

# Call main method to execute all functions
Main







# SIG # Begin signature block
# MIIjhgYJKoZIhvcNAQcCoIIjdzCCI3MCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCwewzc23/ctmX+
# kwMxFbBBqIImYXKMcxtvoMPUpDnhNKCCDYEwggX/MIID56ADAgECAhMzAAABUZ6N
# j0Bxow5BAAAAAAFRMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMTkwNTAyMjEzNzQ2WhcNMjAwNTAyMjEzNzQ2WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQCVWsaGaUcdNB7xVcNmdfZiVBhYFGcn8KMqxgNIvOZWNH9JYQLuhHhmJ5RWISy1
# oey3zTuxqLbkHAdmbeU8NFMo49Pv71MgIS9IG/EtqwOH7upan+lIq6NOcw5fO6Os
# +12R0Q28MzGn+3y7F2mKDnopVu0sEufy453gxz16M8bAw4+QXuv7+fR9WzRJ2CpU
# 62wQKYiFQMfew6Vh5fuPoXloN3k6+Qlz7zgcT4YRmxzx7jMVpP/uvK6sZcBxQ3Wg
# B/WkyXHgxaY19IAzLq2QiPiX2YryiR5EsYBq35BP7U15DlZtpSs2wIYTkkDBxhPJ
# IDJgowZu5GyhHdqrst3OjkSRAgMBAAGjggF+MIIBejAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUV4Iarkq57esagu6FUBb270Zijc8w
# UAYDVR0RBEkwR6RFMEMxKTAnBgNVBAsTIE1pY3Jvc29mdCBPcGVyYXRpb25zIFB1
# ZXJ0byBSaWNvMRYwFAYDVQQFEw0yMzAwMTIrNDU0MTM1MB8GA1UdIwQYMBaAFEhu
# ZOVQBdOCqhc3NyK1bajKdQKVMFQGA1UdHwRNMEswSaBHoEWGQ2h0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY0NvZFNpZ1BDQTIwMTFfMjAxMS0w
# Ny0wOC5jcmwwYQYIKwYBBQUHAQEEVTBTMFEGCCsGAQUFBzAChkVodHRwOi8vd3d3
# Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY0NvZFNpZ1BDQTIwMTFfMjAx
# MS0wNy0wOC5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAgEAWg+A
# rS4Anq7KrogslIQnoMHSXUPr/RqOIhJX+32ObuY3MFvdlRElbSsSJxrRy/OCCZdS
# se+f2AqQ+F/2aYwBDmUQbeMB8n0pYLZnOPifqe78RBH2fVZsvXxyfizbHubWWoUf
# NW/FJlZlLXwJmF3BoL8E2p09K3hagwz/otcKtQ1+Q4+DaOYXWleqJrJUsnHs9UiL
# crVF0leL/Q1V5bshob2OTlZq0qzSdrMDLWdhyrUOxnZ+ojZ7UdTY4VnCuogbZ9Zs
# 9syJbg7ZUS9SVgYkowRsWv5jV4lbqTD+tG4FzhOwcRQwdb6A8zp2Nnd+s7VdCuYF
# sGgI41ucD8oxVfcAMjF9YX5N2s4mltkqnUe3/htVrnxKKDAwSYliaux2L7gKw+bD
# 1kEZ/5ozLRnJ3jjDkomTrPctokY/KaZ1qub0NUnmOKH+3xUK/plWJK8BOQYuU7gK
# YH7Yy9WSKNlP7pKj6i417+3Na/frInjnBkKRCJ/eYTvBH+s5guezpfQWtU4bNo/j
# 8Qw2vpTQ9w7flhH78Rmwd319+YTmhv7TcxDbWlyteaj4RK2wk3pY1oSz2JPE5PNu
# Nmd9Gmf6oePZgy7Ii9JLLq8SnULV7b+IP0UXRY9q+GdRjM2AEX6msZvvPCIoG0aY
# HQu9wZsKEK2jqvWi8/xdeeeSI9FN6K1w4oVQM4Mwggd6MIIFYqADAgECAgphDpDS
# AAAAAAADMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0
# ZSBBdXRob3JpdHkgMjAxMTAeFw0xMTA3MDgyMDU5MDlaFw0yNjA3MDgyMTA5MDla
# MH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMT
# H01pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTEwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQCr8PpyEBwurdhuqoIQTTS68rZYIZ9CGypr6VpQqrgG
# OBoESbp/wwwe3TdrxhLYC/A4wpkGsMg51QEUMULTiQ15ZId+lGAkbK+eSZzpaF7S
# 35tTsgosw6/ZqSuuegmv15ZZymAaBelmdugyUiYSL+erCFDPs0S3XdjELgN1q2jz
# y23zOlyhFvRGuuA4ZKxuZDV4pqBjDy3TQJP4494HDdVceaVJKecNvqATd76UPe/7
# 4ytaEB9NViiienLgEjq3SV7Y7e1DkYPZe7J7hhvZPrGMXeiJT4Qa8qEvWeSQOy2u
# M1jFtz7+MtOzAz2xsq+SOH7SnYAs9U5WkSE1JcM5bmR/U7qcD60ZI4TL9LoDho33
# X/DQUr+MlIe8wCF0JV8YKLbMJyg4JZg5SjbPfLGSrhwjp6lm7GEfauEoSZ1fiOIl
# XdMhSz5SxLVXPyQD8NF6Wy/VI+NwXQ9RRnez+ADhvKwCgl/bwBWzvRvUVUvnOaEP
# 6SNJvBi4RHxF5MHDcnrgcuck379GmcXvwhxX24ON7E1JMKerjt/sW5+v/N2wZuLB
# l4F77dbtS+dJKacTKKanfWeA5opieF+yL4TXV5xcv3coKPHtbcMojyyPQDdPweGF
# RInECUzF1KVDL3SV9274eCBYLBNdYJWaPk8zhNqwiBfenk70lrC8RqBsmNLg1oiM
# CwIDAQABo4IB7TCCAekwEAYJKwYBBAGCNxUBBAMCAQAwHQYDVR0OBBYEFEhuZOVQ
# BdOCqhc3NyK1bajKdQKVMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1Ud
# DwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFHItOgIxkEO5FAVO
# 4eqnxzHRI4k0MFoGA1UdHwRTMFEwT6BNoEuGSWh0dHA6Ly9jcmwubWljcm9zb2Z0
# LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcmwwXgYIKwYBBQUHAQEEUjBQME4GCCsGAQUFBzAChkJodHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcnQwgZ8GA1UdIASBlzCBlDCBkQYJKwYBBAGCNy4DMIGDMD8GCCsGAQUFBwIB
# FjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2RvY3MvcHJpbWFyeWNw
# cy5odG0wQAYIKwYBBQUHAgIwNB4yIB0ATABlAGcAYQBsAF8AcABvAGwAaQBjAHkA
# XwBzAHQAYQB0AGUAbQBlAG4AdAAuIB0wDQYJKoZIhvcNAQELBQADggIBAGfyhqWY
# 4FR5Gi7T2HRnIpsLlhHhY5KZQpZ90nkMkMFlXy4sPvjDctFtg/6+P+gKyju/R6mj
# 82nbY78iNaWXXWWEkH2LRlBV2AySfNIaSxzzPEKLUtCw/WvjPgcuKZvmPRul1LUd
# d5Q54ulkyUQ9eHoj8xN9ppB0g430yyYCRirCihC7pKkFDJvtaPpoLpWgKj8qa1hJ
# Yx8JaW5amJbkg/TAj/NGK978O9C9Ne9uJa7lryft0N3zDq+ZKJeYTQ49C/IIidYf
# wzIY4vDFLc5bnrRJOQrGCsLGra7lstnbFYhRRVg4MnEnGn+x9Cf43iw6IGmYslmJ
# aG5vp7d0w0AFBqYBKig+gj8TTWYLwLNN9eGPfxxvFX1Fp3blQCplo8NdUmKGwx1j
# NpeG39rz+PIWoZon4c2ll9DuXWNB41sHnIc+BncG0QaxdR8UvmFhtfDcxhsEvt9B
# xw4o7t5lL+yX9qFcltgA1qFGvVnzl6UJS0gQmYAf0AApxbGbpT9Fdx41xtKiop96
# eiL6SJUfq/tHI4D1nvi/a7dLl+LrdXga7Oo3mXkYS//WsyNodeav+vyL6wuA6mk7
# r/ww7QRMjt/fdW1jkT3RnVZOT7+AVyKheBEyIXrvQQqxP/uozKRdwaGIm1dxVk5I
# RcBCyZt2WwqASGv9eZ/BvW1taslScxMNelDNMYIVWzCCFVcCAQEwgZUwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMQITMwAAAVGejY9AcaMOQQAAAAABUTAN
# BglghkgBZQMEAgEFAKCBrjAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgor
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgexemkWJs
# OL4knkH+QNtbsLIQvEFiWie1/Qey926fnsYwQgYKKwYBBAGCNwIBDDE0MDKgFIAS
# AE0AaQBjAHIAbwBzAG8AZgB0oRqAGGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbTAN
# BgkqhkiG9w0BAQEFAASCAQCLAvITfd8ao5dzy/k+PYt9H01j/00UK9cyGCZm3wmX
# XKu9NkBLxWm0HbM0AO19IwmF/ODQLBwirwH/UjixSmsYfzB1JQHot317kpDDNEoA
# Y0u6Ap8bJ3dqcoBoMgfn4AYdCkpaOK6tNRN8f0ykeoUfktfqcyXHt0VUAzBidK8k
# GHMU05aID3T23LN8JTbd2OXqQ1vhS+SI/jAYCJcC/SNfEygvsSBlO7xy+Si8kDVP
# r/kgjQQ1mZvlhhaUnMvaypRIfJr59u5eh2QaU8slAREcWkEDrm04E7uBKWSsRKNz
# 6IFrErWCG8CXQggLUTekL6YU8Sctlu/Jfjt+vHG5tYGRoYIS5TCCEuEGCisGAQQB
# gjcDAwExghLRMIISzQYJKoZIhvcNAQcCoIISvjCCEroCAQMxDzANBglghkgBZQME
# AgEFADCCAVEGCyqGSIb3DQEJEAEEoIIBQASCATwwggE4AgEBBgorBgEEAYRZCgMB
# MDEwDQYJYIZIAWUDBAIBBQAEIBGqwyI7bSG2Va4Hl32H2oRjX06KZjmJjwO2exM5
# TMETAgZdl5UJJlIYEzIwMTkxMDA4MjIwNjI1LjUwNFowBIACAfSggdCkgc0wgcox
# CzAJBgNVBAYTAlVTMQswCQYDVQQIEwJXQTEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQg
# SXJlbGFuZCBPcGVyYXRpb25zIExpbWl0ZWQxJjAkBgNVBAsTHVRoYWxlcyBUU1Mg
# RVNOOkEyNDAtNEI4Mi0xMzBFMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFt
# cCBzZXJ2aWNloIIOPDCCBPEwggPZoAMCAQICEzMAAADgship1NHCtPcAAAAAAOAw
# DQYJKoZIhvcNAQELBQAwfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
# b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3Jh
# dGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwHhcN
# MTgwODIzMjAyNzAxWhcNMTkxMTIzMjAyNzAxWjCByjELMAkGA1UEBhMCVVMxCzAJ
# BgNVBAgTAldBMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJhdGlv
# bnMgTGltaXRlZDEmMCQGA1UECxMdVGhhbGVzIFRTUyBFU046QTI0MC00QjgyLTEz
# MEUxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIHNlcnZpY2UwggEiMA0G
# CSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDCl/hnDsembJx7FMS5RuaJsiQJDP0t
# iqyeCettiNlGcNoa428oVtblH6yZatCXZygyhzbnDJofTHIGtdiEzQc5fPhddfTd
# 4hEQgd5ch/BlGITXFEwJ4d/GhHZQ1hbLdiBT/j67Qx15VeuXy5n/jM9PvIbBksSW
# wX8vrkhRT/rqa1xWrmF+SfcYKw+pC+d3tUHrgACo0YaVHuS3jlpXg33A+pul8wib
# ZBcGxMF1CqwlP0AfMW60Dp4qm/JLbWxdx/pOiiOrM/tykFDtEnN07HXRjXDhDhfI
# eBCz4GPiCEFk94AaFxysFeFn9vyz7TyKJxUksXJhtWGq2i4WmPcphyDzAgMBAAGj
# ggEbMIIBFzAdBgNVHQ4EFgQUa0HTCrY5zqzv/p44rWuSbXaAh+gwHwYDVR0jBBgw
# FoAU1WM6XIoxkPNDe3xGG8UzaFqFbVUwVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDov
# L2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljVGltU3RhUENB
# XzIwMTAtMDctMDEuY3JsMFoGCCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0
# cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNUaW1TdGFQQ0FfMjAx
# MC0wNy0wMS5jcnQwDAYDVR0TAQH/BAIwADATBgNVHSUEDDAKBggrBgEFBQcDCDAN
# BgkqhkiG9w0BAQsFAAOCAQEACx/endmS5DW6xgb8fIdEqI963fUB7IoYlYNQU/YZ
# wq155uK1lwhcH5R1CaVMr+lyNIfD8l+lqy8mdou+Zwcrnzo3m2UEGO0uNFd4c8Ie
# w5Z49V+6AojT6z5IGJh3y56uDACQzRZrR+26uCx1nLsYjK/WtxQDq1IHHWeAxfrG
# xsAZO1BdQo25Mx34ZseViVj+usfmy0nUmfvZ0hFcMeNd4i4Kds03kY/CwWVZBw62
# tAjOHK/c81wO7hiutu9JX4MNjaEuFdheiNwmHyAmbpqYmz6K+9IPM75iXELbzjDc
# 6yLJpVOq17gfVDCaweryzgVnC2CIxq7gDGyTM9afwMtESTCCBnEwggRZoAMCAQIC
# CmEJgSoAAAAAAAIwDQYJKoZIhvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRp
# ZmljYXRlIEF1dGhvcml0eSAyMDEwMB4XDTEwMDcwMTIxMzY1NVoXDTI1MDcwMTIx
# NDY1NVowfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQG
# A1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQCpHQ28dxGKOiDs/BOX9fp/aZRrdFQQ1aUKAIKF
# ++18aEssX8XD5WHCdrc+Zitb8BVTJwQxH0EbGpUdzgkTjnxhMFmxMEQP8WCIhFRD
# DNdNuDgIs0Ldk6zWczBXJoKjRQ3Q6vVHgc2/JGAyWGBG8lhHhjKEHnRhZ5FfgVSx
# z5NMksHEpl3RYRNuKMYa+YaAu99h/EbBJx0kZxJyGiGKr0tkiVBisV39dx898Fd1
# rL2KQk1AUdEPnAY+Z3/1ZsADlkR+79BL/W7lmsqxqPJ6Kgox8NpOBpG2iAg16Hgc
# sOmZzTznL0S6p/TcZL2kAcEgCZN4zfy8wMlEXV4WnAEFTyJNAgMBAAGjggHmMIIB
# 4jAQBgkrBgEEAYI3FQEEAwIBADAdBgNVHQ4EFgQU1WM6XIoxkPNDe3xGG8UzaFqF
# bVUwGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwCwYDVR0PBAQDAgGGMA8GA1Ud
# EwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU1fZWy4/oolxiaNE9lJBb186aGMQwVgYD
# VR0fBE8wTTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwv
# cHJvZHVjdHMvTWljUm9vQ2VyQXV0XzIwMTAtMDYtMjMuY3JsMFoGCCsGAQUFBwEB
# BE4wTDBKBggrBgEFBQcwAoY+aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9j
# ZXJ0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcnQwgaAGA1UdIAEB/wSBlTCB
# kjCBjwYJKwYBBAGCNy4DMIGBMD0GCCsGAQUFBwIBFjFodHRwOi8vd3d3Lm1pY3Jv
# c29mdC5jb20vUEtJL2RvY3MvQ1BTL2RlZmF1bHQuaHRtMEAGCCsGAQUFBwICMDQe
# MiAdAEwAZQBnAGEAbABfAFAAbwBsAGkAYwB5AF8AUwB0AGEAdABlAG0AZQBuAHQA
# LiAdMA0GCSqGSIb3DQEBCwUAA4ICAQAH5ohRDeLG4Jg/gXEDPZ2joSFvs+umzPUx
# vs8F4qn++ldtGTCzwsVmyWrf9efweL3HqJ4l4/m87WtUVwgrUYJEEvu5U4zM9GAS
# inbMQEBBm9xcF/9c+V4XNZgkVkt070IQyK+/f8Z/8jd9Wj8c8pl5SpFSAK84Dxf1
# L3mBZdmptWvkx872ynoAb0swRCQiPM/tA6WWj1kpvLb9BOFwnzJKJ/1Vry/+tuWO
# M7tiX5rbV0Dp8c6ZZpCM/2pif93FSguRJuI57BlKcWOdeyFtw5yjojz6f32WapB4
# pm3S4Zz5Hfw42JT0xqUKloakvZ4argRCg7i1gJsiOCC1JeVk7Pf0v35jWSUPei45
# V3aicaoGig+JFrphpxHLmtgOR5qAxdDNp9DvfYPw4TtxCd9ddJgiCGHasFAeb73x
# 4QDf5zEHpJM692VHeOj4qEir995yfmFrb3epgcunCaw5u+zGy9iCtHLNHfS4hQEe
# gPsbiSpUObJb2sgNVZl6h3M7COaYLeqN4DMuEin1wC9UJyH3yKxO2ii4sanblrKn
# QqLJzxlBTeCG+SqaoxFmMNO7dDJL32N79ZmKLxvHIa9Zta7cRDyXUHHXodLFVeNp
# 3lfB0d4wwP3M5k37Db9dT+mdHhk4L7zPWAUu7w2gUDXa7wknHNWzfjUeCLraNtvT
# X4/edIhJEqGCAs4wggI3AgEBMIH4oYHQpIHNMIHKMQswCQYDVQQGEwJVUzELMAkG
# A1UECBMCV0ExEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
# b3Jwb3JhdGlvbjEtMCsGA1UECxMkTWljcm9zb2Z0IElyZWxhbmQgT3BlcmF0aW9u
# cyBMaW1pdGVkMSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVTTjpBMjQwLTRCODItMTMw
# RTElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgc2VydmljZaIjCgEBMAcG
# BSsOAwIaAxUAxnmkjOXedpqyHQqkJGn7ewhSC9GggYMwgYCkfjB8MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQg
# VGltZS1TdGFtcCBQQ0EgMjAxMDANBgkqhkiG9w0BAQUFAAIFAOFHWX4wIhgPMjAx
# OTEwMDkwMjUyNDZaGA8yMDE5MTAxMDAyNTI0NlowdzA9BgorBgEEAYRZCgQBMS8w
# LTAKAgUA4UdZfgIBADAKAgEAAgIGcAIB/zAHAgEAAgISSTAKAgUA4Uiq/gIBADA2
# BgorBgEEAYRZCgQCMSgwJjAMBgorBgEEAYRZCgMCoAowCAIBAAIDB6EgoQowCAIB
# AAIDAYagMA0GCSqGSIb3DQEBBQUAA4GBAKtZuZ1w/2ga+qtAMi945BjS8g15dyV9
# SV5G1ZMrVbjnQ/CBXC+ql6otquikxFzD1sO6NIYrsTIckYyFblVnTbJcKdj/vOep
# /TyFVBqV/qKPrx0A5SLOnIZuiKcQltp33njmfp4Tnaf51+jwMuyVSCrSlGkVeqRT
# 2Blz6Hl5IEg5MYIDDTCCAwkCAQEwgZMwfDELMAkGA1UEBhMCVVMxEzARBgNVBAgT
# Cldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29m
# dCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENB
# IDIwMTACEzMAAADgship1NHCtPcAAAAAAOAwDQYJYIZIAWUDBAIBBQCgggFKMBoG
# CSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAvBgkqhkiG9w0BCQQxIgQgtPnrnnMn
# RQq7A/HH0Q0g3PQuW6S2fRG6Yc4qnzFdVbcwgfoGCyqGSIb3DQEJEAIvMYHqMIHn
# MIHkMIG9BCClgS9VpDMosldyg1GQPVVk5wwNOD+Pcl2aoLvRrEJfkDCBmDCBgKR+
# MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMT
# HU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAAA4LIYqdTRwrT3AAAA
# AADgMCIEILumIgBh5YxOY+Xf5dhi2vE1CPgvzpibhbNQ08+tBOoIMA0GCSqGSIb3
# DQEBCwUABIIBAER89l3W4HODj9wEtVeCerFPGrJxMumJoZpw0bdr3VLzX1+DTZsf
# dh5BAJF9WM4O65AiYBHCbw6n6kXh6YdEwqLJF12FYKNyjLiO3ReBPEe3VNaJAwFT
# YP5J3JSCmwHwdya2Uho2dhyeeSIm5ufe4/A7Cg6NAxH7JTA9DcBjVlx/Ky0mbz73
# UXC4ijPROHUJeE2FJ9S9636sJRRNHkiji/MKWV36p6uCn/yE+tU+QkEQ+ibLSzly
# Tq+tnrrA0rIYC2JSpGXSWQQGJxgm0RZPnRTSreDsq9ML6WPTbmBTxyAq2rXnVYvn
# BxNXVoUjiGnarwhPdVQUHrfU61+nES6U2cs=
# SIG # End signature block
