<#  
.SYNOPSIS  
       Alert on content source thresholds.
.DESCRIPTION  
       The purpose of this script is to report when various thresholds are not met for each content source.

       Set -Confirm:$false if you are running script from a repeating task. Otherwise, you will be prompted
       to overwrite files and the script will not complete.
         
.EXAMPLE  
.\Report-CrawlFreshness.ps1 -CreateSampleCsvInputFile -CsvFilename "$($Env:temp)\CrawlReport.csv" 
.\Report-CrawlFreshness.ps1 -CsvFilename "$($Env:temp)\CrawlReport.csv" 
  
.EXAMPLE
1.	Run Report script with -CreateSampleCsvInputFile parameter to create a sample input file.

>.\Report-CrawlFreshness.ps1 –CreateSampleCsvInputFile -CsvFilename "$($Env:temp)\CrawlReport.csv" 

2.	Edit the input file with your content sources and tests.

3.	Run the Report script with the input file.  By default, the a list named CrawlFreshnessReport will be
    created in the central admin site.  Otherwise, use the -ListName and -SiteName parameters to specify custom
    values. The list will be created in the site if it doesn’t exit.  Report data appends to list if list does exist.  

>.\Report-CrawlFreshness.ps1 -CsvFilename "$($Env:temp)\CrawlReport.csv" 
- or -
>.\Report-CrawlFreshness.ps1 -CsvFilename "$($Env:temp)\CrawlReport.csv" -SiteName http://mysite.com/sites/data -ListName MyList

4.	The title of each row in the list is the datetime stamp of when the script was run.


.LINK  
This Script - http://gallery.technet.microsoft.com/scriptcenter/

.NOTES  
  File Name : Report-CrawlFreshness.ps1  
  Author    : Brent Groom, Eric Dixon, Dan Pandre
#>  
param(
	[Parameter(Mandatory=$True, ParameterSetName='Report')]
	[Parameter(Mandatory=$True, ParameterSetName='CreateSample')]
    [string]$CsvFilename = "",

	[Parameter(Mandatory=$True, ParameterSetName='CreateSample')]
    [switch]$CreateSampleCsvInputFile = $false,

	[Parameter(ParameterSetName='Report')]
    [switch]$Confirm = $true,

	[Parameter(ParameterSetName='Report')]
    [string]$ListName = "CrawlFreshnessReport",

	[Parameter(ParameterSetName='Report')]
    [string]$SiteName = "",

    [switch]$Whatif = $false
)

Add-PSSnapin Microsoft.SharePoint.Powershell

function CreateExampleCsv($Filename)
{
$csv = @"
"ContentSource", "Property", "Threshold","Type","Operator"
"Default","Errors","100","number","lt"
"Default","Duration","12","timespan:hours","lt"
"Default","AverageRepositoryTime","1000","number","lt"
"Default","AverageCrawlRate","5","number","gt"
"Default","EndTime","3","datetime:days","lt"
"Default","Deletes","100","number","lt"
"Default","Successes","10","number","gt"
"@

    $csv | Out-File $Filename

    Write-Host "Wrote sample configuration file $Filename" -ForegroundColor Green
}


function GetSSA()
{
    $ssas = @(Get-SPEnterpriseSearchServiceApplication)
    if($ssas.Count -gt 1)
    {
        $done = $false
        while(-not $done)
        {
            Write-Host ""
            Write-Host "Enter the number of the SSA"
		    for($i=0; $i -lt $ssas.Count; $i++)
		    {
			    Write-Host ("{0}. {1}" -f ($i+1), $ssas[$i].Name)
		    }
            
            try
            {
                $choice = (Read-Host "Select the SSA")
                $choice = [int]$choice
            }
            catch
            {
                Write-Host "Error: Input '$choice' is not a number. Enter a number between 1 and $($ssas.Count)." -ForegroundColor Red
                continue
            }
            if($choice -lt 1 -or $choice -gt $ssas.Count)
            {
                Write-Host "Error: Input '$choice' must be between 1 and $($ssas.Count)." -ForegroundColor Red
            }
            else
            {
                $done = $true
            }
        }
		$ssa = $ssas[$choice-1]

    }
    else
    {
        $ssa = $ssas[0]
    }


    if ($ssa.Status -ne "Online")
    {
        $ssaStat = $global:ssa.Status
        Write-Host "Expected SSA to have status 'Online', found status: $ssaStat" -ForegroundColor Red
		exit
    }

    Write-Host "Selected SSA: $($ssa.Name)"

    return $ssa
}


function ValidateOutputFile($OutputFile)
{
    if(Test-Path $OutputFile)
    {
        Write-Host "Warning: The file $OutputFile already exists." -ForegroundColor Yellow
        if($Confirm)
        {
            $overwrite = Read-Host "Enter 'Y' to overwrite the file and continue. A backup will be made. Entering any other value will cause the script to exit"
            if($overwrite -ne 'Y' -and $overwrite -ne 'YES')
            {
                Write-Host "Exiting script."
                exit
            }
        }

        $now = Get-Date -Format "yyyyMMddHHmmss"
        $backup = Join-Path ([System.IO.Path]::GetDirectoryName($Outputfile)) `
                            ("{0}.{1}{2}" -f [System.IO.Path]::GetFileNameWithoutExtension($Outputfile), $now, [System.IO.Path]::GetExtension($Outputfile))

        Copy-Item $OutputFile $backup
        Write-Host "Made backup file $backup."
    }
}


function Test($Value1, $Operator, $Value2)
{
    switch($Operator)
    {
        "gt" {return $Value1 -gt $Value2;break}
        "lt" {return $Value1 -lt $Value2;break}
        "eq" {return $Value1 -eq $Value2;break}
        "ne" {return $Value1 -ne $Value2;break}
    }
}


function OperatorToEnglish($Operator)
{
    switch($Operator)
    {
        "gt" {return "greater than";break}
        "lt" {return "less than";break}
        "eq" {return "equal to";break}
        "ne" {return "not equal to";break}
    }
}


function GetCrawlHistory
{
    $crawltype = @{
      [int]1 = "Full"
      [int]2 = "Incremental"
    }
    $ssa = GetSSA
    $cl = [Microsoft.Office.Server.Search.Administration.CrawlLog] $ssa
    $history = $ssa | Get-SPEnterpriseSearchCrawlContentSource | % {
        $cl.GetCrawlHistory(1, $_.id) | % {
            $cs = new-object psobject -property @{
                ContentSource = $_.ContentSourceName
                Type = $crawltype[$_.CrawlType]
                StartTime = $_.CrawlStartTime
                EndTime = $_.CrawlEndTime
                Duration = $_.CrawlDuration
                Successes = $_.Successes
                AverageRepositoryTime = $_.AverageRepositoryTime
                AverageCrawlRate = $_.AverageCrawlRate
                Errors = $_.Errors
                Warnings = $_.Warnings
                Deletes = $_.Deletes
            }
            $cs
       }
    }

    return $history
}


function AnalyzeCrawlHistory($History, $Thresholds)
{
    $analyzed = @()
    foreach($cs in $History)
    {
       if(-not $Thresholds.ContainsKey($cs.ContentSource))
        {
            $source = "default"

        }
        else
        {
            $source = $cs.ContentSource
        }
        foreach($prop in $Thresholds[$source])
        {
            $success = $false
            $message = ""
            $ope = OperatorToEnglish $prop.Operator

            switch -wildcard ($prop.Type)
            {
                "Number" {
                    if($success = (Test -Value1 $cs.($prop.Property) -Operator $prop.Operator -Value2 $prop.Threshold))
                    {
                        $message = "$($prop.Property) of $($cs.($prop.Property)) was $ope $($prop.Threshold)."
                    }
                    else
                    {
                        $message = "$($prop.Property) of $($cs.($prop.Property)) was not $ope $($prop.Threshold)."
                    }
                    break
                }
                "timespan:*" {
                    $units = $prop.Type.Split(':')[1]
                    #recreate the timestamp object
                    $ts = [timespan]$($cs.($prop.Property)) 
                    if($success = (Test -Value1 $ts.($units) -Operator $prop.Operator -Value2 $prop.Threshold))
                    {
                        $message = "$($prop.Property) of $($cs.($prop.Property)) $units was $ope $($prop.Threshold) $units."
                    }
                    else
                    {
                        $message = "$($prop.Property) of $($cs.($prop.Property)) $units was not $ope $($prop.Threshold) $units."
                    }
                    break
                }
                "datetime:*" {
                    $units = $prop.Type.Split(':')[1]
                    # Get today's date
                    $now = Get-Date
                    $starttime = [datetime]$cs.($prop.Property)
                    $ts = New-TimeSpan -Start $starttime -End $now
                    $diff = $ts.($units)

                    # check if the difference is greater than threshold
                    if($success = (Test -Value1 $diff -Operator $prop.Operator -Value2 $prop.Threshold))
                    {
                        $message = "The $($prop.Property) of $($cs.($prop.Property)) was $ope $($prop.Threshold) $units threshold. Total timespan was $diff $units."
                    }
                    else
                    {
                        $message = "The $($prop.Property) of $($cs.($prop.Property)) was not $ope $($prop.Threshold) $units threshold. Total timespan was $diff $units."
                    }
                    break
                }
            }

            $obj = $prop.psobject.Copy()
            $obj | Add-Member -Name Success -Value $success -MemberType NoteProperty
            $obj | Add-Member -Name Actual -Value $($cs.($prop.Property)) -MemberType NoteProperty
            $obj | Add-Member -Name Message -Value $message -MemberType NoteProperty

            if($source -eq "Default")
            {
                $obj.ContentSource = $cs.ContentSource
            }

            $analyzed += $obj
        }
    }

    return $analyzed
}


function ReadCSVFile($Filename)
{
    $thresholds = @{}

    $csv = Import-Csv $Filename 
    foreach($entry in $csv)
    {
        if(-not $thresholds.ContainsKey($entry.ContentSource))
        {
            $thresholds[$entry.ContentSource] = @()
        }

        $obj = New-Object -TypeName PSObject
        $obj | Add-Member -Name "ContentSource" -Value $entry.ContentSource -MemberType NoteProperty
        $obj | Add-Member -Name "Property" -Value $entry.Property -MemberType NoteProperty
        $obj | Add-Member -Name "Threshold" -Value $entry.Threshold -MemberType NoteProperty
        $obj | Add-Member -Name "Type" -Value $entry.Type -MemberType NoteProperty
        $obj | Add-Member -Name "Operator" -Value $entry.Operator -MemberType NoteProperty
        $thresholds[$entry.ContentSource] += $obj
    }

    return $thresholds
}


function WriteOutputToCSV($Output)
{
    $now = Get-Date -Format "yyyyMMddHHmmss"
    $outFile = Join-Path ([System.IO.Path]::GetDirectoryName($CsvFilename)) `
                         ("{0}.{1}.csv" -f [System.IO.Path]::GetFileNameWithoutExtension($CsvFilename), 
                         "analyzed-$now")
    $Output | Export-Csv -NoTypeInformation -Path $outFile
    Write-Host "Wrote crawl analysis output to $outFile" -ForegroundColor Green
}


function Init
{
    if (!([System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")))
    {
        Write-Error "No SharePoint CSOM libraries (Client) could be located. Does this machine have them installed?"
        return $false
    }
    if (!([System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")))
    {
        Write-Error "No SharePoint CSOM libraries (Runtime) could be located. Does this machine have them installed?"
        return $false
    }

    return $true
}


function New-SPList
{
    param
    (
        [Microsoft.SharePoint.Client.ClientContext] $Context = $(throw "Context is required."),
        [array]$FieldNames = $(throw "Field names are required")
    )

    try
    {
        $list = $Context.Web.Lists.GetByTitle($ListName)
        $context.Load($list)
        $context.ExecuteQuery()
        
        if ($list)
        {
            Write-Host "Found existing SPList named $ListName."
            return
        }    
    }
    catch [Microsoft.SharePoint.Client.IdcrlException]
    {
        Write-Error "Unable to connect to site  $_"
        return
    }
    catch
    {
        # create list
		$listinfo = New-Object Microsoft.SharePoint.Client.ListCreationInformation
		$listinfo.TemplateType = [Microsoft.SharePoint.Client.ListTemplateType]::GenericList
		$listinfo.QuickLaunchOption = [Microsoft.SharePoint.Client.QuickLaunchOptions]::On
		$listinfo.Title = $ListName

		$list = $Context.Web.Lists.Add($listinfo)
		$context.ExecuteQuery()

        $defTemplate = '<Field Type="Text" DisplayName="{0}" ReadOnly="FALSE" Name="{1}"/>'
		foreach ($field in $FieldNames)
		{
			$deferred = $list.Fields.AddFieldAsXml(($defTemplate -f $field, $field), $true, [Microsoft.SharePoint.Client.AddFieldOptions]::AddToAllContentTypes)
		}
		$list.Update()
		$context.ExecuteQuery()
    } 
}


function Load-SPList
{
    param
    (
        [Microsoft.SharePoint.Client.ClientContext] $Context = $(throw "Context is required."),
        [array]$FieldNames = $(throw "DataFileName is required."),
        [array]$Objects = $(throw "DataFileName is required.")
    )

    # load list
    $list = $context.Web.Lists.GetByTitle($ListName)
    $context.Load($list.Fields)
    $context.ExecuteQuery()
    

    $properties = $FieldNames
    $processed = 0
    $now = Get-Date -Format "yyyyMMddHHmmss"

    foreach ($object in $Objects)
    {
        $listitem = $list.AddItem((New-Object Microsoft.SharePoint.Client.ListItemCreationInformation))
        $listitem["Title"] = $now
        foreach ($property in $properties)
        {
			#"property = $property long name"
            $propertyshort = $property
            #csom has a max length of 32 on the column name
            if($propertyshort.length -gt 32) 
            {
                $propertyshort = $propertyshort.Substring(0,32) 
            }
			$afield = $row.$property
            $listitem[$propertyshort] = $object.$property
        }
        $listitem.Update()

        # upload in batches
        $processed++
        if ($processed -eq 25)
        {
            $Context.ExecuteQuery()
            $processed = 0
        }
    }

    $context.ExecuteQuery()
}


function Main
{
    if($CreateSampleCsvInputFile)
    {
        if(-not $CsvFilename)
        {
            $CsvFilename = ".\CrawlReport.csv"
        }

        ValidateOutputFile -OutputFile $CsvFilename
        CreateExampleCsv -Filename $CsvFilename

        Write-Host "Done."
        exit
    }

    $thresholds = ReadCSVFile -Filename $CsvFilename
    $history = GetCrawlHistory
    $output = AnalyzeCrawlHistory -History $history -Thresholds $thresholds
    WriteOutputToCSV -Output $output

    if (!(Init))
    {
        return
    }
    

    if([string]::IsNullOrEmpty($SiteName))
    {
        $SiteName = Get-SPWebApplication -includecentraladministration | where {$_.IsAdministrationWebApplication} | Select-Object -ExpandProperty Url
    }


    # establish client context
    $context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteName)
    Write-Host "Loaded context for $SiteName."

    # get field names 
    $properties = ($output | select -First 1).psobject.Properties | %{$_.Name}

    New-SPList -Context $context -FieldNames $properties

    Load-SPList -Context $context -FieldNames $properties -Objects $output
    Write-Host "Loaded items into list: $($SiteName)Lists/$ListName."

    Write-Host "Done."
}


Main
