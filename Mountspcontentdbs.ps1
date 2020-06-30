$currentTime = Get-Date -Format o | foreach {$_ -replace ":", "."}
New-Item $PSScriptRoot\MountSPContentDB-$currentTime.txt -type file
$webAppUrl = Read-Host "Please provide Web Application URL"
$databaseServer = Read-Host "Please provide database server name"

$DbArray =  Get-Content -Path $PSScriptRoot\DB.txt
function TestSPContentDB()
{

Start-Transcript -Path "$PSScriptRoot\MountSPContentDB-$currentTime.txt"
foreach ($dbName in $DbArray)
{
       if($dbName -ne "End Point") {

    try{
    Mount-SPContentDatabase $dbName -DatabaseServer $databaseServer -WebApplication $webAppUrl -ErrorAction Stop
    
    Write-Output "$dbName is Mounted"  
    }
    catch [System.Exception] 
    {
    Write-Host $_.Exception.Message
    }
    }
    }
Stop-Transcript
}
MountSPContentDB
