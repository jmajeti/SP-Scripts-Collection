<#
.Synopsis
   Query share connection numbers
.DESCRIPTION
   Returns the number of connections to Shares on local or remote systems
.EXAMPLE
   Get-ShareConnection -ComputerName Server1,Server2
.EXAMPLE
   Get-ShareConnection -CN Localhost
.INPUTS
   A String or Array of ComputerNames
.OUTPUTS
   An OBJECT with the following properties is returned from this function
   ComputerName,Share(Name),Connections(number)
.NOTES
   General
.FUNCTIONALITY
   Using WMI to query the number of open connections to Shares on local or remote systems
#>
function Get-ShareConnection
{
    Param
    (
        # param1 help description
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0)]
        [Alias("cn")] 
        [String[]]$ComputerName
    )

    Begin
    {
    }
    Process
    {
        $ComputerName | ForEach-Object {
           $Computer = $_
           try {
                Get-WmiObject -Class Win32_ConnectionShare  -Namespace root\cimv2 -ComputerName $Computer -EA Stop | 
                Group-Object Antecedent |
                Select-Object @{Name="ComputerName";Expression={$Computer}},
                              @{Name="Share"       ;Expression={(($_.Name -split "=") | 
                                    Select-Object -Index 1).trim('"')}},
                              @{Name="Connections" ;Expression={$_.Count}}
               }
           catch 
               {
                    Write-Host "Cannot connect to $Computer" -BackgroundColor White -ForegroundColor Red
               }
           
           }#ForEach-Object
    }
    End
    {
    }
}