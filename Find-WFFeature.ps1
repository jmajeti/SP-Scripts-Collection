if ((Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null) 
[PSObject[]]$global:resultsarray = @()
                }
                $wFeatureName = $ns2Find.ToLower()
                if($nsDetail -like "*$wFeatureName*")
                        #Find the offending Namespace Prefix
                        $nodes = $wXml.rootworkflowactivitywithdata
                        if ($nodes)
                        {
                            #Get the prefix of the namespaces we're looking for
                            $namespaces = $nodes.Attributes | Where-Object { $_.Prefix -eq 'xmlns' -and $_.Value -like "*$ns2Find*"}
                            foreach ($nameentry in $namespaces)
                            {
                                #$nsPrefix = $nodes.GetPrefixOfNamespace($nameentry.Value)
                                #if ($nsPrefix)
                                #{
                                    #Get the full Namespace
                                    #$fullNamespace = $nodes.GetNamespaceOfPrefix($nsPrefix)
                                    #Create the Namespace Object
                                    $fullns = @{$($nameentry.LocalName)="$($nameentry.Value)"}
                                    #Locate the nodes list
                                    $foundNodes = Select-Xml -Xml $wXml -Namespace $fullns -XPath "//$($nameentry.LocalName):*" | Select -ExpandProperty Node
                                    foreach($foundNode in $foundNodes)
                                        $outObject = new-object PSObject
                                        $outObject | add-member -membertype NoteProperty -name "Notes" -Value ""
                                        $global:resultsarray += $outObject
                                    }
                                #}
                            }
                        }
                }
foreach($webApp in $WebApplications)
                #Write-Host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
                Write-Host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red