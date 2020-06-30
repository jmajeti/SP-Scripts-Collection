#Enter URl with '/' at the end
$siteURL = "https://www.abcdefgxyz.com"

#Do not enter '\' at the end
$outputDirectory = "C:\New Folder"

#Include Library
Add-PSSnapin "Microsoft.SharePoint.PowerShell"

#Group Name in which Term Set is present
$groupName = "exampleGroupName"

#Term Set Name
$termSetName = "exampleTermSetName"

#Function to Export to Terms to CSV
function ExportTermsToCSV() {
    param($SiteURL, $OutputDirectory, $groupLabel, $termSetName)
    $empty = ""
    $taxonomySiteURL = Get-SPSite -Identity $SiteURL
    #Connect to Term Store in the Managed Metadata Service Application
    $taxonomySession = Get-SPTaxonomySession -site $taxonomySiteURL
    $taxonomyTermStore = $taxonomySession.TermStores | Select Name
    $termStore = $taxonomySession.TermStores[$taxonomyTermStore.Name]
    #Mention Group Name in which Term Set Exists
    $group = $termStore.Groups | Where-Object {$_.Name -eq $groupLabel}
    #Mention Term Set Name in which Terms Exists
    $termSet = $group.TermSets | Where-Object {$_.Name -eq $termSetName}
    $outputFile = $OutputDirectory + '\' + $termSet.Name + 'GUID.csv'
    foreach($term in $termSet.GetAllTerms()){
        $eachTerm = $term
        $index = 0
        #Mention Maximun Level of the Term to be Iterated
        $maxTermLevel = 3
        #Increase or Decrease "" into $levelArray depending on the levels present in Term Set
        $levelArray = @("", "", "", "")
        while (!$eachTerm.IsRoot) {
            $eachTerm = $eachTerm.Parent
            $index = $index + 1
        }
        $indexNew = $index
        while ($index -le $maxTermLevel){
            $levelArray[$index] = $empty
            $index = $index + 1
        }
        $eachTerm = $term
        while ($indexNew -ge 0){
            $levelArray[$indexNew] = $eachTerm.Name
            $eachTerm = $eachTerm.Parent
            $indexNew = $indexNew - 1
        }
        $TermHierarchy = $levelArray[0]
        if ($levelArray[1] -ne ""){
            $TermHierarchy += ":" + $levelArray[1]
            if ($levelArray[2] -ne ""){
                $TermHierarchy += ":" + $levelArray[2]
            }
        }
        $TermAndGUID = $term.Name + "|" + $term.Id
        #create a new line in the CSV file
        $newRecord = NewCSVRecord -TermName $term.Name -TermHierarchy $TermHierarchy -TermGUID $term.Id -TermAndGUID $TermAndGUID
        #add the new line
        $allRecords = [Array]$allRecords + $newRecord
    }
    #export all of the terms to a CSV file
    $allRecords | Export-Csv $outputFile -Encoding UTF8
    $taxonomySiteURL.dispose()
}

#Function to Create New Record (Each Term in TermSet) for CSV
function NewCSVRecord() {
    param($TermName, $TermHierarchy, $TermGUID,$TermAndGUID)
    $term = New-Object PSObject
    $term | Add-Member -Name "Term Hierarchy" -MemberType NoteProperty -Value $TermHierarchy
    $term | Add-Member -Name "Term Name" -MemberType NoteProperty -Value $TermName
    $term | Add-Member -Name "Term GUID" -MemberType NoteProperty -Value $TermGUID
    $term | Add-Member -Name "Term|GUID" -MemberType NoteProperty -Value $TermAndGUID
    return $term
}

#Call ExportTermsToCSV Function
ExportTermsToCSV -SiteURL $siteURL -OutputDirectory $outputDirectory -groupLabel $groupName -termSetLabel $termSetName