
$SP_URL =  read-Host -Prompt 'input thr url'
$FileName = Read-Host -Prompt 'Input the fileName'
$restore = Read-Host -prompt "Restore? [y/n]"
Connect-PnPOnline -Url $SP_URL -CurrentCredentials

if($restore -eq 'y'){
Get-PnPRecycleBinItem -firststage | ? {($_.Title -like $FileName)} | Restore-PnpRecycleBinItem -Force
Get-PnPRecycleBinItem -secondstage | ? {($_.Title -like $FileName)} | Restore-PnpRecycleBinItem -Force
}else{
Get-PnPRecycleBinItem -firststage | ? {($_.Title -like $FileName)}
Get-PnPRecycleBinItem -secondstage | ? {($_.Title -like $FileName)}
}

