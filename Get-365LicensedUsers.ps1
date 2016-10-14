Import-Module ActiveDirectory;
Import-Module -Force 'C:\cron\creds.psm1';

$users = Get-MsolUser -All | Where {$_.isLicensed -eq "True" -and ("DPaW:ENTERPRISEPACK" -in $_.licenses.accountSkuId)}
$user=$users | Select UserPrincipalName
Del C:\cron\o365_licensed.csv
$user | Export-Csv C:\cron\o365_licensed.csv