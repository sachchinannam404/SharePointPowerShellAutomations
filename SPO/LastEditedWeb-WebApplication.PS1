Add-PSSnapin Microsoft.sharepoint.powershell -ErrorAction SilentlyContinue
$Webapp = "http://contoso.com" 
$site = Get-SPSite -webapplication $Webapp -Limit All 
$site | Get-SpWeb -Limit All | Format-Table -Property LastItemModifiedDate,URL -AutoSize | Out-string -width 3000 > output_webapp.txt