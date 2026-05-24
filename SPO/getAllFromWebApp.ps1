Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
Get-SPWebApplication http://test-Sitecoll | Get-SPSite -Limit All  | Select Title, URL | 
Export-CSV D:\reports\SharePoint_Sites_APACReport.csv -NoTypeInformation