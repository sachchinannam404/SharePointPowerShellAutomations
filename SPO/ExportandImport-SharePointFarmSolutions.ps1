#Exporting Solutions from Source farm

Add-PSSnapin Microsoft.SharePoint.PowerShell
$farm = Get-SPFarm
$farm.Solutions | ForEach-Object{$_.SolutionFile.SaveAs("c:\ExportFolder\" + $_.SolutionFile.name)}

#Adding Solutions to Another Farm

<p style="padding-left: 30px;">Add-PSSnapin Microsoft.SharePoint.PowerShell
$files = Get-ChildItem "c:\ImportFolder\"
ForEach ($file in $files) {Add-SPSolution $file.FullName}
