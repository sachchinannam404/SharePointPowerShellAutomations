<#
--------------------------------------------

SharePoint 2010 PowerShell script to export
all sites within a given Site Collection to
a separate file.

File		: ExportSites.PS1
Author		: David Webb
Date		: 01/10/2013
Version	: 1.0

Revision history:
-----------------
1.0 Initial version

---------------------------------------------
#>

# -------------
# Set variables
# -------------

$SiteCollection = "http://intranet"
$sites = Get-SPWeb -Site $SiteCollection -Limit All
$ExportFolder = "D:\Exports\"

# ------------
# Begin export
# ------------

ForEach ($site In $sites)
{
   $ExportFile = $ExportFolder + $site.Title + ".cmp"
   Write-Host "Exporting" $site.title "-" $site.url
   Export-SPWeb $site.url -Path $ExportFile -IncludeVersions All -IncludeUserSecurity -Force
   Write-Host "Export to" $ExportFile " complete.  File is" (Get-Item $ExportFile).length "bytes." -ForegroundColor "Green"
}
