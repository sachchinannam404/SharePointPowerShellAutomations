<#
.SYNOPSIS
    Get document counts for all document librares in your
    Sharepoint sites
.DESCRIPTION
    Retrieve the number of documents in all of the document
    libraries and lists in the designated Sharepoint site and 
    all of the sub-sites for that site.

    Script must be run on the Sharepoint server.
.PARAMETER Site
    Full URL of the site you want to report on.  
.PARAMETER RootSite
    Full URL for the root of your Sharepoint site
.PARAMETER ReportPath
    Full path of the directory you want to save the CSV report
.INPUTS
    None
.OUTPUTS
    CSV: SPListReport.CSV
.EXAMPLE
    Get-SPDocumentCount -Site http://SurlySharepoint/ITDept -RootSite http://SurlySharepoint -ReportPath c:\Reports
    Will retrieve all document libraries in the ITDept site on SurlySharepoint 
    and give you the number of documents/items in every list.  Report will
    be saved at c:\Reports.
.EXAMPLE
    Get-SPDocumentCount -Site http://SurlySharepoint/QA
    Will retrieve all document libraries in the QA site on SurlySharepoint,
    this time using the default values for RootSite and ReportPath as
    designated in the PARAM section.
.NOTES
    Author:            Sachin Annam
    
       
    Changelog:
       1.0             Initial Release

#>

#Param (
    #[Parameter(Mandatory=$true)]
    #[string]$Site="http://sitcollec/us/department/test/",
    #[string]$RootSite = "http://sitcollec",
    #[string]$ReportPath = "D:\reports"
#)
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
$ReportPath = "D:\reports"
$Webs = Get-SPWeb http://sitcollec/us/department/test/
$Result = @()
ForEach ($Web in $Webs.Webs)
{   ForEach ( $List in $Web.Lists )
    {   $Result += New-Object PSObject -Property @{
            'Library Title' = $List.Title
            Count = $List.Folders.Count + $List.Items.Count
            'Site Title' = $Web.Title
            URL = $Web.URL
            'Library Type' = $List.BaseType
			'Last Modified'=$List.LastItemModifiedDate
        }
    }
}
$Result | Select 'Site Title',URL,'Library Type','Library Title',Count,'Last Modified' | Export-Csv "$ReportPath\SPListReport1.csv" -NoTypeInformation
#$Result | Select 'Site Title',URL,'Library Type','Library Title',Count | Out-GridView