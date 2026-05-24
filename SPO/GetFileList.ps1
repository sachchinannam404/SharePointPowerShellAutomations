<#
.SYNOPSIS
GetFileList.ps1 iterates SharePoint site collections and sites to return a .CSV file containing a list of documents and their sizes.

.DESCRIPTION
This script expects a site collection or site as input.  You may pass a URL to the -siteCollectionUrl paramter for site collections and to
-site for sites.  Site collections may also be passed by pipe binding (i.e.: Get-SPSite http://myserver/mysitecollection | .\GetFileList.ps1).

The default output of the script will be a CSV file named output.csv.  This can be overridden with the -filename parameter.

The data returned in this script is security trimmed therefore site collection administrator access is required to get complete results.

This script will work with SharePoint 2013 and SharePoint 2010.

WARNING - This script will use a large amount of memory for large site collections.  Use with caution on large site collections in production.

The script takes the following optional input parameters.
-includeVersions (default: true) - includes all versions of each file in the exported list
-includeSystemFiles (default: false) - includes files in system document libraries such as those in /_catalogs
-includePages (default: false) - includes .aspx pages in the exported list
-author (default: null) - includes files for a particular author using user's full name (i.e.: John Smith)

.EXAMPLE
.\GetFileList.ps1 -siteCollectionUrl http://server/sitecollection

This example returns a list of files from the site collection and all subsites including versions but excluding system files and pages.

.EXAMPLE
Get-SPSite http://server/sitecollection | .\GetFileList.ps1 -includeVersions $false

This example uses a SPSite object from a pipe bind and passes it to the script.  Versions are excluded.

.EXAMPLE
.\GetFileList.ps1 -siteUrl http://server/sitecollection/subsite -includeSystemFiles

This example returns a list of files for a subsite and also include system files such as those in the master pages library.

.EXAMPLE
.\GetFileList.ps1 -siteUrl http://server/sitecollection/subsite -author "John Smith"

This example returns a list of files for a subsite for a given author.  Name must match exactly.

#>
Param($siteUrl, 
    $siteCollectionUrl= "http://server/sitecollection",
    $includeVersions = $true, 
    $includeSystemFiles = $false, 
    $includePages = $false, 
    $filename = "output.csv",
    $author,
    [Parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)][Alias('Site')][Microsoft.SharePoint.SPSite]$siteCollection,
    [Parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)][Microsoft.SharePoint.SPWeb]$web
)

Write-Host "Script started at" (Get-Date).ToString()

$global:itemList = New-Object System.Collections.Generic.List[PSObject]

Function GetFilesFromSite([Microsoft.SharePoint.SPWeb] $currentWeb)
{
    foreach($list in $currentWeb.lists)
    {
      Write-Host $list.Title "($($list.ParentWebUrl)/$($list.RootFolder.Url))"
      if ( ($list.BaseType -eq "DocumentLibrary") -and ( ($includeSystemFiles) -or ( ($list.DefaultViewUrl -notlike "*_catalogs*") -and ($list.DefaultViewUrl -notlike "*Style Library*")  ) )  )
      {
        Write-Host "--- Including" $list.Title
        foreach($item in $list.items) {
            if ( ( ($includePages) -or ($item.Url -notlike "*.aspx") ) -and ( ($author -eq $null) -or ($author -like $item.File.Author.DisplayName) ) )
            {
	            $listItem = New-Object PSObject -Property @{SiteCollectionUrl = $currentWeb.Site.Url;
                    SiteUrl = $currentWeb.Url; 
                    DocumentLibrary = $list.Title;
                    Title = $item.Title; 
                    FileSize = ($item.file).length; 
                    Url = $item.Url; 
                    MajorVersion = ($item.file).MajorVersion;
                    ModifiedDate = $item.file.TimeLastModified;
                    Author = $item.File.Author.DisplayName;
                    FileExtension = [System.IO.Path]::GetExtension($item.Url);
                    }
                $global:itemList.Add($listItem)

                if ($includeVersions)
                {
                    foreach($fileVersion in $item.File.Versions)
                    {
                        $listItem = New-Object PSObject -Property @{SiteCollectionUrl = $currentWeb.Site.Url; 
                        SiteUrl = $currentWeb.Url; 
                        DocumentLibrary = $list.Title;
                        Title = $fileVersion.File.Item.Title;
                        FileSize = $fileVersion.Size;
                        Url = $item.Url; 
                        MajorVersion = $fileVersion.VersionLabel;
                        ModifiedDate = $fileVersion.Created;
                        Author = $fileVersion.CreatedBy.DisplayName;
                        FileExtension = [System.IO.Path]::GetExtension($item.Url);
                        }
                        $global:itemList.Add($listItem)
                    }
                }
            }
        }
      } 
    }
}

# get the site collection if it hasn't been passed via pipe but a URL has been
if ( ($siteCollection -eq $null) -and ($siteCollectionUrl -ne $null) ) 
{
    $siteCollection = Get-SPSite -Identity $siteCollectionUrl
}

# no site collection has been passed so check to see if a subsite has been passed instead
if ($siteCollection -eq $null)
{
    if ( ($web -eq $null) -and ($siteUrl -ne $null) )
    {
        $web = Get-SPWeb -Identity $siteUrl  
    }

    GetFilesFromSite $web
    $web.Dispose()
}
else
{
    foreach($web in $siteCollection.AllWebs)
    {
        Write-Host "***" $web.Url "***"
        GetFilesFromSite $web
        $web.Dispose()
    }

    $siteCollection.Dispose()
}

Write-Host "Writing " $filename
$itemList | Select-Object SiteCollectionUrl, SiteUrl, DocumentLibrary, Url, Title, MajorVersion, FileSize, FileExtension, Author, ModifiedDate | Export-Csv $filename -NoType

# release the array since it will be quite large
$itemList = $null

Write-Host "Script finished at" (Get-Date).ToString()