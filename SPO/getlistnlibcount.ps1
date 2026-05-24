Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
$ListsInfo = @{}
$TotalItems = 0
$SiteCollection = Get-SPSite http://sitecoll/sites/test/
ForEach ($Site in $SiteCollection.AllWebs)
{
    ForEach ($List in $Site.Lists)
    {
        $ListURL = $Site.url +"/"+ $List.RootFolder.Url
        $ListsInfo.Add($List.Title + " - " + $ListURL, $List.ItemCount)
        $TotalItems += $List.ItemCount
	$ModifiedAt+=$List.LastItemModifiedDate
    }
}
$ListsInfo.GetEnumerator() | sort name | Export-Csv D:\reports\test.txt