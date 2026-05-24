param($url = "")
Function ReadWebPermissionInheritance($web)
{
	foreach ($subweb in $web.Webs)
	{
	     ReadWebPermissionInheritance($subweb)
	}
	foreach ($list in $web.Lists)
	{
	      ReadListPermissionInheritance($list)
	}
	if ($web.HasUniqueRoleAssignments)
	{
		Write-Host "inheritance broken on WEB:" $web.Url  
	}
}
Function ReadListPermissionInheritance($list)
{
	#read list items => files
	foreach($item in $list.Items)
	{
		if ($item.HasUniqueRoleAssignments)
		{
			Write-Host "WEB:" $list.ParentWebUrl "- LIST:" $list.Title " - inheritance broken on LISTITEM:" $item.Url
		}
	}
	
	#read list folders => folders
	foreach ($folder in $list.Folders) { 
		if ($folder.HasUniqueRoleAssignments)
		{
			Write-Host "WEB:" $list.ParentWebUrl "- LIST:" $list.Title " - inheritance broken on LISTFOLDER:" $folder.Url
		}
	}
	#read the list itself
	if ($list.HasUniqueRoleAssignments)
	{
		Write-Host "WEB:" $list.ParentWebUrl "- inheritance broken on LIST:" $list.Title   
	}
}


if ($url -eq "") 
{ 
    Write-Warning "Please specify a site collection"
    Write-Host "Usage: ./Find-BrokenInheritance.ps1 -url http://portal.contoso.com"
    exit
}

$siteCollection = Get-SPSite $url

$WebApp = $siteCollection.WebApplication
foreach ($Site in $WebApp.Sites)
{
  foreach($spWeb in $Site.AllWebs)
  {
	if (!$spWeb.IsRootWeb)  
	{    
		ReadWebPermissionInheritance($spWeb)
	}
  }	

}





