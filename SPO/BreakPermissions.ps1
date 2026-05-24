 <#  
.SYNOPSIS  
	To break the original permissions on the site/subsite
also deletes the existing groups
.DESCRIPTION  
	
.NOTES  
	File Name  : BreakPermissions.ps1
	Author     : Yogesha H P
	Requires   : PowerShell Version 2.0  
	
.LINK  
	
.PARAMETER File  
	The configuration file
#>

 if ((Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null)
       {
           Add-PSSnapin "Microsoft.SharePoint.PowerShell"
       }

#Get Site collection and open the web

#$spSite = Get-SPSite "http://servername/sites/TestYogesh"
$spSiteUrl = Read-Host "Enter the site url"

$spSite = Get-SPWeb $spSiteUrl
If ($spsite.IsRootWeb -eq $true)
    {
$spWeb = $spSite
Write-Host ($spSite.Url + "This is the Site collection")
# To delete existing permissions presee yes
$yes = Read-Host "Enter Yes to delete existing permissions on site collections"

if($yes -eq "Yes")
        {

       # Remove unnecessary groups/users from the site permissions
        for ($i = 0; $i -le $spWeb.RoleDefinitions.Count; $i++)
                    {
                    $spWeb.RoleDefinitions.Delete($i)
                    }
        }

    }


else

{
$spweb = Get-SPWeb $spSiteUrl
 If ( ($spWeb.HasUniqueRoleAssignments -and $spWeb.HasUniqueRoleDefinitions) -eq $false )

 {
 $spWeb.RoleDefinitions.BreakInheritance($true, $true)
 #$spWeb.BreakRoleInheritance($true,$true)
 $spWeb.Update()
 Write-Host ($spSite.Url + "Inheritance is Just broken")
             # Remove unnecessary groups/users from the site permissions
            for ($i = 0; $i -le $spWeb.RoleDefinitions.Count; $i++)
            {
            $spWeb.RoleDefinitions.Delete($i)
            }
 }
 else
 {
 Write-Host ($spSite.Url + "Inheritance is already broken")
             # Remove unnecessary groups/users from the site permissions
            for ($i = 0; $i -le $spWeb.RoleDefinitions.Count; $i++)
            {

            $spWeb.RoleDefinitions.Delete($i)
            }
            
 }


}


$spWeb.Dispose()
$spsite.Dispose()
