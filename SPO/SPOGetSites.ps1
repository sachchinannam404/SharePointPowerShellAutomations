Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll" 
Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll" 

$siteUrl = “https://mytenant.sharepoint.com/sites/mysitecollection”
$username = "admin@mytenant.onmicrosoft.com"
$password = Read-Host -Prompt "Enter password" -AsSecureString 
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl) 
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $password) 
$ctx.Credentials = $credentials

$rootWeb = $ctx.Web 
$sites  = $rootWeb.Webs

$ctx.Load($rootWeb)
$ctx.Load($sites)
$ctx.ExecuteQuery()

foreach($site in $sites)
{
    $ctx.Load($site)
    $ctx.ExecuteQuery()

    Write-Host $site.Title "-" $site.Url 
}

