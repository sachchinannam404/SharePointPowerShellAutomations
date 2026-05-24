$FolderPath = "\\Backups\Production"
$snapin = Get-PSSnapin | Where-Object {$_.Name -eq 'Microsoft.SharePoint.Powershell'} 
if ($snapin -eq $null) 
{    
                Write-Host "Loading SharePoint Powershell Snapin..."    
                Add-PSSnapin "Microsoft.SharePoint.Powershell" 
}
foreach ($SPWebApplication in (Get-SPWebApplication)) {
    foreach ($Site in $SPWebApplication.Sites) {
        $Filename = $FolderPath + $SPWebApplication.Name.Replace(" ","") + " " + $Site.ServerRelativeUrl.Replace("/","_")+ ".bak"
        Write-Host "$Filename"
        backup-spsite -identity $Site.URL -path $FileName -Force 
    }
}