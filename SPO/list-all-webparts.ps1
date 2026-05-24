#----------------------------------------------------------------------------- 
# Name:             list-all-webparts.ps1  
# Description:      This script will create a list of all web parts in all 
#                   pages in a site collection.
#                     
# Usage:            Run the script passing three paramters: Url, Folder and WP. 
#					Url Accepts multiple values (comma separated).
#                   Folder is the destination of the report.
#					WP is the name of the web part (optional).
# By:               Riccardo Celesti blog.riccardocelesti.it
#----------------------------------------------------------------------------- 


Param([Parameter(Mandatory=$true)] 
      [String] 
      $Url,
      [Parameter(Mandatory=$true)] 
      [ValidateScript({Test-Path $_ -PathType 'Container'})]
      [String] 
      $folder,
	  [Parameter(Mandatory=$false)]
      [String] 
      $WP
) 

if ((gsnp Microsoft.SharePoint.Powershell -EA SilentlyContinue) -eq $null){
    asnp Microsoft.SharePoint.Powershell -EA Stop
}

$filename = "WebPartsReport_" + (Get-Date).ToFileTimeUtc().ToString() + ".csv"
$filenamewp = "WebPartsReport_" + $WP.Replace(" ","-") + "_" + (Get-Date).ToFileTimeUtc().ToString() + ".csv"

$logfile = Join-Path $folder $filename
$logfilewp = Join-Path $folder $filenamewp

$urlArray = $Url.Split(",")

$header = "File Url, Web Part Title, Web Part Type, Visible"

ac $logfile $header
ac $logfilewp $header

$logfilecontrol = $null

foreach ($SPsite in $urlArray){
    Get-SPSite $SPsite| % {

        foreach ($web in $_.AllWebs){
        
            if ([Microsoft.SharePoint.Publishing.PublishingWeb]::IsPublishingWeb($web)){

                $library = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($web)
                $pages = $library.PagesList
            
                foreach ($file in $pages.Items){

                    $fileUrl = $web.Url + "/" + $file.File.Url

                    $manager = $file.file.GetLimitedWebPartManager([System.Web.UI.WebControls.Webparts.PersonalizationScope]::Shared);

                    $webparts = $manager.webparts

                    $webparts | %{
                        ac $logfile "$fileUrl, $($_.DisplayTitle), $($_.GetType().ToString()), $($_.IsVisible)"
						
						if ($_.DisplayTitle -match $WP -and -not [string]::IsNullOrEmpty($WP)){
							ac $logfilewp "$fileUrl, $($_.DisplayTitle), $($_.GetType().ToString()), $($_.IsVisible)"

                            $logfilecontrol = 1
						}
                    }
                }

                $sitepages = [Microsoft.Sharepoint.Utilities.SpUtility]::GetLocalizedString('$Resources:WikiLibDefaultTitle',"core",$web.UICulture.LCID)

                $pages = $null
                $pages = $web.Lists[$sitepages]

                if ($pages -and $pages.ItemCount -gt 0){

                    foreach ($file in $pages.Items) {
                        $fileUrl = $web.Url + "/" + $file.File.Url

                        $manager = $file.file.GetLimitedWebPartManager([System.Web.UI.WebControls.Webparts.PersonalizationScope]::Shared);

                        $webparts = $manager.webparts

                        $webparts | %{
                            ac $logfile "$fileUrl, $($_.DisplayTitle), $($_.GetType().ToString()), $($_.IsVisible)"
							
							if ($_.DisplayTitle -match $WP -and -not [string]::IsNullOrEmpty($WP)){
								ac $logfilewp "$fileUrl, $($_.DisplayTitle), $($_.GetType().ToString()), $($_.IsVisible)"

                                $logfilecontrol = 1
							}
                        }   
                    }
                }
            } else {
                $sitepages = [Microsoft.Sharepoint.Utilities.SpUtility]::GetLocalizedString('$Resources:WikiLibDefaultTitle',"core",$web.UICulture.LCID)

                $pages = $null
                $pages = $web.Lists[$sitepages]

                if ($pages){

                    foreach ($file in $pages.Items) {
                        $fileUrl = $web.Url + "/" + $file.File.Url

                        $manager = $file.file.GetLimitedWebPartManager([System.Web.UI.WebControls.Webparts.PersonalizationScope]::Shared);

                        $webparts = $manager.webparts

                        $webparts | %{
                            ac $logfile "$fileUrl, $($_.DisplayTitle), $($_.GetType().ToString()), $($_.IsVisible)"
                            
                            if ($_.DisplayTitle -match $WP -and -not [string]::IsNullOrEmpty($WP)){
								ac $logfilewp "$fileUrl, $($_.DisplayTitle), $($_.GetType().ToString()), $($_.IsVisible)"

                                $logfilecontrol = 1
							}
                        }
                    }
                }
            }
        }
    }
}

.\notepad.exe $logfile

if ($logfilecontrol -eq 1){
    .\notepad.exe $logfilewp
}