if ( (Get-PSSnapin -Name microsoft.sharepoint.powershell -ErrorAction SilentlyContinue) -eq $null )
{
    Add-PsSnapin microsoft.sharepoint.powershell
}
function Start-AuditReportInterface {
<#
.Synopsis
   This module will extract the audit log events for the specified time period, at the indicated location.
.DESCRIPTION
   This module will extract the audit log events for the specified time period, at the indicated location.
   There are 3 options: 
    - Extract just the <new file uploaded> events from the document libraries;
    - Extract the standard Audit log report with a custom date interval;
    - Extract a the same report as in option two, but with a more user friendly set of data
#>
function Get-DocumentLibStatistics {
    [CmdletBinding()]
    param()

$site=Get-SPWeb (Read-Host "Please provide the site URL")
Write-Host "Retrieved Site object" -ForegroundColor Green
#minimum <created date> value
$format=(get-date).ToShortDateString()
[string]$SaveFormat="Doc_Stats_" + ((get-date).Day).ToString() + "_" + (get-date).Month.ToString() + "_" + (get-date).Year.ToString() + ".csv"
$StartD=(Get-Date -Date "$(Read-Host "Please enter the oldest item date you want to start with in the format: $($format)")").ToShortDateString()
    do {
    [string]$location=(Read-Host "Type the path, where you want to save the results") 
    }
until (Test-Path $location)
$location =$location + $saveformat
    Write-Host "Checking the site lists" -ForegroundColor Green
$lists=$site.Lists
foreach ($list in $lists) {
$count=$list.Items.Count-1
    if ($list.basetemplate -eq "DocumentLibrary" -and $count -ge "1" -and $list.Title -ne "Style Library" -and $list.Title -ne "Site assets") {
    Write-Host "Found document library: $($list.Title)" -ForegroundColor Green
    foreach ($doc in $list.Items) {
     if ($doc.file.TimeCreated -ne $null -and $doc.file.TimeLastModified -ne $null) {
     [datetime]$x=$doc.file.TimeCreated 
     if ($x -ge $StartD -and $x -ne $null) {
      Write-Host "Date created: $($x)"
      Write-Host "File: $($doc.Name)" -ForegroundColor DarkCyan
        $createdD=$x.ToShortDateString()
        $createdT=$x.ToLongTimeString()
        $properties = @{
                        'App Id'="";
                        'Event Data'="";
                        'Event Source'="SharePoint";
                        'Custom Event Name'=" ";
                        'Event'="Created";
                        'Occurred (GMT)'="$($createdD) $($createdT)";
                        'Document Location'="$($site.URL)/$($doc.URL)";
                        'User ID'="$($doc.file.author.UserLogin)";
                        'Item Type'="Document";
                        'Item Id'="$($doc.file.UniqueId)";
                        'Site Id'="$($site.ID)";

        }#properties
        $obj = New-Object –TypeName PSObject –Property $properties
       
    Write-Output $obj|Select-Object 'Site Id','Item Id','Item Type','User ID','Document Location','Occurred (GMT)','Event','Custom Event Name','Event Source','Event Data','App Id' |Export-Csv -Path "$($location)" -Delimiter "," -Encoding UTF8 -Append -NoTypeInformation -Force
    } #time validation
    } #date validation
    } #file loop
    }#template validation
  }#list for loop ends
  Write-Host "Report generated here - $($location)" -ForegroundColor Green
 }# function ends


function Get-AuditReport {
#values for the audit log
$format=(get-date).ToShortDateString()
    Do {
    [string]$location=Read-Host "Type the path, where you want to save the results" }
    until (Test-Path $location)
[string]$SaveFormat="Custom_AuditQuery_" + ((get-date).Day).ToString() + "_" + (get-date).Month.ToString() + "_" + (get-date).Year.ToString() + ".csv"
$location=$location + $SaveFormat
$StartD=Get-Date -Date "$(Read-Host "Please enter the start date in the format: $($format)")"
$EndD=Get-Date -Date "$(Read-Host "Please enter the end date in the format: $($format)")"
$s1 = Get-SPsite (Read-Host "Please enter the site, to run the query against")
$q1 = New-Object Microsoft.SharePoint.SPAuditQuery($s1)
$q1.SetRangeStart($StartD)
$q1.SetRangeEnd($EndD)
$s1.Audit.GetEntries($q1) | select @{label='Site Id';e={"$($_.siteid)"}},@{label='Item Id';e={"$($_.itemid)"}},@{label='Item Type';e={$_.itemtype}},@{label='User Id';e={$_.userid}},@{label='Document location';e={$_.doclocation}},@{label='Occurred (GMT)';e={$_.occurred}},@{label="Event";e={$_.eventname}},@{label='Custom Event Name';e={$_.eventsource}},@{label='Event Data';e={$_.eventdata}},@{label='App Id';e={$_.appprincipalid}} |Export-Csv -Path $location -Delimiter "," -Encoding UTF8 -Append -NoTypeInformation -Force
    Write-Host "Report generated here - $($location)" -ForegroundColor Green
} 

function get-CustomAuditReport {
$format=(get-date).ToShortDateString()
    Do {
    [string]$location=Read-Host "Type the path, where you want to save the results" }
    until (Test-Path $location)
[string]$SaveFormat="CustomFormat_Audit_" + ((get-date).Day).ToString() + "_" + (get-date).Month.ToString() + "_" + (get-date).Year.ToString() + ".csv"
$location=$location + $SaveFormat
$StartD=Get-Date -Date "$(Read-Host "Please enter the start date in the format: $($format)")"
$EndD=Get-Date -Date "$(Read-Host "Please enter the end date in the format: $($format)")"
$s1 = Get-SPsite (Read-Host "Please enter the site, to run the query against")
$q1 = New-Object Microsoft.SharePoint.SPAuditQuery($s1)
$q1.SetRangeStart($StartD)
$q1.SetRangeEnd($EndD)
$entries=$s1.Audit.GetEntries($q1)
$w=get-spweb $s1.Url
Write-Host "Looping through the Audit entries to customize the output. This will take a while."
    foreach ($entry in $entries) {
        Write-Host "*" -NoNewline -ForegroundColor Red
        Write-Host "*" -NoNewline -ForegroundColor Green
        Write-Host "*" -NoNewline -ForegroundColor White
        Write-Host "*" -NoNewline -ForegroundColor Black
        $occurredD=$entry.occurred.toshortdatestring()
        $occurredT=$entry.occurred.ToShortTimeString()
        $occurred=$occurredD + $occurredT
        $uid=$entry.userid
        if ($uid -gt 0) {
            $userID=$W.allusers.getbyid($uid).userlogin + " " + "<" + $W.allusers.getbyid($uid).name + ">"
            $DocLoc=$entry.DocLocation
            $properties = @{
                                    'App Id'="$($entry.appprincipalid)";
                                    'Event Data'="$($entry.eventdata)";
                                    'Event Source'="$($entry.eventsource)";
                                    'Custom Event Name'="$($entry.eventsource)";
                                    'Event'="$($entry.eventname)";
                                    'Occurred (GMT)'="$($occurred)";
                                    'Document Location'="$($DocLoc)";
                                    'User ID'="$($userID)";
                                    'Item Type'="$($entry.itemtype)";
                                    'Item Id'="$($entry.itemId)";
                                    'Site Id'="$($entry.siteID)";
                                    'Site URL'="$($w.Url)";

                    }#properties
                    $obj1 = New-Object –TypeName PSObject –Property $properties
       
                Write-Output $obj1|Select-Object 'Site URL','Site Id','Item Id','Item Type','User ID','Document Location','Occurred (GMT)','Event','Custom Event Name','Event Source','Event Data','App Id' |Export-Csv -Path $location -Delimiter "," -Encoding UTF8 -Append -NoTypeInformation -Force

            }#if end
    }#for loop end
}#function end


do {
  [int]$userMenuChoice = 0
  
  while ( $userMenuChoice -lt 1 -or $userMenuChoice -gt 4) {
   Write-Host " "
    Write-Host "Audit Menu" -BackgroundColor Yellow -ForegroundColor Black
    write-Host "1. File statistics" -ForegroundColor Yellow
    Write-Host "2. Custom Audit Report" -ForegroundColor Yellow
    Write-Host "3. Custom Format Audit Report" -ForegroundColor Yellow
    Write-Host "4. Quit and Exit" -ForegroundColor Yellow

    [int]$userMenuChoice = Read-Host "Please choose an option"

    switch ($userMenuChoice) {
      1{Write-Host "Preparing..." -ForegroundColor DarkGreen
        sleep -Seconds 3
        Get-DocumentLibStatistics
        }
      2{Write-Host "Preparing..." -ForegroundColor DarkGreen
        sleep -Seconds 3
        Get-AuditReport
      }
      3{Write-Host "Preparing..." -ForegroundColor DarkGreen
        sleep -Seconds 3
        get-CustomAuditReport
      }
       4{Write-Host "Thank you for using the module. Have a nice day!" -ForegroundColor DarkGreen
        sleep -Seconds 1
      }
      default {
      (New-Object Media.SoundPlayer "C:\WINDOWS\Media\Windows User Account Control.wav").Play()
      Write-Host "Invalid Choice"}
    }
  }
} while ( $userMenuChoice -ne 4 )
}