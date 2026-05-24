<#
.SYNOPSIS
The script create a new SharePoint group.

.DESCRIPTION
The script create a new SharePoint group on a specified site where you need. 
You should give the site url, group name and the permission level which should be
the basic level as "Full Control","Design","Edit","Contribute","Read","View Only".
You could modify the group deault settings as OnlyAllowMembersViewMembership, 
AllowMembersEditMembership, AllowRequestToJoinLeave, AutoAcceptRequestToJoinLeave.

.EXAMPLE
New-SPGroup.ps1 -SiteUrl "http://site" -GroupName "Site Owners" -PermissionLevel "Contribute"
-Descriptipon "Member of this group can edit lists, document libraries, and pages in the site"
#>

#NAME: New-SPGroup.ps1
#AUTHOR: Tibor Revesz
#DATE: 05/12/2013

param (
    [Parameter(Mandatory = $true, `
		ValueFromPipeline = $true)]
    [string]$SiteUrl = "",
    [Parameter(Mandatory = $true)]
    [string]$GroupName = "",
    [Parameter(Mandatory = $true)]
    [ValidateSet("Full Control","Design",`
		"Edit","Contribute","Read","View Only")]
    [string]$PermissionLevel = "View Only",	
    [string]$Description = "",
    [ValidateSet($true,$false)]
    [string]$OnlyAllowMembersViewMembership = $true,
    [ValidateSet($true,$false)]
    [string]$AllowMembersEditMembership = $false,
    [ValidateSet($true,$false)]
    [string]$AllowRequestToJoinLeave = $false,
    [ValidateSet($true,$false)]
    [string]$AutoAcceptRequestToJoinLeave = $false,
    [System.Net.Mail.MailAddress]$RequestToJoinLeaveEmailSetting 
    )

#Set site url
$web = Get-SPWeb $SiteUrl

#Create a new group
$web.SiteGroups.Add($GroupName,$web.CurrentUser,$web.CurrentUser,$Description)

#Customize the group settings
$Group = $web.SiteGroups[$GroupName]
$Group.OnlyAllowMembersViewMembership = $OnlyAllowMembersViewMembership
$Group.AllowMembersEditMembership = $AllowMembersEditMembership
$Group.AllowRequestToJoinLeave = $AllowRequestToJoinLeave
$Group.AutoAcceptRequestToJoinLeave = $AutoAcceptRequestToJoinLeave
$Group.RequestToJoinLeaveEmailSetting = $RequestToJoinLeaveEmailSetting
$Group.Update()

#Create a new assignment (group and permission level pair)
$GroupAssignment = New-Object Microsoft.SharePoint.SPRoleAssignment($Group)

#Get the permission levels to apply to the new group
$RoleDefinition = $web.Site.RootWeb.RoleDefinitions[$PermissionLevel]

#Assign the appropriate permission level to group
$GroupAssignment.RoleDefinitionBindings.Add($RoleDefinition)

#Add group to the site
$web.RoleAssignments.Add($GroupAssignment)