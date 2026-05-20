<#
.SYNOPSIS
    Get inventory details for SharePoint Online site collections and export to CSV.

.DESCRIPTION
    This script scans SharePoint Online site collections and captures detailed inventory metrics
    including lists, libraries, site pages, subsites, users, groups, storage usage, and more.
    Results are exported to a timestamped CSV file for analysis and reporting.

.PARAMETER TenantAdminUrl
    The SharePoint Tenant Admin URL (required).
    Format: https://tenant-admin.sharepoint.com

.PARAMETER SiteUrl
    Optional. Specific site URL to inventory. If not provided, all sites are scanned.

.PARAMETER OutputPath
    Optional. Path where CSV will be exported.
    Default: .\SiteCollectionInventory_<timestamp>.csv

.EXAMPLE
    # Scan all sites in tenant
    .\Get-SPOSiteCollectionInventory.ps1 -TenantAdminUrl "https://contoso-admin.sharepoint.com"

.EXAMPLE
    # Scan specific site
    .\Get-SPOSiteCollectionInventory.ps1 -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
      -SiteUrl "https://contoso.sharepoint.com/sites/marketing"

.EXAMPLE
    # Specify custom output path
    .\Get-SPOSiteCollectionInventory.ps1 -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
      -OutputPath "C:\Reports\Inventory.csv"

.NOTES
    Author: SharePoint Automation Team
    Version: 1.0.0
    Requires: PowerShell 5.1+, PnP.PowerShell module
#>

param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^https://[a-zA-Z0-9-]+-admin\.sharepoint\.com$')]
    [string]$TenantAdminUrl,

    [Parameter(Mandatory = $false)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath
)

# Set error action preference
$ErrorActionPreference = "Continue"
$WarningPreference = "SilentlyContinue"

# Create log directory if it doesn't exist
$logDirectory = Join-Path -Path (Get-Location) -ChildPath "Logs"
if (-not (Test-Path -Path $logDirectory)) {
    New-Item -ItemType Directory -Path $logDirectory | Out-Null
}

# Create log file
$timestamp = Get-Date -Format "yyyy-MM-dd"
$logFile = Join-Path -Path $logDirectory -ChildPath "SPO-Inventory-$timestamp.log"

# Function to write logs
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "Info"
    )
    $logMessage = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    Write-Host $logMessage
    Add-Content -Path $logFile -Value $logMessage
}

# Function to check and install PnP.PowerShell
function Install-RequiredModules {
    Write-Log "Checking for PnP.PowerShell module..."
    
    $module = Get-Module -Name "PnP.PowerShell" -ListAvailable
    
    if (-not $module) {
        Write-Log "PnP.PowerShell not found. Installing..." "Info"
        try {
            Install-Module -Name PnP.PowerShell -Force -Scope CurrentUser -AllowClobber
            Write-Log "PnP.PowerShell installed successfully" "Success"
        }
        catch {
            Write-Log "Failed to install PnP.PowerShell: $_" "Error"
            exit 1
        }
    }
    else {
        Write-Log "PnP.PowerShell found (Version: $($module.Version))" "Success"
    }
    
    # Import module
    try {
        Import-Module -Name PnP.PowerShell -DisableNameChecking -Force
        Write-Log "PnP.PowerShell module imported successfully" "Success"
    }
    catch {
        Write-Log "Failed to import PnP.PowerShell: $_" "Error"
        exit 1
    }
}

# Function to connect to SharePoint Tenant
function Connect-SPOTenant {
    param([string]$Url)
    
    Write-Log "Connecting to SharePoint Tenant: $Url" "Info"
    try {
        Connect-PnPOnline -Url $Url -Interactive
        Write-Log "Successfully connected to SharePoint Tenant" "Success"
    }
    catch {
        Write-Log "Failed to connect to SharePoint Tenant: $_" "Error"
        exit 1
    }
}

# Function to get all sites
function Get-SPOSitesList {
    Write-Log "Retrieving site collection list..." "Info"
    try {
        $sites = Get-PnPTenantSite
        Write-Log "Found $($sites.Count) site collections" "Success"
        return $sites
    }
    catch {
        Write-Log "Failed to retrieve sites: $_" "Error"
        return @()
    }
}

# Function to get site inventory
function Get-SiteInventory {
    param([string]$SiteUrl)
    
    $inventory = @{
        SiteUrl = $SiteUrl
        SiteTitle = ""
        SiteTemplate = ""
        Owner = ""
        Status = ""
        Lists = 0
        DocumentLibraries = 0
        ImageLibraries = 0
        AssetLibraries = 0
        FormLibraries = 0
        WikiPageLibraries = 0
        SitePages = 0
        Webs = 0
        Users = 0
        Groups = 0
        TotalItems = 0
        StorageUsedMB = 0
        LastModified = ""
        Error = ""
    }
    
    try {
        # Connect to the site
        Connect-PnPOnline -Url $SiteUrl -Interactive
        
        # Get site properties
        $web = Get-PnPWeb -Includes "Title", "WebTemplate", "LastItemModifiedDate", "StorageUsed"
        $inventory.SiteTitle = $web.Title
        $inventory.SiteTemplate = $web.WebTemplate
        $inventory.LastModified = $web.LastItemModifiedDate
        $inventory.StorageUsedMB = [Math]::Round($web.StorageUsed / 1MB, 2)
        $inventory.Status = "Active"
        
        # Get owner
        try {
            $owner = Get-PnPSiteCollectionAdmin | Select-Object -First 1
            $inventory.Owner = $owner.LoginName -replace "i:0#\.f\|membership\|" , ""
        }
        catch {
            $inventory.Owner = "Unknown"
        }
        
        # Get lists and libraries
        $lists = Get-PnPList | Where-Object { $_.Hidden -eq $false }
        
        foreach ($list in $lists) {
            switch ($list.BaseTemplate) {
                100 { $inventory.DocumentLibraries++ }
                101 { $inventory.DocumentLibraries++ }
                109 { $inventory.ImageLibraries++ }
                104 { $inventory.FormLibraries++ }
                123 { $inventory.AssetLibraries++ }
                151 { $inventory.WikiPageLibraries++ }
                default { $inventory.Lists++ }
            }
            
            # Count items
            try {
                $itemCount = (Get-PnPListItem -List $list.Id -PageSize 5000).Count
                $inventory.TotalItems += $itemCount
            }
            catch {
                # Skip if error retrieving items
            }
        }
        
        # Get site pages
        try {
            $pages = Get-PnPListItem -List "Site Pages" -PageSize 5000 | Where-Object { $_["FileLeafRef"] -like "*.aspx" }
            $inventory.SitePages = $pages.Count
        }
        catch {
            $inventory.SitePages = 0
        }
        
        # Get subsites
        try {
            $webs = Get-PnPSubWeb -Recurse
            $inventory.Webs = $webs.Count + 1  # +1 for root web
        }
        catch {
            $inventory.Webs = 1
        }
        
        # Get users
        try {
            $users = Get-PnPUser
            $inventory.Users = $users.Count
        }
        catch {
            $inventory.Users = 0
        }
        
        # Get groups
        try {
            $groups = Get-PnPGroupMembers -All
            $inventory.Groups = $groups.Count
        }
        catch {
            $inventory.Groups = 0
        }
        
        Write-Log "Successfully inventoried site: $SiteUrl (Lists: $($inventory.Lists), Libraries: $($inventory.DocumentLibraries + $inventory.ImageLibraries), Pages: $($inventory.SitePages))" "Success"
        
    }
    catch {
        Write-Log "Error processing site $($SiteUrl): $_" "Error"
        $inventory.Error = $_.Exception.Message
    }
    
    return $inventory
}

# Main execution
try {
    Write-Log "===== SharePoint Site Collection Inventory Scan Started =====" "Info"
    Write-Log "Tenant Admin URL: $TenantAdminUrl" "Info"
    
    # Install required modules
    Install-RequiredModules
    
    # Connect to tenant
    Connect-SPOTenant -Url $TenantAdminUrl
    
    # Get sites to process
    if ($SiteUrl) {
        Write-Log "Processing specific site: $SiteUrl" "Info"
        $sitesToProcess = @($SiteUrl)
    }
    else {
        Write-Log "Processing all sites in tenant" "Info"
        $sitesToProcess = (Get-SPOSitesList).Url
    }
    
    # Process each site
    $inventoryList = @()
    $siteCount = $sitesToProcess.Count
    $currentCount = 0
    
    foreach ($site in $sitesToProcess) {
        $currentCount++
        Write-Log "[$currentCount/$siteCount] Processing site: $site" "Info"
        
        $inventory = Get-SiteInventory -SiteUrl $site
        $inventoryList += $inventory
    }
    
    # Determine output path
    if (-not $OutputPath) {
        $timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
        $OutputPath = Join-Path -Path (Get-Location) -ChildPath "SiteCollectionInventory_$timestamp.csv"
    }
    
    # Export to CSV
    Write-Log "Exporting inventory to CSV: $OutputPath" "Info"
    $inventoryList | Export-Csv -Path $OutputPath -NoTypeInformation -Force
    Write-Log "Successfully exported inventory to: $OutputPath" "Success"
    
    # Disconnect
    Disconnect-PnPOnline
    
    Write-Log "===== SharePoint Site Collection Inventory Scan Completed =====" "Info"
    Write-Log "Total sites processed: $($inventoryList.Count)" "Info"
    Write-Log "Output file: $OutputPath" "Info"
    
}
catch {
    Write-Log "Fatal error: $_" "Error"
    exit 1
}
