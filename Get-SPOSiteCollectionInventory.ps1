<#
.SYNOPSIS
    Get SharePoint Online Site Collection Inventory - Comprehensive audit and inventory script for SharePoint Online.

.DESCRIPTION
    This script audits SharePoint Online site collections and captures detailed inventory metrics including:
    - Lists and Libraries (Document, Image, Asset, Form, Wiki)
    - Site Pages count
    - Subsites
    - Users and Groups
    - Total items and storage usage
    - Site properties and last modified date
    
    Results are exported to a timestamped CSV file for reporting and analysis.

.PARAMETER TenantAdminUrl
    The SharePoint Tenant Admin URL in format: https://tenant-admin.sharepoint.com
    Required for authentication to SharePoint Online.

.PARAMETER SiteUrl
    Optional. Specific site collection URL to inventory. If not provided, all site collections are scanned.
    Example: https://contoso.sharepoint.com/sites/marketing

.PARAMETER OutputPath
    Optional. Path where the CSV file will be exported. 
    Default: .\SiteCollectionInventory_<timestamp>.csv

.EXAMPLE
    # Scan all site collections
    .\Get-SPOSiteCollectionInventory.ps1 -TenantAdminUrl "https://contoso-admin.sharepoint.com"

.EXAMPLE
    # Scan specific site collection
    .\Get-SPOSiteCollectionInventory.ps1 `
      -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
      -SiteUrl "https://contoso.sharepoint.com/sites/marketing"

.EXAMPLE
    # Specify output location
    .\Get-SPOSiteCollectionInventory.ps1 `
      -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
      -OutputPath "C:\Reports\Inventory.csv"

.NOTES
    Author: SharePoint Automation
    Version: 1.0.0
    Created: 2026-05-20
    Requirements:
        - PowerShell 5.1 or higher
        - PnP.PowerShell module
        - SharePoint Online tenant admin access
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$TenantAdminUrl,
    
    [Parameter(Mandatory = $false)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath
)

# ====================================
# Helper Functions
# ====================================

function Write-Log {
    <#
    .SYNOPSIS
        Write log messages to console and log file with timestamp
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("Info", "Success", "Warning", "Error")]
        [string]$Level = "Info"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Write to console
    switch ($Level) {
        "Success" { Write-Host $logMessage -ForegroundColor Green }
        "Warning" { Write-Host $logMessage -ForegroundColor Yellow }
        "Error" { Write-Host $logMessage -ForegroundColor Red }
        default { Write-Host $logMessage }
    }
    
    # Write to log file
    if (-not (Test-Path ".\Logs")) {
        New-Item -ItemType Directory -Path ".\Logs" -ErrorAction SilentlyContinue | Out-Null
    }
    
    $logFile = ".\Logs\SPO-Inventory-$(Get-Date -Format 'yyyy-MM-dd').log"
    Add-Content -Path $logFile -Value $logMessage
}

function Install-RequiredModules {
    <#
    .SYNOPSIS
        Check and install required PowerShell modules
    #>
    Write-Log "Checking for PnP.PowerShell module..." "Info"
    
    $pnpModule = Get-Module -Name PnP.PowerShell -ListAvailable -ErrorAction SilentlyContinue
    
    if ($null -eq $pnpModule) {
        Write-Log "PnP.PowerShell not found. Installing..." "Warning"
        try {
            Install-Module -Name PnP.PowerShell -Force -Scope CurrentUser -AllowClobber
            Write-Log "PnP.PowerShell installed successfully" "Success"
        }
        catch {
            Write-Log "Failed to install PnP.PowerShell: $_" "Error"
            throw
        }
    }
    else {
        Write-Log "PnP.PowerShell module found (Version: $($pnpModule.Version))" "Success"
    }
    
    # Import module
    Import-Module -Name PnP.PowerShell -Force -ErrorAction Stop
    Write-Log "PnP.PowerShell module loaded" "Success"
}

function Connect-SPOTenant {
    <#
    .SYNOPSIS
        Connect to SharePoint Online Tenant Admin
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantAdminUrl
    )
    
    Write-Log "Connecting to SharePoint Tenant: $TenantAdminUrl" "Info"
    
    try {
        $connection = Connect-PnPOnline -Url $TenantAdminUrl -Interactive -ErrorAction Stop
        Write-Log "Successfully connected to SharePoint Tenant Admin" "Success"
        return $connection
    }
    catch {
        Write-Log "Failed to connect to SharePoint Tenant: $_" "Error"
        throw
    }
}

function Get-SPOSitesList {
    <#
    .SYNOPSIS
        Get list of all SharePoint Online site collections
    #>
    Write-Log "Retrieving site collection list..." "Info"
    
    try {
        $sites = Get-PnPTenantSite -Detailed -ErrorAction Stop
        Write-Log "Found $($sites.Count) site collections" "Success"
        return $sites
    }
    catch {
        Write-Log "Failed to retrieve site list: $_" "Error"
        throw
    }
}

function Get-SiteInventory {
    <#
    .SYNOPSIS
        Get inventory details for a single site
    #>
    param(
        [Parameter(Mandatory = $true)]
        $Site
    )
    
    $inventoryObject = [PSCustomObject]@{
        SiteUrl              = $Site.Url
        SiteTitle            = $Site.Title
        SiteTemplate         = $Site.Template
        Owner                = $Site.Owner
        Status               = $Site.Status
        Lists                = 0
        DocumentLibraries    = 0
        ImageLibraries       = 0
        AssetLibraries       = 0
        FormLibraries        = 0
        WikiPageLibraries    = 0
        SitePages            = 0
        Webs                 = 0
        Users                = 0
        Groups               = 0
        TotalItems           = 0
        StorageUsedMB        = [math]::Round($Site.StorageUsageCurrent / 1024, 2)
        LastModified         = $Site.LastContentModifiedDate
        Error                = ""
    }
    
    try {
        # Connect to the specific site
        Connect-PnPOnline -Url $Site.Url -Interactive -ErrorAction Stop | Out-Null
        
        # Get lists and libraries
        $lists = Get-PnPList -ErrorAction SilentlyContinue
        
        foreach ($list in $lists) {
            # Skip system lists
            if ($list.Hidden -eq $false -and $list.Title -notlike "Form Templates" -and $list.BaseTemplate -notlike "*Catalog*") {
                switch ($list.BaseTemplate) {
                    "DocumentLibrary" { $inventoryObject.DocumentLibraries++ }
                    "PictureLibrary" { $inventoryObject.ImageLibraries++ }
                    "AssetLibrary" { $inventoryObject.AssetLibraries++ }
                    "FormLibrary" { $inventoryObject.FormLibraries++ }
                    "WikiPageLibrary" { $inventoryObject.WikiPageLibraries++ }
                    default { $inventoryObject.Lists++ }
                }
                
                # Count items
                try {
                    $itemCount = Get-PnPListItem -List $list.Id -PageSize 5000 -ErrorAction SilentlyContinue | Measure-Object | Select-Object -ExpandProperty Count
                    $inventoryObject.TotalItems += $itemCount
                }
                catch {
                    # Skip if can't access items
                }
            }
        }
        
        # Get site pages
        try {
            $sitePages = Get-PnPListItem -List "Site Pages" -PageSize 5000 -ErrorAction SilentlyContinue | Measure-Object | Select-Object -ExpandProperty Count
            $inventoryObject.SitePages = $sitePages
        }
        catch {
            $inventoryObject.SitePages = 0
        }
        
        # Get webs (subsites)
        try {
            $webs = Get-PnPSubWeb -Recurse -ErrorAction SilentlyContinue
            $inventoryObject.Webs = $webs.Count + 1  # +1 for root web
        }
        catch {
            $inventoryObject.Webs = 1
        }
        
        # Get users
        try {
            $users = Get-PnPUser -ErrorAction SilentlyContinue
            $inventoryObject.Users = $users.Count
        }
        catch {
            $inventoryObject.Users = 0
        }
        
        # Get groups
        try {
            $groups = Get-PnPGroup -ErrorAction SilentlyContinue
            $inventoryObject.Groups = $groups.Count
        }
        catch {
            $inventoryObject.Groups = 0
        }
    }
    catch {
        $inventoryObject.Error = $_.Exception.Message
        Write-Log "Error processing site $($Site.Url): $($_.Exception.Message)" "Warning"
    }
    finally {
        # Disconnect from site
        Disconnect-PnPOnline -ErrorAction SilentlyContinue | Out-Null
    }
    
    return $inventoryObject
}

function Export-InventoryToCSV {
    <#
    .SYNOPSIS
        Export inventory data to CSV file
    #>
    param(
        [Parameter(Mandatory = $true)]
        [array]$InventoryData,
        
        [Parameter(Mandatory = $false)]
        [string]$OutputPath
    )
    
    # Generate output path if not provided
    if ([string]::IsNullOrEmpty($OutputPath)) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $OutputPath = ".\SiteCollectionInventory_$timestamp.csv"
    }
    
    try {
        $InventoryData | Export-Csv -Path $OutputPath -NoTypeInformation -Force -ErrorAction Stop
        Write-Log "Inventory exported successfully to: $OutputPath" "Success"
        return $OutputPath
    }
    catch {
        Write-Log "Failed to export inventory: $_" "Error"
        throw
    }
}

# ====================================
# Main Script Execution
# ====================================

try {
    Write-Log "================================================" "Info"
    Write-Log "SharePoint Online Site Collection Inventory" "Info"
    Write-Log "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" "Info"
    Write-Log "================================================" "Info"
    
    # Check PowerShell version
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        throw "PowerShell 5.0 or higher is required."
    }
    Write-Log "PowerShell version: $($PSVersionTable.PSVersion)" "Info"
    
    # Install required modules
    Install-RequiredModules
    
    # Connect to tenant
    Connect-SPOTenant -TenantAdminUrl $TenantAdminUrl
    
    # Get sites to process
    if ([string]::IsNullOrEmpty($SiteUrl)) {
        $sites = Get-SPOSitesList
        Write-Log "Processing all $($sites.Count) site collections" "Info"
    }
    else {
        Write-Log "Processing specific site: $SiteUrl" "Info"
        try {
            $site = Get-PnPTenantSite -Url $SiteUrl -Detailed -ErrorAction Stop
            $sites = @($site)
        }
        catch {
            Write-Log "Site not found: $SiteUrl" "Error"
            throw
        }
    }
    
    # Process each site
    $inventory = @()
    $siteCount = $sites.Count
    
    foreach ($site in $sites) {
        $currentIndex = $inventory.Count + 1
        Write-Log "[$currentIndex/$siteCount] Processing site: $($site.Url)" "Info"
        
        $siteInventory = Get-SiteInventory -Site $site
        $inventory += $siteInventory
        
        if ([string]::IsNullOrEmpty($siteInventory.Error)) {
            Write-Log "[SUCCESS] $($site.Url) (Lists: $($siteInventory.Lists), Libraries: $($siteInventory.DocumentLibraries + $siteInventory.ImageLibraries + $siteInventory.AssetLibraries + $siteInventory.FormLibraries + $siteInventory.WikiPageLibraries), Pages: $($siteInventory.SitePages))" "Success"
        }
        else {
            Write-Log "[ERROR] $($site.Url) - $($siteInventory.Error)" "Error"
        }
    }
    
    # Export results
    Write-Log "Exporting inventory data to CSV..." "Info"
    $csvPath = Export-InventoryToCSV -InventoryData $inventory -OutputPath $OutputPath
    
    # Summary
    Write-Log "================================================" "Info"
    Write-Log "Inventory Complete" "Success"
    Write-Log "Total sites processed: $($inventory.Count)" "Success"
    Write-Log "Successful: $($inventory | Where-Object { [string]::IsNullOrEmpty($_.Error) } | Measure-Object | Select-Object -ExpandProperty Count)" "Success"
    Write-Log "Failed: $($inventory | Where-Object { -not [string]::IsNullOrEmpty($_.Error) } | Measure-Object | Select-Object -ExpandProperty Count)" "Success"
    Write-Log "Output file: $csvPath" "Info"
    Write-Log "Completed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" "Info"
    Write-Log "================================================" "Info"
}
catch {
    Write-Log "Script execution failed: $_" "Error"
    exit 1
}
finally {
    # Disconnect
    Disconnect-PnPOnline -ErrorAction SilentlyContinue | Out-Null
}
