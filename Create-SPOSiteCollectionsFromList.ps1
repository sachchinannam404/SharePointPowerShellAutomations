<#
.SYNOPSIS
    Creates SharePoint Online Site Collections based on list data and configures hub site architecture.

.DESCRIPTION
    This script automates the creation of SharePoint Online Site Collections from a CSV/Excel list.
    It supports creating site collections and optionally registering them as hub sites or associating
    them with existing hub sites based on list data points.

.PARAMETER ListPath
    Path to the CSV file containing site collection details.
    Required columns: SiteTitle, SiteUrl, IsHub, HubAssociation, Owner, Description

.PARAMETER TenantAdminUrl
    The SharePoint Tenant Admin URL (e.g., https://contoso-admin.sharepoint.com)

.PARAMETER CsvHeaders
    Custom CSV headers if different from default. Default: SiteTitle,SiteUrl,IsHub,HubAssociation,Owner,Description

.EXAMPLE
    .\Create-SPOSiteCollectionsFromList.ps1 -ListPath "C:\SiteCollections.csv" -TenantAdminUrl "https://contoso-admin.sharepoint.com"

.NOTES
    Prerequisites:
    - PnP.PowerShell module installed
    - Appropriate permissions to create site collections and manage hub sites
    - CSV file with required columns
#>

param(
    [Parameter(Mandatory = $true, HelpMessage = "Path to CSV file with site collection data")]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$ListPath,

    [Parameter(Mandatory = $true, HelpMessage = "SharePoint Tenant Admin URL")]
    [ValidateScript({ $_ -match '^https:\/\/.+-admin\.sharepoint\.com$' })]
    [string]$TenantAdminUrl,

    [Parameter(Mandatory = $false)]
    [array]$CsvHeaders = @("SiteTitle", "SiteUrl", "IsHub", "HubAssociation", "Owner", "Description")
)

# ============================================================================
# Configuration
# ============================================================================

$ErrorActionPreference = "Stop"
$VerbosePreference = "Continue"

# Log file configuration
$LogFolder = "$PSScriptRoot\Logs"
$LogFile = "$LogFolder\SPO-SiteCreation-$(Get-Date -Format 'yyyy-MM-dd-HHmmss').log"

if (-not (Test-Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
}

# ============================================================================
# Functions
# ============================================================================

<#
.SYNOPSIS
    Write log messages to both console and log file
#>
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("Info", "Warning", "Error", "Success")]
        [string]$Level = "Info"
    )
    
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [$Level] $Message"
    
    switch ($Level) {
        "Info" { Write-Host $LogMessage -ForegroundColor Cyan }
        "Warning" { Write-Host $LogMessage -ForegroundColor Yellow }
        "Error" { Write-Host $LogMessage -ForegroundColor Red }
        "Success" { Write-Host $LogMessage -ForegroundColor Green }
    }
    
    Add-Content -Path $LogFile -Value $LogMessage
}

<#
.SYNOPSIS
    Connect to SharePoint Online Tenant Admin
#>
function Connect-SPOTenant {
    param(
        [string]$TenantAdminUrl
    )
    
    try {
        Write-Log "Connecting to SharePoint Tenant: $TenantAdminUrl" "Info"
        
        # Check if already connected
        $currentConnection = Get-PnPConnection -ErrorAction SilentlyContinue
        
        if ($currentConnection) {
            Write-Log "Already connected to SharePoint Online" "Info"
            return $true
        }
        
        # Connect to tenant admin
        Connect-PnPOnline -Url $TenantAdminUrl -Interactive
        
        Write-Log "Successfully connected to SharePoint Tenant Admin" "Success"
        return $true
    }
    catch {
        Write-Log "Failed to connect to SharePoint Tenant: $($_.Exception.Message)" "Error"
        throw
    }
}

<#
.SYNOPSIS
    Read and validate CSV data
#>
function Import-SiteCollectionData {
    param(
        [string]$ListPath
    )
    
    try {
        Write-Log "Reading CSV file from: $ListPath" "Info"
        
        $csvData = @()
        $rowNumber = 0
        
        Get-Content $ListPath | ConvertFrom-Csv | ForEach-Object {
            $rowNumber++
            
            # Validate required fields
            if ([string]::IsNullOrWhiteSpace($_.SiteTitle)) {
                Write-Log "Row $rowNumber: SiteTitle is empty - skipping" "Warning"
                return
            }
            
            if ([string]::IsNullOrWhiteSpace($_.SiteUrl)) {
                Write-Log "Row $rowNumber: SiteUrl is empty - skipping" "Warning"
                return
            }
            
            $csvData += $_
        }
        
        Write-Log "Successfully imported $($csvData.Count) site collections" "Success"
        return $csvData
    }
    catch {
        Write-Log "Error reading CSV file: $($_.Exception.Message)" "Error"
        throw
    }
}

<#
.SYNOPSIS
    Create SharePoint Online Site Collection
#>
function New-SPOSiteCollection {
    param(
        [string]$SiteTitle,
        [string]$SiteUrl,
        [string]$Owner,
        [string]$Description
    )
    
    try {
        Write-Log "Creating site collection: $SiteTitle ($SiteUrl)" "Info"
        
        # Verify site doesn't already exist
        try {
            Get-PnPTenantSite -Url $SiteUrl -ErrorAction SilentlyContinue | Out-Null
            Write-Log "Site already exists: $SiteUrl - skipping creation" "Warning"
            return $SiteUrl
        }
        catch {
            # Site doesn't exist, proceed with creation
        }
        
        # Create new site
        $newSite = New-PnPTenantSite -Title $SiteTitle -Url $SiteUrl -Owner $Owner -Template "STS#3"
        
        Write-Log "Successfully created site collection: $SiteTitle" "Success"
        
        # Wait for site to be provisioned
        Write-Log "Waiting for site to be provisioned..." "Info"
        Start-Sleep -Seconds 30
        
        return $SiteUrl
    }
    catch {
        Write-Log "Failed to create site collection $SiteTitle : $($_.Exception.Message)" "Error"
        throw
    }
}

<#
.SYNOPSIS
    Register a site as a hub site
#>
function Register-HubSite {
    param(
        [string]$SiteUrl,
        [string]$HubTitle
    )
    
    try {
        Write-Log "Registering hub site: $SiteUrl (Title: $HubTitle)" "Info"
        
        # Check if already a hub site
        $hubSite = Get-PnPHubSite -ErrorAction SilentlyContinue | Where-Object { $_.SiteUrl -eq $SiteUrl }
        
        if ($hubSite) {
            Write-Log "Site is already registered as hub site: $SiteUrl" "Warning"
            return $hubSite.ID
        }
        
        # Register as hub site
        $hubId = Register-PnPHubSite -Site $SiteUrl -HubSiteTitle $HubTitle
        
        Write-Log "Successfully registered hub site: $SiteUrl (Hub ID: $hubId)" "Success"
        return $hubId
    }
    catch {
        Write-Log "Failed to register hub site $SiteUrl : $($_.Exception.Message)" "Error"
        throw
    }
}

<#
.SYNOPSIS
    Associate a site with a hub site
#>
function Add-SiteToHub {
    param(
        [string]$SiteUrl,
        [string]$HubUrl
    )
    
    try {
        Write-Log "Associating site $SiteUrl with hub $HubUrl" "Info"
        
        # Get hub site ID
        $hubSite = Get-PnPHubSite | Where-Object { $_.SiteUrl -eq $HubUrl }
        
        if (-not $hubSite) {
            Write-Log "Hub site not found: $HubUrl" "Error"
            throw "Hub site not found: $HubUrl"
        }
        
        # Associate the site with hub
        Add-PnPHubSiteAssociation -Site $SiteUrl -HubSite $hubSite.ID
        
        Write-Log "Successfully associated site with hub" "Success"
    }
    catch {
        Write-Log "Failed to associate site with hub: $($_.Exception.Message)" "Error"
        throw
    }
}

<#
.SYNOPSIS
    Validate site URL format
#>
function Test-SiteUrlFormat {
    param(
        [string]$SiteUrl
    )
    
    # Expected format: https://tenant.sharepoint.com/sites/sitename
    $urlPattern = '^https:\/\/[a-z0-9-]+\.sharepoint\.com\/sites\/[a-z0-9-]+$'
    
    if ($SiteUrl -match $urlPattern) {
        return $true
    }
    
    Write-Log "Invalid site URL format: $SiteUrl. Expected format: https://tenant.sharepoint.com/sites/sitename" "Error"
    return $false
}

<#
.SYNOPSIS
    Process all site collections from the list
#>
function Invoke-SiteCollectionCreation {
    param(
        [array]$SiteData
    )
    
    $successCount = 0
    $failureCount = 0
    $hubSites = @{}
    
    Write-Log "======================================================" "Info"
    Write-Log "Starting Site Collection Creation Process" "Info"
    Write-Log "======================================================" "Info"
    
    # First pass: Create all sites and register hub sites
    foreach ($site in $SiteData) {
        try {
            $siteTitle = $site.SiteTitle.Trim()
            $siteUrl = $site.SiteUrl.Trim()
            $isHub = [bool]::Parse($site.IsHub)
            $owner = $site.Owner.Trim()
            $description = $site.Description.Trim()
            
            # Validate URL format
            if (-not (Test-SiteUrlFormat -SiteUrl $siteUrl)) {
                $failureCount++
                continue
            }
            
            # Create site collection
            $createdSiteUrl = New-SPOSiteCollection -SiteTitle $siteTitle -SiteUrl $siteUrl -Owner $owner -Description $description
            
            # Register as hub site if required
            if ($isHub) {
                $hubId = Register-HubSite -SiteUrl $createdSiteUrl -HubTitle $siteTitle
                $hubSites[$createdSiteUrl] = @{
                    Title = $siteTitle
                    Id    = $hubId
                }
            }
            
            $successCount++
        }
        catch {
            $failureCount++
            Write-Log "Exception processing site: $($_.Exception.Message)" "Error"
        }
    }
    
    # Second pass: Associate non-hub sites with their hub sites
    Write-Log "======================================================" "Info"
    Write-Log "Starting Hub Site Association Process" "Info"
    Write-Log "======================================================" "Info"
    
    foreach ($site in $SiteData) {
        try {
            $siteUrl = $site.SiteUrl.Trim()
            $isHub = [bool]::Parse($site.IsHub)
            $hubAssociation = $site.HubAssociation.Trim()
            
            # Skip if this is a hub site or no hub association specified
            if ($isHub -or [string]::IsNullOrWhiteSpace($hubAssociation)) {
                continue
            }
            
            # Validate URL format
            if (-not (Test-SiteUrlFormat -SiteUrl $siteUrl)) {
                continue
            }
            
            # Associate with hub site
            Add-SiteToHub -SiteUrl $siteUrl -HubUrl $hubAssociation
        }
        catch {
            Write-Log "Exception associating site with hub: $($_.Exception.Message)" "Error"
        }
    }
    
    # Summary
    Write-Log "======================================================" "Info"
    Write-Log "Site Collection Creation Summary" "Info"
    Write-Log "======================================================" "Info"
    Write-Log "Total Sites Processed: $($SiteData.Count)" "Info"
    Write-Log "Successful: $successCount" "Success"
    Write-Log "Failed: $failureCount" "Error"
    Write-Log "Hub Sites Created: $($hubSites.Count)" "Info"
}

# ============================================================================
# Main Execution
# ============================================================================

try {
    Write-Log "Script execution started" "Info"
    
    # Import PnP.PowerShell module
    Write-Log "Loading PnP.PowerShell module..." "Info"
    if (-not (Get-Module -ListAvailable -Name "PnP.PowerShell")) {
        Write-Log "Installing PnP.PowerShell module..." "Info"
        Install-Module -Name "PnP.PowerShell" -Force -Scope CurrentUser
    }
    
    Import-Module -Name "PnP.PowerShell" -Force
    
    # Connect to SharePoint
    Connect-SPOTenant -TenantAdminUrl $TenantAdminUrl
    
    # Import site data
    $siteCollections = Import-SiteCollectionData -ListPath $ListPath
    
    if ($siteCollections.Count -eq 0) {
        Write-Log "No valid site collections found in CSV file" "Warning"
        exit 0
    }
    
    # Process site collections
    Invoke-SiteCollectionCreation -SiteData $siteCollections
    
    Write-Log "Script execution completed successfully" "Success"
}
catch {
    Write-Log "Critical error occurred: $($_.Exception.Message)" "Error"
    Write-Log "Stack trace: $($_.ScriptStackTrace)" "Error"
    exit 1
}
finally {
    Write-Log "Log file saved to: $LogFile" "Info"
}
