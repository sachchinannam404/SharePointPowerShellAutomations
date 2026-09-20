<#
.SYNOPSIS
    Creates SharePoint Online Site Collections based on list data and configures hub site architecture.

.DESCRIPTION
    Automates creation of SharePoint Online site collections from a CSV file.
    Supports hub registration and association. Uses PnP.PowerShell with your own Entra ID app.

.PARAMETER ListPath
    Path to the CSV file containing site collection details.
    Columns: SiteTitle, SiteUrl, IsHub, HubAssociation, Owner, Description, Template (optional)

.PARAMETER TenantAdminUrl
    SharePoint Tenant Admin URL (e.g., https://contoso-admin.sharepoint.com)

.PARAMETER ClientId
    Entra ID application (client) ID used with Connect-PnPOnline. Required for current PnP.PowerShell.

.PARAMETER TenantId
    Directory (tenant) ID. Required for certificate-based app-only auth.

.PARAMETER CertificatePath
    Path to a .pfx certificate for app-only authentication. If omitted, interactive login is used.

.PARAMETER CertificatePassword
    SecureString password for the certificate file (if the PFX is protected).

.PARAMETER DefaultTemplate
    Fallback site template when CSV Template is empty. Default: STS#3

.PARAMETER ProvisioningTimeoutSeconds
    Max seconds to wait for a new site to become available. Default: 300

.PARAMETER CsvHeaders
    Optional override of expected header names (advanced).

.EXAMPLE
    .\Create-SPOSiteCollectionsFromList.ps1 -ListPath ".\Sample-SiteCollections.csv" `
      -TenantAdminUrl "https://contoso-admin.sharepoint.com" -ClientId "00000000-0000-0000-0000-000000000000"

.NOTES
    Prerequisites:
    - PowerShell 7.4+ recommended for PnP.PowerShell 3.x
    - Entra ID app registration with appropriate SharePoint permissions
    - SharePoint Admin or equivalent app permissions
#>

param(
    [Parameter(Mandatory = $true, HelpMessage = "Path to CSV file with site collection data")]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$ListPath,

    [Parameter(Mandatory = $true, HelpMessage = "SharePoint Tenant Admin URL")]
    [ValidateScript({ $_ -match '^https:\/\/.+-admin\.sharepoint\.com$' })]
    [string]$TenantAdminUrl,

    [Parameter(Mandatory = $true, HelpMessage = "Entra ID application (client) ID")]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [ValidateScript({ if ($_) { Test-Path $_ -PathType Leaf } else { $true } })]
    [string]$CertificatePath,

    [Parameter(Mandatory = $false)]
    [SecureString]$CertificatePassword,

    [Parameter(Mandatory = $false)]
    [string]$DefaultTemplate = "STS#3",

    [Parameter(Mandatory = $false)]
    [ValidateRange(30, 1800)]
    [int]$ProvisioningTimeoutSeconds = 300,

    [Parameter(Mandatory = $false)]
    [array]$CsvHeaders = @("SiteTitle", "SiteUrl", "IsHub", "HubAssociation", "Owner", "Description", "Template")
)

$ErrorActionPreference = "Stop"
$VerbosePreference = "Continue"

$LogFolder = Join-Path -Path $PSScriptRoot -ChildPath "Logs"
$LogFile = Join-Path -Path $LogFolder -ChildPath ("SPO-SiteCreation-{0}.log" -f (Get-Date -Format "yyyy-MM-dd-HHmmss"))

if (-not (Test-Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
}

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

function Connect-SPOTenant {
    param(
        [string]$TenantAdminUrl,
        [string]$ClientId,
        [string]$TenantId,
        [string]$CertificatePath,
        [SecureString]$CertificatePassword
    )

    try {
        Write-Log "Connecting to SharePoint Tenant: $TenantAdminUrl" "Info"

        $currentConnection = Get-PnPConnection -ErrorAction SilentlyContinue
        if ($currentConnection -and $currentConnection.Url -like "$TenantAdminUrl*") {
            Write-Log "Already connected to SharePoint Online ($($currentConnection.Url))" "Info"
            return $true
        }

        if ($CertificatePath) {
            if ([string]::IsNullOrWhiteSpace($TenantId)) {
                throw "TenantId is required when using CertificatePath for app-only authentication."
            }

            Write-Log "Using certificate-based app-only authentication" "Info"
            $connectParams = @{
                Url            = $TenantAdminUrl
                ClientId       = $ClientId
                Tenant         = $TenantId
                CertificatePath = $CertificatePath
            }
            if ($CertificatePassword) {
                $connectParams["CertificatePassword"] = $CertificatePassword
            }
            Connect-PnPOnline @connectParams
        }
        else {
            Write-Log "Using interactive authentication with ClientId $ClientId" "Info"
            Connect-PnPOnline -Url $TenantAdminUrl -ClientId $ClientId -Interactive
        }

        Write-Log "Successfully connected to SharePoint Tenant Admin" "Success"
        return $true
    }
    catch {
        Write-Log "Failed to connect to SharePoint Tenant: $($_.Exception.Message)" "Error"
        throw
    }
}

function Import-SiteCollectionData {
    param([string]$ListPath)

    try {
        Write-Log "Reading CSV file from: $ListPath" "Info"

        $csvData = [System.Collections.Generic.List[object]]::new()
        $rowNumber = 0

        Import-Csv -Path $ListPath | ForEach-Object {
            $rowNumber++

            if ([string]::IsNullOrWhiteSpace($_.SiteTitle)) {
                Write-Log "Row $rowNumber: SiteTitle is empty - skipping" "Warning"
                return
            }

            if ([string]::IsNullOrWhiteSpace($_.SiteUrl)) {
                Write-Log "Row $rowNumber: SiteUrl is empty - skipping" "Warning"
                return
            }

            $csvData.Add($_)
        }

        Write-Log "Successfully imported $($csvData.Count) site collection row(s)" "Success"
        return $csvData
    }
    catch {
        Write-Log "Error reading CSV file: $($_.Exception.Message)" "Error"
        throw
    }
}

function Test-TenantSiteExists {
    param([string]$SiteUrl)

    try {
        $site = Get-PnPTenantSite -Url $SiteUrl -ErrorAction Stop
        return ($null -ne $site)
    }
    catch {
        return $false
    }
}

function Wait-ForTenantSite {
    param(
        [string]$SiteUrl,
        [int]$TimeoutSeconds = 300,
        [int]$PollSeconds = 15
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    Write-Log "Waiting for site provisioning (timeout ${TimeoutSeconds}s): $SiteUrl" "Info"

    while ((Get-Date) -lt $deadline) {
        if (Test-TenantSiteExists -SiteUrl $SiteUrl) {
            try {
                $site = Get-PnPTenantSite -Url $SiteUrl -ErrorAction Stop
                # Status can vary; presence + non-failed is enough for association steps
                if ($site.Status -and $site.Status -match "Failed|Recycled") {
                    Write-Log "Site reported status '$($site.Status)'" "Warning"
                }
                Write-Log "Site is available: $SiteUrl" "Success"
                return $true
            }
            catch {
                # keep polling
            }
        }
        Start-Sleep -Seconds $PollSeconds
    }

    Write-Log "Timed out waiting for site: $SiteUrl" "Warning"
    return $false
}

function New-SPOSiteCollection {
    param(
        [string]$SiteTitle,
        [string]$SiteUrl,
        [string]$Owner,
        [string]$Description,
        [string]$Template,
        [int]$ProvisioningTimeoutSeconds
    )

    try {
        Write-Log "Creating site collection: $SiteTitle ($SiteUrl) [Template=$Template]" "Info"

        if (Test-TenantSiteExists -SiteUrl $SiteUrl) {
            Write-Log "Site already exists: $SiteUrl - skipping creation" "Warning"
            return $SiteUrl
        }

        $newParams = @{
            Title    = $SiteTitle
            Url      = $SiteUrl
            Owner    = $Owner
            Template = $Template
            Wait     = $true
        }

        # Wait switch availability differs by PnP version; fall back if unsupported
        try {
            New-PnPTenantSite @newParams | Out-Null
        }
        catch {
            if ($_.Exception.Message -match "Wait|parameter") {
                $newParams.Remove("Wait")
                New-PnPTenantSite @newParams | Out-Null
                Wait-ForTenantSite -SiteUrl $SiteUrl -TimeoutSeconds $ProvisioningTimeoutSeconds | Out-Null
            }
            else {
                throw
            }
        }

        if (-not (Test-TenantSiteExists -SiteUrl $SiteUrl)) {
            Wait-ForTenantSite -SiteUrl $SiteUrl -TimeoutSeconds $ProvisioningTimeoutSeconds | Out-Null
        }

        Write-Log "Successfully created site collection: $SiteTitle" "Success"

        if (-not [string]::IsNullOrWhiteSpace($Description)) {
            try {
                Connect-PnPOnline -Url $SiteUrl -ClientId $script:ClientId -Interactive -ErrorAction Stop
                Set-PnPWeb -Description $Description -ErrorAction Stop
                Connect-SPOTenant -TenantAdminUrl $script:TenantAdminUrl -ClientId $script:ClientId `
                    -TenantId $script:TenantId -CertificatePath $script:CertificatePath `
                    -CertificatePassword $script:CertificatePassword | Out-Null
                Write-Log "Applied web description for $SiteUrl" "Info"
            }
            catch {
                Write-Log "Could not set Description on web (site still created): $($_.Exception.Message)" "Warning"
                # Reconnect to admin for subsequent operations
                try {
                    Connect-SPOTenant -TenantAdminUrl $script:TenantAdminUrl -ClientId $script:ClientId `
                        -TenantId $script:TenantId -CertificatePath $script:CertificatePath `
                        -CertificatePassword $script:CertificatePassword | Out-Null
                }
                catch {
                    Write-Log "Failed to restore admin connection: $($_.Exception.Message)" "Error"
                    throw
                }
            }
        }

        return $SiteUrl
    }
    catch {
        Write-Log "Failed to create site collection $SiteTitle : $($_.Exception.Message)" "Error"
        throw
    }
}

function Register-HubSite {
    param(
        [string]$SiteUrl,
        [string]$HubTitle
    )

    try {
        Write-Log "Registering hub site: $SiteUrl (Title: $HubTitle)" "Info"

        $hubSite = Get-PnPHubSite -ErrorAction SilentlyContinue | Where-Object { $_.SiteUrl -eq $SiteUrl }

        if ($hubSite) {
            Write-Log "Site is already registered as hub site: $SiteUrl" "Warning"
            return $hubSite.Id
        }

        $hubId = Register-PnPHubSite -Site $SiteUrl -HubSiteTitle $HubTitle

        Write-Log "Successfully registered hub site: $SiteUrl (Hub ID: $hubId)" "Success"
        return $hubId
    }
    catch {
        Write-Log "Failed to register hub site $SiteUrl : $($_.Exception.Message)" "Error"
        throw
    }
}

function Add-SiteToHub {
    param(
        [string]$SiteUrl,
        [string]$HubUrl
    )

    try {
        Write-Log "Associating site $SiteUrl with hub $HubUrl" "Info"

        $hubSite = Get-PnPHubSite | Where-Object { $_.SiteUrl -eq $HubUrl }

        if (-not $hubSite) {
            Write-Log "Hub site not found: $HubUrl" "Error"
            throw "Hub site not found: $HubUrl"
        }

        Add-PnPHubSiteAssociation -Site $SiteUrl -HubSite $hubSite.Id

        Write-Log "Successfully associated site with hub" "Success"
    }
    catch {
        Write-Log "Failed to associate site with hub: $($_.Exception.Message)" "Error"
        throw
    }
}

function Test-SiteUrlFormat {
    param([string]$SiteUrl)

    $urlPattern = '^https:\/\/[a-z0-9-]+\.sharepoint\.com\/sites\/[a-z0-9-]+$'

    if ($SiteUrl -match $urlPattern) {
        return $true
    }

    Write-Log "Invalid site URL format: $SiteUrl. Expected: https://tenant.sharepoint.com/sites/sitename" "Error"
    return $false
}

function Invoke-SiteCollectionCreation {
    param(
        [System.Collections.IEnumerable]$SiteData,
        [string]$DefaultTemplate,
        [int]$ProvisioningTimeoutSeconds
    )

    $successCount = 0
    $failureCount = 0
    $hubSites = @{}

    Write-Log "======================================================" "Info"
    Write-Log "Starting Site Collection Creation Process" "Info"
    Write-Log "======================================================" "Info"

    foreach ($site in $SiteData) {
        try {
            $siteTitle = $site.SiteTitle.Trim()
            $siteUrl = $site.SiteUrl.Trim()
            $isHub = [bool]::Parse(($site.IsHub.ToString().Trim()))
            $owner = $site.Owner.Trim()
            $description = if ($site.Description) { $site.Description.Trim() } else { "" }
            $template = if ($site.PSObject.Properties.Name -contains "Template" -and -not [string]::IsNullOrWhiteSpace($site.Template)) {
                $site.Template.Trim()
            }
            else {
                $DefaultTemplate
            }

            if (-not (Test-SiteUrlFormat -SiteUrl $siteUrl)) {
                $failureCount++
                continue
            }

            $createdSiteUrl = New-SPOSiteCollection -SiteTitle $siteTitle -SiteUrl $siteUrl -Owner $owner `
                -Description $description -Template $template -ProvisioningTimeoutSeconds $ProvisioningTimeoutSeconds

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

    Write-Log "======================================================" "Info"
    Write-Log "Starting Hub Site Association Process" "Info"
    Write-Log "======================================================" "Info"

    foreach ($site in $SiteData) {
        try {
            $siteUrl = $site.SiteUrl.Trim()
            $isHub = [bool]::Parse(($site.IsHub.ToString().Trim()))
            $hubAssociation = if ($site.HubAssociation) { $site.HubAssociation.Trim() } else { "" }

            if ($isHub -or [string]::IsNullOrWhiteSpace($hubAssociation)) {
                continue
            }

            if (-not (Test-SiteUrlFormat -SiteUrl $siteUrl)) {
                continue
            }

            Add-SiteToHub -SiteUrl $siteUrl -HubUrl $hubAssociation
        }
        catch {
            Write-Log "Exception associating site with hub: $($_.Exception.Message)" "Error"
        }
    }

    Write-Log "======================================================" "Info"
    Write-Log "Site Collection Creation Summary" "Info"
    Write-Log "======================================================" "Info"
    Write-Log "Total Rows: $(($SiteData | Measure-Object).Count)" "Info"
    Write-Log "Successful: $successCount" "Success"
    Write-Log "Failed: $failureCount" "Error"
    Write-Log "Hub Sites Registered (this run): $($hubSites.Count)" "Info"
}

# Script-scoped connection settings for reconnect after site-scoped operations
$script:ClientId = $ClientId
$script:TenantAdminUrl = $TenantAdminUrl
$script:TenantId = $TenantId
$script:CertificatePath = $CertificatePath
$script:CertificatePassword = $CertificatePassword

try {
    Write-Log "Script execution started" "Info"

    Write-Log "Loading PnP.PowerShell module..." "Info"
    if (-not (Get-Module -ListAvailable -Name "PnP.PowerShell")) {
        Write-Log "Installing PnP.PowerShell module (CurrentUser scope)..." "Info"
        Install-Module -Name "PnP.PowerShell" -Force -Scope CurrentUser
    }

    Import-Module -Name "PnP.PowerShell" -Force

    $pnpMod = Get-Module -Name "PnP.PowerShell"
    if ($pnpMod) {
        Write-Log "Using PnP.PowerShell version $($pnpMod.Version)" "Info"
    }

    Connect-SPOTenant -TenantAdminUrl $TenantAdminUrl -ClientId $ClientId -TenantId $TenantId `
        -CertificatePath $CertificatePath -CertificatePassword $CertificatePassword

    $siteCollections = Import-SiteCollectionData -ListPath $ListPath

    if ($siteCollections.Count -eq 0) {
        Write-Log "No valid site collections found in CSV file" "Warning"
        exit 0
    }

    Invoke-SiteCollectionCreation -SiteData $siteCollections -DefaultTemplate $DefaultTemplate `
        -ProvisioningTimeoutSeconds $ProvisioningTimeoutSeconds

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
