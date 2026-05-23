
# Site Collection Inventory Script

A comprehensive PowerShell automation script to audit and extract detailed inventory information from SharePoint Online site collections and export the results to CSV format.

## 📋 Overview

This script scans SharePoint Online site collections and captures detailed inventory metrics including:
- Lists and Libraries (all types)
- Site Pages count
- Subsites (Webs)
- Users and Groups
- Total items across all lists
- Storage usage
- Last modified date
- Site template and owner information

## ✨ Features

- **Complete Site Audit**: Scans all site collections or a specific site
- **Multiple Library Types**: Detects Document, Image, Asset, Form, and Wiki libraries
- **Detailed Metrics**: Captures comprehensive site statistics
- **CSV Export**: Results exported to timestamped CSV files
- **Automatic Module Installation**: Auto-installs PnP.PowerShell if missing
- **Comprehensive Logging**: All operations logged to timestamped log files
- **Error Handling**: Robust error handling with detailed error messages
- **Flexible Targeting**: Scan all sites or target a specific site collection
- **Storage Tracking**: Monitors storage usage per site

## 📊 CSV Output Format

The exported CSV contains the following columns:

| Column | Description | Example |
|--------|-------------|---------|
| `SiteUrl` | Full URL of the site collection | https://contoso.sharepoint.com/sites/marketing |
| `SiteTitle` | Display name of the site | Marketing Hub |
| `SiteTemplate` | SharePoint template type | STS, SITEPAGEPUBLISHING, etc. |
| `Owner` | Email of site owner | owner@contoso.com |
| `Status` | Site status | Active |
| `Lists` | Number of custom lists | 5 |
| `DocumentLibraries` | Number of document libraries | 3 |
| `ImageLibraries` | Number of image/picture libraries | 1 |
| `AssetLibraries` | Number of asset libraries | 0 |
| `FormLibraries` | Number of form libraries | 0 |
| `WikiPageLibraries` | Number of wiki page libraries | 0 |
| `SitePages` | Number of modern site pages | 12 |
| `Webs` | Number of subsites (including root) | 3 |
| `Users` | Number of users with access | 15 |
| `Groups` | Number of SharePoint groups | 8 |
| `TotalItems` | Total number of items across all lists/libraries | 250 |
| `StorageUsedMB` | Storage used in megabytes | 1024.50 |
| `LastModified` | Date/time of last modification | 2026-05-20 14:30:00 |
| `Error` | Any errors encountered during processing | (empty if successful) |

## 🚀 Usage

### Basic Execution - Scan All Sites

```powershell
.\Get-SPOSiteCollectionInventory.ps1 `
  -TenantAdminUrl "https://contoso-admin.sharepoint.com"
```

### Scan Specific Site

```powershell
.\Get-SPOSiteCollectionInventory.ps1 `
  -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
  -SiteUrl "https://contoso.sharepoint.com/sites/marketing"
```

### Specify Output Location

```powershell
.\Get-SPOSiteCollectionInventory.ps1 `
  -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
  -OutputPath "C:\Reports\InventoryReport.csv"
```

### Enable Verbose Logging

```powershell
.\Get-SPOSiteCollectionInventory.ps1 `
  -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
  -Verbose
```

## 📋 Parameters

| Parameter | Type | Required | Description | Default |
|-----------|------|----------|-------------|---------|
| `TenantAdminUrl` | String | Yes | SharePoint Tenant Admin URL | N/A |
| `SiteUrl` | String | No | Specific site URL to inventory (if not provided, all sites are scanned) | N/A |
| `OutputPath` | String | No | Path where CSV will be exported | `.\SiteCollectionInventory_<timestamp>.csv` |
| `Verbose` | Switch | No | Enable detailed logging | False |

## 🔧 Prerequisites

- Windows PowerShell 5.1+ or PowerShell 7+
- PnP.PowerShell module (auto-installed if missing)
- SharePoint Tenant Admin URL access
- Appropriate permissions:
  - Global Admin or SharePoint Admin role
  - Ability to read all site collections
  - Permission to access site properties and lists

## 📁 Log Files

All operations are logged to timestamped files in the `./Logs` folder:

```
Logs/
├── SPO-Inventory-2026-05-20.log
```

### Log File Format

```
[2026-05-20 14:30:22] [Info] Checking for PnP.PowerShell module...
[2026-05-20 14:30:25] [Success] PnP.PowerShell module found (Version: 2.3.0)
[2026-05-20 14:30:26] [Info] Connecting to SharePoint Tenant: https://contoso-admin.sharepoint.com
[2026-05-20 14:30:28] [Success] Successfully connected to SharePoint Tenant
[2026-05-20 14:30:29] [Info] Retrieving site collection list...
[2026-05-20 14:30:31] [Success] Found 15 site collections
[2026-05-20 14:30:32] [1/15] Processing site: https://contoso.sharepoint.com/sites/marketing
[2026-05-20 14:30:45] [Success] Successfully inventoried site: https://contoso.sharepoint.com/sites/marketing (Lists: 5, Libraries: 3, Pages: 12)
```

## 📊 Example Reports

### All Sites Inventory

The script can generate a comprehensive report of all site collections:

```csv
SiteUrl,SiteTitle,SiteTemplate,Owner,Status,Lists,DocumentLibraries,ImageLibraries,AssetLibraries,FormLibraries,WikiPageLibraries,SitePages,Webs,Users,Groups,TotalItems,StorageUsedMB,LastModified,Error
https://contoso.sharepoint.com/sites/marketing,Marketing Hub,STS,marketing@contoso.com,Active,5,3,1,0,0,0,12,2,15,8,250,1024.50,2026-05-20 14:30:00,
https://contoso.sharepoint.com/sites/sales,Sales Hub,STS,sales@contoso.com,Active,3,2,0,1,0,0,8,1,12,6,180,512.25,2026-05-20 14:25:00,
https://contoso.sharepoint.com/sites/hr,HR Hub,STS,hr@contoso.com,Active,4,2,1,0,1,0,10,3,18,10,320,768.75,2026-05-20 14:20:00,
```

## 🔍 What the Script Does

1. **Validates Environment**
   - Checks PowerShell version
   - Installs PnP.PowerShell if missing
   - Connects to SharePoint Tenant Admin

2. **Retrieves Site Collections**
   - Gets list of all sites or specific site
   - Prepares for inventory process

3. **Processes Each Site**
   - Connects to individual site
   - Retrieves site properties (title, owner, template)
   - Counts all types of lists and libraries
   - Counts site pages
   - Counts subsites
   - Retrieves user and group information
   - Calculates total items
   - Tracks storage usage
   - Records last modification date

4. **Exports Results**
   - Creates timestamped CSV file
   - Includes all metrics in standardized format
   - Logs completion status

## ⚠️ Troubleshooting

### Issue: "Invalid site URL format"
**Solution**: Ensure the site URL is in the correct format: `https://tenant.sharepoint.com/sites/sitename`

### Issue: "Access Denied"
**Cause**: Insufficient permissions
**Solution**:
1. Verify you're a Global Admin or SharePoint Admin
2. Check your account isn't blocked from accessing sites
3. Ensure multi-factor authentication is properly configured

### Issue: "Failed to connect to SharePoint Tenant"
**Cause**: Invalid TenantAdminUrl or connectivity issues
**Solution**:
1. Verify TenantAdminUrl format: `https://tenant-admin.sharepoint.com`
2. Check your internet connection
3. Ensure you can access SharePoint in your browser

### Issue: "PnP.PowerShell module not found"
**Cause**: Module installation failed
**Solution**:
```powershell
Install-Module -Name PnP.PowerShell -Force -Scope CurrentUser -AllowClobber
```

### Issue: Script Takes Too Long
**Cause**: Processing many large sites
**Solution**:
- Consider targeting specific sites with `-SiteUrl` parameter
- Run during off-peak hours
- Check network connectivity

## 💡 Tips & Best Practices

1. **Start with a Single Site**
   ```powershell
   .\Get-SPOSiteCollectionInventory.ps1 `
     -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
     -SiteUrl "https://contoso.sharepoint.com/sites/test"
   ```

2. **Save Reports with Meaningful Names**
   ```powershell
   .\Get-SPOSiteCollectionInventory.ps1 `
     -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
     -OutputPath "C:\Reports\Inventory_$(Get-Date -Format 'yyyy-MM-dd').csv"
   ```

3. **Schedule Regular Audits**
   - Use Windows Task Scheduler to run weekly/monthly
   - Compare reports over time to track growth

4. **Combine with Other Tools**
   - Import CSV into Excel for analysis
   - Use Power BI to visualize inventory data
   - Create dashboards showing storage trends

## 🔐 Security Considerations

- Store CSV reports securely; they may contain email addresses
- Limit access to inventory reports
- Use app-only authentication for scheduled runs (PnP app registration)
- Run during business hours for transparency
- Archive old reports for compliance

## 📝 Example Scenarios

### Scenario 1: Monthly Site Audit
```powershell
# Run on the first Monday of each month
$outputFile = "C:\Reports\SiteInventory_$(Get-Date -Format 'yyyy-MM').csv"
.\Get-SPOSiteCollectionInventory.ps1 `
  -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
  -OutputPath $outputFile
```

### Scenario 2: Audit Specific Department Sites
```powershell
# Audit all marketing sites
$sites = @(
  "https://contoso.sharepoint.com/sites/marketing",
  "https://contoso.sharepoint.com/sites/marketing-emea",
  "https://contoso.sharepoint.com/sites/marketing-apac"
)

foreach ($site in $sites) {
  .\Get-SPOSiteCollectionInventory.ps1 `
    -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
    -SiteUrl $site
}
```

### Scenario 3: Storage Analysis
```powershell
# Generate inventory and analyze storage
$outputFile = "C:\Reports\StorageAnalysis.csv"
.\Get-SPOSiteCollectionInventory.ps1 `
  -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
  -OutputPath $outputFile

# Analyze using PowerShell
$report = Import-Csv -Path $outputFile
$report | Sort-Object -Property StorageUsedMB -Descending | 
  Select-Object SiteUrl, SiteTitle, StorageUsedMB | 
  Format-Table -AutoSize
```

## 🔗 Related Scripts

- `Create-SPOSiteCollectionsFromList.ps1` - Bulk create site collections
- `Set-SPOSitePermissions.ps1` - Manage site permissions
- `Archive-SPOSites.ps1` - Archive old sites

## 📞 Support & Feedback

For issues, suggestions, or contributions:
- Open an issue on GitHub
- Submit pull requests for improvements
- Check existing documentation

## 📄 Version History

### v1.0.0 (2026-05-20)
- Initial release
- Complete site inventory functionality
- CSV export capability
- Comprehensive logging
- Error handling and validation

---

**Last Updated**: 2026-05-20  
**Author**: SharePoint Automation Team  
**Compatibility**: PowerShell 5.1+, PowerShell 7+
