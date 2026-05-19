# SharePoint Online Site Collections & Hub Architecture Automation

Complete PowerShell automation solution for creating SharePoint Online Site Collections from list data and configuring hub site architecture.

## 📋 Table of Contents

- [Features](#features)
- [Prerequisites](#prerequisites)
- [Files](#files)
- [Installation](#installation)
- [Usage](#usage)
- [CSV Format](#csv-format)
- [Examples](#examples)
- [Troubleshooting](#troubleshooting)

## ✨ Features

- **Bulk Site Creation**: Create multiple SharePoint Online Site Collections from CSV data
- **Hub Site Registration**: Automatically register sites as hub sites based on `IsHub` flag
- **Hub Association**: Associate regular sites with hub sites based on list data
- **Automatic Module Installation**: Auto-installs PnP.PowerShell if not present
- **Comprehensive Logging**: All operations logged to timestamped log files in `./Logs` folder
- **Error Handling**: Robust error handling with detailed error messages
- **Two-Phase Processing**: 
  - Phase 1: Create all sites and register hub sites
  - Phase 2: Associate non-hub sites with their parent hubs
- **Validation**: URL format and required field validation
- **Duplicate Prevention**: Skips sites that already exist

## 📦 Prerequisites

- Windows PowerShell 5.1+ or PowerShell 7+
- PnP.PowerShell module (auto-installed if missing)
- SharePoint Tenant Admin URL access
- Appropriate permissions:
  - Global Admin or SharePoint Admin role
  - Ability to create site collections
  - Permission to manage hub sites

## 📁 Files

### 1. `Create-SPOSiteCollectionsFromList.ps1`
Main automation script with all functions for site creation and hub management.

**Functions Include:**
- `Connect-SPOTenant` - Connect to SharePoint Tenant Admin
- `Import-SiteCollectionData` - Read and validate CSV data
- `New-SPOSiteCollection` - Create individual site collections
- `Register-HubSite` - Register sites as hub sites
- `Add-SiteToHub` - Associate regular sites with hub sites
- `Test-SiteUrlFormat` - Validate site URL format
- `Write-Log` - Centralized logging
- `Invoke-SiteCollectionCreation` - Orchestrate entire process

### 2. `Sample-SiteCollections.csv`
Template CSV file with example data showing:
- 3 Hub Sites (Marketing, Sales, HR)
- 7 Associated regular sites
- Proper formatting and hierarchy

### 3. `README.md`
This comprehensive documentation file.

## 🚀 Installation

1. **Clone or download the repository**
   ```powershell
   git clone https://github.com/sachchinannam404/SharePointPowerShellAutomations.git
   cd SharePointPowerShellAutomations
   ```

2. **Verify PowerShell version**
   ```powershell
   $PSVersionTable.PSVersion
   ```
   Minimum required: 5.1

3. **Set execution policy** (if needed)
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

## 📊 Usage

### Basic Execution

```powershell
.\Create-SPOSiteCollectionsFromList.ps1 `
  -ListPath ".\Sample-SiteCollections.csv" `
  -TenantAdminUrl "https://contoso-admin.sharepoint.com"
```

### Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `ListPath` | String | Yes | Path to CSV file with site collection data |
| `TenantAdminUrl` | String | Yes | SharePoint Tenant Admin URL (format: https://tenant-admin.sharepoint.com) |
| `CsvHeaders` | Array | No | Custom CSV column headers (default: SiteTitle,SiteUrl,IsHub,HubAssociation,Owner,Description) |

### With Verbose Output

```powershell
.\Create-SPOSiteCollectionsFromList.ps1 `
  -ListPath "C:\SharePoint\SiteCollections.csv" `
  -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
  -Verbose
```

## 📋 CSV Format

### Required Columns

| Column | Type | Description | Example |
|--------|------|-------------|---------|
| `SiteTitle` | String | Display name for the site collection | "Marketing Hub" |
| `SiteUrl` | String | Full URL for the site collection | "https://contoso.sharepoint.com/sites/marketing-hub" |
| `IsHub` | Boolean | Whether to register as hub site (true/false) | "true" |
| `HubAssociation` | String | URL of parent hub site (leave empty for hub sites) | "https://contoso.sharepoint.com/sites/marketing-hub" |
| `Owner` | String | Email of site owner/admin | "marketing@contoso.com" |
| `Description` | String | Site description and purpose | "Central hub for marketing" |

### URL Format Rules

- Must use `https://`
- Format: `https://tenant.sharepoint.com/sites/sitename`
- Site names should be lowercase and use hyphens (not spaces)
- No special characters except hyphens

### Sample CSV Structure

```csv
SiteTitle,SiteUrl,IsHub,HubAssociation,Owner,Description
Marketing Hub,https://contoso.sharepoint.com/sites/marketing-hub,true,,marketing@contoso.com,Central hub for all marketing
US Marketing,https://contoso.sharepoint.com/sites/us-marketing,false,https://contoso.sharepoint.com/sites/marketing-hub,marketing@contoso.com,US Region Marketing
```

## 📝 Examples

### Example 1: Create Multiple Hub Sites with Associated Teams

**CSV Content:**
```csv
SiteTitle,SiteUrl,IsHub,HubAssociation,Owner,Description
Sales Hub,https://contoso.sharepoint.com/sites/sales-hub,true,,sales@contoso.com,Main sales hub
Enterprise Sales,https://contoso.sharepoint.com/sites/enterprise-sales,false,https://contoso.sharepoint.com/sites/sales-hub,sales@contoso.com,Enterprise team
SMB Sales,https://contoso.sharepoint.com/sites/smb-sales,false,https://contoso.sharepoint.com/sites/sales-hub,sales@contoso.com,SMB team
```

**Execution:**
```powershell
.\Create-SPOSiteCollectionsFromList.ps1 `
  -ListPath ".\sales-sites.csv" `
  -TenantAdminUrl "https://contoso-admin.sharepoint.com"
```

### Example 2: Create Organizational Hub Architecture

**Structure:**
```
Marketing Hub
├── US Marketing
└── EMEA Marketing

Sales Hub
├── Enterprise Sales
└── SMB Sales

HR Hub
├── Recruitment
└── Payroll
```

Use the provided `Sample-SiteCollections.csv` which includes this complete structure.

## 📊 Log Files

All operations are logged to timestamped files in the `./Logs` folder:

```
Logs/
├── SPO-SiteCreation-2026-05-19-143022.log
├── SPO-SiteCreation-2026-05-19-150045.log
└── SPO-SiteCreation-2026-05-19-152030.log
```

### Log File Format

```
[2026-05-19 14:30:22] [Info] Connecting to SharePoint Tenant: https://contoso-admin.sharepoint.com
[2026-05-19 14:30:25] [Success] Successfully connected to SharePoint Tenant Admin
[2026-05-19 14:30:26] [Info] Creating site collection: Marketing Hub (https://contoso.sharepoint.com/sites/marketing-hub)
[2026-05-19 14:30:45] [Success] Successfully created site collection: Marketing Hub
```

## 🔧 Troubleshooting

### Issue: "Invalid site URL format"
**Cause**: Site URL doesn't match expected format
**Solution**: Ensure URLs follow format: `https://tenant.sharepoint.com/sites/sitename`

### Issue: "Site already exists"
**Cause**: Site collection already created in tenant
**Solution**: Script will skip and continue. Remove site or use different URL

### Issue: "Hub site not found"
**Cause**: HubAssociation URL doesn't match a registered hub site
**Solution**: 
1. Verify hub site is created first
2. Ensure URL exactly matches hub site URL in CSV
3. Check IsHub is set to "true" for parent site

### Issue: "Failed to connect to SharePoint Tenant"
**Cause**: Invalid TenantAdminUrl or insufficient permissions
**Solution**:
1. Verify you have Global Admin or SharePoint Admin role
2. Check TenantAdminUrl format: `https://tenant-admin.sharepoint.com`
3. Ensure multi-factor authentication is configured

### Issue: "PnP.PowerShell module not found"
**Cause**: Module not installed
**Solution**: Script will auto-install, but if it fails:
```powershell
Install-Module -Name PnP.PowerShell -Force -Scope CurrentUser
```

### Issue: "Access Denied"
**Cause**: Insufficient permissions
**Solution**:
1. Verify you're a Global Admin or SharePoint Admin
2. Check if your account is blocked from creating sites
3. Verify site owner email is valid and active

## 📖 Hub Site Architecture Best Practices

1. **Create hub sites first** - Ensure all hub sites (IsHub=true) are created before associating regular sites
2. **Unique URLs** - Use descriptive, unique URLs for better navigation
3. **Clear ownership** - Assign clear owners for each site
4. **Naming conventions** - Use consistent naming patterns (e.g., lowercase, hyphens)
5. **One parent hub** - Each non-hub site should be associated with only one hub
6. **Test before bulk** - Run with sample data first to verify configuration

## 🔐 Security Considerations

- Store CSV files securely; they may contain email addresses
- Use app-only authentication for production scenarios
- Implement conditional access policies for tenant admin access
- Monitor site creation activities in audit logs
- Use service accounts with minimal required permissions

## 📞 Support & Contribution

For issues, questions, or contributions:
- Open an issue on GitHub
- Submit pull requests for improvements
- Check existing issues before creating new ones

## 📄 License

[Specify your license here]

## 🙏 Acknowledgments

Built with:
- PnP.PowerShell community
- SharePoint Online REST API
- PowerShell 7 best practices

---

**Last Updated**: 2026-05-19  
**Version**: 1.0.0  
**Author**: Sachchin Annam
