# Create SharePoint Online site collections from CSV

Bulk-create site collections and configure **hub registration** / **hub association** using PnP.PowerShell.

## Files

| File | Role |
|------|------|
| `Create-SPOSiteCollectionsFromList.ps1` | Main script |
| `Sample-SiteCollections.csv` | Example input |
| `Logs/` | Created at runtime (gitignored) |

## Requirements

- PowerShell **7.4+** recommended (PnP.PowerShell **3.x**)
- SharePoint Admin (or app with equivalent permissions)
- Your own **Entra ID app** registration (`-ClientId`)

See the [root README](../README.md) for PnP install and auth guidance.

## Usage

```powershell
.\Create-SPOSiteCollectionsFromList.ps1 `
  -ListPath ".\Sample-SiteCollections.csv" `
  -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
  -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
```

### Optional app-only (certificate)

```powershell
.\Create-SPOSiteCollectionsFromList.ps1 `
  -ListPath ".\Sample-SiteCollections.csv" `
  -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
  -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
  -TenantId "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
  -CertificatePath ".\certs\pnpcert.pfx" `
  -CertificatePassword (Read-Host -AsSecureString)
```

## CSV columns

`SiteTitle`, `SiteUrl`, `IsHub`, `HubAssociation`, `Owner`, `Description`, optional `Template` (default `STS#3`).
