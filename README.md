# SharePoint PowerShell Automations

Personal collection of **SharePoint** and related **PowerShell** scripts: modern SharePoint Online automation, archived on-premises samples, and general Windows/ops utilities.

> **Scope:** This is a lab / archive toolkit, not a single product. Prefer the **Modern SPO** folders for current Microsoft 365 work. Treat on-prem folders as historical.

---

## Repository map

| Folder | Status | Purpose |
|--------|--------|--------|
| [**CreateSiteCollections/**](CreateSiteCollections/) | **Modern / maintained** | Bulk create SPO site collections and hub associations from CSV |
| [**Get_SPOInventory/**](Get_SPOInventory/) | **Modern / maintained** | Inventory site collections (lists, storage, users, …) to CSV |
| [**SPO/**](SPO/) | Mixed / archive | Older SharePoint Online–oriented scripts |
| [**SP2013/**](SP2013/) | **Archived** | SharePoint Server 2013 samples |
| [**SP2010/**](SP2010/) | **Archived** | SharePoint Server 2010 samples |
| [**MOSS2007/**](MOSS2007/) | **Archived** | MOSS 2007 / early 2010 OM-style scripts |
| [**SQLServer/**](SQLServer/) | Ops / archive | SQL Server–related PowerShell |
| [**Scripts/**](Scripts/) | General toolbox | Non-SharePoint utilities (FTP, Hyper-V, disk, MySQL, …) |

Each major folder has its own `README.md` with prerequisites and cautions.

---

## Modern SharePoint Online tools

### 1. Create site collections & hub architecture

**Path:** [`CreateSiteCollections/Create-SPOSiteCollectionsFromList.ps1`](CreateSiteCollections/Create-SPOSiteCollectionsFromList.ps1)  
**Sample data:** [`CreateSiteCollections/Sample-SiteCollections.csv`](CreateSiteCollections/Sample-SiteCollections.csv)

**Features**

- Bulk site creation from CSV  
- Hub registration (`IsHub`) and association (`HubAssociation`)  
- Optional **Entra ID app** auth (`-ClientId`, certificate)  
- Optional **Template** column (default `STS#3`)  
- Logging under `CreateSiteCollections/Logs/`  
- Two-phase processing: create/register hubs, then associate child sites  

**Quick start**

```powershell
cd CreateSiteCollections

# Interactive (browser) sign-in — register your own Entra app first (see below)
.\Create-SPOSiteCollectionsFromList.ps1 `
  -ListPath ".\Sample-SiteCollections.csv" `
  -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
  -ClientId "<your-entra-app-client-id>"
```

### 2. Site collection inventory

**Path:** [`Get_SPOInventory/Get-SPOSiteCollectionInventory.ps1`](Get_SPOInventory/Get-SPOSiteCollectionInventory.ps1)

**Docs (start here)**

- [Quick reference](Get_SPOInventory/QUICK-REFERENCE.md)  
- [Full guide](Get_SPOInventory/GET-SITECOLLECTION-INVENTORY-GUIDE.md)  
- [Examples](Get_SPOInventory/INVENTORY-EXAMPLES.md)  
- [Sample output CSV](Get_SPOInventory/Sample-SiteInventory-Output.csv)  

```powershell
cd Get_SPOInventory

.\Get-SPOSiteCollectionInventory.ps1 `
  -TenantAdminUrl "https://contoso-admin.sharepoint.com"
```

---

## Prerequisites (modern SPO scripts)

| Requirement | Notes |
|-------------|--------|
| **PowerShell** | **7.4.0+** recommended for current **PnP.PowerShell 3.x** |
| **Windows PowerShell 5.1** | Only with an older pinned PnP module (e.g. `1.12.0`) |
| **PnP.PowerShell** | Install from Gallery; prefer a known version in production |
| **Permissions** | SharePoint Administrator or Global Administrator (or app with equivalent app permissions) |
| **Entra ID app** | **Required** for PnP in most tenants — register your own app; the multi-tenant PnP Management Shell app was retired |

### PnP.PowerShell and Entra ID

1. Register an Entra ID application for PnP (see [PnP auth docs](https://pnp.github.io/powershell/articles/authentication.html) and [discussion #4249](https://github.com/pnp/powershell/discussions/4249)).  
2. Grant the least-privilege Microsoft Graph / SharePoint permissions your script needs; admin consent as required.  
3. For interactive use, pass **`-ClientId`** to `Connect-PnPOnline` (supported by the updated site-creation script).  
4. For unattended runs, use a **certificate** (app-only), not a client secret, with SharePoint.

```powershell
# Current major line (PowerShell 7.4+)
Install-Module PnP.PowerShell -Scope CurrentUser -Force

# Legacy Windows PowerShell 5.1 pin (example)
Install-Module PnP.PowerShell -RequiredVersion 1.12.0 -Scope CurrentUser -Force
```

Official module docs: https://pnp.github.io/powershell/

---

## CSV format (site creation)

| Column | Required | Description |
|--------|----------|-------------|
| `SiteTitle` | Yes | Display name |
| `SiteUrl` | Yes | `https://tenant.sharepoint.com/sites/name` |
| `IsHub` | Yes | `true` / `false` |
| `HubAssociation` | No | Parent hub URL for non-hub sites |
| `Owner` | Yes | Site owner UPN/email |
| `Description` | No | Applied to the web after creation when possible |
| `Template` | No | SPO template ID (default `STS#3`) |

See the sample CSV for a hub + associated sites example.

---

## Archived & general content

- **MOSS2007 / SP2010 / SP2013:** On-premises server object model or farm-era scripts. **Do not run against Microsoft 365** expecting modern APIs.  
- **SPO/:** Older Online-oriented samples; review before use; prefer the two modern folders above.  
- **SQLServer/** and **Scripts/:** Database and general Windows automation; not SharePoint-specific.

Read each folder’s README before executing anything.

---

## Security notes

- Prefer **app-only certificate** auth for production automation.  
- Do not commit real tenant URLs, client secrets, certificates, or production CSVs.  
- `.gitignore` excludes `Logs/`, `*.log`, and common local/report CSV patterns.  
- See [SECURITY.md](SECURITY.md).

---

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md). Issues and PRs that improve modern SPO scripts or documentation are welcome.

## License

[MIT](LICENSE) © Sachchin Annam

---

**Author:** Sachchin Annam  
**Repo version (docs):** 1.1.0  
