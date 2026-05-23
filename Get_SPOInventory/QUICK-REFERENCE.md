# Get-SPOSiteCollectionInventory - Quick Reference

## ⚡ 30-Second Start

### Scan All Sites
```powershell
.\Get-SPOSiteCollectionInventory.ps1 -TenantAdminUrl "https://contoso-admin.sharepoint.com"
```

### Scan One Site
```powershell
.\Get-SPOSiteCollectionInventory.ps1 -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
  -SiteUrl "https://contoso.sharepoint.com/sites/marketing"
```

---

## 📊 CSV Columns Explained

| Column | What It Means |
|--------|---------------|
| **SiteUrl** | Full SharePoint site URL |
| **SiteTitle** | Display name shown in UI |
| **Lists** | Count of custom lists |
| **DocumentLibraries** | Count of document libraries |
| **SitePages** | Count of modern pages |
| **Webs** | Count of subsites + root |
| **Users** | People with any access |
| **StorageUsedMB** | Storage used in megabytes |
| **TotalItems** | All items across all lists |

---

## 🎯 Common Tasks

### Find Top 10 Largest Sites
```powershell
$report = Import-Csv "SiteCollectionInventory_*.csv"
$report | Sort-Object { [decimal]$_.StorageUsedMB } -Descending | Select-Object -First 10
```

### Count Total Sites Scanned
```powershell
$report = Import-Csv "SiteCollectionInventory_*.csv"
$report.Count
```

### Get Sites with Errors
```powershell
$report = Import-Csv "SiteCollectionInventory_*.csv"
$report | Where-Object { $_.Error -ne "" }
```

### Total Storage Used
```powershell
$report = Import-Csv "SiteCollectionInventory_*.csv"
$report | Measure-Object -Property StorageUsedMB -Sum
```

### Find Sites with Most Lists
```powershell
$report = Import-Csv "SiteCollectionInventory_*.csv"
$report | Sort-Object Lists -Descending | Select-Object -First 5
```

---

## 🛠️ Troubleshooting

### "Access Denied"
- Make sure you're a SharePoint Admin or Global Admin
- Check multi-factor authentication (MFA) is set up

### Script Takes Too Long
- Run against one site first: use `-SiteUrl` parameter
- Run during off-peak hours
- Check your internet connection

### "Invalid site URL format"
- Your TenantAdminUrl must be: `https://tenant-admin.sharepoint.com`
- Don't include `/sites/` or other paths

### No CSV File Created
- Check the `./Logs` folder for error messages
- Make sure you have write permissions in the script folder
- Try specifying full output path: `-OutputPath "C:\Reports\Inventory.csv"`

---

## 📈 Excel Analysis Tips

1. **Open CSV in Excel**
   - File → Open → Select CSV file

2. **Create Pivot Table**
   - Select all data → Insert → Pivot Table
   - Analyze by SiteTemplate, Owner, or by List count

3. **Sort by Storage**
   - Data → Sort → StorageUsedMB (Largest to Smallest)

4. **Create Charts**
   - Select storage column → Insert → Chart
   - Visualize growth trends

---

## 🔄 Automation

### Schedule Weekly Scan (Windows Task Scheduler)
```powershell
# Create a batch file (.bat):
cd C:\Scripts
powershell.exe -ExecutionPolicy Bypass `
  -File ".\Get-SPOSiteCollectionInventory.ps1" `
  -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
  -OutputPath "C:\Reports\Inventory_$(Get-Date -Format 'yyyy-MM-dd').csv"
```

### Monthly Reports
```powershell
# Run monthly and archive
$date = Get-Date -Format "yyyy-MM"
.\Get-SPOSiteCollectionInventory.ps1 `
  -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
  -OutputPath "C:\Reports\Archive\Inventory_$date.csv"
```

---

## 📝 Log File Location

All logs saved to: `./Logs/SPO-Inventory-YYYY-MM-DD.log`

---

## ❓ Need Help?

Check the full guide: `GET-SITECOLLECTION-INVENTORY-GUIDE.md`
