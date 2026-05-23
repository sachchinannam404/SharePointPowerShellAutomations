# Site Collection Inventory - Real-World Examples

## Scenario 1: Department Storage Audit

### Goal
Find which department sites are using the most storage

### Steps
```powershell
# 1. Run the inventory script
.\Get-SPOSiteCollectionInventory.ps1 -TenantAdminUrl "https://contoso-admin.sharepoint.com"

# 2. Import and analyze in PowerShell
$report = Import-Csv "SiteCollectionInventory_*.csv"
$report | Sort-Object { [decimal]$_.StorageUsedMB } -Descending | 
  Format-Table SiteTitle, StorageUsedMB, Users, TotalItems -AutoSize
```

### Sample Output
```
SiteTitle              StorageUsedMB Users TotalItems
-----------            ------------- ----- ----------
IT Shared              3072.25       25    600
Finance Hub            2048.00       20    450
Marketing Hub          1024.50       15    250
HR Hub                 768.75        18    320
Sales Hub              512.25        12    180
```

---

## Scenario 2: Site Governance - Lists and Libraries Audit

### Goal
Identify sites with excessive lists/libraries for consolidation

### Steps
```powershell
# Find sites with too many lists
$report = Import-Csv "SiteCollectionInventory_*.csv"
$report | Where-Object { [int]$_.Lists -gt 5 } | 
  Select-Object SiteTitle, Lists, DocumentLibraries, Users | 
  Sort-Object Lists -Descending

# Export problematic sites
$report | Where-Object { [int]$_.Lists -gt 5 } | 
  Export-Csv -Path "HighListCount.csv" -NoTypeInformation
```

### Sample Finding
```
SiteTitle    Lists DocumentLibraries Users
---------    ----- ------------------- -----
IT Shared    8     5                   25
Finance Hub  6     4                   20
```
**Action:** Review and consolidate redundant lists

---

## Scenario 3: User Access Analysis

### Goal
Identify underutilized sites with low user engagement

### Steps
```powershell
# Find sites with few users
$report = Import-Csv "SiteCollectionInventory_*.csv"
$report | Where-Object { [int]$_.Users -lt 5 } | 
  Select-Object SiteTitle, Users, TotalItems, StorageUsedMB | 
  Sort-Object Users

# Calculate engagement ratio
$report | Select-Object SiteTitle, Users, TotalItems, @{
    Name='Engagement';
    Expression={ if([int]$_.Users -gt 0) { [math]::Round([int]$_.TotalItems/[int]$_.Users, 2) } else { 0 } }
} | Sort-Object Engagement
```

### Sample Finding
```
SiteTitle         Users TotalItems Engagement
---------         ----- ---------- ----------
Operations Hub    8     90         11.25
Corporate         50    150        3.00
IT Shared         25    600        24.00
```
**Action:** Low engagement sites may need archiving or reorganization

---

## Scenario 4: Compliance Report

### Goal
Generate compliance report for audit trail

### Steps
```powershell
# Create comprehensive audit report
$report = Import-Csv "SiteCollectionInventory_*.csv"

$auditReport = $report | Select-Object @{
    Name='Site';
    Expression={$_.SiteTitle}
}, @{
    Name='URL';
    Expression={$_.SiteUrl}
}, @{
    Name='Owner';
    Expression={$_.Owner}
}, @{
    Name='Storage (MB)';
    Expression={[decimal]$_.StorageUsedMB}
}, @{
    Name='Total Items';
    Expression={[int]$_.TotalItems}
}, @{
    Name='Users';
    Expression={[int]$_.Users}
}, @{
    Name='Last Modified';
    Expression={$_.LastModified}
}

$auditReport | Export-Csv -Path "Compliance_Report_$(Get-Date -Format 'yyyy-MM-dd').csv" -NoTypeInformation
```

---

## Scenario 5: Site Template Analysis

### Goal
Understand site template distribution and usage

### Steps
```powershell
# Count sites by template
$report = Import-Csv "SiteCollectionInventory_*.csv"

$report | Group-Object SiteTemplate | Select-Object Name, @{
    Name='Count';
    Expression={$_.Count}
}, @{
    Name='TotalStorage(MB)';
    Expression={ ($_.Group | Measure-Object -Property StorageUsedMB -Sum).Sum }
} | Format-Table -AutoSize
```

### Sample Output
```
Name                    Count TotalStorage(MB)
----                    ----- ----------------
STS                     6     4640.25
SITEPAGEPUBLISHING      1     512.00
```

---

## Scenario 6: Monthly Growth Tracking

### Goal
Track storage growth month-over-month

### Steps
```powershell
# Run script each month
$currentMonth = Get-Date -Format "yyyy-MM"
.\Get-SPOSiteCollectionInventory.ps1 `
  -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
  -OutputPath "C:\Reports\Archive\Inventory_$currentMonth.csv"

# Compare with previous month
$current = Import-Csv "Inventory_2026-05.csv" | 
  Measure-Object -Property StorageUsedMB -Sum
$previous = Import-Csv "Inventory_2026-04.csv" | 
  Measure-Object -Property StorageUsedMB -Sum

$growth = $current.Sum - $previous.Sum
Write-Host "Storage growth this month: $([math]::Round($growth, 2)) MB"
```

---

## Scenario 7: Critical Sites Report

### Goal
Identify critical sites based on multiple factors

### Steps
```powershell
# Mark sites as critical (high users + high items + high storage)
$report = Import-Csv "SiteCollectionInventory_*.csv"

$critical = $report | Where-Object {
    ([int]$_.Users -gt 20) -and 
    ([int]$_.TotalItems -gt 300) -and 
    ([decimal]$_.StorageUsedMB -gt 1000)
} | Select-Object SiteTitle, Users, TotalItems, StorageUsedMB

Write-Host "Critical Sites: $($critical.Count)"
$critical | Format-Table -AutoSize
```

---

## Scenario 8: Export to Excel with Formatting

### Goal
Create professional Excel report with formatting

### Steps
```powershell
# Requires ImportExcel module
Install-Module -Name ImportExcel -Force

$report = Import-Csv "SiteCollectionInventory_*.csv"

$excelParams = @{
    Path = "SiteCollectionInventory_Report.xlsx"
    WorksheetName = "Inventory"
    TableName = "SiteInventory"
    TableStyle = "Light17"
    AutoSize = $true
    FreezeTopRow = $true
}

$report | Export-Excel @excelParams

# Add conditional formatting for storage
Add-ConditionalFormatting -Path "SiteCollectionInventory_Report.xlsx" `
  -WorksheetName "Inventory" `
  -Range "E:E" `
  -ConditionalFormat DatabarGradient `
  -ConditionalFormatType Positive
```

---

## Scenario 9: Team Productivity Analysis

### Goal
Analyze site productivity metrics

### Steps
```powershell
$report = Import-Csv "SiteCollectionInventory_*.csv"

$productivity = $report | Select-Object SiteTitle, @{
    Name='Users';
    Expression={[int]$_.Users}
}, @{
    Name='Items Per User';
    Expression={if([int]$_.Users -gt 0) {[math]::Round([int]$_.TotalItems/[int]$_.Users, 2)} else {0}}
}, @{
    Name='Libraries';
    Expression={[int]$_.DocumentLibraries}
}, @{
    Name='Pages';
    Expression={[int]$_.SitePages}
} | Sort-Object 'Items Per User' -Descending

$productivity | Format-Table -AutoSize
```

---

## Scenario 10: Archival Candidates

### Goal
Identify sites for potential archival

### Steps
```powershell
$report = Import-Csv "SiteCollectionInventory_*.csv"

# Criteria: Few users, few items, not recently modified
$archiveCandidate = $report | Where-Object {
    ([int]$_.Users -lt 3) -and 
    ([int]$_.TotalItems -lt 50) -and 
    ([datetime]$_.LastModified -lt (Get-Date).AddMonths(-6))
} | Select-Object SiteTitle, Users, TotalItems, LastModified, StorageUsedMB

Write-Host "Potential Archive Candidates: $($archiveCandidate.Count)"
$archiveCandidate | Export-Csv "Archive_Candidates.csv" -NoTypeInformation
```

---

## Quick PowerShell Snippets

### Calculate Total Across All Sites
```powershell
$report = Import-Csv "SiteCollectionInventory_*.csv"
$report | Measure-Object -Property StorageUsedMB, TotalItems, Users -Sum
```

### Find Sites Matching Pattern
```powershell
$report = Import-Csv "SiteCollectionInventory_*.csv"
$report | Where-Object { $_.SiteTitle -like "*Hub*" }
```

### Export Specific Columns
```powershell
$report = Import-Csv "SiteCollectionInventory_*.csv"
$report | Select-Object SiteTitle, Owner, StorageUsedMB | Export-Csv "Summary.csv" -NoTypeInformation
```

### Generate Statistics
```powershell
$report = Import-Csv "SiteCollectionInventory_*.csv"
$stats = @{
    TotalSites = $report.Count
    TotalStorage = ($report | Measure-Object -Property StorageUsedMB -Sum).Sum
    TotalUsers = ($report | Measure-Object -Property Users -Sum).Sum
    TotalItems = ($report | Measure-Object -Property TotalItems -Sum).Sum
    AvgStoragePerSite = ($report | Measure-Object -Property StorageUsedMB -Average).Average
}
$stats
```
