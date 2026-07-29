# Azure Environment Discovery for WAF/CAF Workshops

A PowerShell script that uses **Azure Resource Graph** to comprehensively inventory a customer's Azure environment and exports the results to an **Excel workbook** with 23 categorized worksheets.

## Prerequisites

| Requirement | Notes |
|---|---|
| PowerShell 7+ or Windows PowerShell 5.1 | Cloud Shell has this pre-installed |
| `Az.Accounts` module | Auto-installed by the script |
| `Az.ResourceGraph` module | Auto-installed by the script |
| `ImportExcel` module | Auto-installed by the script |
| Azure RBAC | Reader role on target subscriptions |

## Quick Start

### Azure Cloud Shell

```powershell
# Upload the script to Cloud Shell, then:
./Invoke-AzureDiscovery.ps1
```

### Local PowerShell

```powershell
# Authenticate first
Connect-AzAccount

# Run discovery
./Invoke-AzureDiscovery.ps1 -OutputPath "C:\Reports\CustomerDiscovery.xlsx"
```

## Parameters

| Parameter | Required | Description |
|---|---|---|
| `-OutputPath` | No | Output Excel file path. Default: `AzureDiscovery_<timestamp>.xlsx` |
| `-TenantId` | No | Target a specific Azure AD tenant |

## Worksheets Generated

The script produces an Excel workbook covering all **5 WAF pillars**:

### 🏗️ Foundation
| # | Sheet | Content |
|---|---|---|
| 1 | Subscriptions | Tenant subscription inventory |
| 2 | ResourceGroups | All resource groups with tags |
| 3 | ResourceSummary | Resource counts by type & region |

### ⚡ Compute
| # | Sheet | Content |
|---|---|---|
| 4 | VMs | VM sizes, OS, power state, availability zones |
| 5 | AppServices | App Services, Functions, App Service Plans |
| 6 | AKS | Kubernetes version, node pools, network plugin |

### 🌐 Networking
| # | Sheet | Content |
|---|---|---|
| 7 | VNets | Address spaces, subnet count, DDoS protection |
| 8 | NSGs | Rule counts, associated subnets/NICs |
| 9 | LoadBalancers | LBs, App Gateways, Front Door, CDN |
| 10 | Firewalls | Azure Firewall, WAF policies |
| 11 | PrivateEndpoints | PL connections and status |
| 12 | PublicIPs | Allocation, SKU, association status |
| 13 | DNS | DNS Zones & Private DNS |

### 🔐 Security
| # | Sheet | Content |
|---|---|---|
| 14 | Storage | TLS, HTTPS, network rules, public access |
| 15 | KeyVaults | Soft delete, purge protection, RBAC mode |
| 16 | ManagedIdentities | User-assigned identities |

### 📊 Operations & Governance
| # | Sheet | Content |
|---|---|---|
| 17 | Monitoring | Log Analytics, App Insights, Alerts |
| 18 | PolicyAssignments | Policy enforcement and scope |
| 19 | BackupVaults | Recovery Services & Backup vaults |

### 💰 Cost Optimization
| # | Sheet | Content |
|---|---|---|
| 20 | Databases | SQL, PostgreSQL, Cosmos DB, Redis with SKUs |
| 21 | UnattachedDisks | Orphaned disks (wasted spend) |
| 22 | DeallocatedVMs | Stopped VMs still incurring costs |
| 23 | AdvisorRecommendations | Azure Advisor suggestions |

## WAF Pillar Mapping

| Pillar | Relevant Sheets |
|---|---|
| **Reliability** | VMs (zones), AKS, VNets, LBs, BackupVaults, DNS |
| **Security** | NSGs, Firewalls, KeyVaults, Storage, PrivateEndpoints, ManagedIdentities, PolicyAssignments |
| **Cost Optimization** | UnattachedDisks, DeallocatedVMs, ResourceSummary, AdvisorRecommendations |
| **Operational Excellence** | Monitoring, PolicyAssignments, ResourceGroups (tags) |
| **Performance Efficiency** | VMs (sizing), AKS, AppServices, LoadBalancers, Databases |

## Notes

- The script uses Azure Resource Graph for fast, cross-subscription queries.
- All queries are **read-only** — no changes are made to the environment.
- Results are paginated automatically (1000 rows per page).
- Empty categories still produce a worksheet tab with a placeholder row.
