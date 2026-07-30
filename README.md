# Azure WAF/CAF Workshop Discovery Toolkit

A comprehensive suite of PowerShell scripts + Python dashboard agent for **Azure Well-Architected Framework (WAF)** and **Cloud Adoption Framework (CAF)** workshop preparation.

**One launcher → four discovery phases → one consolidated dashboard.**

```
Launch-AzureWorkshop.ps1 → generate-dashboard.py → WAF_Dashboard.html
```

## 🚀 Quick Start

```powershell
# 1. Upload all scripts to Azure Cloud Shell (PowerShell)

# 2. Run the launcher (executes all discovery scripts)
./Launch-AzureWorkshop.ps1

# 3. Generate the consolidated dashboard
python3 generate-dashboard.py ~/AzureWorkshop_<timestamp>

# 4. Download and open WAF_Dashboard.html in your browser
```

## Prerequisites

| Requirement | Notes |
|---|---|
| Azure Cloud Shell (PowerShell) | Or local PS 7+ with Az module |
| `Reader` RBAC role | On Management Group or Subscriptions |
| Python 3.6+ | Pre-installed in Cloud Shell |
| `openpyxl` | Auto-installed by the Python script |
| `ImportExcel` PS module | Auto-installed by the launcher |

> **No Service Principal required.** Everything runs under your logged-in user context.

## Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                  Launch-AzureWorkshop.ps1                    │
│                    (Master Launcher)                         │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  Phase 1               Phase 2              Phase 3         │
│  ┌──────────────┐     ┌──────────────┐    ┌──────────────┐ │
│  │ Discovery    │     │ Advisor      │    │ Metrics      │ │
│  │ (ARG queries)│     │ (By pillar)  │    │ (Right-size) │ │
│  └──────┬───────┘     └──────┬───────┘    └──────┬───────┘ │
│         │                    │                    │         │
│         ▼                    ▼                    ▼         │
│  01_Discovery/        02_Advisor/          03_Metrics/      │
│  AzureDiscovery.xlsx  AzureAdvisor.xlsx    AzureMetrics.xlsx│
│                                                             │
│  Phase 4                                                    │
│  ┌──────────────────────────────────────────────────┐       │
│  │ Governance Visualizer (Lite)                     │       │
│  │ → Interactive HTML + Excel                       │       │
│  └──────────────────┬───────────────────────────────┘       │
│                     ▼                                       │
│              04_Governance/                                  │
│              AzureGovernance.html + .xlsx                    │
└─────────────────────────────────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────────┐
│              generate-dashboard.py                           │
│              (Python Analysis Agent)                         │
├─────────────────────────────────────────────────────────────┤
│  • Reads all Excel outputs                                  │
│  • Scores each WAF pillar (0-100)                          │
│  • Generates prioritized action items                       │
│  • Produces consolidated HTML dashboard                     │
└──────────────────────────┬──────────────────────────────────┘
                           ▼
                    05_Dashboard/
                    WAF_Dashboard.html
```

## Scripts

| File | Purpose | Duration |
|---|---|---|
| `Launch-AzureWorkshop.ps1` | Master launcher — calls all scripts sequentially | Total |
| `Invoke-AzureDiscovery-CloudShell.ps1` | Resource inventory via ARG (23 categories) | ~1 min |
| `Invoke-AzureAdvisor-CloudShell.ps1` | Advisor recommendations by WAF pillar | ~30 sec |
| `Invoke-AzureMetrics-CloudShell.ps1` | Right-sizing via Azure Monitor metrics | ~3-10 min |
| `Invoke-AzureGovernanceViz-CloudShell.ps1` | Governance HTML (AzGovViz-style) | ~2 min |
| `generate-dashboard.py` | Consolidated dashboard with scoring + action items | ~5 sec |

## Output Structure

```
~/AzureWorkshop_<timestamp>/
├── 01_Discovery/
│   └── AzureDiscovery.xlsx          # 23 tabs: VMs, VNets, Storage, DBs, etc.
├── 02_Advisor/
│   └── AzureAdvisor.xlsx            # 8 tabs: By pillar + summaries
├── 03_Metrics/
│   └── AzureMetrics.xlsx            # 6 tabs: VM/SQL/AppPlan sizing, diagnostics
├── 04_Governance/
│   ├── AzureGovernance.html         # Interactive governance report (AzGovViz-style)
│   └── AzureGovernance.xlsx         # MG hierarchy, policies, RBAC, Defender
└── 05_Dashboard/
    └── WAF_Dashboard.html           # ★ Final consolidated dashboard
```

## Dashboard Features

The Python agent produces a **WAF_Dashboard.html** with:

- **Overall WAF Score** (0-100, weighted across 5 pillars)
- **Per-Pillar Scores** with visual progress bars
- **Prioritized Action Items** sorted by severity (Critical → Low)
- **Resource Summary** tables
- **Advisor Findings** aggregated view

### Scoring Logic

| Pillar | What's Evaluated |
|---|---|
| **Reliability** | Availability zones, backup vaults, DDoS, Advisor HA recs |
| **Security** | Secure Score, public storage, Key Vault config, TLS, Advisor |
| **Cost Optimization** | Orphaned resources, deallocated VMs, right-sizing, Advisor |
| **Operational Excellence** | Diagnostic settings, tag coverage, policy compliance |
| **Performance Efficiency** | Saturated VMs/SQL, Advisor performance recs |

## Parameters

### Launch-AzureWorkshop.ps1

| Parameter | Description |
|---|---|
| `-OutputDir` | Custom output directory. Default: `~/AzureWorkshop_<timestamp>` |
| `-SkipMetrics` | Skip the metrics phase (faster, but no right-sizing data) |

### Individual Scripts

Each script can also be run standalone:
```powershell
./Invoke-AzureDiscovery-CloudShell.ps1      # Runs independently
./Invoke-AzureAdvisor-CloudShell.ps1        # Runs independently
./Invoke-AzureMetrics-CloudShell.ps1        # Runs independently (slower)
./Invoke-AzureGovernanceViz-CloudShell.ps1  # Runs independently
```

## Notes

- All queries are **read-only** — no changes are made to the environment.
- Results auto-paginate (1000 rows per page via ARG).
- The metrics script is the slowest (~1-3 min per 100 resources) as it queries per-resource.
- Use `-SkipMetrics` if you're short on time during a workshop.
- The Governance HTML can be used standalone as an AzGovViz-lite alternative.
