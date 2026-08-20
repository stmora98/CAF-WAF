# Azure WAF/CAF Workshop Discovery Toolkit

A comprehensive suite of PowerShell scripts + Python dashboard agent for **Azure Well-Architected Framework (WAF)** and **Cloud Adoption Framework (CAF)** workshop preparation.

**One launcher → seven assessment phases → one consolidated dashboard.**

```
powershell/Launch-AzureWorkshop.ps1 → generate-dashboard.py → WAF_Dashboard.html
```

Point it at a tenant (single subscription or up to ~250), sign in with your own account, and it produces an interactive HTML dashboard scored against the five WAF pillars — no infrastructure to deploy, no Service Principal, no changes made to your environment (every query is read-only).

## Before you start

### 1. Where to install it

Copy or clone this whole folder onto your **local `C:\` drive**, for example `C:\AzureWorkshop-Toolkit`. Avoid running it from `Desktop`, `Documents`, or any path under `OneDrive` — those are commonly synced, and OneDrive can lock or partially encrypt the Excel files while the scripts are still writing to them, which breaks the dashboard step.

- **Cloned with git:** you're done, skip to step 2 below.
- **Downloaded as a ZIP from GitHub:** after extracting, right-click the extracted folder → **Properties** → check **Unblock** (Windows marks files downloaded from the internet as untrusted, which can block the scripts from running) → **OK**.

### 2. Roles and permissions you need

You sign in with your **own Azure AD account** — no Service Principal or app registration required.

**Minimum required (the toolkit will not produce useful results without this):**

| Role | Scope | Why |
|---|---|---|
| `Reader` | Every subscription you want assessed, or a Management Group above them | Lets Resource Graph, Governance, and Checklist scans read resources, policies, and role assignments |

**Optional — each unlocks one extra data source. Skipping any of these is fine: the toolkit reports the gap in the workbook/dashboard instead of failing.**

| Feature | Role / permission needed | What you get |
|---|---|---|
| Cost data (Advisor reservation recs, FinOps actual cost & budgets) | `Cost Management Reader` on each subscription | Real spend by service, configured budgets, reservation/savings plan recommendations |
| Reservation utilization | `Billing Reader` at the billing account (tenant) scope | Utilization % of your existing Reserved Instances |
| Key Vault secret/certificate expiration | Key Vault data-plane **Get** permission on each vault (access policy, or the `Key Vault Secrets User`/`Key Vault Certificates User` RBAC roles) | List of secrets/certificates nearing expiry |
| Defender XDR incidents & alerts | Microsoft Graph delegated scopes `SecurityIncident.Read.All` + `SecurityAlert.Read.All`, plus one of: Security Reader, Global Reader, Security Operator, Security Administrator | Recent security incidents and alerts |
| App credential expiry & guest users | Microsoft Graph delegated scopes `Application.Read.All` + `User.Read.All`, plus one of: Global Reader, Directory Reader, Cloud Application Administrator, Application Administrator | Expiring app registration secrets/certs, guest (B2B) account list |
| Defender for Endpoint data | A Defender for Endpoint license and API access (`Machine.Read.All`, `SecurityRecommendation.Read.All`, `Vulnerability.Read.All`) | Onboarded devices, endpoint recommendations, vulnerabilities |

Missing an optional permission never stops the run — check the `SourceStatus` sheet in `05_Security/AzureSecurity.xlsx` afterward to see exactly what was and wasn't collected.

### 3. What gets installed automatically

The launcher installs everything else for you the first time it runs — you don't need to prepare these yourself:

| Requirement | How it's handled |
|---|---|
| PowerShell 7 | **You must install this yourself** — [aka.ms/powershell](https://aka.ms/powershell) |
| Az PowerShell modules (`Az.Accounts`, `Az.Resources`, `Az.ResourceGraph`, `Az.Monitor`, `ImportExcel`) | Auto-installed for your user on first run |
| Python 3.6+ | Pre-installed in Azure Cloud Shell; install separately for local Windows runs ([python.org](https://www.python.org/downloads/)) |
| `openpyxl` (Python package) | Auto-installed by the launcher when Python is found |

## How to run it

### Option A — Windows, local (recommended)

1. Confirm the toolkit lives under `C:\` (see [Where to install it](#1-where-to-install-it) above) and [PowerShell 7](https://aka.ms/powershell) is installed.
2. Double-click **[`Start-AzureWorkshop.cmd`](Start-AzureWorkshop.cmd)**.
3. A console window opens and checks/installs prerequisites automatically — this only takes a while the first time.
4. An Azure sign-in window appears. Pick (or sign into) the account that has the roles listed above for the tenant/subscriptions you want to assess.
5. Watch the console — it prints `PHASE 1/7` through `PHASE 7/7` as it works. Leave the window open until it says the workshop completed successfully.
6. The dashboard opens automatically in your browser. You can reopen it any time from `AzureWorkshop\07_Dashboard\WAF_Dashboard.html`.

Need a faster first pass, or want to skip an optional phase? Double-clicking can't pass parameters — instead open a PowerShell 7 window in this folder and run it directly, e.g. `./powershell/Launch-AzureWorkshop.ps1 -SkipMetrics` (see [Parameters](#launch-azureworkshopps1) below for the full list of flags).

### Option B — Azure Cloud Shell

1. Go to [shell.azure.com](https://shell.azure.com) (or the Cloud Shell icon in the Azure Portal) and choose **PowerShell**.
2. Upload every file in this folder (use the upload icon in the Cloud Shell toolbar), keeping the `powershell/` and `checklists/` subfolders intact.
3. Run the launcher:
   ```powershell
   ./powershell/Launch-AzureWorkshop.ps1
   ```
4. Generate the dashboard once it finishes:
   ```powershell
   python3 generate-dashboard.py ~/AzureWorkshop
   ```
5. Download `07_Dashboard/WAF_Dashboard.html` (and ideally the whole output folder) using the Cloud Shell download button, then open it in your browser.

## How long does it take

Every phase paginates fully and retries automatically on Azure throttling, so runtime scales with how much you're scanning rather than stopping early. Rough guidance:

| Environment size | Subscriptions | Typical total time |
|---|---|---|
| Small | 1–5 | ~10–15 minutes |
| Medium | 5–50 | ~20–45 minutes |
| Large | 50–250 | ~1–3+ hours |

The **Metrics phase is almost always the bottleneck** — it queries Azure Monitor per resource (VMs, SQL databases, App Service plans, storage accounts), so it scales with resource count, not just subscription count. Use `-SkipMetrics` for a faster first pass if you only need inventory, Advisor, governance, security, and checklist findings.

| Phase | Script | Typical duration |
|---|---|---|
| 1. Resource Discovery | `Invoke-AzureDiscovery-CloudShell.ps1` | ~1 min |
| 2. Advisor Recommendations | `Invoke-AzureAdvisor-CloudShell.ps1` | ~30 sec |
| 2b. FinOps Extended | `Invoke-AzureFinOps-CloudShell.ps1` | ~1–3 min |
| 3. Metrics & Right-Sizing | `Invoke-AzureMetrics-CloudShell.ps1` | ~3–10 min per 100 resources |
| 4. Governance Visualizer | `Invoke-AzureGovernanceViz-CloudShell.ps1` | ~2 min |
| 5. Security Assessment | `Invoke-AzureSecurity-CloudShell.ps1` | ~1–3 min |
| 6. Review Checklists | `Invoke-AzureChecklists-CloudShell.ps1` | ~2–5 min |
| 7. Dashboard generation | `generate-dashboard.py` | ~5 sec |

> **No Service Principal required.** Everything runs under your logged-in user context. Defender for Cloud posture is the baseline; unavailable optional security sources are recorded in `SourceStatus` without aborting the assessment.

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
│                                                             │
│  Phase 5                                                    │
│  ┌──────────────────────────────────────────────────┐       │
│  │ Security Assessment                              │       │
│  │ → CSPM/MCSB + Defender XDR + Endpoint           │       │
│  └──────────────────┬───────────────────────────────┘       │
│                     ▼                                       │
│              05_Security/                                    │
│              AzureSecurity.xlsx                              │
│                                                             │
│  Phase 6                                                    │
│  ┌──────────────────────────────────────────────────┐       │
│  │ Review Checklists (Azure/review-checklists)      │       │
│  │ → Community WAF checks via Resource Graph        │       │
│  └──────────────────┬───────────────────────────────┘       │
│                     ▼                                       │
│              06_Checklists/                                  │
│              AzureChecklists.xlsx                            │
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
                    07_Dashboard/
                    WAF_Dashboard.html
```

## Scripts

| File | Purpose |
|---|---|
| `powershell/Launch-AzureWorkshop.ps1` | Master launcher — calls all scripts sequentially |
| `powershell/Invoke-AzureDiscovery-CloudShell.ps1` | Resource inventory via ARG (23 categories) |
| `powershell/Invoke-AzureAdvisor-CloudShell.ps1` | Advisor recommendations by WAF pillar |
| `powershell/Invoke-AzureFinOps-CloudShell.ps1` | Extended cost data: actual spend, budgets, reservation utilization, waste findings |
| `powershell/Invoke-AzureMetrics-CloudShell.ps1` | Right-sizing via Azure Monitor metrics |
| `powershell/Invoke-AzureGovernanceViz-CloudShell.ps1` | Governance HTML (AzGovViz-style) |
| `powershell/Invoke-AzureSecurity-CloudShell.ps1` | Defender for Cloud CSPM/MCSB, Defender XDR, and Endpoint export |
| `powershell/Invoke-AzureChecklists-CloudShell.ps1` | Community WAF checks from [Azure/review-checklists](https://github.com/Azure/review-checklists) (vendored locally, no network needed) via Resource Graph |
| `generate-dashboard.py` | Consolidated dashboard with scoring + action items |

See [How long does it take](#how-long-does-it-take) above for per-phase duration estimates.

## Output Structure

By default, results land in an `AzureWorkshop` folder at the repository root (pass `-OutputDir` to change this, or `-KeepPrevious` to archive rather than overwrite a prior run). Both the PowerShell and Python launchers use this same location, so `AzureWorkshop.history.json` at the repository root tracks progress across runs from either launcher:

```
AzureWorkshop.history.json                 # Shared score/action history for both launchers
AzureWorkshop/
├── 01_Discovery/
│   └── AzureDiscovery.xlsx          # 23 tabs: VMs, VNets, Storage, DBs, etc.
├── 02_Advisor/
│   ├── AzureAdvisor.xlsx            # 8 tabs: By pillar + summaries
│   └── AzureFinOps.xlsx             # Actual cost, budgets, reservation utilization, waste findings
├── 03_Metrics/
│   └── AzureMetrics.xlsx            # 6 tabs: VM/SQL/AppPlan sizing, diagnostics
├── 04_Governance/
│   ├── AzureGovernance.html         # Interactive governance report (AzGovViz-style)
│   └── AzureGovernance.xlsx         # MG hierarchy, policies, RBAC, Defender
├── 05_Security/
│   └── AzureSecurity.xlsx           # CSPM, MCSB, plans, incidents, alerts, Endpoint, source status
├── 06_Checklists/
│   └── AzureChecklists.xlsx         # Non-compliant findings from Azure/review-checklists ARG checks
└── 07_Dashboard/
    └── WAF_Dashboard.html           # ★ Final consolidated dashboard
```

## Dashboard Features

The Python agent produces a **WAF_Dashboard.html** with:

- **Overall WAF Score** (0-100, weighted across 5 pillars)
- **Per-Pillar Scores** with visual progress bars
- **Prioritized Action Items** sorted by severity (Critical → Low)
- **Resource Summary** tables
- **Advisor Findings** aggregated view
- **Dedicated Security tab** for CSPM recommendations, MCSB compliance, Defender plans, XDR incidents/alerts, Endpoint exposure, and source coverage
- **Governance resource recommendations** derived from the generated findings, with priority, supporting action items, and Microsoft Learn guidance

### Scoring Logic

| Pillar | What's Evaluated |
|---|---|
| **Reliability** | Availability zones, backup vaults, DDoS, Advisor HA recs |
| **Security** | Defender Secure Score, CSPM recommendations, MCSB, Defender plans, XDR incidents/alerts, public storage, Key Vault, and TLS |
| **Cost Optimization** | Orphaned resources, deallocated VMs, right-sizing, Advisor |
| **Operational Excellence** | Diagnostic settings, tag coverage, policy compliance |
| **Performance Efficiency** | Saturated VMs/SQL, Advisor performance recs |

## Parameters

### Launch-AzureWorkshop.ps1

| Parameter | Description |
|---|---|
| `-OutputDir` | Custom output directory. Default: `AzureWorkshop` at the repository root |
| `-SkipMetrics` | Skip the metrics phase (faster, but no right-sizing data) |
| `-SkipFinOps` | Skip the extended cost export (actual cost, budgets, reservation utilization) |
| `-KeepPrevious` | Archive the previous output folder (timestamp suffix) instead of deleting it |
| `-SkipGraphSecurity` | Skip the interactive Microsoft Graph sign-in for Defender XDR incidents/alerts, keeping Defender for Cloud + Endpoint data |
| `-ForceAccountSelection` | Force a full interactive sign-in page if the account picker seems stuck on a cached account |

### Individual Scripts

Each script can also be run standalone:
```powershell
./powershell/Invoke-AzureDiscovery-CloudShell.ps1      # Runs independently
./powershell/Invoke-AzureAdvisor-CloudShell.ps1        # Runs independently
./powershell/Invoke-AzureMetrics-CloudShell.ps1        # Runs independently (slower)
./powershell/Invoke-AzureGovernanceViz-CloudShell.ps1  # Runs independently
./powershell/Invoke-AzureSecurity-CloudShell.ps1       # Runs independently; requests optional Graph scopes
./powershell/Invoke-AzureSecurity-CloudShell.ps1 -LookbackDays 90
./powershell/Invoke-AzureSecurity-CloudShell.ps1 -SkipGraphSecurity     # Keep Cloud + Endpoint; skip Graph sign-in
./powershell/Invoke-AzureSecurity-CloudShell.ps1 -MaxEndpointRecords 10000
./powershell/Invoke-AzureSecurity-CloudShell.ps1 -SkipDefenderPortal  # Defender for Cloud posture only
./powershell/Invoke-AzureChecklists-CloudShell.ps1     # Runs independently; scans all community checklists
./powershell/Invoke-AzureChecklists-CloudShell.ps1 -Services keyvault,aks,storage  # Scope to specific services
```

### Security Data Sources

| Source | Data | Access behavior |
|---|---|---|
| Defender for Cloud via Azure Resource Graph | Secure Score, controls, recommendations, MCSB, regulatory standards, plans, cloud alerts | Uses existing Azure RBAC context |
| Microsoft Graph Security | Defender XDR incidents and alerts | Prompts for delegated consent when scopes are missing |
| Defender for Endpoint API | Machines, security recommendations, vulnerabilities | Requires a compatible license, API role, and token audience |

Check the `SourceStatus` worksheet before interpreting zero findings. `NoData` means the source was queried successfully and returned no records; `Forbidden`, `Unavailable`, `Partial`, or `Skipped` means the assessment does not have complete visibility for that source. Defender for Endpoint datasets default to 5,000 records per sheet; reaching `-MaxEndpointRecords` is reported as `Partial` so the workbook closes predictably.

### Checklist Data Source

`Invoke-AzureChecklists-CloudShell.ps1` runs each community check's embedded Azure Resource Graph query against the current subscription(s), using the checklist JSON files vendored locally under [`checklists/`](checklists/) — sourced from [Azure/review-checklists](https://github.com/Azure/review-checklists). This is **not an official Microsoft dataset**; it's an open-source, community-curated set of WAF checks per Azure service (Key Vault, AKS, Storage, etc.), each tagged with a Reliability/Security/Cost/Operations/Performance pillar.

- No GitHub/network access is required at run time — only Azure Resource Graph, same as every other script in this toolkit. This keeps it safe to run from a locked-down Cloud Shell/launcher session.
- Run `./powershell/Update-ReviewChecklists.ps1` manually (from a machine with internet access) whenever you want to refresh `checklists/` with the latest upstream files, then commit the folder.
- Findings only **add action items** to the dashboard (tagged with their source and a link back to the check's documentation) — they do **not** change the Advisor Score pillar math described below.
- Use `-Services` to scope the scan to specific checklist files (e.g. `-Services keyvault,aks,storage`) if a full catalog scan is too slow.

## Notes

- All queries are **read-only** — no changes are made to the environment.
- Results auto-paginate fully (1,000 rows per page via Azure Resource Graph, looping until every page is retrieved) and automatically retry with backoff on throttling (`429`/`5xx`) — tested for tenants with up to ~250 subscriptions.
- The metrics script is the slowest (~3-10 min per 100 resources) as it queries per-resource; it scales with resource count more than subscription count.
- Use `-SkipMetrics` if you're short on time during a workshop.
- The Governance HTML can be used standalone as an AzGovViz-lite alternative.
- Security API failures are isolated by source and do not prevent the workbook or dashboard from being generated.
- Checklist findings from Azure/review-checklists are additive action items only and never affect pillar scores.

## License

This project is licensed under the [MIT License](LICENSE).
