# Azure WAF/CAF Workshop - Python edition

Functional Python port of the PowerShell workshop (`Launch-AzureWorkshop.ps1` and the
`Invoke-Azure*-CloudShell.ps1` scripts). Same phases, same output folder layout, same
Excel sheet names/columns, so the existing [generate-dashboard.py](../generate-dashboard.py)
consumes either version's output interchangeably.

Unlike the PowerShell version (which needs `pwsh` + Az modules), this port only depends
on cross-platform Python SDKs, so it runs the same way on **Windows, macOS, and Linux**.

## Setup

Nothing to install manually. Every script checks for its required packages on
startup and installs anything missing into whichever Python interpreter runs it
(same self-installing behavior as the PowerShell scripts' `Install-Module`) - just
make sure Python 3.9+ is on your machine, then run a script:

```bash
python launch_workshop.py
```

The first run prints `Installing missing Python dependencies (...)` and installs them
automatically; every run after that starts immediately since the packages are already
present. No venv is required - this works the same whether you use the system Python,
a venv, or any other interpreter, on any OS.

If you prefer an isolated environment (optional, not required):

```bash
# Windows (PowerShell)
python -m venv .venv
.venv\Scripts\Activate.ps1
python launch_workshop.py

# macOS / Linux
python3 -m venv .venv
source .venv/bin/activate
python launch_workshop.py
```

Requires access to the Azure subscriptions being assessed (Reader role is enough for
most data; Cost Management Reader / Billing Reader / Key Vault Get / Microsoft Graph
Security roles unlock the optional cost and security datasets, same as the PowerShell
version).

> On Windows, always invoke scripts as `python launch_workshop.py` (never
> `.\launch_workshop.py` directly) - `.py` files aren't reliably associated with an
> interpreter in every Windows setup, and running the bare path can silently do nothing.

## Run the full workshop

```bash
python launch_workshop.py
```

This signs in once (interactive browser), auto-discovers every enabled subscription in
the tenant (no subscription picker), and runs all 7 phases into `../AzureWorkshop/`:

```
python launch_workshop.py --output-dir /path/to/ContosoAssessment
python launch_workshop.py --skip-metrics --skip-finops
python launch_workshop.py --keep-previous
python launch_workshop.py --skip-graph-security
```

## Run a single phase

Each phase script can also run standalone (it signs in and discovers subscriptions on
its own, writing to the current directory or `$AZWORKSHOP_OUTPUT` if set):

```powershell
python phase1_discovery.py
python phase2_advisor.py
python phase2b_finops.py
python phase3_metrics.py
python phase4_governance.py
python phase5_security.py
python phase6_checklists.py
```

## Design notes

- **Auth**: `azure-identity` `InteractiveBrowserCredential` (equivalent to
  `Connect-AzAccount`). Subscription scope is resolved once by the launcher and shared
  across phases via the `AZWORKSHOP_SUBSCRIPTION_IDS` env var (or auto-discovered per
  phase when run standalone), so the assessment always covers every enabled
  subscription rather than a single selected one.
- **Resource Graph**: `azure-mgmt-resourcegraph` with the same KQL queries as the
  PowerShell scripts, paginated and retried with exponential backoff on 429/5xx.
- **Metrics**: classic `azure-mgmt-monitor` `metrics.list` per resource (the direct
  equivalent of `Get-AzMetric`). Pinned to `azure-mgmt-monitor==6.0.2` because 7.0.0
  removed the `diagnostic_settings` operation group this phase also depends on.
- **Cost Management / Microsoft Graph Security / Defender for Endpoint**: no official
  lightweight synchronous SDK covers these; accessed via direct REST calls using an
  `azure-identity` token, mirroring the PowerShell scripts' `Invoke-RestMethod` calls.
- **Dashboard**: [generate-dashboard.py](../generate-dashboard.py) is reused unchanged
  - it only reads raw Excel cell values, so it doesn't care which implementation
  produced the workbook.
