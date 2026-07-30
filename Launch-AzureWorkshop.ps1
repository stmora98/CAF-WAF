<#
.SYNOPSIS
    Azure WAF/CAF Workshop Launcher - Runs all discovery scripts and consolidates output.

.DESCRIPTION
    Master launcher that sequentially executes:
      1. Invoke-AzureDiscovery-CloudShell.ps1       (Resource inventory)
      2. Invoke-AzureAdvisor-CloudShell.ps1         (Advisor recommendations)
      3. Invoke-AzureMetrics-CloudShell.ps1         (Right-sizing & reliability)
      4. Invoke-AzureGovernanceViz-CloudShell.ps1   (Governance + HTML report)
    
    All outputs are consolidated into a single timestamped folder.
    After completion, run the Python dashboard agent to generate the final report.

.PARAMETER OutputDir
    Base output directory. Defaults to ~/AzureWorkshop_<timestamp>

.PARAMETER SkipMetrics
    Skip the metrics script (it's slower, queries per-resource). Use if short on time.

.EXAMPLE
    ./Launch-AzureWorkshop.ps1
    ./Launch-AzureWorkshop.ps1 -SkipMetrics

.NOTES
    Run in Azure Cloud Shell (PowerShell) with Reader access on subscriptions.
#>

[CmdletBinding()]
param(
    [string]$OutputDir,
    [switch]$SkipMetrics
)

$ErrorActionPreference = 'Continue'
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

if (-not $OutputDir) {
    $OutputDir = "$HOME/AzureWorkshop_$timestamp"
}

New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
New-Item -ItemType Directory -Path "$OutputDir/01_Discovery" -Force | Out-Null
New-Item -ItemType Directory -Path "$OutputDir/02_Advisor" -Force | Out-Null
New-Item -ItemType Directory -Path "$OutputDir/03_Metrics" -Force | Out-Null
New-Item -ItemType Directory -Path "$OutputDir/04_Governance" -Force | Out-Null
New-Item -ItemType Directory -Path "$OutputDir/05_Dashboard" -Force | Out-Null

# ─── Ensure modules ──────────────────────────────────────────────────────────
Write-Host "Checking prerequisites..." -ForegroundColor Yellow

$requiredModules = @(
    'Az.Accounts',
    'Az.Resources',
    'Az.ResourceGraph',
    'Az.Monitor',
    'ImportExcel'
)

foreach ($mod in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Host "  Installing $mod..." -ForegroundColor DarkYellow
        Install-Module -Name $mod -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck
    }
}

Import-Module Az.Accounts -ErrorAction Stop
Import-Module Az.Resources -ErrorAction Stop
Import-Module Az.ResourceGraph -ErrorAction Stop
Import-Module Az.Monitor -ErrorAction SilentlyContinue
Import-Module ImportExcel -ErrorAction Stop

# Ensure Python + openpyxl for the dashboard agent
$pythonCmd = if (Get-Command python3 -ErrorAction SilentlyContinue) { "python3" }
             elseif (Get-Command python -ErrorAction SilentlyContinue) { "python" }
             else { $null }

if ($pythonCmd) {
    Write-Host "  Checking Python openpyxl..." -ForegroundColor DarkGray
    & $pythonCmd -c "import openpyxl" 2>$null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  Installing openpyxl..." -ForegroundColor DarkYellow
        & $pythonCmd -m pip install openpyxl --quiet --user 2>$null
    }
} else {
    Write-Host "  ⚠️ Python not found - dashboard generation will need manual pip install" -ForegroundColor DarkYellow
}

Write-Host "  ✓ All prerequisites satisfied.`n" -ForegroundColor Green

$context = Get-AzContext
if (-not $context) {
    Write-Host "Not authenticated. Running Connect-AzAccount..." -ForegroundColor Yellow
    Connect-AzAccount
    $context = Get-AzContext
}

Write-Host @"

╔═══════════════════════════════════════════════════════════════╗
║         Azure WAF/CAF Workshop - Discovery Launcher          ║
╠═══════════════════════════════════════════════════════════════╣
║  Tenant:  $($context.Tenant.Id)      ║
║  Account: $($context.Account.Id)                      ║
║  Output:  $OutputDir    ║
╚═══════════════════════════════════════════════════════════════╝

"@ -ForegroundColor Cyan

$startTime = Get-Date
$scriptDir = $PSScriptRoot
if (-not $scriptDir) { $scriptDir = (Get-Location).Path }

# ═══════════════════════════════════════════════════════════════════════════════
# PHASE 1: Resource Discovery
# ═══════════════════════════════════════════════════════════════════════════════
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkCyan
Write-Host " PHASE 1/4: Resource Discovery" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkCyan

$env:AZWORKSHOP_OUTPUT = "$OutputDir/01_Discovery"
try {
    # Inline the discovery logic with redirected output path
    $discoveryOutput = "$OutputDir/01_Discovery/AzureDiscovery.xlsx"
    & "$scriptDir/Invoke-AzureDiscovery-CloudShell.ps1" *>&1 | ForEach-Object { Write-Host $_ }
    # Move output file if it was created in $HOME
    $latestDiscovery = Get-ChildItem "$HOME/AzureDiscovery_*.xlsx" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($latestDiscovery) {
        Move-Item $latestDiscovery.FullName "$OutputDir/01_Discovery/AzureDiscovery.xlsx" -Force
    }
} catch {
    Write-Host "  ⚠️ Discovery script encountered an error: $($_.Exception.Message)" -ForegroundColor Yellow
}

# ═══════════════════════════════════════════════════════════════════════════════
# PHASE 2: Advisor Recommendations
# ═══════════════════════════════════════════════════════════════════════════════
Write-Host "`n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkCyan
Write-Host " PHASE 2/4: Advisor Recommendations" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkCyan

try {
    & "$scriptDir/Invoke-AzureAdvisor-CloudShell.ps1" *>&1 | ForEach-Object { Write-Host $_ }
    $latestAdvisor = Get-ChildItem "$HOME/AzureAdvisor_*.xlsx" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($latestAdvisor) {
        Move-Item $latestAdvisor.FullName "$OutputDir/02_Advisor/AzureAdvisor.xlsx" -Force
    }
} catch {
    Write-Host "  ⚠️ Advisor script encountered an error: $($_.Exception.Message)" -ForegroundColor Yellow
}

# ═══════════════════════════════════════════════════════════════════════════════
# PHASE 3: Metrics (Right-Sizing)
# ═══════════════════════════════════════════════════════════════════════════════
if (-not $SkipMetrics) {
    Write-Host "`n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkCyan
    Write-Host " PHASE 3/4: Metrics & Right-Sizing (this takes longer)" -ForegroundColor Cyan
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkCyan

    try {
        & "$scriptDir/Invoke-AzureMetrics-CloudShell.ps1" *>&1 | ForEach-Object { Write-Host $_ }
        $latestMetrics = Get-ChildItem "$HOME/AzureMetrics_*.xlsx" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($latestMetrics) {
            Move-Item $latestMetrics.FullName "$OutputDir/03_Metrics/AzureMetrics.xlsx" -Force
        }
    } catch {
        Write-Host "  ⚠️ Metrics script encountered an error: $($_.Exception.Message)" -ForegroundColor Yellow
    }
} else {
    Write-Host "`n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkYellow
    Write-Host " PHASE 3/4: Metrics - SKIPPED (use -SkipMetrics:$false to include)" -ForegroundColor DarkYellow
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkYellow
}

# ═══════════════════════════════════════════════════════════════════════════════
# PHASE 4: Governance Visualization
# ═══════════════════════════════════════════════════════════════════════════════
Write-Host "`n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkCyan
Write-Host " PHASE 4/4: Governance Visualization" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkCyan

try {
    & "$scriptDir/Invoke-AzureGovernanceViz-CloudShell.ps1" *>&1 | ForEach-Object { Write-Host $_ }
    $latestGov = Get-ChildItem "$HOME/AzGovViz_Lite_*" -Directory -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($latestGov) {
        Get-ChildItem $latestGov.FullName | Move-Item -Destination "$OutputDir/04_Governance/" -Force
        Remove-Item $latestGov.FullName -Force -Recurse -ErrorAction SilentlyContinue
    }
} catch {
    Write-Host "  ⚠️ Governance script encountered an error: $($_.Exception.Message)" -ForegroundColor Yellow
}

# ═══════════════════════════════════════════════════════════════════════════════
# SUMMARY
# ═══════════════════════════════════════════════════════════════════════════════
$elapsed = (Get-Date) - $startTime

Write-Host @"

╔═══════════════════════════════════════════════════════════════╗
║              ✅ ALL PHASES COMPLETE                            ║
╠═══════════════════════════════════════════════════════════════╣
║  Duration: $([math]::Round($elapsed.TotalMinutes, 1)) minutes                                       ║
║  Output:   $OutputDir   ║
╠═══════════════════════════════════════════════════════════════╣
║  📁 01_Discovery/    - Resource inventory (Excel)             ║
║  📁 02_Advisor/      - Advisor recommendations (Excel)        ║
║  📁 03_Metrics/      - Right-sizing analysis (Excel)          ║
║  📁 04_Governance/   - Governance report (HTML + Excel)       ║
║  📁 05_Dashboard/    - (Run Python agent next)                ║
╠═══════════════════════════════════════════════════════════════╣
║                                                               ║
║  NEXT STEP: Generate the consolidated dashboard:              ║
║                                                               ║
║    python3 generate-dashboard.py "$OutputDir"                 ║
║                                                               ║
╚═══════════════════════════════════════════════════════════════╝
"@ -ForegroundColor Green
