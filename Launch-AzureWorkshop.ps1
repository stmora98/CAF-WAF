<#
.SYNOPSIS
    Azure WAF/CAF Workshop Launcher - Runs all discovery scripts and consolidates output.

.DESCRIPTION
    Master launcher that sequentially executes:
      1. Invoke-AzureDiscovery-CloudShell.ps1       (Resource inventory)
      2. Invoke-AzureAdvisor-CloudShell.ps1         (Advisor recommendations)
      3. Invoke-AzureMetrics-CloudShell.ps1         (Right-sizing & reliability)
      4. Invoke-AzureGovernanceViz-CloudShell.ps1   (Governance + HTML report)
      5. generate-dashboard.py                      (Consolidated dashboard + action items)
    
    All outputs are consolidated into a single timestamped folder
    created alongside the scripts.

.PARAMETER OutputDir
    Base output directory. Defaults to ./AzureWorkshop_<timestamp> (same folder as scripts)

.PARAMETER SkipMetrics
    Skip the metrics script (it is slower, queries per-resource). Use if short on time.

.EXAMPLE
    ./Launch-AzureWorkshop.ps1
    ./Launch-AzureWorkshop.ps1 -SkipMetrics

.NOTES
    Run in Azure Cloud Shell (PowerShell) or locally with Az modules installed.
#>

[CmdletBinding()]
param(
    [string]$OutputDir,
    [switch]$SkipMetrics
)

$ErrorActionPreference = 'Continue'
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$scriptDir = $PSScriptRoot
if (-not $scriptDir) { $scriptDir = (Get-Location).Path }

if (-not $OutputDir) {
    $OutputDir = Join-Path $scriptDir "AzureWorkshop_$timestamp"
}

New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
New-Item -ItemType Directory -Path "$OutputDir/01_Discovery" -Force | Out-Null
New-Item -ItemType Directory -Path "$OutputDir/02_Advisor" -Force | Out-Null
New-Item -ItemType Directory -Path "$OutputDir/03_Metrics" -Force | Out-Null
New-Item -ItemType Directory -Path "$OutputDir/04_Governance" -Force | Out-Null
New-Item -ItemType Directory -Path "$OutputDir/05_Dashboard" -Force | Out-Null

# --- Ensure modules ---
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
    Write-Host "  Python not found - dashboard generation will be skipped" -ForegroundColor DarkYellow
}

Write-Host "  All prerequisites satisfied.`n" -ForegroundColor Green

$context = Get-AzContext
if (-not $context) {
    Write-Host "Not authenticated. Running Connect-AzAccount..." -ForegroundColor Yellow
    Connect-AzAccount
    $context = Get-AzContext
}

Write-Host "Azure WAF/CAF Workshop - Discovery Launcher" -ForegroundColor Cyan
Write-Host "Tenant:  $($context.Tenant.Id)" -ForegroundColor Cyan
Write-Host "Account: $($context.Account.Id)" -ForegroundColor Cyan
Write-Host "Output:  $OutputDir`n" -ForegroundColor Cyan

$startTime = Get-Date

# === PHASE 1: Resource Discovery ===
Write-Host " PHASE 1/5: Resource Discovery" -ForegroundColor Cyan
$env:AZWORKSHOP_OUTPUT = "$OutputDir/01_Discovery"
try {
    & "$scriptDir/Invoke-AzureDiscovery-CloudShell.ps1" *>&1 | ForEach-Object { Write-Host $_ }
} catch {
    Write-Host "  Discovery error: $($_.Exception.Message)" -ForegroundColor Yellow
}

# === PHASE 2: Advisor Recommendations ===
Write-Host "`n PHASE 2/5: Advisor Recommendations" -ForegroundColor Cyan
$env:AZWORKSHOP_OUTPUT = "$OutputDir/02_Advisor"
try {
    & "$scriptDir/Invoke-AzureAdvisor-CloudShell.ps1" *>&1 | ForEach-Object { Write-Host $_ }
} catch {
    Write-Host "  Advisor error: $($_.Exception.Message)" -ForegroundColor Yellow
}

# === PHASE 3: Metrics (Right-Sizing) ===
if (-not $SkipMetrics) {
    Write-Host "`n PHASE 3/5: Metrics & Right-Sizing (this takes longer)" -ForegroundColor Cyan
    $env:AZWORKSHOP_OUTPUT = "$OutputDir/03_Metrics"
    try {
        & "$scriptDir/Invoke-AzureMetrics-CloudShell.ps1" *>&1 | ForEach-Object { Write-Host $_ }
    } catch {
        Write-Host "  Metrics error: $($_.Exception.Message)" -ForegroundColor Yellow
    }
} else {
    Write-Host "`n PHASE 3/5: Metrics - SKIPPED" -ForegroundColor DarkYellow
}

# === PHASE 4: Governance Visualization ===
Write-Host "`n PHASE 4/5: Governance Visualization" -ForegroundColor Cyan
$env:AZWORKSHOP_OUTPUT = "$OutputDir/04_Governance"
try {
    & "$scriptDir/Invoke-AzureGovernanceViz-CloudShell.ps1" *>&1 | ForEach-Object { Write-Host $_ }
} catch {
    Write-Host "  Governance error: $($_.Exception.Message)" -ForegroundColor Yellow
}

# === PHASE 5: Consolidated Dashboard ===
Write-Host "`n PHASE 5/5: Generating Consolidated Dashboard" -ForegroundColor Cyan
if ($pythonCmd) {
    try {
        & $pythonCmd "$scriptDir/generate-dashboard.py" $OutputDir *>&1 | ForEach-Object { Write-Host $_ }
        Write-Host "  Dashboard generated in $OutputDir/05_Dashboard/" -ForegroundColor Green
    } catch {
        Write-Host "  Dashboard generation failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }
} else {
    Write-Host "  Python not available. Run manually: python3 generate-dashboard.py `"$OutputDir`"" -ForegroundColor Yellow
}

# Clean up env var
Remove-Item Env:\AZWORKSHOP_OUTPUT -ErrorAction SilentlyContinue

# === SUMMARY ===
$elapsed = (Get-Date) - $startTime
Write-Host "`nALL PHASES COMPLETE" -ForegroundColor Green
Write-Host "Duration: $([math]::Round($elapsed.TotalMinutes, 1)) minutes" -ForegroundColor Green
Write-Host "Output:   $OutputDir" -ForegroundColor Green
Write-Host "  01_Discovery/    - Resource inventory (Excel)" -ForegroundColor White
Write-Host "  02_Advisor/      - Advisor recommendations (Excel)" -ForegroundColor White
Write-Host "  03_Metrics/      - Right-sizing analysis (Excel)" -ForegroundColor White
Write-Host "  04_Governance/   - Governance report (HTML + Excel)" -ForegroundColor White
Write-Host "  05_Dashboard/    - Consolidated dashboard (HTML)" -ForegroundColor White
