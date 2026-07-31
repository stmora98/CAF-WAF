<#
  Azure Monitor Metrics Discovery - Right-Sizing & Reliability Gaps
  Paste into Cloud Shell (PowerShell). Queries metrics per resource (slower than ARG).
  Covers last 30 days of metric data.
#>

$_outDir = if ($env:AZWORKSHOP_OUTPUT) { $env:AZWORKSHOP_OUTPUT } else { $HOME }
$outputFile = Join-Path $_outDir "AzureMetrics_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
$daysBack = 30
$startTime = (Get-Date).AddDays(-$daysBack)
$endTime = Get-Date

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel
Import-Module Az.Monitor

function Export-Sheet {
    param([object[]]$Data, [string]$Sheet)
    if ($Data.Count -eq 0) { $Data = @([PSCustomObject]@{ Result = "No data found" }) }
    $Data | Export-Excel -Path $outputFile -WorksheetName $Sheet -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle Medium6
    Write-Host "  [$Sheet] $($Data.Count) rows" -ForegroundColor Gray
}

function Get-AvgMetric {
    param([string]$ResourceId, [string]$MetricName, [string]$Aggregation = "Average")
    try {
        $metric = Get-AzMetric -ResourceId $ResourceId -MetricName $MetricName `
            -StartTime $startTime -EndTime $endTime -AggregationType $Aggregation `
            -TimeGrain 1.00:00:00 -WarningAction SilentlyContinue -ErrorAction Stop
        $values = $metric.Data | Where-Object { $null -ne $_.Average } | Select-Object -ExpandProperty Average
        if ($values.Count -gt 0) {
            return [math]::Round(($values | Measure-Object -Average).Average, 2)
        }
    } catch { }
    return $null
}

function Get-MaxMetric {
    param([string]$ResourceId, [string]$MetricName)
    try {
        $metric = Get-AzMetric -ResourceId $ResourceId -MetricName $MetricName `
            -StartTime $startTime -EndTime $endTime -AggregationType Maximum `
            -TimeGrain 1.00:00:00 -WarningAction SilentlyContinue -ErrorAction Stop
        $values = $metric.Data | Where-Object { $null -ne $_.Maximum } | Select-Object -ExpandProperty Maximum
        if ($values.Count -gt 0) {
            return [math]::Round(($values | Measure-Object -Maximum).Maximum, 2)
        }
    } catch { }
    return $null
}

Write-Host "`n=== Azure Metrics Discovery (Last $daysBack days) ===" -ForegroundColor Cyan
Write-Host "Output: $outputFile`n" -ForegroundColor Yellow

# ─── 1. VM Right-Sizing ──────────────────────────────────────────────────────
Write-Host "[1/6] VM CPU & Memory Utilization..." -ForegroundColor Green

$vms = Search-AzGraph -Query '
resources
| where type == "microsoft.compute/virtualmachines"
| where properties.extended.instanceView.powerState.displayStatus == "VM running"
| project id, name, resourceGroup, subscriptionId, location,
          vmSize=properties.hardwareProfile.vmSize
| order by name asc' -First 1000

$vmMetrics = @()
$i = 0
foreach ($vm in $vms) {
    $i++
    Write-Progress -Activity "Querying VM metrics" -Status "$i of $($vms.Count): $($vm.name)" -PercentComplete (($i/$vms.Count)*100)

    $cpuAvg = Get-AvgMetric -ResourceId $vm.id -MetricName "Percentage CPU"
    $cpuMax = Get-MaxMetric -ResourceId $vm.id -MetricName "Percentage CPU"

    $sizing = if ($null -eq $cpuAvg) { "No data" }
              elseif ($cpuAvg -lt 5)  { "Idle (<5%)" }
              elseif ($cpuAvg -lt 15) { "Underutilized (<15%)" }
              elseif ($cpuAvg -lt 80) { "Right-sized" }
              else                    { "Saturated (>80%)" }

    $vmMetrics += [PSCustomObject]@{
        Name          = $vm.name
        ResourceGroup = $vm.resourceGroup
        Subscription  = $vm.subscriptionId
        Location      = $vm.location
        VMSize        = $vm.vmSize
        AvgCPU_Pct    = $cpuAvg
        MaxCPU_Pct    = $cpuMax
        Assessment    = $sizing
    }
}
Write-Progress -Activity "Querying VM metrics" -Completed
Export-Sheet -Data $vmMetrics -Sheet "VM_RightSizing"

# ─── 2. SQL Database Utilization ─────────────────────────────────────────────
Write-Host "[2/6] SQL Database DTU/CPU Usage..." -ForegroundColor Green

$sqlDbs = Search-AzGraph -Query '
resources
| where type == "microsoft.sql/servers/databases"
| where name != "master"
| project id, name, resourceGroup, subscriptionId, location,
          skuName=sku.name, skuTier=sku.tier
| order by name asc' -First 1000

$sqlMetrics = @()
$i = 0
foreach ($db in $sqlDbs) {
    $i++
    Write-Progress -Activity "Querying SQL metrics" -Status "$i of $($sqlDbs.Count): $($db.name)" -PercentComplete (($i/$sqlDbs.Count)*100)

    $dtuAvg = Get-AvgMetric -ResourceId $db.id -MetricName "dtu_consumption_percent"
    $cpuAvg = if ($null -eq $dtuAvg) { Get-AvgMetric -ResourceId $db.id -MetricName "cpu_percent" } else { $null }
    $storageUsed = Get-AvgMetric -ResourceId $db.id -MetricName "storage_percent"

    $usageMetric = if ($null -ne $dtuAvg) { $dtuAvg } else { $cpuAvg }
    $metricType = if ($null -ne $dtuAvg) { "DTU%" } else { "CPU%" }

    $sizing = if ($null -eq $usageMetric) { "No data" }
              elseif ($usageMetric -lt 10) { "Oversized (<10%)" }
              elseif ($usageMetric -lt 30) { "Underutilized (<30%)" }
              elseif ($usageMetric -lt 80) { "Right-sized" }
              else                         { "Saturated (>80%)" }

    $sqlMetrics += [PSCustomObject]@{
        Name          = $db.name
        ResourceGroup = $db.resourceGroup
        Subscription  = $db.subscriptionId
        SKU           = $db.skuName
        Tier          = $db.skuTier
        MetricType    = $metricType
        AvgUsage_Pct  = $usageMetric
        Storage_Pct   = $storageUsed
        Assessment    = $sizing
    }
}
Write-Progress -Activity "Querying SQL metrics" -Completed
Export-Sheet -Data $sqlMetrics -Sheet "SQL_RightSizing"

# ─── 3. App Service Plan Utilization ─────────────────────────────────────────
Write-Host "[3/6] App Service Plan CPU..." -ForegroundColor Green

$plans = Search-AzGraph -Query '
resources
| where type == "microsoft.web/serverfarms"
| where sku.tier != "Free" and sku.tier != "Shared"
| project id, name, resourceGroup, subscriptionId, location,
          skuName=sku.name, skuTier=sku.tier, workers=properties.numberOfWorkers
| order by name asc' -First 1000

$planMetrics = @()
$i = 0
foreach ($plan in $plans) {
    $i++
    Write-Progress -Activity "Querying App Plan metrics" -Status "$i of $($plans.Count): $($plan.name)" -PercentComplete (($i/$plans.Count)*100)

    $cpuAvg = Get-AvgMetric -ResourceId $plan.id -MetricName "CpuPercentage"
    $memAvg = Get-AvgMetric -ResourceId $plan.id -MetricName "MemoryPercentage"

    $sizing = if ($null -eq $cpuAvg) { "No data" }
              elseif ($cpuAvg -lt 5)  { "Idle (<5%)" }
              elseif ($cpuAvg -lt 20) { "Underutilized (<20%)" }
              elseif ($cpuAvg -lt 80) { "Right-sized" }
              else                    { "Saturated (>80%)" }

    $planMetrics += [PSCustomObject]@{
        Name          = $plan.name
        ResourceGroup = $plan.resourceGroup
        Subscription  = $plan.subscriptionId
        SKU           = $plan.skuName
        Tier          = $plan.skuTier
        Workers       = $plan.workers
        AvgCPU_Pct    = $cpuAvg
        AvgMemory_Pct = $memAvg
        Assessment    = $sizing
    }
}
Write-Progress -Activity "Querying App Plan metrics" -Completed
Export-Sheet -Data $planMetrics -Sheet "AppPlan_RightSizing"

# ─── 4. Storage Account Activity ─────────────────────────────────────────────
Write-Host "[4/6] Storage Account Activity..." -ForegroundColor Green

$storageAccounts = Search-AzGraph -Query '
resources
| where type == "microsoft.storage/storageaccounts"
| project id, name, resourceGroup, subscriptionId, location, skuName=sku.name
| order by name asc' -First 1000

$storageMetrics = @()
$i = 0
foreach ($sa in $storageAccounts) {
    $i++
    Write-Progress -Activity "Querying Storage metrics" -Status "$i of $($storageAccounts.Count): $($sa.name)" -PercentComplete (($i/$storageAccounts.Count)*100)

    $transactions = Get-AvgMetric -ResourceId $sa.id -MetricName "Transactions"
    $availability = Get-AvgMetric -ResourceId $sa.id -MetricName "Availability"

    $activity = if ($null -eq $transactions) { "No data" }
                elseif ($transactions -eq 0)  { "Zero activity" }
                elseif ($transactions -lt 10) { "Minimal activity" }
                else                          { "Active" }

    $storageMetrics += [PSCustomObject]@{
        Name             = $sa.name
        ResourceGroup    = $sa.resourceGroup
        Subscription     = $sa.subscriptionId
        SKU              = $sa.skuName
        AvgDailyTxns     = $transactions
        Availability_Pct = $availability
        Assessment       = $activity
    }
}
Write-Progress -Activity "Querying Storage metrics" -Completed
Export-Sheet -Data $storageMetrics -Sheet "Storage_Activity"

# ─── 5. Diagnostic Settings Coverage ────────────────────────────────────────
Write-Host "[5/6] Diagnostic Settings Coverage (ARG)..." -ForegroundColor Green

$diagCoverage = Search-AzGraph -Query '
resources
| where type in ("microsoft.compute/virtualmachines",
                 "microsoft.web/sites",
                 "microsoft.sql/servers/databases",
                 "microsoft.network/applicationgateways",
                 "microsoft.network/azurefirewalls",
                 "microsoft.keyvault/vaults",
                 "microsoft.containerservice/managedclusters")
| project id, name, type, resourceGroup, subscriptionId, location
| order by type asc, name asc' -First 1000

$diagResults = @()
$i = 0
foreach ($res in $diagCoverage) {
    $i++
    Write-Progress -Activity "Checking Diagnostic Settings" -Status "$i of $($diagCoverage.Count): $($res.name)" -PercentComplete (($i/$diagCoverage.Count)*100)

    try {
        $diag = Get-AzDiagnosticSetting -ResourceId $res.id -ErrorAction Stop
        $hasDiag = $diag.Count -gt 0
        $destinations = if ($hasDiag) {
            ($diag | ForEach-Object {
                $d = @()
                if ($_.WorkspaceId) { $d += "LogAnalytics" }
                if ($_.StorageAccountId) { $d += "Storage" }
                if ($_.EventHubAuthorizationRuleId) { $d += "EventHub" }
                $d -join "+"
            }) -join "; "
        } else { "None" }
    } catch {
        $hasDiag = $false
        $destinations = "Error/NotSupported"
    }

    $diagResults += [PSCustomObject]@{
        Name           = $res.name
        Type           = $res.type
        ResourceGroup  = $res.resourceGroup
        Subscription   = $res.subscriptionId
        HasDiagnostics = $hasDiag
        Destinations   = $destinations
        Gap            = if (-not $hasDiag) { "No diagnostics configured" } else { "OK" }
    }
}
Write-Progress -Activity "Checking Diagnostic Settings" -Completed
Export-Sheet -Data $diagResults -Sheet "DiagnosticsCoverage"

# ─── 6. Alert Rules Coverage ────────────────────────────────────────────────
Write-Host "[6/6] Alert Rules Summary (ARG)..." -ForegroundColor Green

$alertSummary = Search-AzGraph -Query '
resources
| where type in ("microsoft.insights/metricalerts",
                 "microsoft.insights/activitylogalerts",
                 "microsoft.insights/scheduledqueryrules")
| extend enabled = properties.enabled,
         severity = properties.severity,
         targetResourceType = tostring(properties.targetResourceType),
         scopes = tostring(properties.scopes),
         description_ = tostring(properties.description)
| project name, type, resourceGroup, subscriptionId, enabled, severity,
          targetResourceType, scopes, description_
| order by type asc, name asc' -First 1000

Export-Sheet -Data $alertSummary -Sheet "AlertRules"

# ─── Done ────────────────────────────────────────────────────────────────────
Write-Host "`n✅ Metrics discovery complete!" -ForegroundColor Green
Write-Host "📁 File: $outputFile" -ForegroundColor Cyan
Write-Host @"

Tabs generated:
  1. VM_RightSizing       - CPU utilization & sizing assessment
  2. SQL_RightSizing      - DTU/CPU usage per database
  3. AppPlan_RightSizing  - App Service Plan utilization
  4. Storage_Activity     - Accounts with zero/minimal activity
  5. DiagnosticsCoverage  - Resources missing diagnostic settings
  6. AlertRules           - All configured alert rules

Note: This script queries metrics per-resource, so it takes
longer than the Resource Graph-only scripts (~1-3 min per 100 resources).
"@ -ForegroundColor Yellow
