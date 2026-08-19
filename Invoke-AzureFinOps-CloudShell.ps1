<#
  Azure FinOps Extended Export - Paste into Cloud Shell (PowerShell)
  Approximates the data FinOps Hub (FinOps Toolkit) surfaces, using only
  subscription-level Reader access (no billing-account exports required):
    - ActualCost              Real monthly spend by service & region, last 6 months (Cost Management Query API)
    - ReservationDetails      Reservation utilization (billing-account scope, best-effort)
    - Budgets                 Configured budgets vs. current spend (Consumption Budgets API)
    - UnattachedPublicIPs     Static public IPs not attached to any resource
    - StoppedVMs              VMs stopped but not deallocated (still billed for compute)
    - BackendlessAppGateways  Application Gateways with no backend pool
    - BackendlessLoadBalancers Load Balancers with no backend pool
    - EmptySqlElasticPools    SQL Elastic Pools with zero databases
    - NonSpotAKSPools         Autoscaling AKS node pools not using Spot VMs
    - VMsWithoutHybridBenefit Windows VMs not leveraging Azure Hybrid Benefit
    - SqlVMsWithoutHybridBenefit SQL VMs on pay-as-you-go licensing
#>

$_outDir = if ($env:AZWORKSHOP_OUTPUT) { $env:AZWORKSHOP_OUTPUT } else { $HOME }
$outputFile = Join-Path $_outDir "AzureFinOps_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
$subscriptionIds = @($env:AZWORKSHOP_SUBSCRIPTION_IDS -split ',' | Where-Object { $_ })
if ($subscriptionIds.Count -eq 0) {
    $subscriptionIds = @(Get-AzSubscription -ErrorAction Stop | Where-Object { $_.State -eq 'Enabled' } | Select-Object -ExpandProperty Id)
}
if ($subscriptionIds.Count -eq 0) { throw "No enabled subscriptions are accessible in the current tenant." }

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel

# ─── Resource Graph retry (handles 429/5xx throttling across large multi-subscription tenants) ──
function Invoke-SearchAzGraphWithRetry {
    param([hashtable]$Parameters, [int]$MaxAttempts = 6)
    $attempt = 0
    do {
        try {
            return Search-AzGraph @Parameters -ErrorAction Stop
        } catch {
            $attempt++
            $statusCode = try { [int]$_.Exception.Response.StatusCode } catch { 0 }
            $transient = ($statusCode -in 429, 408, 500, 502, 503, 504) -or ($_.Exception.Message -match '429|too many requests|throttl|gateway timeout|server error|service unavailable')
            if ($attempt -ge $MaxAttempts -or -not $transient) { throw }
            $delaySeconds = [math]::Min(60, [math]::Pow(2, $attempt))
            Write-Host "  Resource Graph throttled or unavailable. Retrying in $delaySeconds seconds (attempt $attempt/$MaxAttempts)..." -ForegroundColor DarkYellow
            Start-Sleep -Seconds $delaySeconds
        }
    } while ($true)
}

function Invoke-GraphQuery {
    param([string]$Query, [string]$Sheet)
    $all = @(); $skip = $null
    do {
        $p = @{ Query = $Query; Subscription = $subscriptionIds; First = 1000 }
        if ($skip) { $p['SkipToken'] = $skip }
        $r = Invoke-SearchAzGraphWithRetry -Parameters $p
        $all += $r.Data
        $skip = $r.SkipToken
    } while ($skip)
    if ($all.Count -eq 0) { $all = @([PSCustomObject]@{ Result = "No findings" }) }
    $all | Export-Excel -Path $outputFile -WorksheetName $Sheet -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle Medium6
    Write-Host "  [$Sheet] $($all.Count) rows" -ForegroundColor Gray
}

Write-Host "`n=== Azure FinOps Extended Export ===" -ForegroundColor Cyan

# --- Shared bearer token for REST calls (Cost Management / Consumption / Billing) ---
$accessToken = $null
try {
    $tokenResponse = Get-AzAccessToken -ResourceUrl 'https://management.azure.com/'
    $accessToken = if ($tokenResponse.Token -is [System.Security.SecureString]) {
        [System.Net.NetworkCredential]::new('', $tokenResponse.Token).Password
    } else { [string]$tokenResponse.Token }
} catch {
    Write-Host "  ! Could not acquire an access token: $($_.Exception.Message)" -ForegroundColor DarkYellow
}
$headers = if ($accessToken) { @{ Authorization = "Bearer $accessToken"; 'Content-Type' = 'application/json' } } else { $null }

function Invoke-RestMethodWithRetry {
    param([string]$Uri, [string]$Method, [hashtable]$Headers, [string]$Body, [int]$TimeoutSec = 60, [int]$MaxAttempts = 5)

    $attempt = 0
    do {
        try {
            return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $Headers -Body $Body -TimeoutSec $TimeoutSec -ErrorAction Stop
        } catch {
            $attempt++
            $statusCode = try { [int]$_.Exception.Response.StatusCode } catch { 0 }
            if ($attempt -ge $MaxAttempts -or $statusCode -notin @(429, 408, 500, 502, 503, 504)) { throw }
            $retryAfter = try { [int]$_.Exception.Response.Headers['Retry-After'] } catch { 0 }
            $delaySeconds = if ($retryAfter -gt 0) { $retryAfter } else { [math]::Min(30, [math]::Pow(2, $attempt)) }
            Write-Host "  Cost Management API returned HTTP $statusCode. Retrying in $delaySeconds seconds (attempt $attempt/$MaxAttempts)..." -ForegroundColor DarkYellow
            Start-Sleep -Seconds $delaySeconds
        }
    } while ($true)
}

# ═══════════════════════════════════════════════════════════════════════════════
# ActualCost - real monthly spend by service (Cost Management Query API)
# ═══════════════════════════════════════════════════════════════════════════════
Write-Host "`nQuerying actual cost (last 6 months)..." -ForegroundColor Cyan
$costRows = @()
if ($headers) {
    $from = (Get-Date).AddMonths(-6).ToString('yyyy-MM-01')
    $to = (Get-Date).ToString('yyyy-MM-dd')
    $body = @{
        type      = "ActualCost"
        timeframe = "Custom"
        timePeriod = @{ from = $from; to = $to }
        dataset   = @{
            granularity = "Monthly"
            aggregation = @{ totalCost = @{ name = "Cost"; function = "Sum" } }
            grouping    = @(
                @{ type = "Dimension"; name = "ServiceName" },
                @{ type = "Dimension"; name = "ResourceLocation" }
            )
        }
    } | ConvertTo-Json -Depth 10

    try {
        $subs = @(Get-AzSubscription -ErrorAction Stop | Where-Object { $_.Id -in $subscriptionIds })
        foreach ($sub in $subs) {
            $uri = "https://management.azure.com/subscriptions/$($sub.Id)/providers/Microsoft.CostManagement/query?api-version=2023-11-01"
            try {
                $resp = Invoke-RestMethodWithRetry -Method POST -Uri $uri -Headers $headers -Body $body -TimeoutSec 60
                $columns = @($resp.properties.columns.name)
                foreach ($row in @($resp.properties.rows)) {
                    $record = [ordered]@{ Subscription = $sub.Name; SubscriptionId = $sub.Id }
                    for ($i = 0; $i -lt $columns.Count; $i++) { $record[$columns[$i]] = $row[$i] }
                    $costRows += [PSCustomObject]$record
                }
            } catch {
                $status = try { [int]$_.Exception.Response.StatusCode } catch { 0 }
                if ($status -in @(403, 401)) {
                    Write-Host "  ! No permission to read cost data for subscription $($sub.Name) (needs Cost Management Reader)." -ForegroundColor DarkYellow
                } else {
                    Write-Host "  ! Actual cost unavailable for subscription $($sub.Name): $($_.Exception.Message)" -ForegroundColor DarkYellow
                }
            }
        }
    } catch {
        Write-Host "  ! Could not enumerate subscriptions for cost query: $($_.Exception.Message)" -ForegroundColor DarkYellow
    }
}
if ($costRows.Count -eq 0) {
    $costRows = @([PSCustomObject]@{ Result = "No actual cost data found (or permission unavailable - requires Cost Management Reader)" })
}
$costRows | Export-Excel -Path $outputFile -WorksheetName "ActualCost" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle Medium6
Write-Host "  [ActualCost] $($costRows.Count) rows" -ForegroundColor Gray

# ═══════════════════════════════════════════════════════════════════════════════
# ReservationDetails - utilization (billing-account scope, best-effort)
# ═══════════════════════════════════════════════════════════════════════════════
Write-Host "`nChecking reservation utilization (requires billing account access)..." -ForegroundColor Cyan
$reservationDetailRows = @()
if ($headers) {
    try {
        $baUri = "https://management.azure.com/providers/Microsoft.Billing/billingAccounts?api-version=2020-05-01"
        $baResp = Invoke-RestMethodWithRetry -Uri $baUri -Method GET -Headers $headers -TimeoutSec 60
        $from = (Get-Date).AddDays(-30).ToString('yyyy-MM-dd')
        $to = (Get-Date).ToString('yyyy-MM-dd')
        foreach ($ba in @($baResp.value)) {
            $filter = "properties/UsageDate ge '$from' and properties/UsageDate le '$to'"
            $uri = "https://management.azure.com$($ba.id)/providers/Microsoft.Consumption/reservationDetails?api-version=2023-05-01&`$filter=$([uri]::EscapeDataString($filter))"
            try {
                $resp = Invoke-RestMethodWithRetry -Uri $uri -Method GET -Headers $headers -TimeoutSec 60
                foreach ($item in @($resp.value)) {
                    $p = $item.properties
                    $reservationDetailRows += [PSCustomObject]@{
                        BillingAccount   = $ba.name
                        ReservationId    = $p.reservationId
                        SkuName          = $p.skuName
                        InstanceFlexibility = $p.instanceFlexibility
                        TotalReservedQty = $p.totalReservedQuantity
                        UsedHours        = $p.usedHours
                        UtilizationPct   = $p.utilizationPercentage
                        UsageDate        = $p.usageDate
                    }
                }
            } catch {
                $status = try { [int]$_.Exception.Response.StatusCode } catch { 0 }
                if ($status -in @(403, 401)) {
                    Write-Host "  ! No permission to read reservation details for billing account $($ba.name) (needs Enterprise/Billing Reader)." -ForegroundColor DarkYellow
                } else {
                    Write-Host "  ! Reservation details unavailable for billing account $($ba.name): $($_.Exception.Message)" -ForegroundColor DarkYellow
                }
            }
        }
    } catch {
        Write-Host "  ! Could not enumerate billing accounts (needs Billing Reader at tenant scope): $($_.Exception.Message)" -ForegroundColor DarkYellow
    }
}
if ($reservationDetailRows.Count -eq 0) {
    $reservationDetailRows = @([PSCustomObject]@{ Result = "No reservation utilization data found (or permission unavailable - requires Billing Reader)" })
}
$reservationDetailRows | Export-Excel -Path $outputFile -WorksheetName "ReservationDetails" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle Medium6
Write-Host "  [ReservationDetails] $($reservationDetailRows.Count) rows" -ForegroundColor Gray

# ═══════════════════════════════════════════════════════════════════════════════
# Budgets - configured budgets vs. current spend (Consumption Budgets API, subscription scope)
# ═══════════════════════════════════════════════════════════════════════════════
Write-Host "`nChecking configured budgets..." -ForegroundColor Cyan
$budgetRows = @()
if ($headers) {
    try {
        $subs = @(Get-AzSubscription -ErrorAction Stop | Where-Object { $_.Id -in $subscriptionIds })
        foreach ($sub in $subs) {
            $uri = "https://management.azure.com/subscriptions/$($sub.Id)/providers/Microsoft.Consumption/budgets?api-version=2023-05-01"
            try {
                $resp = Invoke-RestMethodWithRetry -Uri $uri -Method GET -Headers $headers -TimeoutSec 60
                foreach ($b in @($resp.value)) {
                    $p = $b.properties
                    $budgetRows += [PSCustomObject]@{
                        Subscription   = $sub.Name
                        SubscriptionId = $sub.Id
                        BudgetName     = $b.name
                        Category       = $p.category
                        Amount         = $p.amount
                        CurrentSpend   = $p.currentSpend.amount
                        Currency       = $p.currentSpend.unit
                        TimeGrain      = $p.timeGrain
                        StartDate      = $p.timePeriod.startDate
                        EndDate        = $p.timePeriod.endDate
                    }
                }
            } catch {
                $status = try { [int]$_.Exception.Response.StatusCode } catch { 0 }
                if ($status -in @(403, 401)) {
                    Write-Host "  ! No permission to read budgets for subscription $($sub.Name) (needs Cost Management Reader)." -ForegroundColor DarkYellow
                } else {
                    Write-Host "  ! Budgets unavailable for subscription $($sub.Name): $($_.Exception.Message)" -ForegroundColor DarkYellow
                }
            }
        }
    } catch {
        Write-Host "  ! Could not enumerate subscriptions for budget query: $($_.Exception.Message)" -ForegroundColor DarkYellow
    }
}
if ($budgetRows.Count -eq 0) {
    $budgetRows = @([PSCustomObject]@{ Result = "No budgets configured (or permission unavailable - requires Cost Management Reader)" })
}
$budgetRows | Export-Excel -Path $outputFile -WorksheetName "Budgets" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle Medium6
Write-Host "  [Budgets] $($budgetRows.Count) rows" -ForegroundColor Gray

# ═══════════════════════════════════════════════════════════════════════════════
# Extended optimization recommendations (Azure Resource Graph)
# Same query set FinOps Hub runs daily against Azure Resource Graph.
# ═══════════════════════════════════════════════════════════════════════════════
Invoke-GraphQuery -Sheet "UnattachedPublicIPs" -Query '
resources
| where type =~ "microsoft.network/publicipaddresses"
| where isnull(properties.ipConfiguration)
| project name, resourceGroup, subscriptionId, location,
          sku = tostring(sku.name), allocationMethod = tostring(properties.publicIPAllocationMethod)'

Invoke-GraphQuery -Sheet "StoppedVMs" -Query '
resources
| where type =~ "microsoft.compute/virtualmachines"
| extend powerState = tostring(properties.extended.instanceView.powerState.code)
| where powerState == "PowerState/stopped"
| project name, resourceGroup, subscriptionId, location, powerState,
          vmSize = tostring(properties.hardwareProfile.vmSize)'

Invoke-GraphQuery -Sheet "BackendlessAppGateways" -Query '
resources
| where type =~ "microsoft.network/applicationgateways"
| extend backendPoolCount = array_length(properties.backendAddressPools)
| where backendPoolCount == 0
| project name, resourceGroup, subscriptionId, location,
          sku = tostring(properties.sku.name)'

Invoke-GraphQuery -Sheet "BackendlessLoadBalancers" -Query '
resources
| where type =~ "microsoft.network/loadbalancers"
| extend backendPoolCount = array_length(properties.backendAddressPools)
| where backendPoolCount == 0
| project name, resourceGroup, subscriptionId, location,
          sku = tostring(sku.name)'

Invoke-GraphQuery -Sheet "EmptySqlElasticPools" -Query '
resources
| where type =~ "microsoft.sql/servers/elasticpools"
| project poolId = id, name, resourceGroup, subscriptionId, location
| join kind=leftouter (
    resources
    | where type =~ "microsoft.sql/servers/databases"
    | extend poolId = tostring(properties.elasticPoolId)
    | where isnotempty(poolId)
    | project poolId, dbName = name
) on poolId
| summarize dbCount = countif(isnotempty(dbName)) by poolId, name, resourceGroup, subscriptionId, location
| where dbCount == 0
| project-away poolId'

Invoke-GraphQuery -Sheet "NonSpotAKSPools" -Query '
resources
| where type =~ "microsoft.containerservice/managedclusters"
| mv-expand pool = properties.agentPoolProfiles
| extend poolName = tostring(pool.name),
         enableAutoScaling = tobool(pool.enableAutoScaling),
         priority = tostring(pool.scaleSetPriority)
| where enableAutoScaling == true and priority != "Spot"
| project name, poolName, resourceGroup, subscriptionId, location'

Invoke-GraphQuery -Sheet "VMsWithoutHybridBenefit" -Query '
resources
| where type =~ "microsoft.compute/virtualmachines"
| extend licenseType = tostring(properties.licenseType),
         osType = tostring(properties.storageProfile.osDisk.osType)
| where osType == "Windows" and licenseType !in ("Windows_Server", "Windows_Client")
| project name, resourceGroup, subscriptionId, location,
          vmSize = tostring(properties.hardwareProfile.vmSize)'

Invoke-GraphQuery -Sheet "SqlVMsWithoutHybridBenefit" -Query '
resources
| where type =~ "microsoft.sqlvirtualmachine/sqlvirtualmachines"
| extend licenseType = tostring(properties.sqlServerLicenseType)
| where licenseType == "PAYG"
| project name, resourceGroup, subscriptionId, location, licenseType'

Write-Host "`n✅ FinOps extended export complete!" -ForegroundColor Green
Write-Host "📁 File: $outputFile" -ForegroundColor Cyan
Write-Host "💡 In Cloud Shell, click download icon to get the file.`n" -ForegroundColor Yellow
