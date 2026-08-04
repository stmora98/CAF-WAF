<#
  Azure Advisor Full Export - Paste into Cloud Shell (PowerShell)
  Exports all Advisor recommendations to Excel grouped by pillar.
#>

$_outDir = if ($env:AZWORKSHOP_OUTPUT) { $env:AZWORKSHOP_OUTPUT } else { $HOME }
$outputFile = Join-Path $_outDir "AzureAdvisor_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel

function Run-Query {
    param([string]$Query, [string]$Sheet)
    $all = @(); $skip = $null
    do {
        $p = @{ Query = $Query; First = 1000 }
        if ($skip) { $p['SkipToken'] = $skip }
        $r = Search-AzGraph @p
        $all += $r.Data
        $skip = $r.SkipToken
    } while ($skip)
    if ($all.Count -eq 0) { $all = @([PSCustomObject]@{ Result = "No recommendations" }) }
    $all | Export-Excel -Path $outputFile -WorksheetName $Sheet -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle Medium6
    Write-Host "  [$Sheet] $($all.Count) rows" -ForegroundColor Gray
}

Write-Host "`n=== Azure Advisor Export ===" -ForegroundColor Cyan

# All recommendations (overview)
Run-Query -Sheet "AllRecommendations" -Query '
advisorresources
| where type == "microsoft.advisor/recommendations"
| extend category = tostring(properties.category),
         impact = tostring(properties.impact),
         impactedType = tostring(properties.impactedField),
         impactedResource = tostring(properties.impactedValue),
         problem = tostring(properties.shortDescription.problem),
         solution = tostring(properties.shortDescription.solution)
| project category, impact, impactedType, impactedResource, problem, solution,
          resourceGroup, subscriptionId
| order by category asc, impact desc'

# Reliability
Run-Query -Sheet "Reliability" -Query '
advisorresources
| where type == "microsoft.advisor/recommendations"
| where properties.category == "HighAvailability"
| extend impact = tostring(properties.impact),
         impactedType = tostring(properties.impactedField),
         impactedResource = tostring(properties.impactedValue),
         problem = tostring(properties.shortDescription.problem),
         solution = tostring(properties.shortDescription.solution)
| project impact, impactedType, impactedResource, problem, solution,
          resourceGroup, subscriptionId
| order by impact desc'

# Security
Run-Query -Sheet "Security" -Query '
advisorresources
| where type == "microsoft.advisor/recommendations"
| where properties.category == "Security"
| extend impact = tostring(properties.impact),
         impactedType = tostring(properties.impactedField),
         impactedResource = tostring(properties.impactedValue),
         problem = tostring(properties.shortDescription.problem),
         solution = tostring(properties.shortDescription.solution)
| project impact, impactedType, impactedResource, problem, solution,
          resourceGroup, subscriptionId
| order by impact desc'

# Cost (with savings estimates)
Run-Query -Sheet "Cost" -Query '
advisorresources
| where type == "microsoft.advisor/recommendations"
| where properties.category == "Cost"
| extend impact = tostring(properties.impact),
         impactedType = tostring(properties.impactedField),
         impactedResource = tostring(properties.impactedValue),
         problem = tostring(properties.shortDescription.problem),
         solution = tostring(properties.shortDescription.solution),
         savingsAmount = tostring(properties.extendedProperties.savingsAmount),
         savingsCurrency = tostring(properties.extendedProperties.savingsCurrency),
         annualSavings = tostring(properties.extendedProperties.annualSavingsAmount)
| project impact, impactedType, impactedResource, problem, solution,
          savingsAmount, savingsCurrency, annualSavings,
          resourceGroup, subscriptionId
| order by impact desc'

# Operational Excellence
Run-Query -Sheet "OperationalExcellence" -Query '
advisorresources
| where type == "microsoft.advisor/recommendations"
| where properties.category == "OperationalExcellence"
| extend impact = tostring(properties.impact),
         impactedType = tostring(properties.impactedField),
         impactedResource = tostring(properties.impactedValue),
         problem = tostring(properties.shortDescription.problem),
         solution = tostring(properties.shortDescription.solution)
| project impact, impactedType, impactedResource, problem, solution,
          resourceGroup, subscriptionId
| order by impact desc'

# Performance
Run-Query -Sheet "Performance" -Query '
advisorresources
| where type == "microsoft.advisor/recommendations"
| where properties.category == "Performance"
| extend impact = tostring(properties.impact),
         impactedType = tostring(properties.impactedField),
         impactedResource = tostring(properties.impactedValue),
         problem = tostring(properties.shortDescription.problem),
         solution = tostring(properties.shortDescription.solution)
| project impact, impactedType, impactedResource, problem, solution,
          resourceGroup, subscriptionId
| order by impact desc'

# Summary by category and impact
Run-Query -Sheet "SummaryByCategory" -Query '
advisorresources
| where type == "microsoft.advisor/recommendations"
| extend category = tostring(properties.category),
         impact = tostring(properties.impact)
| summarize Count=count() by category, impact
| order by category asc, impact desc'

# Summary by resource type
Run-Query -Sheet "SummaryByResource" -Query '
advisorresources
| where type == "microsoft.advisor/recommendations"
| extend category = tostring(properties.category),
         impactedType = tostring(properties.impactedField)
| summarize Count=count() by impactedType, category
| order by Count desc'

# Reservation / Savings Plan recommendations (rate optimization) - subscription-scoped
# REST call, not Resource Graph, since these come from Microsoft.Consumption, not ARG.
Write-Host "`nChecking reservation & savings plan recommendations..." -ForegroundColor Cyan
$reservationRows = @()
try {
    $tokenResponse = Get-AzAccessToken -ResourceUrl 'https://management.azure.com/'
    $accessToken = if ($tokenResponse.Token -is [System.Security.SecureString]) {
        [System.Net.NetworkCredential]::new('', $tokenResponse.Token).Password
    } else { [string]$tokenResponse.Token }
    $headers = @{ Authorization = "Bearer $accessToken" }

    $subs = Get-AzSubscription -ErrorAction Stop
    foreach ($sub in $subs) {
        $uri = "https://management.azure.com/subscriptions/$($sub.Id)/providers/Microsoft.Consumption/reservationRecommendations?api-version=2021-10-01"
        try {
            $resp = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers -TimeoutSec 60 -ErrorAction Stop
            foreach ($item in @($resp.value)) {
                $p = $item.properties
                $reservationRows += [PSCustomObject]@{
                    Subscription      = $sub.Name
                    SubscriptionId    = $sub.Id
                    ResourceType      = $p.resourceType
                    SkuName           = $p.skuName
                    Location          = $p.location
                    Term              = $p.term
                    LookBackPeriod    = $p.lookBackPeriod
                    RecommendedQty    = $p.recommendedQuantity
                    CostWithNoRI      = $p.costWithNoReservedInstances
                    CostWithRI        = $p.totalCostWithReservedInstances
                    NetSavings        = $p.netSavings
                    Currency          = $p.currency
                    Scope             = $p.scope
                }
            }
        } catch {
            $status = try { [int]$_.Exception.Response.StatusCode } catch { 0 }
            if ($status -in @(403, 401)) {
                Write-Host "  ! No permission to read reservation recommendations for subscription $($sub.Name) (needs Cost Management Reader)." -ForegroundColor DarkYellow
            } else {
                Write-Host "  ! Reservation recommendations unavailable for subscription $($sub.Name): $($_.Exception.Message)" -ForegroundColor DarkYellow
            }
        }
    }
} catch {
    Write-Host "  ! Could not check reservation recommendations: $($_.Exception.Message)" -ForegroundColor DarkYellow
}

if ($reservationRows.Count -eq 0) {
    $reservationRows = @([PSCustomObject]@{ Result = "No reservation/savings plan recommendations found (or permission unavailable - requires Cost Management Reader)" })
}
$reservationRows | Export-Excel -Path $outputFile -WorksheetName "ReservationRecommendations" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle Medium6
Write-Host "  [ReservationRecommendations] $($reservationRows.Count) rows" -ForegroundColor Gray

Write-Host "`n✅ Advisor export complete!" -ForegroundColor Green
Write-Host "📁 File: $outputFile" -ForegroundColor Cyan
Write-Host "💡 In Cloud Shell, click download icon to get the file.`n" -ForegroundColor Yellow
