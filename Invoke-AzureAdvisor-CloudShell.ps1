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

Write-Host "`n✅ Advisor export complete!" -ForegroundColor Green
Write-Host "📁 File: $outputFile" -ForegroundColor Cyan
Write-Host "💡 In Cloud Shell, click download icon to get the file.`n" -ForegroundColor Yellow
