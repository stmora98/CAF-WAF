<#
.SYNOPSIS
    Azure Governance Visualizer (Lite) - Single paste-and-run script for Cloud Shell.
    Generates an interactive HTML report similar to AzGovViz with hierarchy map,
    tenant summary, and scope insights.

.DESCRIPTION
    No service principal required. Uses logged-in user context.
    Outputs:
      - Interactive HTML file (AzGovViz-style with 4 panels)
      - Excel workbook with raw data

.EXAMPLE
    # Paste into Azure Cloud Shell (PowerShell) and press Enter
    # Download output files from ~/AzGovViz_Lite_<timestamp>/

.NOTES
    Prerequisites: Reader role at Management Group or Subscription scope.
#>

# ═══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════════════

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$_baseDir = if ($env:AZWORKSHOP_OUTPUT) { $env:AZWORKSHOP_OUTPUT } else { "$HOME/AzGovViz_Lite_$timestamp" }
$outputDir = $_baseDir
$htmlFile = "$outputDir/AzureGovernance.html"
$excelFile = "$outputDir/AzureGovernance.xlsx"
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

# ═══════════════════════════════════════════════════════════════════════════════
# MODULES
# ═══════════════════════════════════════════════════════════════════════════════

if (-not (Get-Module -ListAvailable -Name ImportExcel)) { Install-Module ImportExcel -Scope CurrentUser -Force }
Import-Module ImportExcel
Import-Module Az.Resources
Import-Module Az.ResourceGraph

$context = Get-AzContext
$tenantId = $context.Tenant.Id

Write-Host "`n╔══════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║  Azure Governance Visualizer (Lite)          ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host "  Tenant:  $tenantId" -ForegroundColor Gray
Write-Host "  Account: $($context.Account.Id)" -ForegroundColor Gray
Write-Host "  Output:  $outputDir`n" -ForegroundColor Yellow

# ═══════════════════════════════════════════════════════════════════════════════
# HELPER FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

function Run-ARG {
    param([string]$Query)
    $all = @(); $skip = $null
    do {
        $p = @{ Query = $Query; First = 1000 }
        if ($skip) { $p['SkipToken'] = $skip }
        $r = Search-AzGraph @p
        $all += $r.Data
        $skip = $r.SkipToken
    } while ($skip)
    return $all
}

function Export-Sheet {
    param([object[]]$Data, [string]$Sheet)
    if (-not $Data -or $Data.Count -eq 0) { $Data = @([PSCustomObject]@{ Result = "No data found" }) }
    $Data | Export-Excel -Path $excelFile -WorksheetName $Sheet -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle Medium6
}

# ═══════════════════════════════════════════════════════════════════════════════
# DATA COLLECTION
# ═══════════════════════════════════════════════════════════════════════════════

Write-Host "📊 Collecting data..." -ForegroundColor Green

# ─── Management Group Hierarchy ──────────────────────────────────────────────
Write-Host "  [1/12] Management Group Hierarchy..." -ForegroundColor Gray

$mgData = @()
$mgPolicyCount = @{}
$mgRoleCount = @{}

function Collect-MGHierarchy {
    param([string]$GroupName, [int]$Level = 0, [string]$ParentPath = "")
    try {
        $mg = Get-AzManagementGroup -GroupName $GroupName -Expand -ErrorAction Stop
        $currentPath = if ($ParentPath) { "$ParentPath/$($mg.DisplayName)" } else { $mg.DisplayName }
        
        $childMGs = @($mg.Children | Where-Object { $_.Type -match 'managementGroups' })
        $childSubs = @($mg.Children | Where-Object { $_.Type -match 'subscriptions' })

        $script:mgData += [PSCustomObject]@{
            Level         = $Level
            DisplayName   = $mg.DisplayName
            Id            = $mg.Name
            Path          = $currentPath
            ChildMGs      = $childMGs.Count
            Subscriptions = $childSubs.Count
            SubNames      = ($childSubs | ForEach-Object { $_.DisplayName }) -join "; "
            ParentId      = if ($ParentPath) { $GroupName } else { "" }
        }

        foreach ($child in $childMGs) {
            Collect-MGHierarchy -GroupName $child.Name -Level ($Level + 1) -ParentPath $currentPath
        }
    } catch {
        Write-Host "    ⚠️ Cannot expand MG: $GroupName ($($_.Exception.Message))" -ForegroundColor DarkYellow
    }
}

try {
    Collect-MGHierarchy -GroupName $tenantId
} catch {
    Write-Host "    Falling back to ARG for MG data..." -ForegroundColor DarkYellow
    $mgData = Run-ARG -Query '
    resourcecontainers
    | where type == "microsoft.management/managementgroups"
    | extend displayName = tostring(properties.displayName),
             parent = tostring(properties.details.parent.id)
    | project name, displayName, parent'
}
Export-Sheet -Data $mgData -Sheet "MgmtGroups"

# ─── Subscriptions ───────────────────────────────────────────────────────────
Write-Host "  [2/12] Subscriptions..." -ForegroundColor Gray

$subscriptions = Run-ARG -Query '
resourcecontainers
| where type == "microsoft.resources/subscriptions"
| extend state = tostring(properties.state),
         quotaId = tostring(properties.subscriptionPolicies.quotaId),
         mgParent = tostring(properties.managementGroupAncestorsChain[0].displayName)
| project subscriptionId, name, state, quotaId, mgParent, tags'
Export-Sheet -Data $subscriptions -Sheet "Subscriptions"

# ─── Policy Assignments ──────────────────────────────────────────────────────
Write-Host "  [3/12] Policy Assignments..." -ForegroundColor Gray

$policyAssignments = Run-ARG -Query '
policyresources
| where type == "microsoft.authorization/policyassignments"
| extend displayName = tostring(properties.displayName),
         enforcement = tostring(properties.enforcementMode),
         scope_ = tostring(properties.scope),
         policyDefId = tostring(properties.policyDefinitionId),
         identity_ = tostring(identity.type),
         assignedBy = tostring(properties.metadata.assignedBy)
| project name, displayName, enforcement, scope_, policyDefId, identity_, assignedBy, subscriptionId
| order by scope_ asc'
Export-Sheet -Data $policyAssignments -Sheet "PolicyAssignments"

# ─── Custom Policy Definitions ───────────────────────────────────────────────
Write-Host "  [4/12] Custom Policy Definitions..." -ForegroundColor Gray

$customPolicies = Run-ARG -Query '
policyresources
| where type == "microsoft.authorization/policydefinitions"
| where properties.policyType == "Custom"
| extend displayName = tostring(properties.displayName),
         effect = tostring(properties.policyRule.then.effect),
         category = tostring(properties.metadata.category),
         deprecated = tostring(properties.metadata.deprecated)
| project name, displayName, effect, category, deprecated, subscriptionId, id'
Export-Sheet -Data $customPolicies -Sheet "CustomPolicies"

# ─── Policy Compliance ───────────────────────────────────────────────────────
Write-Host "  [5/12] Policy Compliance..." -ForegroundColor Gray

$policyCompliance = Run-ARG -Query '
policyresources
| where type == "microsoft.policyinsights/policystates"
| where properties.complianceState != "Compliant"
| extend complianceState = tostring(properties.complianceState),
         policyAssignment = tostring(properties.policyAssignmentName),
         policyDefinition = tostring(properties.policyDefinitionName)
| summarize NonCompliantCount=count() by policyAssignment, policyDefinition, complianceState, subscriptionId
| order by NonCompliantCount desc'
Export-Sheet -Data $policyCompliance -Sheet "PolicyCompliance"

# ─── RBAC Role Assignments ───────────────────────────────────────────────────
Write-Host "  [6/12] RBAC Role Assignments..." -ForegroundColor Gray

$roleAssignments = Run-ARG -Query '
authorizationresources
| where type == "microsoft.authorization/roleassignments"
| extend principalId = tostring(properties.principalId),
         principalType = tostring(properties.principalType),
         roleDefId = tostring(properties.roleDefinitionId),
         scope_ = tostring(properties.scope),
         createdOn = tostring(properties.createdOn)
| project principalId, principalType, roleDefId, scope_, createdOn, subscriptionId'
Export-Sheet -Data $roleAssignments -Sheet "RoleAssignments"

# ─── Custom Role Definitions ─────────────────────────────────────────────────
Write-Host "  [7/12] Custom Role Definitions..." -ForegroundColor Gray

$customRoles = Run-ARG -Query '
authorizationresources
| where type == "microsoft.authorization/roledefinitions"
| where properties.type == "CustomRole"
| extend roleName = tostring(properties.roleName),
         description_ = tostring(properties.description),
         scopes = tostring(properties.assignableScopes)
| project roleName, description_, scopes, id'
Export-Sheet -Data $customRoles -Sheet "CustomRoles"

# ─── Defender for Cloud ──────────────────────────────────────────────────────
Write-Host "  [8/12] Microsoft Defender for Cloud..." -ForegroundColor Gray

$defenderPlans = Run-ARG -Query '
securityresources
| where type == "microsoft.security/pricings"
| extend tier = tostring(properties.pricingTier),
         subPlan = tostring(properties.subPlan)
| project subscriptionId, name, tier, subPlan
| order by subscriptionId asc, name asc'
Export-Sheet -Data $defenderPlans -Sheet "DefenderPlans"

$secureScores = Run-ARG -Query '
securityresources
| where type == "microsoft.security/securescores"
| extend current_ = todouble(properties.score.current),
         max_ = todouble(properties.score.max),
         pct = todouble(properties.score.percentage)
| project subscriptionId, current_, max_, pct
| order by pct asc'
Export-Sheet -Data $secureScores -Sheet "SecureScores"

# ─── Resources Summary ───────────────────────────────────────────────────────
Write-Host "  [9/12] Resources Summary..." -ForegroundColor Gray

$resourceSummary = Run-ARG -Query '
resources
| summarize Count=count() by type, location, subscriptionId
| order by Count desc'
Export-Sheet -Data $resourceSummary -Sheet "ResourceSummary"

# ─── Orphaned Resources ──────────────────────────────────────────────────────
Write-Host "  [10/12] Orphaned Resources..." -ForegroundColor Gray

$orphanedDisks = Run-ARG -Query '
resources
| where type == "microsoft.compute/disks"
| where isnull(managedBy) or managedBy == ""
| extend diskSizeGB=properties.diskSizeGB, skuName=sku.name
| project name, resourceGroup, subscriptionId, location, diskSizeGB, skuName'

$orphanedNICs = Run-ARG -Query '
resources
| where type == "microsoft.network/networkinterfaces"
| where isnull(properties.virtualMachine) and isnull(properties.privateEndpoint)
| project name, resourceGroup, subscriptionId, location'

$orphanedPIPs = Run-ARG -Query '
resources
| where type == "microsoft.network/publicipaddresses"
| where isnull(properties.ipConfiguration) and isnull(properties.natGateway)
| project name, resourceGroup, subscriptionId, location'

$orphanedNSGs = Run-ARG -Query '
resources
| where type == "microsoft.network/networksecuritygroups"
| where isnull(properties.networkInterfaces) and isnull(properties.subnets)
| project name, resourceGroup, subscriptionId, location'

$allOrphaned = @()
foreach ($d in $orphanedDisks) { $allOrphaned += [PSCustomObject]@{ Type="Disk"; Name=$d.name; ResourceGroup=$d.resourceGroup; Subscription=$d.subscriptionId; Detail="$($d.diskSizeGB) GB ($($d.skuName))" } }
foreach ($n in $orphanedNICs) { $allOrphaned += [PSCustomObject]@{ Type="NIC"; Name=$n.name; ResourceGroup=$n.resourceGroup; Subscription=$n.subscriptionId; Detail="" } }
foreach ($p in $orphanedPIPs) { $allOrphaned += [PSCustomObject]@{ Type="PublicIP"; Name=$p.name; ResourceGroup=$p.resourceGroup; Subscription=$p.subscriptionId; Detail="" } }
foreach ($s in $orphanedNSGs) { $allOrphaned += [PSCustomObject]@{ Type="NSG"; Name=$s.name; ResourceGroup=$s.resourceGroup; Subscription=$s.subscriptionId; Detail="" } }
Export-Sheet -Data $allOrphaned -Sheet "OrphanedResources"

# ─── Network Topology ────────────────────────────────────────────────────────
Write-Host "  [11/12] Network Topology..." -ForegroundColor Gray

$vnets = Run-ARG -Query '
resources
| where type == "microsoft.network/virtualnetworks"
| extend addressSpace = tostring(properties.addressSpace.addressPrefixes),
         subnets = array_length(properties.subnets),
         peerings = array_length(properties.virtualNetworkPeerings),
         ddos = properties.enableDdosProtection
| project name, resourceGroup, subscriptionId, location, addressSpace, subnets, peerings, ddos'
Export-Sheet -Data $vnets -Sheet "VNets"

# ─── Resource Locks ──────────────────────────────────────────────────────────
Write-Host "  [12/12] Resource Locks..." -ForegroundColor Gray

$locks = Run-ARG -Query '
resources
| where type == "microsoft.authorization/locks"
| extend lockLevel = tostring(properties.level)
| project name, lockLevel, resourceGroup, subscriptionId'
Export-Sheet -Data $locks -Sheet "Locks"

Write-Host "`n📊 Data collection complete. Building HTML..." -ForegroundColor Green

# ═══════════════════════════════════════════════════════════════════════════════
# HTML GENERATION
# ═══════════════════════════════════════════════════════════════════════════════

# Count policy/role assignments per MG scope for badges
$mgPolicyCounts = @{}
$mgRoleCounts = @{}
foreach ($pa in $policyAssignments) {
    $scope = $pa.scope_
    if ($scope -match '/providers/Microsoft.Management/managementGroups/([^/]+)$') {
        $mgId = $Matches[1]
        if (-not $mgPolicyCounts[$mgId]) { $mgPolicyCounts[$mgId] = 0 }
        $mgPolicyCounts[$mgId]++
    }
}
foreach ($ra in $roleAssignments) {
    $scope = $ra.scope_
    if ($scope -match '/providers/Microsoft.Management/managementGroups/([^/]+)$') {
        $mgId = $Matches[1]
        if (-not $mgRoleCounts[$mgId]) { $mgRoleCounts[$mgId] = 0 }
        $mgRoleCounts[$mgId]++
    }
}

# ─── Build Hierarchy Tree HTML ────────────────────────────────────────────────

function Build-MGTreeHTML {
    param([string]$MGId, [int]$Level = 0)
    
    $mg = $mgData | Where-Object { $_.Id -eq $MGId }
    if (-not $mg) { return "" }
    
    $pCount = if ($mgPolicyCounts[$MGId]) { $mgPolicyCounts[$MGId] } else { 0 }
    $rCount = if ($mgRoleCounts[$MGId]) { $mgRoleCounts[$MGId] } else { 0 }
    $subCount = $mg.Subscriptions
    
    $childMGs = $mgData | Where-Object { $_.Level -eq ($Level + 1) -and $_.Path -match [regex]::Escape($mg.Path) -and $_.Id -ne $MGId }
    
    $html = @"
<li>
  <a class="mg-node" href="#scope_$MGId" title="$($mg.DisplayName) ($MGId)">
    <div class="node-main">
      <div class="badge-left">
        <span class="badge badge-policy" title="$pCount Policy assignments">$pCount</span>
      </div>
      <div class="node-icon">
        <svg width="20" height="20" viewBox="0 0 18 18"><path d="M9 1L1 5v3h16V5L9 1zM2 9v6l7 3 7-3V9H2z" fill="#0078D4" opacity="0.8"/></svg>
      </div>
      <div class="badge-right">
        <span class="badge badge-rbac" title="$rCount Role assignments">$rCount</span>
      </div>
    </div>
    <div class="node-label">$($mg.DisplayName)</div>
  </a>
"@
    
    # Children
    $directChildren = $mgData | Where-Object { $_.Level -eq ($Level + 1) -and $_.Path.StartsWith("$($mg.Path)/") }
    
    if ($directChildren.Count -gt 0 -or $subCount -gt 0) {
        $html += "  <ul>`n"
        foreach ($child in $directChildren) {
            $html += Build-MGTreeHTML -MGId $child.Id -Level ($Level + 1)
        }
        if ($subCount -gt 0) {
            $html += @"
    <li class="sub-leaf">
      <a class="sub-node" href="#scope_${MGId}_subs" title="$subCount Subscriptions">
        <div class="sub-icon">
          <svg width="16" height="16" viewBox="0 0 18 18"><rect x="2" y="2" width="14" height="14" rx="2" fill="#FFB900" opacity="0.8"/><path d="M5 7h8M5 9h8M5 11h6" stroke="#fff" stroke-width="1.2"/></svg>
          <span class="sub-count">${subCount}x</span>
        </div>
      </a>
    </li>
"@
        }
        $html += "  </ul>`n"
    }
    
    $html += "</li>`n"
    return $html
}

$rootMG = $mgData | Where-Object { $_.Level -eq 0 } | Select-Object -First 1
$hierarchyTreeHTML = ""
if ($rootMG) {
    $hierarchyTreeHTML = Build-MGTreeHTML -MGId $rootMG.Id -Level 0
}

# ─── Build TenantSummary Tables HTML ─────────────────────────────────────────

function Build-TableHTML {
    param([object[]]$Data, [string]$TableId, [string]$Title, [string]$Icon = "fa-table")
    
    if (-not $Data -or $Data.Count -eq 0) {
        return @"
<button type="button" class="collapsible">
  <i class="fa $Icon"></i> <span>$Title (0)</span>
</button>
<div class="content-section"><p class="no-data">No data found</p></div>
"@
    }
    
    $columns = $Data[0].PSObject.Properties.Name
    $headerCells = ($columns | ForEach-Object { "<th>$_</th>" }) -join ""
    
    $rows = ""
    foreach ($item in $Data) {
        $cells = ($columns | ForEach-Object { "<td>$($item.$_)</td>" }) -join ""
        $rows += "    <tr>$cells</tr>`n"
    }
    
    return @"
<button type="button" class="collapsible">
  <i class="fa $Icon"></i> <span>$Title ($($Data.Count))</span>
</button>
<div class="content-section">
  <div class="table-controls">
    <input type="text" class="table-search" onkeyup="filterTable('$TableId', this)" placeholder="🔍 Filter...">
    <button class="btn-csv" onclick="exportCSV('$TableId')">📥 CSV</button>
  </div>
  <table id="$TableId" class="data-table">
    <thead><tr>$headerCells</tr></thead>
    <tbody>
$rows
    </tbody>
  </table>
</div>
"@
}

# Policy section tables
$policySummaryHTML = Build-TableHTML -Data $policyAssignments -TableId "tblPolicyAssignments" -Title "Policy Assignments" -Icon "fa-check-square"
$policySummaryHTML += Build-TableHTML -Data $customPolicies -TableId "tblCustomPolicies" -Title "Custom Policy Definitions" -Icon "fa-file-text"
$policySummaryHTML += Build-TableHTML -Data $policyCompliance -TableId "tblCompliance" -Title "Non-Compliant Resources" -Icon "fa-exclamation-triangle"

# RBAC section tables
$rbacSummaryHTML = Build-TableHTML -Data $roleAssignments -TableId "tblRoleAssignments" -Title "Role Assignments" -Icon "fa-users"
$rbacSummaryHTML += Build-TableHTML -Data $customRoles -TableId "tblCustomRoles" -Title "Custom Role Definitions" -Icon "fa-key"

# Security section
$securityHTML = Build-TableHTML -Data $defenderPlans -TableId "tblDefender" -Title "Microsoft Defender for Cloud Plans" -Icon "fa-shield"
$securityHTML += Build-TableHTML -Data $secureScores -TableId "tblSecureScore" -Title "Secure Score by Subscription" -Icon "fa-tachometer"

# Resources section
$resourcesHTML = Build-TableHTML -Data $resourceSummary -TableId "tblResources" -Title "Resources by Type & Location" -Icon "fa-cubes"
$resourcesHTML += Build-TableHTML -Data $allOrphaned -TableId "tblOrphaned" -Title "Orphaned Resources (Cost Savings)" -Icon "fa-trash"

# Network section
$networkHTML = Build-TableHTML -Data $vnets -TableId "tblVNets" -Title "Virtual Networks" -Icon "fa-sitemap"

# Governance section
$governanceHTML = Build-TableHTML -Data $locks -TableId "tblLocks" -Title "Resource Locks" -Icon "fa-lock"

# ─── Build ScopeInsights HTML ─────────────────────────────────────────────────

$scopeInsightsHTML = ""
foreach ($mg in $mgData) {
    $mgId = $mg.Id
    $mgPolicies = $policyAssignments | Where-Object { $_.scope_ -match $mgId }
    $mgRoles = $roleAssignments | Where-Object { $_.scope_ -match $mgId }
    
    $scopeInsightsHTML += @"
<button type="button" class="collapsible scope-mg" id="scope_$mgId">
  <span class="scope-indent">$("&mdash;" * $mg.Level)</span>
  <svg width="16" height="16" viewBox="0 0 18 18" style="vertical-align:middle"><path d="M9 1L1 5v3h16V5L9 1zM2 9v6l7 3 7-3V9H2z" fill="#0078D4" opacity="0.8"/></svg>
  <strong>$($mg.DisplayName)</strong> <span class="scope-id">($mgId)</span>
  <span class="scope-badges">
    <span class="badge badge-sm policy">$($mgPolicies.Count) policies</span>
    <span class="badge badge-sm rbac">$($mgRoles.Count) roles</span>
    <span class="badge badge-sm subs">$($mg.Subscriptions) subs</span>
  </span>
</button>
<div class="content-section scope-detail">
  <table class="scope-info-table">
    <tr><td><strong>Path:</strong></td><td>$($mg.Path)</td></tr>
    <tr><td><strong>Child MGs:</strong></td><td>$($mg.ChildMGs)</td></tr>
    <tr><td><strong>Subscriptions:</strong></td><td>$($mg.Subscriptions)</td></tr>
  </table>
"@
    
    if ($mgPolicies.Count -gt 0) {
        $scopeInsightsHTML += "  <h4>Policy Assignments at this scope ($($mgPolicies.Count))</h4>`n  <table class='data-table mini'><thead><tr><th>Name</th><th>Enforcement</th><th>Identity</th></tr></thead><tbody>`n"
        foreach ($p in $mgPolicies | Select-Object -First 50) {
            $scopeInsightsHTML += "    <tr><td>$($p.displayName)</td><td>$($p.enforcement)</td><td>$($p.identity_)</td></tr>`n"
        }
        $scopeInsightsHTML += "  </tbody></table>`n"
    }
    
    $scopeInsightsHTML += "</div>`n"
}

# ─── Assemble Final HTML ──────────────────────────────────────────────────────

$totalResources = ($resourceSummary | Measure-Object -Property Count -Sum).Sum
$totalPolicies = $policyAssignments.Count
$totalRoles = $roleAssignments.Count
$totalOrphaned = $allOrphaned.Count

$fullHTML = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Azure Governance Report - $(Get-Date -Format "yyyy-MM-dd")</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
<style>
/* ═══ Base ═══ */
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
    font-family: "Segoe UI", "Helvetica Neue", Arial, sans-serif;
    font-size: 12px;
    display: flex; flex-direction: column;
    height: 100vh; overflow: hidden;
    background: #eee;
}

/* ═══ Panel Layout ═══ */
.panel {
    overflow-y: auto; padding: 16px;
    resize: vertical; min-height: 100px;
    border-bottom: 2px solid #ccc;
}
.panel-hierarchy { background: #fff; flex: 2; }
.panel-summary { background: #e0f2ff; flex: 3; }
.panel-scope { background: #eee; flex: 2; }

.panel-header {
    font-size: 16px; font-weight: 700;
    padding: 8px 16px; margin: -16px -16px 16px -16px;
    border-bottom: 1px solid rgba(0,0,0,0.1);
    position: sticky; top: -16px; z-index: 10;
}
.panel-hierarchy .panel-header { background: #fff; color: #0078D4; }
.panel-summary .panel-header { background: #e0f2ff; color: #005A9E; }
.panel-scope .panel-header { background: #eee; color: #333; }

/* ═══ Summary Stats Bar ═══ */
.stats-bar {
    display: flex; gap: 12px; flex-wrap: wrap;
    padding: 12px 0; margin-bottom: 12px;
    border-bottom: 1px solid rgba(0,0,0,0.1);
}
.stat-card {
    background: #fff; border-radius: 8px; padding: 12px 16px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1); min-width: 140px;
}
.stat-card .stat-value { font-size: 24px; font-weight: 700; color: #0078D4; }
.stat-card .stat-label { font-size: 11px; color: #666; margin-top: 2px; }

/* ═══ Hierarchy Tree ═══ */
.tree { padding: 20px; }
.tree ul { padding-top: 20px; position: relative; transition: all 0.3s; }
.tree li {
    float: left; text-align: center;
    list-style-type: none; position: relative;
    padding: 20px 8px 0 8px; transition: all 0.3s;
}
.tree li::before, .tree li::after {
    content: ''; position: absolute; top: 0; right: 50%;
    border-top: 2px solid #B9B9B7; width: 50%; height: 20px;
}
.tree li::after { right: auto; left: 50%; border-left: 2px solid #B9B9B7; }
.tree li:only-child::before, .tree li:only-child::after { display: none; }
.tree li:only-child { padding-top: 0; }
.tree li:first-child::before, .tree li:last-child::after { border: 0 none; }
.tree li:last-child::before { border-right: 2px solid #B9B9B7; border-radius: 0 5px 0 0; }
.tree li:first-child::after { border-radius: 5px 0 0 0; }
.tree ul ul::before {
    content: ''; position: absolute; top: 0; left: 50%;
    border-left: 2px solid #B9B9B7; width: 0; height: 20px;
}
#first { padding-top: 0; }
#first::before, #first::after { border: 0 none; }

.mg-node {
    display: inline-block; border: 2px solid #AEDEFE;
    border-radius: 8px; padding: 6px 10px;
    text-decoration: none; color: #333; background: #f8fbff;
    transition: all 0.2s; min-width: 80px;
}
.mg-node:hover { background: #AEDEFE; transform: translateY(-2px); box-shadow: 0 3px 8px rgba(0,0,0,0.15); }
.node-main { display: flex; align-items: center; justify-content: center; gap: 6px; }
.node-label { font-size: 10px; font-weight: 600; margin-top: 4px; max-width: 120px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.badge { display: inline-block; font-size: 9px; padding: 2px 5px; border-radius: 10px; font-weight: 700; }
.badge-policy { background: #e0f2ff; color: #005A9E; }
.badge-rbac { background: #FFF3CD; color: #856404; }

.sub-leaf { }
.sub-node {
    display: inline-block; padding: 4px 8px;
    border: 1px solid #FFB900; border-radius: 4px;
    text-decoration: none; color: #333; background: #FFFDF5;
    font-size: 10px;
}
.sub-icon { display: flex; align-items: center; gap: 4px; }
.sub-count { font-weight: 700; }

/* ═══ Section Headers ═══ */
.section-header {
    font-size: 14px; font-weight: 700; margin: 16px 0 8px 0;
    padding: 8px 12px; border-radius: 4px; color: #fff;
}
.section-policy { background: #0078D4; }
.section-rbac { background: #107C10; }
.section-security { background: #D83B01; }
.section-resources { background: #5C2D91; }
.section-network { background: #008272; }
.section-governance { background: #4A4A4A; }

/* ═══ Collapsible ═══ */
.collapsible {
    background: #f5f5f5; color: #333;
    cursor: pointer; padding: 10px 14px;
    width: 100%; border: 1px solid #ddd;
    border-radius: 4px; text-align: left;
    font-size: 12px; font-weight: 600;
    margin: 4px 0; transition: background 0.2s;
    display: flex; align-items: center; gap: 8px;
}
.collapsible:hover { background: #e8e8e8; }
.collapsible::before { content: '▶'; font-size: 9px; transition: transform 0.2s; }
.collapsible.active::before { transform: rotate(90deg); }
.content-section { display: none; padding: 8px 12px; border-left: 3px solid #0078D4; margin: 0 0 8px 6px; }

/* ═══ Data Tables ═══ */
.table-controls { display: flex; gap: 8px; margin-bottom: 8px; align-items: center; }
.table-search {
    padding: 6px 10px; border: 1px solid #ccc; border-radius: 4px;
    font-size: 11px; width: 250px;
}
.btn-csv {
    padding: 5px 10px; border: 1px solid #0078D4; border-radius: 4px;
    background: #fff; color: #0078D4; cursor: pointer; font-size: 11px;
}
.btn-csv:hover { background: #0078D4; color: #fff; }

.data-table {
    width: 100%; border-collapse: collapse; font-size: 11px; margin-bottom: 12px;
}
.data-table th {
    background: #f0f0f0; padding: 6px 8px; border: 1px solid #ddd;
    font-weight: 700; text-align: left; position: sticky; top: 0;
    cursor: pointer;
}
.data-table th:hover { background: #e0e0e0; }
.data-table td { padding: 5px 8px; border: 1px solid #eee; }
.data-table tbody tr:nth-child(even) { background: #fafafa; }
.data-table tbody tr:hover { background: #e8f4fd; }
.data-table.mini { font-size: 10px; }
.data-table.mini th { padding: 4px 6px; }
.data-table.mini td { padding: 3px 6px; }
.no-data { color: #888; font-style: italic; padding: 8px; }

/* ═══ Scope Insights ═══ */
.scope-mg {
    border-left: 3px solid #0078D4;
}
.scope-indent { color: #aaa; margin-right: 4px; }
.scope-id { color: #888; font-weight: normal; font-size: 10px; }
.scope-badges { margin-left: auto; }
.badge-sm { font-size: 9px; padding: 2px 6px; border-radius: 8px; margin-left: 4px; }
.badge-sm.policy { background: #e0f2ff; color: #005A9E; }
.badge-sm.rbac { background: #E6F4E1; color: #107C10; }
.badge-sm.subs { background: #FFF3CD; color: #856404; }
.scope-detail { border-left-color: #AEDEFE; }
.scope-info-table td { padding: 3px 8px; }

/* ═══ Footer ═══ */
.footer { padding: 8px 16px; background: #333; color: #aaa; font-size: 10px; text-align: center; }
</style>
</head>
<body>

<!-- ═══ PANEL 1: HIERARCHY MAP ═══ -->
<div class="panel panel-hierarchy">
  <div class="panel-header">
    <i class="fa fa-sitemap"></i> HierarchyMap
    <span style="font-size:11px;font-weight:normal;margin-left:12px;">Management Group hierarchy with Policy/RBAC assignment counts</span>
  </div>
  <div class="tree">
    <ul>
      <li id="first">
$hierarchyTreeHTML
      </li>
    </ul>
  </div>
</div>

<!-- ═══ PANEL 2: TENANT SUMMARY ═══ -->
<div class="panel panel-summary">
  <div class="panel-header">
    <i class="fa fa-dashboard"></i> TenantSummary
  </div>
  
  <div class="stats-bar">
    <div class="stat-card"><div class="stat-value">$($subscriptions.Count)</div><div class="stat-label">Subscriptions</div></div>
    <div class="stat-card"><div class="stat-value">$($mgData.Count)</div><div class="stat-label">Management Groups</div></div>
    <div class="stat-card"><div class="stat-value">$totalResources</div><div class="stat-label">Total Resources</div></div>
    <div class="stat-card"><div class="stat-value">$totalPolicies</div><div class="stat-label">Policy Assignments</div></div>
    <div class="stat-card"><div class="stat-value">$totalRoles</div><div class="stat-label">Role Assignments</div></div>
    <div class="stat-card"><div class="stat-value">$($customPolicies.Count)</div><div class="stat-label">Custom Policies</div></div>
    <div class="stat-card"><div class="stat-value">$($customRoles.Count)</div><div class="stat-label">Custom Roles</div></div>
    <div class="stat-card"><div class="stat-value">$totalOrphaned</div><div class="stat-label">Orphaned Resources</div></div>
  </div>

  <div class="section-header section-policy"><i class="fa fa-check-square"></i> Policy</div>
  $policySummaryHTML

  <div class="section-header section-rbac"><i class="fa fa-users"></i> RBAC</div>
  $rbacSummaryHTML

  <div class="section-header section-security"><i class="fa fa-shield"></i> Security</div>
  $securityHTML

  <div class="section-header section-resources"><i class="fa fa-cubes"></i> Resources</div>
  $resourcesHTML

  <div class="section-header section-network"><i class="fa fa-sitemap"></i> Network</div>
  $networkHTML

  <div class="section-header section-governance"><i class="fa fa-lock"></i> Governance</div>
  $governanceHTML
</div>

<!-- ═══ PANEL 3: SCOPE INSIGHTS ═══ -->
<div class="panel panel-scope">
  <div class="panel-header">
    <i class="fa fa-th-list"></i> ScopeInsights
    <span style="font-size:11px;font-weight:normal;margin-left:12px;">Per Management Group detail</span>
  </div>
  $scopeInsightsHTML
</div>

<!-- ═══ FOOTER ═══ -->
<div class="footer">
  Azure Governance Report (Lite) | Generated $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") | Tenant: $tenantId | Account: $($context.Account.Id)
</div>

<!-- ═══ JAVASCRIPT ═══ -->
<script>
// Collapsible toggle
document.querySelectorAll('.collapsible').forEach(function(btn) {
    btn.addEventListener('click', function() {
        this.classList.toggle('active');
        var content = this.nextElementSibling;
        if (content && content.classList.contains('content-section')) {
            content.style.display = content.style.display === 'block' ? 'none' : 'block';
        }
        if (content && content.classList.contains('scope-detail')) {
            content.style.display = content.style.display === 'block' ? 'none' : 'block';
        }
    });
});

// Table filter
function filterTable(tableId, input) {
    var filter = input.value.toUpperCase();
    var table = document.getElementById(tableId);
    var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
    for (var i = 0; i < rows.length; i++) {
        var text = rows[i].textContent || rows[i].innerText;
        rows[i].style.display = text.toUpperCase().indexOf(filter) > -1 ? '' : 'none';
    }
}

// Table sort
document.querySelectorAll('.data-table th').forEach(function(th) {
    th.addEventListener('click', function() {
        var table = th.closest('table');
        var tbody = table.querySelector('tbody');
        var rows = Array.from(tbody.querySelectorAll('tr'));
        var idx = Array.from(th.parentNode.children).indexOf(th);
        var asc = th.dataset.sort !== 'asc';
        th.dataset.sort = asc ? 'asc' : 'desc';
        rows.sort(function(a, b) {
            var aVal = a.children[idx].textContent.trim();
            var bVal = b.children[idx].textContent.trim();
            var aNum = parseFloat(aVal), bNum = parseFloat(bVal);
            if (!isNaN(aNum) && !isNaN(bNum)) return asc ? aNum - bNum : bNum - aNum;
            return asc ? aVal.localeCompare(bVal) : bVal.localeCompare(aVal);
        });
        rows.forEach(function(row) { tbody.appendChild(row); });
    });
});

// CSV export
function exportCSV(tableId) {
    var table = document.getElementById(tableId);
    var rows = table.querySelectorAll('tr');
    var csv = [];
    rows.forEach(function(row) {
        var cols = row.querySelectorAll('td, th');
        var rowData = [];
        cols.forEach(function(col) { rowData.push('"' + col.textContent.replace(/"/g, '""') + '"'); });
        csv.push(rowData.join(','));
    });
    var blob = new Blob([csv.join('\n')], { type: 'text/csv' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = tableId + '.csv'; a.click();
    URL.revokeObjectURL(url);
}
</script>
</body>
</html>
"@

# Write HTML file
$fullHTML | Set-Content -Path $htmlFile -Encoding UTF8

# ═══════════════════════════════════════════════════════════════════════════════
# DONE
# ═══════════════════════════════════════════════════════════════════════════════

Write-Host "`n╔══════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║  ✅ Governance Report Complete!               ║" -ForegroundColor Green
Write-Host "╚══════════════════════════════════════════════╝" -ForegroundColor Green
Write-Host @"

📁 Output directory: $outputDir
   • AzureGovernance.html  - Interactive HTML report (AzGovViz-style)
   • AzureGovernance.xlsx  - Raw data Excel workbook

📊 HTML Report Panels:
   1. HierarchyMap    - MG tree with policy/RBAC badges
   2. TenantSummary   - Policy, RBAC, Security, Resources, Network, Governance
   3. ScopeInsights   - Per-MG detail with drill-down

💡 In Cloud Shell, use the download icon to get the files.
   The HTML file is self-contained - just open in any browser.
"@ -ForegroundColor Cyan
