<#
.SYNOPSIS
    Azure Governance Visualizer (Lite) - Single paste-and-run script for Cloud Shell.
    Collects management group hierarchy, policy, RBAC, and governance data
    (AzGovViz-style scope).

.DESCRIPTION
    No service principal required. Uses logged-in user context.
    Outputs:
      - Excel workbook with raw data (consumed by the consolidated dashboard)

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
$subscriptionIds = @($env:AZWORKSHOP_SUBSCRIPTION_IDS -split ',' | Where-Object { $_ })
if ($subscriptionIds.Count -eq 0) {
    $subscriptionIds = @(Get-AzSubscription -TenantId $tenantId -ErrorAction Stop | Where-Object { $_.State -eq 'Enabled' } | Select-Object -ExpandProperty Id)
}
if ($subscriptionIds.Count -eq 0) { throw "No enabled subscriptions are accessible in tenant $tenantId." }

Write-Host "`n╔══════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║  Azure Governance Visualizer (Lite)          ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host "  Tenant:  $tenantId" -ForegroundColor Gray
Write-Host "  Account: $($context.Account.Id)" -ForegroundColor Gray
Write-Host "  Output:  $outputDir`n" -ForegroundColor Yellow

# ═══════════════════════════════════════════════════════════════════════════════
# HELPER FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

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

function Invoke-ARGQuery {
    param([string]$Query)
    $all = @(); $skip = $null
    do {
        $p = @{ Query = $Query; Subscription = $subscriptionIds; First = 1000 }
        if ($skip) { $p['SkipToken'] = $skip }
        $r = Invoke-SearchAzGraphWithRetry -Parameters $p
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

# Computes Level/Path/ChildMGs/ParentId from a flat {name,displayName,parent} MG list (the ARG
# fallback shape) by walking each row's parent chain - the ARG query alone has no depth/path
# info, and the dashboard's hierarchy diagram requires a real Level per row to render.
function Resolve-MgLevels {
    param([object[]]$FlatMgs)
    if (-not $FlatMgs -or $FlatMgs.Count -eq 0) { return @() }
    $byId = @{}
    foreach ($m in $FlatMgs) { $byId[$m.name] = $m }
    $childCounts = @{}
    foreach ($m in $FlatMgs) {
        $parentId = if ($m.parent) { ($m.parent -split '/')[-1] } else { $null }
        if ($parentId -and $byId.ContainsKey($parentId)) {
            if (-not $childCounts.ContainsKey($parentId)) { $childCounts[$parentId] = 0 }
            $childCounts[$parentId]++
        }
    }
    $levelCache = @{}
    $pathCache = @{}
    function Resolve-MgNode {
        param([string]$Id, [int]$Guard = 0)
        if ($levelCache.ContainsKey($Id)) { return }
        $node = $byId[$Id]
        $parentId = if ($node.parent) { ($node.parent -split '/')[-1] } else { $null }
        if (-not $parentId -or -not $byId.ContainsKey($parentId) -or $Guard -gt 20) {
            $levelCache[$Id] = 0
            $pathCache[$Id] = $node.displayName
            return
        }
        Resolve-MgNode -Id $parentId -Guard ($Guard + 1)
        $levelCache[$Id] = $levelCache[$parentId] + 1
        $pathCache[$Id] = "$($pathCache[$parentId])/$($node.displayName)"
    }
    foreach ($m in $FlatMgs) { Resolve-MgNode -Id $m.name }
    return $FlatMgs | ForEach-Object {
        $parentId = if ($_.parent) { ($_.parent -split '/')[-1] } else { "" }
        [PSCustomObject]@{
            Level         = $levelCache[$_.name]
            DisplayName   = $_.displayName
            Id            = $_.name
            Path          = $pathCache[$_.name]
            ChildMGs      = if ($childCounts.ContainsKey($_.name)) { $childCounts[$_.name] } else { 0 }
            Subscriptions = 0
            SubNames      = ""
            ParentId      = $parentId
        }
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# DATA COLLECTION
# ═══════════════════════════════════════════════════════════════════════════════

Write-Host "📊 Collecting data..." -ForegroundColor Green

# ─── Management Group Hierarchy ──────────────────────────────────────────────
Write-Host "  [1/12] Management Group Hierarchy..." -ForegroundColor Gray

$mgData = @()

function Get-MGHierarchy {
    param([string]$GroupName, [int]$Level = 0, [string]$ParentPath = "")
    try {
        $mgAttempt = 0
        $mg = $null
        do {
            try {
                $mg = Get-AzManagementGroup -GroupName $GroupName -Expand -ErrorAction Stop
                break
            } catch {
                $mgAttempt++
                $statusCode = try { [int]$_.Exception.Response.StatusCode } catch { 0 }
                $transient = ($statusCode -in 429, 408, 500, 502, 503, 504) -or ($_.Exception.Message -match '429|too many requests|throttl|gateway timeout|server error|service unavailable')
                if ($mgAttempt -ge 6 -or -not $transient) { throw }
                Start-Sleep -Seconds ([math]::Min(60, [math]::Pow(2, $mgAttempt)))
            }
        } while ($true)
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
            Get-MGHierarchy -GroupName $child.Name -Level ($Level + 1) -ParentPath $currentPath
        }
    } catch {
        Write-Host "    ⚠️ Cannot expand MG: $GroupName ($($_.Exception.Message))" -ForegroundColor DarkYellow
    }
}

try {
    Get-MGHierarchy -GroupName $tenantId
} catch {
    Write-Host "    Root MG traversal threw an exception - falling back to ARG for MG data..." -ForegroundColor DarkYellow
}
if (-not $mgData -or $mgData.Count -eq 0) {
    # Get-MGHierarchy swallows its own per-node errors (see its internal catch above), so a
    # caller without Reader at the TENANT ROOT (but with Reader at a child MG) ends up here
    # with an empty $mgData instead of a thrown exception - always re-check emptiness, not
    # just rely on the try/catch, or partial MG access silently renders as "no data".
    Write-Host "    No management groups discovered from the tenant root (no access at that scope?) - falling back to ARG for MG data..." -ForegroundColor DarkYellow
    $flatMgs = Invoke-ARGQuery -Query '
    resourcecontainers
    | where type == "microsoft.management/managementgroups"
    | extend displayName = tostring(properties.displayName),
             parent = tostring(properties.details.parent.id)
    | project name, displayName, parent'
    $mgData = Resolve-MgLevels -FlatMgs $flatMgs
}
Export-Sheet -Data $mgData -Sheet "MgmtGroups"

# ─── Subscriptions ───────────────────────────────────────────────────────────
Write-Host "  [2/12] Subscriptions..." -ForegroundColor Gray

$subscriptions = Invoke-ARGQuery -Query '
resourcecontainers
| where type == "microsoft.resources/subscriptions"
| extend state = tostring(properties.state),
         quotaId = tostring(properties.subscriptionPolicies.quotaId),
         mgParent = tostring(properties.managementGroupAncestorsChain[0].displayName)
| project subscriptionId, name, state, quotaId, mgParent, tags'
Export-Sheet -Data $subscriptions -Sheet "Subscriptions"

# ─── Policy Assignments ──────────────────────────────────────────────────────
Write-Host "  [3/12] Policy Assignments..." -ForegroundColor Gray

$policyAssignments = Invoke-ARGQuery -Query '
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

$customPolicies = Invoke-ARGQuery -Query '
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

$policyCompliance = Invoke-ARGQuery -Query '
policyresources
| where type == "microsoft.policyinsights/policystates"
| where properties.complianceState != "Compliant"
| extend complianceState = tostring(properties.complianceState),
         policyAssignment = tostring(properties.policyAssignmentName),
         policyDefinition = tostring(properties.policyDefinitionName)
| summarize NonCompliantCount=count() by policyAssignment, policyDefinition, complianceState, subscriptionId
| order by NonCompliantCount desc'
Export-Sheet -Data $policyCompliance -Sheet "PolicyCompliance"

# ─── Policy Exemptions ────────────────────────────────────────────────────────
Write-Host "  [5b/12] Policy Exemptions..." -ForegroundColor Gray

$policyExemptions = Invoke-ARGQuery -Query '
policyresources
| where type == "microsoft.authorization/policyexemptions"
| extend displayName = tostring(properties.displayName),
         exemptionCategory = tostring(properties.exemptionCategory),
         policyAssignmentId = tostring(properties.policyAssignmentId),
         expiresOn = tostring(properties.expiresOn)
| project name, displayName, exemptionCategory, policyAssignmentId, expiresOn, subscriptionId, id
| order by expiresOn asc'
Export-Sheet -Data $policyExemptions -Sheet "PolicyExemptions"

# ─── RBAC Role Assignments ───────────────────────────────────────────────────
Write-Host "  [6/12] RBAC Role Assignments..." -ForegroundColor Gray

$roleAssignments = Invoke-ARGQuery -Query '
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

$customRoles = Invoke-ARGQuery -Query '
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

$defenderPlans = Invoke-ARGQuery -Query '
securityresources
| where type == "microsoft.security/pricings"
| extend tier = tostring(properties.pricingTier),
         subPlan = tostring(properties.subPlan)
| project subscriptionId, name, tier, subPlan
| order by subscriptionId asc, name asc'
Export-Sheet -Data $defenderPlans -Sheet "DefenderPlans"

$secureScores = Invoke-ARGQuery -Query '
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

$resourceSummary = Invoke-ARGQuery -Query '
resources
| summarize Count=count() by type, location, subscriptionId
| order by Count desc'
Export-Sheet -Data $resourceSummary -Sheet "ResourceSummary"

# ─── Orphaned Resources ──────────────────────────────────────────────────────
Write-Host "  [10/12] Orphaned Resources..." -ForegroundColor Gray

$orphanedDisks = Invoke-ARGQuery -Query '
resources
| where type == "microsoft.compute/disks"
| where isnull(managedBy) or managedBy == ""
| extend diskSizeGB=properties.diskSizeGB, skuName=sku.name
| project name, resourceGroup, subscriptionId, location, diskSizeGB, skuName'

$orphanedNICs = Invoke-ARGQuery -Query '
resources
| where type == "microsoft.network/networkinterfaces"
| where isnull(properties.virtualMachine) and isnull(properties.privateEndpoint)
| project name, resourceGroup, subscriptionId, location'

$orphanedPIPs = Invoke-ARGQuery -Query '
resources
| where type == "microsoft.network/publicipaddresses"
| where isnull(properties.ipConfiguration) and isnull(properties.natGateway)
| project name, resourceGroup, subscriptionId, location'

$orphanedNSGs = Invoke-ARGQuery -Query '
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

$vnets = Invoke-ARGQuery -Query '
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

$locks = Invoke-ARGQuery -Query '
resources
| where type == "microsoft.authorization/locks"
| extend lockLevel = tostring(properties.level)
| project name, lockLevel, resourceGroup, subscriptionId'
Export-Sheet -Data $locks -Sheet "Locks"

Write-Host "`n📊 Data collection complete." -ForegroundColor Green

# ═══════════════════════════════════════════════════════════════════════════════
# DONE
# ═══════════════════════════════════════════════════════════════════════════════

Write-Host "`n╔══════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║  ✅ Governance Export Complete!               ║" -ForegroundColor Green
Write-Host "╚══════════════════════════════════════════════╝" -ForegroundColor Green
Write-Host @"

📁 Output directory: $outputDir
   • AzureGovernance.xlsx  - Raw data Excel workbook (Hierarchy, Policy, RBAC, Security, Resources, Network, Locks)

💡 In Cloud Shell, use the download icon to get the file.
"@ -ForegroundColor Cyan
