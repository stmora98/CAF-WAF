<#
.SYNOPSIS
    Azure Environment Discovery Script for WAF/CAF Workshops.

.DESCRIPTION
    Queries Azure Resource Graph and ARM APIs to collect a comprehensive inventory
    of a customer's Azure environment across all WAF pillars (Reliability, Security,
    Cost Optimization, Operational Excellence, Performance Efficiency).
    Results are exported to an Excel workbook with one tab per category.

.PARAMETER OutputPath
    Path for the output Excel file. Defaults to "AzureDiscovery_<timestamp>.xlsx"
    in the current directory.

.PARAMETER TenantId
    Optional. Specify a tenant ID if you need to target a specific tenant.

.EXAMPLE
    # Run in Azure Cloud Shell or local PowerShell with Az module
    .\Invoke-AzureDiscovery.ps1

.EXAMPLE
    .\Invoke-AzureDiscovery.ps1 -OutputPath "C:\Reports\CustomerDiscovery.xlsx"

.NOTES
    Prerequisites:
      - Az PowerShell module (Az.Accounts, Az.ResourceGraph)
      - ImportExcel module (auto-installed if missing)
      - Authenticated Azure session (Connect-AzAccount)
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$OutputPath,

    [Parameter()]
    [string]$TenantId,

    # Clears any cached Az context and disables WAM's silent single-account SSO for this
    # run only, so the full sign-in page is shown and you can pick/switch accounts.
    [Parameter()]
    [switch]$ForceAccountSelection
)

#region ─── Setup ───────────────────────────────────────────────────────────────

$ErrorActionPreference = 'Stop'

if (-not $OutputPath) {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $OutputPath = Join-Path $PWD "AzureDiscovery_$timestamp.xlsx"
}

# Ensure required modules
function Ensure-Module {
    param([string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Host "Installing module $Name..." -ForegroundColor Yellow
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $Name -Force
}

Ensure-Module 'Az.Accounts'
Ensure-Module 'Az.ResourceGraph'
Ensure-Module 'ImportExcel'

# Prefers the WAM broker (Windows Hello/Conditional Access support); falls back to device-code login if the broker fails.
# -ForceAccountSelection disables WAM (process scope only, doesn't touch the saved user preference) so
# WAM's silent single-cached-account SSO is bypassed and the full account-picker sign-in page is shown.
function Connect-AzAccountWithWamFallback {
    param(
        [string]$TenantId,
        [switch]$ForceAccountSelection
    )

    $connectParams = @{}
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }

    if ($ForceAccountSelection) {
        Write-Host "  Clearing cached sign-in and forcing the account picker..." -ForegroundColor DarkGray
        Disconnect-AzAccount -ErrorAction SilentlyContinue | Out-Null
        if (Get-Command Update-AzConfig -ErrorAction SilentlyContinue) {
            try { Update-AzConfig -EnableLoginByWam $false -Scope Process -ErrorAction Stop | Out-Null } catch { }
        }
        Connect-AzAccount @connectParams -ErrorAction Stop
        return
    }

    if (Get-Command Update-AzConfig -ErrorAction SilentlyContinue) {
        try {
            Update-AzConfig -EnableLoginByWam $true -ErrorAction Stop | Out-Null
        } catch {
            Write-Host "  Could not enable WAM broker login ($($_.Exception.Message)) - continuing with default login method." -ForegroundColor DarkYellow
        }
    }
    try {
        Connect-AzAccount @connectParams -ErrorAction Stop
    } catch {
        Write-Host "  WAM sign-in failed ($($_.Exception.Message)). Retrying with device code authentication..." -ForegroundColor DarkYellow
        if (Get-Command Update-AzConfig -ErrorAction SilentlyContinue) {
            try { Update-AzConfig -EnableLoginByWam $false -ErrorAction Stop | Out-Null } catch { }
        }
        Connect-AzAccount @connectParams -UseDeviceAuthentication
    }
}

# Verify authentication
$context = Get-AzContext
if ($ForceAccountSelection -or -not $context) {
    Write-Host "Not authenticated. Running Connect-AzAccount..." -ForegroundColor Yellow
    Connect-AzAccountWithWamFallback -TenantId $TenantId -ForceAccountSelection:$ForceAccountSelection
    $context = Get-AzContext
}

Write-Host "`n=== Azure Discovery for WAF/CAF Workshop ===" -ForegroundColor Cyan
Write-Host "Tenant:  $($context.Tenant.Id)" -ForegroundColor Gray
Write-Host "Account: $($context.Account.Id)" -ForegroundColor Gray
Write-Host "Output:  $OutputPath`n" -ForegroundColor Gray

#endregion

#region ─── Helper Functions ────────────────────────────────────────────────────

function Invoke-ResourceGraphQuery {
    <#
    .SYNOPSIS
        Runs an ARG query with automatic pagination (1000 rows per page).
    #>
    param(
        [Parameter(Mandatory)][string]$Query,
        [string]$Description = "query"
    )

    Write-Host "  Querying: $Description..." -ForegroundColor DarkGray
    $results = @()
    $skipToken = $null

    do {
        $params = @{
            Query = $Query
            First = 1000
        }
        if ($skipToken) { $params['SkipToken'] = $skipToken }

        $response = Search-AzGraph @params
        $results += $response.Data
        $skipToken = $response.SkipToken
    } while ($skipToken)

    Write-Host "    Found $($results.Count) records." -ForegroundColor DarkGray
    return $results
}

function Export-ToExcel {
    param(
        [Parameter(Mandatory)][object[]]$Data,
        [Parameter(Mandatory)][string]$WorksheetName,
        [string]$Path = $script:OutputPath
    )

    if ($Data.Count -eq 0) {
        # Write a placeholder row so the tab still exists
        $Data = @([PSCustomObject]@{ Info = "No resources found" })
    }

    $excelParams = @{
        Path          = $Path
        WorksheetName = $WorksheetName
        AutoSize      = $true
        AutoFilter    = $true
        FreezeTopRow  = $true
        BoldTopRow    = $true
        TableStyle    = 'Medium6'
    }

    $Data | Export-Excel @excelParams
}

#endregion

#region ─── 1. Subscriptions Overview ──────────────────────────────────────────

Write-Host "[1/14] Subscriptions Overview" -ForegroundColor Green

$subscriptions = Invoke-ResourceGraphQuery -Description "Subscriptions" -Query @"
resourcecontainers
| where type == 'microsoft.resources/subscriptions'
| project subscriptionId, name, properties.state, tags
| order by name asc
"@

Export-ToExcel -Data $subscriptions -WorksheetName "Subscriptions"

#endregion

#region ─── 2. Resource Groups ─────────────────────────────────────────────────

Write-Host "[2/14] Resource Groups" -ForegroundColor Green

$resourceGroups = Invoke-ResourceGraphQuery -Description "Resource Groups" -Query @"
resourcecontainers
| where type == 'microsoft.resources/subscriptions/resourcegroups'
| project name, location, subscriptionId, tags, properties.provisioningState
| order by subscriptionId asc, name asc
"@

Export-ToExcel -Data $resourceGroups -WorksheetName "ResourceGroups"

#endregion

#region ─── 3. Resource Summary by Type ────────────────────────────────────────

Write-Host "[3/14] Resource Summary by Type" -ForegroundColor Green

$resourceSummary = Invoke-ResourceGraphQuery -Description "Resource counts by type" -Query @"
resources
| summarize Count=count() by type, location
| order by Count desc
"@

Export-ToExcel -Data $resourceSummary -WorksheetName "ResourceSummary"

#endregion

#region ─── 4. Compute (VMs, VMSS, App Services, AKS, Functions) ───────────────

Write-Host "[4/14] Compute Resources" -ForegroundColor Green

$vms = Invoke-ResourceGraphQuery -Description "Virtual Machines" -Query @"
resources
| where type == 'microsoft.compute/virtualmachines'
| extend vmSize = properties.hardwareProfile.vmSize,
         osType = properties.storageProfile.osDisk.osType,
         osPublisher = properties.storageProfile.imageReference.publisher,
         osSku = properties.storageProfile.imageReference.sku,
         powerState = properties.extended.instanceView.powerState.displayStatus,
         availabilityZone = tostring(zones[0])
| project name, resourceGroup, subscriptionId, location, vmSize, osType,
          osPublisher, osSku, powerState, availabilityZone, tags
| order by subscriptionId asc, name asc
"@

Export-ToExcel -Data $vms -WorksheetName "VMs"

$appServices = Invoke-ResourceGraphQuery -Description "App Services & Functions" -Query @"
resources
| where type in ('microsoft.web/sites', 'microsoft.web/serverfarms')
| extend kind_ = kind,
         sku = properties.sku,
         state = properties.state,
         httpsOnly = properties.httpsOnly,
         reserved = properties.reserved
| project name, type, resourceGroup, subscriptionId, location, kind_,
          sku, state, httpsOnly, reserved, tags
| order by type asc, name asc
"@

Export-ToExcel -Data $appServices -WorksheetName "AppServices"

$aks = Invoke-ResourceGraphQuery -Description "AKS Clusters" -Query @"
resources
| where type == 'microsoft.containerservice/managedclusters'
| extend kubernetesVersion = properties.kubernetesVersion,
         nodeCount = properties.agentPoolProfiles[0].count,
         nodeVmSize = properties.agentPoolProfiles[0].vmSize,
         networkPlugin = properties.networkProfile.networkPlugin,
         tier = sku.tier
| project name, resourceGroup, subscriptionId, location, kubernetesVersion,
          nodeCount, nodeVmSize, networkPlugin, tier, tags
| order by name asc
"@

Export-ToExcel -Data $aks -WorksheetName "AKS"

#endregion

#region ─── 5. Networking ──────────────────────────────────────────────────────

Write-Host "[5/14] Networking" -ForegroundColor Green

$vnets = Invoke-ResourceGraphQuery -Description "Virtual Networks" -Query @"
resources
| where type == 'microsoft.network/virtualnetworks'
| extend addressSpace = tostring(properties.addressSpace.addressPrefixes),
         subnetsCount = array_length(properties.subnets),
         enableDdosProtection = properties.enableDdosProtection
| project name, resourceGroup, subscriptionId, location, addressSpace,
          subnetsCount, enableDdosProtection, tags
| order by name asc
"@

Export-ToExcel -Data $vnets -WorksheetName "VNets"

$nsgs = Invoke-ResourceGraphQuery -Description "Network Security Groups" -Query @"
resources
| where type == 'microsoft.network/networksecuritygroups'
| extend rulesCount = array_length(properties.securityRules),
         subnets = array_length(properties.subnets),
         nics = array_length(properties.networkInterfaces)
| project name, resourceGroup, subscriptionId, location, rulesCount, subnets, nics, tags
| order by name asc
"@

Export-ToExcel -Data $nsgs -WorksheetName "NSGs"

$lbs = Invoke-ResourceGraphQuery -Description "Load Balancers & App Gateways" -Query @"
resources
| where type in ('microsoft.network/loadbalancers',
                 'microsoft.network/applicationgateways',
                 'microsoft.network/frontdoors',
                 'microsoft.cdn/profiles')
| extend skuName = sku.name, skuTier = sku.tier
| project name, type, resourceGroup, subscriptionId, location, skuName, skuTier, tags
| order by type asc, name asc
"@

Export-ToExcel -Data $lbs -WorksheetName "LoadBalancers"

$firewalls = Invoke-ResourceGraphQuery -Description "Firewalls & WAFs" -Query @"
resources
| where type in ('microsoft.network/azurefirewalls',
                 'microsoft.network/firewallpolicies',
                 'microsoft.network/applicationgatewaywebapplicationfirewallpolicies')
| extend skuName = sku.name, skuTier = sku.tier,
         threatIntelMode = properties.threatIntelMode
| project name, type, resourceGroup, subscriptionId, location, skuName, skuTier,
          threatIntelMode, tags
| order by type asc, name asc
"@

Export-ToExcel -Data $firewalls -WorksheetName "Firewalls"

$privateEndpoints = Invoke-ResourceGraphQuery -Description "Private Endpoints" -Query @"
resources
| where type == 'microsoft.network/privateendpoints'
| extend targetResource = tostring(properties.privateLinkServiceConnections[0].properties.privateLinkServiceId),
         connectionStatus = properties.privateLinkServiceConnections[0].properties.privateLinkServiceConnectionState.status
| project name, resourceGroup, subscriptionId, location, targetResource, connectionStatus, tags
| order by name asc
"@

Export-ToExcel -Data $privateEndpoints -WorksheetName "PrivateEndpoints"

#endregion

#region ─── 6. Storage ─────────────────────────────────────────────────────────

Write-Host "[6/14] Storage" -ForegroundColor Green

$storage = Invoke-ResourceGraphQuery -Description "Storage Accounts" -Query @"
resources
| where type == 'microsoft.storage/storageaccounts'
| extend skuName = sku.name,
         kind_ = kind,
         httpsOnly = properties.supportsHttpsTrafficOnly,
         minimumTlsVersion = properties.minimumTlsVersion,
         networkDefaultAction = properties.networkAcls.defaultAction,
         allowBlobPublicAccess = properties.allowBlobPublicAccess,
         replication = sku.name
| project name, resourceGroup, subscriptionId, location, skuName, kind_,
          httpsOnly, minimumTlsVersion, networkDefaultAction,
          allowBlobPublicAccess, tags
| order by name asc
"@

Export-ToExcel -Data $storage -WorksheetName "Storage"

#endregion

#region ─── 7. Databases ───────────────────────────────────────────────────────

Write-Host "[7/14] Databases" -ForegroundColor Green

$databases = Invoke-ResourceGraphQuery -Description "Database Services" -Query @"
resources
| where type in ('microsoft.sql/servers',
                 'microsoft.sql/servers/databases',
                 'microsoft.dbformysql/flexibleservers',
                 'microsoft.dbforpostgresql/flexibleservers',
                 'microsoft.documentdb/databaseaccounts',
                 'microsoft.cache/redis')
| extend skuName = sku.name,
         skuTier = sku.tier,
         skuCapacity = sku.capacity
| project name, type, resourceGroup, subscriptionId, location,
          skuName, skuTier, skuCapacity, tags
| order by type asc, name asc
"@

Export-ToExcel -Data $databases -WorksheetName "Databases"

#endregion

#region ─── 8. Security ────────────────────────────────────────────────────────

Write-Host "[8/14] Security Resources" -ForegroundColor Green

$keyVaults = Invoke-ResourceGraphQuery -Description "Key Vaults" -Query @"
resources
| where type == 'microsoft.keyvault/vaults'
| extend skuName = properties.sku.name,
         enableSoftDelete = properties.enableSoftDelete,
         enablePurgeProtection = properties.enablePurgeProtection,
         enableRbac = properties.enableRbacAuthorization,
         networkDefaultAction = properties.networkAcls.defaultAction
| project name, resourceGroup, subscriptionId, location, skuName,
          enableSoftDelete, enablePurgeProtection, enableRbac,
          networkDefaultAction, tags
| order by name asc
"@

Export-ToExcel -Data $keyVaults -WorksheetName "KeyVaults"

$managedIdentities = Invoke-ResourceGraphQuery -Description "Managed Identities" -Query @"
resources
| where type in ('microsoft.managedidentity/userassignedidentities')
| project name, resourceGroup, subscriptionId, location, tags
| order by name asc
"@

Export-ToExcel -Data $managedIdentities -WorksheetName "ManagedIdentities"

#endregion

#region ─── 9. Monitoring & Observability ──────────────────────────────────────

Write-Host "[9/14] Monitoring & Observability" -ForegroundColor Green

$monitoring = Invoke-ResourceGraphQuery -Description "Monitoring Resources" -Query @"
resources
| where type in ('microsoft.insights/components',
                 'microsoft.operationalinsights/workspaces',
                 'microsoft.insights/activitylogalerts',
                 'microsoft.insights/metricalerts',
                 'microsoft.insights/actiongroups',
                 'microsoft.monitor/accounts')
| extend skuName = sku.name, retentionDays = properties.retentionInDays
| project name, type, resourceGroup, subscriptionId, location, skuName, retentionDays, tags
| order by type asc, name asc
"@

Export-ToExcel -Data $monitoring -WorksheetName "Monitoring"

#endregion

#region ─── 10. Governance (Policies) ──────────────────────────────────────────

Write-Host "[10/14] Governance - Policy Assignments" -ForegroundColor Green

$policies = Invoke-ResourceGraphQuery -Description "Policy Assignments" -Query @"
policyresources
| where type == 'microsoft.authorization/policyassignments'
| extend displayName = properties.displayName,
         enforcementMode = properties.enforcementMode,
         policyDefinitionId = properties.policyDefinitionId,
         scope = properties.scope
| project name, displayName, enforcementMode, policyDefinitionId, scope, subscriptionId
| order by displayName asc
"@

Export-ToExcel -Data $policies -WorksheetName "PolicyAssignments"

#endregion

#region ─── 11. Reliability - Backup & DR ──────────────────────────────────────

Write-Host "[11/14] Reliability - Backup & Recovery" -ForegroundColor Green

$backupVaults = Invoke-ResourceGraphQuery -Description "Recovery Services Vaults" -Query @"
resources
| where type in ('microsoft.recoveryservices/vaults',
                 'microsoft.dataprotection/backupvaults')
| extend skuName = sku.name, redundancy = properties.storageSettings[0].type
| project name, type, resourceGroup, subscriptionId, location, skuName, redundancy, tags
| order by type asc, name asc
"@

Export-ToExcel -Data $backupVaults -WorksheetName "BackupVaults"

#endregion

#region ─── 12. Public IPs & DNS ───────────────────────────────────────────────

Write-Host "[12/14] Public IPs & DNS" -ForegroundColor Green

$publicIPs = Invoke-ResourceGraphQuery -Description "Public IP Addresses" -Query @"
resources
| where type == 'microsoft.network/publicipaddresses'
| extend ipAddress = properties.ipAddress,
         allocationMethod = properties.publicIPAllocationMethod,
         skuName = sku.name,
         associatedTo = coalesce(
             tostring(properties.ipConfiguration.id),
             'Unassociated'
         )
| project name, resourceGroup, subscriptionId, location, ipAddress,
          allocationMethod, skuName, associatedTo, tags
| order by name asc
"@

Export-ToExcel -Data $publicIPs -WorksheetName "PublicIPs"

$dns = Invoke-ResourceGraphQuery -Description "DNS Zones" -Query @"
resources
| where type in ('microsoft.network/dnszones', 'microsoft.network/privatednszones')
| extend recordSetCount = properties.numberOfRecordSets
| project name, type, resourceGroup, subscriptionId, location, recordSetCount, tags
| order by type asc, name asc
"@

Export-ToExcel -Data $dns -WorksheetName "DNS"

#endregion

#region ─── 13. Cost Optimization - Unused/Unattached Resources ────────────────

Write-Host "[13/14] Cost Optimization - Potentially Unused Resources" -ForegroundColor Green

$unattachedDisks = Invoke-ResourceGraphQuery -Description "Unattached Managed Disks" -Query @"
resources
| where type == 'microsoft.compute/disks'
| where isnull(managedBy) or managedBy == ''
| extend diskSizeGB = properties.diskSizeGB,
         skuName = sku.name,
         diskState = properties.diskState
| project name, resourceGroup, subscriptionId, location, diskSizeGB, skuName, diskState, tags
| order by diskSizeGB desc
"@

Export-ToExcel -Data $unattachedDisks -WorksheetName "UnattachedDisks"

$stoppedVMs = Invoke-ResourceGraphQuery -Description "Deallocated VMs" -Query @"
resources
| where type == 'microsoft.compute/virtualmachines'
| where properties.extended.instanceView.powerState.displayStatus == 'VM deallocated'
| extend vmSize = properties.hardwareProfile.vmSize
| project name, resourceGroup, subscriptionId, location, vmSize, tags
| order by name asc
"@

Export-ToExcel -Data $stoppedVMs -WorksheetName "DeallocatedVMs"

#endregion

#region ─── 14. Advisor Recommendations ────────────────────────────────────────

Write-Host "[14/14] Azure Advisor Recommendations" -ForegroundColor Green

$advisor = Invoke-ResourceGraphQuery -Description "Advisor Recommendations" -Query @"
advisorresources
| where type == 'microsoft.advisor/recommendations'
| extend category = properties.category,
         impact = properties.impact,
         impactedField = properties.impactedField,
         impactedValue = properties.impactedValue,
         shortDescription = tostring(properties.shortDescription.solution)
| project category, impact, impactedField, impactedValue, shortDescription, subscriptionId
| order by category asc, impact desc
"@

Export-ToExcel -Data $advisor -WorksheetName "AdvisorRecommendations"

#endregion

#region ─── Summary & Finish ───────────────────────────────────────────────────

Write-Host "`n✅ Discovery complete!" -ForegroundColor Green
Write-Host "   Output: $OutputPath" -ForegroundColor Cyan
Write-Host "`n   Worksheets included:" -ForegroundColor White
Write-Host "     1.  Subscriptions        - Tenant subscription inventory"
Write-Host "     2.  ResourceGroups        - All resource groups"
Write-Host "     3.  ResourceSummary       - Resource counts by type & region"
Write-Host "     4.  VMs                   - Virtual machines detail"
Write-Host "     5.  AppServices           - App Services, Functions, Plans"
Write-Host "     6.  AKS                   - Kubernetes clusters"
Write-Host "     7.  VNets                 - Virtual networks"
Write-Host "     8.  NSGs                  - Network Security Groups"
Write-Host "     9.  LoadBalancers         - LBs, AppGW, Front Door, CDN"
Write-Host "     10. Firewalls             - Azure Firewalls & WAF policies"
Write-Host "     11. PrivateEndpoints      - Private endpoint connections"
Write-Host "     12. Storage               - Storage accounts & config"
Write-Host "     13. Databases             - SQL, PostgreSQL, CosmosDB, Redis"
Write-Host "     14. KeyVaults             - Key Vault configurations"
Write-Host "     15. ManagedIdentities     - User-assigned identities"
Write-Host "     16. Monitoring            - Log Analytics, App Insights, Alerts"
Write-Host "     17. PolicyAssignments     - Azure Policy assignments"
Write-Host "     18. BackupVaults          - Recovery Services & Backup vaults"
Write-Host "     19. PublicIPs             - Public IP addresses"
Write-Host "     20. DNS                   - DNS & Private DNS zones"
Write-Host "     21. UnattachedDisks       - Orphaned managed disks"
Write-Host "     22. DeallocatedVMs        - Stopped/deallocated VMs"
Write-Host "     23. AdvisorRecommendations- Advisor suggestions"
Write-Host ""

#endregion
