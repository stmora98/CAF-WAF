<#
.SYNOPSIS
    Azure WAF/CAF Discovery - Single paste-and-run script for Cloud Shell.
.DESCRIPTION
    Paste this entire script into Azure Cloud Shell (PowerShell).
    It runs 23 Resource Graph queries and exports results to one Excel file.
.EXAMPLE
    # Just paste the whole script into Cloud Shell and press Enter.
    # Download the output file from: ~/AzureDiscovery_<timestamp>.xlsx
#>

# ─── Config ──────────────────────────────────────────────────────────────────
$_outDir = if ($env:AZWORKSHOP_OUTPUT) { $env:AZWORKSHOP_OUTPUT } else { $HOME }
$outputFile = Join-Path $_outDir "AzureDiscovery_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
$subscriptionIds = @($env:AZWORKSHOP_SUBSCRIPTION_IDS -split ',' | Where-Object { $_ })
if ($subscriptionIds.Count -eq 0) {
    $subscriptionIds = @(Get-AzSubscription -ErrorAction Stop | Where-Object { $_.State -eq 'Enabled' } | Select-Object -ExpandProperty Id)
}
if ($subscriptionIds.Count -eq 0) { throw "No enabled subscriptions are accessible in the current tenant." }

# ─── Install ImportExcel if missing ──────────────────────────────────────────
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

# ─── Query runner with pagination ────────────────────────────────────────────
function Get-GraphRows {
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

function Write-Sheet {
    param([object[]]$Rows, [string]$Sheet)
    $all = @($Rows)
    if ($all.Count -eq 0) { $all = @([PSCustomObject]@{ Result = "No resources found" }) }
    $all | Export-Excel -Path $outputFile -WorksheetName $Sheet -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle Medium6
    Write-Host "  [$Sheet] $($all.Count) rows" -ForegroundColor Gray
    return $all
}

function Run-Query {
    param([string]$Query, [string]$Sheet)
    return Write-Sheet -Rows @(Get-GraphRows -Query $Query) -Sheet $Sheet
}

Write-Host "`n=== Azure WAF/CAF Discovery ===" -ForegroundColor Cyan
Write-Host "Output: $outputFile`n" -ForegroundColor Yellow

# ─── Run all queries ─────────────────────────────────────────────────────────

Write-Host "Querying environment..." -ForegroundColor Green

Run-Query -Sheet "Subscriptions" -Query '
resourcecontainers
| where type == "microsoft.resources/subscriptions"
| project id, subscriptionId, name, state=properties.state, tags
| order by name asc'

Run-Query -Sheet "ResourceGroups" -Query '
resourcecontainers
| where type == "microsoft.resources/subscriptions/resourcegroups"
| project id, name, location, subscriptionId, tags, provisioningState=properties.provisioningState
| order by subscriptionId asc, name asc'

Run-Query -Sheet "ResourceSummary" -Query '
resources
| summarize Count=count() by type, location
| order by Count desc'

Run-Query -Sheet "VMs" -Query '
resources
| where type == "microsoft.compute/virtualmachines"
| extend vmSize = properties.hardwareProfile.vmSize,
         osType = properties.storageProfile.osDisk.osType,
         osSku = properties.storageProfile.imageReference.sku,
         powerState = properties.extended.instanceView.powerState.displayStatus,
         zone = tostring(zones[0])
| project id, name, resourceGroup, subscriptionId, location, vmSize, osType, osSku, powerState, zone, tags
| order by name asc'

Run-Query -Sheet "AppServices" -Query '
resources
| where type in ("microsoft.web/sites", "microsoft.web/serverfarms")
| extend kind_=kind, state=properties.state, httpsOnly=properties.httpsOnly
| project id, name, type, resourceGroup, subscriptionId, location, kind_, state, httpsOnly, tags
| order by type asc, name asc'

Run-Query -Sheet "AKS" -Query '
resources
| where type == "microsoft.containerservice/managedclusters"
| extend k8sVersion = properties.kubernetesVersion,
         nodeCount = toint(properties.agentPoolProfiles[0]["count"]),
         nodeVmSize = properties.agentPoolProfiles[0].vmSize,
         networkPlugin = properties.networkProfile.networkPlugin,
         tier = sku.tier
| project id, name, resourceGroup, subscriptionId, location, k8sVersion, nodeCount, nodeVmSize, networkPlugin, tier, tags
| order by name asc'

Run-Query -Sheet "VNets" -Query '
resources
| where type == "microsoft.network/virtualnetworks"
| extend addressSpace = tostring(properties.addressSpace.addressPrefixes),
         subnets = array_length(properties.subnets),
         ddosProtection = properties.enableDdosProtection
| project id, name, resourceGroup, subscriptionId, location, addressSpace, subnets, ddosProtection, tags
| order by name asc'

Run-Query -Sheet "NSGs" -Query '
resources
| where type == "microsoft.network/networksecuritygroups"
| extend rules = array_length(properties.securityRules),
         associatedSubnets = array_length(properties.subnets),
         associatedNICs = array_length(properties.networkInterfaces)
| project id, name, resourceGroup, subscriptionId, location, rules, associatedSubnets, associatedNICs, tags
| order by name asc'

Run-Query -Sheet "LoadBalancers" -Query '
resources
| where type in ("microsoft.network/loadbalancers",
                 "microsoft.network/applicationgateways",
                 "microsoft.network/frontdoors",
                 "microsoft.cdn/profiles")
| extend skuName = sku.name, skuTier = sku.tier
| project id, name, type, resourceGroup, subscriptionId, location, skuName, skuTier, tags
| order by type asc, name asc'

Run-Query -Sheet "Firewalls" -Query '
resources
| where type in ("microsoft.network/azurefirewalls",
                 "microsoft.network/firewallpolicies",
                 "microsoft.network/applicationgatewaywebapplicationfirewallpolicies")
| extend skuName = sku.name, skuTier = sku.tier, threatIntel = properties.threatIntelMode
| project id, name, type, resourceGroup, subscriptionId, location, skuName, skuTier, threatIntel, tags
| order by type asc, name asc'

Run-Query -Sheet "PrivateEndpoints" -Query '
resources
| where type == "microsoft.network/privateendpoints"
| extend target = tostring(properties.privateLinkServiceConnections[0].properties.privateLinkServiceId),
         status = properties.privateLinkServiceConnections[0].properties.privateLinkServiceConnectionState.status
| project id, name, resourceGroup, subscriptionId, location, target, status, tags
| order by name asc'

Run-Query -Sheet "PublicIPs" -Query '
resources
| where type == "microsoft.network/publicipaddresses"
| extend ip = properties.ipAddress,
         allocation = properties.publicIPAllocationMethod,
         skuName = sku.name,
         associatedTo = coalesce(tostring(properties.ipConfiguration.id), "Unassociated")
| project id, name, resourceGroup, subscriptionId, location, ip, allocation, skuName, associatedTo, tags
| order by name asc'

# Which VNets actually host internet-reachable public IPs (used to prioritize DDoS Protection findings).
# Resource Graph can't express this in one query (no `let`, max 3 union/join/mv-expand), so each
# exposure path is queried separately and merged here. A failure only skips the sheet; the
# dashboard then falls back to its default DDoS severity.
try {
    $vnetOf = { param($subnetId) ([string]$subnetId -split '/subnets/')[0].ToLowerInvariant() }
    $links = [System.Collections.Generic.List[object]]::new()
    $addLinks = {
        param($rows, $via)
        foreach ($row in $rows) {
            $links.Add([PSCustomObject]@{
                VNetId = & $vnetOf $row.subnetId
                PublicIpId = ([string]$row.publicIpId).ToLowerInvariant()
                Via = if ($row.via) { $row.via } else { $via }
            })
        }
    }

    & $addLinks (Get-GraphRows -Query '
resources
| where type =~ "microsoft.network/networkinterfaces"
| mv-expand ipc = properties.ipConfigurations
| extend subnetId = tostring(ipc.properties.subnet.id), publicIpId = tostring(ipc.properties.publicIPAddress.id)
| where isnotempty(subnetId) and isnotempty(publicIpId)
| project subnetId, publicIpId') 'Network interface'

    & $addLinks (Get-GraphRows -Query '
resources
| where type =~ "microsoft.network/applicationgateways"
| extend subnetId = tostring(properties.gatewayIPConfigurations[0].properties.subnet.id)
| mv-expand fe = properties.frontendIPConfigurations
| extend publicIpId = tostring(fe.properties.publicIPAddress.id)
| where isnotempty(subnetId) and isnotempty(publicIpId)
| project subnetId, publicIpId') 'Application Gateway'

    & $addLinks (Get-GraphRows -Query '
resources
| where type in~ ("microsoft.network/azurefirewalls", "microsoft.network/bastionhosts", "microsoft.network/virtualnetworkgateways")
| mv-expand ipc = properties.ipConfigurations
| extend subnetId = tostring(ipc.properties.subnet.id), publicIpId = tostring(ipc.properties.publicIPAddress.id),
         via = case(type =~ "microsoft.network/azurefirewalls", "Azure Firewall", type =~ "microsoft.network/bastionhosts", "Azure Bastion", "Virtual network gateway")
| where isnotempty(subnetId) and isnotempty(publicIpId)
| project subnetId, publicIpId, via') 'Edge service'

    # Public load balancers expose the VNets that host their backend pool members (NICs and VMSS/AKS nodes).
    $lbFrontends = @{}
    foreach ($row in (Get-GraphRows -Query '
resources
| where type =~ "microsoft.network/loadbalancers"
| mv-expand fe = properties.frontendIPConfigurations
| extend publicIpId = tostring(fe.properties.publicIPAddress.id)
| where isnotempty(publicIpId)
| project lbId = tolower(id), publicIpId')) {
        if (-not $lbFrontends.ContainsKey($row.lbId)) { $lbFrontends[$row.lbId] = [System.Collections.Generic.List[string]]::new() }
        $lbFrontends[$row.lbId].Add([string]$row.publicIpId)
    }
    $lbMembers = @()
    $lbMembers += Get-GraphRows -Query '
resources
| where type =~ "microsoft.network/networkinterfaces"
| mv-expand ipc = properties.ipConfigurations
| mv-expand pool = ipc.properties.loadBalancerBackendAddressPools
| extend subnetId = tostring(ipc.properties.subnet.id), lbId = tolower(tostring(split(tostring(pool.id), "/backendAddressPools/")[0]))
| where isnotempty(subnetId) and isnotempty(lbId)
| project subnetId, lbId'
    $lbMembers += Get-GraphRows -Query '
resources
| where type =~ "microsoft.compute/virtualmachinescalesets"
| mv-expand nic = properties.virtualMachineProfile.networkProfile.networkInterfaceConfigurations
| mv-expand ipc = nic.properties.ipConfigurations
| mv-expand pool = ipc.properties.loadBalancerBackendAddressPools
| extend subnetId = tostring(ipc.properties.subnet.id), lbId = tolower(tostring(split(tostring(pool.id), "/backendAddressPools/")[0]))
| where isnotempty(subnetId) and isnotempty(lbId)
| project subnetId, lbId'
    foreach ($member in $lbMembers) {
        if ($lbFrontends.ContainsKey($member.lbId)) {
            foreach ($publicIpId in $lbFrontends[$member.lbId]) {
                $links.Add([PSCustomObject]@{ VNetId = & $vnetOf $member.subnetId; PublicIpId = $publicIpId.ToLowerInvariant(); Via = 'Public load balancer' })
            }
        }
    }

    $exposure = @($links | Group-Object VNetId | ForEach-Object {
        [PSCustomObject]@{
            VNetId = $_.Name
            VNetName = ($_.Name -split '/')[-1]
            PublicIPCount = @($_.Group.PublicIpId | Sort-Object -Unique).Count
            ExposedVia = (@($_.Group.Via | Sort-Object -Unique) -join ', ')
        }
    } | Sort-Object VNetName)
    Write-Sheet -Rows $exposure -Sheet "VNetPublicExposure" | Out-Null
} catch {
    Write-Host "  [VNetPublicExposure] skipped: $($_.Exception.Message)" -ForegroundColor DarkYellow
}
Run-Query -Sheet "DNS" -Query '
resources
| where type in ("microsoft.network/dnszones", "microsoft.network/privatednszones")
| extend recordSets = properties.numberOfRecordSets
| project id, name, type, resourceGroup, subscriptionId, location, recordSets, tags
| order by type asc, name asc'

Run-Query -Sheet "Storage" -Query '
resources
| where type == "microsoft.storage/storageaccounts"
| extend skuName = sku.name, kind_=kind,
         httpsOnly = properties.supportsHttpsTrafficOnly,
         tlsVersion = properties.minimumTlsVersion,
         networkAccess = properties.networkAcls.defaultAction,
         publicBlob = properties.allowBlobPublicAccess
| project id, name, resourceGroup, subscriptionId, location, skuName, kind_, httpsOnly, tlsVersion, networkAccess, publicBlob, tags
| order by name asc'

Run-Query -Sheet "Databases" -Query '
resources
| where type in ("microsoft.sql/servers",
                 "microsoft.sql/servers/databases",
                 "microsoft.dbformysql/flexibleservers",
                 "microsoft.dbforpostgresql/flexibleservers",
                 "microsoft.documentdb/databaseaccounts",
                 "microsoft.cache/redis")
| extend skuName=sku.name, skuTier=sku.tier, skuCapacity=sku.capacity
| project id, name, type, resourceGroup, subscriptionId, location, skuName, skuTier, skuCapacity, tags
| order by type asc, name asc'

$keyVaultList = Run-Query -Sheet "KeyVaults" -Query '
resources
| where type == "microsoft.keyvault/vaults"
| extend sku_=properties.sku.name,
         softDelete=properties.enableSoftDelete,
         purgeProtection=properties.enablePurgeProtection,
         rbac=properties.enableRbacAuthorization,
         networkAccess=properties.networkAcls.defaultAction
| project id, name, resourceGroup, subscriptionId, location, sku_, softDelete, purgeProtection, rbac, networkAccess, tags
| order by name asc'

# ─── Key Vault secret/certificate expiration (data-plane, needs Key Vault Get permission per vault) ──
Write-Host "Checking Key Vault secret/certificate expiration..." -ForegroundColor Green
$kvExpirationRows = @()
if ($keyVaultList -and -not ($keyVaultList.Count -eq 1 -and $keyVaultList[0].PSObject.Properties.Name -contains "Result")) {
    if (-not (Get-Module -ListAvailable -Name Az.KeyVault)) {
        Install-Module Az.KeyVault -Scope CurrentUser -Force
    }
    Import-Module Az.KeyVault
    $now = Get-Date
    foreach ($vault in $keyVaultList) {
        try {
            $secrets = Get-AzKeyVaultSecret -VaultName $vault.name -ErrorAction Stop
            foreach ($s in $secrets) {
                $days = if ($s.Expires) { [math]::Round(($s.Expires - $now).TotalDays) } else { $null }
                $kvExpirationRows += [PSCustomObject]@{
                    VaultName = $vault.name; ItemType = "Secret"; ItemName = $s.Name
                    Enabled = $s.Enabled; ExpiresOn = $s.Expires; DaysUntilExpiry = $days
                    ResourceGroup = $vault.resourceGroup; SubscriptionId = $vault.subscriptionId
                }
            }
        } catch {
            Write-Host "  ! Could not read secrets for vault $($vault.name): $($_.Exception.Message)" -ForegroundColor DarkYellow
        }
        try {
            $certs = Get-AzKeyVaultCertificate -VaultName $vault.name -ErrorAction Stop
            foreach ($c in $certs) {
                $days = if ($c.Expires) { [math]::Round(($c.Expires - $now).TotalDays) } else { $null }
                $kvExpirationRows += [PSCustomObject]@{
                    VaultName = $vault.name; ItemType = "Certificate"; ItemName = $c.Name
                    Enabled = $c.Enabled; ExpiresOn = $c.Expires; DaysUntilExpiry = $days
                    ResourceGroup = $vault.resourceGroup; SubscriptionId = $vault.subscriptionId
                }
            }
        } catch {
            Write-Host "  ! Could not read certificates for vault $($vault.name): $($_.Exception.Message)" -ForegroundColor DarkYellow
        }
    }
}
if ($kvExpirationRows.Count -eq 0) {
    $kvExpirationRows = @([PSCustomObject]@{ Result = "No Key Vault secrets/certificates found (or permission unavailable - requires Key Vault Get permission on each vault)" })
}
$kvExpirationRows | Export-Excel -Path $outputFile -WorksheetName "KeyVaultExpirations" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle Medium6
Write-Host "  [KeyVaultExpirations] $($kvExpirationRows.Count) rows" -ForegroundColor Gray

Run-Query -Sheet "ManagedIdentities" -Query '
resources
| where type == "microsoft.managedidentity/userassignedidentities"
| project id, name, resourceGroup, subscriptionId, location, tags
| order by name asc'

Run-Query -Sheet "Monitoring" -Query '
resources
| where type in ("microsoft.insights/components",
                 "microsoft.operationalinsights/workspaces",
                 "microsoft.insights/activitylogalerts",
                 "microsoft.insights/metricalerts",
                 "microsoft.insights/actiongroups")
| extend skuName=sku.name, retention=properties.retentionInDays
| project id, name, type, resourceGroup, subscriptionId, location, skuName, retention, tags
| order by type asc, name asc'

Run-Query -Sheet "PolicyAssignments" -Query '
policyresources
| where type == "microsoft.authorization/policyassignments"
| extend displayName=tostring(properties.displayName),
         enforcement=tostring(properties.enforcementMode),
         definition=tostring(properties.policyDefinitionId),
         scope_=tostring(properties.scope)
| project id, name, displayName, enforcement, definition, scope_, subscriptionId
| order by displayName asc'

Run-Query -Sheet "BackupVaults" -Query '
resources
| where type in ("microsoft.recoveryservices/vaults", "microsoft.dataprotection/backupvaults")
| extend skuName=sku.name, redundancy=properties.storageSettings[0].type
| project id, name, type, resourceGroup, subscriptionId, location, skuName, redundancy, tags
| order by type asc, name asc'

Run-Query -Sheet "BackupProtectedItems" -Query '
resources
| where type == "microsoft.recoveryservices/vaults/backupfabrics/protectioncontainers/protecteditems"
| extend friendlyName=tostring(properties.friendlyName),
         protectionStatus=tostring(properties.protectionStatus),
         lastBackupStatus=tostring(properties.lastBackupStatus),
         lastBackupTime=tostring(properties.lastBackupTime),
         policyName=tostring(properties.policyName),
         protectedItemType=tostring(properties.protectedItemType),
         workloadType=tostring(properties.workloadType)
| project id, friendlyName, protectionStatus, lastBackupStatus, lastBackupTime, policyName, protectedItemType, workloadType, resourceGroup, subscriptionId
| order by protectionStatus asc'

Run-Query -Sheet "ResourceHealthEvents" -Query '
resources
| where type == "microsoft.resourcehealth/events"
| extend eventType=tostring(properties.eventType),
         status=tostring(properties.status),
         title_=tostring(properties["title"]),
         summary=tostring(properties.summary),
         level=tostring(properties.level),
         impactStartTime=tostring(properties.impactStartTime),
         impactMitigationTime=tostring(properties.impactMitigationTime)
| project id, title_, eventType, status, level, impactStartTime, impactMitigationTime, summary, subscriptionId
| order by impactStartTime desc'

Run-Query -Sheet "UnattachedDisks" -Query '
resources
| where type == "microsoft.compute/disks"
| where isnull(managedBy) or managedBy == ""
| extend diskSizeGB=toint(properties.diskSizeGB), skuName=tostring(sku.name), diskState=tostring(properties.diskState)
| project id, name, resourceGroup, subscriptionId, location, diskSizeGB, skuName, diskState, tags
| order by diskSizeGB desc'

Run-Query -Sheet "DeallocatedVMs" -Query '
resources
| where type == "microsoft.compute/virtualmachines"
| where properties.extended.instanceView.powerState.displayStatus == "VM deallocated"
| extend vmSize = properties.hardwareProfile.vmSize
| project id, name, resourceGroup, subscriptionId, location, vmSize, tags
| order by name asc'

Run-Query -Sheet "AdvisorRecommendations" -Query '
advisorresources
| where type == "microsoft.advisor/recommendations"
| extend category=tostring(properties.category),
         impact=tostring(properties.impact),
         impactedField=tostring(properties.impactedField),
         impactedValue=tostring(properties.impactedValue),
         description=tostring(properties.shortDescription.solution)
| project id, category, impact, impactedField, impactedValue, description, subscriptionId
| order by category asc, impact desc'

# ─── Done ────────────────────────────────────────────────────────────────────
Write-Host "`n✅ Discovery complete!" -ForegroundColor Green
Write-Host "📁 File: $outputFile" -ForegroundColor Cyan
Write-Host "💡 In Cloud Shell, click the upload/download icon to download the file.`n" -ForegroundColor Yellow
