"""Azure Governance Visualizer (Lite) - Python port of Invoke-AzureGovernanceViz-CloudShell.ps1.

Collects management group hierarchy, policy, RBAC, and governance data
(AzGovViz-style scope). No service principal required - uses the signed-in user context.
"""
from __future__ import annotations

import bootstrap
bootstrap.ensure_dependencies()

from azure.mgmt.managementgroups import ManagementGroupsMgmtClient

from common.argquery import run_query
from common.auth import get_credential, resolve_subscription_ids
from common.excelio import add_sheet, make_output_path, new_workbook, save


def _flatten_management_groups(node, level: int = 0, parent_path: str = "", parent_id: str = "") -> list:
    """Recursively flattens the ManagementGroup tree returned by a single recurse=True
    call into the same row shape the PowerShell version builds via repeated API calls.
    """
    rows = []
    display_name = getattr(node, "display_name", None) or getattr(node, "name", "")
    current_path = f"{parent_path}/{display_name}" if parent_path else display_name

    children = list(getattr(node, "children", None) or [])
    child_mgs = [c for c in children if "managementgroups" in (getattr(c, "type", "") or "").lower()]
    child_subs = [c for c in children if "subscriptions" in (getattr(c, "type", "") or "").lower()]

    rows.append({
        "Level": level,
        "DisplayName": display_name,
        "Id": getattr(node, "name", ""),
        "Path": current_path,
        "ChildMGs": len(child_mgs),
        "Subscriptions": len(child_subs),
        "SubNames": "; ".join(getattr(c, "display_name", "") or "" for c in child_subs),
        "ParentId": parent_id,
    })

    for child in child_mgs:
        rows.extend(_flatten_management_groups(child, level + 1, current_path, getattr(node, "name", "")))

    return rows


def run(credential=None, subscription_ids=None) -> str:
    credential = credential or get_credential()
    subscription_ids = subscription_ids or resolve_subscription_ids(credential)
    output_path = make_output_path("AzureGovernance")
    wb = new_workbook()

    from azure.mgmt.subscription import SubscriptionClient
    tenant_id = next(iter(SubscriptionClient(credential).tenants.list())).tenant_id

    print("\n=== Azure Governance Visualizer (Lite) ===")
    print(f"Tenant:  {tenant_id}")
    print(f"Output:  {output_path}\n")
    print("Collecting data...")

    def export(sheet: str, rows):
        count = add_sheet(wb, sheet, rows, empty_message="No data found")
        print(f"  [{sheet}] {count} rows")

    # ─── Management Group Hierarchy ──────────────────────────────────────
    print("  [1/12] Management Group Hierarchy...")
    mg_rows = []
    try:
        mg_client = ManagementGroupsMgmtClient(credential)
        root = mg_client.management_groups.get(tenant_id, expand="children", recurse=True)
        mg_rows = _flatten_management_groups(root)
    except Exception as exc:
        print(f"    Falling back to ARG for MG data ({exc})...")
        mg_rows = run_query(credential, subscription_ids, """
resourcecontainers
| where type == "microsoft.management/managementgroups"
| extend displayName = tostring(properties.displayName),
         parent = tostring(properties.details.parent.id)
| project name, displayName, parent""")
    export("MgmtGroups", mg_rows)

    # ─── Subscriptions ────────────────────────────────────────────────────
    print("  [2/12] Subscriptions...")
    export("Subscriptions", run_query(credential, subscription_ids, """
resourcecontainers
| where type == "microsoft.resources/subscriptions"
| extend state = tostring(properties.state),
         quotaId = tostring(properties.subscriptionPolicies.quotaId),
         mgParent = tostring(properties.managementGroupAncestorsChain[0].displayName)
| project subscriptionId, name, state, quotaId, mgParent, tags"""))

    # ─── Policy Assignments ───────────────────────────────────────────────
    print("  [3/12] Policy Assignments...")
    export("PolicyAssignments", run_query(credential, subscription_ids, """
policyresources
| where type == "microsoft.authorization/policyassignments"
| extend displayName = tostring(properties.displayName),
         enforcement = tostring(properties.enforcementMode),
         scope_ = tostring(properties.scope),
         policyDefId = tostring(properties.policyDefinitionId),
         identity_ = tostring(identity.type),
         assignedBy = tostring(properties.metadata.assignedBy)
| project name, displayName, enforcement, scope_, policyDefId, identity_, assignedBy, subscriptionId
| order by scope_ asc"""))

    # ─── Custom Policy Definitions ────────────────────────────────────────
    print("  [4/12] Custom Policy Definitions...")
    export("CustomPolicies", run_query(credential, subscription_ids, """
policyresources
| where type == "microsoft.authorization/policydefinitions"
| where properties.policyType == "Custom"
| extend displayName = tostring(properties.displayName),
         effect = tostring(properties.policyRule.then.effect),
         category = tostring(properties.metadata.category),
         deprecated = tostring(properties.metadata.deprecated)
| project name, displayName, effect, category, deprecated, subscriptionId, id"""))

    # ─── Policy Compliance ────────────────────────────────────────────────
    print("  [5/12] Policy Compliance...")
    export("PolicyCompliance", run_query(credential, subscription_ids, """
policyresources
| where type == "microsoft.policyinsights/policystates"
| where properties.complianceState != "Compliant"
| extend complianceState = tostring(properties.complianceState),
         policyAssignment = tostring(properties.policyAssignmentName),
         policyDefinition = tostring(properties.policyDefinitionName)
| summarize NonCompliantCount=count() by policyAssignment, policyDefinition, complianceState, subscriptionId
| order by NonCompliantCount desc"""))

    # ─── Policy Exemptions ────────────────────────────────────────────────
    print("  [5b/12] Policy Exemptions...")
    export("PolicyExemptions", run_query(credential, subscription_ids, """
policyresources
| where type == "microsoft.authorization/policyexemptions"
| extend displayName = tostring(properties.displayName),
         exemptionCategory = tostring(properties.exemptionCategory),
         policyAssignmentId = tostring(properties.policyAssignmentId),
         expiresOn = tostring(properties.expiresOn)
| project name, displayName, exemptionCategory, policyAssignmentId, expiresOn, subscriptionId, id
| order by expiresOn asc"""))

    # ─── RBAC Role Assignments ────────────────────────────────────────────
    print("  [6/12] RBAC Role Assignments...")
    export("RoleAssignments", run_query(credential, subscription_ids, """
authorizationresources
| where type == "microsoft.authorization/roleassignments"
| extend principalId = tostring(properties.principalId),
         principalType = tostring(properties.principalType),
         roleDefId = tostring(properties.roleDefinitionId),
         scope_ = tostring(properties.scope),
         createdOn = tostring(properties.createdOn)
| project principalId, principalType, roleDefId, scope_, createdOn, subscriptionId"""))

    # ─── Custom Role Definitions ──────────────────────────────────────────
    print("  [7/12] Custom Role Definitions...")
    export("CustomRoles", run_query(credential, subscription_ids, """
authorizationresources
| where type == "microsoft.authorization/roledefinitions"
| where properties.type == "CustomRole"
| extend roleName = tostring(properties.roleName),
         description_ = tostring(properties.description),
         scopes = tostring(properties.assignableScopes)
| project roleName, description_, scopes, id"""))

    # ─── Defender for Cloud ───────────────────────────────────────────────
    print("  [8/12] Microsoft Defender for Cloud...")
    export("DefenderPlans", run_query(credential, subscription_ids, """
securityresources
| where type == "microsoft.security/pricings"
| extend tier = tostring(properties.pricingTier),
         subPlan = tostring(properties.subPlan)
| project subscriptionId, name, tier, subPlan
| order by subscriptionId asc, name asc"""))

    export("SecureScores", run_query(credential, subscription_ids, """
securityresources
| where type == "microsoft.security/securescores"
| extend current_ = todouble(properties.score.current),
         max_ = todouble(properties.score.max),
         pct = todouble(properties.score.percentage)
| project subscriptionId, current_, max_, pct
| order by pct asc"""))

    # ─── Resources Summary ────────────────────────────────────────────────
    print("  [9/12] Resources Summary...")
    export("ResourceSummary", run_query(credential, subscription_ids, """
resources
| summarize Count=count() by type, location, subscriptionId
| order by Count desc"""))

    # ─── Orphaned Resources ───────────────────────────────────────────────
    print("  [10/12] Orphaned Resources...")
    orphaned_disks = run_query(credential, subscription_ids, """
resources
| where type == "microsoft.compute/disks"
| where isnull(managedBy) or managedBy == ""
| extend diskSizeGB=properties.diskSizeGB, skuName=sku.name
| project name, resourceGroup, subscriptionId, location, diskSizeGB, skuName""")

    orphaned_nics = run_query(credential, subscription_ids, """
resources
| where type == "microsoft.network/networkinterfaces"
| where isnull(properties.virtualMachine) and isnull(properties.privateEndpoint)
| project name, resourceGroup, subscriptionId, location""")

    orphaned_pips = run_query(credential, subscription_ids, """
resources
| where type == "microsoft.network/publicipaddresses"
| where isnull(properties.ipConfiguration) and isnull(properties.natGateway)
| project name, resourceGroup, subscriptionId, location""")

    orphaned_nsgs = run_query(credential, subscription_ids, """
resources
| where type == "microsoft.network/networksecuritygroups"
| where isnull(properties.networkInterfaces) and isnull(properties.subnets)
| project name, resourceGroup, subscriptionId, location""")

    all_orphaned = []
    for d in orphaned_disks:
        all_orphaned.append({"Type": "Disk", "Name": d.get("name"), "ResourceGroup": d.get("resourceGroup"),
                              "Subscription": d.get("subscriptionId"),
                              "Detail": f"{d.get('diskSizeGB')} GB ({d.get('skuName')})"})
    for n in orphaned_nics:
        all_orphaned.append({"Type": "NIC", "Name": n.get("name"), "ResourceGroup": n.get("resourceGroup"),
                              "Subscription": n.get("subscriptionId"), "Detail": ""})
    for p in orphaned_pips:
        all_orphaned.append({"Type": "PublicIP", "Name": p.get("name"), "ResourceGroup": p.get("resourceGroup"),
                              "Subscription": p.get("subscriptionId"), "Detail": ""})
    for s in orphaned_nsgs:
        all_orphaned.append({"Type": "NSG", "Name": s.get("name"), "ResourceGroup": s.get("resourceGroup"),
                              "Subscription": s.get("subscriptionId"), "Detail": ""})
    export("OrphanedResources", all_orphaned)

    # ─── Network Topology ─────────────────────────────────────────────────
    print("  [11/12] Network Topology...")
    export("VNets", run_query(credential, subscription_ids, """
resources
| where type == "microsoft.network/virtualnetworks"
| extend addressSpace = tostring(properties.addressSpace.addressPrefixes),
         subnets = array_length(properties.subnets),
         peerings = array_length(properties.virtualNetworkPeerings),
         ddos = properties.enableDdosProtection
| project name, resourceGroup, subscriptionId, location, addressSpace, subnets, peerings, ddos"""))

    # ─── Resource Locks ───────────────────────────────────────────────────
    print("  [12/12] Resource Locks...")
    export("Locks", run_query(credential, subscription_ids, """
resources
| where type == "microsoft.authorization/locks"
| extend lockLevel = tostring(properties.level)
| project name, lockLevel, resourceGroup, subscriptionId"""))

    save(wb, output_path)
    print("\nGovernance export complete!")
    print(f"File: {output_path}\n")
    return output_path


if __name__ == "__main__":
    run()
