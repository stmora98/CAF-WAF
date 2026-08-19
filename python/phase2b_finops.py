"""Azure FinOps Extended Export - Python port of Invoke-AzureFinOps-CloudShell.ps1.

Approximates the data FinOps Hub (FinOps Toolkit) surfaces, using only
subscription-level Reader access (no billing-account exports required).
"""
from __future__ import annotations

import bootstrap
bootstrap.ensure_dependencies()

from datetime import datetime, timedelta

from common.argquery import run_query
from common.auth import get_credential, list_enabled_subscriptions, resolve_subscription_ids
from common.excelio import add_sheet, make_output_path, new_workbook, save
from common.httpretry import get_bearer_token, request_with_retry


def run(credential=None, subscription_ids=None) -> str:
    credential = credential or get_credential()
    subscription_ids = subscription_ids or resolve_subscription_ids(credential)
    output_path = make_output_path("AzureFinOps")
    wb = new_workbook()

    print("\n=== Azure FinOps Extended Export ===")

    def q(sheet: str, query: str):
        rows = run_query(credential, subscription_ids, query)
        count = add_sheet(wb, sheet, rows, empty_message="No findings")
        print(f"  [{sheet}] {count} rows")
        return rows

    headers = None
    try:
        token = get_bearer_token(credential)
        headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    except Exception as exc:
        print(f"  ! Could not acquire an access token: {exc}")

    # ─── ActualCost - real monthly spend by service (Cost Management Query API) ──
    print("\nQuerying actual cost (last 6 months)...")
    cost_rows = []
    if headers:
        from_date = (datetime.now() - timedelta(days=183)).strftime("%Y-%m-01")
        to_date = datetime.now().strftime("%Y-%m-%d")
        body = {
            "type": "ActualCost",
            "timeframe": "Custom",
            "timePeriod": {"from": from_date, "to": to_date},
            "dataset": {
                "granularity": "Monthly",
                "aggregation": {"totalCost": {"name": "Cost", "function": "Sum"}},
                "grouping": [
                    {"type": "Dimension", "name": "ServiceName"},
                    {"type": "Dimension", "name": "ResourceLocation"},
                ],
            },
        }
        subs = [s for s in list_enabled_subscriptions(credential) if s.subscription_id in subscription_ids]
        for sub in subs:
            url = (f"https://management.azure.com/subscriptions/{sub.subscription_id}"
                   f"/providers/Microsoft.CostManagement/query?api-version=2023-11-01")
            try:
                resp = request_with_retry("POST", url, headers=headers, json_body=body, timeout=60)
                payload = resp.json().get("properties", {})
                columns = [c.get("name") for c in payload.get("columns", [])]
                for row in payload.get("rows", []):
                    record = {"Subscription": sub.display_name, "SubscriptionId": sub.subscription_id}
                    for i, col in enumerate(columns):
                        record[col] = row[i] if i < len(row) else None
                    cost_rows.append(record)
            except Exception as exc:
                status = getattr(getattr(exc, "response", None), "status_code", 0)
                if status in (401, 403):
                    print(f"  ! No permission to read cost data for subscription {sub.display_name} "
                          f"(needs Cost Management Reader).")
                else:
                    print(f"  ! Actual cost unavailable for subscription {sub.display_name}: {exc}")

    count = add_sheet(
        wb, "ActualCost", cost_rows,
        empty_message="No actual cost data found (or permission unavailable - requires Cost Management Reader)",
    )
    print(f"  [ActualCost] {count} rows")

    # ─── ReservationDetails - utilization (billing-account scope, best-effort) ──
    print("\nChecking reservation utilization (requires billing account access)...")
    reservation_detail_rows = []
    if headers:
        try:
            ba_url = "https://management.azure.com/providers/Microsoft.Billing/billingAccounts?api-version=2020-05-01"
            ba_resp = request_with_retry("GET", ba_url, headers=headers, timeout=60)
            from_date = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")
            to_date = datetime.now().strftime("%Y-%m-%d")
            for ba in ba_resp.json().get("value", []):
                filter_expr = f"properties/UsageDate ge '{from_date}' and properties/UsageDate le '{to_date}'"
                import urllib.parse
                url = (f"https://management.azure.com{ba.get('id')}/providers/Microsoft.Consumption/"
                       f"reservationDetails?api-version=2023-05-01&$filter={urllib.parse.quote(filter_expr)}")
                try:
                    resp = request_with_retry("GET", url, headers=headers, timeout=60)
                    for item in resp.json().get("value", []):
                        p = item.get("properties", {})
                        reservation_detail_rows.append({
                            "BillingAccount": ba.get("name"),
                            "ReservationId": p.get("reservationId"),
                            "SkuName": p.get("skuName"),
                            "InstanceFlexibility": p.get("instanceFlexibility"),
                            "TotalReservedQty": p.get("totalReservedQuantity"),
                            "UsedHours": p.get("usedHours"),
                            "UtilizationPct": p.get("utilizationPercentage"),
                            "UsageDate": p.get("usageDate"),
                        })
                except Exception as exc:
                    status = getattr(getattr(exc, "response", None), "status_code", 0)
                    if status in (401, 403):
                        print(f"  ! No permission to read reservation details for billing account "
                              f"{ba.get('name')} (needs Enterprise/Billing Reader).")
                    else:
                        print(f"  ! Reservation details unavailable for billing account {ba.get('name')}: {exc}")
        except Exception as exc:
            print(f"  ! Could not enumerate billing accounts (needs Billing Reader at tenant scope): {exc}")

    count = add_sheet(
        wb, "ReservationDetails", reservation_detail_rows,
        empty_message="No reservation utilization data found (or permission unavailable - requires Billing Reader)",
    )
    print(f"  [ReservationDetails] {count} rows")

    # ─── Budgets - configured budgets vs. current spend ──
    print("\nChecking configured budgets...")
    budget_rows = []
    if headers:
        subs = [s for s in list_enabled_subscriptions(credential) if s.subscription_id in subscription_ids]
        for sub in subs:
            url = (f"https://management.azure.com/subscriptions/{sub.subscription_id}"
                   f"/providers/Microsoft.Consumption/budgets?api-version=2023-05-01")
            try:
                resp = request_with_retry("GET", url, headers=headers, timeout=60)
                for b in resp.json().get("value", []):
                    p = b.get("properties", {})
                    current_spend = p.get("currentSpend", {}) or {}
                    time_period = p.get("timePeriod", {}) or {}
                    budget_rows.append({
                        "Subscription": sub.display_name,
                        "SubscriptionId": sub.subscription_id,
                        "BudgetName": b.get("name"),
                        "Category": p.get("category"),
                        "Amount": p.get("amount"),
                        "CurrentSpend": current_spend.get("amount"),
                        "Currency": current_spend.get("unit"),
                        "TimeGrain": p.get("timeGrain"),
                        "StartDate": time_period.get("startDate"),
                        "EndDate": time_period.get("endDate"),
                    })
            except Exception as exc:
                status = getattr(getattr(exc, "response", None), "status_code", 0)
                if status in (401, 403):
                    print(f"  ! No permission to read budgets for subscription {sub.display_name} "
                          f"(needs Cost Management Reader).")
                else:
                    print(f"  ! Budgets unavailable for subscription {sub.display_name}: {exc}")

    count = add_sheet(
        wb, "Budgets", budget_rows,
        empty_message="No budgets configured (or permission unavailable - requires Cost Management Reader)",
    )
    print(f"  [Budgets] {count} rows")

    # ─── Extended optimization recommendations (Azure Resource Graph) ──
    q("UnattachedPublicIPs", """
resources
| where type =~ "microsoft.network/publicipaddresses"
| where isnull(properties.ipConfiguration)
| project name, resourceGroup, subscriptionId, location,
          sku = tostring(sku.name), allocationMethod = tostring(properties.publicIPAllocationMethod)""")

    q("StoppedVMs", """
resources
| where type =~ "microsoft.compute/virtualmachines"
| extend powerState = tostring(properties.extended.instanceView.powerState.code)
| where powerState == "PowerState/stopped"
| project name, resourceGroup, subscriptionId, location, powerState,
          vmSize = tostring(properties.hardwareProfile.vmSize)""")

    q("BackendlessAppGateways", """
resources
| where type =~ "microsoft.network/applicationgateways"
| extend backendPoolCount = array_length(properties.backendAddressPools)
| where backendPoolCount == 0
| project name, resourceGroup, subscriptionId, location,
          sku = tostring(properties.sku.name)""")

    q("BackendlessLoadBalancers", """
resources
| where type =~ "microsoft.network/loadbalancers"
| extend backendPoolCount = array_length(properties.backendAddressPools)
| where backendPoolCount == 0
| project name, resourceGroup, subscriptionId, location,
          sku = tostring(sku.name)""")

    q("EmptySqlElasticPools", """
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
| project-away poolId""")

    q("NonSpotAKSPools", """
resources
| where type =~ "microsoft.containerservice/managedclusters"
| mv-expand pool = properties.agentPoolProfiles
| extend poolName = tostring(pool.name),
         enableAutoScaling = tobool(pool.enableAutoScaling),
         priority = tostring(pool.scaleSetPriority)
| where enableAutoScaling == true and priority != "Spot"
| project name, poolName, resourceGroup, subscriptionId, location""")

    q("VMsWithoutHybridBenefit", """
resources
| where type =~ "microsoft.compute/virtualmachines"
| extend licenseType = tostring(properties.licenseType),
         osType = tostring(properties.storageProfile.osDisk.osType)
| where osType == "Windows" and licenseType !in ("Windows_Server", "Windows_Client")
| project name, resourceGroup, subscriptionId, location,
          vmSize = tostring(properties.hardwareProfile.vmSize)""")

    q("SqlVMsWithoutHybridBenefit", """
resources
| where type =~ "microsoft.sqlvirtualmachine/sqlvirtualmachines"
| extend licenseType = tostring(properties.sqlServerLicenseType)
| where licenseType == "PAYG"
| project name, resourceGroup, subscriptionId, location, licenseType""")

    save(wb, output_path)
    print("\nFinOps extended export complete!")
    print(f"File: {output_path}\n")
    return output_path


if __name__ == "__main__":
    run()
