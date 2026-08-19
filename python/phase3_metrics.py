"""Azure Monitor Metrics Discovery - Python port of Invoke-AzureMetrics-CloudShell.ps1.

Right-sizing & reliability gaps. Queries metrics per resource (slower than ARG alone).
Covers the last 30 days of metric data.
"""
from __future__ import annotations

import bootstrap
bootstrap.ensure_dependencies()

from datetime import datetime, timedelta, timezone

from azure.mgmt.monitor import MonitorManagementClient

from common.argquery import run_query
from common.auth import get_credential, resolve_subscription_ids
from common.excelio import add_sheet, make_output_path, new_workbook, save

DAYS_BACK = 30


def _avg_max(monitor_client: MonitorManagementClient, resource_id: str, metric_name: str, start, end):
    """Returns (average, maximum) for a metric over the window, or (None, None) on failure.
    Uses the classic per-resource Metrics REST API (azure-mgmt-monitor), the direct
    equivalent of Get-AzMetric -ResourceId, with a daily time grain matching the
    PowerShell script's -TimeGrain 1.00:00:00.
    """
    timespan = f"{start.isoformat()}/{end.isoformat()}"
    try:
        response = monitor_client.metrics.list(
            resource_id,
            timespan=timespan,
            interval="P1D",
            metricnames=metric_name,
            aggregation="Average,Maximum",
        )
    except Exception:
        return None, None

    averages, maximums = [], []
    for metric in response.value:
        for series in metric.timeseries or []:
            for dp in series.data or []:
                if dp.average is not None:
                    averages.append(dp.average)
                if dp.maximum is not None:
                    maximums.append(dp.maximum)

    avg = round(sum(averages) / len(averages), 2) if averages else None
    mx = round(max(maximums), 2) if maximums else None
    return avg, mx


def _avg_only(monitor_client: MonitorManagementClient, resource_id: str, metric_name: str, start, end):
    avg, _ = _avg_max(monitor_client, resource_id, metric_name, start, end)
    return avg


def run(credential=None, subscription_ids=None) -> str:
    credential = credential or get_credential()
    subscription_ids = subscription_ids or resolve_subscription_ids(credential)
    output_path = make_output_path("AzureMetrics")
    wb = new_workbook()

    end_time = datetime.now(timezone.utc)
    start_time = end_time - timedelta(days=DAYS_BACK)

    # resource_uri passed to each call carries its own subscription, so the client's
    # subscription_id is just a placeholder required by the SDK constructor.
    monitor_client = MonitorManagementClient(credential, subscription_id=subscription_ids[0])

    print(f"\n=== Azure Metrics Discovery (Last {DAYS_BACK} days) ===")
    print(f"Output: {output_path}\n")

    def export(sheet: str, rows):
        count = add_sheet(wb, sheet, rows, empty_message="No data found")
        print(f"  [{sheet}] {count} rows")

    # ─── 1. VM Right-Sizing ──────────────────────────────────────────────
    print("[1/6] VM CPU & Memory Utilization...")
    vms = run_query(credential, subscription_ids, """
resources
| where type == "microsoft.compute/virtualmachines"
| where properties.extended.instanceView.powerState.displayStatus == "VM running"
| project id, name, resourceGroup, subscriptionId, location,
          vmSize=properties.hardwareProfile.vmSize
| order by name asc""")

    vm_rows = []
    for i, vm in enumerate(vms, start=1):
        print(f"  ({i}/{len(vms)}) {vm.get('name')}", end="\r")
        cpu_avg, cpu_max = _avg_max(monitor_client, vm["id"], "Percentage CPU", start_time, end_time)
        if cpu_avg is None:
            sizing = "No data"
        elif cpu_avg < 5:
            sizing = "Idle (<5%)"
        elif cpu_avg < 15:
            sizing = "Underutilized (<15%)"
        elif cpu_avg < 80:
            sizing = "Right-sized"
        else:
            sizing = "Saturated (>80%)"
        vm_rows.append({
            "Name": vm.get("name"), "ResourceGroup": vm.get("resourceGroup"),
            "Subscription": vm.get("subscriptionId"), "Location": vm.get("location"),
            "VMSize": vm.get("vmSize"), "AvgCPU_Pct": cpu_avg, "MaxCPU_Pct": cpu_max,
            "Assessment": sizing,
        })
    print()
    export("VM_RightSizing", vm_rows)

    # ─── 2. SQL Database Utilization ─────────────────────────────────────
    print("[2/6] SQL Database DTU/CPU Usage...")
    sql_dbs = run_query(credential, subscription_ids, """
resources
| where type == "microsoft.sql/servers/databases"
| where name != "master"
| project id, name, resourceGroup, subscriptionId, location,
          skuName=sku.name, skuTier=sku.tier
| order by name asc""")

    sql_rows = []
    for i, db in enumerate(sql_dbs, start=1):
        print(f"  ({i}/{len(sql_dbs)}) {db.get('name')}", end="\r")
        dtu_avg = _avg_only(monitor_client, db["id"], "dtu_consumption_percent", start_time, end_time)
        cpu_avg = _avg_only(monitor_client, db["id"], "cpu_percent", start_time, end_time) if dtu_avg is None else None
        storage_used = _avg_only(monitor_client, db["id"], "storage_percent", start_time, end_time)

        usage_metric = dtu_avg if dtu_avg is not None else cpu_avg
        metric_type = "DTU%" if dtu_avg is not None else "CPU%"

        if usage_metric is None:
            sizing = "No data"
        elif usage_metric < 10:
            sizing = "Oversized (<10%)"
        elif usage_metric < 30:
            sizing = "Underutilized (<30%)"
        elif usage_metric < 80:
            sizing = "Right-sized"
        else:
            sizing = "Saturated (>80%)"

        sql_rows.append({
            "Name": db.get("name"), "ResourceGroup": db.get("resourceGroup"),
            "Subscription": db.get("subscriptionId"), "SKU": db.get("skuName"), "Tier": db.get("skuTier"),
            "MetricType": metric_type, "AvgUsage_Pct": usage_metric, "Storage_Pct": storage_used,
            "Assessment": sizing,
        })
    print()
    export("SQL_RightSizing", sql_rows)

    # ─── 3. App Service Plan Utilization ─────────────────────────────────
    print("[3/6] App Service Plan CPU...")
    plans = run_query(credential, subscription_ids, """
resources
| where type == "microsoft.web/serverfarms"
| where sku.tier != "Free" and sku.tier != "Shared"
| project id, name, resourceGroup, subscriptionId, location,
          skuName=sku.name, skuTier=sku.tier, workers=properties.numberOfWorkers
| order by name asc""")

    plan_rows = []
    for i, plan in enumerate(plans, start=1):
        print(f"  ({i}/{len(plans)}) {plan.get('name')}", end="\r")
        cpu_avg = _avg_only(monitor_client, plan["id"], "CpuPercentage", start_time, end_time)
        mem_avg = _avg_only(monitor_client, plan["id"], "MemoryPercentage", start_time, end_time)

        if cpu_avg is None:
            sizing = "No data"
        elif cpu_avg < 5:
            sizing = "Idle (<5%)"
        elif cpu_avg < 20:
            sizing = "Underutilized (<20%)"
        elif cpu_avg < 80:
            sizing = "Right-sized"
        else:
            sizing = "Saturated (>80%)"

        plan_rows.append({
            "Name": plan.get("name"), "ResourceGroup": plan.get("resourceGroup"),
            "Subscription": plan.get("subscriptionId"), "SKU": plan.get("skuName"), "Tier": plan.get("skuTier"),
            "Workers": plan.get("workers"), "AvgCPU_Pct": cpu_avg, "AvgMemory_Pct": mem_avg,
            "Assessment": sizing,
        })
    print()
    export("AppPlan_RightSizing", plan_rows)

    # ─── 4. Storage Account Activity ─────────────────────────────────────
    print("[4/6] Storage Account Activity...")
    storage_accounts = run_query(credential, subscription_ids, """
resources
| where type == "microsoft.storage/storageaccounts"
| project id, name, resourceGroup, subscriptionId, location, skuName=sku.name
| order by name asc""")

    storage_rows = []
    for i, sa in enumerate(storage_accounts, start=1):
        print(f"  ({i}/{len(storage_accounts)}) {sa.get('name')}", end="\r")
        transactions = _avg_only(monitor_client, sa["id"], "Transactions", start_time, end_time)
        availability = _avg_only(monitor_client, sa["id"], "Availability", start_time, end_time)

        if transactions is None:
            activity = "No data"
        elif transactions == 0:
            activity = "Zero activity"
        elif transactions < 10:
            activity = "Minimal activity"
        else:
            activity = "Active"

        storage_rows.append({
            "Name": sa.get("name"), "ResourceGroup": sa.get("resourceGroup"),
            "Subscription": sa.get("subscriptionId"), "SKU": sa.get("skuName"),
            "AvgDailyTxns": transactions, "Availability_Pct": availability, "Assessment": activity,
        })
    print()
    export("Storage_Activity", storage_rows)

    # ─── 5. Diagnostic Settings Coverage ────────────────────────────────
    print("[5/6] Diagnostic Settings Coverage...")
    diag_coverage = run_query(credential, subscription_ids, """
resources
| where type in ("microsoft.compute/virtualmachines",
                 "microsoft.web/sites",
                 "microsoft.sql/servers/databases",
                 "microsoft.network/applicationgateways",
                 "microsoft.network/azurefirewalls",
                 "microsoft.keyvault/vaults",
                 "microsoft.containerservice/managedclusters")
| project id, name, type, resourceGroup, subscriptionId, location
| order by type asc, name asc""")

    diag_rows = []
    for i, res in enumerate(diag_coverage, start=1):
        print(f"  ({i}/{len(diag_coverage)}) {res.get('name')}", end="\r")
        try:
            settings = list(monitor_client.diagnostic_settings.list(res["id"]))
            has_diag = len(settings) > 0
            destinations = []
            for s in settings:
                dest = []
                if s.workspace_id:
                    dest.append("LogAnalytics")
                if s.storage_account_id:
                    dest.append("Storage")
                if s.event_hub_authorization_rule_id:
                    dest.append("EventHub")
                destinations.append("+".join(dest))
            destinations_str = "; ".join(destinations) if has_diag else "None"
        except Exception:
            has_diag = False
            destinations_str = "Error/NotSupported"

        diag_rows.append({
            "Name": res.get("name"), "Type": res.get("type"), "ResourceGroup": res.get("resourceGroup"),
            "Subscription": res.get("subscriptionId"), "HasDiagnostics": has_diag,
            "Destinations": destinations_str,
            "Gap": "OK" if has_diag else "No diagnostics configured",
        })
    print()
    export("DiagnosticsCoverage", diag_rows)

    # ─── 6. Alert Rules Coverage ─────────────────────────────────────────
    print("[6/6] Alert Rules Summary...")
    alert_summary = run_query(credential, subscription_ids, """
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
| order by type asc, name asc""")
    export("AlertRules", alert_summary)

    save(wb, output_path)
    print("\nMetrics discovery complete!")
    print(f"File: {output_path}")
    print("\nNote: this script queries metrics per-resource, so it takes longer than the "
          "Resource Graph-only phases (~1-3 min per 100 resources).\n")
    return output_path


if __name__ == "__main__":
    run()
